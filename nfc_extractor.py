#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SIGA SEFAZ-CE - Extrator Completo de NFC-e e Compilador de XMLs (Threads e ZIP)

Este script unifica o fluxo:
1. Coleta as chaves de acesso detalhadas das NFC-e no SIGA para o mês selecionado.
2. Associa com códigos internos da planilha 'empresas cod.xlsx'.
3. Realiza o download multithread (40 threads) dos XMLs fiscais da SEFAZ.
4. Limpa namespaces dos XMLs e organiza em pastas 'EMPRESA-/MMAAAA/'.
5. Compacta os XMLs em arquivos ZIP separados por empresa.
"""

import os
import sys
import json
import time
import base64
import csv
import re
import threading
import zipfile
import xml.etree.ElementTree as ET
import requests
import urllib3

# Ocultar avisos de SSL (InsecureRequestWarning) já que a chamada usa verify=False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    import openpyxl
except ImportError:
    print("\n❌ A biblioteca 'openpyxl' é necessária para criar arquivos Excel.")
    print("Execute o seguinte comando no terminal para instalá-la:")
    print("pip install openpyxl")
    sys.exit(1)

CONFIG_FILE = "config.json"
EXCEL_FILE = "empresas cod.xlsx"
COMPANIES_CSV = "siga_empresas.csv"
TOKEN_FILE = "api_token.txt"

def load_config():
    """Carrega as credenciais em cache."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️ Erro ao carregar config.json: {e}")
    return {}

def save_config(config):
    """Grava as credenciais em cache."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"⚠️ Erro ao salvar config.json: {e}")

def load_api_token():
    """Carrega o token da API da SEFAZ do arquivo api_token.txt."""
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            print(f"⚠️ Erro ao carregar {TOKEN_FILE}: {e}")
    return ""

def save_api_token(token):
    """Grava o token da API da SEFAZ no arquivo api_token.txt."""
    try:
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(token)
    except Exception as e:
        print(f"⚠️ Erro ao gravar {TOKEN_FILE}: {e}")

def is_token_expired(token):
    """Verifica se o token JWT de acesso do SIGA está expirado ou prestes a expirar (< 30s)."""
    if not token:
        return True
    try:
        parts = token.split('.')
        if len(parts) != 3:
            return True
        payload_b64 = parts[1]
        payload_b64 += '=' * (-len(payload_b64) % 4)
        payload_json = base64.urlsafe_b64decode(payload_b64).decode('utf-8')
        payload = json.loads(payload_json)
        
        exp = payload.get('exp', 0)
        return (exp - time.time()) < 30
    except Exception:
        return True

def refresh_access_token(token_url, refresh_token, client_id):
    """Renova o token de acesso do SIGA no Keycloak."""
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': client_id
    }
    try:
        response = requests.post(token_url, headers=headers, data=data, timeout=15)
        if response.status_code == 200:
            res_json = response.json()
            return res_json.get('access_token'), res_json.get('refresh_token', refresh_token)
        else:
            print(f"\n❌ Erro ao renovar token SIGA (HTTP {response.status_code}): {response.text}")
            return None, None
    except Exception as e:
        print(f"\n❌ Erro de conexão ao renovar token SIGA: {e}")
        return None, None

def normalize_name(name):
    """Normaliza o nome da empresa para cruzamento flexível."""
    if not name:
        return ""
    name = name.upper()
    name = re.sub(r'[^A-Z0-9\s]', '', name)
    endings = [
        r'\bLTDA\b', r'\bME\b', r'\bEPP\b', r'\bEIRELI\b', r'\bSA\b', r'\bS\b',
        r'\bLIMITADA\b', r'\bMICROEMPRESA\b', r'\bMICRO\b', r'\bEMPRESA\b'
    ]
    for ending in endings:
        name = re.sub(ending, '', name)
    return " ".join(name.split())

def sanitize_filename(filename):
    """Remove caracteres inválidos no Windows para evitar erros de gravação de arquivos."""
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def clean_input_val(val):
    """Limpa o valor inserido removendo aspas extras, vírgulas ou chaves de JSON."""
    if not val:
        return ""
    val = val.strip()
    # Se o usuário colou a linha inteira do JSON: "access_token": "ey..."
    if ":" in val:
        parts = val.split(":", 1)
        val = parts[1].strip()
    # Remove aspas extras, vírgulas e espaços
    val = val.strip(" \t\n\r\"',;")
    return val

def ensure_valid_token(config, access_token, refresh_token, token_url, client_id):
    """Garante que o token de acesso seja válido, renovando se necessário ou pedindo novas credenciais."""
    if not is_token_expired(access_token):
        return access_token, refresh_token
        
    print("\n🔄 Token de Acesso expirado ou prestes a expirar. Renovando... ", end="", flush=True)
    if refresh_token:
        new_acc, new_ref = refresh_access_token(token_url, refresh_token, client_id)
        if new_acc:
            access_token = new_acc
            refresh_token = new_ref
            config.update({"access_token": access_token, "refresh_token": refresh_token})
            save_config(config)
            print("Renovado!")
            return access_token, refresh_token
            
    # Se falhou a renovação automatizada, pede ao usuário
    print("\n⚠️ O Refresh Token expirou ou é inválido (erro ao renovar).")
    print("Por favor, obtenha novos tokens na aba Network (F12) do SIGA.")
    while True:
        novo_access = clean_input_val(input("Cole o novo access_token: "))
        novo_refresh = clean_input_val(input("Cole o novo refresh_token: "))
        
        if novo_access and novo_refresh:
            access_token = novo_access
            refresh_token = novo_refresh
            config.update({"access_token": access_token, "refresh_token": refresh_token})
            save_config(config)
            print("✅ Tokens atualizados! Retomando extração...")
            return access_token, refresh_token
        else:
            opcao = input("Entrada inválida. Deseja tentar novamente? (S/n): ").strip().lower()
            if opcao == 'n':
                print("❌ Falha crítica ao obter credenciais. Saindo.")
                sys.exit(1)

def load_excel_mapping(filepath):
    """Lê a planilha Excel com os códigos internos e nomes das empresas."""
    if not os.path.exists(filepath):
        print(f"❌ Planilha Excel '{filepath}' não encontrada no diretório atual.")
        sys.exit(1)
        
    print(f"📖 Carregando planilha '{filepath}'...")
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb.active
    
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        print("❌ A planilha Excel está vazia.")
        sys.exit(1)
        
    # Mapeia colunas dinamicamente
    headers = [str(h).strip().lower() for h in rows[0] if h is not None]
    code_idx = 0
    name_idx = 1
    
    for idx, h in enumerate(headers):
        if "cod" in h:
            code_idx = idx
        elif "nome" in h or "empresa" in h or "raz" in h:
            name_idx = idx
            
    print(f"📌 Colunas identificadas: Código em '{rows[0][code_idx]}', Nome em '{rows[0][name_idx]}'")
    
    mapping = {}
    for r in rows[1:]:
        code = r[code_idx]
        name = r[name_idx]
        if code is not None and name is not None:
            norm_name = normalize_name(str(name))
            mapping[norm_name] = {
                "codigo": code,
                "nome_original": name
            }
    print(f"✅ Mapeados {len(mapping)} códigos de empresas da planilha.")
    return mapping

def find_company_code(razao_social, mapping):
    """Busca o código interno da planilha a partir da Razão Social do SIGA."""
    siga_norm = normalize_name(razao_social)
    if siga_norm in mapping:
        return mapping[siga_norm]["codigo"], mapping[siga_norm]["nome_original"]
        
    for excel_norm, info in mapping.items():
        if len(siga_norm) > 4 and len(excel_norm) > 4:
            if siga_norm in excel_norm or excel_norm in siga_norm:
                return info["codigo"], info["nome_original"]
                
    return "", ""

# ====================================================
# =============== Função de download =================
# ====================================================

def processar_lote_xml(lote, api_token, thread_id):
    """Processa um lote de chaves fazendo o download dos XMLs e limpando namespaces."""
    total = len(lote)
    count = 0
    
    session = requests.Session()
    
    for task in lote:
        chave, company_label, date_subfolder = task
        folder = os.path.join("siga_nfce", f"{company_label}-", date_subfolder)
        filepath = os.path.join(folder, f"{chave}.xml")
        
        # Ignora se já foi baixado anteriormente
        if os.path.exists(filepath):
            count += 1
            continue
            
        time.sleep(0.5)  # Atraso de cortesia exigido pela API
        
        try:
            # 1. Obter o ID da NFC-e
            link_id = f'https://cfe.sefaz.ce.gov.br:8443/portalcfews/nfce/coupons/extract/{str(chave).replace(" ", "")}?apiKey={api_token}'
            r_id = session.get(link_id, verify=False, timeout=20)
            
            if r_id.status_code != 200:
                print(f"❌ [Thread {thread_id}] Erro ao buscar ID da chave {chave} (HTTP {r_id.status_code})")
                continue
                
            res_json = r_id.json()
            id_nfe = res_json.get('idNfe')
            if not id_nfe:
                print(f"❌ [Thread {thread_id}] ID NFe não encontrado no retorno da chave {chave}")
                continue
                
            # 2. Baixar o XML fiscal
            link_xml = f'https://cfe.sefaz.ce.gov.br:8443/portalcfews/nfce/fiscal-coupons/xml/{id_nfe}?chaveAcesso={str(chave).replace(" ", "")}&apiKey={api_token}'
            r_xml = session.get(link_xml, verify=False, timeout=20)
            
            if r_xml.status_code != 200:
                print(f"❌ [Thread {thread_id}] Erro ao baixar XML da chave {chave} (HTTP {r_xml.status_code})")
                continue
                
            # 3. Limpar namespaces do XML (mantendo compatibilidade com seu código original)
            xml_data = r_xml.content.decode('utf-8')
            xml_data = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', xml_data)  # remove declarações xmlns
            xml_data = re.sub(r'(<\/?)(\w+:)', r'\1', xml_data)  # remove prefixos de namespace
            
            root = ET.fromstring(xml_data)
            tree = ET.ElementTree(root)
            
            # Criar pasta da empresa e gravar arquivo
            os.makedirs(folder, exist_ok=True)
            tree.write(filepath, encoding='utf-8', xml_declaration=True)
            r_xml.close()
            
            count += 1
            percent = (count / total) * 100
            print(f"[Thread {thread_id}] {percent:.2f}% concluído - Faltam {total - count}")
            
        except Exception as e:
            print(f"❌ [Thread {thread_id}] Erro inesperado na chave {chave}: {e}")

# ====================================================
# ============== Código principal ====================
# ====================================================

def main():
    print("=" * 60)
    print("     SIGA SEFAZ-CE - COLETOR COMPLETO DE XMLS (NFC-E)")
    print("=" * 60)
    
    # 1. Validar e carregar base de empresas salvas do SIGA
    if not os.path.exists(COMPANIES_CSV):
        print(f"❌ Base de empresas '{COMPANIES_CSV}' não encontrada.")
        print("Por favor, execute o script 'siga_extractor.py' primeiro para coletar os CNPJs.")
        sys.exit(1)
        
    siga_companies = []
    with open(COMPANIES_CSV, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            siga_companies.append(row)
            
    print(f"✅ Carregadas {len(siga_companies)} empresas do arquivo '{COMPANIES_CSV}'.")
    
    # 2. Carregar de-para de códigos da planilha Excel
    excel_mapping = load_excel_mapping(EXCEL_FILE)
    
    # 3. Carregar API Token da SEFAZ CFe
    api_token = load_api_token()
    if not api_token:
        print("\nChave da API SEFAZ CFe ausente.")
        api_token = input("Cole a sua apiKey da SEFAZ Ceará: ").strip()
        save_api_token(api_token)
    else:
        print("ℹ️ Token da API SEFAZ carregado com sucesso do api_token.txt.")
        
    # 4. Carregar configurações de autenticação do SIGA
    config = load_config()
    access_token = config.get("access_token")
    refresh_token = config.get("refresh_token")
    client_id = config.get("client_id") or "painelind-frontend"
    token_url = config.get("token_url") or "https://sso.sefaz.ce.gov.br/auth/realms/sefaz-ad-realm/protocol/openid-connect/token"
    
    # Se os tokens não existem, ou se o refresh_token já expirou, solicita novos tokens
    if not access_token or not refresh_token or is_token_expired(refresh_token):
        if not access_token or not refresh_token:
            print("❌ Credenciais de acesso do SIGA ausentes no 'config.json'.")
        else:
            print("\n⚠️ O seu Refresh Token do SIGA expirou ou é inválido.")
            
        print("Por favor, forneça novos tokens obtidos na aba Network (F12) do SIGA.")
        access_token, refresh_token = ensure_valid_token(config, "", "", token_url, client_id)
        
    # 5. Coletar período desejado pelo usuário
    try:
        ano = int(input("\nDigite o ano de referência (ex: 2026): ").strip())
        mes_num = int(input("Digite o número do mês de referência (1 a 12): ").strip())
        if mes_num < 1 or mes_num > 12:
            raise ValueError()
    except ValueError:
        print("❌ Ano ou mês inválido. Saindo.")
        sys.exit(1)
        
    ref_date = f"{ano:04d}-{mes_num:02d}-01"
    mes_str = f"{mes_num:02d}"
    date_subfolder = f"{mes_str}{ano}"
    
    print(f"\n📅 Período de consulta: Mês {target_month_name if 'target_month_name' in locals() else mes_str}/{ano}")
    
    # 5.1. Perguntar se deseja processar empresas específicas
    opcao_empresa = input("\nDeseja puxar dados de alguma empresa específica? (s/N): ").strip().lower()
    if opcao_empresa in ['s', 'sim']:
        cnpjs_input = input("Digite ou cole os CNPJs das empresas: ").strip()
        
        # Extrai CNPJs usando regex para suportar formatos colados, separados por vírgula, espaços, etc.
        cnpjs_encontrados = re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{14}|\d{8}', cnpjs_input)
        cnpjs_filtrar = []
        for c in cnpjs_encontrados:
            c_clean = "".join(filter(str.isdigit, c))
            if c_clean:
                cnpjs_filtrar.append(c_clean)
        
        if cnpjs_filtrar:
            filtradas = []
            for comp in siga_companies:
                cnpj_comp = "".join(filter(str.isdigit, comp.get("cnpj", "")))
                # Compara se o CNPJ filtrado é um subconjunto ou igual ao CNPJ da empresa
                if any(f in cnpj_comp for f in cnpjs_filtrar):
                    filtradas.append(comp)
            
            if filtradas:
                siga_companies = filtradas
                print(f"🎯 Filtrado! {len(siga_companies)} empresa(s) selecionada(s) para processar.")
            else:
                print("⚠️ Nenhuma empresa correspondente encontrada com os CNPJs fornecidos. Usando a lista completa.")
        else:
            print("⚠️ Nenhum CNPJ válido identificado no texto digitado. Usando a lista completa.")
            
    # 6. Loop de extração de chaves do SIGA
    xml_tasks = []
    session = requests.Session()
    
    print(f"\n🔎 Buscando chaves de acesso no painel do SIGA...")
    
    for idx, comp in enumerate(siga_companies):
        cnpj = comp.get("cnpj", "").strip()
        razao_social = comp.get("razaoSocial", "").strip()
        cgf = comp.get("cgf", "").strip()
        cnpj_clean = "".join(filter(str.isdigit, cnpj))
        
        if not cnpj_clean:
            continue
            
        cod_excel, nome_excel = find_company_code(razao_social, excel_mapping)
        company_label = f"{cod_excel} - {nome_excel}" if cod_excel else f"{cnpj_clean} - {razao_social}"
        company_label = sanitize_filename(company_label)
        
        print(f"⏳ [{idx+1}/{len(siga_companies)}] Coletando chaves: {company_label[:35]}... ", end="", flush=True)
        
        company_nfc_keys = []
        page = 0
        size = 1000
        has_more = True
        sem_procuracao_detected = False
        
        while has_more:
            # Renovar token SIGA se necessário
            access_token, refresh_token = ensure_valid_token(config, access_token, refresh_token, token_url, client_id)
                    
            # Chamada da API detalhada de NFC-e do SIGA
            api_url = f"https://siga.sefaz.ce.gov.br/api/v1/unidades/{cnpj_clean}/documentos-fiscais/nfc-e"
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Accept': 'application/json, text/plain, */*',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Origin': 'https://siga.sefaz.ce.gov.br',
                'Referer': 'https://siga.sefaz.ce.gov.br/ui/selecao-contribuinte/contribuinte'
            }
            params = {
                'size': size,
                'page': page,
                'sort': 'datEmissao,asc',
                'dat-referencia': ref_date,
                'tipo-operacao': 'SAIDA'
            }
            
            try:
                response = session.get(api_url, headers=headers, params=params, timeout=25)
                
                if response.status_code == 401:
                    print("\n🔄 Erro 401 (Não Autorizado). Forçando renovação emergencial...")
                    access_token, refresh_token = ensure_valid_token(config, "", refresh_token, token_url, client_id)
                    continue
                
                # Se der 403, significa que não temos procuração de acesso para esta empresa
                if response.status_code == 403:
                    sem_procuracao_detected = True
                    break
                        
                if response.status_code != 200:
                    break
                    
                data = response.json()
                content = []
                if isinstance(data, list):
                    content = data
                elif isinstance(data, dict):
                    if "content" in data and isinstance(data["content"], list):
                        content = data["content"]
                    elif "data" in data and isinstance(data["data"], list):
                        content = data["data"]
                        
                if not content:
                    has_more = False
                else:
                    for item in content:
                        chave = item.get("chaveAcesso") or item.get("chave")
                        if chave:
                            company_nfc_keys.append(chave)
                    
                    if len(content) < size:
                        has_more = False
                    else:
                        page += 1
                        
                time.sleep(0.1)
                
            except Exception:
                break
                
        if sem_procuracao_detected:
            print("❌ Sem procuração! (Adicionado na pasta 'Sem Procuracao')")
            # Salvar o aviso numa pasta diferente
            folder_sem = os.path.join("siga_nfce", "Sem Procuracao")
            os.makedirs(folder_sem, exist_ok=True)
            filepath_sem = os.path.join(folder_sem, f"{company_label}.txt")
            with open(filepath_sem, "w", encoding="utf-8") as f_sem:
                f_sem.write(f"Código/Nome: {company_label}\nCNPJ: {cnpj}\nCGF: {cgf}\nRazão Social: {razao_social}\nStatus: Sem procuração de acesso cadastrada na SEFAZ para este usuário.")
            continue
            
        print(f"OK! ({len(company_nfc_keys)} chaves)")
        
        # Adicionar as chaves na lista de downloads
        for k in company_nfc_keys:
            xml_tasks.append((k, company_label, date_subfolder))
            
    # 7. Filtro de XMLs faltantes
    baixados = []
    print("\n🔍 Escaneando diretórios locais para verificar arquivos já baixados...")
    if os.path.exists("siga_nfce"):
        for root, dirs, files in os.walk("siga_nfce"):
            for name in files:
                if name.endswith(".xml"):
                    baixados.append(name.replace(".xml", ""))
                
    print(f"ℹ️ Total de XMLs locais identificados: {len(baixados)}")
    
    # Filtrar chaves pendentes
    tarefas_faltando = [t for t in xml_tasks if t[0] not in baixados]
    print(f"ℹ️ Total de XMLs pendentes de download: {len(tarefas_faltando)}")
    
    if not tarefas_faltando:
        print("\n🎉 Excelente! Todos os XMLs de todas as empresas para este período já estão baixados.")
        return
        
    # 8. Download multithread de XMLs (divisão em 40 threads)
    num_threads = min(40, len(tarefas_faltando))
    partes = [tarefas_faltando[i::num_threads] for i in range(num_threads)]
    
    print(f"\n⚡ Iniciando download de {len(tarefas_faltando)} XMLs divididos em {num_threads} threads...")
    
    threads = []
    for idx, parte in enumerate(partes):
        if parte:
            t = threading.Thread(target=processar_lote_xml, args=(parte, api_token, idx+1))
            t.start()
            threads.append(t)
            
    for t in threads:
        t.join()
        
    print("\n🎉 Download de XMLs concluído!")

if __name__ == "__main__":
    main()
