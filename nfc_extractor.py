#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SIGA SEFAZ-CE - Extrator Completo de NFC-e e Compilador de XMLs (Threads e ZIP)

Este script unifica o fluxo:
1. Carrega histórico de chaves já cadastradas no banco a partir de 'chaves_ja_cadastradas.csv' (somente leitura).
2. Coleta as chaves de acesso detalhadas das NFC-e no SIGA para o mês/ano selecionado.
3. Filtra documentos cancelados conforme configuração individual de cada empresa.
4. Descarta chaves que já constem no CSV de chaves cadastradas para aquele CNPJ.
5. Associa com códigos internos da planilha 'empresas cod.xlsx'.
6. Realiza o download multithread (40 threads) dos XMLs fiscais pendentes da SEFAZ em lotes de até 4.000.
7. Limpa namespaces dos XMLs e organiza em pastas 'EMPRESA-/MMAAAA/NFCe_Parte_X/'.
8. Registra as chaves processadas nesta execução em 'chaves_baixadas_nesta_execucao.csv'.
"""

import os
import sys
import json
import time
import base64
import csv
import re
import threading
import xml.etree.ElementTree as ET
import requests
import urllib3

# Garante compatibilidade de caracteres especiais/emojis no console Windows
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

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
CSV_JA_CADASTRADAS = "chaves_ja_cadastradas.csv"
CSV_EXECUCAO = "chaves_baixadas_nesta_execucao.csv"

def load_config():
    """Carrega as credenciais e configurações em cache."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️ Erro ao carregar config.json: {e}")
    return {}

def save_config(config):
    """Grava as credenciais e configurações em cache."""
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
    if ":" in val:
        parts = val.split(":", 1)
        val = parts[1].strip()
    val = val.strip(" \t\n\r\"',;")
    return val

def clean_cnpj(cnpj):
    """Remove formatação do CNPJ e retorna apenas dígitos numéricos (normalizado para 14 dígitos)."""
    if not cnpj:
        return ""
    digits = "".join(filter(str.isdigit, str(cnpj)))
    if len(digits) <= 14 and len(digits) >= 8:
        return digits.zfill(14)
    return digits

def format_cnpj(cnpj):
    """Formata CNPJ no padrão 00.000.000/0000-00."""
    c = clean_cnpj(cnpj)
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    return str(cnpj)

def clean_chave(chave):
    """Remove qualquer caractere não numérico e retorna a chave com 44 dígitos."""
    if not chave:
        return ""
    return "".join(filter(str.isdigit, str(chave)))

def format_num(n):
    """Formata número com separador de milhar brasileiro (ponto)."""
    return f"{n:,}".replace(",", ".")

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

def escanear_xmls_locais(pasta_base="siga_nfce"):
    """Varre a pasta local procurando XMLs já salvos no disco."""
    chaves_locais = {}
    if not os.path.exists(pasta_base):
        return chaves_locais
    for root, _, files in os.walk(pasta_base):
        for name in files:
            if name.lower().endswith(".xml"):
                chave = clean_chave(name[:-4])
                if len(chave) == 44:
                    cnpj = clean_cnpj(chave[6:20])
                    if cnpj not in chaves_locais:
                        chaves_locais[cnpj] = set()
                    chaves_locais[cnpj].add(chave)
    return chaves_locais

def carregar_chaves_ja_cadastradas(filepath=CSV_JA_CADASTRADAS):
    """
    Carrega o arquivo CSV de chaves já cadastradas no banco de dados.
    Este arquivo é utilizado EXCLUSIVAMENTE em modo de LEITURA e NUNCA é modificado pelo código.
    
    Retorna uma estrutura de indexação em memória: dict[CNPJ_limpo, set[CHAVE_limpa]].
    Suporta:
      - Arquivos com colunas (CNPJ, CHAVE) ou (CHAVE, CNPJ)
      - Qualquer delimitador (vírgula, ponto-e-vírgula, tabulação, barra vertical)
      - CNPJs com ou sem pontuação (normalizados para 14 dígitos numéricos)
      - Chaves com ou sem formatação (normalizadas para 44 dígitos numéricos)
      - Linhas contendo apenas a CHAVE (extrai o CNPJ dos dígitos 6 a 19 da chave de 44 dígitos)
    """
    chaves_por_cnpj = {}
    total_carregadas = 0
    
    # Se o arquivo principal não existir mas o legado existir, usa como fallback
    target_file = filepath
    if not os.path.exists(target_file) and os.path.exists("chaves_ja_baixadas.csv"):
        target_file = "chaves_ja_baixadas.csv"
        
    if os.path.exists(target_file) and os.path.getsize(target_file) > 0:
        print(f"📖 Carregando chaves cadastradas de '{target_file}' (somente leitura)...")
        try:
            with open(target_file, mode="r", encoding="utf-8-sig") as f:
                for line_num, line in enumerate(f, 1):
                    raw_line = line.strip()
                    if not raw_line:
                        continue
                    # Se for linha de cabeçalho
                    if line_num <= 3 and ("cnpj" in raw_line.lower() or "chave" in raw_line.lower()):
                        digitos = "".join(filter(str.isdigit, raw_line))
                        if len(digitos) < 20:
                            continue
                            
                    partes = re.split(r'[,;\t|\s]+', raw_line)
                    c_cnpj = ""
                    c_chave = ""
                    
                    for p in partes:
                        digits = "".join(filter(str.isdigit, p))
                        if len(digits) == 44 and not c_chave:
                            c_chave = digits
                        elif (8 <= len(digits) <= 14) and not c_cnpj:
                            c_cnpj = clean_cnpj(digits)
                            
                    if c_chave and len(c_chave) == 44:
                        if not c_cnpj:
                            c_cnpj = clean_cnpj(c_chave[6:20])
                            
                        if c_cnpj not in chaves_por_cnpj:
                            chaves_por_cnpj[c_cnpj] = set()
                        chaves_por_cnpj[c_cnpj].add(c_chave)
                        total_carregadas += 1
                        
            print(f"✅ Chaves carregadas do CSV: {format_num(total_carregadas)} chaves em {len(chaves_por_cnpj)} CNPJ(s).\n")
        except Exception as e:
            print(f"⚠️ Erro ao ler '{target_file}': {e}\n")
    else:
        print(f"ℹ️ Arquivo '{filepath}' não encontrado ou vazio. Nenhuma chave pré-cadastrada carregada.\n")
        
    return chaves_por_cnpj

def registrar_execucao(registros, filepath=CSV_EXECUCAO):
    """
    Registra as chaves processadas nesta execução em arquivo CSV para auditoria.
    registros: lista de tuplas (cnpj, chave, status)
    """
    deve_escrever_cabecalho = not os.path.exists(filepath) or os.path.getsize(filepath) == 0
    try:
        with open(filepath, mode="a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            if deve_escrever_cabecalho:
                writer.writerow(["CNPJ", "CHAVE", "STATUS", "DATA_HORA"])
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            for cnpj, chave, status in registros:
                writer.writerow([clean_cnpj(cnpj), clean_chave(chave), status, timestamp])
    except Exception as e:
        print(f"⚠️ Erro ao gravar em '{filepath}': {e}")

# ====================================================
# =============== Função de download =================
# ====================================================

def processar_lote_xml(lote, api_token, thread_id, stats_lock=None, stats=None, registros_lock=None, registros_execucao=None, cnpj_clean=""):
    """Processa um lote de chaves fazendo o download dos XMLs e limpando namespaces."""
    total = len(lote)
    count = 0
    
    session = requests.Session()
    
    for task in lote:
        chave, folder = task
        filepath = os.path.join(folder, f"{chave}.xml")
        
        # Ignora se já foi baixado anteriormente
        if os.path.exists(filepath):
            count += 1
            if stats_lock and stats is not None:
                with stats_lock:
                    stats["baixados"] += 1
            if registros_lock and registros_execucao is not None:
                with registros_lock:
                    registros_execucao.append((cnpj_clean, chave, "SUCESSO"))
            continue
            
        time.sleep(0.5)  # Atraso de cortesia exigido pela API
        
        try:
            # 1. Obter o ID da NFC-e
            clean_chave_str = clean_chave(chave)
            link_id = f'https://cfe.sefaz.ce.gov.br:8443/portalcfews/nfce/coupons/extract/{clean_chave_str}?apiKey={api_token}'
            r_id = session.get(link_id, verify=False, timeout=20)
            
            if r_id.status_code != 200:
                print(f"❌ [Thread {thread_id}] Erro ao buscar ID da chave {chave} (HTTP {r_id.status_code})")
                if stats_lock and stats is not None:
                    with stats_lock:
                        stats["falhas"] += 1
                if registros_lock and registros_execucao is not None:
                    with registros_lock:
                        registros_execucao.append((cnpj_clean, chave, "ERRO"))
                continue
                
            res_json = r_id.json()
            id_nfe = res_json.get('idNfe')
            if not id_nfe:
                print(f"❌ [Thread {thread_id}] ID NFe não encontrado no retorno da chave {chave}")
                if stats_lock and stats is not None:
                    with stats_lock:
                        stats["falhas"] += 1
                if registros_lock and registros_execucao is not None:
                    with registros_lock:
                        registros_execucao.append((cnpj_clean, chave, "ERRO"))
                continue
                
            # 2. Baixar o XML fiscal
            link_xml = f'https://cfe.sefaz.ce.gov.br:8443/portalcfews/nfce/fiscal-coupons/xml/{id_nfe}?chaveAcesso={clean_chave_str}&apiKey={api_token}'
            r_xml = session.get(link_xml, verify=False, timeout=20)
            
            if r_xml.status_code != 200:
                print(f"❌ [Thread {thread_id}] Erro ao baixar XML da chave {chave} (HTTP {r_xml.status_code})")
                if stats_lock and stats is not None:
                    with stats_lock:
                        stats["falhas"] += 1
                if registros_lock and registros_execucao is not None:
                    with registros_lock:
                        registros_execucao.append((cnpj_clean, chave, "ERRO"))
                continue
                
            # 3. Limpar namespaces do XML
            xml_data = r_xml.content.decode('utf-8')
            xml_data = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', xml_data)
            xml_data = re.sub(r'(<\/?)(\w+:)', r'\1', xml_data)
            
            root = ET.fromstring(xml_data)
            tree = ET.ElementTree(root)
            
            # Criar pasta da empresa e gravar arquivo
            os.makedirs(folder, exist_ok=True)
            tree.write(filepath, encoding='utf-8', xml_declaration=True)
            r_xml.close()
            
            count += 1
            if stats_lock and stats is not None:
                with stats_lock:
                    stats["baixados"] += 1
            if registros_lock and registros_execucao is not None:
                with registros_lock:
                    registros_execucao.append((cnpj_clean, chave, "SUCESSO"))
            percent = (count / total) * 100
            print(f"[Thread {thread_id}] {percent:.2f}% concluído - Faltam {total - count}")
            
        except Exception as e:
            print(f"❌ [Thread {thread_id}] Erro inesperado na chave {chave}: {e}")
            if stats_lock and stats is not None:
                with stats_lock:
                    stats["falhas"] += 1
            if registros_lock and registros_execucao is not None:
                with registros_lock:
                    registros_execucao.append((cnpj_clean, chave, "ERRO"))

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
            
    print(f" Carregadas {len(siga_companies)} empresas do arquivo '{COMPANIES_CSV}'.")
    
    # 2. Carregar de-para de códigos da planilha Excel
    excel_mapping = load_excel_mapping(EXCEL_FILE)
    
    # 3. Carregar API Token da SEFAZ CFe
    api_token = load_api_token()
    if not api_token:
        print("\nChave da API SEFAZ CFe ausente.")
        api_token = clean_input_val(input("Cole a sua apiKey da SEFAZ Ceará: "))
        save_api_token(api_token)
    else:
        print("ℹ️ Token da API SEFAZ carregado com sucesso do api_token.txt.")
        
    # 4. Carregar configurações de autenticação do SIGA
    config = load_config()
    access_token = config.get("access_token")
    refresh_token = config.get("refresh_token")
    client_id = config.get("client_id") or "painelind-frontend"
    token_url = config.get("token_url") or "https://sso.sefaz.ce.gov.br/auth/realms/sefaz-ad-realm/protocol/openid-connect/token"
    
    # Pergunta se deseja trocar os tokens do SIGA
    if access_token and refresh_token and not is_token_expired(refresh_token):
        trocar_siga = input("Deseja trocar os tokens de login/F12 do SIGA (access_token/refresh_token)? (s/N): ").strip().lower()
        if trocar_siga in ['s', 'sim']:
            print("\n⚠️ Atenção: Os tokens do SIGA são códigos longos (JWT) obtidos na aba Network (F12), NÃO são CNPJs.")
            novo_acc = clean_input_val(input("Cole o novo access_token (longo): "))
            if len(novo_acc) < 50:
                print("⚠️ O access_token colado é muito curto ou inválido. Mantendo o token atual.")
            else:
                novo_ref = clean_input_val(input("Cole o novo refresh_token (longo): "))
                if novo_acc and novo_ref:
                    access_token = novo_acc
                    refresh_token = novo_ref
                    config.update({"access_token": access_token, "refresh_token": refresh_token})
                    save_config(config)
                    print("✅ Tokens do SIGA atualizados com sucesso no config.json!")
    
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
    
    print(f"\n📅 Período de consulta: Mês {mes_str}/{ano}")
    
    # 6. Carregar histórico de chaves já existentes no banco (CSV - Somente Leitura)
    chaves_por_cnpj = carregar_chaves_ja_cadastradas(CSV_JA_CADASTRADAS)
    
    # 7. Selecionar empresas e configurar tratamento de canceladas
    mapa_siga = {}
    for comp in siga_companies:
        c_clean = clean_cnpj(comp.get("cnpj", ""))
        if c_clean:
            mapa_siga[c_clean] = comp
            
    empresas_alvo = []
    
    # Verifica se há lista pré-configurada no config.json
    config_empresas = config.get("empresas_nfc", [])
    if config_empresas:
        print(f"\nℹ️ Encontradas {len(config_empresas)} empresas configuradas no config.json:")
        for idx_e, emp_cfg in enumerate(config_empresas, 1):
            emp_cnpj = clean_cnpj(emp_cfg.get("cnpj", ""))
            emp_canc = emp_cfg.get("processar_canceladas", False)
            canc_txt = "SIM" if emp_canc else "NÃO"
            print(f"  {idx_e}. CNPJ: {format_cnpj(emp_cnpj)} (Processar canceladas: {canc_txt})")
            
        usar_cfg = input("\nDeseja processar as empresas configuradas no config.json? (S/n): ").strip().lower()
        if usar_cfg not in ['n', 'nao', 'não']:
            for emp_cfg in config_empresas:
                c_clean = clean_cnpj(emp_cfg.get("cnpj", ""))
                if c_clean:
                    emp_info = mapa_siga.get(c_clean, {"cnpj": c_clean, "razaoSocial": "", "cgf": ""})
                    empresas_alvo.append({
                        "cnpj": c_clean,
                        "razaoSocial": emp_info.get("razaoSocial", ""),
                        "cgf": emp_info.get("cgf", ""),
                        "processar_canceladas": emp_cfg.get("processar_canceladas", False)
                    })
                    
    if not empresas_alvo:
        opcao_empresa = input("\nDeseja puxar dados de alguma empresa específica? (s/N): ").strip().lower()
        if opcao_empresa in ['s', 'sim']:
            cnpjs_input = input("Digite ou cole os CNPJs das empresas: ").strip()
            cnpjs_encontrados = re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{14}|\d{8}', cnpjs_input)
            cnpjs_filtrar = []
            for c in cnpjs_encontrados:
                c_clean = clean_cnpj(c)
                if c_clean and c_clean not in cnpjs_filtrar:
                    cnpjs_filtrar.append(c_clean)
            
            if cnpjs_filtrar:
                for c_clean in cnpjs_filtrar:
                    emp_info = mapa_siga.get(c_clean, {"cnpj": c_clean, "razaoSocial": "", "cgf": ""})
                    # Pergunta ou usa False por padrão para canceladas
                    empresas_alvo.append({
                        "cnpj": c_clean,
                        "razaoSocial": emp_info.get("razaoSocial", ""),
                        "cgf": emp_info.get("cgf", ""),
                        "processar_canceladas": False
                    })
                print(f"🎯 Total: {len(empresas_alvo)} empresa(s) selecionada(s) para processar.")
            else:
                print("⚠️ Nenhum CNPJ válido identificado. Usando todas as empresas salvas.")
                
        if not empresas_alvo:
            for comp in siga_companies:
                c_clean = clean_cnpj(comp.get("cnpj", ""))
                if c_clean:
                    empresas_alvo.append({
                        "cnpj": c_clean,
                        "razaoSocial": comp.get("razaoSocial", ""),
                        "cgf": comp.get("cgf", ""),
                        "processar_canceladas": False
                    })
                    
    # 8. Execução e download por CNPJ
    session = requests.Session()
    resumo_empresas = []
    
    for emp in empresas_alvo:
        cnpj_clean = emp["cnpj"]
        razao_social = emp.get("razaoSocial", "")
        cgf = emp.get("cgf", "")
        processar_canceladas = emp.get("processar_canceladas", False)
        
        cod_excel, nome_excel = find_company_code(razao_social, excel_mapping) if razao_social else ("", "")
        if cod_excel:
            company_label = f"{cod_excel} - {nome_excel}"
        elif razao_social:
            company_label = f"{cnpj_clean} - {razao_social}"
        else:
            company_label = f"{cnpj_clean}"
        company_label = sanitize_filename(company_label)
        
        print("\n" + "=" * 50)
        print(f"PROCESSANDO: {format_cnpj(cnpj_clean)}")
        if company_label:
            print(f"Empresa: {company_label}")
        if processar_canceladas:
            print("Processar canceladas: SIM")
        print("=" * 50)
        
        print(f"\n🔍 Consultando listagem de notas no SIGA ({mes_str}/{ano})...")
        
        company_nfc_keys = []
        page = 0
        size = 1000
        has_more = True
        sem_procuracao_detected = False
        
        while has_more:
            # Renovar token SIGA se necessário
            access_token, refresh_token = ensure_valid_token(config, access_token, refresh_token, token_url, client_id)
                    
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
                    print("\n⚠️ Erro 401 (Não Autorizado). Forçando renovação emergencial...")
                    access_token, refresh_token = ensure_valid_token(config, "", refresh_token, token_url, client_id)
                    continue
                
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
                        situacao = str(item.get("situacao") or item.get("status") or "").upper()
                        is_cancelada = "CANCELAD" in situacao or item.get("cancelada") is True
                        
                        if not processar_canceladas and is_cancelada:
                            continue
                            
                        if chave:
                            c_chave = clean_chave(chave)
                            if len(c_chave) == 44:
                                company_nfc_keys.append(c_chave)
                    
                    print(f"  → Página {page + 1}: {len(company_nfc_keys)} chaves localizadas no SIGA...", end="\r")
                    
                    if len(content) < size:
                        has_more = False
                    else:
                        page += 1
                        
                time.sleep(0.1)
                
            except Exception:
                break
                
        if sem_procuracao_detected:
            print("❌ Sem procuração de acesso! (Registrado em 'Sem Procuracao')")
            folder_sem = os.path.join("siga_nfce", "Sem Procuracao")
            os.makedirs(folder_sem, exist_ok=True)
            filepath_sem = os.path.join(folder_sem, f"{company_label}.txt")
            with open(filepath_sem, "w", encoding="utf-8") as f_sem:
                f_sem.write(f"Código/Nome: {company_label}\nCNPJ: {cnpj_clean}\nCGF: {cgf}\nRazão Social: {razao_social}\nStatus: Sem procuração de acesso cadastrada na SEFAZ para este usuário.")
            resumo_empresas.append({
                "cnpj": cnpj_clean,
                "siga": 0,
                "existentes": 0,
                "pendentes": 0,
                "baixadas": 0,
                "falhas": 0,
                "status": "Sem Procuração"
            })
            continue
            
        # Deduplicação mantendo ordem original
        chaves_siga = list(dict.fromkeys(company_nfc_keys))
        duplicadas_removidas = len(company_nfc_keys) - len(chaves_siga)
        
        print(f"\nChaves encontradas na API: {format_num(len(chaves_siga))}")
        if duplicadas_removidas > 0:
            print(f"   (Chaves duplicadas na API ignoradas: {format_num(duplicadas_removidas)})")
        
        # Comparação com chaves_ja_cadastradas.csv
        print(f"\nComparando com {CSV_JA_CADASTRADAS}...")
        
        # Busca no arquivo SOMENTE as chaves relacionadas a esse CNPJ específico
        chaves_cadastradas_cnpj = chaves_por_cnpj.get(cnpj_clean, set())
        
        # Identifica chaves cadastradas e pendentes para este CNPJ
        chaves_ignoradas = [k for k in chaves_siga if k in chaves_cadastradas_cnpj]
        chaves_pendentes = [k for k in chaves_siga if k not in chaves_cadastradas_cnpj]
        
        print(f"\nChaves já cadastradas: {format_num(len(chaves_ignoradas))}")
        print(f"Chaves ignoradas: {format_num(len(chaves_ignoradas))}")
        print(f"Chaves pendentes: {format_num(len(chaves_pendentes))}")
        
        if not chaves_pendentes:
            print("\nTodas as chaves desse CNPJ já estão cadastradas.")
            print("Nenhum download necessário.")
            resumo_empresas.append({
                "cnpj": cnpj_clean,
                "siga": len(chaves_siga),
                "existentes": len(chaves_ignoradas),
                "pendentes": 0,
                "baixadas": 0,
                "falhas": 0,
                "status": "OK"
            })
            continue
            
        print(f"\nBaixando somente as {format_num(len(chaves_pendentes))} chaves pendentes...")
        
        # Divisão em lotes de até 4.000
        TAMANHO_LOTE = 4000
        lotes = [chaves_pendentes[i:i + TAMANHO_LOTE] for i in range(0, len(chaves_pendentes), TAMANHO_LOTE)]
        total_lotes = len(lotes)
        
        print(f"\nTotal de lotes: {total_lotes}")
        for l_i, l_c in enumerate(lotes, 1):
            print(f"Parte {l_i} → {format_num(len(l_c))}")
                
        baixados_empresa = 0
        falhas_empresa = 0
        
        for lote_idx, chaves_lote in enumerate(lotes, 1):
            pasta_nome = f"NFCe_Parte_{lote_idx}"
            folder = os.path.join("siga_nfce", f"{company_label}-", date_subfolder, pasta_nome)
            
            print(f"\nLote {lote_idx}/{total_lotes}")
            print(f"Chaves: {format_num(len(chaves_lote))}")
            print(f"Pasta: {pasta_nome}\n")
            print("Iniciando downloads...")
            
            tasks_lote = [(chave, folder) for chave in chaves_lote]
            num_threads = min(40, len(tasks_lote))
            partes = [tasks_lote[i::num_threads] for i in range(num_threads)]
            
            stats_lock = threading.Lock()
            stats = {"baixados": 0, "falhas": 0}
            registros_lock = threading.Lock()
            registros_execucao = []
            
            threads = []
            for idx, parte in enumerate(partes):
                if parte:
                    t = threading.Thread(
                        target=processar_lote_xml,
                        args=(parte, api_token, idx + 1, stats_lock, stats, registros_lock, registros_execucao, cnpj_clean)
                    )
                    t.start()
                    threads.append(t)
                    
            for t in threads:
                t.join()
                
            baixados_empresa += stats["baixados"]
            falhas_empresa += stats["falhas"]
            registrar_execucao(registros_execucao)
            
            print(f"\nLote {lote_idx}/{total_lotes} concluído: {format_num(stats['baixados'])}/{format_num(len(chaves_lote))}")
            
        resumo_empresas.append({
            "cnpj": cnpj_clean,
            "siga": len(chaves_siga),
            "existentes": len(chaves_ignoradas),
            "pendentes": len(chaves_pendentes),
            "baixadas": baixados_empresa,
            "falhas": falhas_empresa,
            "status": "OK"
        })
        
    # 9. Exibição do Resumo Final
    print("\n" + "=" * 50)
    print("RESUMO FINAL")
    print("=" * 50)
    
    total_siga_geral = sum(r["siga"] for r in resumo_empresas)
    total_existentes_geral = sum(r["existentes"] for r in resumo_empresas)
    total_pendentes_geral = sum(r["pendentes"] for r in resumo_empresas)
    total_baixadas_geral = sum(r["baixadas"] for r in resumo_empresas)
    total_falhas_geral = sum(r["falhas"] for r in resumo_empresas)
    
    for idx, r in enumerate(resumo_empresas):
        print(f"\nCNPJ: {format_cnpj(r['cnpj'])}")
        if r.get("status") == "Sem Procuração":
            print("Status: Sem procuração de acesso cadastrada na SEFAZ.")
        else:
            print(f"SIGA: {format_num(r['siga'])}")
            print(f"Já existentes: {format_num(r['existentes'])}")
            print(f"Pendentes: {format_num(r['pendentes'])}")
            print(f"Baixadas: {format_num(r['baixadas'])}")
            if r["falhas"] > 0 or r["pendentes"] > 0:
                print(f"Falhas: {format_num(r['falhas'])}")
        if idx < len(resumo_empresas) - 1:
            print("\n" + "-" * 40)
            
    print("\n" + "=" * 50)
    print("\nTOTAL:\n")
    print(f"Chaves encontradas no SIGA: {format_num(total_siga_geral)}")
    print(f"Já existentes no banco/CSV: {format_num(total_existentes_geral)}")
    print(f"Novas identificadas: {format_num(total_pendentes_geral)}")
    print(f"Baixadas com sucesso: {format_num(total_baixadas_geral)}")
    print(f"Falhas: {format_num(total_falhas_geral)}")
    print("\n" + "=" * 50)

if __name__ == "__main__":
    main()
