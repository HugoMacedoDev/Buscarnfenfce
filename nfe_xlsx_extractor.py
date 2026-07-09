#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SIGA SEFAZ-CE - Extrator Automático de Tabelas NF-e (XLSX)
Este script automatiza as solicitações e downloads de planilhas XLSX contendo
todas as chaves de NF-e (Emissor/Saídas e Destinatário/Entradas) no SIGA.
"""

import os
import sys

# Garante compatibilidade de caracteres especiais/emojis no console Windows
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import json
import time
import base64
import csv
import re
import requests
import urllib3

# Desativa avisos de certificado (caso utilize verify=False em algum momento)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    import openpyxl
except ImportError:
    print("\n❌ A biblioteca 'openpyxl' é necessária para ler a planilha Excel.")
    print("Execute o seguinte comando no terminal para instalá-la:")
    print("pip install openpyxl")
    sys.exit(1)

CONFIG_FILE = "config.json"
EXCEL_FILE = "empresas cod.xlsx"
COMPANIES_CSV = "siga_empresas.csv"

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

def is_token_expired(token):
    """Verifica se o token JWT de acesso está expirado."""
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
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0'
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
        print(f"\n❌ Erro de conexão ao renovar token: {e}")
        return None, None

def ensure_valid_token(config, access_token, refresh_token, token_url, client_id):
    """Garante que o token de acesso seja válido, renovando ou pedindo novos dados."""
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
    """Remove caracteres inválidos no Windows."""
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def clean_input_val(val):
    """Limpa aspas e outros caracteres de campos colados."""
    if not val:
        return ""
    val = val.strip()
    if ":" in val:
        parts = val.split(":", 1)
        val = parts[1].strip()
    return val.strip(" \t\n\r\"',;")

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

def get_existing_solicitation(session, headers, cnpj, ref_date, doc_type, tipo_operacao):
    """Busca se já existe uma solicitação com as características desejadas na fila do SIGA."""
    url = "https://siga.sefaz.ce.gov.br/api/v1/solicitacoes"
    params = {
        'size': 1000,
        'sort': 'criacao,desc'
    }
    try:
        response = session.get(url, headers=headers, params=params, timeout=20)
        if response.status_code == 200:
            data = response.json()
            items = data if isinstance(data, list) else data.get("data", [])
            for item in items:
                if item.get("cnpj") == cnpj and item.get("tipo") == doc_type:
                    if doc_type == "NF_E":
                        filtro = item.get("filtro", {})
                        if (filtro.get("datReferencia") == [ref_date] and 
                            filtro.get("tipoOperacao") == tipo_operacao and 
                            filtro.get("formatoArquivo") == "xlsx"):
                            return item
                    else: # NFC_E
                        # Para NFC-e, como o filtro vem vazio do backend do SIGA,
                        # nós retornamos a mais recente encontrada para este CNPJ (o primeiro match)
                        return item
    except Exception as e:
        print(f"\n⚠️ Erro ao consultar lista de solicitações: {e}")
    return None

def download_file(session, headers, sol_id, dest_filepath):
    """Realiza o download da planilha XLSX a partir do ID da solicitação concluída."""
    url = f"https://siga.sefaz.ce.gov.br/api/v1/solicitacoes/{sol_id}/download"
    try:
        r = session.get(url, headers=headers, timeout=60)
        if r.status_code == 200:
            # Garante a criação da pasta antes de escrever
            os.makedirs(os.path.dirname(dest_filepath), exist_ok=True)
            with open(dest_filepath, "wb") as f:
                f.write(r.content)
            return True
        else:
            print(f"\n❌ Erro HTTP ao fazer download ({r.status_code}): {r.text}")
    except Exception as e:
        print(f"\n❌ Erro durante o download do arquivo: {e}")
    return False

def main():
    print("=" * 60)
    print("     SIGA SEFAZ-CE - COLETOR DE PLANILHAS NF-E (XLSX)")
    print("=" * 60)
    
    # 1. Carregar base de empresas salvas do SIGA
    if not os.path.exists(COMPANIES_CSV):
        print(f"❌ Base de empresas '{COMPANIES_CSV}' não encontrada.")
        print("Por favor, execute o script 'siga_extractor.py' primeiro.")
        sys.exit(1)
        
    siga_companies = []
    with open(COMPANIES_CSV, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            siga_companies.append(row)
            
    print(f"✅ Carregadas {len(siga_companies)} empresas do arquivo '{COMPANIES_CSV}'.")
    
    # 2. Carregar Excel de-para
    excel_mapping = load_excel_mapping(EXCEL_FILE)
    
    # 3. Carregar configurações de autenticação do SIGA
    config = load_config()
    access_token = config.get("access_token")
    refresh_token = config.get("refresh_token")
    client_id = config.get("client_id") or "painelind-frontend"
    token_url = config.get("token_url") or "https://sso.sefaz.ce.gov.br/auth/realms/sefaz-ad-realm/protocol/openid-connect/token"
    
    # Garantir token válido logo no início
    access_token, refresh_token = ensure_valid_token(config, access_token, refresh_token, token_url, client_id)
    
    # 4. Obter período desejado pelo usuário
    try:
        ano = int(input("\nDigite o ano de referência (ex: 2026): ").strip())
        mes_num = int(input("Digite o número do mês de referência (1 a 12): ").strip())
        if mes_num < 1 or mes_num > 12:
            raise ValueError()
    except ValueError:
        print("❌ Ano ou mês inválido. Saindo.")
        sys.exit(1)
        
    ref_date = f"{ano:04d}-{mes_num:02d}-01"
    date_subfolder = f"{mes_num:02d}{ano}"
    
    print(f"\n📅 Período selecionado: {mes_num:02d}/{ano}")
    
    # 4.1. Filtro opcional por empresa
    opcao_empresa = input("\nDeseja puxar dados de alguma empresa específica? (s/N): ").strip().lower()
    if opcao_empresa in ['s', 'sim']:
        cnpjs_input = input("Digite ou cole os CNPJs das empresas: ").strip()
        cnpjs_encontrados = re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{14}|\d{8}', cnpjs_input)
        cnpjs_filtrar = ["".join(filter(str.isdigit, c)) for c in cnpjs_encontrados]
        
        if cnpjs_filtrar:
            filtradas = []
            for comp in siga_companies:
                cnpj_comp = "".join(filter(str.isdigit, comp.get("cnpj", "")))
                if any(f in cnpj_comp for f in cnpjs_filtrar):
                    filtradas.append(comp)
            
            if filtradas:
                siga_companies = filtradas
                print(f"🎯 Filtrado! {len(siga_companies)} empresa(s) selecionada(s) para processar.")
            else:
                print("⚠️ Nenhuma empresa correspondente encontrada. Usando a lista completa.")
        else:
            print("⚠️ Nenhum CNPJ válido identificado. Usando a lista completa.")
            
    # 4.2. Escolha do tipo de documento
    print("\nTipos de planilhas para baixar:")
    print("1 - Apenas NF-e")
    print("2 - Apenas NFC-e")
    print("3 - Ambas (NF-e e NFC-e)")
    opcao_tipo = input("Selecione uma opção (1, 2 ou 3) [Padrão: 3]: ").strip()
    if opcao_tipo not in ["1", "2", "3"]:
        opcao_tipo = "3"
        
    # 5. Processamento principal
    session = requests.Session()
    
    for idx, comp in enumerate(siga_companies):
        cnpj = comp.get("cnpj", "").strip()
        razao_social = comp.get("razaoSocial", "").strip()
        cnpj_clean = "".join(filter(str.isdigit, cnpj))
        
        if not cnpj_clean:
            continue
            
        cod_excel, nome_excel = find_company_code(razao_social, excel_mapping)
        company_label = f"{cod_excel} - {nome_excel}" if cod_excel else f"{cnpj_clean} - {razao_social}"
        company_label = sanitize_filename(company_label)
        
        # Estrutura de pasta padrão: pasta-/[MMAAAA]/
        folder = os.path.join(f"{company_label}-", date_subfolder)
        
        print(f"\n⏳ [{idx+1}/{len(siga_companies)}] Processando: {company_label}")
        
        # Configuração de tipos de documentos e operações
        jobs = []
        if opcao_tipo in ["1", "3"]:
            jobs.extend([
                {"doc_type": "NF_E", "op": "SAIDA", "label": "nfe_emissor", "api_path": "nf-e"},
                {"doc_type": "NF_E", "op": "ENTRADA", "label": "nfe_destinatario", "api_path": "nf-e"}
            ])
        if opcao_tipo in ["2", "3"]:
            jobs.extend([
                {"doc_type": "NFC_E", "op": "SAIDA", "label": "nfce_emissor", "api_path": "nfc-e"},
                {"doc_type": "NFC_E", "op": "ENTRADA", "label": "nfce_destinatario", "api_path": "nfc-e"}
            ])
            
        for job in jobs:
            doc_type = job["doc_type"]
            operacao = job["op"]
            op_label = job["label"]
            api_path = job["api_path"]
            
            filepath = os.path.join(folder, f"{op_label}_{date_subfolder}.xlsx")
            
            # Se já foi baixado anteriormente
            if os.path.exists(filepath):
                print(f"  ➜ {op_label.upper()} já existe localmente. Pulando.")
                continue
                
            print(f"  ➜ Obtendo planilha de {op_label.upper()}... ", end="", flush=True)
            
            # Atualizar token se necessário antes de cada chamada
            access_token, refresh_token = ensure_valid_token(config, access_token, refresh_token, token_url, client_id)
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json',
                'Accept': 'application/json, text/plain, */*',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            # 1. Verifica se já existe solicitação pendente ou concluída no SIGA
            solicitação = get_existing_solicitation(session, headers, cnpj_clean, ref_date, doc_type, operacao)
            
            if solicitação:
                sol_id = solicitação.get("id")
                sol_status = solicitação.get("status")
                print(f"encontrada solicitação ativa (ID={sol_id} | Status={sol_status})")
            else:
                # 2. Se não existir, faz a requisição POST para gerar o download
                print("criando nova solicitação... ", end="", flush=True)
                post_url = f"https://siga.sefaz.ce.gov.br/api/v1/unidades/{cnpj_clean}/documentos-fiscais/{api_path}/download"
                payload = {
                    "datReferencia": [ref_date],
                    "tipoOperacao": operacao,
                    "formatoArquivo": "xlsx"
                }
                
                try:
                    r_post = session.post(post_url, headers=headers, json=payload, timeout=20)
                    if r_post.status_code == 202:
                        print("criada! ", end="", flush=True)
                        time.sleep(2)
                        solicitação = get_existing_solicitation(session, headers, cnpj_clean, ref_date, doc_type, operacao)
                        if solicitação:
                            sol_id = solicitação.get("id")
                            sol_status = solicitação.get("status")
                        else:
                            print("⚠️ Erro ao recuperar ID da nova solicitação.")
                            continue
                    elif r_post.status_code == 403:
                        print("❌ Sem procuração de acesso.")
                        folder_sem = "Sem Procuracao"
                        os.makedirs(folder_sem, exist_ok=True)
                        filepath_sem = os.path.join(folder_sem, f"{company_label}.txt")
                        with open(filepath_sem, "w", encoding="utf-8") as f_sem:
                            f_sem.write(f"Código/Nome: {company_label}\nCNPJ: {cnpj}\nRazão Social: {razao_social}\nStatus: Sem procuração de acesso cadastrada na SEFAZ para este usuário.")
                        break # Pula esta empresa para as demais operações
                    elif r_post.status_code == 409:
                        time.sleep(1)
                        solicitação = get_existing_solicitation(session, headers, cnpj_clean, ref_date, doc_type, operacao)
                        if solicitação:
                            sol_id = solicitação.get("id")
                            sol_status = solicitação.get("status")
                            print(f"recuperada após conflito (ID={sol_id})")
                        else:
                            print(f"❌ Conflito (HTTP 409) e falha ao obter ID.")
                            continue
                    else:
                        print(f"❌ Erro HTTP ao criar ({r_post.status_code})")
                        continue
                except Exception as e:
                    print(f"❌ Erro na requisição POST: {e}")
                    continue
            
            # 3. Monitoramento do status e Download
            concluido = False
            for tentativa in range(12): # Monitora por até 60 segundos por arquivo
                if sol_status == "CONCLUIDO":
                    concluido = True
                    break
                elif sol_status in ["ERRO", "FALHA"]:
                    print(f"    ❌ Solicitação falhou no processamento da SEFAZ (Status={sol_status})")
                    break
                
                print(".", end="", flush=True)
                time.sleep(5)
                
                solicitação = get_existing_solicitation(session, headers, cnpj_clean, ref_date, doc_type, operacao)
                if solicitação:
                    sol_status = solicitação.get("status")
            
            if concluido:
                print(" baixando planilha... ", end="", flush=True)
                if download_file(session, headers, sol_id, filepath):
                    print("✅ Sucesso!")
                else:
                    print("❌ Falha no download (talvez o arquivo esteja vazio ou indisponível).")
            else:
                print(" ⏳ Tempo esgotado (ainda processando na SEFAZ). Tente rodar o script novamente mais tarde.")
                
    print("\n🎉 Processamento concluído!")

if __name__ == "__main__":
    main()
