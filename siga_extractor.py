#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SIGA SEFAZ-CE - Extrator Automático de Contribuintes em Python

Este script roda localmente e realiza a extração automatizada de todas as empresas
do painel do SIGA. Ele gerencia a expiração do token de 300 segundos realizando a 
renovação automática (refresh_token) via OAuth2/Keycloak da SEFAZ em segundo plano.
"""

import os
import sys
import json
import time
import base64
import csv
import requests

CONFIG_FILE = "config.json"
DEFAULT_CLIENT_ID = "painelind-frontend"
DEFAULT_TOKEN_URL = "https://sso.sefaz.ce.gov.br/auth/realms/sefaz-ad-realm/protocol/openid-connect/token"
DEFAULT_API_URL = "https://siga.sefaz.ce.gov.br/api/v1/unidades-resumo-malha"
OUTPUT_CSV = "siga_empresas.csv"

def load_config():
    """Carrega as configurações e tokens salvos em cache."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️ Erro ao carregar cache de configuração: {e}")
    return {}

def save_config(config):
    """Grava as configurações e tokens em cache para evitar redigitação."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"⚠️ Erro ao salvar cache de configuração: {e}")

def is_token_expired(token):
    """Verifica se o token JWT de acesso está expirado ou prestes a expirar (< 30 segundos)."""
    if not token:
        return True
    try:
        parts = token.split('.')
        if len(parts) != 3:
            return True
        
        # Decodifica o payload do JWT (segunda parte)
        payload_b64 = parts[1]
        payload_b64 += '=' * (-len(payload_b64) % 4) # Corrige preenchimento base64
        payload_json = base64.urlsafe_b64decode(payload_b64).decode('utf-8')
        payload = json.loads(payload_json)
        
        exp = payload.get('exp', 0)
        now = time.time()
        
        # Retorna True se o token expirar em menos de 30 segundos
        return (exp - now) < 30
    except Exception:
        return True # Assume expirado caso ocorra falha de decodificação

def refresh_access_token(token_url, refresh_token, client_id):
    """Executa uma chamada HTTP para a SEFAZ e renova o token de acesso."""
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
            new_access_token = res_json.get('access_token')
            new_refresh_token = res_json.get('refresh_token', refresh_token)
            return new_access_token, new_refresh_token
        else:
            print(f"\n❌ Erro ao renovar token (HTTP {response.status_code}).")
            print(f"Resposta do servidor: {response.text}")
            return None, None
    except Exception as e:
        print(f"\n❌ Erro de conexão ao renovar token: {e}")
        return None, None

def save_to_csv(companies, filename):
    """Exporta a lista de dicionários para um arquivo CSV formatado para Excel Brasil."""
    if not companies:
        print("⚠️ Nenhuma empresa para salvar.")
        return
    
    # Coleta todas as chaves únicas encontradas nas empresas para servirem de colunas
    headers = set()
    for comp in companies:
        headers.update(comp.keys())
    headers = sorted(list(headers))
    
    try:
        # utf-8-sig adiciona o BOM do UTF-8 para que o Excel abra os acentos corretamente
        with open(filename, mode="w", newline="", encoding="utf-8-sig") as f:
            # Separador ponto-e-vírgula é o padrão para o Excel em português
            writer = csv.DictWriter(f, fieldnames=headers, delimiter=";")
            writer.writeheader()
            
            for comp in companies:
                # Transforma valores que são listas/objetos em strings JSON
                row = {}
                for k, v in comp.items():
                    if v is None:
                        row[k] = ""
                    elif isinstance(v, (list, dict)):
                        row[k] = json.dumps(v, ensure_ascii=False)
                    else:
                        row[k] = str(v)
                writer.writerow(row)
                
        print(f"\n💾 Sucesso! {len(companies)} empresas salvas no arquivo: {os.path.abspath(filename)}")
    except Exception as e:
        print(f"\n❌ Erro ao salvar arquivo CSV: {e}")

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

def main():
    print("=" * 60)
    print("      SIGA SEFAZ-CE - EXTRATOR DE CONTRIBUINTES EM PYTHON")
    print("=" * 60)
    
    # Carregar cache de tokens
    config = load_config()
    
    access_token = config.get("access_token")
    refresh_token = config.get("refresh_token")
    client_id = config.get("client_id", DEFAULT_CLIENT_ID)
    token_url = config.get("token_url", DEFAULT_TOKEN_URL)
    api_url = config.get("api_url", DEFAULT_API_URL)
    
    # Solicitar ou alterar configurações
    alterar = False
    if access_token:
        print(f"ℹ️ Configurações carregadas do cache:")
        print(f"   URL da API: {api_url}")
        print(f"   Token de Acesso: {access_token[:20]}...{access_token[-20:] if len(access_token) > 40 else ''}")
        
        # Verificar se o refresh_token do cache já expirou
        if is_token_expired(refresh_token):
            print("\n⚠️ O seu Refresh Token salvo expirou. Você precisará fornecer novos tokens.")
            alterar = True
        else:
            opcao = input("\nDeseja alterar as credenciais ou a URL da API? (s/N): ").strip().lower()
            if opcao == 's':
                alterar = True
                
    if not access_token or alterar:
        print("\n--- MENU DE CONFIGURAÇÃO ---")
        
        # Alterar API
        nova_api = clean_input_val(input(f"Cole a URL da API [{api_url}]: "))
        if nova_api:
            api_url = nova_api
            
        # Alterar tokens
        print("\nComo obter os tokens: Abra a aba Network (F12) no SIGA, localize a chamada '/token' e copie os valores.")
        novo_access = clean_input_val(input("Cole o novo access_token (ou Enter para manter o atual): "))
        if novo_access:
            access_token = novo_access
            
        novo_refresh = clean_input_val(input("Cole o novo refresh_token (ou Enter para manter o atual): "))
        if novo_refresh:
            refresh_token = novo_refresh
            
    # Salvar nas configurações atuais
    config.update({
        "access_token": access_token,
        "refresh_token": refresh_token,
        "client_id": client_id,
        "token_url": token_url,
        "api_url": api_url
    })
    save_config(config)

    # Iniciar processo de extração
    companies = []
    page = 0
    size = 1000  # Tamanho seguro por lote (aumentado para 1000 para ser mais rápido)
    total_pages = 1
    total_elements = 0
    
    session = requests.Session()
    
    print("\n🚀 Iniciando extração total...")
    
    try:
        while page < total_pages:
            print(f"⏳ Requisitando página {page + 1} de {total_pages}... ", end="", flush=True)
            
            # 1. Verificar e renovar token de acesso se necessário
            if is_token_expired(access_token):
                print("\n🔄 Token de Acesso expirado ou prestes a expirar. Renovando... ", end="", flush=True)
                new_acc, new_ref = refresh_access_token(token_url, refresh_token, client_id)
                if new_acc:
                    access_token = new_acc
                    refresh_token = new_ref
                    config.update({"access_token": access_token, "refresh_token": refresh_token})
                    save_config(config)
                    print("Renovado! Continuando requisição... ", end="", flush=True)
                else:
                    print("\n❌ Falha na renovação do token. Parando extração para evitar erros.")
                    break
            
            # 2. Configurar requisição
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Accept': 'application/json, text/plain, */*',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Origin': 'https://siga.sefaz.ce.gov.br',
                'Referer': 'https://siga.sefaz.ce.gov.br/ui/selecao-contribuinte/contribuinte'
            }
            
            params = {
                'page': page,
                'size': size
            }
            
            try:
                response = session.get(api_url, headers=headers, params=params, timeout=20)
                
                # Se der 401, tenta renovação emergencial na hora
                if response.status_code == 401:
                    print("\n🔄 Erro 401 (Não Autorizado). Forçando renovação imediata... ", end="", flush=True)
                    new_acc, new_ref = refresh_access_token(token_url, refresh_token, client_id)
                    if new_acc:
                        access_token = new_acc
                        refresh_token = new_ref
                        config.update({"access_token": access_token, "refresh_token": refresh_token})
                        save_config(config)
                        print("Renovado! Tentando novamente... ", end="", flush=True)
                        continue # Refaz a iteração da mesma página
                    else:
                        print("\n❌ Falha crítica ao renovar token. Extração abortada.")
                        break
                
                if response.status_code != 200:
                    print(f"\n❌ Erro HTTP {response.status_code}: {response.text}")
                    print("Aguardando 5 segundos antes de tentar novamente...")
                    time.sleep(5)
                    continue
                
                # 3. Parsear dados da resposta
                data = response.json()
                
                # Extrair array de dados
                content = []
                if isinstance(data, list):
                    content = data
                elif isinstance(data, dict):
                    if "content" in data and isinstance(data["content"], list):
                        content = data["content"]
                    elif "data" in data and isinstance(data["data"], list):
                        content = data["data"]
                    else:
                        for k, v in data.items():
                            if isinstance(v, list):
                                content = v
                                break
                                
                if not content:
                    print("Página vazia. Concluído.")
                    break
                
                companies.extend(content)
                print(f"OK! (+{len(content)} empresas. Total: {len(companies)})")
                
                # Parsear paginação
                if isinstance(data, dict):
                    if "totalPages" in data:
                        total_pages = data["totalPages"]
                    elif "pageCount" in data:
                        total_pages = data["pageCount"]
                        
                    if "totalElements" in data:
                        total_elements = data["totalElements"]
                    elif "total" in data:
                        total_elements = data["total"]
                
                # Se não houver paginação no json, para quando retornar menos registros que o lote
                if total_pages == 1 and len(content) < size:
                    break
                    
                page += 1
                
                # Atraso de cortesia de 0.1 segundo (reduzido para acelerar o processo)
                time.sleep(0.1)
                
            except requests.exceptions.RequestException as re:
                print(f"\n⚠️ Erro de rede ao buscar página: {re}")
                print("Aguardando 5 segundos antes de tentar novamente...")
                time.sleep(5)
                
    except KeyboardInterrupt:
        print("\n\n🛑 Extração interrompida pelo usuário via Ctrl+C!")
    
    # 4. Salvar resultados coletados
    save_to_csv(companies, OUTPUT_CSV)

if __name__ == "__main__":
    main()
