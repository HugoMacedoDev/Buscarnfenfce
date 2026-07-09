# SIGA SEFAZ-CE - Extrator de Contribuintes e NFC-e

Este repositório contém um conjunto de ferramentas em Python para automação e extração de dados do portal **SIGA da SEFAZ-CE** (Secretaria da Fazenda do Estado do Ceará), com suporte a renovação automática de sessão, download em massa multithread de XMLs fiscais de NFC-e (Nota Fiscal de Consumidor Eletrônica), estruturação de pastas e compactação automática em arquivos ZIP.

---

## 🛠️ Funcionalidades Principais

### 1. Extrator de Contribuintes (`siga_extractor.py`)
* Realiza a extração automatizada de todas as empresas do painel do SIGA.
* Gerencia a expiração do token de acesso de 300 segundos, realizando a **renovação automática** (`refresh_token`) via OAuth2/Keycloak da SEFAZ-CE em segundo plano.
* Salva os resultados estruturados em um arquivo CSV (`siga_empresas.csv`).

### 2. Extrator de NFC-e (`nfc_extractor.py`)
* Coleta chaves de acesso detalhadas das NFC-e no SIGA para o mês selecionado.
* Associa os dados com os códigos internos das empresas a partir de uma planilha auxiliar (`empresas cod.xlsx`).
* Realiza o **download multithread de alta performance** (configurado para 40 threads paralelas) dos XMLs fiscais diretamente da SEFAZ.
* Limpa os namespaces dos XMLs para facilitar manipulações futuras.
* Organiza os arquivos de saída em pastas bem estruturadas: `[CÓDIGO/CNPJ] - [NOME_EMPRESA]/MMAAAA/`.
* Compacta automaticamente os XMLs baixados em arquivos ZIP individuais para cada empresa.

### 3. Extrator de NF-e XLSX (`nfe_xlsx_extractor.py`)
* Automatiza a geração e o download de planilhas XLSX contendo todas as chaves de **NF-e** das empresas.
* Dá suporte a ambas as modalidades: **Emissor** (saídas) e **Destinatário** (entradas).
* Verifica se a solicitação de exportação já existe para evitar reprocessamentos desnecessários (erro 409).
* Monitora o processamento da SEFAZ até que o arquivo esteja pronto para download e salva-o estruturadamente na pasta de cada empresa.

### 4. Coletor de XMLs de NF-e (`nfe_xml_downloader.py`)
* Escaneia e identifica certificados digitais A1 (`.pfx` ou `.p12`) salvos na raiz do diretório.
* Lê as chaves de acesso diretamente das planilhas XLSX geradas no passo anterior.
* Conecta via SOAP ao WebService oficial da SEFAZ Nacional utilizando o certificado correspondente e baixa os XMLs das NF-e.
* Salva os arquivos XML completos diretamente na pasta do respectivo mês de cada empresa.

---

## 📂 Estrutura de Arquivos

*   `siga_extractor.py`: Script para extração da lista de empresas e monitoramento de tokens.
*   `nfc_extractor.py`: Script de download multithread e compactação das NFC-es.
*   `nfe_xlsx_extractor.py`: Script para download automatizado de relatórios XLSX de NF-e (Emissor e Destinatário).
*   `nfe_xml_downloader.py`: Script para download dos XMLs completos das NF-e utilizando os certificados A1.
*   `config.json`: Cache local das configurações de conexão e tokens OAuth2 (ignorado pelo Git).
*   `api_token.txt`: Token de autorização ativo para as requisições à API da SEFAZ (ignorado pelo Git).
*   `empresas cod.xlsx`: Planilha com os códigos e identificadores internos das empresas.
*   `.gitignore`: Arquivo para evitar o rastreamento de chaves de API, tokens, bancos de dados temporários e dados baixados locais.

---

## 🚀 Como Configurar e Executar

### Pré-requisitos
*   Python 3.10 ou superior
*   Bibliotecas adicionais:
    ```bash
    pip install requests openpyxl
    ```

### Configuração
1.  Obtenha os dados de acesso e configure o arquivo `config.json` na raiz do projeto:
    ```json
    {
      "access_token": "SEU_ACCESS_TOKEN",
      "refresh_token": "SEU_REFRESH_TOKEN",
      "client_id": "painelind-frontend",
      "token_url": "https://sso.sefaz.ce.gov.br/auth/realms/sefaz-ad-realm/protocol/openid-connect/token",
      "api_url": "https://siga.sefaz.ce.gov.br/api/v1/unidades-resumo-malha"
    }
    ```
2.  Insira o token de autenticação no arquivo `api_token.txt`.
3.  Garanta que a planilha `empresas cod.xlsx` esteja preenchida com as empresas a extrair.

### Execução
Para extrair a lista de empresas do painel do SIGA:
```bash
python siga_extractor.py
```

Para extrair e compilar os XMLs de NFC-e:
```bash
python nfc_extractor.py
```

Para baixar as planilhas XLSX das NF-e (Emissor e Destinatário):
```bash
python nfe_xlsx_extractor.py
```

Para baixar os XMLs completos das NF-e usando os certificados A1:
```bash
python nfe_xml_downloader.py
```

---

## 🛡️ Segurança e LGPD

Este repositório possui regras estritas no arquivo `.gitignore` para assegurar que:
1.  Nenhum token ou chave de API seja exposto publicamente (`api_token.txt` e `config.json` são ignorados).
2.  Os dados fiscais e cadastrais extraídos das empresas **permaneçam apenas localmente** na sua máquina e nunca sejam commitados para o repositório público do GitHub.
