# SIGA SEFAZ-CE - Automação Fiscal e Extrator de NFC-e / NF-e

Este repositório contém uma suíte de ferramentas em Python para automação, consulta e extração de documentos fiscais eletrônicos diretamente do portal **SIGA da SEFAZ-CE** (Secretaria da Fazenda do Estado do Ceará), com suporte a renovação automática de autenticação OAuth2, download multithread de XMLs, deduplicação por histórico prévio, geração de planilhas e auditoria de execuções.

---

## Ferramentas e Funcionalidades

### 1. Extrator de Contribuintes (`siga_extractor.py`)
* Consulta paginada de todas as empresas e estabelecimentos vinculados no painel do SIGA.
* Gerenciamento de sessão com **renovação automática do token** (`refresh_token`) via Keycloak/OAuth2 da SEFAZ-CE.
* Gera o arquivo estruturado `siga_empresas.csv` contendo Razão Social, CNPJ, CGF e dados cadastrais.

### 2. Extrator e Baixador de NFC-e (`nfc_extractor.py`)
* **Consulta de Chaves de NFC-e**: Coleta as chaves de acesso das NFC-e emitidas para o mês/ano selecionado.
* **Filtro de Documentos Cancelados**: Configuração flexível por empresa para incluir ou ignorar notas fiscais canceladas.
* **Controle Inteligente de Histórico (Chaves Cadastradas)**:
  * Lê a base de chaves já existentes a partir de `chaves_ja_cadastradas.csv` (ou `chaves_ja_baixadas.csv`) em modo somente leitura (*read-only*).
  * Reconhece arquivos com ou sem cabeçalho, com 2 colunas (`CNPJ, CHAVE`) ou apenas 1 coluna (`CHAVE`), extraindo automaticamente o CNPJ do emitente contido na própria chave (posições 7 a 20).
  * Compara e filtra as chaves em memória **especificamente por CNPJ**, evitando downloads duplicados e economizando requisições.
* **Divisão em Lotes (Partes de até 4.000 XMLs)**:
  * Agrupa as chaves pendentes de cada empresa em lotes de até 4.000 arquivos (ex: `NFCe_Parte_1`, `NFCe_Parte_2`, etc.).
* **Download Multithread de Alta Performance**:
  * Realiza o download paralelo dos XMLs utilizando até 40 threads simultâneas.
  * Remove namespaces dos XMLs para simplificar a leitura por sistemas ERPs e bancos de dados.
* **Auditoria de Execução**:
  * Registra todas as chaves processadas no arquivo `chaves_baixadas_nesta_execucao.csv`.
* **Tratamento de Empresas Sem Procuração**:
  * Identifica bloqueios de permissão (HTTP 403) e gera relatórios automáticos na pasta `Sem Procuracao/`.
* **Resumo Consolidado**:
  * Ao final, exibe um painel analítico por empresa e o balanço total (chaves no SIGA, já existentes, novas identificadas, baixadas e falhas).

### 3. Extrator de NF-e em Excel (`nfe_xlsx_extractor.py`)
* Consulta direta via API paginada das **NF-e (modelo 55)** para os tipos de operação:
  * **Saída / Emissor**
  * **Entrada / Destinatário**
* Compila os registros em planilhas Excel (`.xlsx`) formatadas com os campos oficiais da SEFAZ (CNPJ/Razão Social, UF, Número da Nota, Data de Emissão, Status de Processamento, Valor R$ e Chave de Acesso).
* Salva os arquivos organizados na pasta de cada empresa e mês correspondente.

### 4. Coletor de XMLs de NF-e via WebService (`nfe_xml_downloader.py`)
* Identifica certificados digitais A1 (`.pfx` / `.p12`) salvos no diretório.
* Lê as chaves de acesso das planilhas Excel geradas pelo `nfe_xlsx_extractor.py`.
* Realiza a comunicação SOAP segura com os WebServices da SEFAZ Nacional para baixar os XMLs completos das NF-e.

---

## Estrutura do Projeto

```text
├── nfc_extractor.py                   # Script principal de extração e download multithread de NFC-e
├── nfe_xlsx_extractor.py              # Extrator de relatórios de NF-e em formato Excel (.xlsx)
├── nfe_xml_downloader.py              # Download de XMLs de NF-e via certificado digital A1
├── siga_extractor.py                  # Extrator do cadastro geral de empresas vinculadas
├── empresas cod.xlsx                  # Planilha com o mapeamento e códigos internos das empresas
├── chaves_ja_cadastradas.csv          # Histórico de chaves já salvas no banco (somente leitura)
├── chaves_baixadas_nesta_execucao.csv # Registro/auditoria das chaves baixadas na execução atual
├── siga_empresas.csv                  # Lista exportada de empresas obtidas pelo siga_extractor.py
├── api_token.txt                      # Token JWT ativo da SEFAZ (ignorado pelo Git)
├── config.json                        # Configurações de autenticação e endpoints OAuth2 (ignorado pelo Git)
├── requirements.txt                   # Dependências do projeto
├── siga_nfce/                         # Diretório onde os XMLs de NFC-e são salvos e organizados
│   └── [EMPRESA_COD_NOME]/
│       └── [MMAAAA]/
│           ├── NFCe_Parte_1/
│           │   └── *.xml
│           └── NFCe_Parte_2/
│               └── *.xml
└── Sem Procuracao/                    # Relatório de empresas sem autorização ativa na SEFAZ
```

---

## Instalação e Configuração

### 1. Pré-requisitos
* **Python 3.10** ou superior instalado.
* Instale as dependências executando:
```bash
pip install -r requirements.txt
```

### 2. Configuração de Credenciais
Crie ou configure o arquivo `config.json` na raiz do projeto com as credenciais do SIGA:
```json
{
  "access_token": "SEU_ACCESS_TOKEN",
  "refresh_token": "SEU_REFRESH_TOKEN",
  "client_id": "painelind-frontend",
  "token_url": "https://sso.sefaz.ce.gov.br/auth/realms/sefaz-ad-realm/protocol/openid-connect/token",
  "api_url": "https://siga.sefaz.ce.gov.br/api/v1/unidades-resumo-malha"
}
```

Insira também o token de autenticação ativo no arquivo `api_token.txt`.

### 3. Planilha de Empresas
Preencha a planilha `empresas cod.xlsx` com os códigos, CNPJs e nomes das empresas que deseja processar.

---

## Como Executar

### Extrair cadastro de empresas:
```bash
python siga_extractor.py
```

### Extrair e baixar XMLs de NFC-e (com deduplicação e multithread):
```bash
python nfc_extractor.py
```

### Extrair relatórios de NF-e em Excel (Saídas e Entradas):
```bash
python nfe_xlsx_extractor.py
```

### Baixar XMLs de NF-e com Certificado Digital A1:
```bash
python nfe_xml_downloader.py
```

---

## Segurança e LGPD

* Os arquivos com tokens, chaves de autenticação e caches locais (`api_token.txt`, `config.json`) estão configurados no `.gitignore` e **nunca são enviados ao repositório**.
* Os arquivos fiscais e XMLs baixados permanecem armazenados estritamente na máquina local.
