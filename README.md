# CONTROLE DE FERRAMENTAS

Sistema web desenvolvido em **Python e Flask** para cadastro e acompanhamento de defeitos em ferramentas, com armazenamento de informações em banco de dados PostgreSQL e exportação de registros para Excel.

## Tecnologias

- **Python** – linguagem principal.
- **PostgreSQL** – Banco de dados para armazenamento dos registros.
- **HTML, CSS e JavaScript** – interface do sistema.

## Bibliotecas python
- **Flask** – desenvolvimento da aplicação web.
- **Psycopg2** – conexão entre Python e PostgreSQL.
- **Pillow (PIL)** – manipulação de imagens.
- **OpenPyXL** – leitura e edição de planilhas Excel.


## Funcionalidades

- Cadastro de defeitos em ferramentas com registro de descrição, data, ciclo e evidências fotográficas.
- Consulta e filtragem de registros.
- Acompanhamento e atualização do status de manutenção.
- Exportação dos registros para Excel.

## Como executar localmente

### 1. Pré-requisitos

Instale os seguintes programas:

- [Python](https://www.python.org/downloads/)
- [PostgreSQL](https://www.postgresql.org/download/)

### 2. Clone o repositório e entre na pasta do projeto:

```bash
cd "CONTROLE DE FERRAMENTAS"
```

### 3. Crie um ambiente virtual

```bash
python -m venv .venv
```

Ative o ambiente virtual.

**Windows:**

```bash
.venv\Scripts\activate
```

**Linux:**

```bash
source .venv/bin/activate
```

### 4. Instale as dependências

```bash
pip install Flask psycopg2 openpyxl Pillow
```

### 5. Configure o banco de dados

Crie um banco de dados PostgreSQL chamado:

```sql
CREATE DATABASE controle_de_ferramentas;
```

No arquivo `main.py`, localize a variável `DB_CONFIG` e configure os dados de acesso ao seu PostgreSQL:

```python
DB_CONFIG = {
    "host": "localhost",
    "database": "controle_de_ferramentas",
    "user": "SEU_USUARIO",
    "password": "SUA_SENHA",
    "port": "5432"
}
```

**Importante:** a estrutura das tabelas do banco deve corresponder à utilizada pelas consultas do projeto (o sistema não cria a tabela caso ela não exista, isso será alterado futuramente). Por isso, crie a seguinte tabela no postgreSQL:

```sql
CREATE TABLE IF NOT EXISTS defeito (
    id SERIAL PRIMARY KEY,
    ferramenta VARCHAR(255) NOT NULL,
    descricao TEXT,
    evidencia_foto VARCHAR(255),
    ciclo VARCHAR(100),
    data DATE,
    em_manutencao BOOLEAN DEFAULT FALSE
);
```

### 6. Execute a aplicação

Com o ambiente virtual ativado e o banco configurado, execute:

```bash
python main.py
```

Acesse pelo navegador:

```text
http://127.0.0.1:5000
```

## Estrutura do projeto

```text
CONTROLE DE FERRAMENTAS/
│
├── main.py
├── ferramentas.xlsx
│
├── static/
│   ├── style.css
│   ├── uploads/
│   └── ...
│
└── templates/
    ├── index.html
    ├── registros.html
    ├── manutencao.html
    └── ...
```

## Observações

- O PostgreSQL precisa estar instalado e em execução.
- Os arquivos Excel devem estar disponíveis conforme os caminhos utilizados pela aplicação.
- A pasta `static/uploads/` deve existir e permitir a gravação de arquivos.
- As credenciais do banco devem ser configuradas de acordo com o ambiente local.
---

Ainda em desenvolvimento...