# ⚡ exelToSql — Gerador Dinâmico de Scripts SQL a partir de Excel

<div align="center">

  <p align="center">
    <strong>Utilitário web para automação de cargas massivas, saneamento de dados e geração de scripts SQL parametrizados a partir de planilhas <code>.xlsx</code>.</strong>
  </p>

  <p align="center">
    <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python" />
    <img src="https://img.shields.io/badge/Flask-Web_Framework-000000?style=for-the-badge&logo=flask&logoColor=white" alt="Flask" />
    <img src="https://img.shields.io/badge/Pandas-Data_Processing-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas" />
    <img src="https://img.shields.io/badge/SQL-Server_%2F_Firebird_%2F_PostgreSQL-CC292B?style=for-the-badge&logo=sqlite&logoColor=white" alt="SQL" />
    <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="MIT License" />
  </p>

</div>

---

## 📌 Contexto & Por que este projeto existe?

Em ambientes corporativos de suporte avançado e desenvolvimento de software, é rotineiro receber **planilhas do Excel enviadas por clientes** contendo cadastros legados, tabelas financeiras ou dados que precisam ser migrados para o banco de dados relacional.

Tradicionalmente, desenvolvedores acabam recorrendo a fórmulas manuais no Excel (como `=CONCATENAR("INSERT INTO...")`), um processo que é:
- Lento e tedioso;
- Suscetível a erros de escape de aspas e formatação de nulos;
- Difícil de reproduzir com consistência em planilhas volumosas.

O **exelToSql** automatiza esse fluxo por completo: você envia a planilha, seleciona visualmente a aba e as colunas, digita um modelo de comando SQL com placeholders `{NomeDaColuna}` e o sistema gera instantaneamente o arquivo `.sql` pronto para execução em **SQL Server, Firebird, PostgreSQL, Oracle ou MySQL**.

---

## ✨ Funcionalidades Principais

- 📂 **Upload e Inspeção de Planilhas:** Suporte nativo a arquivos `.xlsx`.
- 📑 **Detecção Dinâmica de Abas & Colunas:** Lê os metadados da planilha via **Pandas** e disponibiliza as colunas em tela sem recarregar a página (Fetch/AJAX).
- 🧩 **Templates SQL Personalizados:** Suporte flexível para comandos `INSERT`, `UPDATE`, `UPSERT` ou chamadas de `Stored Procedures` utilizando `{NomeDaColuna}` como variáveis.
- 💾 **Exportação Pronta para Produção:** Download de arquivo `.sql` consolidado e formatado.
- 🔒 **Isolamento por Sessão:** Cada requisição opera em uma sessão exclusiva via `uuid.uuid4()`, impedindo que os arquivos de um usuário sejam sobrescritos por outro.
- 🧹 **Auto-Purge de Temporários:** Rotina periódica que descarta arquivos de sessões inativas após 15 minutos, prevenindo consumo excessivo de disco no servidor.

---

## 💡 Como Funciona (Exemplo Prático)

Suponha que você recebeu uma planilha com a seguinte aba de alunos:

| Matricula | Nome | Cidade | CPF |
| :---: | :--- | :--- | :---: |
| 1001 | João Silva | Goiânia | 12345678901 |
| 1002 | Maria Souza | Anápolis | 98765432100 |

### 1. Template Configurado na Ferramenta:
```sql
INSERT INTO TBALUNO (MATRICULA, NOME, CIDACODIGO, CPF) 
VALUES ({Matricula}, '{Nome}', (SELECT CIDACODIGO FROM TBCIDADE WHERE DESCRICAO = '{Cidade}'), '{CPF}');
```

### 2. Script SQL Gerado Automaticamente:
```sql
INSERT INTO TBALUNO (MATRICULA, NOME, CIDACODIGO, CPF) 
VALUES (1001, 'João Silva', (SELECT CIDACODIGO FROM TBCIDADE WHERE DESCRICAO = 'Goiânia'), '12345678901');

INSERT INTO TBALUNO (MATRICULA, NOME, CIDACODIGO, CPF) 
VALUES (1002, 'Maria Souza', (SELECT CIDACODIGO FROM TBCIDADE WHERE DESCRICAO = 'Anápolis'), '98765432100');
```

---

## 🏛️ Estrutura do Projeto

```
exelToSql/
├── 📂 static/           # Estilos CSS e scripts JavaScript interativos (AJAX)
├── 📂 templates/        # Interface HTML renderizada com Jinja2 (Flask)
├── 📂 uploads/          # Diretório temporário dos arquivos Excel enviados
├── 📂 generated/        # Diretório temporário dos arquivos SQL gerados
├── 📄 app.py            # Servidor Flask, rotas e rotina de limpeza temporária
├── 📄 requirements.txt  # Dependências do ecossistema Python (Flask, Pandas, Openpyxl)
└── 📄 LICENSE           # Licença MIT
```

---

## 🚀 Como Executar Localmente

### 📋 Pré-requisitos
- [Python 3.10+](https://www.python.org/downloads/) instalado.
- Gerenciador de pacotes `pip`.

---

### ⚙️ Passo a Passo

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/FellypeAlexandre/exelToSql.git
   cd exelToSql
   ```

2. **Crie e ative um ambiente virtual (recomendado):**
   ```bash
   # No Windows (PowerShell):
   python -m venv venv
   .\venv\Scripts\Activate.ps1

   # No Linux / macOS:
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Execute o servidor:**
   ```bash
   python app.py
   ```

5. **Acesse no navegador:**
   Abra [http://localhost:5000](http://localhost:5000) e faça o upload da sua planilha!

---

## 🛠️ Tecnologias Utilizadas

| Tecnologia | Finalidade |
| :--- | :--- |
| **Python** | Linguagem base para processamento de dados e servidor |
| **Flask** | Micro-framework web leve e veloz para roteamento e templates |
| **Pandas & OpenPyXL** | Leitura analítica de estruturas tabulares e manipulação de DataFrames |
| **JavaScript (ES6)** | Requisições assíncronas assíncronas (Fetch API) para interação fluida |
| **HTML5 & CSS3** | Interface de usuário limpa e intuitiva |

---

## 👨‍💻 Autor

Desenvolvido por **Fellype Alexandre**:
- 💼 [LinkedIn](https://linkedin.com/in/fellype-alexandre-02762517a)
- 🐙 [GitHub](https://github.com/FellypeAlexandre)
- ✉️ [fellype200313@gmail.com](mailto:fellype200313@gmail.com)

---

## 📄 Licença

Este projeto está sob a licença [MIT](LICENSE).
