cat << 'EOF' > README.md
# Extrator de Metadados de Documentos (.docx)

Este projeto automatiza a extração de metadados de documentos Word (.docx), convertendo informações brutas em planilhas Excel organizadas. O objetivo é facilitar a conferência de autoria e títulos em grandes volumes de arquivos dentro de órgãos públicos ou empresas.

## 🚀 Objetivo
Otimizar o fluxo de trabalho administrativo, permitindo a leitura em massa de arquivos para identificar autores e títulos sem a necessidade de abertura manual de cada documento.

## 🛠️ Tecnologias Utilizadas
* **Linguagem**: Python 3.12
* **Bibliotecas Principais**:
    * `pandas`: Manipulação de dados e geração de planilhas Excel.
    * `python-docx`: Leitura de propriedades e metadados de arquivos Word.
    * `os`: Gerenciamento de arquivos e caminhos no sistema operacional.
* **Sistema Operacional**: Pop!_OS (Linux) com ambiente virtual `.venv`.
* **Versionamento**: Git e GitHub com autenticação via SSH.

## ⚙️ Funcionalidades
* **Resiliência**: Tratamento de exceções (Try/Except) para processar pastas contendo arquivos vazios ou corrompidos sem interromper a execução.
* **Dicionários Dinâmicos**: Organização dos metadados em estruturas de dados Python antes da exportação.
* **Exportação Direta**: Criação automática do arquivo `relatorio_documentos.xlsx` na raiz do projeto.

## 🔧 Como Executar

1. **Clone o repositório**:
   ```bash
   git clone git@github.com:andrebonfim/extrator-prefeitura.git
   ```

2. **Configure o ambiente virtual**:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

3. **Instale as dependências**:
   ```bash
   pip install pandas python-docx openpyxl
   ```

4. **Execute o extrator**:
   ```bash
   python3 main.py
   ```
