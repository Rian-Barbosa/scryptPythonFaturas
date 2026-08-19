# Automação de Faturas em Python

Este projeto automatiza a extração de dados de faturas em formato PDF e o preenchimento dessas informações em uma planilha do Excel (`.xlsx`). 

Ele foi desenvolvido para agilizar o processo de conciliação financeira, eliminando a digitação manual de identificadores e valores de transação.

---

## 🛠️ Tecnologias Utilizadas

*   **Python 3.11+**
*   **[PyMuPDF (fitz)](https://pymupdf.readthedocs.io/)** — Para extração e leitura rápida de textos dos arquivos PDF.
*   **[openpyxl](https://openpyxl.readthedocs.io/)** — Para manipulação e preenchimento da planilha Excel modelo.

---

## 📁 Estrutura do Projeto

*   `processar_faturas.py`: O script principal que contém toda a lógica de leitura e preenchimento.
*   `Fatura detalhada para preencher.xlsx`: A planilha modelo (template) contendo as colunas a serem preenchidas (ex: ID da transação, Awareness, Performance, Post).
*   `instalar_requisitos.bat`: Script automatizado para instalar o Python (se necessário) e as dependências do projeto.
*   `rodar_automacao.bat`: Executa o script principal com apenas dois cliques no Windows.
*   `pdfs/`: Pasta onde devem ser colocados os arquivos de fatura originais em formato PDF.

---

## 🚀 Como Usar

### Passo 1: Instalação das dependências
Dê um duplo clique no arquivo `instalar_requisitos.bat`.
*   Se você não tiver o Python instalado, ele fará o download e iniciará o instalador oficial. *Lembre-se de marcar a caixa "Add python.exe to PATH"*.
*   Se o Python já estiver presente, ele instalará automaticamente as bibliotecas necessárias (`pymupdf` e `openpyxl`).

### Passo 2: Preparação dos arquivos
1. Salve as faturas em PDF que deseja processar dentro da pasta `pdfs/`. Você pode colocar quantos PDFs quiser.
2. Certifique-se de que a planilha modelo `Fatura detalhada para preencher.xlsx` está na mesma pasta que o script.

### Passo 3: Execução
Dê um duplo clique em `rodar_automacao.bat`. 

O script irá:
1. Ler e analisar os PDFs na pasta `pdfs/`.
2. Identificar e extrair os dados de ID da Transação, Subtotal e campanhas agrupadas por tipo (*Awareness, Performance, Post*).
3. Gerar um novo arquivo chamado `Fatura detalhada preenchida resultado.xlsx` com todos os valores preenchidos corretamente, sem alterar a planilha modelo original.
