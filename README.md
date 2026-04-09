# 📄 Leitor de Boletos PDF (RPA)

Automação desenvolvida em Python para leitura de boletos em PDF, extração de informações relevantes e geração automática de uma planilha Excel com os dados organizados.

---

## 🚀 Objetivo

Automatizar o processo manual de leitura de boletos condominiais, reduzindo erros humanos e aumentando a produtividade no lançamento de informações no sistema.

---

## ⚙️ Funcionalidades

* 📥 Leitura automática de arquivos PDF em um diretório
* 🔍 Extração de dados via Regex:

  * CNPJ
  * Data de vencimento
  * Valor do documento
  * Número do documento
* 📊 Geração de planilha Excel automatizada
* 🧠 Tratamento de variações de layout (quebras de linha e formatos diferentes)
* 📂 Abertura automática do arquivo Excel ao final do processamento

---

## 🛠️ Tecnologias utilizadas

* 🐍 Python 3
* 📄 pdfplumber (extração de texto de PDFs)
* 📊 openpyxl (geração de planilhas Excel)
* 🔎 re (Regex para extração de dados)

---

## 📁 Estrutura do projeto

```
rpa-pdf/
│
├── pdfs/                 # Pasta com os boletos em PDF
├── main.py               # Script principal
├── Sicovimed.xlsx       # Planilha gerada automaticamente
└── README.md            # Documentação do projeto
```

---

## ▶️ Como executar o projeto

### 1. Clone o repositório

```bash
git clone https://github.com/Jordana-Metzler/Leitor-de-PDF.git
cd Leitor-de-PDF
```

---

### 2. Crie um ambiente virtual

```bash
python -m venv .venv
```

Ative o ambiente:

* Windows:

```bash
.venv\Scripts\activate
```

---

### 3. Instale as dependências

```bash
pip install pdfplumber openpyxl
```

---

### 4. Adicione os PDFs

Coloque os arquivos na pasta:

```
pdfs/
```

---

### 5. Execute o script

```bash
python main.py
```

---

## 📊 Saída do sistema

O sistema gera automaticamente:

```
Planilha.xlsx
```

Com as colunas:

| CNPJ | Vencimento | Valor | Número do Documento | Status |
| ---- | ---------- | ----- | ------------------- | ------ |

---

## ⚠️ Observações importantes

* O script foi adaptado para lidar com variações de layout comuns em PDFs
* Regex foi ajustado para trabalhar com:

  * Quebras de linha
  * Texto em maiúsculo/minúsculo
  * Espaços inconsistentes
* Caso novos modelos de boleto sejam utilizados, pode ser necessário ajustar os padrões de extração

---

## 💡 Melhorias futuras

* 📌 Extração detalhada de taxas (água, condomínio, gás, etc.)
* 📌 Agrupamento e soma de taxas duplicadas
* 📌 Integração com sistemas via API/Webservice
* 📌 Interface gráfica para usuários não técnicos
* 📌 Logs estruturados por arquivo processado

---

## 👩‍💻 Autora

Desenvolvido por **Jordana Metzler**

🔗 GitHub: https://github.com/Jordana-Metzler
---

## ⭐ Contribuição

Sinta-se à vontade para contribuir com melhorias ou sugestões!

---

## 📌 Objetivo do projeto

🚧 Aprender ferramentas de extração de dados de pdf, além de regex e envio automático para planilhas.
