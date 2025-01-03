﻿# InvoiceReader
Este projeto é uma ferramenta para extrair dados de faturas (boletos) em formato PDF e salvá-los em uma planilha Excel. Utiliza as bibliotecas pdfplumber e openpyxl para manipulação de PDFs e criação de arquivos Excel, respectivamente. O script é capaz de extrair o número da fatura, data de emissão e status, e armazenar essas informações organizadamente em uma planilha.

Recursos:
- Leitura de arquivos PDF de um diretório especificado
- Extração de informações usando expressões regulares (regex)
- Geração automática de planilhas Excel com os dados extraídos
- Tratamento de exceções para garantir robustez no processamento dos arquivos

Requisitos:
- Python 3.x
- pdfplumber
- openpyxl

Como usar:
1. Coloque seus arquivos PDF de boletos na pasta `pdf_invoices`.
2. Execute o script Python para processar os arquivos e gerar a planilha Excel.

Este projeto é ideal para automatizar o processamento de faturas e facilitar a análise de dados financeiros.

Contribuições são bem-vindas!
