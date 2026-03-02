# 📊 Mini-Projeto - Automação em Python para Atualização de Base no Power BI

## 📌 Sobre o Projeto

Este projeto tem como objetivo automatizar o processo de atualização de uma base de dados utilizada em um dashboard no Power BI.

A automação foi desenvolvida em Python com foco em:

- Organização de projeto
- Manipulação de arquivos Excel
- Comparação de colunas entre planilhas
- Atualização automática de base de dados
- Criação de estrutura profissional de pastas

O script elimina o processo manual de copiar e colar dados, tornando a atualização mais rápida, segura e eficiente.



## 🚀 O Que o Projeto Faz

O script realiza as seguintes etapas:

1. Identifica automaticamente o diretório do projeto.
2. Cria as pastas necessárias (`dados` e `arquivoautomatizado`).
3. Lê duas planilhas Excel:
   - Base semanal de vendas
   - Base utilizada pelo Power BI
4. Compara as colunas das duas planilhas.
5. Mantém apenas as colunas que existem em ambas.
6. Atualiza automaticamente a base do Power BI.
7. Gera uma cópia do arquivo atualizado na pasta de saída.



## 🛠 Tecnologias Utilizadas

- Python
- Pandas
- Pathlib
- OpenPyXL
- Shutil



## 📂 Estrutura do Projeto
automacao_powerbi/
│
├── dados/
│ ├── base_semanal_vendas.xlsx
│ ├── powerbi_vendas.xlsx
│
├── arquivoautomatizado/
│
├── scriptautomacao.py
└── README.md




## ⚙️ Como Funciona a Lógica

- As duas planilhas são carregadas como DataFrames.
- O código identifica as colunas em comum entre elas.
- A base principal é filtrada para manter apenas essas colunas.
- O arquivo do Power BI é reescrito automaticamente com os novos dados.
- Uma versão atualizada é salva na pasta de saída.


CÓDIGO:

import pandas as pd
from pathlib import Path
import shutil

def executar():
    base_projeto = Path(__file__).parent
    print(base_projeto)
    pasta_dados = base_projeto / "dados"
    pasta_saida = base_projeto / "arquivoautomatizado"

    pasta_dados.mkdir(exist_ok=True,parents=True)
    pasta_saida.mkdir(exist_ok=True,parents=True)

    arquivo_base = pasta_dados / "base_semanal_vendas.xlsx"
    arquivo_powerbi = pasta_dados / "powerbi_vendas.xlsx"

    dados_base = pd.read_excel(arquivo_base,sheet_name="Base_Vendas")
    dados_powerbi = pd.read_excel(arquivo_powerbi,sheet_name="Base_Vendas")

    print("Base vendas carregada com sucesso: ")
    print(dados_base.head())

    print("Base Power Bi carregada com sucesso: ")
    print(dados_powerbi.head())

    print("Colunas da base de vendas: ")
    print(dados_base.columns)

    print("colunas da base power bi :")
    print(dados_powerbi.columns)

    colunas_iguais = dados_base.columns.intersection(dados_powerbi.columns)
    
    print("Colunas em comum: ")
    print(colunas_iguais)

    dados_filtrados = dados_base[colunas_iguais]

    with pd.ExcelWriter(arquivo_powerbi,engine="openpyxl",mode="w") as writer:
        dados_filtrados.to_excel(writer,sheet_name="Base_Vendas",index=False)

    print("Base Power Bi Atualizada com Sucesso")

    caminho_origem = arquivo_powerbi
    caminho_destino  = pasta_saida / "powerbi_vendas_atualizado.xlsx"

    shutil.copy(caminho_origem, caminho_destino)

if __name__ == "__main__":
    executar()
