### PROJETO 1 - Automação em Python

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


