import pandas as pd
from core.preenchedor import preencher_campos

# Substitua por um caminho real se necessário
caminho_planilha = "D:\Projetos AO Arquiteto\Cód.72 - Raquel e Moares - Avaliação Imobiliária\Avaliação por Inferência estatística\Infer-32\base_de_dados_210625.xlsx"
dados = pd.read_excel(caminho_planilha)

preencher_campos(dados)