# %%
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import pandas as pd
from pathlib import Path


dic_quinzena = {
    '1': 'Primeira Quinzena',
    '2': 'Segunda Quinzena'
}
dic_mes = {
    '01': 'Janeiro',
    '02': 'Fevereiro',
    '03': 'Março',
    '04': 'Abril',
    '05': 'Maio',
    '06': 'Junho',
    '07': 'Julho',
    '08': 'Agosto',
    '09': 'Setembro',
    '10': 'Outubro',
    '11': 'Novembro',
    '12': 'Dezembro'
}

def cria_quinzena(a):
    if a == 0:
        return 0
    else:
        quinzena = str(a)[0]
        mes = str(a)[1:]
        nome_quinzena = f'{dic_quinzena[quinzena]} de {dic_mes[mes]}'
        return nome_quinzena

def find_closest_sku(sku, sku_list):
    matches = get_close_matches(
        sku, sku_list, n=1, cutoff=0.20
    )  # Ajuste o cutoff conforme necessário
    return matches[0] if matches else None


url_valores = "https://docs.google.com/spreadsheets/d/1JkafGyVeOQjCvMmePSfrgW3rSrTs6bzcL01G1Spge4s/export?format=csv&gid=2104680401#gid=2104680401"

df_valores = pd.read_csv(url_valores)
df_valores["MO"] = df_valores["MO"].str.title()

url_inicio = "https://docs.google.com/spreadsheets/d/1cGeQrjvsnuj9K1S_uPrYwxDKnUoyHnQEvFJjttU4Pcw/"
url_fim = "gid=887283048#gid=887283048"
url = f"{url_inicio}export?format=csv&{url_fim}"

# %%
df = pd.read_csv(url)

quinzena_escolhida = 111

df = df[df['Observações'] == quinzena_escolhida]
df = df.drop(columns=['Descrição do produto', 'Destino', '% Completude','Dias na MO', 'Setor de finalização', 'Preço unitário', 'Concluído', 'Observações'])
df = df.rename(columns={
    'Unnamed: 0': 'data', 'Código (SKU) ': 'sku', 'MAXHOME': 'mo',
    'Requisição': 'requisicao', 'Número da nota': 'numero da nota', 'Quantidade': 'quantidade'
})
df['sku'] = df['sku'].fillna('0')
df_valores = df_valores.rename(columns={'SKU': 'sku_mapeado', 'VALOR': 'valor', 'MO': 'mo'})
colunas = ['data', 'numero da nota', 'requisicao', 'sku', 'quantidade', 'mo']
df = df[colunas]
df['sku_mapeado'] = df['sku'].apply(lambda x: find_closest_sku(x, df_valores["sku"]))

# %%

df = df.merge(
    right=df_valores,
    left_on=df['sku_mapeado'],
    right_on=df_valores['sku'],
    suffixes=['_producao', '_valores'],
    how='left'
)


df = df.drop(columns=['sku_mapeado','sku_valores', 'mo_valores'])
df = df.rename(columns={
    'sku_producao': 'sku',
    'mo_producao': 'mo',
})
df
# %%
df['quantidade'] = df['quantidade'].astype(float)
df['valor'] = df['valor'].str.replace(',','.').astype(float)
df['total'] = df['quantidade'] * df['valor']
df
# %%
quinzena = cria_quinzena(str(quinzena_escolhida))
print(quinzena)

# %%
lista_mo = df['mo'].unique()
for mo in lista_mo:
    display(df[df['mo'] == mo])

# %%

# df = df[df['Observações'] == 111]
df_dado = df[df['mo'] == 'DADO']

df_dado
# %%
df_valores_dado = df_valores[df_valores['mo'] == 'Dado']
# %%
df_final = pd.merge(df_dado, df_valores_dado, left_on=['sku_mapeado'], right_on=['sku'], how='left')
df_final
