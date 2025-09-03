# %%
from difflib import get_close_matches
import pandas as pd
import streamlit as st


def find_closest_sku(sku, sku_list):
    matches = get_close_matches(
        sku, sku_list, n=1, cutoff=0.20
    )  # Ajuste o cutoff conforme necessário
    return matches[0] if matches else None


st.set_page_config(layout="wide")
st.title("Fechamento - Mão de Obra")

url_valores = "https://docs.google.com/spreadsheets/d/1JkafGyVeOQjCvMmePSfrgW3rSrTs6bzcL01G1Spge4s/export?format=csv&gid=2104680401#gid=2104680401"

df_valores = pd.read_csv(url_valores)
df_valores["MO"] = df_valores["MO"].str.title()

url_inicio = "https://docs.google.com/spreadsheets/d/1cGeQrjvsnuj9K1S_uPrYwxDKnUoyHnQEvFJjttU4Pcw/"
url_fim = "gid=887283048#gid=887283048"
url = f"{url_inicio}export?format=csv&{url_fim}"

df = pd.read_csv(url)
df["MAXHOME"] = df["MAXHOME"].str.title()
df = df.drop(
    columns=[
        "Descrição do produto",
        "Destino",
        "% Completude",
        "Dias na MO",
        "Setor de finalização",
        "Preço unitário",
        "Concluído",
    ]
)

df["Observações"] = df["Observações"].fillna(0).astype(int)

# ---------------- FILTRO 1 - Quinzena -------------

filtro1 = df["Observações"] == 208

df = df[filtro1]

# ---------------- FILTRO 2 - Mao de obra -------------
col1, col2, col3, col4 = st.columns(4)

mo_quinzena = df["MAXHOME"].unique().tolist()
mo_quinzena = sorted(mo_quinzena)
mo_selecionada = col1.selectbox(label="Mão de Obra", options=mo_quinzena)

filtro2 = df["MAXHOME"] == mo_selecionada

df = df[filtro2]

filtro3 = df_valores["MO"] == mo_selecionada
df_valores = df_valores[filtro3]

df["Sku_mapeado"] = df["Código (SKU) "].apply(
    lambda x: find_closest_sku(x, df_valores["SKU"])
)

df = df.merge(
    right=df_valores,
    left_on=["Sku_mapeado"],
    right_on=["SKU"],
    how="left",
    suffixes=["Producao", "Valores"],
)

df = df.drop(columns=["MAXHOME", "Observações", "Sku_mapeado", "SKU", "MO"])
df["VALOR"] = df["VALOR"].str.replace(",", ".").astype(float)

df["Total"] = df["Quantidade"] * df["VALOR"]

quantidade_total = df["Quantidade"].sum()

total = df["Total"].sum()
col2.metric(label="Valor", value=f"R$ {total:.2f}")
col3.metric(label="Quantidade entregue", value=quantidade_total)

df.columns = ["Data", "Requisição", "SKU", "Quantidade", "Pedido", "Valor", "Total"]
ordem_colunas = ["Data", "Pedido", "Requisição", "SKU", "Quantidade", "Valor", "Total"]
df = df[ordem_colunas]
df = df.set_index(df["Data"]).drop(columns="Data")

columns_config = {
    "Valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
    "Total": st.column_config.NumberColumn("Total", format="R$ %.2f"),
}
df["Valor"] = df["Valor"].astype(str).str.replace(".", ",")
df["Total"]

df["Total"] = df["Total"].round(2)
df["Total"] = df["Total"].astype(str).str.replace(".", ",")

st.dataframe(df, column_config=columns_config)
