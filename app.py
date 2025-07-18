# %%
import pandas as pd
import streamlit as st

st.set_page_config(layout="wide")

st.title(f"Fechamento - Mão de Obra")

url_inicio = "https://docs.google.com/spreadsheets/d/1cGeQrjvsnuj9K1S_uPrYwxDKnUoyHnQEvFJjttU4Pcw/"
url_fim = "gid=887283048#gid=887283048"
url = f"{url_inicio}export?format=csv&{url_fim}"

df = pd.read_csv(url)
df["MAXHOME"] = df["MAXHOME"].str.title()
df["Observações"] = df["Observações"].fillna(0).astype(int)
df["Observações"] = df["Observações"].astype(int)

df_selecionado = df[df["Observações"] == 107]


# --------------------- filtro para quinzena ----------------
# opcoes = df["Observações"].unique()
# opcao_selecionada = st.sidebar.selectbox(label="Quinzena", options=opcoes)
# df_selecionado = df[df["Observações"] == 'opcao_selecionada']
# --------------------- filtro para quinzena ----------------

# --------------------- filtro para mao de obra ----------------
col1, col2, col3, col4, col5, col6, col7 = st.columns(7)

mo_quinzena = df_selecionado["MAXHOME"].unique().tolist()
mo_quinzena = sorted(mo_quinzena)

mo_selecionada = col1.selectbox(label="Mão de Obra", options=mo_quinzena)

df_final = df_selecionado[df_selecionado["MAXHOME"] == mo_selecionada]
# --------------------- filtro para mao de obra ----------------
# mo_selecionada = "Maxhome"

df_final = df_selecionado[df_selecionado["MAXHOME"] == mo_selecionada]

df_final = df_final.drop(
    columns=[
        "Descrição do produto",
        "MAXHOME",
        "Destino",
        "% Completude",
        "Dias na MO",
        "Setor de finalização",
        "Preço unitário",
        "Concluído",
        "Observações",
    ]
)

df_final = df_final.set_index("Data da notinha")

total = df_final["Quantidade"].sum()


st.text(f"Quantidade total produzida: {total}")

df_final = df_final[["Número da nota", "Requisição", "Código (SKU) ", "Quantidade"]]

st.dataframe(df_final)
