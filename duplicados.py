import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from itertools import combinations
import unicodedata
import io

st.set_page_config(page_title="Identificação de Duplicatas", layout="wide")
st.title("🧠 Identificação Automática de Duplicatas")

st.markdown("""
Este aplicativo identifica registros duplicados em planilhas de produções com base na similaridade dos campos **NM_PRODUCAO** (título) e **AUTOR(ES)**.  
Produções duplicadas receberão:
- `"SIM"` na coluna **Produção duplicada**
- O **ID principal** na coluna **ID_UNIFICADO**
- Os IDs dos pares duplicados nas colunas **ID_VEICULO1** e **ID_VEICULO2**
""")

def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = unicodedata.normalize('NFKD', text)
    return "".join([c for c in text if not unicodedata.combining(c)]).strip()

uploaded_file = st.file_uploader("📄 Faça upload da planilha (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    required_cols = ['NM_PRODUCAO', 'AUTOR(ES)', 'NM_SUBTIPO_PRODUCAO', 'ID_ADD_PRODUCAO_INTELECTUAL']
    if not all(col in df.columns for col in required_cols):
        st.error(f"A planilha deve conter as colunas: {required_cols}")
    else:
        df["titulo_norm"] = df["NM_PRODUCAO"].apply(normalize_text)
        df["autor_norm"] = df["AUTOR(ES)"].apply(normalize_text)
        df["Produção duplicada"] = "NÃO"
        df["ID_UNIFICADO"] = ""
        df["ID_VEICULO1"] = ""
        df["ID_VEICULO2"] = ""

        for subtipo in df["NM_SUBTIPO_PRODUCAO"].unique():
            sub_df = df[df["NM_SUBTIPO_PRODUCAO"] == subtipo]
            indices = sub_df.index.tolist()

            for idx1, idx2 in combinations(indices, 2):
                titulo_sim = fuzz.token_set_ratio(df.at[idx1, "titulo_norm"], df.at[idx2, "titulo_norm"])
                autor_sim = fuzz.token_set_ratio(df.at[idx1, "autor_norm"], df.at[idx2, "autor_norm"])

                if titulo_sim >= 85 and autor_sim >= 85:
                    idx_keep = min(idx1, idx2)
                    idx_mark = max(idx1, idx2)

                    df.at[idx_mark, "Produção duplicada"] = "SIM"
                    df.at[idx_mark, "ID_UNIFICADO"] = df.at[idx_keep, "ID_ADD_PRODUCAO_INTELECTUAL"]
                    df.at[idx_mark, "ID_VEICULO1"] = df.at[idx1, "ID_ADD_PRODUCAO_INTELECTUAL"]
                    df.at[idx_mark, "ID_VEICULO2"] = df.at[idx2, "ID_ADD_PRODUCAO_INTELECTUAL"]

        df.drop(columns=["titulo_norm", "autor_norm"], inplace=True)

        st.success("✅ Análise concluída. Baixe a planilha com os resultados.")
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)

        st.download_button(
            "⬇️ Baixar planilha com marcação de duplicatas",
            data=buffer,
            file_name="duplicatas_marcadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
