import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from itertools import combinations
import unicodedata
import os

st.set_page_config(page_title="Revisão de Duplicatas por Subtipo", layout="wide")
st.title("🧠 Triagem de Duplicatas por Subtipo de Produção")

st.markdown("""
Este aplicativo tem como objetivo identificar registros duplicados em planilhas de produções, com base na similaridade dos campos **Título** e **Autor(es)**.
Serão considerados como duplicatas os pares com **≥ 85% de similaridade** nos dois campos.
""")

# Função para normalizar texto
def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = unicodedata.normalize('NFKD', text)
    return "".join([c for c in text if not unicodedata.combining(c)]).strip()

# Pasta de saída
output_dir = os.path.join(os.getcwd(), "outputs")
os.makedirs(output_dir, exist_ok=True)

# Upload do arquivo principal
uploaded_file = st.file_uploader("📄 Faça upload da planilha (.xlsx)", type=["xlsx"])

if uploaded_file:
    if "iniciado" not in st.session_state:
        st.session_state["iniciado"] = False

    if not st.session_state["iniciado"]:
        limpar = st.checkbox("🧹 Limpar resultados anteriores (outputs) antes de iniciar nova revisão")
        st.markdown("""
💡 Recomendado apenas se você deseja **reiniciar a revisão do zero**. Caso contrário, o sistema irá **continuar de onde parou** com base nos arquivos salvos na pasta `outputs`.
""")

        if st.button("➡️ Prosseguir para identificação dos subtipos"):
            if limpar:
                for f in os.listdir(output_dir):
                    os.remove(os.path.join(output_dir, f))
                st.success("✅ Pasta outputs limpa. Agora você pode iniciar uma nova revisão.")
            st.session_state["iniciado"] = True
            st.rerun()

    if st.session_state["iniciado"]:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.replace("'", "").str.strip()

        required_cols = ['Título', 'AUTOR(ES)', 'NM_SUBTIPO_PRODUCAO']
        if not all(col in df.columns for col in required_cols):
            st.error(f"A planilha deve conter as colunas obrigatórias: {required_cols}")
        else:
            df["titulo_norm"] = df["Título"].apply(normalize_text)
            df["autor_norm"] = df["AUTOR(ES)"].apply(normalize_text)

            subtipos_info = df["NM_SUBTIPO_PRODUCAO"].value_counts().reset_index()
            subtipos_info.columns = ["Subtipo", "Registros"]

            st.markdown("### 📊 Subtipos encontrados na planilha:")
            st.dataframe(subtipos_info)

            st.markdown(f"""
🗭 **Como funciona este processo:**

- O sistema irá processar os subtipos **um por vez**, na ordem em que aparecem.
- Ao final da revisão de cada subtipo:
  - Uma planilha com os registros limpos será **salva automaticamente**.
  - Uma planilha com os pares comparados será salva para **auditoria**.
- Ao concluir todos os subtipos, você poderá **consolidar os resultados**.

📁 Os arquivos serão salvos automaticamente na pasta: `{output_dir}`

⚠️ **Importante:** Guarde os arquivos `limpo_*.xlsx` e `auditoria_*.xlsx` gerados para consolidar tudo ao final.
            """)

            processados = [f.replace("limpo_", "").replace(".xlsx", "") for f in os.listdir(output_dir) if f.startswith("limpo_")]
            proximo_subtipo = None
            for subtipo in subtipos_info["Subtipo"]:
                nome_padrao = subtipo.replace(" ", "_").replace("/", "_")
                if nome_padrao not in processados:
                    proximo_subtipo = subtipo
                    break

            if not proximo_subtipo:
                st.success("🎉 Todos os subtipos foram processados!")
                st.info(f"📂 Os arquivos foram salvos na pasta: `{output_dir}`.")

                if st.button("📅 Consolidar planilhas finais"):
                    limpos = [f for f in os.listdir(output_dir) if f.startswith("limpo_")]
                    auditorias = [f for f in os.listdir(output_dir) if f.startswith("auditoria_")]

                    df_consolidados = pd.concat(
                        [pd.read_excel(os.path.join(output_dir, f)) for f in limpos],
                        ignore_index=True
                    ) if limpos else pd.DataFrame()

                    df_auditoria = pd.concat(
                        [pd.read_excel(os.path.join(output_dir, f)) for f in auditorias],
                        ignore_index=True
                    ) if auditorias else pd.DataFrame()

                    path_final = os.path.join(output_dir, "consolidado_completo.xlsx")
                    path_auditoria = os.path.join(output_dir, "consolidado_auditoria.xlsx")

                    df_consolidados.to_excel(path_final, index=False)
                    df_auditoria.to_excel(path_auditoria, index=False)

                    st.success("✅ Consolidação concluída!")
                    with open(path_final, "rb") as f1:
                        st.download_button("⬇️ Baixar planilha final consolidada", f1, "consolidado_completo.xlsx")

                    with open(path_auditoria, "rb") as f2:
                        st.download_button("⬇️ Baixar planilha de auditoria consolidada", f2, "consolidado_auditoria.xlsx")

                total_registros = len(df)
                total_processados = sum(
                    pd.read_excel(os.path.join(output_dir, f)).shape[0]
                    for f in os.listdir(output_dir) if f.startswith("limpo_")
                )
                total_removidos = total_registros - total_processados

                st.markdown("### 📈 Estatísticas Gerais")
                st.write(f"🔢 Total de registros originais: {total_registros}")
                st.write(f"✅ Total de registros mantidos: {total_processados}")
                st.write(f"❌ Total de registros removidos como duplicatas: {total_removidos}")

            else:
                st.header(f"📌 Próximo subtipo a ser analisado: **{proximo_subtipo}**")
                nome_sub = proximo_subtipo.replace(" ", "_").replace("/", "_")

                if f"analise_{nome_sub}" not in st.session_state:
                    st.session_state[f"analise_{nome_sub}"] = False

                if not st.session_state[f"analise_{nome_sub}"]:
                    if st.button("🚀 Iniciar análise deste subtipo"):
                        st.session_state[f"analise_{nome_sub}"] = True
                        st.rerun()

                elif st.session_state[f"analise_{nome_sub}"]:
                    df_sub = df[df["NM_SUBTIPO_PRODUCAO"] == proximo_subtipo].reset_index(drop=True)

                    duplicatas = []
                    for idx1, idx2 in combinations(df_sub.index, 2):
                        titulo_sim = fuzz.token_set_ratio(df_sub.loc[idx1, "titulo_norm"], df_sub.loc[idx2, "titulo_norm"])
                        autor_sim = fuzz.token_set_ratio(df_sub.loc[idx1, "autor_norm"], df_sub.loc[idx2, "autor_norm"])
                        if titulo_sim >= 85 and autor_sim >= 85:
                            duplicatas.append({
                                "idx1": idx1,
                                "idx2": idx2,
                                "linha_excel_1": idx1 + 2,
                                "linha_excel_2": idx2 + 2,
                                "titulo_1": df_sub.loc[idx1, "Título"],
                                "titulo_2": df_sub.loc[idx2, "Título"],
                                "autor_1": df_sub.loc[idx1, "AUTOR(ES)"],
                                "autor_2": df_sub.loc[idx2, "AUTOR(ES)"],
                                "sim_titulo": titulo_sim,
                                "sim_autor": autor_sim,
                            })

                    if not duplicatas:
                        st.info("✅ Nenhuma duplicata encontrada para este subtipo.")
                        df_sub.drop(columns=["titulo_norm", "autor_norm"]).to_excel(os.path.join(output_dir, f"limpo_{nome_sub}.xlsx"), index=False)
                        pd.DataFrame().to_excel(os.path.join(output_dir, f"auditoria_{nome_sub}.xlsx"), index=False)
                        st.success(f"Subtipo '{proximo_subtipo}' salvo sem duplicatas detectadas.")
                        st.info(f"📁 Arquivos salvos em: `{output_dir}`")
                        st.session_state[f"analise_{nome_sub}"] = False
                        st.rerun()
                    else:
                        st.success(f"{len(duplicatas)} pares encontrados. Faça a seleção abaixo:")

                        historico = []
                        indices_remover = set()

                        for i, dup in enumerate(duplicatas):
                            st.markdown(f"#### 🔁 Par {i+1}")
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write(f"🅰 **Linha {dup['linha_excel_1']}**")
                                st.write(f"📘 {dup['titulo_1']}")
                                st.write(f"👤 {dup['autor_1']}")
                            with col2:
                                st.write(f"🅱 **Linha {dup['linha_excel_2']}**")
                                st.write(f"📘 {dup['titulo_2']}")
                                st.write(f"👤 {dup['autor_2']}")

                            st.radio(
                                f"Escolha para o par {i+1}",
                                [f"🅰 Manter A (remover B)",
                                 f"🅱 Manter B (remover A)",
                                 "✅ Manter ambos"],
                                key=f"escolha_{nome_sub}_{i}",
                                index=None
                            )

                        respostas_dadas = all(
                            st.session_state.get(f"escolha_{nome_sub}_{i}") is not None
                            for i in range(len(duplicatas))
                        )

                        if not respostas_dadas:
                            st.warning("⚠️ Responda todos os pares antes de continuar.")
                        else:
                            if st.button("➡️ Próximo subtipo"):
                                for i, dup in enumerate(duplicatas):
                                    esc = st.session_state.get(f"escolha_{nome_sub}_{i}")
                                    if esc:
                                        if "remover B" in esc:
                                            indices_remover.add(dup["idx2"])
                                            acao = "Manter A"
                                        elif "remover A" in esc:
                                            indices_remover.add(dup["idx1"])
                                            acao = "Manter B"
                                        else:
                                            acao = "Manter ambos"
                                        historico.append({**dup, "decisao": acao})

                                df_limpo = df_sub.drop(list(indices_remover)).drop(columns=["titulo_norm", "autor_norm"])
                                df_limpo.to_excel(os.path.join(output_dir, f"limpo_{nome_sub}.xlsx"), index=False)
                                pd.DataFrame(historico).to_excel(os.path.join(output_dir, f"auditoria_{nome_sub}.xlsx"), index=False)

                                st.success(f"Subtipo '{proximo_subtipo}' salvo com decisões aplicadas.")
                                st.info(f"📁 Arquivos salvos em: `{output_dir}`")
                                st.session_state[f"analise_{nome_sub}"] = False
                                st.rerun()