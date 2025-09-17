import pandas as pd
import streamlit as st
from io import BytesIO
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Validador de Duplicados", layout="centered")
st.title("📊 Validador de Duplicados")
st.write("Suba uma planilha Excel ou informe o link público do Google Sheets para validar duplicados.")

# ---------------- Funções ----------------
def ler_planilha(caminho_ou_link):
    # Caso seja arquivo enviado via upload
    if not isinstance(caminho_ou_link, str):
        return pd.read_excel(caminho_ou_link)

    # Caso seja link do Google Sheets
    if caminho_ou_link.startswith("http"):
        if "docs.google.com/spreadsheets" in caminho_ou_link:
            try:
                # Extrair ID e montar link de exportação
                sheet_id = caminho_ou_link.split("/d/")[1].split("/")[0]
                export_link = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
                resp = requests.get(export_link)
                if resp.status_code == 200:
                    return pd.read_excel(BytesIO(resp.content))
                else:
                    st.error(f"❌ Erro ao acessar o link. Status: {resp.status_code}")
                    return None
            except Exception as e:
                st.error(f"❌ Não foi possível processar o link: {e}")
                return None
        else:
            st.error("❌ O link não parece ser do Google Sheets.")
            return None
    else:
        # Caminho local
        return pd.read_excel(caminho_ou_link)


def marcar_duplicados(df):
    df["Duplicado"] = df.duplicated(keep=False)

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    col_dup = df.columns.get_loc("Duplicado") + 1

    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=col_dup).value == True:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = amarelo

    ws.delete_cols(col_dup)

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output, df["Duplicado"].sum()

# ---------------- Interface ----------------
tab1, tab2 = st.tabs(["📂 Upload Excel", "🔗 Link Google Sheets"])

df = None

with tab1:
    uploaded_file = st.file_uploader("Selecione um arquivo Excel", type=["xlsx"])
    if uploaded_file is not None:
        df = ler_planilha(uploaded_file)

with tab2:
    link = st.text_input("Cole o link público do Google Sheets:")
    if link:
        df = ler_planilha(link)

if df is not None:
    st.subheader("📑 Pré-visualização dos dados")
    st.dataframe(df.head())

    if st.button("🔎 Validar Duplicados"):
        arquivo_final, qtd_dup = marcar_duplicados(df)

        if qtd_dup > 0:
            st.success(f"✅ Foram encontradas {qtd_dup} linhas duplicadas.")
        else:
            st.info("Nenhuma linha duplicada encontrada.")

        st.download_button(
            label="📥 Baixar planilha validada",
            data=arquivo_final,
            file_name="planilha_validada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
