import pandas as pd
import streamlit as st
from io import BytesIO
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Validador de Duplicados", layout="centered")
st.title("📊 Validador de Duplicados Avançado")
st.write("Suba uma planilha Excel ou informe o link público do Google Sheets para validar duplicados.")

# ---------------- Funções ----------------
def ler_planilha(caminho_ou_link):
    if not isinstance(caminho_ou_link, str):
        return pd.read_excel(caminho_ou_link)
    if caminho_ou_link.startswith("http"):
        if "docs.google.com/spreadsheets" in caminho_ou_link:
            try:
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
        return pd.read_excel(caminho_ou_link)

def marcar_duplicados_avancado(df):
    # Inicializar coluna de referência
    df["Duplicado_Linha"] = ""
    
    # Criar um dicionário para rastrear primeira ocorrência de cada linha
    primeira_ocorrencia = {}
    
    for idx, row in df.iterrows():
        # Converter conteúdo da linha para tupla para poder usar como chave
        conteudo = tuple(row)
        # Remover a coluna Duplicado_Linha da chave
        conteudo = conteudo[:-1]
        
        if conteudo in primeira_ocorrencia:
            # Se já apareceu, indicar a linha original (idx começa em 0)
            df.at[idx, "Duplicado_Linha"] = f"Conteúdo já presente na linha {primeira_ocorrencia[conteudo]+2}" 
            # +2 porque planilha Excel começa em 1 e tem cabeçalho
        else:
            # Primeira ocorrência
            primeira_ocorrencia[conteudo] = idx

    # Salvar temporário
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    col_dup = df.columns.get_loc("Duplicado_Linha") + 1

    # Colorir apenas as linhas duplicadas (segunda ocorrência em diante)
    for row in range(2, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=col_dup).value
        if cell_value != "":
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = amarelo

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    qtd_dup = (df["Duplicado_Linha"] != "").sum()
    return final_output, qtd_dup

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
        arquivo_final, qtd_dup = marcar_duplicados_avancado(df)

        if qtd_dup > 0:
            st.success(f"✅ Foram encontradas {qtd_dup} linhas duplicadas (a partir da segunda ocorrência).")
        else:
            st.info("Nenhuma linha duplicada encontrada.")

        st.download_button(
            label="📥 Baixar planilha validada",
            data=arquivo_final,
            file_name="planilha_validada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
