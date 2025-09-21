import pandas as pd
import streamlit as st
from io import BytesIO
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Validador de Duplicados", layout="centered")
st.title("🎩✨ Validador de Duplicados")
st.write("Suba uma planilha Excel ou CSV, ou informe o link público do Google Sheets para validar duplicados.")

# ---------------- Funções ----------------
def ler_planilha(caminho_ou_link):
    try:
        if hasattr(caminho_ou_link, 'name') and caminho_ou_link.name.endswith('.csv'):
            df = pd.read_csv(caminho_ou_link)
        elif not isinstance(caminho_ou_link, str):
            df = pd.read_excel(caminho_ou_link)
        elif caminho_ou_link.startswith("http"):
            if "docs.google.com/spreadsheets" in caminho_ou_link:
                sheet_id = caminho_ou_link.split("/d/")[1].split("/")[0]
                export_link = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
                resp = requests.get(export_link)
                if resp.status_code == 200:
                    df = pd.read_excel(BytesIO(resp.content))
                else:
                    st.error(f"❌ Erro ao acessar o link. Status: {resp.status_code}")
                    return None
            else:
                st.error("❌ O link não parece ser do Google Sheets.")
                return None
        else:
            df = pd.read_excel(caminho_ou_link)
        
        # Limpa os nomes das colunas de espaços em branco
        df.columns = df.columns.str.strip()
        return df

    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo: {e}")
        return None


def marcar_duplicados_vermelho(df):
    # Colunas para verificar duplicados
    colunas_para_verificar = ["Cliente:", "Valor:", "Carimbo de data/hora"]

    # Verificar se as colunas existem no DataFrame
    for col in colunas_para_verificar:
        if col not in df.columns:
            st.error(f"A coluna '{col}' não foi encontrada na planilha. Verifique os nomes das colunas.")
            return None, 0

    # Prepara as colunas para a comparação, evitando erros de tipo de dado
    df_temp = df.copy()
    df_temp['Valor:'] = pd.to_numeric(df_temp['Valor:'], errors='coerce')
    df_temp['Carimbo de data/hora'] = df_temp['Carimbo de data/hora'].astype(str)
    df_temp['Cliente:'] = df_temp['Cliente:'].astype(str)


    # Marcar todas as ocorrências de duplicados
    duplicados = df_temp.duplicated(subset=colunas_para_verificar, keep=False)
    df["Status"] = ""
    df.loc[duplicados, "Status"] = "Duplicada"

    # Salvar em um arquivo Excel em memória
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    
    # Encontra a coluna "Status" pelo nome
    status_col_index = None
    for i, cell in enumerate(ws[1]):
        if cell.value == "Status":
            status_col_index = i + 1
            break
            
    if status_col_index is None:
        st.error("Não foi possível encontrar a coluna 'Status' para aplicar a formatação.")
        return None, 0

    # Pinta de vermelho as linhas marcadas como "Duplicada"
    for row_idx in range(2, ws.max_row + 1):
        cell_value = ws.cell(row=row_idx, column=status_col_index).value
        if cell_value == "Duplicada":
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = vermelho

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    qtd_dup = duplicados.sum()
    return final_output, qtd_dup

# ---------------- Interface ----------------
tab1, tab2 = st.tabs(["📂 Upload de Arquivo", "🔗 Link Google Sheets"])
df = None

with tab1:
    uploaded_file = st.file_uploader("Selecione um arquivo Excel ou CSV", type=["xlsx", "csv"])
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
        arquivo_final, qtd_dup = marcar_duplicados_vermelho(df)

        if qtd_dup > 0:
            st.success(f"✅ Foram encontradas {qtd_dup} linhas duplicadas (todas as ocorrências foram marcadas).")
        else:
            st.info("🎉 Nenhuma linha duplicada encontrada.")

        if arquivo_final:
            st.download_button(
                label="📥 Baixar planilha validada",
                data=arquivo_final,
                file_name="planilha_validada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
