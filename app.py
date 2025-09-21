import re
import numpy as np
import pandas as pd
import streamlit as st
from io import BytesIO
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Validador de Duplicados", layout="centered")
st.title("🎩✨ Validador de Duplicados")
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


def _detectar_coluna(df, termos):
    """Detecta primeira coluna cujo nome contenha qualquer um dos termos (case-insensitive)."""
    for t in termos:
        for c in df.columns:
            if t in c.lower():
                return c
    return None


def _parse_valor(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    s = s.replace("R$", "").replace("r$", "").replace(" ", "")
    if "," in s:
        s2 = s.replace(".", "").replace(",", ".")
    else:
        s2 = s.replace(",", "")
    s2 = re.sub(r"[^\d\.\-+]", "", s2)
    try:
        return float(s2)
    except:
        return None


def _normalize_cliente(v):
    if pd.isna(v):
        return ""
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip().lower()


def _normalizar_data_valor_cliente(row, date_col, client_col, value_col):
    raw_date = row.get(date_col, None)
    try:
        dt = pd.to_datetime(raw_date, dayfirst=True, errors="coerce")
        d = dt.date() if not pd.isna(dt) else None
    except:
        d = None

    cliente = _normalize_cliente(row.get(client_col, None))
    valor = _parse_valor(row.get(value_col, None))
    return d, cliente, valor


def marcar_duplicados_vermelho(df):
    # Detectar colunas mais prováveis
    date_col = _detectar_coluna(df, ["data", "carimbo", "timestamp", "date"]) or df.columns[0]
    client_col = _detectar_coluna(df, ["cliente", "client", "cod", "codigo"]) or df.columns[1]
    value_col = _detectar_coluna(df, ["valor", "value", "amount", "total"]) or df.columns[2]

    # Criar cópia e calcular duplicados
    df = df.copy()
    df["Duplicado_Linha"] = ""

    primeira_ocorrencia = {}
    for idx, row in df.iterrows():
        d, cliente_norm, valor_num = _normalizar_data_valor_cliente(row, date_col, client_col, value_col)
        key = (d, cliente_norm, None if valor_num is None else round(valor_num, 2))
        if key in primeira_ocorrencia and key[0] and key[2] and key[1]:
            first_idx = primeira_ocorrencia[key]
            df.at[idx, "Duplicado_Linha"] = f"Duplicado da linha {first_idx + 2}"
        else:
            primeira_ocorrencia[key] = idx

    # Exportar temporário
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    # Reabrir com openpyxl para manipular estilo
    wb = load_workbook(output)
    ws = wb.active

    # Inserir a coluna depois de "Conferido"
    col_conferido = None
    for idx, cell in enumerate(ws[1], start=1):
        if str(cell.value).lower().startswith("confer"):
            col_conferido = idx
            break

    if col_conferido is None:
        col_conferido = ws.max_column  # se não achar, adiciona no fim

    ws.insert_cols(col_conferido + 1)
    ws.cell(row=1, column=col_conferido + 1, value="Duplicado_Linha")

    vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # Preencher valores e colorir duplicados
    qtd_dup = 0
    for r in range(2, ws.max_row + 1):
        val = df.iloc[r - 2]["Duplicado_Linha"]
        if val != "":
            ws.cell(row=r, column=col_conferido + 1, value=val)
            qtd_dup += 1
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).fill = vermelho  # preserva bordas/fontes

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

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
        arquivo_final, qtd_dup = marcar_duplicados_vermelho(df)

        if qtd_dup > 0:
            st.success(f"✅ Foram encontradas {qtd_dup} linhas duplicadas (segunda ocorrência em diante).")
        else:
            st.info("Nenhuma linha duplicada encontrada.")

        st.download_button(
            label="📥 Baixar planilha validada",
            data=arquivo_final,
            file_name="planilha_validada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
