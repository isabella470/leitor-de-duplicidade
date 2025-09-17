from io import BytesIO
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def marcar_duplicados_verde(df):
    # Inicializar coluna de referência
    df["Duplicado_Linha"] = ""
    
    primeira_ocorrencia = {}
    
    # Preencher a coluna Duplicado_Linha
    for idx, row in df.iterrows():
        conteudo = tuple(row.drop("Duplicado_Linha"))
        if conteudo in primeira_ocorrencia:
            df.at[idx, "Duplicado_Linha"] = f"Conteúdo já presente na linha {primeira_ocorrencia[conteudo]+2}" 
        else:
            primeira_ocorrencia[conteudo] = idx

    # Salvar temporário em memória
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    verde = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # duplicadas
    col_dup = df.columns.get_loc("Duplicado_Linha") + 1

    # Pintar apenas linhas que têm comentário na coluna Duplicado_Linha
    for row_idx in range(2, ws.max_row + 1):
        cell_value = ws.cell(row=row_idx, column=col_dup).value
        if cell_value and str(cell_value).strip() != "":
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col).fill = verde

    # Salvar resultado final em memória
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    qtd_dup = (df["Duplicado_Linha"] != "").sum()
    return final_output, qtd_dup
