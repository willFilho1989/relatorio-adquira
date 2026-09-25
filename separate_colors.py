import openpyxl
from copy import copy

# Carregar arquivo original
wb_source = openpyxl.load_workbook('bases\\Adquira.xlsx')
ws_source = wb_source.active

# Separar linhas por cor
gray_rows = []
blue_rows = []

max_row = ws_source.max_row
max_col = ws_source.max_column

print(f"Total de linhas: {max_row}, Colunas: {max_col}")

for row_idx in range(1, max_row + 1):
    row_data = []
    first_cell_color = None
    
    for col_idx in range(1, max_col + 1):
        cell = ws_source.cell(row=row_idx, column=col_idx)
        row_data.append(cell)
        
        if col_idx == 1:
            fill = cell.fill
            if fill and fill.start_color:
                first_cell_color = fill.start_color.rgb
    
    # Classificar por cor
    if first_cell_color:
        color_str = str(first_cell_color)
        print(f"Linha {row_idx}: Cor={color_str}")
        if 'D0CECE' in color_str or 'D0CECE' in color_str or '99' in color_str:
            gray_rows.append(row_data)
        else:
            blue_rows.append(row_data)
    else:
        blue_rows.append(row_data)

print(f"\nLinhas cinza: {len(gray_rows)}")
print(f"Linhas azul: {len(blue_rows)}")

# Criar novo workbook
wb_new = openpyxl.Workbook()
wb_new.remove(wb_new.active)

def copy_cell(source_cell, target_cell):
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

# Aba para cinza
ws_gray = wb_new.create_sheet('Pedidos (Cinza)')
for row_idx, row_data in enumerate(gray_rows, 1):
    for col_idx, source_cell in enumerate(row_data, 1):
        target_cell = ws_gray.cell(row=row_idx, column=col_idx)
        copy_cell(source_cell, target_cell)

# Aba para azul
ws_blue = wb_new.create_sheet('Itens (Azul)')
for row_idx, row_data in enumerate(blue_rows, 1):
    for col_idx, source_cell in enumerate(row_data, 1):
        target_cell = ws_blue.cell(row=row_idx, column=col_idx)
        copy_cell(source_cell, target_cell)

# Ajustar largura das colunas
for ws in [ws_gray, ws_blue]:
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

wb_new.save('bases\\Relatorio_Separado.xlsx')
print("\n✓ Arquivo criado com sucesso: bases\\Relatorio_Separado.xlsx")
