from copy import copy
import csv
import math
from pathlib import Path
import re
import sys

import openpyxl
from openpyxl import load_workbook


CASA_CLIENTE_PATH_HINT = Path(r"C:\Users\WILLIAM.FILHO\Desktop\Faturamento\Vivo_Casa_Cliente")
IMPLANTACAO_PATH_HINT = Path(r"C:\Users\WILLIAM.FILHO\Desktop\Faturamento\Vivo_Implantação")


def resolve_bases_dir():
    script_dir = Path(__file__).resolve().parent
    candidates = [Path.cwd() / "bases", script_dir / "bases"]

    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates = [exe_dir / "bases", exe_dir.parent / "bases", Path.cwd() / "bases"] + candidates

    for path in candidates:
        if (path / "Adquira.xlsx").exists() and (path / "Demanda.xlsx").exists():
            return path

    checked = "\n".join(str(p) for p in candidates)
    raise FileNotFoundError(
        "Não encontrei a pasta 'bases' com Adquira.xlsx e Demanda.xlsx.\nLocais verificados:\n" + checked
    )


BASES_DIR = resolve_bases_dir()
SOURCE_PATH = BASES_DIR / "Adquira.xlsx"
DEMANDA_PATH = BASES_DIR / "Demanda.xlsx"
OUTPUT_PATH = BASES_DIR / "Relatorio_Separado.xlsx"


def color_key(cell):
    fill = cell.fill
    if fill is None:
        return ("none", None)
    color = fill.fgColor or fill.start_color
    if color is None:
        return ("none", None)
    ctype = color.type or "none"
    if ctype == "indexed":
        return ("indexed", color.indexed)
    if ctype == "theme":
        return ("theme", (color.theme, color.tint))
    return ("rgb", color.rgb)


def copy_cell(source_cell, target_cell):
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)


def parse_brl_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = text.replace("R$", "").replace(" ", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def resolve_tabular_paths(path_hint: Path):
    supported_ext = {".xlsx", ".xlsm", ".csv"}
    candidates = [path_hint]
    if not path_hint.suffix:
        for ext in (".xlsx", ".xlsm", ".csv"):
            candidates.append(path_hint.with_suffix(ext))
    if not path_hint.exists() and path_hint.parent.exists():
        for match in path_hint.parent.glob(path_hint.name + "*"):
            if match.suffix.lower() in supported_ext or match.is_dir():
                candidates.append(match)

    found_files = []
    for candidate in candidates:
        if not candidate.exists():
            continue
        if candidate.is_file() and candidate.suffix.lower() in supported_ext and not candidate.name.startswith("~$"):
            found_files.append(candidate)
            continue
        if candidate.is_dir():
            direct_files = [
                file_path
                for file_path in candidate.iterdir()
                if (
                    file_path.is_file()
                    and file_path.suffix.lower() in supported_ext
                    and not file_path.name.startswith("~$")
                )
            ]
            if direct_files:
                found_files.extend(direct_files)
                continue

            recursive_files = []
            for file_path in candidate.rglob("*"):
                if (
                    file_path.is_file()
                    and file_path.suffix.lower() in supported_ext
                    and not file_path.name.startswith("~$")
                ):
                    recursive_files.append(file_path)
            recursive_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            found_files.extend(recursive_files[:30])

    if found_files:
        unique_sorted = sorted(set(found_files), key=lambda p: str(p).lower())
        return unique_sorted

    checked = "\n".join(str(p) for p in candidates)
    raise FileNotFoundError(
        f"Não encontrei a base em '{path_hint}'.\nArquivos verificados:\n{checked}"
    )


def normalize_header(value):
    return re.sub(r"\s+", " ", str(value or "").strip().upper())


def normalize_order_id(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return None
        if value.is_integer():
            return str(int(value))
        text = format(value, ".15g")
        return text.replace(",", ".").strip().upper()

    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"\s+", "", text).upper()
    if re.fullmatch(r"[+-]?\d+([.,]\d+)?", text):
        text = text.replace(",", ".")
        try:
            as_float = float(text)
            if as_float.is_integer():
                return str(int(as_float))
        except ValueError:
            pass
    return text


def find_column_index(ws, expected_header):
    expected = normalize_header(expected_header)
    max_scan_rows = min(5, ws.max_row)
    for row in range(1, max_scan_rows + 1):
        for col in range(1, ws.max_column + 1):
            header = normalize_header(ws.cell(row=row, column=col).value)
            if header == expected:
                return row, col
    return None, None


def load_ids_from_excel(base_path: Path, column_name: str):
    wb = load_workbook(base_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        header_row, col_idx = find_column_index(ws, column_name)
        if col_idx is None:
            return set(), False

        ids = set()
        for row in range(header_row + 1, ws.max_row + 1):
            normalized = normalize_order_id(ws.cell(row=row, column=col_idx).value)
            if normalized:
                ids.add(normalized)
        return ids, True
    finally:
        wb.close()


def load_ids_from_csv(base_path: Path, column_name: str):
    ids = set()
    found_column = False
    with base_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        if reader.fieldnames:
            normalized_headers = {normalize_header(h): h for h in reader.fieldnames}
            col_name = normalized_headers.get(normalize_header(column_name))
            if col_name:
                found_column = True
                for row in reader:
                    normalized = normalize_order_id(row.get(col_name))
                    if normalized:
                        ids.add(normalized)
    return ids, found_column


def load_ids_from_base(path_hint: Path, column_name: str):
    base_paths = resolve_tabular_paths(path_hint)
    ids = set()
    found_in_any_file = False
    for base_path in base_paths:
        suffix = base_path.suffix.lower()
        if suffix == ".csv":
            file_ids, found_column = load_ids_from_csv(base_path, column_name)
        else:
            file_ids, found_column = load_ids_from_excel(base_path, column_name)
        if found_column:
            found_in_any_file = True
            ids.update(file_ids)

    if not found_in_any_file:
        raise ValueError(
            f"Coluna '{column_name}' não encontrada em nenhum arquivo da base '{path_hint}'."
        )
    return ids


def create_din_with_excel(file_path: Path):
    import pythoncom
    import win32com.client

    xl_database = 1
    xl_pivot_table_version15 = 6
    xl_page_field = 3
    xl_row_field = 1
    xl_sum = -4157
    xl_tabular_row = 1

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    wb = None
    original_error = None
    try:
        try:
            excel.Visible = False
        except Exception:
            pass
        try:
            excel.DisplayAlerts = False
        except Exception:
            pass

        wb = excel.Workbooks.Open(str(file_path))
        data_ws = wb.Worksheets("Verde")

        for ws in wb.Worksheets:
            if ws.Name == "DIN":
                ws.Delete()
                break

        din_ws = wb.Worksheets.Add()
        din_ws.Name = "DIN"

        used = data_ws.UsedRange
        rows = used.Rows.Count
        cols = used.Columns.Count
        src = f"'{data_ws.Name}'!R1C1:R{rows}C{cols}"

        pc1 = wb.PivotCaches().Create(xl_database, src, xl_pivot_table_version15)
        pt1 = pc1.CreatePivotTable(din_ws.Range("A3"), "PT_UF")
        pt1.PivotFields("FRENTE").Orientation = xl_page_field
        pt1.PivotFields("UF").Orientation = xl_row_field
        df1 = pt1.AddDataField(pt1.PivotFields("TRIBUTÁVEL"), "Soma de TRIBUTÁVEL", xl_sum)

        pc2 = wb.PivotCaches().Create(xl_database, src, xl_pivot_table_version15)
        pt2 = pc2.CreatePivotTable(din_ws.Range("F3"), "PT_PEDIDO_UF")
        pt2.PivotFields("FRENTE").Orientation = xl_page_field
        pt2.PivotFields("ID. PEDIDO COMPRADOR").Orientation = xl_row_field
        pt2.PivotFields("UF").Orientation = xl_row_field
        pt2.PivotFields("UF").Position = 2
        df2 = pt2.AddDataField(pt2.PivotFields("TRIBUTÁVEL"), "Soma de TRIBUTÁVEL", xl_sum)

        pt2.RowAxisLayout(xl_tabular_row)
        no_subtotals = (False,) * 12
        pt2.PivotFields("ID. PEDIDO COMPRADOR").Subtotals = no_subtotals
        pt2.PivotFields("UF").Subtotals = no_subtotals
        pt2.RowGrand = False
        pt2.ColumnGrand = False

        # Formato final exibido em PT-BR: R$ 3.889,17
        for df in (df1, df2):
            df.NumberFormat = "R$ #,##0.00"

        if pt1.DataBodyRange is not None:
            pt1.DataBodyRange.NumberFormat = "R$ #,##0.00"
        if pt2.DataBodyRange is not None:
            pt2.DataBodyRange.NumberFormat = "R$ #,##0.00"

        # Garante formato local diretamente nas colunas de valor da aba DIN
        # (pivot 1 em B, pivot 2 em H no layout atual)
        din_ws.Columns("B:B").NumberFormatLocal = "R$ #.##0,00"
        din_ws.Columns("H:H").NumberFormatLocal = "R$ #.##0,00"

        # Reaplica após refresh para evitar retorno de 5 casas
        pt1.RefreshTable()
        pt2.RefreshTable()
        if pt1.DataBodyRange is not None:
            pt1.DataBodyRange.NumberFormatLocal = "R$ #.##0,00"
        if pt2.DataBodyRange is not None:
            pt2.DataBodyRange.NumberFormatLocal = "R$ #.##0,00"
        din_ws.Columns("B:B").NumberFormatLocal = "R$ #.##0,00"
        din_ws.Columns("H:H").NumberFormatLocal = "R$ #.##0,00"

        din_ws.Columns("A:J").AutoFit()

        wb.Save()
    except Exception as exc:
        original_error = exc
    finally:
        if wb is not None:
            try:
                wb.Close(True)
            except Exception as close_exc:
                if original_error is None:
                    original_error = close_exc
        try:
            excel.Quit()
        except Exception as quit_exc:
            if original_error is None:
                original_error = quit_exc
        pythoncom.CoUninitialize()
    if original_error is not None:
        raise original_error


wb_source = openpyxl.load_workbook(SOURCE_PATH)
ws_source = wb_source["Órdenes"] if "Órdenes" in wb_source.sheetnames else wb_source.active
max_row = ws_source.max_row
max_col = ws_source.max_column

header_pedidos_color = color_key(ws_source.cell(row=1, column=1))
header_itens_color = color_key(ws_source.cell(row=2, column=1))

gray_rows = []
blue_rows = []

print(f"Total de linhas: {max_row}, Colunas: {max_col}")
print(f"Cor Pedidos: {header_pedidos_color}")
print(f"Cor Itens: {header_itens_color}")

for row_idx in range(1, max_row + 1):
    row_data = [ws_source.cell(row=row_idx, column=col_idx) for col_idx in range(1, max_col + 1)]
    row_color = color_key(row_data[0])

    if row_color == header_pedidos_color:
        gray_rows.append(row_data)
    elif row_color == header_itens_color:
        blue_rows.append(row_data)

print(f"\nLinhas cinza: {len(gray_rows)}")
print(f"Linhas azul: {len(blue_rows)}")

wb_new = openpyxl.Workbook()
wb_new.remove(wb_new.active)
ws_gray = wb_new.create_sheet("Azul")
ws_blue = wb_new.create_sheet("Verde")

for row_idx, row_data in enumerate(gray_rows, 1):
    for col_idx, source_cell in enumerate(row_data, 1):
        copy_cell(source_cell, ws_gray.cell(row=row_idx, column=col_idx))

for row_idx, row_data in enumerate(blue_rows, 1):
    for col_idx, source_cell in enumerate(row_data, 1):
        copy_cell(source_cell, ws_blue.cell(row=row_idx, column=col_idx))

# Colunas UF e FRENTE na aba Itens (Azul)
demanda_wb = openpyxl.load_workbook(DEMANDA_PATH, data_only=True)
demanda_ws = demanda_wb.active
demanda_map = []
for r in range(2, demanda_ws.max_row + 1):
    artigo_base = demanda_ws.cell(row=r, column=1).value
    frente = demanda_ws.cell(row=r, column=2).value
    if artigo_base and frente:
        key = str(artigo_base).strip().upper()
        value = str(frente).strip()
        if key and value:
            demanda_map.append((key, value, len(key)))
demanda_map.sort(key=lambda x: x[2], reverse=True)

headers = [str(ws_blue.cell(row=1, column=c).value or "").strip().upper() for c in range(1, ws_blue.max_column + 1)]
artigo_col = headers.index("ARTIGO") + 1
endereco_col = headers.index("ENDEREÇO DE ENTREGA") + 1
tributavel_col = headers.index("TRIBUTÁVEL") + 1
pedido_col = headers.index("ID. PEDIDO COMPRADOR") + 1
uf_col = ws_blue.max_column + 1
frente_col = ws_blue.max_column + 2
ws_blue.cell(row=1, column=uf_col, value="UF")
ws_blue.cell(row=1, column=frente_col, value="FRENTE")
ws_blue.cell(row=1, column=uf_col)._style = copy(ws_blue.cell(row=1, column=artigo_col)._style)
ws_blue.cell(row=1, column=frente_col)._style = copy(ws_blue.cell(row=1, column=artigo_col)._style)

faturado_ids = set()
faturado_ids.update(load_ids_from_base(CASA_CLIENTE_PATH_HINT, "SGP_NUMERO_PEDIDO"))
faturado_ids.update(load_ids_from_base(IMPLANTACAO_PATH_HINT, "PEDIDO"))

frente_count = 0
tributavel_count = 0
faturado_count = 0
for r in range(2, ws_blue.max_row + 1):
    endereco = ws_blue.cell(row=r, column=endereco_col).value
    artigo = ws_blue.cell(row=r, column=artigo_col).value
    pedido = ws_blue.cell(row=r, column=pedido_col).value

    uf = ""
    if endereco:
        endereco_txt = str(endereco).strip()
        match = re.search(r"([A-Za-z]{2})\s*$", endereco_txt)
        if match:
            uf = match.group(1).upper()
        elif len(endereco_txt) >= 2:
            uf = endereco_txt[-2:].upper()
    ws_blue.cell(row=r, column=uf_col, value=uf)

    # Normaliza TRIBUTÁVEL com 2 casas ainda na aba Verde.
    trib_cell = ws_blue.cell(row=r, column=tributavel_col)
    trib_value = parse_brl_number(trib_cell.value)
    if trib_value is not None:
        trib_cell.value = round(trib_value, 2)
        trib_cell.number_format = "R$ #,##0.00"
        tributavel_count += 1

    frente_val = ""
    if artigo:
        artigo_txt = str(artigo).upper()
        for key, value, _ in demanda_map:
            if key in artigo_txt:
                frente_val = value
                break
    pedido_normalizado = normalize_order_id(pedido)
    if pedido_normalizado and pedido_normalizado in faturado_ids:
        if frente_val:
            if not str(frente_val).upper().endswith("_FATURADO"):
                frente_val = f"{frente_val}_FATURADO"
        else:
            frente_val = "FATURADO"
        faturado_count += 1
    if frente_val:
        frente_count += 1
    ws_blue.cell(row=r, column=frente_col, value=frente_val)

for ws in (ws_gray, ws_blue):
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].hidden = False
    for c in range(1, ws.max_column + 1):
        letter = openpyxl.utils.get_column_letter(c)
        ws.column_dimensions[letter].hidden = False

saved_path = OUTPUT_PATH
try:
    wb_new.save(saved_path)
except PermissionError:
    saved_path = BASES_DIR / "Relatorio_Separado_novo.xlsx"
    wb_new.save(saved_path)

print(f"\n✓ Arquivo criado com sucesso: {saved_path}")
print(f"FRENTE preenchida em {frente_count} linhas.")
print(f"FRENTE = FATURADO em {faturado_count} linhas.")
print(f"TRIBUTÁVEL normalizado em {tributavel_count} linhas.")
try:
    create_din_with_excel(saved_path)
    print("Aba DIN criada com as duas dinâmicas.")
except Exception as exc:
    print(f"Aba DIN não foi criada automaticamente: {exc}")
    raise
