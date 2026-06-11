from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Dict, List, Optional, Tuple, Union
from copy import deepcopy, copy as copy_obj
import base64
import io
import openpyxl
from openpyxl.cell.cell import MergedCell
import datetime
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


class FieldMapping(BaseModel):
    field_name: str
    mapped_value: Optional[str] = None
    status: str
    row: Optional[int] = None
    col: Optional[Union[int, str]] = None


class ProductMappings(BaseModel):
    product_name: str
    mappings: List[FieldMapping]


class FillRequest(BaseModel):
    file_base64: str
    fill_mode: str = 'tabs'  # 'tabs' | 'columns' | 'rows'
    products: List[ProductMappings]
    retailer_name: str


def col_letter_to_index(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def norm(label) -> str:
    return str(label).strip().lower()


def safe_write(ws, row_num: int, col: int, value) -> None:
    """Write to a cell, redirecting merged cells to the top-left of their range."""
    cell = ws.cell(row=row_num, column=col)
    if isinstance(cell, MergedCell):
        for merge_range in ws.merged_cells.ranges:
            if (
                merge_range.min_row <= row_num <= merge_range.max_row
                and merge_range.min_col <= col <= merge_range.max_col
            ):
                top_left = ws.cell(row=merge_range.min_row, column=merge_range.min_col)
                top_left.value = value
                return
        return
    cell.value = value


def find_template_sheet(wb) -> str:
    product_sheet = None
    for name in wb.sheetnames:
        nl = name.lower()
        if (
            nl.startswith('product')
            or nl.startswith('sku')
            or nl == 'template'
            or nl == 'new line'
            or nl == 'new line form'
            or nl == 'nlf'
            or nl == 'form'
            or 'line form' in nl
            or 'new line' in nl
            or 'submission' in nl
        ):
            product_sheet = name
            break

    if not product_sheet:
        best = None
        best_count = 0
        for name in wb.sheetnames:
            ws_temp = wb[name]
            count = sum(
                1
                for row in ws_temp.iter_rows()
                for cell in row
                if cell.value and not isinstance(cell, MergedCell)
            )
            if count > best_count:
                best_count = count
                best = name
        product_sheet = best or wb.sheetnames[0]

    return product_sheet


def build_label_map(ws) -> Dict[str, Tuple[int, int]]:
    """Map normalised field labels to (row, label_column). Skips merged cells."""
    label_map: Dict[str, Tuple[int, int]] = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value:
                label = norm(cell.value)
                if label and label not in label_map:
                    label_map[label] = (cell.row, cell.column)
    return label_map


def resolve_label_pos(
    mapping: FieldMapping, label_map: Dict[str, Tuple[int, int]]
) -> Optional[Tuple[int, int]]:
    search = norm(mapping.field_name)
    pos = label_map.get(search)
    if pos is None:
        for label, candidate in label_map.items():
            if search in label or label in search:
                pos = candidate
                break
    return pos


def fillable(mapping: FieldMapping) -> bool:
    return bool(mapping.mapped_value) and mapping.status != 'missing'


def resolve_col_index(mapping: FieldMapping, force_col: Optional[int] = None) -> Optional[int]:
    if mapping.col is not None:
        if isinstance(mapping.col, int):
            return mapping.col
        return col_letter_to_index(str(mapping.col))
    return force_col


def fill_sheet(
    ws,
    mappings: List[FieldMapping],
    force_col: Optional[int] = None,
    strict_coords: bool = False,
) -> int:
    """Pure writer: use row+col coordinates when provided; label fallback otherwise."""
    label_map = build_label_map(ws)
    filled = 0

    for mapping in mappings:
        if not fillable(mapping):
            continue

        # Direct coordinate write — trust Claude row; allow force_col to override column in columns mode
        if mapping.row is not None and (mapping.col is not None or force_col is not None):
            col_idx = force_col if force_col is not None else resolve_col_index(mapping)
            if col_idx is not None:
                safe_write(ws, mapping.row, col_idx, mapping.mapped_value)
                filled += 1
            continue

        if strict_coords:
            continue

        row_num: Optional[int] = None
        value_col: Optional[int] = None

        if mapping.row:
            row_num = mapping.row
            value_col = resolve_col_index(mapping, force_col) or 3
        else:
            pos = resolve_label_pos(mapping, label_map)
            if pos is None:
                continue
            row_num, label_col = pos
            value_col = force_col if force_col is not None else label_col + 1

        if row_num is None or value_col is None:
            continue

        safe_write(ws, row_num, value_col, mapping.mapped_value)
        filled += 1

    return filled


def find_data_row_range(ws, label_col: int = 2) -> Tuple[int, int]:
    """First and last rows with field labels in the label column."""
    start: Optional[int] = None
    end: Optional[int] = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.column != label_col or isinstance(cell, MergedCell):
                continue
            if cell.value and str(cell.value).strip():
                if start is None:
                    start = cell.row
                end = cell.row
    return start or 5, end or (ws.max_row or 5)


def copy_column_format(ws, template_col: int, target_col: int, data_start_row: int, data_end_row: int) -> None:
    """Copy solid fill (and number format) from template product column to target column."""
    for row in range(data_start_row, data_end_row + 1):
        template_cell = ws.cell(row=row, column=template_col)
        target_cell = ws.cell(row=row, column=target_col)
        if isinstance(target_cell, MergedCell):
            continue
        if template_cell.fill and template_cell.fill.fill_type == 'solid':
            target_cell.fill = copy_obj(template_cell.fill)
        if template_cell.number_format:
            target_cell.number_format = template_cell.number_format


def copy_template_sheet(wb, template_ws, title: str):
    """Copy the template sheet (openpyxl's copy_worksheet does not copy
    data validations, so dropdowns are re-added manually)."""
    new_ws = wb.copy_worksheet(template_ws)
    new_ws.title = title
    for dv in template_ws.data_validations.dataValidation:
        new_ws.add_data_validation(deepcopy(dv))
    return new_ws


def find_header_row(ws, scan_rows: int = 30) -> int:
    """Header row = the row with the most non-empty cells (horizontal NLFs)."""
    best_row, best_count = 1, 0
    max_row = min(scan_rows, ws.max_row or scan_rows)
    for row in ws.iter_rows(min_row=1, max_row=max_row):
        count = sum(
            1
            for cell in row
            if not isinstance(cell, MergedCell) and cell.value not in (None, '')
        )
        if count > best_count:
            best_row, best_count = row[0].row, count
    return best_row


def fill_tabs(wb, products: List[ProductMappings]) -> int:
    template_name = find_template_sheet(wb)
    template_ws = wb[template_name]
    template_idx = wb.sheetnames.index(template_name)
    filled = 0
    for i, prod in enumerate(products):
        if i == 0:
            ws = template_ws
        else:
            target_name = f"Product {i + 1}"
            if target_name in wb.sheetnames:
                ws = wb[target_name]
            else:
                ws = copy_template_sheet(wb, template_ws, target_name)
                current_idx = wb.sheetnames.index(ws.title)
                wb.move_sheet(ws.title, offset=(template_idx + i) - current_idx)
        filled += fill_sheet(ws, prod.mappings)
    return filled


def fill_columns(wb, products: List[ProductMappings]) -> int:
    ws = wb[find_template_sheet(wb)]
    template_col = 3
    data_start, data_end = find_data_row_range(ws, label_col=2)
    filled = 0
    for i, prod in enumerate(products):
        target_col = 3 + i
        fillable_mappings = [m for m in prod.mappings if fillable(m)]
        has_rows = bool(fillable_mappings) and all(m.row is not None for m in fillable_mappings)
        if has_rows:
            filled += fill_sheet(ws, prod.mappings, force_col=target_col, strict_coords=True)
        else:
            filled += fill_sheet(ws, prod.mappings, force_col=target_col)
        if target_col != template_col:
            copy_column_format(ws, template_col, target_col, data_start, data_end)
    return filled


def fill_rows(wb, products: List[ProductMappings]) -> int:
    ws = wb[find_template_sheet(wb)]
    header_row = find_header_row(ws)

    col_to_field: Dict[int, str] = {}
    for cell in ws[header_row]:
        if isinstance(cell, MergedCell):
            continue
        if cell.value not in (None, ''):
            col_to_field[cell.column] = norm(cell.value)

    filled = 0
    for i, prod in enumerate(products):
        row_num = header_row + 1 + i
        values = {norm(m.field_name): m.mapped_value for m in prod.mappings if fillable(m)}
        for col_idx, field_label in col_to_field.items():
            value = values.get(field_label)
            if value is None:
                for name, v in values.items():
                    if name in field_label or field_label in name:
                        value = v
                        break
            if value is not None:
                safe_write(ws, row_num, col_idx, value)
                filled += 1
    return filled


@app.post("/fill")
async def fill_nlf(req: FillRequest):
    try:
        if not req.products:
            raise HTTPException(status_code=422, detail="No products provided")

        file_bytes = base64.b64decode(req.file_base64)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

        if req.fill_mode == 'columns':
            filled = fill_columns(wb, req.products)
        elif req.fill_mode == 'rows':
            filled = fill_rows(wb, req.products)
        else:
            filled = fill_tabs(wb, req.products)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        date_str = datetime.date.today().isoformat()
        if len(req.products) > 1:
            base = f"{req.retailer_name}_NLF_{len(req.products)}_products_{date_str}.xlsx"
        else:
            base = f"{req.retailer_name}_{req.products[0].product_name}_{date_str}.xlsx"
        filename = base.replace(" ", "_")

        return {
            "file_base64": base64.b64encode(output.read()).decode(),
            "filename": filename,
            "fields_filled": filled,
            "products_filled": len(req.products),
            "fill_mode": req.fill_mode,
        }

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/health")
async def health():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run(app, host="0.0.0.0", port=port)
