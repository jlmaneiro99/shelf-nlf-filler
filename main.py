from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Dict, List, Optional
from copy import deepcopy
import base64
import io
import openpyxl
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
    # Optional precise location (1-based row, column letter) from Step 3 mapping
    row: Optional[int] = None
    col: Optional[str] = None


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


def find_template_sheet(wb) -> str:
    for name in wb.sheetnames:
        nl = norm(name)
        if nl.startswith('product') or nl == 'template' or nl == 'new line':
            return name
    return wb.sheetnames[0]


def build_label_to_row(ws) -> Dict[str, int]:
    label_to_row = {}
    for row in ws.iter_rows(min_col=2, max_col=2):
        for cell in row:
            if cell.value:
                label = norm(cell.value)
                if label not in label_to_row:
                    label_to_row[label] = cell.row
    return label_to_row


def resolve_row(mapping: FieldMapping, label_to_row: Dict[str, int]) -> Optional[int]:
    """Find the target row: explicit row first, then exact label, then fuzzy."""
    if mapping.row:
        return mapping.row
    search = norm(mapping.field_name)
    row_num = label_to_row.get(search)
    if row_num is None:
        for label, rn in label_to_row.items():
            if search in label or label in search:
                row_num = rn
                break
    return row_num


def fillable(mapping: FieldMapping) -> bool:
    return bool(mapping.mapped_value) and mapping.status != 'missing'


def fill_sheet(ws, mappings: List[FieldMapping], force_col: Optional[int] = None) -> int:
    """Write mapped values into a vertical (label in column B) sheet.
    Only cell .value is touched — styles, fills and data validation
    are never modified."""
    label_to_row = build_label_to_row(ws)
    filled = 0
    for mapping in mappings:
        if not fillable(mapping):
            continue
        row_num = resolve_row(mapping, label_to_row)
        if row_num is None:
            continue
        if force_col is not None:
            col_idx = force_col
        elif mapping.col:
            col_idx = col_letter_to_index(mapping.col)
        else:
            col_idx = 3
        ws.cell(row=row_num, column=col_idx).value = mapping.mapped_value
        filled += 1
    return filled


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
    for row in ws.iter_rows(min_row=1, max_row=min(scan_rows, ws.max_row)):
        count = sum(1 for cell in row if cell.value not in (None, ''))
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
                # Place it right after the previous product sheet
                current_idx = wb.sheetnames.index(ws.title)
                wb.move_sheet(ws.title, offset=(template_idx + i) - current_idx)
        filled += fill_sheet(ws, prod.mappings)
    return filled


def fill_columns(wb, products: List[ProductMappings]) -> int:
    ws = wb[find_template_sheet(wb)]
    filled = 0
    for i, prod in enumerate(products):
        filled += fill_sheet(ws, prod.mappings, force_col=3 + i)
    return filled


def fill_rows(wb, products: List[ProductMappings]) -> int:
    ws = wb[find_template_sheet(wb)]
    header_row = find_header_row(ws)

    col_to_field: Dict[int, str] = {}
    for cell in ws[header_row]:
        if cell.value not in (None, ''):
            col_to_field[cell.column] = norm(cell.value)

    filled = 0
    for i, prod in enumerate(products):
        row_num = header_row + 1 + i
        values = {norm(m.field_name): m.mapped_value for m in prod.mappings if fillable(m)}
        for col_idx, field_label in col_to_field.items():
            value = values.get(field_label)
            if value is None:
                # Fuzzy match: header label contains field name or vice versa
                for name, v in values.items():
                    if name in field_label or field_label in name:
                        value = v
                        break
            if value is not None:
                ws.cell(row=row_num, column=col_idx).value = value
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
