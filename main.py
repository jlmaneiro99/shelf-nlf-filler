from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
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


class ProductFill(BaseModel):
    product_name: str
    mappings: List[FieldMapping]


class FillRequest(BaseModel):
    file_base64: str
    fill_mode: str = 'tabs'  # 'tabs' or 'columns'
    products: Optional[List[ProductFill]] = None
    retailer_name: str
    # Legacy single-product payload (kept for backwards compatibility)
    mappings: Optional[List[FieldMapping]] = None
    product_name: Optional[str] = None


def col_letter_to_index(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def find_template_sheet(wb) -> str:
    for name in wb.sheetnames:
        nl = name.lower().strip()
        if nl.startswith('product') or nl == 'template' or nl == 'new line':
            return name
    return wb.sheetnames[0]


def build_label_to_row(ws) -> dict:
    label_to_row = {}
    for row in ws.iter_rows(min_col=2, max_col=2):
        for cell in row:
            if cell.value:
                label = str(cell.value).strip().lower()
                if label not in label_to_row:
                    label_to_row[label] = cell.row
    return label_to_row


def resolve_row(mapping: FieldMapping, label_to_row: dict) -> Optional[int]:
    """Find the target row: explicit row first, then exact label, then fuzzy."""
    if mapping.row:
        return mapping.row
    search = mapping.field_name.strip().lower()
    row_num = label_to_row.get(search)
    if row_num is None:
        for label, rn in label_to_row.items():
            if search in label or label in search:
                row_num = rn
                break
    return row_num


def fill_sheet(ws, mappings: List[FieldMapping], force_col: Optional[int] = None) -> int:
    """Write mapped values into a sheet. Only cell .value is touched —
    styles, fills and data validation are never modified."""
    label_to_row = build_label_to_row(ws)
    filled = 0
    for mapping in mappings:
        if not mapping.mapped_value or mapping.status == 'missing':
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


@app.post("/fill")
async def fill_nlf(req: FillRequest):
    try:
        products = req.products
        if not products:
            # Legacy single-product payload
            if req.mappings is None:
                raise HTTPException(status_code=422, detail="No products or mappings provided")
            products = [ProductFill(
                product_name=req.product_name or 'Product',
                mappings=req.mappings,
            )]

        file_bytes = base64.b64decode(req.file_base64)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

        template_name = find_template_sheet(wb)
        template_ws = wb[template_name]
        template_idx = wb.sheetnames.index(template_name)

        filled = 0

        if req.fill_mode == 'columns' and len(products) > 1:
            # One column per product: C = product 1, D = product 2, ...
            for i, prod in enumerate(products):
                filled += fill_sheet(template_ws, prod.mappings, force_col=3 + i)
        else:
            # One tab per product (default). All other sheets stay untouched.
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

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        date_str = datetime.date.today().isoformat()
        if len(products) > 1:
            base = f"{req.retailer_name}_NLF_{len(products)}_products_{date_str}.xlsx"
        else:
            base = f"{req.retailer_name}_{products[0].product_name}_{date_str}.xlsx"
        filename = base.replace(" ", "_")

        return {
            "file_base64": base64.b64encode(output.read()).decode(),
            "filename": filename,
            "fields_filled": filled,
            "products_filled": len(products),
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
