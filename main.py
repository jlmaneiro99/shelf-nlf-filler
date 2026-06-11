from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
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

class FillRequest(BaseModel):
    file_base64: str
    mappings: List[FieldMapping]
    product_name: str
    retailer_name: str

@app.post("/fill")
async def fill_nlf(req: FillRequest):
    try:
        file_bytes = base64.b64decode(req.file_base64)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        product_sheet = None
        for name in wb.sheetnames:
            nl = name.lower()
            if nl.startswith('product') or nl == 'template' or nl == 'new line':
                product_sheet = name
                break
        if not product_sheet:
            product_sheet = wb.sheetnames[0]
        ws = wb[product_sheet]
        label_to_row = {}
        for row in ws.iter_rows(min_col=2, max_col=2):
            for cell in row:
                if cell.value:
                    label = str(cell.value).strip().lower()
                    label_to_row[label] = cell.row
        filled = 0
        for mapping in req.mappings:
            if not mapping.mapped_value or mapping.status == 'missing':
                continue
            search = mapping.field_name.strip().lower()
            row_num = label_to_row.get(search)
            if row_num is None:
                for label, rn in label_to_row.items():
                    if search in label or label in search:
                        row_num = rn
                        break
            if row_num is not None:
                ws.cell(row=row_num, column=3).value = mapping.mapped_value
                filled += 1
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        date_str = datetime.date.today().isoformat()
        filename = f"{req.retailer_name}_{req.product_name}_{date_str}.xlsx".replace(" ", "_")
        return {
            "file_base64": base64.b64encode(output.read()).decode(),
            "filename": filename,
            "fields_filled": filled
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
async def health():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run(app, host="0.0.0.0", port=port)

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
