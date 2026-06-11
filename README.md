# Shelf NLF Filler API

FastAPI microservice that fills retailer New Line Forms (Excel) using openpyxl,
preserving all formatting and data-validation dropdowns.

## Endpoints

- `POST /fill` — body: `{ file_base64, mappings, product_name, retailer_name }`,
  returns `{ file_base64, filename, fields_filled }`
- `GET /health` — health check

## Run locally

```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

## Deploy (Railway)

1. Push this folder to a GitHub repo
2. railway.app → New Project → Deploy from GitHub repo
3. Copy the generated URL into the app's `.env` as `VITE_NLF_FILLER_URL`
