# Shelf NLF Filler API

FastAPI microservice that fills retailer New Line Forms (Excel) using a hybrid engine:

1. **FormSpec** (from Supabase `nlf-mapper` + Claude) — structural map of the workbook
2. **Deterministic rules** (`map_field`) — known fields filled instantly
3. **Claude fallback** — one batched call per product for unknown labels
4. **Verification gate** — type checks, formula protection, no garbage output

## Endpoints

- `POST /fill` — body: `{ file_base64, products, retailer_name, fill_mode?, form_spec? }`
- `GET /health` — health check

## Environment variables

| Variable | Required | Purpose |
|----------|----------|---------|
| `ANTHROPIC_API_KEY` | Optional on Railway | Unknown-field reasoning (Layer 3). Known fields fill without it. |
| `PORT` | Auto on Railway | Server port (default 8000) |

**Never** expose `ANTHROPIC_API_KEY` in frontend code or commit it to git.

## Run locally

```bash
cp .env.example .env
# Edit .env and set ANTHROPIC_API_KEY (optional for rule-only fills)

pip install -r requirements.txt
python3 test_filler.py && python3 test_fill.py
uvicorn main:app --reload
```

## Deploy (Railway)

1. Connect this repo to Railway (`shelf-nlf-filler`)
2. **Variables → New Variable:** `ANTHROPIC_API_KEY` = your Anthropic secret key
3. Deploy; copy the service URL into the main app `.env` as `VITE_NLF_FILLER_URL`

Supabase Edge Function `nlf-mapper` uses the same `ANTHROPIC_API_KEY` secret (Dashboard → Edge Functions → Secrets) for FormSpec analysis.

## Tests

```bash
python3 test_filler.py   # 50+ unit + integration tests
python3 test_fill.py     # tabs / columns / rows smoke tests
```
