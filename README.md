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

Backend service: **`shelf-nlf-filler`** → `https://shelf-nlf-filler-production.up.railway.app`

The frontend (`VITE_NLF_FILLER_URL`) only points at this URL — **do not** put `ANTHROPIC_API_KEY` in the frontend or root `.env` for Vite.

### Set Anthropic key on the backend service

1. [railway.app](https://railway.app) → project → service **`shelf-nlf-filler`** (Python/Dockerfile, not a static site)
2. **Variables** → **Production** environment
3. Add **`ANTHROPIC_API_KEY`** (exact name, no `VITE_` prefix)
4. **Deploy** → Redeploy latest (required after adding/changing variables)

Verify without guessing:

```bash
curl -s https://shelf-nlf-filler-production.up.railway.app/health/config
# Expect: {"anthropic_key_present":true,"anthropic_key_source":"env",...}
```

Or CLI (after `railway login` and linking this repo):

```bash
./scripts/sync_railway_anthropic.sh
python3 verify_live_railway.py
```

## Tests

```bash
python3 test_filler.py   # 54 unit + integration tests
python3 test_fill.py     # tabs / columns / rows smoke tests
```

## Live verification (Railway)

```bash
python3 verify_live_railway.py
# Uses VITE_NLF_FILLER_URL from ../.env or defaults to production URL
```

Checks `/health`, `/health/config` (anthropic_key_present), and a live `/fill` with Claude fallback.
