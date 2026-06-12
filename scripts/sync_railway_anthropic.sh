#!/usr/bin/env bash
# Sync ANTHROPIC_API_KEY from local .env to Railway (requires: railway login).
# Never commits or prints the key.
set -euo pipefail
cd "$(dirname "$0")"

if ! command -v railway >/dev/null 2>&1; then
  echo "Install Railway CLI: npm i -g @railway/cli  OR  brew install railway"
  echo "Then run: railway login"
  exit 1
fi

if [[ ! -f .env ]]; then
  echo "Missing nlf-filler-api/.env — copy .env.example and set ANTHROPIC_API_KEY"
  exit 1
fi

# shellcheck disable=SC1091
set -a
source .env
set +a

if [[ -z "${ANTHROPIC_API_KEY:-}" ]]; then
  echo "ANTHROPIC_API_KEY is empty in .env"
  exit 1
fi

echo "Setting ANTHROPIC_API_KEY on linked Railway service (value hidden)..."
railway variables set "ANTHROPIC_API_KEY=${ANTHROPIC_API_KEY}"

echo "Redeploying service..."
railway up --detach 2>/dev/null || railway redeploy --yes 2>/dev/null || {
  echo "Variable set. Trigger redeploy manually in Railway dashboard if needed."
}

echo "Verify with: python3 verify_live_railway.py"
