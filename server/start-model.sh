#!/usr/bin/env bash
# Starts llama-server with the settings from .env. Run from anywhere.
set -euo pipefail
REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
set -a; source "$REPO_DIR/.env"; set +a

ARGS=(
  -m "$MODEL_PATH"
  --alias "$MODEL_ALIAS"
  --host "$MODEL_HOST"
  --port "$MODEL_PORT"
  -c "$MODEL_CTX"
  -np 1                      # one slot, so the whole context window goes to one chat
  --flash-attn on
  --cache-type-k q8_0        # 8-bit KV cache roughly halves its memory use
  --cache-type-v q8_0
  --jinja                    # use the model's own chat template, needed for tool calls
)
if [ -n "${CHAT_TEMPLATE_KWARGS:-}" ]; then
  ARGS+=(--chat-template-kwargs "$CHAT_TEMPLATE_KWARGS")
fi
exec "$LLAMA_SERVER" "${ARGS[@]}"
