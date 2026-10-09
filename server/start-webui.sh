#!/usr/bin/env bash
# Starts Open WebUI pointed at the local llama-server. Run from anywhere.
set -euo pipefail
REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
set -a; source "$REPO_DIR/.env"; set +a

export DATA_DIR="$WEBUI_DATA_DIR"
export OPENAI_API_BASE_URL="http://$MODEL_HOST:$MODEL_PORT/v1"
export OPENAI_API_KEY
export ENABLE_OLLAMA_API=false
export HF_HUB_OFFLINE
export GRADER_DIR GRADER_PACKETS_DIR GRADER_CONFIG

mkdir -p "$DATA_DIR"
# Open WebUI's default port is 8080. Host and port are passed as flags here because
# the HOST and PORT environment variables were not honored in the tested version.
exec "$WEBUI_VENV/bin/open-webui" serve --host "$WEBUI_HOST" --port "$WEBUI_PORT"
