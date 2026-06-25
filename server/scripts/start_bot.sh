#!/bin/bash
# Source env vars и пробросить их в sudo для botuser
set -a
source /home/botuser/.bot_env
set +a
export LLM_BASE_URL LLM_API_KEY
exec sudo -u botuser env LLM_BASE_URL="$LLM_BASE_URL" LLM_API_KEY="$LLM_API_KEY" /opt/transcriber/venv/bin/python /opt/transcriber/bot.py
