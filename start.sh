#!/usr/bin/env bash
set -e

# Start licensing API locally inside the same container
uvicorn licensing_api.main:app --host 0.0.0.0 --port 8000 &

# Start Streamlit (Render provides PORT)
streamlit run aplicacion.py --server.port="${PORT:-8501}" --server.address=0.0.0.0
