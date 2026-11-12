#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
if [ -x "/home/im-redmond/Desktop/STREAMLIT_APP/.venv/bin/streamlit" ]; then
    exec /home/im-redmond/Desktop/STREAMLIT_APP/.venv/bin/streamlit run app.py "$@"
else
    exec streamlit run app.py "$@"
fi
