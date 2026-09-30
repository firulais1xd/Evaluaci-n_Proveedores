#!/usr/bin/env bash
# Mac / Linux: ejecutar ./iniciar.sh
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "Creando entorno de Python por primera vez..."
  python3 -m venv .venv
fi
source .venv/bin/activate
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt
python -m streamlit run app.py
