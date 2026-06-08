#!/usr/bin/env bash
# Запуск приложения на macOS / Linux.
set -e
cd "$(dirname "$0")"

if [ ! -d "venv" ]; then
    echo "Создаю виртуальное окружение..."
    python3 -m venv venv
    source venv/bin/activate
    pip install --upgrade pip
    pip install -r requirements.txt
else
    source venv/bin/activate
fi

streamlit run app.py
