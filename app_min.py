"""Минимальный test — проверка что Streamlit работает на этой машине."""
import streamlit as st

st.set_page_config(page_title="Test", layout="wide")
st.title("✅ Streamlit работает")
st.write("Если вы видите этот текст — фронтенд и бэкенд связаны корректно.")
st.success("Тогда проблема в нашем app.py — разберёмся.")
