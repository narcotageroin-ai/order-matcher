import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

st.set_page_config(page_title="Order Matcher", layout="wide")

def read_file(file):
    ext = os.path.splitext(file.name)[1].lower()
    if ext == ".csv":
        return pd.read_csv(file), "csv", None
    elif ext == ".xls":
        return pd.read_excel(file, engine="xlrd"), "xls", None
    elif ext == ".xlsx":
        df = pd.read_excel(file, engine="openpyxl")
        file.seek(0)
        wb = load_workbook(file)
        return df, "xlsx", wb
    else:
        st.error(f"❌ Неподдерживаемый формат файла: {ext}")
        return None, None, None

st.title("📦 Автоматическое формирование заказов поставщикам")

file_order = st.file_uploader("Загрузите файл заказа (CSV/XLS/XLSX)", type=["csv", "xls", "xlsx"])
file_supplier = st.file_uploader("Загрузите файл прайс-листа поставщика (CSV/XLS/XLSX)", type=["csv", "xls", "xlsx"])

if file_order and file_supplier:
    df_order, order_type, _ = read_file(file_order)
    df_supplier, supplier_type, wb = read_file(file_supplier)
    if df_order is not None and df_supplier is not None:
        st.success("✅ Файлы успешно загружены и распознаны!")
        st.write("### Заказ")
        st.dataframe(df_order.head())
        st.write("### Прайс поставщика")
        st.dataframe(df_supplier.head())
