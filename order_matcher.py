import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Формирование заказов", page_icon="📦", layout="centered")
st.title("📦 Автоматическое формирование заказов поставщикам")

# Загружаем файлы
file_order = st.file_uploader("Загрузите файл заказа (Excel)", type=["xls", "xlsx"])
file_supplier = st.file_uploader("Загрузите файл поставщика (Excel)", type=["xls", "xlsx"])

if file_order and file_supplier:
    df_order = pd.read_excel(file_order)
    df_supplier = pd.read_excel(file_supplier)

    st.subheader("Наш файл заказа:")
    st.dataframe(df_order.head())

    st.subheader("Файл поставщика:")
    st.dataframe(df_supplier.head())

    key_col = st.selectbox("Столбец для связывания (наш файл)", df_order.columns)
    qty_col = st.selectbox("Столбец с количеством (наш файл)", df_order.columns)
    supplier_key_col = st.selectbox("Столбец для поиска у поставщика", df_supplier.columns)
    supplier_qty_col = st.selectbox("Столбец для вставки количества у поставщика", df_supplier.columns)

    if st.button("Сформировать файл"):
        qty_dict = dict(zip(df_order[key_col].astype(str), df_order[qty_col]))
        df_supplier[supplier_key_col] = df_supplier[supplier_key_col].astype(str)
        df_supplier[supplier_qty_col] = df_supplier[supplier_key_col].map(qty_dict)

        # Сохраняем в память
        output = BytesIO()
        df_supplier.to_excel(output, index=False, engine="openpyxl")
        st.success("✅ Готово! Скачайте результат ниже.")

        st.download_button(
            label="⬇ Скачать готовый Excel",
            data=output.getvalue(),
            file_name="supplier_with_qty.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
