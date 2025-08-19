import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Формирование заказов", page_icon="📦", layout="centered")
st.title("📦 Автоматическое формирование заказов поставщикам")

def load_excel(file):
    """Определяем, как открыть файл по расширению"""
    if file.name.endswith(".xlsx"):
        return pd.read_excel(file, engine="openpyxl")
    elif file.name.endswith(".xls"):
        return pd.read_excel(file, engine="xlrd")
    elif file.name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        st.error("❌ Поддерживаются только файлы .xls, .xlsx или .csv")
        return None

# Загружаем файлы
file_order = st.file_uploader("Загрузите файл заказа (Excel или CSV)", type=["xls", "xlsx", "csv"])
file_supplier = st.file_uploader("Загрузите файл поставщика (Excel или CSV)", type=["xls", "xlsx", "csv"])

if file_order and file_supplier:
    df_order = load_excel(file_order)
    df_supplier = load_excel(file_supplier)

    if df_order is not None and df_supplier is not None:
        st.subheader("Наш файл заказа:")
        st.dataframe(df_order.head())

        st.subheader("Файл поставщика:")
        st.dataframe(df_supplier.head())

        key_col = st.selectbox("Столбец для связывания (наш файл)", df_order.columns)
        qty_col = st.selectbox("Столбец с количеством (наш файл)", df_order.columns)
        supplier_key_col = st.selectbox("Столбец для поиска у поставщика", df_supplier.columns)
        supplier_qty_col = st.selectbox("Столбец для вставки количества у поставщика", df_supplier.columns)

        if st.button("Сформировать файл"):
            # Словарь ключ -> количество
            qty_dict = dict(zip(df_order[key_col].astype(str), df_order[qty_col]))
            
            # Сохраняем оригинальные данные для статистики
            total = len(df_supplier)
            df_supplier[supplier_key_col] = df_supplier[supplier_key_col].astype(str)

            # Подстановка количества
            df_supplier[supplier_qty_col] = df_supplier[supplier_key_col].map(qty_dict)

            # Считаем статистику
            updated = df_supplier[supplier_qty_col].notna().sum()
            not_found = total - updated

            # Показ статистики
            st.subheader("📊 Результаты обработки")
            st.write(f"Всего товаров у поставщика: **{total}**")
            st.write(f"✅ Найдено и обновлено: **{updated}**")
            st.write(f"⚠️ Не найдено в заказе: **{not_found}**")

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
