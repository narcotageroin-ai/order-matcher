import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Формирование заказов", page_icon="📦", layout="centered")
st.title("📦 Автоматическое формирование заказов поставщикам")

def load_excel(file):
    if file.name.endswith(".xlsx"):
        return pd.read_excel(file, engine="openpyxl")
    elif file.name.endswith(".xls"):
        return pd.read_excel(file, engine="xlrd")
    elif file.name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        st.error("❌ Поддерживаются только файлы .xls, .xlsx или .csv")
        return None

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

        order_keys = st.multiselect("Столбцы для поиска (наш файл)", df_order.columns)
        supplier_keys = st.multiselect("Столбцы для поиска (файл поставщика)", df_supplier.columns)
        qty_col = st.selectbox("Столбец с количеством (наш файл)", df_order.columns)
        supplier_qty_col = st.selectbox("Столбец для вставки количества у поставщика", df_supplier.columns)

        if len(order_keys) != len(supplier_keys):
            st.warning("⚠️ Нужно выбрать одинаковое количество ключевых столбцов в обоих файлах.")
        elif st.button("Сформировать файл"):
            stats = {}
            matched = 0

            file_supplier.seek(0)
            wb = load_workbook(file_supplier)
            ws = wb.active

            supplier_columns = list(df_supplier.columns)
            supplier_col_index_map = {col: idx+1 for idx, col in enumerate(supplier_columns)}
            supplier_qty_col_idx = supplier_col_index_map[supplier_qty_col]

            for idx, (okey, skey) in enumerate(zip(order_keys, supplier_keys), start=1):
                qty_dict = dict(zip(df_order[okey].astype(str), df_order[qty_col]))
                found = 0

                for row in range(2, ws.max_row + 1):
                    key_val = str(ws.cell(row=row, column=supplier_col_index_map[skey]).value)
                    if key_val in qty_dict and ws.cell(row=row, column=supplier_qty_col_idx).value is None:
                        ws.cell(row=row, column=supplier_qty_col_idx).value = qty_dict[key_val]
                        found += 1

                key_name = f"{order_keys[idx-1]} -> {supplier_keys[idx-1]}"
                stats[f"Найдено по ключу {idx} ({key_name})"] = found
                matched += found

            total = len(df_supplier)
            not_found = total - matched

            st.subheader("📊 Результаты обработки")
            st.write(f"Всего товаров у поставщика: **{total}**")
            for k, v in stats.items():
                st.write(f"🔍 {k}: **{v}**")
            st.write(f"⚠️ Не найдено: **{not_found}**")

            output = BytesIO()
            wb.save(output)

            st.success("✅ Готово! Скачайте результат ниже.")
            st.download_button(
                label="⬇ Скачать готовый Excel (с сохранением формата)",
                data=output.getvalue(),
                file_name="supplier_with_qty.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
