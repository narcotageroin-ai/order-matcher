import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Order Matcher", layout="wide")

st.title("📦 Order Matcher — автоматизация заказов")

st.write("""
Загрузите два файла:
1. Наш файл заказа (с артикулами/штрихкодами и количеством).
2. Прайс поставщика.
""")

file_order = st.file_uploader("Наш файл заказа", type=["csv","xls","xlsx"])
file_supplier = st.file_uploader("Файл поставщика", type=["csv","xls","xlsx"])

def read_file(file):
    ext = os.path.splitext(file.name)[1].lower()
    if ext == ".csv":
        return pd.read_csv(file), "csv", None
    elif ext in [".xls", ".xlsx"]:
        df = pd.read_excel(file, engine="openpyxl")
        wb = None
        if ext == ".xlsx":
            file.seek(0)
            wb = load_workbook(file)
        return df, ext, wb
    else:
        st.error(f"❌ Неподдерживаемый формат файла: {ext}")
        return None, None, None

if file_order and file_supplier:
    df_order, order_type, _ = read_file(file_order)
    df_supplier, supplier_type, supplier_wb = read_file(file_supplier)

    if df_order is not None and df_supplier is not None:
        st.success("Файлы загружены успешно ✅")

        # выбор столбцов
        st.subheader("Настройка сопоставления")
        order_keys = st.multiselect("Колонки заказа для поиска (артикул, штрихкод и т.п.)", df_order.columns.tolist())
        supplier_keys = st.multiselect("Колонки поставщика для поиска", df_supplier.columns.tolist())

        qty_col_order = st.selectbox("Колонка количества в заказе", df_order.columns.tolist())
        qty_col_supplier = st.text_input("Название/позиция колонки количества у поставщика")

        overwrite = st.checkbox("Перезаписывать уже заполненные значения")

        if st.button("🚀 Обработать"):
            found, not_found = 0, 0

            for idx, row in df_order.iterrows():
                qty = row[qty_col_order]
                match = None

                for ok, sk in zip(order_keys, supplier_keys):
                    supplier_match = df_supplier[df_supplier[sk].astype(str) == str(row[ok])]
                    if not supplier_match.empty:
                        match = supplier_match.index[0]
                        break

                if match is not None:
                    if qty_col_supplier in df_supplier.columns:
                        if overwrite or pd.isna(df_supplier.at[match, qty_col_supplier]):
                            df_supplier.at[match, qty_col_supplier] = qty
                    else:
                        try:
                            col_idx = int(qty_col_supplier)
                            if overwrite or pd.isna(df_supplier.iat[match, col_idx]):
                                df_supplier.iat[match, col_idx] = qty
                        except:
                            st.error(f"❌ Колонка {qty_col_supplier} не найдена")
                    found += 1
                else:
                    not_found += 1

            st.info(f"Обработано {len(df_order)} строк. Найдено: {found}, Не найдено: {not_found}")

            # сохраняем результат
            out_file = BytesIO()
            if supplier_type == "csv":
                df_supplier.to_csv(out_file, index=False, encoding="utf-8-sig")
                mime="text/csv"
                fname="supplier_updated.csv"
            elif supplier_type in ["xls","xlsx"]:
                df_supplier.to_excel(out_file, index=False, engine="openpyxl")
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                fname="supplier_updated.xlsx"
            else:
                mime="application/octet-stream"
                fname="supplier_updated.dat"

            st.download_button("💾 Скачать результат", out_file.getvalue(), file_name=fname, mime=mime)
