
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Формирование заказов", page_icon="📦", layout="centered")
st.title("📦 Автоматическое формирование заказов поставщикам")

# ---------- Утилиты ----------
def load_table(file_bytes: bytes, filename: str):
    """Читаем заказ/прайс в DataFrame по расширению файла."""
    bio = BytesIO(file_bytes)
    if filename.lower().endswith(".xlsx") or filename.lower().endswith(".xlsm"):
        return pd.read_excel(bio, engine="openpyxl")
    elif filename.lower().endswith(".xls"):
        return pd.read_excel(bio, engine="xlrd")
    elif filename.lower().endswith(".csv"):
        return pd.read_csv(bio)
    else:
        return None

def normalize_key(v):
    if pd.isna(v):
        return None
    return str(v).strip()

def detect_header_row(ws, headers, max_scan=20):
    """Ищем строку заголовков на листе по списку имён столбцов (из pandas).
    Возвращаем 1-базовый индекс строки. Если не нашли — 1.
    """
    headers_set = {str(h) for h in headers}
    best_row = 1
    best_hits = 0
    for r in range(1, min(ws.max_row, max_scan) + 1):
        row_vals = [str(c.value) if c.value is not None else "" for c in ws[r]]
        hits = sum(1 for v in row_vals if v in headers_set)
        if hits > best_hits:
            best_hits = hits
            best_row = r
            if hits == len(headers_set):
                break
    return best_row

def build_col_index_map(ws, header_row):
    """Строим словарь: Заголовок -> номер столбца (1-базовый)."""
    m = {}
    for idx, cell in enumerate(ws[header_row], start=1):
        if cell.value is not None:
            m[str(cell.value)] = idx
    return m

# ---------- Загрузка файлов ----------
file_order = st.file_uploader("Загрузите файл заказа (Excel/CSV)", type=["xls", "xlsx", "xlsm", "csv"], key="order")
file_supplier = st.file_uploader("Загрузите файл поставщика (Excel/CSV)", type=["xls", "xlsx", "xlsm", "csv"], key="supplier")

if file_order and file_supplier:
    order_bytes = file_order.getvalue()
    supplier_bytes = file_supplier.getvalue()

    df_order = load_table(order_bytes, file_order.name)
    df_supplier = load_table(supplier_bytes, file_supplier.name)

    if df_order is None or df_supplier is None:
        st.error("❌ Неподдерживаемый формат. Используйте .xls, .xlsx, .xlsm или .csv")
        st.stop()

    st.subheader("Наш файл заказа:")
    st.dataframe(df_order.head())

    st.subheader("Файл поставщика:")
    st.dataframe(df_supplier.head())

    # Выбор нескольких ключей и столбца количества
    order_keys = st.multiselect("🔑 Поля для поиска (наш файл, по очереди приоритетов)", df_order.columns)
    supplier_keys = st.multiselect("🔍 Поля для поиска (файл поставщика, в том же порядке)", df_supplier.columns)
    qty_col = st.selectbox("📦 Столбец с количеством (наш файл)", df_order.columns)
    supplier_qty_col = st.selectbox("✏️ Столбец для записи количества (файл поставщика)", df_supplier.columns)

    preserve_format = file_supplier.name.lower().endswith((".xlsx", ".xlsm"))
    if not preserve_format:
        st.info("ℹ️ Файл поставщика не в формате .xlsx/.xlsm — форматирование будет пересохранено стандартно.")

    auto_header = st.checkbox("Автоопределение строки заголовков (для .xlsx/.xlsm)", value=True, disabled=not preserve_format)
    manual_header = 1
    if preserve_format and not auto_header:
        manual_header = st.number_input("Номер строки заголовков на листе (1-базовый)", min_value=1, value=1, step=1)

    if len(order_keys) != len(supplier_keys):
        st.warning("⚠️ Выберите одинаковое количество ключевых столбцов в обоих файлах.")
        st.stop()

    # Кнопка запуска
    if st.button("Сформировать файл"):
        # Подготовка словарей для поиска по приоритетам
        dicts = []
        for okey in order_keys:
            d = {}
            for a, q in zip(df_order[okey], df_order[qty_col]):
                na = normalize_key(a)
                if na is not None:
                    d[na] = q
            dicts.append(d)

        total = len(df_supplier)
        stats = {}
        matched = 0

        if preserve_format:
            # Работаем через openpyxl, чтобы сохранить внешний вид
            wb = load_workbook(BytesIO(supplier_bytes))
            ws = wb.active  # по умолчанию активный лист

            header_row = detect_header_row(ws, df_supplier.columns) if auto_header else manual_header
            col_map = build_col_index_map(ws, header_row)

            # Проверим, что выбранные пользователем имена столбцов реально есть в найденной строке заголовков
            missing = [c for c in [*supplier_keys, supplier_qty_col] if str(c) not in col_map]
            if missing:
                st.error("❌ Эти столбцы не найдены на листе (в строке заголовков): " + ", ".join(map(str, missing)))
                st.stop()

            qty_col_idx = col_map[str(supplier_qty_col)]
            key_cols_idx = [col_map[str(c)] for c in supplier_keys]

            # Проходим по строкам после заголовка
            for r in range(header_row + 1, ws.max_row + 1):
                # Если уже есть значение в колонке количества — пропускаем
                if ws.cell(row=r, column=qty_col_idx).value is not None:
                    continue

                # Ищем по порядку ключей
                found_here = False
                for dict_idx, k_col_idx in enumerate(key_cols_idx, start=1):
                    key_val = normalize_key(ws.cell(row=r, column=k_col_idx).value)
                    if key_val is not None and key_val in dicts[dict_idx - 1]:
                        ws.cell(row=r, column=qty_col_idx).value = dicts[dict_idx - 1][key_val]
                        stats[f"Найдено по ключу {dict_idx} ({order_keys[dict_idx-1]} -> {supplier_keys[dict_idx-1]})"] =                                 stats.get(f"Найдено по ключу {dict_idx} ({order_keys[dict_idx-1]} -> {supplier_keys[dict_idx-1]})", 0) + 1
                        matched += 1
                        found_here = True
                        break

            not_found = total - matched

            # Вывод статистики
            st.subheader("📊 Результаты обработки")
            st.write(f"Всего строк у поставщика: **{total}**")
            for k in sorted(stats.keys()):
                st.write(f"🔍 {k}: **{stats[k]}**")
            st.write(f"⚠️ Не найдено: **{not_found}**")

            # Сохраняем в память
            out = BytesIO()
            wb.save(out)
            st.success("✅ Готово! Скачайте результат ниже.")
            st.download_button(
                "⬇ Скачать Excel (с сохранением формата)",
                data=out.getvalue(),
                file_name="supplier_with_qty.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        else:
            # Фоллбэк через pandas (форматирование не сохраняется)
            df = df_supplier.copy()
            df[supplier_qty_col] = None

            for dict_idx, (okey, skey) in enumerate(zip(order_keys, supplier_keys), start=1):
                d = dicts[dict_idx - 1]
                before = df[supplier_qty_col].notna().sum()
                df.loc[df[supplier_qty_col].isna(), supplier_qty_col] = (
                    df.loc[df[supplier_qty_col].isna(), skey].map(lambda x: d.get(normalize_key(x)))
                )
                after = df[supplier_qty_col].notna().sum()
                stats[f"Найдено по ключу {dict_idx} ({okey} -> {skey})"] = int(after - before)

            matched = int(df[supplier_qty_col].notna().sum())
            not_found = int(total - matched)

            # Статистика
            st.subheader("📊 Результаты обработки")
            st.write(f"Всего строк у поставщика: **{total}**")
            for k in sorted(stats.keys()):
                st.write(f"🔍 {k}: **{stats[k]}**")
            st.write(f"⚠️ Не найдено: **{not_found}**")

            out = BytesIO()
            df.to_excel(out, index=False, engine="openpyxl")
            st.success("✅ Готово! Скачайте результат ниже.")
            st.download_button(
                "⬇ Скачать Excel",
                data=out.getvalue(),
                file_name="supplier_with_qty.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
