
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
    name = filename.lower()
    if name.endswith(('.xlsx', '.xlsm')):
        return pd.read_excel(bio, engine='openpyxl')
    elif name.endswith('.xls'):
        return pd.read_excel(bio, engine='xlrd')
    elif name.endswith('.csv'):
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
    headers_set = {str(h) for h in headers if str(h) != 'Unnamed: 0'}
    best_row = 1
    best_hits = 0
    for r in range(1, min(ws.max_row, max_scan) + 1):
        row_vals = [str(c.value) if c.value is not None else '' for c in ws[r]]
        hits = sum(1 for v in row_vals if v in headers_set and v != '')
        if hits > best_hits:
            best_hits = hits
            best_row = r
            if hits >= max(1, len(headers_set) // 2):
                # достаточно совпадений, чтобы считать строку заголовком
                pass
    return best_row

def build_col_index_map(ws, header_row):
    """Строим словарь: Заголовок -> номер столбца (1-базовый)."""
    m = {}
    for idx, cell in enumerate(ws[header_row], start=1):
        if cell.value is not None:
            m[str(cell.value)] = idx
    return m

# ---------- Загрузка файлов ----------
file_order = st.file_uploader('Загрузите файл заказа (Excel/CSV)', type=['xls', 'xlsx', 'xlsm', 'csv'], key='order')
file_supplier = st.file_uploader('Загрузите файл поставщика (Excel/CSV)', type=['xls', 'xlsx', 'xlsm', 'csv'], key='supplier')

if file_order and file_supplier:
    order_bytes = file_order.getvalue()
    supplier_bytes = file_supplier.getvalue()

    df_order = load_table(order_bytes, file_order.name)
    df_supplier = load_table(supplier_bytes, file_supplier.name)

    if df_order is None or df_supplier is None:
        st.error('❌ Неподдерживаемый формат. Используйте .xls, .xlsx, .xlsm или .csv')
        st.stop()

    st.subheader('Наш файл заказа:')
    st.dataframe(df_order.head())

    st.subheader('Файл поставщика:')
    st.dataframe(df_supplier.head())

    # Выбор листа (для .xlsx/.xlsm с сохранением формата)
    preserve_format = file_supplier.name.lower().endswith(('.xlsx', '.xlsm'))
    sheet_name = None
    if preserve_format:
        with st.expander('Доп. настройки листа'):
            sheet_name = st.text_input('Имя листа (если пусто — активный)', value='')

    # Поля
    order_keys = st.multiselect('🔑 Поля для поиска (наш файл, по очереди приоритетов)', df_order.columns)
    supplier_keys = st.multiselect('🔍 Поля для поиска (файл поставщика, в том же порядке)', df_supplier.columns)
    qty_col = st.selectbox('📦 Столбец с количеством (наш файл)', df_order.columns)
    supplier_qty_col = st.selectbox('✏️ Столбец для записи количества (файл поставщика)', df_supplier.columns)

    # Поведение перезаписи
    overwrite_existing = st.checkbox('Перезаписывать уже заполненные значения', value=False)

    # Заголовки
    auto_header = st.checkbox('Автоопределение строки заголовков (для .xlsx/.xlsm)', value=True, disabled=not preserve_format)
    manual_header = 1
    if preserve_format and not auto_header:
        manual_header = st.number_input('Номер строки заголовков на листе (1-базовый)', min_value=1, value=1, step=1)

    if len(order_keys) != len(supplier_keys):
        st.warning('⚠️ Выберите одинаковое количество ключевых столбцов в обоих файлах.')
        st.stop()

    # Кнопка запуска
    if st.button('Сформировать файл'):
        # Подготовка словарей (по приоритетам)
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

        # Фоллбэк карта позиций из DataFrame
        supplier_pos_index_map = {str(col): idx + 1 for idx, col in enumerate(df_supplier.columns)}

        if preserve_format:
            keep_vba = file_supplier.name.lower().endswith('.xlsm')
            wb = load_workbook(BytesIO(supplier_bytes), keep_vba=keep_vba)
            ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active

            header_row = detect_header_row(ws, df_supplier.columns) if auto_header else manual_header
            col_map = build_col_index_map(ws, header_row)

            fallback_used = []

            def get_col_idx(col_name):
                s = str(col_name)
                if s in col_map:
                    return col_map[s]
                # Fallback по позиции, если заголовок пустой/Unnamed
                if s in supplier_pos_index_map:
                    idx = supplier_pos_index_map[s]
                    if idx <= ws.max_column:
                        fallback_used.append(s)
                        return idx
                return None

            qty_col_idx = get_col_idx(supplier_qty_col)
            key_cols_idx = [get_col_idx(c) for c in supplier_keys]

            missing = [n for n, i in zip(['QTY'] + list(supplier_keys), [qty_col_idx] + key_cols_idx) if i is None]
            if missing:
                st.error('❌ Не удалось сопоставить столбцы (даже по позиции). Проверьте выбор: ' + ', '.join(map(str, missing[1:])))
                st.stop()

            if fallback_used:
                st.warning('⚠️ Некоторые столбцы не имели заголовков в Excel и были сопоставлены по позиции: ' + ', '.join(fallback_used))

            # Проходим по строкам после заголовка
            start_row = header_row + 1
            for r in range(start_row, ws.max_row + 1):
                if not overwrite_existing and ws.cell(row=r, column=qty_col_idx).value is not None:
                    continue

                # Ищем по ключам по очереди
                assigned = False
                for dict_idx, k_col_idx in enumerate(key_cols_idx, start=1):
                    key_val = normalize_key(ws.cell(row=r, column=k_col_idx).value)
                    if key_val is not None:
                        val = dicts[dict_idx - 1].get(key_val)
                        if val is not None:
                            ws.cell(row=r, column=qty_col_idx).value = val
                            key_name = f"{order_keys[dict_idx-1]} -> {supplier_keys[dict_idx-1]}">
                            stats[key_name] = stats.get(key_name, 0) + 1
                            matched += 1
                            assigned = True
                            break

            not_found = total - matched

            # Статистика
            st.subheader('📊 Результаты обработки')
            st.write(f'Всего строк у поставщика: **{total}**')
            for k in stats:
                st.write(f'🔍 Найдено по {k}: **{stats[k]}**')
            st.write(f'⚠️ Не найдено: **{not_found}**')

            out = BytesIO()
            wb.save(out)
            st.success('✅ Готово! Скачайте результат ниже.')
            st.download_button(
                '⬇ Скачать Excel (с сохранением формата)',
                data=out.getvalue(),
                file_name='supplier_with_qty.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )

        else:
            # Фоллбэк через pandas
            df = df_supplier.copy()
            if not overwrite_existing:
                mask = df[supplier_qty_col].isna()
            else:
                mask = pd.Series([True] * len(df))

            for dict_idx, (okey, skey) in enumerate(zip(order_keys, supplier_keys), start=1):
                d = dicts[dict_idx - 1]
                before = df.loc[mask, supplier_qty_col].notna().sum()
                df.loc[mask, supplier_qty_col] = (
                    df.loc[mask, skey].map(lambda x: d.get(normalize_key(x)))
                )
                after = df.loc[mask, supplier_qty_col].notna().sum()
                stats[f'{okey} -> {skey}'] = int(after - before)

            matched = int(df[supplier_qty_col].notna().sum())
            not_found = int(total - matched)

            st.subheader('📊 Результаты обработки')
            st.write(f'Всего строк у поставщика: **{total}**')
            for k in stats:
                st.write(f'🔍 Найдено по {k}: **{stats[k]}**')
            st.write(f'⚠️ Не найдено: **{not_found}**')

            out = BytesIO()
            df.to_excel(out, index=False, engine='openpyxl')
            st.success('✅ Готово! Скачайте результат ниже.')
            st.download_button(
                '⬇ Скачать Excel',
                data=out.getvalue(),
                file_name='supplier_with_qty.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
