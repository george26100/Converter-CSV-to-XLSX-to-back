import io
import time
import pandas as pd
import streamlit as st
import plotly.express as px

# ---------- Конфигурация страницы ----------
st.set_page_config(
    page_title="CSV → XLSX Конвертор",
    page_icon="📊",
    layout="wide",
)

# ---------- CSS-анимации ----------
st.markdown(
    """
<style>
@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(30px); }
    to   { opacity: 1; transform: translateY(0); }
}
@keyframes pulseGlow {
    0%, 100% { box-shadow: 0 0 0 0 rgba(255,127,80,0.7); }
    50%      { box-shadow: 0 0 20px 8px rgba(255,127,80,0.25); }
}
@keyframes gradientShift {
    0%   { background-position: 0% 50%; }
    50%  { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}
.hero {
    background: linear-gradient(90deg, #ff7f50, #ff1493, #00bfff, #ff7f50);
    background-size: 300% 300%;
    animation: gradientShift 10s ease infinite;
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-weight: 800;
    font-size: 2.6rem;
    text-align: center;
    padding: 10px 0;
}
.fade { animation: fadeInUp 0.6s ease-out; }
.card {
    padding: 16px 20px;
    border-radius: 12px;
    background: rgba(255,127,80,0.08);
    border-left: 4px solid #ff7f50;
    animation: fadeInUp 0.5s ease-out;
    margin-bottom: 12px;
}
.pulse { animation: pulseGlow 2s infinite; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown('<h1 class="hero">📊 CSV → XLSX Конвертор</h1>', unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center;opacity:.75;'>Загрузите файл → отредактируйте → скачайте XLSX</p>",
    unsafe_allow_html=True,
)

# ---------- Session state ----------
for k, v in {
    "df": None,
    "filename": "data",
    "loaded_name": None,
    "history": [],
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


def load_file(file, name: str):
    """Загружает CSV или XLSX с прогресс-анимацией."""
    progress = st.progress(0, text="🔄 Читаем файл...")
    for i in range(100):
        time.sleep(0.004)
        progress.progress(i + 1)
    try:
        if name.lower().endswith(".csv"):
            try:
                df = pd.read_csv(file, encoding="utf-8")
            except UnicodeDecodeError:
                file.seek(0)
                df = pd.read_csv(file, encoding="cp1251")
        else:  # xlsx / xls
            df = pd.read_excel(file)
    finally:
        progress.empty()
    return df


def push_history(df: pd.DataFrame):
    """Храним до 10 снимков для отмены."""
    st.session_state.history.append(df.copy())
    if len(st.session_state.history) > 10:
        st.session_state.history.pop(0)


# ---------- Загрузка файла ----------
uploaded = st.file_uploader(
    "Загрузите CSV или XLSX",
    type=["csv", "xlsx", "xls"],
    help="CSV в UTF-8 или CP1251, XLSX/XLS через pandas/openpyxl",
)

if uploaded is not None and st.session_state.loaded_name != uploaded.name:
    df = load_file(uploaded, uploaded.name)
    st.session_state.df = df
    st.session_state.loaded_name = uploaded.name
    st.session_state.filename = uploaded.name.rsplit(".", 1)[0]
    st.session_state.history = []

# ---------- Основной блок ----------
if st.session_state.df is not None:
    df = st.session_state.df

    # Статистика (анимированные карточки)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f'<div class="card fade"><b>Строк</b><br><span style="font-size:1.6rem">{len(df)}</span></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="card fade"><b>Столбцов</b><br><span style="font-size:1.6rem">{len(df.columns)}</span></div>',
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f'<div class="card fade"><b>Пустых ячеек</b><br><span style="font-size:1.6rem">{int(df.isna().sum().sum())}</span></div>',
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f'<div class="card fade"><b>Числовых столбцов</b><br><span style="font-size:1.6rem">{len(df.select_dtypes("number").columns)}</span></div>',
            unsafe_allow_html=True,
        )

    # ---------- Sidebar: редактирование ----------
    with st.sidebar:
        st.header("🎛️ Редактирование")

        # -- Сортировка --
        st.subheader("↕️ Сортировка")
        sort_col = st.selectbox("Столбец", ["—"] + list(df.columns), key="sort_col")
        order = st.radio(
            "Порядок",
            ["По возрастанию", "По убыванию"],
            horizontal=True,
            key="sort_order",
        )
        if sort_col != "—" and st.button("Применить сортировку", use_container_width=True):
            push_history(df)
            try:
                df_sorted = df.sort_values(
                    by=sort_col,
                    ascending=(order == "По возрастанию"),
                    na_position="last",
                ).reset_index(drop=True)
                st.session_state.df = df_sorted
                st.toast(f"✅ Отсортировано по «{sort_col}»", icon="↕️")
                time.sleep(0.3)
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка сортировки: {e}")

        st.divider()

        # -- Удаление столбцов --
        st.subheader("🗑️ Удаление столбцов")
        cols_to_del = st.multiselect(
            "Выберите столбцы",
            options=list(df.columns),
            key="cols_del",
        )
        if cols_to_del and st.button("Удалить столбцы", use_container_width=True):
            if len(cols_to_del) >= len(df.columns):
                st.error("Нельзя удалить все столбцы")
            else:
                push_history(df)
                st.session_state.df = df.drop(columns=cols_to_del)
                st.toast(f"🗑️ Удалено столбцов: {len(cols_to_del)}", icon="🗑️")
                time.sleep(0.3)
                st.rerun()

        st.divider()

        # -- Удаление строк --
        st.subheader("🗑️ Удаление строк")
        rows_to_del = st.multiselect(
            "Индексы строк (0, 1, 2 …)",
            options=list(df.index),
            format_func=lambda x: f"Строка {x}",
            key="rows_del",
        )
        if rows_to_del and st.button("Удалить строки", use_container_width=True):
            if len(rows_to_del) >= len(df):
                st.error("Нельзя удалить все строки")
            else:
                push_history(df)
                st.session_state.df = df.drop(index=rows_to_del).reset_index(drop=True)
                st.toast(f"🗑️ Удалено строк: {len(rows_to_del)}", icon="🗑️")
                time.sleep(0.3)
                st.rerun()

        st.divider()

        # -- Отмена / сброс --
        colA, colB = st.columns(2)
        with colA:
            if st.button("↩️ Отменить", use_container_width=True, disabled=not st.session_state.history):
                st.session_state.df = st.session_state.history.pop()
                st.toast("↩️ Изменение отменено", icon="↩️")
                time.sleep(0.3)
                st.rerun()
        with colB:
            if st.button("🔄 Сбросить", use_container_width=True):
                st.session_state.df = None
                st.session_state.loaded_name = None
                st.session_state.history = []
                st.rerun()

    df = st.session_state.df  # берём актуальное

    # ---------- Вкладки ----------
    tab_table, tab_chart, tab_export = st.tabs(
        ["📋 Таблица", "📈 Визуализация", "💾 Экспорт в XLSX"]
    )

    # ---- Таблица ----
    with tab_table:
        st.dataframe(
            df,
            use_container_width=True,
            height=520,
            hide_index=False,
        )

    # ---- Визуализация ----
    with tab_chart:
        numeric = df.select_dtypes(include="number").columns.tolist()
        if not numeric:
            st.info("Нет числовых столбцов — постройте визуализацию после конвертации типов.")
        else:
            cc1, cc2, cc3 = st.columns(3)
            with cc1:
                chart_kind = st.selectbox(
                    "Тип графика", ["Линия", "Столбцы", "Точки", "Площадь"]
                )
            with cc2:
                x_col = st.selectbox("Ось X", df.columns.tolist())
            with cc3:
                y_col = st.selectbox("Ось Y", numeric)

            chart_map = {
                "Линия": px.line,
                "Столбцы": px.bar,
                "Точки": px.scatter,
                "Площадь": px.area,
            }
            try:
                fig = chart_map[chart_kind](df, x=x_col, y=y_col, markers=True)
                fig.update_layout(
                    template="plotly_white",
                    transition_duration=500,
                    height=520,
                    margin=dict(l=10, r=10, t=30, b=10),
                )
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Не удалось построить график: {e}")

    # ---- Экспорт ----
    with tab_export:
        st.markdown("### 📥 Конвертация в XLSX")
        st.write(
            "Файл будет сохранён с авто-шириной столбцов и правильным форматом дат."
        )

        if st.button("🚀 Конвертировать в XLSX", type="primary", use_container_width=True):
            progress = st.progress(0, text="🔨 Конвертация...")
            buf = io.BytesIO()
            try:
                for i in range(100):
                    time.sleep(0.006)
                    progress.progress(i + 1)

                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Data")
                    ws = writer.sheets["Data"]
                    for col in ws.columns:
                        max_len = max(
                            (len(str(c.value)) for c in col if c.value is not None),
                            default=8,
                        )
                        ws.column_dimensions[col[0].column_letter].width = min(
                            max_len + 2, 60
                        )
                    ws.freeze_panes = "A2"

                buf.seek(0)
                st.session_state["xlsx_bytes"] = buf.getvalue()
                progress.empty()
                st.success("✅ Конвертация завершена!")
                st.balloons()
            except Exception as e:
                progress.empty()
                st.error(f"Ошибка: {e}")

        if st.session_state.get("xlsx_bytes"):
            st.download_button(
                "⬇️ Скачать XLSX",
                data=st.session_state["xlsx_bytes"],
                file_name=f"{st.session_state.filename}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

else:
    # Плейсхолдер до загрузки
    st.markdown(
        """
        <div style="text-align:center; padding:60px 20px; opacity:.8" class="fade">
            <div style="font-size:5rem" class="pulse">📄</div>
            <h3>Загрузите CSV или XLSX файл выше</h3>
            <p>После загрузки вы сможете:</p>
            <ul style="display:inline-block;text-align:left">
              <li>↕️ Сортировать по любому столбцу</li>
              <li>🗑️ Удалять строки и столбцы</li>
              <li>📈 Строить интерактивные графики</li>
              <li>💾 Скачать результат в XLSX</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )