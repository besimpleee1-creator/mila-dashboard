import hashlib
import io
import os
import re
import sys
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

# Google Таблица (доступ «по ссылке — просмотр»).
DEFAULT_SHEET_ID = "179x6oJAD_rm3b3kOgtfkbX_J-jrh9a49S_46d_no27k"
DEFAULT_SRC = os.environ.get(
    "MILA_SHEET_URL",
    f"https://docs.google.com/spreadsheets/export?format=xlsx&id={DEFAULT_SHEET_ID}",
)
# Как часто проверять источник на изменения (секунды).
WATCH_INTERVAL = int(os.environ.get("MILA_WATCH_INTERVAL", "15"))

DAY_NAME_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
# Яркая палитра для сотрудников и графиков.
PALETTE = [
    "#6366F1", "#8B5CF6", "#EC4899", "#F59E0B",
    "#10B981", "#3B82F6", "#EF4444", "#14B8A6",
    "#F97316", "#06B6D4", "#A855F7", "#84CC16",
]
KPI_META = [
    ("⌛", "Всего часов", "#3B82F6"),
    ("💰", "Сумма выплат", "#10B981"),
    ("🧑‍🤝‍🧑", "Активных сотрудников", "#8B5CF6"),
    ("🎯", "Средняя ставка", "#F59E0B"),
    ("📈", "Средняя выработка", "#EC4899"),
]


def log(msg: str) -> None:
    print(msg, flush=True)


# ─────────────────────────────────────────────────────────────
# 1. ЗАГРУЗКА ДАННЫХ: URL (Google/файл по ссылке) или локальный путь
# ─────────────────────────────────────────────────────────────
def fetch_bytes(src: str) -> bytes:
    """Скачивает книгу по URL (HTTP) или читает локальный файл.

    Для Google-таблиц умеет подбирать рабочую форму адреса экспорта.
    """
    if not src.lower().startswith(("http://", "https://")):
        with open(src, "rb") as fh:
            return fh.read()

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 "
            "Safari/537.36"
        )
    }

    urls = [src]
    sheet_id = ""
    m1 = re.search(r"/d/([\w-]{20,})", src, re.IGNORECASE)
    m2 = re.search(r"[?&]id=([\w-]{20,})", src)
    if m1:
        sheet_id = m1.group(1)
    elif m2:
        sheet_id = m2.group(1)
    if sheet_id:
        alt = f"https://docs.google.com/spreadsheets/export?format=xlsx&id={sheet_id}"
        if alt != src:
            urls.append(alt)

    last_err = None
    for url in urls:
        try:
            log(f"[fetch] GET {url}")
            r = requests.get(url, headers=headers, timeout=90)
            log(f"[fetch] status={r.status_code} ctype={r.headers.get('Content-Type')}")
            r.raise_for_status()
            data = r.content
            log(f"[fetch] len={len(data)} magic={data[:4]!r}")
            if data[:2] == b"PK":
                return data
            raise ValueError(
                "сервер вернул не книгу Excel, а другую страницу "
                "(возможно, страницу входа Google)"
            )
        except Exception as exc:
            log(f"[fetch] attempt failed: {type(exc).__name__}: {exc}")
            last_err = exc
    raise RuntimeError(f"Не удалось скачать книгу по адресу: {last_err}")


def parse_workbook(data: bytes) -> pd.DataFrame:
    """Разбирает все недельные блоки во всех листах книги.

    Помимо часов запоминает отметки «*» (выходной/пропуск) и факт работы.
    """
    xls = pd.ExcelFile(io.BytesIO(data))
    records = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        i = 0
        while i < len(df):
            row = df.iloc[i]
            # Шапка недели: в столбце D стоит «Дата»
            if pd.notna(row[3]) and str(row[3]).strip() == "Дата" and i + 1 < len(df):
                date_row = df.iloc[i + 1]
                dates = [
                    pd.to_datetime(date_row[col]) if pd.notna(date_row[col]) else pd.NaT
                    for col in range(4, 11)  # E..K
                ]

                j = i + 2
                while j < len(df):
                    emp = df.iloc[j]
                    name, rate = emp[1], emp[2]

                    if pd.notna(emp[3]) and str(emp[3]).strip() == "Дата":
                        break

                    if pd.notna(name) and pd.notna(rate):
                        for k, d in enumerate(dates):
                            if pd.isna(d):
                                continue
                            raw = emp[4 + k]
                            if pd.isna(raw) or str(raw).strip() == "*":
                                if pd.isna(raw):
                                    continue
                                records.append({
                                    "Date": d,
                                    "Employee": str(name).strip(),
                                    "Rate": float(rate),
                                    "Hours": 0.0,
                                    "Sheet": sheet,
                                    "Star": True,
                                    "Worked": False,
                                })
                                continue
                            try:
                                hours = float(raw)
                            except Exception:
                                continue
                            records.append({
                                "Date": d,
                                "Employee": str(name).strip(),
                                "Rate": float(rate),
                                "Hours": hours,
                                "Sheet": sheet,
                                "Star": False,
                                "Worked": True,
                            })
                    j += 1
                i = j
            else:
                i += 1

    df = pd.DataFrame(records)
    log(f"[parse] sheets={len(xls.sheet_names)} records={len(records)}")
    if df.empty:
        return df

    df["Employee"] = df["Employee"].replace("", np.nan).fillna("Без имени")
    df = df.drop_duplicates(subset=["Date", "Employee"], keep="first")
    df["Amount"] = df["Hours"] * df["Rate"]
    df["WeekStart"] = df["Date"] - pd.to_timedelta(df["Date"].dt.dayofweek, unit="D")
    df["DayOfWeek"] = df["Date"].dt.dayofweek
    return df


def data_hash(df: pd.DataFrame) -> str:
    """Стабильный отпечаток ДАННЫХ (одинаковые данные = одинаковый хэш).

    Сам файл от Google каждый раз чуть отличается (служебные поля), поэтому
    сравнивать файлы нельзя — только сами данные.
    """
    if df is None or df.empty:
        return "empty"
    return hashlib.sha1(df.to_csv(index=False).encode("utf-8")).hexdigest()


def load_df(src: str) -> pd.DataFrame:
    """Скачивает книгу и разбирает её в таблицу."""
    data = fetch_bytes(src)
    return parse_workbook(data)


# ─────────────────────────────────────────────────────────────
# 2. АВТО-ОБНОВЛЕНИЕ ПРИ ИЗМЕНЕНИИ ТАБЛИЦЫ
# ─────────────────────────────────────────────────────────────
@st.fragment(run_every=WATCH_INTERVAL)
def watch_source(src: str):
    """Каждые N секунд скачивает книгу; при изменении ДАННЫХ — перерисовка всего приложения."""
    last = st.session_state.get("mila_polled_at")
    if last is not None:
        elapsed = (datetime.now() - last).total_seconds()
        if elapsed < WATCH_INTERVAL / 2:
            return
    try:
        df = load_df(src)
    except Exception:
        return
    st.session_state["mila_polled_at"] = datetime.now()
    h = data_hash(df)
    if st.session_state.get("mila_data_hash") != h:
        st.session_state["mila_df"] = df
        st.session_state["mila_data_hash"] = h
        st.session_state["mila_fetched_at"] = datetime.now()
        st.rerun(scope="app")


# ─────────────────────────────────────────────────────────────
# 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ОФОРМЛЕНИЯ
# ─────────────────────────────────────────────────────────────
def fmt_hours(x: float) -> str:
    return f"{x:,.1f}".replace(",", " ") + " ч"


def fmt_rub(x: float) -> str:
    return f"{x:,.0f}".replace(",", " ") + " ₽"


CSS = """
<style>
    .stApp { background: #F2F4FA; }

    .hero {
        background: linear-gradient(120deg, #6366F1 0%, #8B5CF6 55%, #EC4899 100%);
        border-radius: 20px; padding: 26px 30px; color: #ffffff;
        box-shadow: 0 8px 24px rgba(99,102,241,.35); margin-bottom: 18px;
    }
    .hero h1 { color: #ffffff; margin: 0 0 4px 0; font-size: 28px; letter-spacing: .2px; }
    .hero p  { color: rgba(255,255,255,.92); margin: 0; font-size: 14px; }
    .hero .date-chip {
        display: inline-block; background: rgba(255,255,255,.18);
        border: 1px solid rgba(255,255,255,.35); padding: 3px 12px;
        border-radius: 20px; font-size: 13px; margin-top: 10px;
    }

    .kpi {
        background: #FFFFFF; border-radius: 16px; padding: 14px 18px;
        border: 1px solid #E7EAF5; height: 100%;
        box-shadow: 0 3px 10px rgba(31,41,85,.06); position: relative;
    }
    .kpi .ico { font-size: 22px; line-height: 1; }
    .kpi .lbl { font-size: 12px; color: #6B7280; letter-spacing: .4px;
                text-transform: uppercase; margin-top: 6px; white-space: nowrap; }
    .kpi .val { font-size: 25px; font-weight: 800; margin-top: 2px; }
    .kpi .sub { font-size: 12px; color: #6B7280; margin-top: 4px; }
    .kpi .bar  { height: 4px; border-radius: 4px; margin-top: 10px; }

    .lboard-item {
        display: flex; align-items: center; gap: 12px; padding: 9px 12px;
        background: #FFFFFF; border: 1px solid #E7EAF5; border-radius: 12px;
        margin-bottom: 8px; box-shadow: 0 2px 6px rgba(31,41,85,.05);
    }
    .lboard-medal { font-size: 20px; width: 32px; text-align: center; }
    .lboard-name { flex: 1; font-weight: 600; color: #1F2937; }
    .lboard-hours { font-size: 15px; font-weight: 700; color: #1F2937; }
    .lboard-sum { font-size: 12px; color: #6B7280; min-width: 76px; text-align: right; }
    .lboard-track { position: relative; height: 6px; border-radius: 6px;
                    background: #EEF0F8; overflow: hidden; margin-top: 6px; }

    .cap { color: #6B7280; font-size: 13px; margin: 4px 0 12px 0; }
    h1, h2, h3 { color: #1E2A5A; }
    div[data-testid="stTabs"] button p { font-weight: 600; font-size: 15px; }
</style>
"""


def kpi_card(icon: str, label: str, value: str, sub: str, color: str) -> str:
    return f"""
    <div class="kpi">
        <div class="ico">{icon}</div>
        <div class="lbl">{label}</div>
        <div class="val" style="color:{color}">{value}</div>
        <div class="sub">{sub}</div>
        <div class="bar" style="background:linear-gradient(90deg,{color},transparent)"></div>
    </div>"""


def leaderboard_html(pairs: list, color: str, unit: str = "часов") -> str:
    """pairs: [(name, primary_value, secondary_text)] — top-N."""
    if not pairs:
        return "<div class='cap'>Нет данных по фильтрам.</div>"
    medals = ["🥇", "🥈", "🥉"] + [str(i) for i in range(4, len(pairs) + 1)]
    mx = max(p[1] for p in pairs) or 1
    rows = []
    for i, (name, val, sec) in enumerate(pairs):
        w = max(4, round(100 * val / mx))
        rows.append(
            f"""<div class="lboard-item">
                    <div class="lboard-medal">{medals[i]}</div>
                    <div style="flex:1">
                        <div style="display:flex;align-items:center;gap:8px">
                            <span class="lboard-name">{name}</span>
                            <span class="lboard-hours">{val:,.1f}</span>
                            <span style="font-size:11px;color:#9CA3AF">{unit}</span>
                        </div>
                        <div class="lboard-track">
                            <div style="width:{w}%;height:100%;border-radius:6px;
                                       background:{color}"></div>
                        </div>
                    </div>
                    <div class="lboard-sum">{sec}</div>
                </div>"""
        )
    return "".join(rows)


# ─────────────────────────────────────────────────────────────
# 4. СТРАНИЦА
# ─────────────────────────────────────────────────────────────
log(f"[run] start script, src={DEFAULT_SRC[:85]}")
st.set_page_config(page_title="Мила — Статистика", page_icon="📊", layout="wide")

st.markdown(CSS, unsafe_allow_html=True)

# ── Необязательная защита кодом (если задан секрет mila_pin) ──
try:
    ACCESS_PIN = st.secrets.get("mila_pin")
except Exception:
    ACCESS_PIN = None

if ACCESS_PIN:
    if "unlocked" not in st.session_state:
        st.session_state["unlocked"] = False
    if not st.session_state["unlocked"]:
        st.title("🔒 Доступ ограничен")
        code = st.text_input("Введите код доступа", type="password")
        if st.button("Войти"):
            if code == ACCESS_PIN:
                st.session_state["unlocked"] = True
                st.rerun()
            else:
                st.error("Неверный код")
        st.stop()

# ── Источник данных ──────────────────────────────────────────
src = st.session_state.get("mila_src", DEFAULT_SRC)

with st.sidebar:
    st.header("⚙️ Источник данных")
    st.text_input(
        "Ссылка на Google-таблицу (или путь к xlsx)",
        value=src, key="mila_src",
    )
    src = st.session_state["mila_src"]

    # Скачиваем и разбираем книгу при старте или при смене адреса
    if st.session_state.get("mila_src_state") != src:
        try:
            df = load_df(src)
            st.session_state["mila_df"] = df
            st.session_state["mila_data_hash"] = data_hash(df)
            st.session_state["mila_src_state"] = src
            st.session_state["mila_fetched_at"] = datetime.now()
        except Exception as exc:
            log(f"[error] initial load failed: {type(exc).__name__}: {exc}")
            st.session_state.pop("mila_df", None)
            st.session_state["mila_src_state"] = src
            st.error("Не удалось загрузить книгу по этому адресу.")
            st.caption("Проверьте, что файл открыт «по ссылке» для просмотра, "
                       "и что адрес начинается с http://")
            st.stop()

    watch_source(src)

    fetched = st.session_state.get("mila_fetched_at")
    st.caption(
        f"⚡ Проверка изменений каждые {WATCH_INTERVAL} с\n\n"
        f"Последнее обновление: {fetched:%d.%m.%Y %H:%M:%S}"
    )
    if st.button("🔄 Обновить сейчас", use_container_width=True):
        st.session_state["mila_src_state"] = None
        st.rerun()

    st.markdown("---")
    st.header("🔎 Фильтры")

# ── Данные ────────────────────────────────────────────────────
df = st.session_state.get("mila_df")
if df is None:
    st.error(
        "Нет данных. Откройте страницу, перезагрузите её или нажмите "
        "«🔄 Обновить сейчас»."
    )
    st.stop()

log(f"[run] df shape={df.shape} employees={df['Employee'].nunique()} "
    f"dates={df['Date'].min().date()}..{df['Date'].max().date()}")

if df.empty:
    st.warning("Не удалось найти данные в книге — проверьте разметку листов.")
    st.stop()

dmin, dmax = df["Date"].min().date(), df["Date"].max().date()

with st.sidebar:
    period = st.date_input(
        "Период", [dmin, dmax], min_value=dmin, max_value=dmax
    )
    employees = st.multiselect(
        "Сотрудники",
        sorted(df["Employee"].unique()),
        default=sorted(df["Employee"].unique()),
    )
    day_map = {"Пн": 0, "Вт": 1, "Ср": 2, "Чт": 3, "Пт": 4, "Сб": 5, "Вс": 6}
    days = st.multiselect("Дни недели", list(day_map), default=list(day_map))

if period and len(period) >= 2:
    p0, p1 = period[0], period[1]
else:
    p0, p1 = dmin, dmax

sel_days = [day_map[d] for d in days]
mask = (
    (df["Date"].dt.date >= p0)
    & (df["Date"].dt.date <= p1)
    & (df["Employee"].isin(employees))
    & (df["DayOfWeek"].isin(sel_days))
)
f = df[mask].copy()

if f.empty:
    st.info("По выбранным фильтрам данных нет — измените период, сотрудников или дни.")
    st.stop()

# ── Шапка ─────────────────────────────────────────────────────
st.markdown(
    f"""
    <div class="hero">
        <h1>📊 Мила — статистика часов и выплат</h1>
        <p>Живой дашборд по данным таблицы: кто и сколько работал, сколько это стоило.</p>
        <span class="date-chip">🗓 {p0:%d.%m.%Y} — {p1:%d.%m.%Y}</span>
        <span class="date-chip">👥 {f['Employee'].nunique()} сотрудников в выборке</span>
        <span class="date-chip">🔴 LIVE · автообновление каждые {WATCH_INTERVAL} с</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── KPI ───────────────────────────────────────────────────────
total_h = f["Hours"].sum()
total_a = f["Amount"].sum()
n_emp = f["Employee"].nunique()
n_days = f["Date"].nunique()
avg_rate = f.groupby("Employee")["Rate"].first().mean() if n_emp else 0
avg_per_emp = total_h / n_emp if n_emp else 0
avg_per_day = total_h / n_days if n_days else 0

kpi_values = [
    (KPI_META[0], (fmt_hours(total_h),
     f"{n_days} дн. с данными · в ср. {fmt_hours(avg_per_day)}/день")),
    (KPI_META[1], (fmt_rub(total_a),
     f"≈ {fmt_rub(total_a / n_days)} в день")),
    (KPI_META[2], (str(n_emp),
     f"из {df['Employee'].nunique()} в базе таблицы")),
    (KPI_META[3], (fmt_rub(avg_rate),
     f"ставки {fmt_rub(f.groupby('Employee')['Rate'].first().min())} – "
     f"{fmt_rub(f.groupby('Employee')['Rate'].first().max())}")),
    (KPI_META[4], (fmt_hours(avg_per_emp),
     f"на одного сотрудника за период")),
]

cols = st.columns(len(kpi_values))
ic, ok = 0, 0
for col, ((icon, label, color), (value, sub)) in zip(cols, kpi_values):
    with col:
        st.markdown(kpi_card(icon, label, value, sub, color), unsafe_allow_html=True)

st.markdown("<div class='cap'>Все показатели — за выбранный период и по выбранным фильтрам.</div>",
            unsafe_allow_html=True)
st.markdown("---")

# ── Вкладки ───────────────────────────────────────────────────
tabs = st.tabs(["📊 Обзор", "🧑‍💻 По людям", "📅 По календарю", "🚫 Пропуски (*)"])

# ════════════════ ОБЗОР ════════════════
with tabs[0]:
    c1, c2 = st.columns([3, 2])

    with c1:
        st.subheader("📈 Часы и выплаты по неделям")
        weekly = (
            f.groupby("WeekStart")
            .agg(hours=("Hours", "sum"), amount=("Amount", "sum"))
            .reset_index()
            .sort_values("WeekStart")
        )
        if len(weekly) >= 2:
            last = weekly.iloc[-1]
            prev = weekly.iloc[-2]
            delta = (last["hours"] - prev["hours"]) / prev["hours"] * 100 if prev["hours"] else 0
            sign = "▲" if delta >= 0 else "▼"
            st.caption(
                f"Неделя {last['WeekStart'].date():%d.%m}: {fmt_hours(last['hours'])} "
                f"({sign} {abs(delta):.0f}% к неделе {prev['WeekStart'].date():%d.%m})"
            )
        fig = go.Figure()
        fig.add_bar(x=weekly["WeekStart"], y=weekly["hours"],
                    name="Часы", marker_color="#6366F1", marker_line_width=0)
        fig.add_scatter(x=weekly["WeekStart"], y=weekly["amount"],
                        name="Выплаты", mode="lines+markers",
                        yaxis="y2", line=dict(color="#10B981", width=3))
        fig.update_layout(
            barmode="group", paper_bgcolor="#F2F4FA", plot_bgcolor="#FFFFFF",
            font=dict(size=12), margin=dict(l=10, r=10, t=10, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
            yaxis=dict(title="Часы", gridcolor="#EEF0F8"),
            yaxis2=dict(title="Выплаты, ₽", overlaying="y", side="right",
                        showgrid=False),
        )
        fig.update_xaxes(tickformat="%d.%m", gridcolor="#EEF0F8")
        st.plotly_chart(fig, width="stretch")
        st.markdown(
            "<div class='cap'>Столбики — суммарные часы за неделю, линия — выплаты "
            "(правая ось). Видно, растёт ли нагрузка и расходы неделя к неделе.</div>",
            unsafe_allow_html=True,
        )

    with c2:
        st.subheader("🍩 Доля выплат по людям")
        donut = f.groupby("Employee")["Amount"].sum().sort_values()
        fig = px.pie(values=donut.values, names=donut.index, hole=0.55,
                     color_discrete_sequence=PALETTE)
        fig.update_traces(textinfo="percent", textfont_size=12)
        fig.update_layout(
            paper_bgcolor="#F2F4FA", showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, font=dict(size=11)),
            margin=dict(l=10, r=10, t=10, b=10),
        )
        st.plotly_chart(fig, width="stretch")
        st.markdown(
            "<div class='cap'>Кому ушло больше всего денег за период.</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    c3, c4 = st.columns([1, 1])
    with c3:
        st.subheader("🏆 Топ сотрудников по часам")
        top = (
            f.groupby("Employee").agg(h=("Hours", "sum"), s=("Amount", "sum"))
            .sort_values("h", ascending=False)
        )
        pairs = [
            (emp, row["h"], fmt_rub(row["s"]))
            for emp, row in top.head(8).iterrows()
        ]
        st.markdown(leaderboard_html(pairs, "#6366F1"), unsafe_allow_html=True)

    with c4:
        st.subheader("🌡 Нагрузка по дням недели")
        days_agg = (
            f.groupby("DayOfWeek")["Hours"].sum().reindex(range(7), fill_value=0)
        )
        colors = [PALETTE[i % len(PALETTE)] for i in range(7)]
        fig = px.bar(x=[DAY_NAME_RU[d] for d in range(7)], y=days_agg.values,
                     color=[DAY_NAME_RU[d] for d in range(7)],
                     color_discrete_sequence=colors)
        fig.update_traces(marker_line_width=0, showlegend=False, width=0.6)
        fig.update_layout(
            paper_bgcolor="#F2F4FA", plot_bgcolor="#FFFFFF",
            font=dict(size=12), margin=dict(l=10, r=10, t=10, b=10),
            yaxis=dict(title="Часы", gridcolor="#EEF0F8"), xaxis_title="",
        )
        st.plotly_chart(fig, width="stretch")
        st.markdown(
            "<div class='cap'>В какие дни недели приходится пик работы.</div>",
            unsafe_allow_html=True,
        )

# ════════════════ ПО ЛЮДЯМ ════════════════
with tabs[1]:
    st.subheader("🧑‍💻 Сводка по сотрудникам")
    person = (
        f.groupby("Employee")
        .agg(
            Дней=("Date", "nunique"),
            Часы=("Hours", "sum"),
            Сумма=("Amount", "sum"),
            Ставка=("Rate", "first"),
        )
        .reset_index()
        .sort_values("Часы", ascending=False)
    )
    person["Ср_в_день"] = person["Часы"] / person["Дней"].clip(lower=1)
    person["Доля_выплат"] = 100 * person["Сумма"] / person["Сумма"].sum()
    person["Часы"] = person["Часы"].round(1)
    person["Сумма"] = person["Сумма"].round(0)
    person["Ставка"] = person["Ставка"].round(0)
    person["Ср_в_день"] = person["Ср_в_день"].round(1)

    styled = person.rename(columns={"Ср_в_день": "Ср.ч/день", "Доля_выплат": "Доля, %"})
    st.dataframe(
        styled[["Employee", "Дней", "Часы", "Сумма", "Ставка", "Ср.ч/день", "Доля, %"]]
        .rename(columns={"Employee": "Сотрудник", "Дней": "Дней работы",
                         "Часы": "Часы, ч", "Сумма": "Сумма, ₽",
                         "Ставка": "Ставка, ₽/ч"}),
        width="stretch", hide_index=True,
        column_config={
            "Доля, %": st.column_config.ProgressColumn(
                "Доля выплат, %", min_value=0, max_value=100, format="%.1f%%"
            ),
        },
    )
    st.markdown(
        "<div class='cap'>Кто больше всех работает и сколько это стоит. "
        "Столбец «Доля выплат» — процент от всех выплат за период.</div>",
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.subheader("🔍 Досье сотрудника")
    pick = st.selectbox(
        "Выберите сотрудника",
        sorted(f["Employee"].unique()),
        key="mila_pick",
    )
    person_df = f[f["Employee"] == pick]
    days_n = person_df["Date"].nunique()
    h_sum = person_df["Hours"].sum()
    a_sum = person_df["Amount"].sum()
    rate = person_df["Rate"].iloc[0]
    stars = person_df["Star"].sum()
    d1, d2, d3, d4, d5 = st.columns(5)
    stats = [
        ("⌛", "Часов", fmt_hours(h_sum), "#3B82F6"),
        ("💰", "Выплат", fmt_rub(a_sum), "#10B981"),
        ("🎯", "Ставка", fmt_rub(rate), "#F59E0B"),
        ("📅", "Дней работы", str(days_n), "#8B5CF6"),
        ("🚫", "Пропусков (*)", str(int(stars)), "#EF4444"),
    ]
    for col, (icon, lbl, val, color) in zip([d1, d2, d3, d4, d5], stats):
        with col:
            st.markdown(kpi_card(icon, lbl, val, "", color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    cA, cB = st.columns([3, 2])
    with cA:
        st.markdown("**📈 Часы по неделям**")
        emp_week = (
            person_df.groupby("WeekStart")["Hours"].sum().reset_index()
        )
        fig = px.line(emp_week, x="WeekStart", y="Hours", markers=True,
                      color_discrete_sequence=["#8B5CF6"])
        fig.update_traces(fill="tozeroy", fillcolor="rgba(139,92,246,0.12)")
        fig.update_layout(
            paper_bgcolor="#F2F4FA", plot_bgcolor="#FFFFFF",
            yaxis=dict(title="Часы", gridcolor="#EEF0F8"), xaxis_title="",
            margin=dict(l=10, r=10, t=10, b=10), font=dict(size=12),
        )
        fig.update_xaxes(tickformat="%d.%m")
        st.plotly_chart(fig, width="stretch")
    with cB:
        st.markdown("**🗓 Распределение по дням недели**")
        emp_days = person_df.groupby("DayOfWeek")["Hours"].sum().reindex(
            range(7), fill_value=0
        )
        fig = px.pie(values=emp_days.values,
                     names=[DAY_NAME_RU[d] for d in range(7)], hole=0.5,
                     color_discrete_sequence=PALETTE)
        fig.update_layout(
            paper_bgcolor="#F2F4FA",
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, font=dict(size=11)),
            margin=dict(l=10, r=10, t=10, b=10),
        )
        st.plotly_chart(fig, width="stretch")

    st.markdown("---")
    st.markdown("**📋 Дни и суммы**")
    detail = (
        person_df[["Date", "DayOfWeek", "Hours", "Rate", "Amount", "Star"]]
        .sort_values("Date")
        .assign(День=lambda d: d["DayOfWeek"].map(lambda x: DAY_NAME_RU[x]))
        .assign(Отметка=lambda d: np.where(d["Star"], "🚫 пропуск", "✔ работа"))
        .drop(columns=["DayOfWeek", "Star"])
        .rename(columns={"Date": "Дата", "Hours": "Часы", "Rate": "Ставка",
                         "Amount": "Сумма"})
    )
    detail["Дата"] = detail["Дата"].dt.strftime("%d.%m.%Y")
    detail["Сумма"] = detail["Сумма"].map(lambda x: f"{x:,.0f} ₽")
    detail["Ставка"] = detail["Ставка"].map(lambda x: f"{x:,.0f} ₽")
    detail["Часы"] = detail["Часы"].map(lambda x: f"{x:,.1f}".replace(",", " ").strip())
    detail = detail.rename(columns={"Отметка": "Статус"})
    st.dataframe(detail, width="stretch", hide_index=True)

# ════════════════ ПО КАЛЕНДАРЮ ════════════════
with tabs[2]:
    st.subheader("🔥 Матрица «Сотрудники × Дни» (часы)")
    pivot = f.pivot_table(
        index="Employee", columns="Date", values="Hours",
        aggfunc="sum", fill_value=np.nan,
    )
    pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]

    fig = px.imshow(
        pivot,
        aspect="auto",
        color_continuous_scale=["#EEF0F8", "#A5B4FC", "#6366F1", "#312E81"],
        labels=dict(x="Дата", y="Сотрудник", color="Часы"),
    )
    fig.update_layout(
        height=max(260, 34 * len(pivot)),
        paper_bgcolor="#F2F4FA", plot_bgcolor="#FFFFFF",
        font=dict(size=12), margin=dict(l=10, r=10, t=10, b=10),
    )
    fig.update_xaxes(tickformat="%d.%m", tickangle=-45)
    st.plotly_chart(fig, width="stretch")
    st.markdown(
        "<div class='cap'>Чем темнее ячейка — тем больше часов. Пустая ячейка — "
        "нет данных в таблице.</div>",
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.subheader("🌡 Часы неделя × день")
    heat = f.pivot_table(
        index="DayOfWeek", columns="WeekStart", values="Hours",
        aggfunc="sum", fill_value=0,
    )
    heat.index = [DAY_NAME_RU[i] for i in heat.index.get_level_values(0)]
    fig = px.imshow(heat, aspect="auto",
                    color_continuous_scale=["#EEF0F8", "#FDE68A", "#F59E0B", "#EF4444"],
                    labels=dict(x="Неделя", y="День", color="Часы"))
    fig.update_layout(
        paper_bgcolor="#F2F4FA", plot_bgcolor="#FFFFFF",
        font=dict(size=12), margin=dict(l=10, r=10, t=10, b=10),
    )
    fig.update_xaxes(tickformat="%d.%m", tickangle=-45)
    st.plotly_chart(fig, width="stretch")
    st.markdown(
        "<div class='cap'>Как распределяется нагрузка по дням недели и неделям.</div>",
        unsafe_allow_html=True,
    )

# ════════════════ ПРОПУСКИ ════════════════
with tabs[3]:
    st.subheader("🚫 Пропуски и отметки «*»")
    if "Star" in f.columns and f["Star"].any():
        miss = (
            f[f["Star"]]
            .groupby("Employee")
            .agg(
                Пропусков=("Star", "sum"),
                Часы=("Hours", "sum"),
            )
            .reset_index()
            .sort_values("Пропусков", ascending=False)
        )
        miss["Дней_работы"] = miss["Employee"].map(
            f.groupby("Employee")["Worked"].sum()
        )
        miss["Часы"] = miss["Часы"].round(1)
        miss = miss[["Employee", "Пропусков", "Дней_работы", "Часы"]].rename(
            columns={"Employee": "Сотрудник", "Пропусков": "Пропусков (*)",
                     "Дней_работы": "Дней работы", "Часы": "Часы, ч"}
        )
        st.dataframe(miss, width="stretch", hide_index=True)
        st.markdown(
            "<div class='cap'>Отметка «*» в таблице = выходной/пропуск "
            "(не учитывается в часах и выплатах).</div>",
            unsafe_allow_html=True,
        )

        miss_top = miss.set_index("Сотрудник")["Пропусков (*)"].head(8)
        pairs = [(emp, int(v), "") for emp, v in miss_top.items()]
        st.markdown(leaderboard_html(pairs, "#EF4444", unit="пропусков"),
                    unsafe_allow_html=True)
    else:
        st.success("За выбранный период пропусков (отметок «*») нет!")

# ── Подвал ────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "Легенда: `*` = выходной/пропуск и не учитывается · дубли дат удаляются · "
    "строки без имени → «Без имени». Дашборд автоматически обновляется "
    "при изменении Google-таблицы."
)
cols = st.columns([2, 2, 3])
with cols[0]:
    st.download_button(
        "⬇ Экспорт CSV (выборка)",
        f.to_csv(index=False).encode("utf-8-sig"),
        "mila_clean.csv",
        "text/csv",
    )
with cols[1]:
    st.download_button(
        "⬇ Экспорт CSV (все данные)",
        df.to_csv(index=False).encode("utf-8-sig"),
        "mila_full.csv",
        "text/csv",
    )
with cols[2]:
    st.caption(
        f"Данных в базе: {len(df):,} строк · {df['Employee'].nunique()} человек · "
        f"листов: {df['Sheet'].nunique()}".replace(",", " ")
    )