import hashlib
import io
import os
import re
import sys
from datetime import date, datetime, timedelta

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
    ("⌛", "Всего часов", "#2563EB"),
    ("💰", "Сумма выплат", "#059669"),
    ("🧑‍🤝‍🧑", "Активных сотрудников", "#7C3AED"),
    ("🎯", "Средняя ставка", "#D97706"),
    ("📈", "Средняя выработка", "#DB2777"),
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
        background: #FFFFFF; border: 1px solid #E7EAF5; border-radius: 18px;
        padding: 26px 30px 18px 30px; position: relative; overflow: hidden;
        box-shadow: 0 8px 22px rgba(31,41,85,.08); margin-bottom: 18px;
    }
    .hero::before {
        content: ""; position: absolute; inset: 0 0 auto 0; height: 5px;
        background: linear-gradient(90deg, #4F46E5, #7C3AED, #2563EB);
    }
    .hero h1 { color: #1E2A5A; margin: 8px 0 4px 0; font-size: 28px;
               letter-spacing: .2px; }
    .hero p  { color: #4B5563; margin: 0; font-size: 14px; }
    .hero .date-chip {
        display: inline-block; background: #EEF2FF;
        border: 1px solid #C7D2FE; color: #3730A3; padding: 3px 12px;
        border-radius: 20px; font-size: 13px; font-weight: 600; margin-top: 10px;
    }
    .hero .live-chip { background: #D1FAE5; border-color: #6EE7B7; color: #047857; }

    .filter-line {
        margin-top: 10px; padding: 8px 14px; background: #EEF2FF;
        border: 1px solid #C7D2FE; border-radius: 10px; color: #3730A3;
        font-size: 13px;
    }
    .filter-line b { color: #1E2A5A; }

    .kpi {
        background: #FFFFFF; border-radius: 16px; padding: 14px 18px;
        border: 1px solid #E7EAF5; height: 100%;
        box-shadow: 0 3px 10px rgba(31,41,85,.06); position: relative;
    }
    .kpi .ico { font-size: 22px; line-height: 1; }
    .kpi .lbl { font-size: 12px; color: #374151; letter-spacing: .4px;
                text-transform: uppercase; margin-top: 6px; white-space: nowrap; }
    .kpi .val { font-size: 25px; font-weight: 800; margin-top: 2px; }
    .kpi .sub { font-size: 12px; color: #4B5563; margin-top: 4px; }
    .kpi .bar  { height: 4px; border-radius: 4px; margin-top: 10px; }

    .lboard-item {
        display: flex; align-items: center; gap: 12px; padding: 9px 12px;
        background: #FFFFFF; border: 1px solid #E7EAF5; border-radius: 12px;
        margin-bottom: 8px; box-shadow: 0 2px 6px rgba(31,41,85,.05);
    }
    .lboard-medal { font-size: 20px; width: 32px; text-align: center; }
    .lboard-name { flex: 1; font-weight: 600; color: #1F2937; }
    .lboard-hours { font-size: 15px; font-weight: 700; color: #1F2937; }
    .lboard-sum { font-size: 12px; color: #4B5563; min-width: 76px; text-align: right; }
    .lboard-track { position: relative; height: 6px; border-radius: 6px;
                    background: #EEF0F8; overflow: hidden; margin-top: 6px; }

    .cap { color: #4B5563; font-size: 13px; margin: 4px 0 12px 0; }
    h1, h2, h3 { color: #1E2A5A; }
    div[data-testid="stTabs"] button p { font-weight: 600; font-size: 15px; }
    .dossier-note {
        background: #FFFFFF; border: 1px solid #E7EAF5; border-radius: 12px;
        padding: 10px 16px; color: #1F2937; font-size: 14px;
        box-shadow: 0 2px 8px rgba(31,41,85,.05);
    }
    .dossier-note b { color: #7C3AED; }
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
# 3b. ДОСЬЕ СОТРУДНИКА (общая функция для вкладок)
# ─────────────────────────────────────────────────────────────
def render_dossier(pick: str):
    person_df = f[f["Employee"] == pick]
    days_n = person_df["Date"].nunique()
    total_days_in_period = (p1 - p0).days + 1
    h_sum = person_df["Hours"].sum()
    a_sum = person_df["Amount"].sum()
    rate = person_df["Rate"].iloc[0]
    avg_day = h_sum / days_n if days_n else 0
    stars = int(person_df["Star"].sum())
    share_h = 100 * h_sum / total_h if total_h else 0
    share_a = 100 * a_sum / total_a if total_a else 0
    weeks_worked = person_df["WeekStart"].nunique()

    d1, d2, d3, d4, d5, d6 = st.columns(6)
    stats = [
        ("⌛", "Часов", fmt_hours(h_sum), "#2563EB"),
        ("💰", "Выплат", fmt_rub(a_sum), "#059669"),
        ("📅", "Дней работы", str(days_n), "#7C3AED"),
        ("🎯", "Ср. ч/день", fmt_hours(avg_day), "#D97706"),
        ("🗓", "Недель в работе", str(weeks_worked), "#0891B2"),
        ("🚫", "Пропусков (*)", str(stars), "#DC2626"),
    ]
    for col, (icon, lbl, val, color) in zip([d1, d2, d3, d4, d5, d6], stats):
        with col:
            st.markdown(kpi_card(icon, lbl, val, "", color), unsafe_allow_html=True)

    st.markdown(
        f"<div class='dossier-note'>Вклад в команду за период: "
        f"<b>{share_h:.1f}%</b> всех часов и <b>{share_a:.1f}%</b> всех выплат. "
        f"Работал(-а) <b>{days_n}</b> из {total_days_in_period} дней периода.</div>",
        unsafe_allow_html=True,
    )

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


# ─────────────────────────────────────────────────────────────
# 3c. КНОПКА ОБНОВЛЕНИЯ
# ─────────────────────────────────────────────────────────────
def _refresh_click():
    st.session_state["mila_src_state"] = None


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

# ── Источник данных (сворачиваемый блок на странице) ─────────
src = st.session_state.get("mila_src", DEFAULT_SRC)

with st.expander("⚙️ Источник данных", expanded=False):
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
    ref_c, ref_b = st.columns([3, 1])
    ref_c.caption(
        f"⚡ Автопроверка каждые {WATCH_INTERVAL} с · "
        f"Обновлено: {fetched:%d.%m.%Y %H:%M:%S}"
    )
    ref_b.button("🔄 Обновить сейчас", use_container_width=True,
                 on_click=_refresh_click)

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

# ── Фильтры на самой странице ─────────────────────────────────
with st.expander("🎛 Фильтры: период · сотрудники · дни", expanded=True):
    fc1, fc2, fc3 = st.columns([1, 1, 1])

    with fc1:
        preset = st.selectbox(
            "Быстрый период",
            ["Все данные", "Последние 7 дней", "Последние 30 дней",
             "Последние 90 дней", "Последний год", "Произвольный"],
            key="mila_preset",
        )
        today = date.today()
        days_ago = {
            "Все данные": None,
            "Последние 7 дней": 7,
            "Последние 30 дней": 30,
            "Последние 90 дней": 90,
            "Последний год": 365,
            "Произвольный": None,
        }[preset]
        if days_ago is not None:
            p0 = max(dmin, today - timedelta(days=days_ago - 1))
            p1 = min(dmax, today)
            period = None
        else:
            period = st.date_input(
                "Период (произвольный)", [dmin, dmax],
                min_value=dmin, max_value=dmax, key="mila_custom_period",
            )
            if period and len(period) >= 2:
                p0, p1 = period[0], period[1]
            else:
                p0, p1 = dmin, dmax

    with fc2:
        all_names = sorted(df["Employee"].unique())
        quick = st.selectbox(
            "Быстро выбрать сотрудника",
            ["— Все сотрудники —"] + all_names,
            key="mila_quick",
        )
        employees = st.multiselect(
            "Или вручную",
            all_names,
            default=all_names,
            disabled=(quick != "— Все сотрудники —"),
            key="mila_employees",
        )
        if quick != "— Все сотрудники —":
            employees = [quick]

    with fc3:
        st.caption("📆 Дни недели")
        day_map = {"Пн": 0, "Вт": 1, "Ср": 2, "Чт": 3, "Пт": 4, "Сб": 5, "Вс": 6}
        days = st.multiselect("Дни недели", list(day_map), default=list(day_map))

    st.markdown(
        f"<div class='filter-line'>Применено: период "
        f"<b>{p0:%d.%m.%Y} — {p1:%d.%m.%Y}</b> · сотрудников "
        f"<b>{len(employees)}</b> · дни <b>{', '.join(days)}</b></div>",
        unsafe_allow_html=True,
    )

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
        <span class="date-chip live-chip">🔴 LIVE · автообнов. каждые {WATCH_INTERVAL} с</span>
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

# ── Вкладки (динамические: при выборе сотрудника добавляется вкладка с его досье) ──
single = None
if quick != "— Все сотрудники —":
    single = quick
elif len(employees) == 1:
    single = employees[0]

tab_labels = ["📊 Обзор"]
if single is not None:
    tab_labels.append(f"👤 {single}")
tab_labels += ["🧑💻 Сводка по людям", "📅 По календарю", "🚫 Пропуски (*)"]
tabs = st.tabs(tab_labels)

person_i = 1 if single is not None else -1
summary_i = 1 if single is None else 2
cal_i = summary_i + 1
skip_i = cal_i + 1

# ════════════════ ОБЗОР ════════════════
with tabs[0]:
    if single is not None:
        st.markdown(
            f"<div class='filter-line'>Показаны данные только сотрудника "
            f"<b>{single}</b> — подробно о нём во вкладке «👤 {single}».</div>",
            unsafe_allow_html=True,
        )
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

# ════════════════ СВОДКА ПО ЛЮДЯМ ════════════════
with tabs[summary_i]:
    st.subheader("🧑💻 Сводка по сотрудникам")
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
    person["Доля_часов"] = 100 * person["Часы"] / person["Часы"].sum()
    person["Доля_выплат"] = 100 * person["Сумма"] / person["Сумма"].sum()
    person["Часы"] = person["Часы"].round(1)
    person["Сумма"] = person["Сумма"].round(0)
    person["Ставка"] = person["Ставка"].round(0)
    person["Ср_в_день"] = person["Ср_в_день"].round(1)
    person["Доля_часов"] = person["Доля_часов"].round(1)
    person["Доля_выплат"] = person["Доля_выплат"].round(1)

    styled = person.rename(columns={
        "Ср_в_день": "Ср.ч/день", "Доля_часов": "Часы, %",
        "Доля_выплат": "Выплаты, %",
    })
    st.dataframe(
        styled[["Employee", "Дней", "Часы", "Сумма", "Ставка", "Ср.ч/день",
                "Часы, %", "Выплаты, %"]]
        .rename(columns={"Employee": "Сотрудник", "Дней": "Дней работы",
                         "Часы": "Часы, ч", "Сумма": "Сумма, ₽",
                         "Ставка": "Ставка, ₽/ч"}),
        width="stretch", hide_index=True,
        column_config={
            "Часы, %": st.column_config.ProgressColumn(
                "Доля часов, %", min_value=0, max_value=100, format="%.1f%%"
            ),
            "Выплаты, %": st.column_config.ProgressColumn(
                "Доля выплат, %", min_value=0, max_value=100, format="%.1f%%"
            ),
        },
    )
    st.markdown(
        "<div class='cap'>«Доля часов/выплат» — сколько процентов от всей "
        "команды приходится на человека за период. «Ср.ч/день» — его обычный "
        "объём работы в день.</div>",
        unsafe_allow_html=True,
    )

    st.markdown("---")
    if single is None:
        st.subheader("🔍 Досье сотрудника")
        pick = st.selectbox(
            "Выберите сотрудника",
            sorted(f["Employee"].unique()),
            key="mila_pick",
        )
        render_dossier(pick)

if single is not None:
    # ════════════════ ДОСЬЕ ВЫБРАННОГО ════════════════
    with tabs[person_i]:
        st.subheader(f"👤 {single}")
        render_dossier(single)

# ════════════════ ПО КАЛЕНДАРЮ ════════════════
with tabs[cal_i]:
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
with tabs[skip_i]:
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