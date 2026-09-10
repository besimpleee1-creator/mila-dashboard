import hashlib
import io
import os
import re
import sys
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
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


def log(msg: str) -> None:
    print(msg, flush=True)

DAY_NAME_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
PALETTE = ["#E8F0FE", "#7BA7D9", "#1F3A5F"]


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
        urls.append(
            f"https://docs.google.com/spreadsheets/export"
            f"?format=xlsx&id={sheet_id}"
        )

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
    """Разбирает все недельные блоки во всех листах книги."""
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


def snapshot(src: str) -> tuple:
    """Скачивает книгу и возвращает (payload, sha1-хэш)."""
    data = fetch_bytes(src)
    return data, hashlib.sha1(data).hexdigest()


@st.cache_data(show_spinner=False)
def build_frame(token: str, payload: bytes) -> pd.DataFrame:
    """Парсит скачанную книгу. token (хэш содержимого) — ключ кэша."""
    return parse_workbook(payload)


# ─────────────────────────────────────────────────────────────
# 2. АВТО-ОБНОВЛЕНИЕ ПРИ ИЗМЕНЕНИИ ТАБЛИЦЫ
# ─────────────────────────────────────────────────────────────
@st.fragment(run_every=WATCH_INTERVAL)
def watch_source(src: str):
    """Каждые N секунд скачивает книгу; при изменении содержимого — перерисовка всего приложения."""
    try:
        payload = fetch_bytes(src)
    except Exception:
        return
    h = hashlib.sha1(payload).hexdigest()
    if st.session_state.get("mila_hash") != h:
        st.session_state["mila_payload"] = payload
        st.session_state["mila_hash"] = h
        st.session_state["mila_fetched_at"] = datetime.now()
        st.rerun(scope="app")


# ─────────────────────────────────────────────────────────────
# 3. СТРАНИЦА
# ─────────────────────────────────────────────────────────────
log(f"[run] start script, src={DEFAULT_SRC[:85]}")
st.set_page_config(page_title="Мила — Статистика", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .stApp { background: #F7F9FC; }
    .kpi {
        background: #FFFFFF; border-radius: 14px; padding: 16px 20px;
        border: 1px solid #E6ECF4;
        box-shadow: 0 2px 8px rgba(31,58,95,0.08);
        border-left: 4px solid #1F3A5F;
    }
    .kpi-label { font-size: 12px; color: #6B7280; letter-spacing: .6px;
                 text-transform: uppercase; }
    .kpi-value { font-size: 24px; font-weight: 700; color: #1F3A5F; }
    h1, h2, h3 { color: #1F3A5F; }
    .badge { display: inline-block; padding: 2px 10px; border-radius: 20px;
             font-size: 12px; font-weight: 600; }
    .badge-live { background: #E7F5EC; color: #18794E; }
    .badge-warn { background: #FFF1E0; color: #B5531B; }
</style>
""", unsafe_allow_html=True)

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

    # Скачиваем книгу при старте или при смене адреса
    if st.session_state.get("mila_src_state") != src:
        try:
            payload, h = snapshot(src)
            st.session_state["mila_payload"] = payload
            st.session_state["mila_hash"] = h
            st.session_state["mila_src_state"] = src
            st.session_state["mila_fetched_at"] = datetime.now()
        except Exception:
            st.session_state.pop("mila_payload", None)
            st.session_state.pop("mila_hash", None)
            st.session_state["mila_src_state"] = src
            st.error("Не удалось загрузить книгу по этому адресу.")
            st.caption("Проверьте, что файл открыт «по ссылке» для просмотра, "
                       "и что адрес начинается с http://")
            st.stop()

    watch_source(src)

    fetched = st.session_state.get("mila_fetched_at")
    st.caption(
        f"Источник: **Google Таблица / файл** · "
        f"проверка изменений каждые {WATCH_INTERVAL} с\n\n"
        f"Последнее обновление: {fetched:%d.%m.%Y %H:%M:%S}"
    )
    st.caption(
        f"<span class='badge badge-live'>🔴 LIVE</span> "
        f"хэш данных: `{st.session_state['mila_hash'][:10]}…`",
        unsafe_allow_html=True,
    )
    if st.button("🔄 Обновить сейчас", use_container_width=True):
        build_frame.clear()
        st.session_state["mila_src_state"] = None
        st.rerun()

    st.markdown("---")
    st.header("🔎 Фильтры")

# ── Данные ────────────────────────────────────────────────────
try:
    df = build_frame(st.session_state["mila_hash"], st.session_state["mila_payload"])
except Exception as exc:
    st.error("Не удалось прочитать содержимое книги.")
    st.caption(
        "Возможно, Google запросил вход по ссылке. Убедитесь, что доступ "
        "к таблице открыт «для всех, у кого есть ссылка», и обновите страницу. "
        f"\n\nТехнически: {type(exc).__name__}: {exc}"
    )
    st.stop()

if df.empty:
    st.warning("Не удалось найти данные в книге — проверьте разметку листов.")
    st.stop()

log(f"[run] df shape={df.shape} employees={df['Employee'].nunique()} "
    f"dates={df['Date'].min().date()}..{df['Date'].max().date()}")

dmin, dmax = df["Date"].min().date(), df["Date"].max().date()

with st.sidebar:
    period = st.date_input("Период", [dmin, dmax], min_value=dmin, max_value=dmax)
    employees = st.multiselect(
        "Сотрудники",
        sorted(df["Employee"].unique()),
        default=sorted(df["Employee"].unique()),
    )
    day_map = {"Пн": 0, "Вт": 1, "Ср": 2, "Чт": 3, "Пт": 4, "Сб": 5, "Вс": 6}
    days = st.multiselect("Дни недели", list(day_map), default=list(day_map))
    if st.button("Сбросить фильтры", use_container_width=True):
        st.rerun()

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

st.title("📊 Мила — Статистика часов и выплат")
st.caption(
    f"Период: **{p0:%d.%m.%Y} — {p1:%d.%m.%Y}** · "
    f"листов в книге: **{df['Sheet'].nunique()}** · "
    f"строк с данными: **{len(df):,}**".replace(",", " ")
)

# ─────────────────────────────────────────────────────────────
# 4. KPI
# ─────────────────────────────────────────────────────────────
total_h = f["Hours"].sum()
total_a = f["Amount"].sum()
n_emp   = f["Employee"].nunique()
avg_rate = f.groupby("Employee")["Rate"].first().mean() if n_emp else 0
avg_out  = total_h / n_emp if n_emp else 0

c1, c2, c3, c4, c5 = st.columns(5)
for col, label, val in zip(
    [c1, c2, c3, c4, c5],
    ["Всего часов", "Сумма выплат", "Активных", "Ср. ставка", "Ср. выработка"],
    [
        f"{total_h:,.1f}",
        f"{total_a:,.0f} ₽",
        f"{n_emp}",
        f"{avg_rate:,.0f} ₽",
        f"{avg_out:,.1f} ч",
    ],
):
    with col:
        st.markdown(
            f"""<div class="kpi">
                <div class="kpi-label">{label}</div>
                <div class="kpi-value">{val}</div>
            </div>""",
            unsafe_allow_html=True,
        )

st.markdown("---")

if f.empty:
    st.info("По выбранным фильтрам данных нет — измените период, сотрудников или дни.")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 5. МАТРИЦА «СОТРУДНИКИ × ДНИ» + ИТОГИ
# ─────────────────────────────────────────────────────────────
st.subheader("🔥 Матрица «Сотрудники × Дни» (часы)")

left, right = st.columns([3, 2])

with left:
    pivot = f.pivot_table(
        index="Employee", columns="Date", values="Hours",
        aggfunc="sum", fill_value=np.nan,
    )
    pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]

    fig = px.imshow(
        pivot,
        aspect="auto",
        color_continuous_scale=PALETTE,
        labels=dict(x="Дата", y="Сотрудник", color="Часы"),
    )
    fig.update_layout(
        height=max(300, 34 * len(pivot)),
        paper_bgcolor="#F7F9FC",
        plot_bgcolor="#FFFFFF",
        font=dict(size=12),
        margin=dict(l=10, r=10, t=10, b=10),
    )
    fig.update_xaxes(tickformat="%d.%m")
    st.plotly_chart(fig, width="stretch")

with right:
    summary = f.groupby("Employee").agg(
        Итог_часов=("Hours", "sum"),
        Итог_сумм=("Amount", "sum"),
        Ставка=("Rate", "first"),
    ).sort_values("Итог_часов", ascending=False)
    summary["Итог_сумм"] = summary["Итог_сумм"].map(lambda x: f"{x:,.0f} ₽")
    summary["Ставка"] = summary["Ставка"].map(lambda x: f"{x:,.0f} ₽")
    summary = summary.rename(columns={
        "Итог_часов": "Часы",
        "Итог_сумм": "Сумма",
        "Ставка": "Ставка/ч",
    })
    st.dataframe(summary, width="stretch")
    st.caption("Сортировка по количеству часов (по убыванию).")

st.markdown("---")

# ─────────────────────────────────────────────────────────────
# 6. ГРАФИКИ
# ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("📊 Часы по сотрудникам")
    bar = f.groupby("Employee")["Hours"].sum().sort_values()
    fig = px.bar(
        x=bar.values, y=bar.index, orientation="h",
        color=bar.values,
        color_continuous_scale=PALETTE[1:],
    )
    fig.update_traces(marker_line_width=0, showlegend=False)
    fig.update_layout(
        coloraxis_showscale=False,
        paper_bgcolor="#F7F9FC", plot_bgcolor="#FFFFFF",
        xaxis_title="Часы", yaxis_title="",
        margin=dict(l=0, r=10, t=10, b=0),
    )
    st.plotly_chart(fig, width="stretch")

with col2:
    st.subheader("📈 Динамика по неделям")
    weekly = f.groupby("WeekStart")["Hours"].sum().reset_index()
    fig = px.line(weekly, x="WeekStart", y="Hours", markers=True,
                  color_discrete_sequence=["#1F3A5F"])
    fig.update_traces(fill="tozeroy", fillcolor="rgba(31,58,95,0.10)")
    fig.update_layout(
        paper_bgcolor="#F7F9FC", plot_bgcolor="#FFFFFF",
        xaxis_title="Неделя", yaxis_title="Часы",
        margin=dict(l=0, r=10, t=10, b=0),
    )
    fig.update_xaxes(tickformat="%d.%m")
    st.plotly_chart(fig, width="stretch")

col3, col4 = st.columns(2)

with col3:
    st.subheader("🍩 Доля выплат")
    donut = f.groupby("Employee")["Amount"].sum().sort_values()
    fig = px.pie(values=donut.values, names=donut.index, hole=0.5,
                 color_discrete_sequence=px.colors.sequential.Blues_r)
    fig.update_layout(
        paper_bgcolor="#F7F9FC",
        legend=dict(orientation="h", yanchor="bottom", y=-0.25),
        margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig, width="stretch")

with col4:
    st.subheader("🌡 Нагрузка по дням недели")
    heat = f.pivot_table(index="DayOfWeek", columns="WeekStart",
                         values="Hours", aggfunc="sum", fill_value=0)
    heat.index = [DAY_NAME_RU[i] for i in heat.index.get_level_values(0)]
    fig = px.imshow(heat, aspect="auto", color_continuous_scale=PALETTE,
                    labels=dict(x="Неделя", y="День", color="Часы"))
    fig.update_layout(
        paper_bgcolor="#F7F9FC", plot_bgcolor="#FFFFFF",
        margin=dict(l=10, r=10, t=10, b=10),
    )
    fig.update_xaxes(tickformat="%d.%m")
    st.plotly_chart(fig, width="stretch")

st.markdown("---")

# ─────────────────────────────────────────────────────────────
# 7. ДЕТАЛИ (drill-down)
# ─────────────────────────────────────────────────────────────
st.subheader("📋 Детали по сотруднику")
pick = st.selectbox("Выберите сотрудника", sorted(f["Employee"].unique()))
detail = (
    f[f["Employee"] == pick][["Date", "DayOfWeek", "Hours", "Rate", "Amount"]]
    .sort_values("Date")
    .assign(День=lambda d: d["DayOfWeek"].map(lambda x: DAY_NAME_RU[x]))
    .drop(columns="DayOfWeek")
    .rename(columns={"Date": "Дата", "Hours": "Часы", "Rate": "Ставка", "Amount": "Сумма"})
)
detail["Дата"] = detail["Дата"].dt.strftime("%d.%m.%Y")
detail["Сумма"] = detail["Сумма"].map(lambda x: f"{x:,.0f} ₽")
detail["Ставка"] = detail["Ставка"].map(lambda x: f"{x:,.0f} ₽")
detail["Часы"] = detail["Часы"].map(lambda x: f"{x:.1f}")
st.dataframe(detail, width="stretch", hide_index=True)

# ─────────────────────────────────────────────────────────────
# 8. ПОДВАЛ
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "Легенда: `*` = выходной/пропуск и не учитывается · дубли дат удаляются · "
    "строки без имени → «Без имени». Дашборд автоматически обновляется "
    "при изменении Google-таблицы."
)
st.download_button(
    "⬇ Экспорт в CSV",
    f.to_csv(index=False).encode("utf-8-sig"),
    "mila_clean.csv",
    "text/csv",
)