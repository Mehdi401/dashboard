from __future__ import annotations

import io
import re
import warnings
from datetime import date
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# Universal Executive Dashboard | Management Reporting Edition
# ============================================================

st.set_page_config(
    page_title="داشبورد گزارش‌دهی مدیریتی",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MAX_UPLOAD_MB = 200
MAX_CORRELATION_COLUMNS = 30
EXCEL_ROW_LIMIT = 1_000_000

st.markdown(
    """
    <style>
    @import url('https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css');
    html, body, [class*="css"], [class*="st-"] {font-family: Vazirmatn, Tahoma, sans-serif;}
    .stApp {direction: rtl;}
    h1,h2,h3,h4,p,label,.stMarkdown {text-align:right;}
    [data-testid="stDataFrame"], .stPlotlyChart {direction:ltr;}
    [data-testid="stMetric"] {
        background: linear-gradient(135deg,#ffffff,#f3f6fb);
        border:1px solid #dbe3ef; border-right:5px solid #2563eb;
        border-radius:12px; padding:14px 16px;
    }
    [data-testid="stMetricLabel"] {font-weight:600;color:#475569;}
    [data-testid="stMetricValue"] {font-weight:800;color:#0f172a;}
    .hero {
        background:linear-gradient(120deg,#102a56,#2563eb);
        color:white;padding:24px 28px;border-radius:16px;margin-bottom:18px;
    }
    .hero h1,.hero p {color:white;margin:0;text-align:right;}
    .hero p {margin-top:8px;opacity:.92;}
    </style>
    """,
    unsafe_allow_html=True,
)


def fmt(value: Any, digits: int = 2) -> str:
    """Format numeric values for display without failing on missing values."""
    if value is None or pd.isna(value):
        return "—"
    try:
        value = float(value)
    except (TypeError, ValueError):
        return str(value)
    if abs(value) >= 1_000_000_000:
        return f"{value / 1_000_000_000:,.2f} میلیارد"
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:,.2f} میلیون"
    if value.is_integer():
        return f"{int(value):,}"
    return f"{value:,.{digits}f}"


def load_csv(file_bytes: bytes) -> pd.DataFrame:
    errors = []
    for encoding in ("utf-8-sig", "utf-8", "cp1256", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding, sep=None, engine="python")
        except (UnicodeDecodeError, pd.errors.ParserError) as exc:
            errors.append(f"{encoding}: {exc}")
    raise ValueError("خواندن CSV ناموفق بود. جداکننده یا Encoding فایل را بررسی کنید.")


@st.cache_data(show_spinner="در حال خواندن فایل...")
def read_excel_bytes(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)


@st.cache_data(show_spinner=False)
def excel_sheets(file_bytes: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


def normalize_dataframe(raw: pd.DataFrame, remove_duplicates: bool) -> tuple[pd.DataFrame, dict]:
    df = raw.copy()
    original_rows, original_cols = df.shape
    # Normalize column labels and make duplicate labels unique.
    labels = []
    seen = {}
    for i, column in enumerate(df.columns):
        name = str(column).strip() or f"ستون_بدون_نام_{i+1}"
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1:
            name = f"{name}_{seen[name]}"
        labels.append(name)
    df.columns = labels

    converted_numeric, converted_dates = [], []
    for col in df.columns:
        if df[col].dtype != "object":
            continue
        series = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)
        series = series.replace(r"^\s*$", np.nan, regex=True)
        non_null = series.dropna()
        if non_null.empty:
            df[col] = series
            continue

        # Conservative numeric conversion: only convert when nearly all values parse.
        cleaned = non_null.astype(str).str.replace(",", "", regex=False)
        numeric = pd.to_numeric(cleaned, errors="coerce")
        if numeric.notna().mean() >= 0.95:
            df[col] = pd.to_numeric(series.astype(str).str.replace(",", "", regex=False), errors="coerce")
            converted_numeric.append(col)
            continue

        # Conservative date conversion, avoiding forced conversion of ordinary text.
        sample = non_null.astype(str).head(200)
        looks_like_date = sample.str.match(
            r"^\s*\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(\s|$)"
        ).mean() >= 0.90
        if looks_like_date:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                parsed = pd.to_datetime(series, errors="coerce")
            if parsed.notna().sum() >= 0.90 * len(non_null):
                df[col] = parsed
                converted_dates.append(col)
            else:
                df[col] = series
        else:
            df[col] = series

    duplicates_found = int(df.duplicated().sum())
    removed = 0
    if remove_duplicates:
        before = len(df)
        df = df.drop_duplicates().reset_index(drop=True)
        removed = before - len(df)

    report = {
        "original_rows": original_rows,
        "original_cols": original_cols,
        "duplicates_found": duplicates_found,
        "duplicates_removed": removed,
        "numeric_converted": converted_numeric,
        "dates_converted": converted_dates,
    }
    return df, report


def classify_columns(df: pd.DataFrame) -> tuple[list[str], list[str], list[str]]:
    numeric = df.select_dtypes(include=np.number).columns.tolist()
    dates = df.select_dtypes(include=["datetime", "datetimetz"]).columns.tolist()
    categorical = [c for c in df.columns if c not in numeric and c not in dates]
    return numeric, categorical, dates


def quality_report(df: pd.DataFrame) -> pd.DataFrame:
    n = max(len(df), 1)
    return pd.DataFrame({
        "ستون": df.columns,
        "نوع داده": [str(t) for t in df.dtypes],
        "تعداد مقادیر گمشده": df.isna().sum().values,
        "درصد مقادیر گمشده": (df.isna().mean().values * 100).round(2),
        "تعداد مقادیر یکتا": df.nunique(dropna=True).values,
        "درصد یکتایی نسبت به ردیف‌ها": (df.nunique(dropna=True).values / n * 100).round(2),
    })


def quality_score(df: pd.DataFrame) -> dict:
    if df.empty or df.size == 0:
        return {"score": 0.0, "completeness": 0.0, "uniqueness": 0.0, "constant_columns": 0}
    completeness = float(1 - df.isna().sum().sum() / df.size)
    uniqueness = float(1 - df.duplicated().sum() / max(len(df), 1))
    constant_columns = int((df.nunique(dropna=True) <= 1).sum())
    # This is an indicative score, not a certification of business-data accuracy.
    score = 100 * (0.65 * completeness + 0.25 * uniqueness + 0.10 * (1 - constant_columns / max(df.shape[1], 1)))
    return {
        "score": round(score, 1),
        "completeness": round(completeness * 100, 1),
        "uniqueness": round(uniqueness * 100, 1),
        "constant_columns": constant_columns,
    }


def outlier_summary(df: pd.DataFrame, numeric_cols: list[str]) -> pd.DataFrame:
    rows = []
    for col in numeric_cols:
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        if s.empty:
            continue
        q1, q3 = s.quantile([0.25, 0.75])
        iqr = q3 - q1
        low, high = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        count = int(((s < low) | (s > high)).sum())
        rows.append({
            "متغیر": col,
            "تعداد داده معتبر": len(s),
            "حد پایین IQR": low,
            "حد بالا IQR": high,
            "تعداد داده پرت": count,
            "درصد داده پرت": round(100 * count / max(len(s), 1), 2),
        })
    return pd.DataFrame(rows).sort_values("تعداد داده پرت", ascending=False) if rows else pd.DataFrame()


def safe_filename(name: str) -> str:
    name = re.sub(r"[^\w\-.]+", "_", name, flags=re.UNICODE).strip("._")
    return name[:80] or "management_report"


def style_fig(fig, height: int = 420):
    fig.update_layout(
        template="plotly_white", height=height,
        margin=dict(l=25, r=25, t=65, b=35),
        font=dict(family="Vazirmatn, Tahoma, sans-serif", size=12),
        title=dict(x=0.5, xanchor="center"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
    )
    return fig


def make_excel_report(data: pd.DataFrame, source: str, prep: dict, kpi: str | None) -> bytes:
    buffer = io.BytesIO()
    numeric, _, _ = classify_columns(data)
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        data.head(EXCEL_ROW_LIMIT).to_excel(writer, sheet_name="Filtered_Data", index=False)
        quality_report(data).to_excel(writer, sheet_name="Data_Quality", index=False)
        if numeric:
            data[numeric].describe().T.round(4).to_excel(writer, sheet_name="Descriptive_Stats")
            outlier_summary(data, numeric).to_excel(writer, sheet_name="Outliers", index=False)
        summary = pd.DataFrame([
            {"شاخص": "منبع داده", "مقدار": source},
            {"شاخص": "تاریخ تهیه گزارش", "مقدار": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")},
            {"شاخص": "تعداد ردیف پس از فیلتر", "مقدار": len(data)},
            {"شاخص": "تعداد ستون", "مقدار": data.shape[1]},
            {"شاخص": "درصد کامل بودن", "مقدار": round(100 * (1 - data.isna().sum().sum() / max(data.size, 1)), 2)},
            {"شاخص": "تعداد تکراری حذف‌شده در آماده‌سازی", "مقدار": prep.get("duplicates_removed", 0)},
            {"شاخص": "شاخص انتخاب‌شده", "مقدار": kpi or "انتخاب نشده"},
        ])
        summary.to_excel(writer, sheet_name="Executive_Summary", index=False)
    return buffer.getvalue()


# -------------------------- Sidebar --------------------------
st.sidebar.title("📂 تنظیمات گزارش")
uploaded = st.sidebar.file_uploader("بارگذاری فایل Excel یا CSV", type=["xlsx", "xls", "csv"])
use_sample = st.sidebar.toggle("استفاده از داده نمونه نمایشی", value=False)
remove_duplicates = st.sidebar.checkbox("حذف ردیف‌های تکراری", value=False)
st.sidebar.caption(f"حداکثر حجم فایل: {MAX_UPLOAD_MB} مگابایت")

raw = None
source = ""
prep = {}
if uploaded is not None:
    if uploaded.size > MAX_UPLOAD_MB * 1024 * 1024:
        st.error(f"حجم فایل از سقف {MAX_UPLOAD_MB} مگابایت بیشتر است.")
        st.stop()
    try:
        content = uploaded.getvalue()
        source = uploaded.name
        if uploaded.name.lower().endswith(".csv"):
            raw = load_csv(content)
        else:
            sheets = excel_sheets(content)
            sheet = st.sidebar.selectbox("انتخاب شیت Excel", sheets)
            raw = read_excel_bytes(content, sheet)
    except Exception as exc:
        st.error(f"فایل خوانده نشد: {exc}")
        st.stop()
elif use_sample:
    rng = np.random.default_rng(21)
    n = 900
    dates = pd.date_range("2025-01-01", periods=n, freq="D")
    region = rng.choice(["شمال", "جنوب", "مرکز", "غرب"], n)
    quantity = rng.integers(1, 25, n)
    unit_price = rng.normal(250_000, 45_000, n).clip(30_000, None).round(0)
    raw = pd.DataFrame({
        "تاریخ": dates, "منطقه": region, "تعداد": quantity,
        "قیمت واحد": unit_price, "فروش": quantity * unit_price,
        "رضایت مشتری": rng.normal(4.1, .55, n).clip(1, 5).round(1),
    })
    raw.loc[rng.choice(n, 30, replace=False), "رضایت مشتری"] = np.nan
    source = "داده نمونه نمایشی (ساختگی)"
else:
    st.markdown("""
    <div class="hero">
      <h1>📊 داشبورد گزارش‌دهی مدیریتی</h1>
      <p>تحلیل داده، پایش شاخص‌ها، شناسایی ریسک‌ها و آماده‌سازی گزارش جلسه</p>
    </div>
    """, unsafe_allow_html=True)
    st.info("برای شروع، از نوار کناری یک فایل Excel/CSV بارگذاری کنید یا داده نمونه نمایشی را فعال کنید.")
    st.stop()

if raw is None or raw.empty:
    st.error("فایل خالی است یا داده‌ای برای تحلیل وجود ندارد.")
    st.stop()

df, prep = normalize_dataframe(raw, remove_duplicates)
if df.empty or df.shape[1] == 0:
    st.error("پس از آماده‌سازی، داده قابل تحلیلی باقی نمانده است.")
    st.stop()

numeric_cols, categorical_cols, date_cols = classify_columns(df)

st.markdown(f"""
<div class="hero">
  <h1>📊 داشبورد گزارش‌دهی مدیریتی</h1>
  <p>منبع: {source} &nbsp; | &nbsp; تهیه گزارش: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}</p>
</div>
""", unsafe_allow_html=True)

# -------------------------- Global filters --------------------------
with st.sidebar:
    st.divider()
    st.subheader("🔎 فیلترهای سراسری")
    filter_columns = st.multiselect("انتخاب ستون برای فیلتر", df.columns.tolist())
    mask = pd.Series(True, index=df.index)
    for col in filter_columns:
        if col in date_cols:
            valid = df[col].dropna()
            if valid.empty:
                continue
            date_range = st.date_input(
                f"بازه {col}",
                value=(valid.min().date(), valid.max().date()),
                key=f"date_filter_{col}",
            )
            if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
                start, end = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]) + pd.Timedelta(days=1)
                mask &= df[col].between(start, end, inclusive="left")
        elif col in numeric_cols:
            valid = pd.to_numeric(df[col], errors="coerce").dropna()
            if valid.empty or valid.min() == valid.max():
                continue
            low, high = st.slider(
                f"بازه {col}", float(valid.min()), float(valid.max()),
                (float(valid.min()), float(valid.max())), key=f"num_filter_{col}",
            )
            mask &= df[col].between(low, high) | df[col].isna()
        else:
            values = sorted(df[col].dropna().astype(str).unique().tolist())
            if len(values) > 300:
                st.caption(f"فیلتر «{col}» به دلیل بیش از ۳۰۰ مقدار یکتا نمایش داده نشد.")
                continue
            selected = st.multiselect(f"مقادیر {col}", values, default=values, key=f"cat_filter_{col}")
            mask &= df[col].astype(str).isin(selected) | df[col].isna()

filtered = df.loc[mask].copy()
if filtered.empty:
    st.warning("با فیلترهای فعلی هیچ ردیفی باقی نمانده است. فیلترها را تغییر دهید.")
    st.stop()

nnum, ncat, ndates = classify_columns(filtered)
qscore = quality_score(filtered)
missing_total = int(filtered.isna().sum().sum())
duplicate_current = int(filtered.duplicated().sum())

st.caption(f"تعداد رکوردهای اولیه: {len(df):,}  |  رکوردهای پس از فیلتر: {len(filtered):,}  |  ستون‌ها: {filtered.shape[1]:,}")

tabs = st.tabs([
    "خلاصه مدیریتی", "تحلیل نموداری", "مقایسه گروه‌ها",
    "روند زمانی", "کیفیت داده", "خروجی گزارش"
])

# -------------------------- Executive summary --------------------------
with tabs[0]:
    st.subheader("نمای کلی برای جلسه مدیریتی")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("تعداد رکورد", f"{len(filtered):,}")
    k2.metric("تعداد شاخص/ستون", f"{filtered.shape[1]:,}")
    k3.metric("کامل بودن داده", f"{qscore['completeness']}%")
    k4.metric("مقادیر گمشده", f"{missing_total:,}")
    k5.metric("امتیاز کیفیت تخمینی", f"{qscore['score']}/100")

    st.caption("امتیاز کیفیت یک شاخص داخلی و تقریبی است؛ جایگزین اعتبارسنجی تخصصی یا کنترل صحت کسب‌وکار نیست.")
    if nnum:
        st.markdown("#### شاخص کلیدی عملکرد (KPI)")
        selected_kpi = st.selectbox("شاخص عددی مورد بررسی", nnum, key="executive_kpi")
        s = pd.to_numeric(filtered[selected_kpi], errors="coerce").dropna()
        a, b, c, d, e = st.columns(5)
        a.metric("مجموع", fmt(s.sum()) if not s.empty else "—")
        b.metric("میانگین", fmt(s.mean()) if not s.empty else "—")
        c.metric("میانه", fmt(s.median()) if not s.empty else "—")
        d.metric("بیشینه", fmt(s.max()) if not s.empty else "—")
        e.metric("کمینه", fmt(s.min()) if not s.empty else "—")

        st.markdown("#### مقایسه دو دوره (در صورت وجود ستون تاریخ)")
        if ndates:
            date_col = st.selectbox("ستون تاریخ برای مقایسه", ndates, key="period_date")
            valid_dates = filtered[date_col].dropna()
            if len(valid_dates) >= 2:
                max_date = valid_dates.max()
                min_date = valid_dates.min()
                span_days = max((max_date - min_date).days + 1, 2)
                period_days = st.number_input(
                    "طول هر دوره مقایسه (روز)", min_value=1,
                    max_value=min(3650, span_days), value=max(1, min(30, span_days // 2)),
                )
                cut = max_date - pd.Timedelta(days=int(period_days) - 1)
                previous_cut = cut - pd.Timedelta(days=1)
                previous_start = previous_cut - pd.Timedelta(days=int(period_days) - 1)
                current_values = pd.to_numeric(
                    filtered.loc[filtered[date_col].between(cut, max_date), selected_kpi], errors="coerce"
                ).dropna()
                previous_values = pd.to_numeric(
                    filtered.loc[filtered[date_col].between(previous_start, previous_cut), selected_kpi], errors="coerce"
                ).dropna()
                agg = st.selectbox("روش مقایسه", ["مجموع", "میانگین", "میانه", "تعداد"], key="period_agg")
                def aggregate(series):
                    if agg == "مجموع": return series.sum() if len(series) else np.nan
                    if agg == "میانگین": return series.mean() if len(series) else np.nan
                    if agg == "میانه": return series.median() if len(series) else np.nan
                    return series.count()
                current_value, previous_value = aggregate(current_values), aggregate(previous_values)
                delta = np.nan
                if pd.notna(previous_value) and previous_value != 0 and pd.notna(current_value):
                    delta = (current_value - previous_value) / abs(previous_value) * 100
                p1, p2, p3 = st.columns(3)
                p1.metric("دوره اخیر", fmt(current_value), help=f"از {cut.date()} تا {max_date.date()}")
                p2.metric("دوره قبل", fmt(previous_value), help=f"از {previous_start.date()} تا {previous_cut.date()}")
                p3.metric("تغییر نسبت به دوره قبل", f"{delta:+.2f}%" if pd.notna(delta) else "قابل محاسبه نیست")
                if pd.isna(delta):
                    st.info("برای محاسبه تغییر درصدی، هر دو دوره باید داده معتبر داشته باشند و مقدار دوره قبل صفر نباشد.")
        else:
            st.info("برای مقایسه دوره‌ای، فایل باید حداقل یک ستون تاریخ معتبر داشته باشد.")

    st.divider()
    st.markdown("#### نکات قابل طرح در جلسه")
    if missing_total:
        st.warning(f"در داده‌های فیلترشده {missing_total:,} مقدار گمشده وجود دارد؛ پیش از تصمیم‌گیری، اثر آن بر شاخص‌های اصلی بررسی شود.")
    else:
        st.success("در داده‌های فیلترشده مقدار گمشده‌ای شناسایی نشد.")
    if duplicate_current:
        st.warning(f"{duplicate_current:,} ردیف تکراری در داده‌های فعلی دیده می‌شود.")
    if nnum:
        outs = outlier_summary(filtered, nnum)
        if not outs.empty and outs["تعداد داده پرت"].max() > 0:
            worst = outs.iloc[0]
            st.info(f"بیشترین تعداد مشاهده خارج از محدوده IQR در ستون «{worst['متغیر']}» دیده شده است ({int(worst['تعداد داده پرت']):,} مورد). این موارد الزاماً خطا نیستند و باید بررسی شوند.")
    if not nnum:
        st.info("ستون عددی شناسایی نشد؛ KPI عددی و تحلیل‌های کمی در دسترس نیست.")

# -------------------------- Chart explorer --------------------------
with tabs[1]:
    st.subheader("تحلیل نموداری اکتشافی")
    chart_type = st.selectbox(
        "نوع تحلیل",
        ["توزیع عددی", "نمودار جعبه‌ای", "پراکندگی دو متغیر", "فراوانی دسته‌ها", "پارتو دسته‌ها"],
    )
    if chart_type in ("توزیع عددی", "نمودار جعبه‌ای") and nnum:
        col = st.selectbox("متغیر عددی", nnum, key="chart_num")
        valid = pd.to_numeric(filtered[col], errors="coerce").dropna()
        if valid.empty:
            st.info("برای این ستون مقدار عددی معتبر وجود ندارد.")
        elif chart_type == "توزیع عددی":
            bins = st.slider("تعداد بازه‌های هیستوگرام", 5, 100, 30)
            fig = px.histogram(filtered, x=col, nbins=bins, marginal="box", title=f"توزیع {col}")
            show = style_fig(fig, 470)
            st.plotly_chart(show, use_container_width=True)
            st.caption(f"تعداد معتبر: {len(valid):,} | انحراف معیار: {fmt(valid.std())} | چولگی: {fmt(valid.skew())}")
        else:
            group = st.selectbox("گروه‌بندی اختیاری", ["بدون گروه"] + ncat, key="box_group")
            fig = px.box(filtered, y=col, x=None if group == "بدون گروه" else group, points="outliers", title=f"نمودار جعبه‌ای {col}")
            st.plotly_chart(style_fig(fig, 470), use_container_width=True)
    elif chart_type == "پراکندگی دو متغیر" and len(nnum) >= 2:
        xcol, ycol = st.columns(2)
        x = xcol.selectbox("محور افقی (X)", nnum, index=0)
        y_default = 1 if len(nnum) > 1 else 0
        y = ycol.selectbox("محور عمودی (Y)", nnum, index=y_default)
        color_by = st.selectbox("رنگ بر اساس دسته (اختیاری)", ["بدون دسته"] + ncat)
        fig = px.scatter(
            filtered, x=x, y=y,
            color=None if color_by == "بدون دسته" else color_by,
            trendline="ols" if st.checkbox("نمایش خط روند خطی", value=False) else None,
            title=f"رابطه {x} و {y}",
        )
        st.plotly_chart(style_fig(fig, 480), use_container_width=True)
        corr = filtered[[x, y]].corr().iloc[0, 1]
        st.metric("ضریب همبستگی پیرسون", fmt(corr))
        st.caption("همبستگی به‌تنهایی رابطه علت و معلولی را اثبات نمی‌کند.")
    elif chart_type in ("فراوانی دسته‌ها", "پارتو دسته‌ها") and ncat:
        col = st.selectbox("متغیر دسته‌ای", ncat, key="chart_cat")
        top_n = st.slider("تعداد دسته‌های نمایش‌داده‌شده", 3, 30, 10)
        counts = filtered[col].fillna("(مفقود)").astype(str).value_counts().head(top_n).rename_axis("دسته").reset_index(name="تعداد")
        counts["سهم تجمعی (%)"] = counts["تعداد"].cumsum() / max(counts["تعداد"].sum(), 1) * 100
        if chart_type == "پارتو دسته‌ها":
            fig = go.Figure()
            fig.add_bar(x=counts["دسته"], y=counts["تعداد"], name="تعداد")
            fig.add_scatter(x=counts["دسته"], y=counts["سهم تجمعی (%)"], name="سهم تجمعی (%)", yaxis="y2", mode="lines+markers")
            fig.update_layout(
                title=f"تحلیل پارتو {col}",
                yaxis=dict(title="تعداد"),
                yaxis2=dict(title="سهم تجمعی (%)", overlaying="y", side="right", range=[0, 100]),
            )
            st.plotly_chart(style_fig(fig, 480), use_container_width=True)
        else:
            fig = px.bar(counts, x="تعداد", y="دسته", orientation="h", text="تعداد", title=f"فراوانی {col}")
            fig.update_yaxes(categoryorder="total ascending")
            st.plotly_chart(style_fig(fig, 450), use_container_width=True)
        st.dataframe(counts, use_container_width=True)
    else:
        st.info("برای این نمودار، ستون‌های موردنیاز در داده فعلی وجود ندارند.")

# -------------------------- Group comparison --------------------------
with tabs[2]:
    st.subheader("مقایسه عملکرد گروه‌ها")
    if nnum and ncat:
        c1, c2, c3 = st.columns(3)
        group_col = c1.selectbox("ستون گروه‌بندی", ncat, key="group_col")
        value_col = c2.selectbox("شاخص عددی", nnum, key="group_val")
        aggregation = c3.selectbox("تجمیع", ["مجموع", "میانگین", "میانه", "تعداد", "بیشینه", "کمینه"])
        agg_map = {"مجموع": "sum", "میانگین": "mean", "میانه": "median", "تعداد": "count", "بیشینه": "max", "کمینه": "min"}
        result = filtered.groupby(group_col, dropna=False)[value_col].agg(agg_map[aggregation]).reset_index()
        result = result.sort_values(value_col, ascending=False)
        result[group_col] = result[group_col].fillna("(مفقود)").astype(str)
        fig = px.bar(result.head(30), x=value_col, y=group_col, orientation="h", text_auto=".3s", title=f"{aggregation} {value_col} به تفکیک {group_col}")
        fig.update_yaxes(categoryorder="total ascending")
        st.plotly_chart(style_fig(fig, 480), use_container_width=True)
        st.dataframe(result, use_container_width=True)

        if len(nnum) >= 2:
            st.markdown("#### ماتریس همبستگی")
            corr_cols = st.multiselect("متغیرهای همبستگی (حداکثر ۳۰ ستون)", nnum, default=nnum[:min(10, len(nnum))])
            corr_cols = corr_cols[:MAX_CORRELATION_COLUMNS]
            method = st.selectbox("روش همبستگی", ["pearson", "spearman", "kendall"])
            if len(corr_cols) >= 2:
                corr = filtered[corr_cols].corr(method=method)
                fig = px.imshow(corr, text_auto=".2f", zmin=-1, zmax=1, color_continuous_scale="RdBu_r", title=f"همبستگی ({method})")
                st.plotly_chart(style_fig(fig, 500), use_container_width=True)
            else:
                st.info("حداقل دو متغیر انتخاب کنید.")
    else:
        st.info("برای مقایسه گروهی، حداقل یک ستون دسته‌ای و یک ستون عددی لازم است.")

# -------------------------- Time series --------------------------
with tabs[3]:
    st.subheader("پایش روند زمانی")
    if ndates and nnum:
        dcol, vcol = st.columns(2)
        date_col = dcol.selectbox("ستون تاریخ", ndates, key="trend_date")
        value_col = vcol.selectbox("شاخص عددی", nnum, key="trend_value")
        frequency = st.selectbox("بازه زمانی", ["روزانه", "هفتگی", "ماهانه", "فصلی", "سالانه"], index=2)
        agg_label = st.selectbox("روش تجمیع", ["مجموع", "میانگین", "میانه", "تعداد"])
        window = st.slider("طول پنجره میانگین متحرک (تعداد دوره)", 1, 24, 3)
        freq_map = {"روزانه": "D", "هفتگی": "W", "ماهانه": "MS", "فصلی": "QS", "سالانه": "YS"}
        agg_map = {"مجموع": "sum", "میانگین": "mean", "میانه": "median", "تعداد": "count"}
        ts_data = filtered[[date_col, value_col]].dropna().copy()
        if not ts_data.empty:
            ts_data[date_col] = pd.to_datetime(ts_data[date_col], errors="coerce")
            ts_data = ts_data.dropna(subset=[date_col]).set_index(date_col).sort_index()
            series = ts_data[value_col].resample(freq_map[frequency]).agg(agg_map[agg_label])
            if len(series) >= 1:
                trend = series.to_frame("مقدار")
                trend["میانگین متحرک"] = trend["مقدار"].rolling(window, min_periods=1).mean()
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=trend.index, y=trend["مقدار"], mode="lines+markers", name="مقدار"))
                fig.add_trace(go.Scatter(x=trend.index, y=trend["میانگین متحرک"], mode="lines", name="میانگین متحرک", line=dict(dash="dash", width=3)))
                fig.update_layout(title=f"روند {agg_label} {value_col}", xaxis_title="تاریخ", yaxis_title=value_col)
                st.plotly_chart(style_fig(fig, 480), use_container_width=True)
                st.dataframe(trend.reset_index(), use_container_width=True)
                st.caption("میانگین متحرک برای هموارسازی نوسانات است و پیش‌بینی آینده محسوب نمی‌شود.")
            else:
                st.info("داده زمانی کافی برای رسم نمودار وجود ندارد.")
        else:
            st.info("در ستون‌های انتخاب‌شده داده معتبر کافی وجود ندارد.")
    else:
        st.info("برای تحلیل زمانی به یک ستون تاریخ معتبر و یک ستون عددی نیاز است.")

# -------------------------- Data quality --------------------------
with tabs[4]:
    st.subheader("گزارش کیفیت داده")
    quality = quality_report(filtered)
    st.dataframe(quality, use_container_width=True, hide_index=True)
    missing = quality[quality["تعداد مقادیر گمشده"] > 0].sort_values("درصد مقادیر گمشده", ascending=False)
    if not missing.empty:
        left, right = st.columns([3, 2])
        with left:
            fig = px.bar(missing, x="درصد مقادیر گمشده", y="ستون", orientation="h", text="درصد مقادیر گمشده", title="درصد داده گمشده به تفکیک ستون")
            fig.update_yaxes(categoryorder="total ascending")
            st.plotly_chart(style_fig(fig, 430), use_container_width=True)
        with right:
            st.markdown("**ستون‌های دارای مقادیر گمشده**")
            st.dataframe(missing[["ستون", "تعداد مقادیر گمشده", "درصد مقادیر گمشده"]], use_container_width=True, hide_index=True)
    else:
        st.success("هیچ مقدار گمشده‌ای در داده‌های فیلترشده شناسایی نشد.")

    if nnum:
        st.markdown("#### تشخیص داده‌های پرت با روش IQR")
        outliers = outlier_summary(filtered, nnum)
        st.caption("داده پرت به معنی داده اشتباه نیست؛ ممکن است رخداد واقعی، خطای ثبت یا موردی نیازمند بررسی باشد.")
        st.dataframe(outliers, use_container_width=True, hide_index=True)
    st.markdown("#### گزارش آماده‌سازی")
    st.write(f"ردیف‌های اولیه: **{prep['original_rows']:,}**")
    st.write(f"ستون‌های اولیه: **{prep['original_cols']:,}**")
    st.write(f"تبدیل به عدد: **{len(prep['numeric_converted'])}** ستون")
    st.write(f"تبدیل به تاریخ: **{len(prep['dates_converted'])}** ستون")
    st.write(f"تکراری‌های شناسایی‌شده پیش از حذف: **{prep['duplicates_found']:,}**")
    st.write(f"تکراری‌های حذف‌شده در آماده‌سازی: **{prep['duplicates_removed']:,}**")

# -------------------------- Export --------------------------
with tabs[5]:
    st.subheader("خروجی مناسب گزارش جلسه")
    st.write("خروجی‌ها بر اساس فیلترهای فعلی ساخته می‌شوند.")
    st.caption("فایل Excel شامل خلاصه مدیریتی، داده فیلترشده، کیفیت داده، آمار توصیفی و گزارش داده‌های پرت است.")
    base_name = safe_filename(re.sub(r"\.[^.]+$", "", source))
    csv_bytes = filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ دانلود داده فیلترشده CSV", data=csv_bytes,
        file_name=f"{base_name}_filtered.csv", mime="text/csv",
        use_container_width=True,
    )
    try:
        kpi_for_report = st.session_state.get("executive_kpi")
        excel_bytes = make_excel_report(filtered, source, prep, kpi_for_report)
        st.download_button(
            "⬇️ دانلود گزارش مدیریتی چندشیتی Excel", data=excel_bytes,
            file_name=f"{base_name}_management_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as exc:
        st.error(f"ساخت گزارش Excel با خطا مواجه شد: {exc}")

    st.divider()
    st.markdown("#### محدودیت‌های تفسیر")
    st.write("- داشبورد الگوهای آماری را نشان می‌دهد؛ به‌تنهایی علت رخدادها را اثبات نمی‌کند.")
    st.write("- تعریف KPI، واحد اندازه‌گیری، اهداف سازمانی و قواعد اعتبارسنجی باید با مدیر یا مالک داده تأیید شوند.")
    st.write("- قبل از توزیع گزارش، وجود داده محرمانه یا اطلاعات شخصی در فایل بررسی شود.")
