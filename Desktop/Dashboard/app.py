from __future__ import annotations

import html
import io
import re
import warnings
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# Universal Executive Dashboard | Bilingual (FA / EN)
# Run:  streamlit run app.py
# Requirements: streamlit plotly pandas numpy openpyxl xlrd statsmodels
# ============================================================

st.set_page_config(
    page_title="Executive Dashboard | داشبورد مدیریتی",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MAX_UPLOAD_MB = 200
MAX_CORRELATION_COLUMNS = 30
EXCEL_ROW_LIMIT = 1_000_000

# ------------------------------------------------------------
# Streamlit version compatibility (use_container_width -> width)
# ------------------------------------------------------------
_ver = tuple(int(x) for x in re.findall(r"\d+", st.__version__)[:2])
STRETCH = {"width": "stretch"} if _ver >= (1, 50) else {"use_container_width": True}

# ------------------------------------------------------------
# Translations
# ------------------------------------------------------------
TR: dict[str, dict[str, str]] = {
    "fa": {
        "title": "داشبورد گزارش‌دهی مدیریتی",
        "tagline": "تحلیل داده، پایش شاخص‌ها، شناسایی ریسک‌ها و آماده‌سازی گزارش جلسه",
        "settings": "📂 تنظیمات گزارش",
        "upload": "بارگذاری فایل Excel یا CSV",
        "sample": "استفاده از داده نمونه نمایشی",
        "dedup": "حذف ردیف‌های تکراری",
        "max_size": "حداکثر حجم فایل: {mb} مگابایت",
        "too_big": "حجم فایل از سقف {mb} مگابایت بیشتر است.",
        "pick_sheet": "انتخاب شیت Excel",
        "read_fail": "فایل خوانده نشد: {err}",
        "csv_fail": "خواندن CSV ناموفق بود. جداکننده یا Encoding فایل را بررسی کنید.",
        "reading": "در حال خواندن فایل...",
        "welcome": "برای شروع، از نوار کناری یک فایل Excel/CSV بارگذاری کنید یا داده نمونه نمایشی را فعال کنید.",
        "empty_raw": "فایل خالی است یا داده‌ای برای تحلیل وجود ندارد.",
        "empty_prep": "پس از آماده‌سازی، داده قابل تحلیلی باقی نمانده است.",
        "sample_source": "داده نمونه نمایشی (ساختگی)",
        "source_line": "منبع: {src} &nbsp;|&nbsp; تهیه گزارش: {ts}",
        "filters": "🔎 فیلترهای سراسری",
        "filter_cols": "انتخاب ستون برای فیلتر",
        "range_of": "بازه {col}",
        "values_of": "مقادیر {col}",
        "filter_skipped": "فیلتر «{col}» به دلیل بیش از ۳۰۰ مقدار یکتا نمایش داده نشد.",
        "no_rows": "با فیلترهای فعلی هیچ ردیفی باقی نمانده است. فیلترها را تغییر دهید.",
        "records_line": "رکوردهای اولیه: {a} | رکوردهای پس از فیلتر: {b} | ستون‌ها: {c}",
        "billion": "میلیارد",
        "million": "میلیون",
        "missing": "(مفقود)",
        "tab_summary": "📌 خلاصه مدیریتی",
        "tab_charts": "📈 تحلیل نموداری",
        "tab_groups": "🧩 مقایسه گروه‌ها",
        "tab_trend": "⏱️ روند زمانی",
        "tab_quality": "🛡️ کیفیت داده",
        "tab_export": "📤 خروجی گزارش",
        # summary
        "overview": "نمای کلی برای جلسه مدیریتی",
        "m_records": "تعداد رکورد",
        "m_cols": "تعداد ستون",
        "m_complete": "کامل بودن داده",
        "m_missing": "مقادیر گمشده",
        "m_quality": "امتیاز کیفیت تخمینی",
        "quality_note": "امتیاز کیفیت یک شاخص داخلی و تقریبی است؛ جایگزین اعتبارسنجی تخصصی یا کنترل صحت کسب‌وکار نیست.",
        "kpi_title": "شاخص کلیدی عملکرد (KPI)",
        "kpi_pick": "شاخص عددی مورد بررسی",
        "m_sum": "مجموع",
        "m_mean": "میانگین",
        "m_median": "میانه",
        "m_max": "بیشینه",
        "m_min": "کمینه",
        "period_title": "مقایسه دو دوره (در صورت وجود ستون تاریخ)",
        "date_pick": "ستون تاریخ برای مقایسه",
        "period_len": "طول هر دوره مقایسه (روز)",
        "period_method": "روش مقایسه",
        "m_recent": "دوره اخیر",
        "m_prev": "دوره قبل",
        "m_change": "تغییر نسبت به دوره قبل",
        "na_change": "قابل محاسبه نیست",
        "delta_info": "برای محاسبه تغییر درصدی، هر دو دوره باید داده معتبر داشته باشند و مقدار دوره قبل صفر نباشد.",
        "range_help": "از {a} تا {b}",
        "need_date": "برای مقایسه دوره‌ای، فایل باید حداقل یک ستون تاریخ معتبر داشته باشد.",
        "insights": "برداشت‌های خودکار",
        "insight_top": "بیشترین سهم «{kpi}» متعلق به «{grp}» در ستون «{col}» است ({share}% از کل).",
        "notes_title": "نکات قابل طرح در جلسه",
        "warn_missing": "در داده‌های فیلترشده {n} مقدار گمشده وجود دارد؛ پیش از تصمیم‌گیری، اثر آن بر شاخص‌های اصلی بررسی شود.",
        "ok_missing": "در داده‌های فیلترشده مقدار گمشده‌ای شناسایی نشد.",
        "warn_dups": "{n} ردیف تکراری در داده‌های فعلی دیده می‌شود.",
        "info_outlier": "بیشترین تعداد مشاهده خارج از محدوده IQR در ستون «{col}» دیده شده است ({n} مورد). این موارد الزاماً خطا نیستند و باید بررسی شوند.",
        "no_numeric": "ستون عددی شناسایی نشد؛ KPI عددی و تحلیل‌های کمی در دسترس نیست.",
        "preview": "👁️ پیش‌نمایش داده (۵۰ ردیف اول)",
        # charts
        "chart_title": "تحلیل نموداری اکتشافی",
        "analysis_type": "نوع تحلیل",
        "ct_hist": "توزیع عددی",
        "ct_box": "نمودار جعبه‌ای",
        "ct_scatter": "پراکندگی دو متغیر",
        "ct_freq": "فراوانی دسته‌ها",
        "ct_pareto": "پارتو دسته‌ها",
        "num_var": "متغیر عددی",
        "no_valid_num": "برای این ستون مقدار عددی معتبر وجود ندارد.",
        "bins": "تعداد بازه‌های هیستوگرام",
        "hist_title": "توزیع {col}",
        "stats_line": "تعداد معتبر: {n} | انحراف معیار: {sd} | چولگی: {sk}",
        "group_opt": "گروه‌بندی اختیاری",
        "no_group": "بدون گروه",
        "box_title": "نمودار جعبه‌ای {col}",
        "x_axis": "محور افقی (X)",
        "y_axis": "محور عمودی (Y)",
        "color_opt": "رنگ بر اساس دسته (اختیاری)",
        "no_color": "بدون دسته",
        "show_trend": "نمایش خط روند خطی",
        "scatter_title": "رابطه {x} و {y}",
        "trend_needs_sm": "برای خط روند باید بسته statsmodels نصب باشد: pip install statsmodels",
        "pearson": "ضریب همبستگی پیرسون",
        "corr_note": "همبستگی به‌تنهایی رابطه علت و معلولی را اثبات نمی‌کند.",
        "cat_var": "متغیر دسته‌ای",
        "top_n": "تعداد دسته‌های نمایش‌داده‌شده",
        "cat": "دسته",
        "count": "تعداد",
        "cum_share": "سهم تجمعی (%)",
        "pareto_title": "تحلیل پارتو {col}",
        "freq_title": "فراوانی {col}",
        "chart_missing_cols": "برای این نمودار، ستون‌های موردنیاز در داده فعلی وجود ندارند.",
        # groups
        "groups_title": "مقایسه عملکرد گروه‌ها",
        "group_col": "ستون گروه‌بندی",
        "value_col": "شاخص عددی",
        "aggregation": "تجمیع",
        "group_bar_title": "{agg} {val} به تفکیک {grp}",
        "corr_matrix": "ماتریس همبستگی",
        "corr_vars": "متغیرهای همبستگی (حداکثر ۳۰ ستون)",
        "corr_method": "روش همبستگی",
        "corr_title": "همبستگی ({m})",
        "pick_two": "حداقل دو متغیر انتخاب کنید.",
        "groups_need": "برای مقایسه گروهی، حداقل یک ستون دسته‌ای و یک ستون عددی لازم است.",
        # trend
        "trend_title": "پایش روند زمانی",
        "date_col": "ستون تاریخ",
        "freq": "بازه زمانی",
        "agg_method": "روش تجمیع",
        "window": "طول پنجره میانگین متحرک (تعداد دوره)",
        "f_D": "روزانه",
        "f_W": "هفتگی",
        "f_M": "ماهانه",
        "f_Q": "فصلی",
        "f_Y": "سالانه",
        "value": "مقدار",
        "moving_avg": "میانگین متحرک",
        "date": "تاریخ",
        "trend_fig_title": "روند {agg} {val}",
        "trend_note": "میانگین متحرک برای هموارسازی نوسانات است و پیش‌بینی آینده محسوب نمی‌شود.",
        "trend_few": "داده زمانی کافی برای رسم نمودار وجود ندارد.",
        "trend_invalid": "در ستون‌های انتخاب‌شده داده معتبر کافی وجود ندارد.",
        "trend_need": "برای تحلیل زمانی به یک ستون تاریخ معتبر و یک ستون عددی نیاز است.",
        # quality
        "quality_title": "گزارش کیفیت داده",
        "q_col": "ستون",
        "q_dtype": "نوع داده",
        "q_missing_n": "تعداد مقادیر گمشده",
        "q_missing_pct": "درصد مقادیر گمشده",
        "q_unique_n": "تعداد مقادیر یکتا",
        "q_unique_pct": "درصد یکتایی نسبت به ردیف‌ها",
        "missing_chart": "درصد داده گمشده به تفکیک ستون",
        "missing_cols": "ستون‌های دارای مقادیر گمشده",
        "no_missing": "هیچ مقدار گمشده‌ای در داده‌های فیلترشده شناسایی نشد.",
        "outliers_title": "تشخیص داده‌های پرت با روش IQR",
        "outliers_note": "داده پرت به معنی داده اشتباه نیست؛ ممکن است رخداد واقعی، خطای ثبت یا موردی نیازمند بررسی باشد.",
        "o_var": "متغیر",
        "o_valid": "تعداد داده معتبر",
        "o_low": "حد پایین IQR",
        "o_high": "حد بالا IQR",
        "o_count": "تعداد داده پرت",
        "o_pct": "درصد داده پرت",
        "prep_title": "گزارش آماده‌سازی",
        "prep_rows": "ردیف‌های اولیه",
        "prep_cols": "ستون‌های اولیه",
        "prep_num": "ستون‌های تبدیل‌شده به عدد",
        "prep_dates": "ستون‌های تبدیل‌شده به تاریخ",
        "prep_dups_found": "تکراری‌های شناسایی‌شده پیش از حذف",
        "prep_dups_removed": "تکراری‌های حذف‌شده در آماده‌سازی",
        # export
        "export_title": "خروجی مناسب گزارش جلسه",
        "export_note1": "خروجی‌ها بر اساس فیلترهای فعلی ساخته می‌شوند.",
        "export_note2": "فایل Excel شامل خلاصه مدیریتی، داده فیلترشده، کیفیت داده، آمار توصیفی و گزارش داده‌های پرت است.",
        "dl_csv": "⬇️ دانلود داده فیلترشده (CSV)",
        "dl_xlsx": "⬇️ دانلود گزارش مدیریتی چندشیتی (Excel)",
        "xlsx_fail": "ساخت گزارش Excel با خطا مواجه شد: {err}",
        "limits_title": "محدودیت‌های تفسیر",
        "limit1": "داشبورد الگوهای آماری را نشان می‌دهد؛ به‌تنهایی علت رخدادها را اثبات نمی‌کند.",
        "limit2": "تعریف KPI، واحد اندازه‌گیری، اهداف سازمانی و قواعد اعتبارسنجی باید با مدیر یا مالک داده تأیید شوند.",
        "limit3": "قبل از توزیع گزارش، وجود داده محرمانه یا اطلاعات شخصی در فایل بررسی شود.",
        # excel summary
        "x_indicator": "شاخص",
        "x_value": "مقدار",
        "x_source": "منبع داده",
        "x_date": "تاریخ تهیه گزارش",
        "x_rows": "تعداد ردیف پس از فیلتر",
        "x_cols": "تعداد ستون",
        "x_complete": "درصد کامل بودن",
        "x_dups_removed": "تعداد تکراری حذف‌شده در آماده‌سازی",
        "x_kpi": "شاخص انتخاب‌شده",
        "x_none": "انتخاب نشده",
        "unnamed": "ستون_بدون_نام",
        # aggs
        "agg_sum": "مجموع",
        "agg_mean": "میانگین",
        "agg_median": "میانه",
        "agg_count": "تعداد",
        "agg_max": "بیشینه",
        "agg_min": "کمینه",
        # sample
        "s_date": "تاریخ",
        "s_region": "منطقه",
        "s_qty": "تعداد",
        "s_price": "قیمت واحد",
        "s_sales": "فروش",
        "s_csat": "رضایت مشتری",
        "s_regions": "شمال|جنوب|مرکز|غرب",
        "footer": "ساخته‌شده با Streamlit · داشبورد گزارش‌دهی مدیریتی",
    },
    "en": {
        "title": "Management Reporting Dashboard",
        "tagline": "Data analysis, KPI monitoring, risk detection and meeting-ready reporting",
        "settings": "📂 Report settings",
        "upload": "Upload an Excel or CSV file",
        "sample": "Use demo sample data",
        "dedup": "Remove duplicate rows",
        "max_size": "Max file size: {mb} MB",
        "too_big": "File exceeds the {mb} MB limit.",
        "pick_sheet": "Select Excel sheet",
        "read_fail": "Could not read the file: {err}",
        "csv_fail": "Failed to read CSV. Check the delimiter or file encoding.",
        "reading": "Reading file...",
        "welcome": "To start, upload an Excel/CSV file from the sidebar or enable the demo sample data.",
        "empty_raw": "The file is empty or contains no data to analyze.",
        "empty_prep": "No analyzable data remains after preparation.",
        "sample_source": "Demo sample data (synthetic)",
        "source_line": "Source: {src} &nbsp;|&nbsp; Report generated: {ts}",
        "filters": "🔎 Global filters",
        "filter_cols": "Select columns to filter",
        "range_of": "Range of {col}",
        "values_of": "Values of {col}",
        "filter_skipped": "Filter “{col}” hidden: more than 300 unique values.",
        "no_rows": "No rows left with the current filters. Please adjust the filters.",
        "records_line": "Initial records: {a} | After filters: {b} | Columns: {c}",
        "billion": "B",
        "million": "M",
        "missing": "(missing)",
        "tab_summary": "📌 Executive summary",
        "tab_charts": "📈 Chart analysis",
        "tab_groups": "🧩 Group comparison",
        "tab_trend": "⏱️ Time trend",
        "tab_quality": "🛡️ Data quality",
        "tab_export": "📤 Report export",
        "overview": "Overview for the management meeting",
        "m_records": "Records",
        "m_cols": "Columns",
        "m_complete": "Data completeness",
        "m_missing": "Missing values",
        "m_quality": "Estimated quality score",
        "quality_note": "The quality score is an internal, indicative metric; it is not a substitute for expert validation or business-accuracy checks.",
        "kpi_title": "Key Performance Indicator (KPI)",
        "kpi_pick": "Numeric indicator to review",
        "m_sum": "Total",
        "m_mean": "Mean",
        "m_median": "Median",
        "m_max": "Maximum",
        "m_min": "Minimum",
        "period_title": "Two-period comparison (requires a date column)",
        "date_pick": "Date column for comparison",
        "period_len": "Length of each period (days)",
        "period_method": "Comparison method",
        "m_recent": "Recent period",
        "m_prev": "Previous period",
        "m_change": "Change vs. previous period",
        "na_change": "Not computable",
        "delta_info": "To compute a percentage change, both periods need valid data and the previous value must not be zero.",
        "range_help": "From {a} to {b}",
        "need_date": "Period comparison needs at least one valid date column in the file.",
        "insights": "Automatic insights",
        "insight_top": "The largest share of “{kpi}” belongs to “{grp}” in column “{col}” ({share}% of total).",
        "notes_title": "Talking points for the meeting",
        "warn_missing": "The filtered data has {n} missing values; review their impact on key indicators before deciding.",
        "ok_missing": "No missing values were found in the filtered data.",
        "warn_dups": "{n} duplicate rows were found in the current data.",
        "info_outlier": "Most IQR outliers were found in column “{col}” ({n} cases). These are not necessarily errors and should be reviewed.",
        "no_numeric": "No numeric column detected; numeric KPIs and quantitative analyses are unavailable.",
        "preview": "👁️ Data preview (first 50 rows)",
        "chart_title": "Exploratory chart analysis",
        "analysis_type": "Analysis type",
        "ct_hist": "Numeric distribution",
        "ct_box": "Box plot",
        "ct_scatter": "Two-variable scatter",
        "ct_freq": "Category frequency",
        "ct_pareto": "Category Pareto",
        "num_var": "Numeric variable",
        "no_valid_num": "This column has no valid numeric values.",
        "bins": "Number of histogram bins",
        "hist_title": "Distribution of {col}",
        "stats_line": "Valid count: {n} | Std. dev.: {sd} | Skewness: {sk}",
        "group_opt": "Optional grouping",
        "no_group": "No grouping",
        "box_title": "Box plot of {col}",
        "x_axis": "X axis",
        "y_axis": "Y axis",
        "color_opt": "Color by category (optional)",
        "no_color": "No category",
        "show_trend": "Show linear trendline",
        "scatter_title": "Relationship between {x} and {y}",
        "trend_needs_sm": "The trendline requires statsmodels: pip install statsmodels",
        "pearson": "Pearson correlation",
        "corr_note": "Correlation alone does not prove causation.",
        "cat_var": "Categorical variable",
        "top_n": "Number of categories shown",
        "cat": "Category",
        "count": "Count",
        "cum_share": "Cumulative share (%)",
        "pareto_title": "Pareto analysis of {col}",
        "freq_title": "Frequency of {col}",
        "chart_missing_cols": "The columns required for this chart are not available in the current data.",
        "groups_title": "Group performance comparison",
        "group_col": "Grouping column",
        "value_col": "Numeric indicator",
        "aggregation": "Aggregation",
        "group_bar_title": "{agg} of {val} by {grp}",
        "corr_matrix": "Correlation matrix",
        "corr_vars": "Correlation variables (max 30 columns)",
        "corr_method": "Correlation method",
        "corr_title": "Correlation ({m})",
        "pick_two": "Select at least two variables.",
        "groups_need": "Group comparison needs at least one categorical and one numeric column.",
        "trend_title": "Time trend monitoring",
        "date_col": "Date column",
        "freq": "Time interval",
        "agg_method": "Aggregation method",
        "window": "Moving-average window (periods)",
        "f_D": "Daily",
        "f_W": "Weekly",
        "f_M": "Monthly",
        "f_Q": "Quarterly",
        "f_Y": "Yearly",
        "value": "Value",
        "moving_avg": "Moving average",
        "date": "Date",
        "trend_fig_title": "{agg} trend of {val}",
        "trend_note": "The moving average smooths fluctuations; it is not a forecast.",
        "trend_few": "Not enough time-series data to draw the chart.",
        "trend_invalid": "Not enough valid data in the selected columns.",
        "trend_need": "Time analysis needs a valid date column and a numeric column.",
        "quality_title": "Data quality report",
        "q_col": "Column",
        "q_dtype": "Data type",
        "q_missing_n": "Missing values",
        "q_missing_pct": "Missing %",
        "q_unique_n": "Unique values",
        "q_unique_pct": "Uniqueness % of rows",
        "missing_chart": "Missing data % by column",
        "missing_cols": "Columns with missing values",
        "no_missing": "No missing values found in the filtered data.",
        "outliers_title": "Outlier detection (IQR method)",
        "outliers_note": "An outlier is not necessarily wrong; it may be a real event, an entry error, or a case needing review.",
        "o_var": "Variable",
        "o_valid": "Valid values",
        "o_low": "IQR lower bound",
        "o_high": "IQR upper bound",
        "o_count": "Outlier count",
        "o_pct": "Outlier %",
        "prep_title": "Preparation report",
        "prep_rows": "Original rows",
        "prep_cols": "Original columns",
        "prep_num": "Columns converted to numeric",
        "prep_dates": "Columns converted to date",
        "prep_dups_found": "Duplicates found before removal",
        "prep_dups_removed": "Duplicates removed during preparation",
        "export_title": "Meeting-ready export",
        "export_note1": "Exports are built from the current filters.",
        "export_note2": "The Excel file contains the executive summary, filtered data, data quality, descriptive statistics and outlier report.",
        "dl_csv": "⬇️ Download filtered data (CSV)",
        "dl_xlsx": "⬇️ Download multi-sheet management report (Excel)",
        "xlsx_fail": "Excel report generation failed: {err}",
        "limits_title": "Interpretation limits",
        "limit1": "The dashboard shows statistical patterns; it does not by itself prove the causes of events.",
        "limit2": "KPI definitions, units, organizational targets and validation rules must be confirmed with the manager or data owner.",
        "limit3": "Before distributing the report, check the file for confidential or personal data.",
        "x_indicator": "Indicator",
        "x_value": "Value",
        "x_source": "Data source",
        "x_date": "Report date",
        "x_rows": "Rows after filters",
        "x_cols": "Columns",
        "x_complete": "Completeness %",
        "x_dups_removed": "Duplicates removed during preparation",
        "x_kpi": "Selected KPI",
        "x_none": "Not selected",
        "unnamed": "unnamed_column",
        "agg_sum": "Sum",
        "agg_mean": "Mean",
        "agg_median": "Median",
        "agg_count": "Count",
        "agg_max": "Maximum",
        "agg_min": "Minimum",
        "s_date": "Date",
        "s_region": "Region",
        "s_qty": "Quantity",
        "s_price": "Unit price",
        "s_sales": "Sales",
        "s_csat": "Customer satisfaction",
        "s_regions": "North|South|Central|West",
        "footer": "Built with Streamlit · Management Reporting Dashboard",
    },
}

LANG = "fa"


def t(key: str, **kw: Any) -> str:
    text = TR[LANG].get(key) or TR["en"].get(key) or key
    return text.format(**kw) if kw else text


# ------------------------------------------------------------
# Sidebar: language switch first (drives the whole page)
# ------------------------------------------------------------
with st.sidebar:
    st.radio(
        "🌐 Language / زبان",
        ["fa", "en"],
        format_func=lambda x: {"fa": "فارسی", "en": "English"}[x],
        key="lang",
        horizontal=True,
    )
LANG = st.session_state.get("lang", "fa")
IS_RTL = LANG == "fa"

# ------------------------------------------------------------
# Styling: sidebar is ALWAYS on the right, text direction follows language
# ------------------------------------------------------------
CSS = """
<style>
@import url('https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css');
html, body, .stApp,
.stApp :is(p, h1, h2, h3, h4, h5, h6, label, li, input, textarea, button, td, th,
           [data-testid="stMetricLabel"], [data-testid="stMetricValue"], [data-testid="stCaptionContainer"]) {
    font-family: Vazirmatn, Tahoma, 'Segoe UI', sans-serif;
}
.stApp {direction: __DIR__;}

/* Keep the sidebar on the right-hand side in both languages */
[data-testid="stAppViewContainer"] {flex-direction: row-reverse; direction: ltr;}
[data-testid="stMain"], section.main, [data-testid="stSidebar"] {direction: __DIR__;}
[data-testid="stSidebar"] {border-left: 1px solid #e2e8f0;}
.stApp :is(h1, h2, h3, h4, h5, p, label, li, .stMarkdown, .stCaption) {text-align: __ALIGN__;}

/* Charts and tables always render left-to-right internally */
[data-testid="stDataFrame"], .stPlotlyChart, [data-testid="stPlotlyChart"] {direction: ltr;}

/* On desktop the sidebar stays pinned (no collapse button that would flip sides) */
@media (min-width: 992px) {
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="stExpandSidebarButton"] {display: none !important;}
}

[data-testid="stMetric"] {
    background: linear-gradient(135deg, #ffffff, #f3f6fb);
    border: 1px solid #dbe3ef; border-__SIDE__: 5px solid #2563eb;
    border-radius: 12px; padding: 14px 16px;
    box-shadow: 0 2px 8px rgba(15, 23, 42, .06);
}
[data-testid="stMetricLabel"] {font-weight: 600; color: #475569;}
[data-testid="stMetricValue"] {font-weight: 800; color: #0f172a;}
.hero {
    background: linear-gradient(120deg, #102a56, #2563eb);
    color: white; padding: 24px 28px; border-radius: 16px; margin-bottom: 18px;
    box-shadow: 0 6px 20px rgba(37, 99, 235, .25);
}
.hero h1, .hero p {color: white; margin: 0; text-align: __ALIGN__;}
.hero p {margin-top: 8px; opacity: .92;}
.footer {text-align: center; color: #94a3b8; font-size: 12px; margin-top: 32px;}

/* Responsive layout for phones and small tablets */
@media (max-width: 768px) {
    /* Do not force the sidebar to sit beside the page on small screens. */
    [data-testid="stAppViewContainer"] {
        flex-direction: row;
        width: 100%;
        min-width: 0;
    }
    [data-testid="stMain"], section.main, [data-testid="stMainBlockContainer"] {
        min-width: 0 !important;
        width: 100% !important;
        max-width: 100% !important;
    }
    [data-testid="stMainBlockContainer"] {
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        padding-top: 1rem !important;
    }
    [data-testid="stSidebar"] {
        max-width: min(86vw, 340px);
    }
    .hero {
        padding: 18px 16px;
        border-radius: 12px;
        margin-bottom: 12px;
        overflow-wrap: anywhere;
    }
    .hero h1 {font-size: 1.35rem !important; line-height: 1.5;}
    .hero p {font-size: 0.9rem !important; line-height: 1.8;}
    [data-testid="stMetric"] {
        padding: 10px 9px;
        border-radius: 10px;
        min-width: 0;
    }
    [data-testid="stMetricLabel"] {font-size: 0.78rem !important;}
    [data-testid="stMetricValue"] {font-size: clamp(1rem, 4vw, 1.4rem) !important; overflow-wrap: anywhere;}
    [data-testid="stTabs"] [role="tablist"] {
        gap: 0.25rem;
        overflow-x: auto !important;
        flex-wrap: nowrap !important;
        scrollbar-width: thin;
    }
    [data-testid="stTabs"] button[role="tab"] {
        flex: 0 0 auto;
        white-space: nowrap;
        font-size: 0.82rem;
        padding: 0.45rem 0.65rem;
    }
    [data-testid="stPlotlyChart"], .stPlotlyChart,
    [data-testid="stDataFrame"], [data-testid="stTable"] {
        max-width: 100% !important;
        min-width: 0 !important;
    }
    [data-testid="stDataFrame"] {overflow-x: auto !important;}
    .stButton button, .stDownloadButton button {width: 100%; white-space: normal;}
    h1 {font-size: 1.45rem !important;}
    h2 {font-size: 1.25rem !important;}
    h3 {font-size: 1.1rem !important;}
}
</style>
"""
st.markdown(
    CSS.replace("__DIR__", "rtl" if IS_RTL else "ltr")
    .replace("__ALIGN__", "right" if IS_RTL else "left")
    .replace("__SIDE__", "right" if IS_RTL else "left"),
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩٫٬", "01234567890123456789.,")


def fmt(value: Any, digits: int = 2) -> str:
    """Format numbers for display without failing on missing values."""
    if value is None or pd.isna(value):
        return "—"
    try:
        value = float(value)
    except (TypeError, ValueError):
        return str(value)
    if abs(value) >= 1_000_000_000:
        return f"{value / 1_000_000_000:,.2f} {t('billion')}"
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:,.2f} {t('million')}"
    if value.is_integer():
        return f"{int(value):,}"
    return f"{value:,.{digits}f}"


def load_csv(file_bytes: bytes) -> pd.DataFrame:
    for encoding in ("utf-8-sig", "utf-8", "cp1256", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding, sep=None, engine="python")
        except Exception:  # noqa: BLE001 - try the next encoding
            continue
    raise ValueError(t("csv_fail"))


@st.cache_data(show_spinner=False)
def read_excel_bytes(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)


@st.cache_data(show_spinner=False)
def excel_sheets(file_bytes: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


def _is_text_dtype(s: pd.Series) -> bool:
    return pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s)


def normalize_dataframe(raw: pd.DataFrame, remove_duplicates: bool, unnamed: str) -> tuple[pd.DataFrame, dict]:
    df = raw.copy()
    original_rows, original_cols = df.shape

    labels, seen = [], {}
    for i, column in enumerate(df.columns):
        name = str(column).strip() or f"{unnamed}_{i + 1}"
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1:
            name = f"{name}_{seen[name]}"
        labels.append(name)
    df.columns = labels

    converted_numeric, converted_dates = [], []
    for col in df.columns:
        if not _is_text_dtype(df[col]):
            continue
        series = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)
        series = series.replace(r"^\s*$", np.nan, regex=True)
        non_null = series.dropna()
        if non_null.empty:
            df[col] = series
            continue

        # Conservative numeric conversion (supports Persian/Arabic digits and separators)
        cleaned = series.astype("string").str.translate(_DIGITS).str.replace(",", "", regex=False)
        numeric = pd.to_numeric(cleaned, errors="coerce").astype("float64")
        ratio = numeric.notna().sum() / max(len(non_null), 1)
        has_leading_zero_codes = cleaned.dropna().str.match(r"^0\d+").any()  # e.g. phone / code columns
        if ratio >= 0.95 and not has_leading_zero_codes:
            df[col] = numeric
            converted_numeric.append(col)
            continue

        # Conservative date conversion
        sample = non_null.astype(str).head(200)
        looks_like_date = sample.str.match(r"^\s*\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(\s|$)").mean() >= 0.90
        if looks_like_date:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                parsed = pd.to_datetime(series, errors="coerce")
            if parsed.notna().sum() >= 0.90 * len(non_null):
                df[col] = parsed
                converted_dates.append(col)
                continue
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
    uniq = df.nunique(dropna=True).values
    return pd.DataFrame({
        t("q_col"): df.columns,
        t("q_dtype"): [str(x) for x in df.dtypes],
        t("q_missing_n"): df.isna().sum().values,
        t("q_missing_pct"): (df.isna().mean().values * 100).round(2),
        t("q_unique_n"): uniq,
        t("q_unique_pct"): (uniq / n * 100).round(2),
    })


def quality_score(df: pd.DataFrame) -> dict:
    if df.empty or df.size == 0:
        return {"score": 0.0, "completeness": 0.0, "uniqueness": 0.0, "constant_columns": 0}
    completeness = float(1 - df.isna().sum().sum() / df.size)
    uniqueness = float(1 - df.duplicated().sum() / max(len(df), 1))
    constant_columns = int((df.nunique(dropna=True) <= 1).sum())
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
            t("o_var"): col,
            t("o_valid"): len(s),
            t("o_low"): low,
            t("o_high"): high,
            t("o_count"): count,
            t("o_pct"): round(100 * count / max(len(s), 1), 2),
        })
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(t("o_count"), ascending=False).reset_index(drop=True)


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


def show(fig, height: int = 420, key: str | None = None):
    st.plotly_chart(style_fig(fig, height), key=key, **STRETCH)


def table(data: pd.DataFrame, hide_index: bool = False):
    st.dataframe(data, hide_index=hide_index, **STRETCH)


def agg_series(series: pd.Series, how: str) -> float:
    if how == "count":
        return float(series.count())
    if series.empty:
        return np.nan
    return float(getattr(series, how)())


def make_excel_report(data: pd.DataFrame, source: str, prep: dict, kpi: str | None) -> bytes:
    buffer = io.BytesIO()
    numeric, _, _ = classify_columns(data)
    summary = pd.DataFrame([
        {t("x_indicator"): t("x_source"), t("x_value"): source},
        {t("x_indicator"): t("x_date"), t("x_value"): pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")},
        {t("x_indicator"): t("x_rows"), t("x_value"): len(data)},
        {t("x_indicator"): t("x_cols"), t("x_value"): data.shape[1]},
        {t("x_indicator"): t("x_complete"), t("x_value"): round(100 * (1 - data.isna().sum().sum() / max(data.size, 1)), 2)},
        {t("x_indicator"): t("x_dups_removed"), t("x_value"): prep.get("duplicates_removed", 0)},
        {t("x_indicator"): t("x_kpi"), t("x_value"): kpi or t("x_none")},
    ])
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Executive_Summary", index=False)
        data.head(EXCEL_ROW_LIMIT).to_excel(writer, sheet_name="Filtered_Data", index=False)
        quality_report(data).to_excel(writer, sheet_name="Data_Quality", index=False)
        if numeric:
            data[numeric].describe().T.round(4).to_excel(writer, sheet_name="Descriptive_Stats")
            outs = outlier_summary(data, numeric)
            if not outs.empty:
                outs.to_excel(writer, sheet_name="Outliers", index=False)
        for ws in writer.book.worksheets:  # readable column widths
            for column_cells in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in list(column_cells)[:200])
                ws.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 10), 50)
            ws.sheet_view.rightToLeft = IS_RTL
    return buffer.getvalue()


# ------------------------------------------------------------
# Sidebar: data source
# ------------------------------------------------------------
st.sidebar.title(t("settings"))
uploaded = st.sidebar.file_uploader(t("upload"), type=["xlsx", "xls", "csv"])
use_sample = st.sidebar.toggle(t("sample"), value=False)
remove_duplicates = st.sidebar.checkbox(t("dedup"), value=False)
st.sidebar.caption(t("max_size", mb=MAX_UPLOAD_MB))


def hero(subtitle_html: str) -> None:
    st.markdown(
        f'<div class="hero"><h1>📊 {t("title")}</h1><p>{subtitle_html}</p></div>',
        unsafe_allow_html=True,
    )


raw = None
source = ""
prep: dict = {}
if uploaded is not None:
    if uploaded.size > MAX_UPLOAD_MB * 1024 * 1024:
        st.error(t("too_big", mb=MAX_UPLOAD_MB))
        st.stop()
    try:
        content = uploaded.getvalue()
        source = uploaded.name
        if uploaded.name.lower().endswith(".csv"):
            with st.spinner(t("reading")):
                raw = load_csv(content)
        else:
            sheets = excel_sheets(content)
            sheet = st.sidebar.selectbox(t("pick_sheet"), sheets)
            with st.spinner(t("reading")):
                raw = read_excel_bytes(content, sheet)
    except Exception as exc:  # noqa: BLE001
        st.error(t("read_fail", err=exc))
        st.stop()
elif use_sample:
    rng = np.random.default_rng(21)
    n = 900
    dates = pd.date_range("2025-01-01", periods=n, freq="D")
    region = rng.choice(t("s_regions").split("|"), n)
    quantity = rng.integers(1, 25, n)
    unit_price = rng.normal(250_000, 45_000, n).clip(30_000, None).round(0)
    raw = pd.DataFrame({
        t("s_date"): dates,
        t("s_region"): region,
        t("s_qty"): quantity,
        t("s_price"): unit_price,
        t("s_sales"): quantity * unit_price,
        t("s_csat"): rng.normal(4.1, .55, n).clip(1, 5).round(1),
    })
    raw.loc[rng.choice(n, 30, replace=False), t("s_csat")] = np.nan
    source = t("sample_source")
else:
    hero(t("tagline"))
    st.info(t("welcome"))
    st.stop()

if raw is None or raw.empty:
    st.error(t("empty_raw"))
    st.stop()

df, prep = normalize_dataframe(raw, remove_duplicates, t("unnamed"))
if df.empty or df.shape[1] == 0:
    st.error(t("empty_prep"))
    st.stop()

numeric_cols, categorical_cols, date_cols = classify_columns(df)

hero(t("source_line", src=html.escape(str(source)), ts=pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")))

# ------------------------------------------------------------
# Global filters
# ------------------------------------------------------------
with st.sidebar:
    st.divider()
    st.subheader(t("filters"))
    filter_columns = st.multiselect(t("filter_cols"), df.columns.tolist())
    mask = pd.Series(True, index=df.index)
    for col in filter_columns:
        if col in date_cols:
            valid = df[col].dropna()
            if valid.empty:
                continue
            date_range = st.date_input(
                t("range_of", col=col),
                value=(valid.min().date(), valid.max().date()),
                key=f"date_filter_{col}",
            )
            if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
                start = pd.Timestamp(date_range[0])
                end = pd.Timestamp(date_range[1]) + pd.Timedelta(days=1)
                mask &= df[col].between(start, end, inclusive="left") | df[col].isna()
        elif col in numeric_cols:
            valid = pd.to_numeric(df[col], errors="coerce").dropna()
            if valid.empty or valid.min() == valid.max():
                continue
            low, high = st.slider(
                t("range_of", col=col), float(valid.min()), float(valid.max()),
                (float(valid.min()), float(valid.max())), key=f"num_filter_{col}",
            )
            mask &= df[col].between(low, high) | df[col].isna()
        else:
            values = sorted(df[col].dropna().astype(str).unique().tolist())
            if len(values) > 300:
                st.caption(t("filter_skipped", col=col))
                continue
            selected = st.multiselect(t("values_of", col=col), values, default=values, key=f"cat_filter_{col}")
            mask &= df[col].astype(str).isin(selected) | df[col].isna()

filtered = df.loc[mask].copy()
if filtered.empty:
    st.warning(t("no_rows"))
    st.stop()

nnum, ncat, ndates = classify_columns(filtered)
qscore = quality_score(filtered)
missing_total = int(filtered.isna().sum().sum())
duplicate_current = int(filtered.duplicated().sum())

st.caption(t("records_line", a=f"{len(df):,}", b=f"{len(filtered):,}", c=f"{filtered.shape[1]:,}"))

tabs = st.tabs([
    t("tab_summary"), t("tab_charts"), t("tab_groups"),
    t("tab_trend"), t("tab_quality"), t("tab_export"),
])

# ------------------------------------------------------------
# 1) Executive summary
# ------------------------------------------------------------
with tabs[0]:
    st.subheader(t("overview"))
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(t("m_records"), f"{len(filtered):,}")
    k2.metric(t("m_cols"), f"{filtered.shape[1]:,}")
    k3.metric(t("m_complete"), f"{qscore['completeness']}%")
    k4.metric(t("m_missing"), f"{missing_total:,}")
    k5.metric(t("m_quality"), f"{qscore['score']}/100")
    st.caption(t("quality_note"))

    if nnum:
        st.markdown(f"#### {t('kpi_title')}")
        selected_kpi = st.selectbox(t("kpi_pick"), nnum, key="executive_kpi")
        s = pd.to_numeric(filtered[selected_kpi], errors="coerce").dropna()
        a, b, c, d, e = st.columns(5)
        a.metric(t("m_sum"), fmt(s.sum()) if not s.empty else "—")
        b.metric(t("m_mean"), fmt(s.mean()) if not s.empty else "—")
        c.metric(t("m_median"), fmt(s.median()) if not s.empty else "—")
        d.metric(t("m_max"), fmt(s.max()) if not s.empty else "—")
        e.metric(t("m_min"), fmt(s.min()) if not s.empty else "—")

        st.markdown(f"#### {t('period_title')}")
        if ndates:
            date_col = st.selectbox(t("date_pick"), ndates, key="period_date")
            valid_dates = filtered[date_col].dropna()
            if len(valid_dates) >= 2:
                max_date, min_date = valid_dates.max(), valid_dates.min()
                span_days = max((max_date - min_date).days + 1, 2)
                p_cols = st.columns(2)
                period_days = p_cols[0].number_input(
                    t("period_len"), min_value=1, max_value=min(3650, span_days),
                    value=max(1, min(30, span_days // 2)),
                )
                agg = p_cols[1].selectbox(
                    t("period_method"), ["sum", "mean", "median", "count"],
                    format_func=lambda k: t(f"agg_{k}"), key="period_agg",
                )
                cut = max_date - pd.Timedelta(days=int(period_days) - 1)
                previous_cut = cut - pd.Timedelta(days=1)
                previous_start = previous_cut - pd.Timedelta(days=int(period_days) - 1)
                cur_vals = pd.to_numeric(
                    filtered.loc[filtered[date_col].between(cut, max_date), selected_kpi], errors="coerce").dropna()
                prev_vals = pd.to_numeric(
                    filtered.loc[filtered[date_col].between(previous_start, previous_cut), selected_kpi], errors="coerce").dropna()
                current_value, previous_value = agg_series(cur_vals, agg), agg_series(prev_vals, agg)
                delta = np.nan
                if pd.notna(previous_value) and previous_value != 0 and pd.notna(current_value):
                    delta = (current_value - previous_value) / abs(previous_value) * 100
                p1, p2, p3 = st.columns(3)
                p1.metric(t("m_recent"), fmt(current_value),
                          help=t("range_help", a=cut.date(), b=max_date.date()))
                p2.metric(t("m_prev"), fmt(previous_value),
                          help=t("range_help", a=previous_start.date(), b=previous_cut.date()))
                p3.metric(t("m_change"), f"{delta:+.2f}%" if pd.notna(delta) else t("na_change"))
                if pd.isna(delta):
                    st.info(t("delta_info"))
        else:
            st.info(t("need_date"))

        # Automatic insight: which group drives the KPI
        candidates = [c for c in ncat if 2 <= filtered[c].nunique(dropna=True) <= 30]
        if candidates and not s.empty:
            g_col = candidates[0]
            grp = filtered.groupby(g_col)[selected_kpi].sum().sort_values(ascending=False)
            total = grp.sum()
            if len(grp) and total > 0 and grp.iloc[0] >= 0:
                st.markdown(f"#### {t('insights')}")
                st.success(t("insight_top", kpi=selected_kpi, grp=grp.index[0], col=g_col,
                             share=f"{grp.iloc[0] / total * 100:.1f}"))

    st.divider()
    st.markdown(f"#### {t('notes_title')}")
    if missing_total:
        st.warning(t("warn_missing", n=f"{missing_total:,}"))
    else:
        st.success(t("ok_missing"))
    if duplicate_current:
        st.warning(t("warn_dups", n=f"{duplicate_current:,}"))
    if nnum:
        outs = outlier_summary(filtered, nnum)
        if not outs.empty and outs[t("o_count")].max() > 0:
            worst = outs.iloc[0]
            st.info(t("info_outlier", col=worst[t("o_var")], n=f"{int(worst[t('o_count')]):,}"))
    else:
        st.info(t("no_numeric"))

    with st.expander(t("preview")):
        table(filtered.head(50), hide_index=True)

# ------------------------------------------------------------
# 2) Chart explorer
# ------------------------------------------------------------
with tabs[1]:
    st.subheader(t("chart_title"))
    chart_type = st.selectbox(
        t("analysis_type"), ["hist", "box", "scatter", "freq", "pareto"],
        format_func=lambda k: t(f"ct_{k}"),
    )
    if chart_type in ("hist", "box") and nnum:
        col = st.selectbox(t("num_var"), nnum, key="chart_num")
        valid = pd.to_numeric(filtered[col], errors="coerce").dropna()
        if valid.empty:
            st.info(t("no_valid_num"))
        elif chart_type == "hist":
            bins = st.slider(t("bins"), 5, 100, 30)
            fig = px.histogram(filtered, x=col, nbins=bins, marginal="box", title=t("hist_title", col=col))
            show(fig, 470)
            st.caption(t("stats_line", n=f"{len(valid):,}", sd=fmt(valid.std()), sk=fmt(valid.skew())))
        else:
            group = st.selectbox(t("group_opt"), [None] + ncat, format_func=lambda x: t("no_group") if x is None else x, key="box_group")
            fig = px.box(filtered, y=col, x=group, points="outliers", title=t("box_title", col=col))
            show(fig, 470)
    elif chart_type == "scatter" and len(nnum) >= 2:
        cx, cy = st.columns(2)
        x = cx.selectbox(t("x_axis"), nnum, index=0)
        y = cy.selectbox(t("y_axis"), nnum, index=1)
        color_by = st.selectbox(t("color_opt"), [None] + ncat, format_func=lambda v: t("no_color") if v is None else v)
        want_trend = st.checkbox(t("show_trend"), value=False)
        title = t("scatter_title", x=x, y=y)
        try:
            fig = px.scatter(filtered, x=x, y=y, color=color_by, trendline="ols" if want_trend else None, title=title)
        except Exception:  # noqa: BLE001 - statsmodels missing or fit failed
            fig = px.scatter(filtered, x=x, y=y, color=color_by, title=title)
            st.warning(t("trend_needs_sm"))
        show(fig, 480)
        corr_value = filtered[[x, y]].corr().iloc[0, 1]
        st.metric(t("pearson"), fmt(corr_value))
        st.caption(t("corr_note"))
    elif chart_type in ("freq", "pareto") and ncat:
        col = st.selectbox(t("cat_var"), ncat, key="chart_cat")
        top_n = st.slider(t("top_n"), 3, 30, 10)
        counts = (filtered[col].fillna(t("missing")).astype(str).value_counts().head(top_n)
                  .rename_axis(t("cat")).reset_index(name=t("count")))
        counts[t("cum_share")] = counts[t("count")].cumsum() / max(counts[t("count")].sum(), 1) * 100
        if chart_type == "pareto":
            fig = go.Figure()
            fig.add_bar(x=counts[t("cat")], y=counts[t("count")], name=t("count"))
            fig.add_scatter(x=counts[t("cat")], y=counts[t("cum_share")], name=t("cum_share"),
                            yaxis="y2", mode="lines+markers")
            fig.update_layout(
                title=t("pareto_title", col=col),
                yaxis=dict(title=t("count")),
                yaxis2=dict(title=t("cum_share"), overlaying="y", side="right", range=[0, 100]),
            )
            show(fig, 480)
        else:
            fig = px.bar(counts, x=t("count"), y=t("cat"), orientation="h", text=t("count"),
                         title=t("freq_title", col=col))
            fig.update_yaxes(categoryorder="total ascending")
            show(fig, 450)
        table(counts)
    else:
        st.info(t("chart_missing_cols"))

# ------------------------------------------------------------
# 3) Group comparison
# ------------------------------------------------------------
with tabs[2]:
    st.subheader(t("groups_title"))
    if nnum and ncat:
        c1, c2, c3 = st.columns(3)
        group_col = c1.selectbox(t("group_col"), ncat, key="group_col")
        value_col = c2.selectbox(t("value_col"), nnum, key="group_val")
        aggregation = c3.selectbox(t("aggregation"), ["sum", "mean", "median", "count", "max", "min"],
                                   format_func=lambda k: t(f"agg_{k}"), key="group_agg")
        result = filtered.groupby(group_col, dropna=False)[value_col].agg(aggregation).reset_index()
        result = result.sort_values(value_col, ascending=False)
        result[group_col] = result[group_col].astype(object).where(result[group_col].notna(), t("missing")).astype(str)
        fig = px.bar(result.head(30), x=value_col, y=group_col, orientation="h", text_auto=".3s",
                     title=t("group_bar_title", agg=t(f"agg_{aggregation}"), val=value_col, grp=group_col))
        fig.update_yaxes(categoryorder="total ascending")
        show(fig, 480)
        table(result, hide_index=True)

        if len(nnum) >= 2:
            st.markdown(f"#### {t('corr_matrix')}")
            corr_cols = st.multiselect(t("corr_vars"), nnum, default=nnum[:min(10, len(nnum))])
            corr_cols = corr_cols[:MAX_CORRELATION_COLUMNS]
            method = st.selectbox(t("corr_method"), ["pearson", "spearman", "kendall"])
            if len(corr_cols) >= 2:
                corr = filtered[corr_cols].corr(method=method)
                fig = px.imshow(corr, text_auto=".2f", zmin=-1, zmax=1, color_continuous_scale="RdBu_r",
                                title=t("corr_title", m=method))
                show(fig, 500)
            else:
                st.info(t("pick_two"))
    else:
        st.info(t("groups_need"))

# ------------------------------------------------------------
# 4) Time trend
# ------------------------------------------------------------
with tabs[3]:
    st.subheader(t("trend_title"))
    if ndates and nnum:
        dc, vc = st.columns(2)
        t_date = dc.selectbox(t("date_col"), ndates, key="trend_date")
        t_value = vc.selectbox(t("value_col"), nnum, key="trend_value")
        o1, o2, o3 = st.columns(3)
        frequency = o1.selectbox(t("freq"), ["D", "W", "M", "Q", "Y"], index=2,
                                 format_func=lambda k: t(f"f_{k}"))
        t_agg = o2.selectbox(t("agg_method"), ["sum", "mean", "median", "count"],
                             format_func=lambda k: t(f"agg_{k}"), key="trend_agg")
        window = o3.slider(t("window"), 1, 24, 3)
        freq_map = {"D": "D", "W": "W", "M": "MS", "Q": "QS", "Y": "YS"}
        ts_data = filtered[[t_date, t_value]].dropna().copy()
        if not ts_data.empty:
            ts_data[t_date] = pd.to_datetime(ts_data[t_date], errors="coerce")
            ts_data = ts_data.dropna(subset=[t_date]).set_index(t_date).sort_index()
            series = ts_data[t_value].resample(freq_map[frequency]).agg(t_agg)
            if len(series) >= 1:
                trend = series.to_frame(t("value"))
                trend[t("moving_avg")] = trend[t("value")].rolling(window, min_periods=1).mean()
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=trend.index, y=trend[t("value")], mode="lines+markers", name=t("value")))
                fig.add_trace(go.Scatter(x=trend.index, y=trend[t("moving_avg")], mode="lines",
                                         name=t("moving_avg"), line=dict(dash="dash", width=3)))
                fig.update_layout(title=t("trend_fig_title", agg=t(f"agg_{t_agg}"), val=t_value),
                                  xaxis_title=t("date"), yaxis_title=t_value)
                show(fig, 480)
                table(trend.reset_index().rename(columns={trend.reset_index().columns[0]: t("date")}))
                st.caption(t("trend_note"))
            else:
                st.info(t("trend_few"))
        else:
            st.info(t("trend_invalid"))
    else:
        st.info(t("trend_need"))

# ------------------------------------------------------------
# 5) Data quality
# ------------------------------------------------------------
with tabs[4]:
    st.subheader(t("quality_title"))
    quality = quality_report(filtered)
    table(quality, hide_index=True)
    miss_n, miss_p, q_col = t("q_missing_n"), t("q_missing_pct"), t("q_col")
    missing = quality[quality[miss_n] > 0].sort_values(miss_p, ascending=False)
    if not missing.empty:
        left, right = st.columns([3, 2])
        with left:
            fig = px.bar(missing, x=miss_p, y=q_col, orientation="h", text=miss_p, title=t("missing_chart"))
            fig.update_yaxes(categoryorder="total ascending")
            show(fig, 430)
        with right:
            st.markdown(f"**{t('missing_cols')}**")
            table(missing[[q_col, miss_n, miss_p]], hide_index=True)
    else:
        st.success(t("no_missing"))

    if nnum:
        st.markdown(f"#### {t('outliers_title')}")
        st.caption(t("outliers_note"))
        table(outlier_summary(filtered, nnum), hide_index=True)

    st.markdown(f"#### {t('prep_title')}")
    st.write(f"{t('prep_rows')}: **{prep['original_rows']:,}**")
    st.write(f"{t('prep_cols')}: **{prep['original_cols']:,}**")
    st.write(f"{t('prep_num')}: **{len(prep['numeric_converted'])}**")
    st.write(f"{t('prep_dates')}: **{len(prep['dates_converted'])}**")
    st.write(f"{t('prep_dups_found')}: **{prep['duplicates_found']:,}**")
    st.write(f"{t('prep_dups_removed')}: **{prep['duplicates_removed']:,}**")

# ------------------------------------------------------------
# 6) Export
# ------------------------------------------------------------
with tabs[5]:
    st.subheader(t("export_title"))
    st.write(t("export_note1"))
    st.caption(t("export_note2"))
    base_name = safe_filename(re.sub(r"\.[^.]+$", "", source))
    csv_bytes = filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        t("dl_csv"), data=csv_bytes, file_name=f"{base_name}_filtered.csv",
        mime="text/csv", **STRETCH,
    )
    try:
        kpi_for_report = st.session_state.get("executive_kpi")
        excel_bytes = make_excel_report(filtered, source, prep, kpi_for_report)
        st.download_button(
            t("dl_xlsx"), data=excel_bytes, file_name=f"{base_name}_management_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **STRETCH,
        )
    except Exception as exc:  # noqa: BLE001
        st.error(t("xlsx_fail", err=exc))

    st.divider()
    st.markdown(f"#### {t('limits_title')}")
    st.write(f"- {t('limit1')}")
    st.write(f"- {t('limit2')}")
    st.write(f"- {t('limit3')}")

st.markdown(f'<div class="footer">{t("footer")}</div>', unsafe_allow_html=True)