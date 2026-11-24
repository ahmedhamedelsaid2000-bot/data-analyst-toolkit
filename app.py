"""
Data Analyst Toolkit - تطبيق تحليل بيانات (Streamlit)
يغطي: تنظيف البيانات، KPIs، داشبورد، SQL، تصدير Excel

التشغيل:
    pip install streamlit pandas openpyxl
    streamlit run app.py
"""
import io
import sqlite3

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Data Analyst Toolkit", page_icon="📊", layout="wide")


# ---------- دوال التحليل ----------
@st.cache_data
def load_file(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)


def sample_data() -> pd.DataFrame:
    rng = np.random.default_rng(42)
    n = 500
    df = pd.DataFrame(
        {
            "date": pd.date_range("2025-01-01", periods=n, freq="D").astype(str),
            "region": rng.choice(["Cairo", "Kuwait", "Riyadh", "Dubai"], n),
            "product": rng.choice(["A", "B", "C", "D"], n),
            "quantity": rng.integers(1, 50, n).astype(float),
            "price": rng.uniform(10, 200, n).round(2),
        }
    )
    df["sales"] = (df["quantity"] * df["price"]).round(2)
    # نضيف بيانات وسخة عشان نجرب التنظيف
    df.loc[rng.choice(n, 20, replace=False), "quantity"] = np.nan
    df = pd.concat([df, df.sample(15, random_state=1)], ignore_index=True)
    return df


def clean_data(df, drop_dupes, fill_method, parse_dates):
    report = {}
    out = df.copy()
    out.columns = [c.strip().lower().replace(" ", "_") for c in out.columns]

    if drop_dupes:
        before = len(out)
        out = out.drop_duplicates()
        report["صفوف مكررة اتحذفت"] = before - len(out)

    report["قيم ناقصة قبل التنظيف"] = int(out.isna().sum().sum())

    num_cols = out.select_dtypes(include="number").columns
    obj_cols = out.select_dtypes(exclude="number").columns
    if fill_method == "المتوسط":
        out[num_cols] = out[num_cols].fillna(out[num_cols].mean())
    elif fill_method == "الوسيط":
        out[num_cols] = out[num_cols].fillna(out[num_cols].median())
    elif fill_method == "صفر":
        out[num_cols] = out[num_cols].fillna(0)
    elif fill_method == "حذف الصفوف":
        out = out.dropna()
    out[obj_cols] = out[obj_cols].fillna("غير معروف") if fill_method != "حذف الصفوف" else out[obj_cols]

    if parse_dates:
        for c in out.columns:
            if "date" in c or "تاريخ" in c:
                out[c] = pd.to_datetime(out[c], errors="coerce")

    report["قيم ناقصة بعد التنظيف"] = int(out.isna().sum().sum())
    return out, report


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cleaned Data")
        df.describe(include="all").to_excel(writer, sheet_name="Summary")
    return buf.getvalue()


def run_sql(df: pd.DataFrame, query: str) -> pd.DataFrame:
    conn = sqlite3.connect(":memory:")
    df.to_sql("data", conn, index=False, if_exists="replace")
    try:
        return pd.read_sql_query(query, conn)
    finally:
        conn.close()


# ---------- الواجهة ----------
st.title("📊 Data Analyst Toolkit")
st.caption("ارفع ملفك → نظّفه → حلّله → اعمل داشبورد → صدّر النتيجة")

with st.sidebar:
    st.header("1) البيانات")
    uploaded = st.file_uploader("ارفع CSV أو Excel", type=["csv", "xlsx", "xls"])
    use_sample = st.checkbox("استخدم بيانات تجريبية", value=uploaded is None)

    st.header("2) إعدادات التنظيف")
    drop_dupes = st.checkbox("حذف الصفوف المكررة", value=True)
    fill_method = st.selectbox("التعامل مع القيم الناقصة", ["المتوسط", "الوسيط", "صفر", "حذف الصفوف", "بدون"])
    parse_dates = st.checkbox("تحويل أعمدة التاريخ تلقائيًا", value=True)

if uploaded is not None:
    raw = load_file(uploaded)
elif use_sample:
    raw = sample_data()
else:
    st.info("ارفع ملف أو فعّل البيانات التجريبية من الشريط الجانبي.")
    st.stop()

df, report = clean_data(raw, drop_dupes, fill_method, parse_dates)

tab1, tab2, tab3, tab4 = st.tabs(["🧹 التنظيف", "📈 الداشبورد", "🗄️ SQL", "💾 التصدير"])

# --- التنظيف ---
with tab1:
    c1, c2 = st.columns(2)
    c1.metric("الصفوف قبل", len(raw))
    c2.metric("الصفوف بعد", len(df))
    st.subheader("تقرير التنظيف")
    st.json(report)
    st.subheader("معاينة البيانات")
    st.dataframe(df.head(50), use_container_width=True)

# --- الداشبورد ---
with tab2:
    num_cols = df.select_dtypes(include="number").columns.tolist()
    cat_cols = df.select_dtypes(exclude="number").columns.tolist()
    date_cols = df.select_dtypes(include="datetime").columns.tolist()

    if not num_cols:
        st.warning("مفيش أعمدة رقمية للتحليل.")
    else:
        kpi_col = st.selectbox("العمود الرقمي (KPI)", num_cols)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("المجموع", f"{df[kpi_col].sum():,.2f}")
        k2.metric("المتوسط", f"{df[kpi_col].mean():,.2f}")
        k3.metric("الأعلى", f"{df[kpi_col].max():,.2f}")
        k4.metric("الأقل", f"{df[kpi_col].min():,.2f}")

        left, right = st.columns(2)
        group_cols = [c for c in cat_cols if df[c].nunique() <= 50]
        if group_cols:
            grp = left.selectbox("تجميع حسب", group_cols)
            agg = left.radio("الدالة", ["sum", "mean", "count"], horizontal=True)
            grouped = df.groupby(grp)[kpi_col].agg(agg).sort_values(ascending=False)
            left.bar_chart(grouped)
        if date_cols:
            dcol = date_cols[0]
            trend = df.groupby(df[dcol].dt.to_period("M").astype(str))[kpi_col].sum()
            right.line_chart(trend)
            right.caption(f"الاتجاه الشهري لـ {kpi_col}")

        st.subheader("Pivot Table")
        if len(group_cols) >= 2:
            p_rows = st.selectbox("الصفوف", group_cols, index=0, key="pr")
            p_cols = st.selectbox("الأعمدة", group_cols, index=1, key="pc")
            pivot = pd.pivot_table(df, values=kpi_col, index=p_rows, columns=p_cols, aggfunc="sum", fill_value=0)
            st.dataframe(pivot, use_container_width=True)

# --- SQL ---
with tab3:
    st.write("اسم الجدول: `data`")
    st.code(", ".join(df.columns))
    default_q = "SELECT * FROM data LIMIT 10"
    query = st.text_area("اكتب استعلام SQL", default_q, height=120)
    if st.button("تشغيل"):
        try:
            st.dataframe(run_sql(df, query), use_container_width=True)
        except Exception as e:
            st.error(f"خطأ في الاستعلام: {e}")

# --- التصدير ---
with tab4:
    st.download_button(
        "⬇️ تحميل Excel (بيانات + ملخص)",
        data=to_excel_bytes(df),
        file_name="cleaned_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        "⬇️ تحميل CSV",
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name="cleaned_data.csv",
        mime="text/csv",
  )
