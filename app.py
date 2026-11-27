"""
app.py - ระบบวิเคราะห์ยอดขายและพฤติกรรมลูกค้า (SmartRetail Thailand)
พัฒนาด้วย Python + Streamlit + Altair
ต้องติดตั้ง: streamlit pandas altair openpyxl
"""
import hashlib
import hmac
import io
import re
import time
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

CSV_FILE = Path(__file__).with_name("cleaned_transactions.csv")
CUST_FILE = Path(__file__).with_name("customers_data.csv")
PAGES = ["ภาพรวมยอดขาย", "รายละเอียดสินค้า", "พฤติกรรมลูกค้า", "สำรวจข้อมูล (EDA)"]
PAY_TH = {"Cash": "เงินสด", "Credit Card": "บัตรเครดิต", "PromptPay": "พร้อมเพย์"}
GENDER_TH = {"Male": "ชาย", "Female": "หญิง"}

MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]
SHORT = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
CH_TH = {"In-Store": "หน้าร้าน", "Online": "ออนไลน์"}
CAT_TH = {
    "Dry Food": "อาหารแห้ง",
    "Beverages": "เครื่องดื่ม",
    "Personal Care": "ของใช้ส่วนตัว",
    "Household": "ของใช้ในบ้าน",
    "Snacks": "ขนมขบเคี้ยว",
}
REG_TH = {
    "Central": "ภาคกลาง",
    "North": "ภาคเหนือ",
    "South": "ภาคใต้",
    "Northeast": "ภาคตะวันออกเฉียงเหนือ",
    "West": "ภาคตะวันตก",
}
ALL_CH, ALL_CAT, ALL_REG = "ทุกช่องทาง", "ทุกหมวดสินค้า", "ทุกภูมิภาค"

USERS = {
    "judge": {
        "name": "คณะกรรมการตัดสิน (Judge)",
        "hash": hashlib.sha256("1234".encode()).hexdigest(),
    },
    "admin": {
        "name": "ผู้ดูแลระบบ (SmartRetail Analyst)",
        "hash": hashlib.sha256("smartretail2025".encode()).hexdigest(),
    },
}

st.set_page_config(page_title="SmartRetail – ระบบวิเคราะห์ข้อมูล", page_icon="📊", layout="wide")


# ---- รองรับ Streamlit ทั้งเวอร์ชันเก่าและใหม่ ----
# เวอร์ชันใหม่ (>=1.50) ยกเลิก use_container_width แล้วใช้ width="stretch" แทน
def _st_version():
    try:
        return tuple(int(x) for x in st.__version__.split(".")[:2])
    except Exception:
        return (0, 0)


W = {"width": "stretch"} if _st_version() >= (1, 50) else {"use_container_width": True}


def _ensure_light_theme():
    try:
        cfg = Path.cwd() / ".streamlit" / "config.toml"
        if not cfg.exists():
            cfg.parent.mkdir(exist_ok=True)
            cfg.write_text('[theme]\nbase = "light"\nprimaryColor = "#1d4ed8"\n', encoding="utf-8")
    except Exception:
        pass


_ensure_light_theme()

CSS = """
<style>
html, body, .stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp button,
.stApp [data-baseweb="select"] div, .stApp [data-testid="stMarkdownContainer"] {
  font-family: "Sarabun","Leelawadee UI","Tahoma",sans-serif;
}
.stApp { background: linear-gradient(180deg,#e9f0fb 0%,#f7f9fc 45%,#e9f0fb 100%); color:#1b2433; }
header[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer, .stAppDeployButton, [data-testid="stDeployButton"] { visibility: hidden; display:none; }
.block-container { max-width: 1280px; padding-top: 1.4rem; padding-bottom: 3rem; }
@keyframes rise { from {opacity:0; transform: translateY(10px);} to {opacity:1; transform:none;} }

.hero { position:relative; overflow:hidden; border-radius:22px; padding:26px 32px; color:#fff;
  background: linear-gradient(120deg,#0b2e6b 0%,#1d4ed8 55%,#0ea5a4 100%);
  box-shadow: 0 22px 44px rgba(29,78,216,.30), 0 4px 10px rgba(11,46,107,.25); animation: rise .5s ease both; }
.hero::before { content:""; position:absolute; right:-60px; top:-80px; width:300px; height:300px; border-radius:50%;
  background: radial-gradient(circle at 30% 30%, rgba(255,255,255,.35), rgba(255,255,255,0) 70%); }
.hero::after { content:""; position:absolute; right:140px; bottom:-110px; width:240px; height:240px; border-radius:50%;
  background: rgba(255,255,255,.08); }
.hero h1 { color:#fff !important; font-size:2rem; font-weight:800; margin:0 0 4px 0; padding:0; text-shadow:0 2px 8px rgba(0,0,0,.25); }
.hero p { color:#dbeafe !important; margin:0 0 12px 0; font-size:1.05rem; }
.chip { display:inline-block; margin:0 8px 0 0; padding:5px 14px; border-radius:999px; font-size:.88rem; font-weight:600;
  background: rgba(255,255,255,.18); border:1px solid rgba(255,255,255,.35); backdrop-filter: blur(4px); }

.st-key-page [role="radiogroup"] { display:inline-flex; flex-wrap:wrap; gap:6px; background:#fff; padding:6px; border-radius:18px;
  box-shadow: 0 8px 22px rgba(15,61,122,.12); margin: 6px 0 4px 0; }
.st-key-page label { margin:0; padding:9px 22px; border-radius:13px; cursor:pointer; transition: all .2s; }
.st-key-page label > div:first-of-type > div:first-child, .st-key-page label > div:first-child > div:first-child { display:none; }
.st-key-page label:hover { background:#eaf1ff; }
.st-key-page label[data-selected="true"], .st-key-page label:has(input:checked) {
  background: linear-gradient(135deg,#1d4ed8,#0ea5a4); box-shadow: 0 8px 18px rgba(29,78,216,.35); }
.st-key-page label[data-selected="true"] *, .st-key-page label:has(input:checked) * { color:#fff !important; font-weight:700; }
.st-key-page label p { font-size:1rem; margin:0; }

.st-key-filters { background:#fff; border:none !important; border-radius:18px; padding:8px 10px;
  box-shadow: 0 10px 26px rgba(15,61,122,.10), 0 2px 6px rgba(15,61,122,.06); animation: rise .5s .05s ease both; }
.stApp [data-baseweb="select"] > div { border-radius:12px; border-color:#d5deec; background:#f8fafd; }
.stApp [data-testid="stWidgetLabel"] p { font-weight:700; color:#0f3d7a; }

.stApp .stButton > button, .stApp .stDownloadButton > button, .stApp [data-testid="stFormSubmitButton"] > button { border:none; border-radius:12px; font-weight:700; color:#fff;
  background: linear-gradient(135deg,#1d4ed8,#2563eb); box-shadow: 0 8px 18px rgba(29,78,216,.30); transition: all .2s; }
.stApp .stButton > button:hover, .stApp .stDownloadButton > button:hover { transform: translateY(-2px);
  box-shadow: 0 14px 26px rgba(29,78,216,.40); color:#fff; }
.stApp .stDownloadButton > button { background: linear-gradient(135deg,#0f766e,#14b8a6); box-shadow: 0 8px 18px rgba(15,118,110,.30); }

.kpi { position:relative; overflow:hidden; background:#fff; border-radius:20px; padding:18px 20px 16px 24px; margin:6px 0 12px 0;
  border:1px solid #e6ecf5; box-shadow: 0 12px 28px rgba(15,61,122,.10), 0 2px 6px rgba(15,61,122,.06);
  transition: transform .25s, box-shadow .25s; animation: rise .55s ease both; }
.kpi::before { content:""; position:absolute; left:0; top:0; bottom:0; width:7px; background: var(--c); }
.kpi::after { content:""; position:absolute; right:-30px; top:-30px; width:110px; height:110px; border-radius:50%; background: var(--c2); opacity:.7; }
.kpi:hover { transform: translateY(-6px); box-shadow: 0 22px 40px rgba(15,61,122,.20); }
.kpi .ic { position:relative; z-index:1; width:46px; height:46px; border-radius:14px; display:flex; align-items:center; justify-content:center;
  font-size:23px; background: var(--c2); box-shadow: inset 0 -3px 0 rgba(0,0,0,.06); margin-bottom:10px; }
.kpi .l { position:relative; z-index:1; color:#64748b; font-weight:600; font-size:.95rem; }
.kpi .v { position:relative; z-index:1; color:#0f172a; font-weight:800; font-size:clamp(1.35rem,2.1vw,1.9rem); line-height:1.25; }
.kpi .u { position:relative; z-index:1; color:var(--c); font-weight:700; font-size:.85rem; }

.insight-card { background:#fff; border:1px solid #dbeafe; border-left:6px solid #1d4ed8; border-radius:14px;
  padding:16px 20px; margin:4px 0 18px 0; box-shadow:0 6px 18px rgba(15,61,122,.07); font-size:.98rem; line-height:1.65; }
.insight-card strong { color:#0f3d7a; }

.stApp h3 { color:#0f3d7a; font-weight:800; font-size:1.35rem; border-left:6px solid #1d4ed8; padding-left:12px; margin-top:22px; }
.stElementContainer:has(.vega-embed) { background:#fff; border-radius:20px; padding:16px 14px 10px 14px; margin:6px 0 8px 0;
  border:1px solid #e6ecf5; box-shadow: 0 12px 28px rgba(15,61,122,.10), 0 2px 6px rgba(15,61,122,.06); animation: rise .6s ease both; }
.stApp [data-testid="stDataFrame"] { border-radius:16px; overflow:hidden; box-shadow: 0 8px 22px rgba(15,61,122,.10); }
.stApp [data-testid="stExpander"] { background:#fff; border-radius:16px; border:1px solid #e6ecf5; box-shadow: 0 8px 22px rgba(15,61,122,.08); }
.stApp [data-testid="stAlert"] { border-radius:14px; box-shadow: 0 4px 12px rgba(15,61,122,.08); }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

KPI_COLORS = [("#2563eb", "#dbeafe"), ("#f97316", "#ffedd5"), ("#10b981", "#d1fae5"), ("#8b5cf6", "#ede9fe")]
# สหสัมพันธ์ต้องอยู่บนสุด ไม่งั้นจะถูกจับคู่กับคำว่า "ยอดขาย" ก่อน
KPI_ICONS = [
    ("สหสัมพันธ์", "🔗"), ("IQR", "📐"), ("เฉลี่ย", "📈"), ("ยอดขายของหมวด", "💰"), ("ยอดขาย", "💰"),
    ("ยอดซื้อ", "💳"), ("หมวด", "🏆"), ("รายการ", "🧾"), ("ชิ้น", "📦"), ("ซื้อซ้ำ", "🔁"),
    ("ลูกค้า", "👥"), ("ซื้อ", "🛍️"),
]
_kpi_n = [0]


def kpi(col, label, value):
    m = re.match(r"^(.*?)\s*\((.*)\)$", label)
    title, unit = (m.group(1), m.group(2)) if m else (label, "")
    icon = next((ic for k, ic in KPI_ICONS if k in title), "📊")
    c, c2 = KPI_COLORS[_kpi_n[0] % 4]
    _kpi_n[0] += 1
    col.markdown(
        f'<div class="kpi" style="--c:{c};--c2:{c2}"><div class="ic">{icon}</div>'
        f'<div class="l">{title}</div><div class="v">{value}</div><div class="u">{unit or "&nbsp;"}</div></div>',
        unsafe_allow_html=True,
    )


def insight_box(what_text: str, so_what_text: str, now_what_text: str):
    st.markdown(
        '<div class="insight-card">'
        '<div style="font-weight:800; color:#1d4ed8; margin-bottom:6px; font-size:1.05rem;">'
        '💡 สรุปการวิเคราะห์ (What → So What → Now What)</div>'
        f'<div>• <strong>What (เกิดอะไรขึ้น):</strong> {what_text}</div>'
        f'<div>• <strong>So What (สำคัญอย่างไร):</strong> {so_what_text}</div>'
        f'<div>• <strong>Now What (ควรทำอะไรต่อ):</strong> {now_what_text}</div>'
        '</div>',
        unsafe_allow_html=True,
    )


_orig_altair_chart = st.altair_chart


def _pretty_altair_chart(chart, **kw):
    chart = (
        chart.configure_view(strokeWidth=0)
        .configure_axis(
            gridColor="#e8edf5", domainColor="#cbd5e1", tickColor="#cbd5e1",
            labelColor="#334155", titleColor="#334155", labelFontSize=13,
            titleFontSize=13, labelFont="Sarabun, Leelawadee UI, Tahoma",
        )
        .configure_legend(
            labelFontSize=13, titleFontSize=13, orient="top",
            labelColor="#334155", titleColor="#334155",
        )
    )
    return _orig_altair_chart(chart, **kw)


st.altair_chart = _pretty_altair_chart


# ======================= ล็อกอิน =======================
def verify_login(username: str, password: str) -> bool:
    user = USERS.get(username.strip().lower())
    if not user:
        return False
    pw_hash = hashlib.sha256(password.encode()).hexdigest()
    return hmac.compare_digest(user["hash"], pw_hash)


if "auth_user" not in st.session_state:
    st.session_state.auth_user = None

if st.session_state.auth_user is None:
    st.markdown(
        '<div class="hero"><h1>🛒 SmartRetail Thailand</h1>'
        '<p>ระบบวิเคราะห์พฤติกรรมลูกค้าและประสิทธิภาพยอดขาย (ปวช. ระดับจังหวัด)</p>'
        '<span class="chip">🔒 ระบบยืนยันตัวตนก่อนเข้าใช้งาน</span></div>',
        unsafe_allow_html=True,
    )
    st.write("")
    _, login_col, _ = st.columns([1, 1.3, 1])
    with login_col:
        with st.container(border=True, key="filters"):
            st.subheader("🔑 เข้าสู่ระบบ (Login)")
            st.caption("กรุณากรอกชื่อผู้ใช้และรหัสผ่านสำหรับเข้าใช้งานแดชบอร์ด")
            with st.form("login_form"):
                u_in = st.text_input("ชื่อผู้ใช้ (Username)", placeholder="เช่น judge หรือ admin")
                p_in = st.text_input("รหัสผ่าน (Password)", type="password", placeholder="กรอกรหัสผ่าน")
                submitted = st.form_submit_button("เข้าสู่ระบบ", **W)
                if submitted:
                    if verify_login(u_in, p_in):
                        st.session_state.auth_user = u_in.strip().lower()
                        st.rerun()
                    else:
                        st.error("❌ ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")
            st.info(
                "📌 **บัญชีสำหรับทดสอบระบบ (คณะกรรมการ):**\n"
                "- ชื่อผู้ใช้: `judge` | รหัสผ่าน: `1234`\n"
                "- ชื่อผู้ใช้: `admin` | รหัสผ่าน: `smartretail2025`"
            )
    st.stop()


# ======================= โหลดข้อมูล =======================
@st.cache_resource(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["Transaction_Date"] = pd.to_datetime(df["Transaction_Date"], errors="coerce")
    df = df.dropna(subset=["Transaction_Date", "Total_Amount"])
    df["month"] = df["Transaction_Date"].dt.month - 1
    df["ช่องทาง"] = df["Channel"].map(CH_TH).fillna(df["Channel"])
    df["หมวด"] = df["Category"].map(CAT_TH).fillna(df["Category"])
    df["ภูมิภาค"] = df["Region"].map(REG_TH).fillna(df["Region"])
    df["วิธีชำระเงิน"] = df["Payment_Method"].map(PAY_TH).fillna(df["Payment_Method"])
    return df


@st.cache_resource(show_spinner=False)
def load_customers(path: str):
    try:
        c = pd.read_csv(path)
        if "Customer_ID" not in c.columns:
            return None
        return c.drop_duplicates("Customer_ID")
    except Exception:
        return None


def apply_filters(key):
    m_from, m_to, ch, cat, reg = key
    base = load_data(str(CSV_FILE))
    x = base[(base["month"] >= m_from) & (base["month"] <= m_to)]
    if ch != ALL_CH:
        x = x[x["ช่องทาง"] == ch]
    if cat != ALL_CAT:
        x = x[x["หมวด"] == cat]
    if reg != ALL_REG:
        x = x[x["ภูมิภาค"] == reg]
    return x


@st.cache_data(show_spinner=False, max_entries=8)
def make_csv(key):
    x = apply_filters(key)
    return x[[c for c in x.columns if c != "month"]].to_csv(index=False).encode("utf-8-sig")


def to_excel_bytes(df_filtered: pd.DataFrame):
    """สร้างไฟล์ Excel (แคชไว้ ไม่สร้างใหม่ทุกครั้งที่กดเมนู) คืน None ถ้าไม่มี openpyxl"""
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        return None
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_m = (
            df_filtered.groupby(["month", "ช่องทาง"], as_index=False)
            .agg(ยอดขายรวม=("Total_Amount", "sum"),
                 **{"จำนวนรายการ": ("Transaction_ID", "count"),
                    "จำนวนลูกค้าไม่นับซ้ำ": ("Customer_ID", "nunique")})
            .sort_values("month")
        )
        summary_m["เดือน"] = summary_m["month"].map(lambda i: MONTHS[i] + " 2025")
        summary_m = summary_m[["เดือน", "ช่องทาง", "ยอดขายรวม", "จำนวนรายการ", "จำนวนลูกค้าไม่นับซ้ำ"]]
        summary_m.to_excel(writer, sheet_name="สรุปรายเดือน", index=False)

        summary_cat = (
            df_filtered.groupby("หมวด", as_index=False)
            .agg(ยอดขายรวม=("Total_Amount", "sum"),
                 **{"จำนวนชิ้น": ("Units_Sold", "sum"),
                    "จำนวนรายการ": ("Transaction_ID", "count")})
            .sort_values("ยอดขายรวม", ascending=False)
        )
        summary_cat.to_excel(writer, sheet_name="สรุปตามหมวดสินค้า", index=False)

        export_cols = [c for c in df_filtered.columns if c != "month"]
        df_filtered[export_cols].to_excel(writer, sheet_name="รายการขายที่กรอง", index=False)
    return output.getvalue()


def reset_filters():
    st.session_state.update(m_from=MONTHS[0], m_to=MONTHS[-1], f_ch=ALL_CH, f_cat=ALL_CAT, f_reg=ALL_REG)


def fmt0(x):
    return f"{x:,.0f}"


def fmt2(x):
    return f"{x:,.2f}"


t_start = time.perf_counter()
user_info = USERS[st.session_state.auth_user]

top_left, top_right = st.columns([5, 1])
with top_left:
    st.markdown(
        f'<div class="hero"><h1>🛒 SmartRetail Thailand</h1>'
        f'<p>ระบบวิเคราะห์ยอดขายและพฤติกรรมลูกค้า — ผู้ใช้งาน: {user_info["name"]}</p>'
        f'<span class="chip">📅 ข้อมูลปี 2025</span><span class="chip">🏬 หน้าร้าน + 🌐 ออนไลน์</span>'
        f'<span class="chip">📍 5 ภูมิภาค</span></div>',
        unsafe_allow_html=True,
    )
with top_right:
    st.write("")
    st.write("")
    if st.button("🚪 ออกจากระบบ", **W):
        st.session_state.auth_user = None
        st.rerun()

st.radio("เมนู", PAGES, horizontal=True, key="page", label_visibility="collapsed")

if not CSV_FILE.exists():
    st.error(f"❌ ไม่พบไฟล์ {CSV_FILE.name} กรุณาวางไว้โฟลเดอร์เดียวกับ app.py")
    st.stop()
try:
    with st.spinner("⏳ กำลังโหลดข้อมูล..."):
        df = load_data(str(CSV_FILE))
except Exception as e:
    st.error(f"❌ โหลดข้อมูลไม่สำเร็จ: {e}")
    st.stop()

if "m_from" not in st.session_state:
    reset_filters()
with st.container(border=True, key="filters"):
    c1, c2, c3, c4, c5, c6 = st.columns([1.2, 1.2, 1, 1.1, 1.3, 0.9])
    c1.selectbox("ตั้งแต่เดือน", MONTHS, key="m_from")
    c2.selectbox("ถึงเดือน", MONTHS, key="m_to")
    c3.selectbox("ช่องทางขาย", [ALL_CH] + list(CH_TH.values()), key="f_ch")
    c4.selectbox("หมวดสินค้า", [ALL_CAT] + list(CAT_TH.values()), key="f_cat")
    c5.selectbox("ภูมิภาค", [ALL_REG] + list(REG_TH.values()), key="f_reg")
    c6.write("")
    c6.button("🔄 ล้างตัวกรอง", on_click=reset_filters, **W)

m_from, m_to = MONTHS.index(st.session_state.m_from), MONTHS.index(st.session_state.m_to)
if m_from > m_to:
    st.error("⚠️ เดือนเริ่มต้นต้องไม่อยู่หลังเดือนสิ้นสุด กรุณาเลือกช่วงเดือนใหม่")
    st.stop()

fkey = (m_from, m_to, st.session_state.f_ch, st.session_state.f_cat, st.session_state.f_reg)
d = apply_filters(fkey)

if d.empty:
    st.warning('ℹ️ ไม่พบข้อมูลตามตัวกรองที่เลือก ลองกด "ล้างตัวกรอง"')
    st.stop()


# ======================= หน้าต่างๆ =======================
def page_overview(d, m_from, m_to):
    total, count = d["Total_Amount"].sum(), len(d)
    k1, k2, k3, k4 = st.columns(4)
    kpi(k1, "ยอดขายรวม (บาท)", fmt2(total))
    kpi(k2, "จำนวนรายการขาย (รายการ)", fmt0(count))
    kpi(k3, "ยอดเฉลี่ยต่อรายการ (บาท)", fmt2(total / count))
    kpi(k4, "จำนวนลูกค้าที่ซื้อ ไม่นับซ้ำ (คน)", fmt0(d["Customer_ID"].nunique()))

    st.subheader("ยอดขายรายเดือน แยกตามช่องทาง")
    monthly = d.groupby(["month", "ช่องทาง"], as_index=False).agg(
        ยอดขาย=("Total_Amount", "sum"), รายการ=("Transaction_ID", "count")
    )
    monthly["เดือน"] = monthly["month"].map(lambda i: SHORT[i])
    chart = (
        alt.Chart(monthly)
        .mark_bar(cornerRadiusTopLeft=7, cornerRadiusTopRight=7)
        .encode(
            x=alt.X("เดือน:N", sort=SHORT[m_from: m_to + 1], title="เดือน (ปี 2025)", axis=alt.Axis(labelAngle=0)),
            xOffset=alt.XOffset("ช่องทาง:N"),
            y=alt.Y("ยอดขาย:Q", title="ยอดขาย (บาท)", axis=alt.Axis(format=",.0f")),
            color=alt.Color(
                "ช่องทาง:N", title="ช่องทาง",
                scale=alt.Scale(domain=["หน้าร้าน", "ออนไลน์"], range=["#2563eb", "#f97316"]),
            ),
            tooltip=["เดือน", "ช่องทาง", alt.Tooltip("ยอดขาย:Q", format=",.2f"), alt.Tooltip("รายการ:Q", format=",")],
        )
        .properties(height=380)
    )
    st.altair_chart(chart, **W)

    peak_row = monthly.loc[monthly["ยอดขาย"].idxmax()]
    low_row = monthly.loc[monthly["ยอดขาย"].idxmin()]
    ch_share = d.groupby("ช่องทาง")["Total_Amount"].sum()
    top_ch = ch_share.idxmax()
    top_ch_pct = ch_share.max() / total * 100
    insight_box(
        what_text=(
            f"ช่วงเวลาที่เลือกมียอดขายรวม {fmt2(total)} บาท ช่องทาง <b>{top_ch}</b> สัดส่วนสูงสุด {top_ch_pct:.1f}% "
            f"เดือนที่ยอดสูงสุดคือ <b>{MONTHS[int(peak_row['month'])]} ({peak_row['ช่องทาง']})</b> {fmt2(peak_row['ยอดขาย'])} บาท "
            f"และต่ำสุดคือ {MONTHS[int(low_row['month'])]} ({low_row['ช่องทาง']}) {fmt2(low_row['ยอดขาย'])} บาท"
        ),
        so_what_text=(
            f"ยอดขายสูงสุดต่างจากต่ำสุด {peak_row['ยอดขาย'] / low_row['ยอดขาย']:.1f} เท่า "
            "แสดงว่ารายได้ผันผวนตามเดือนและช่องทาง จึงควรวางแผนสต็อกและงบการตลาดให้สอดคล้องกัน"
        ),
        now_what_text=(
            f"รักษาฐานยอดขายช่องทาง {top_ch} ด้วยโปรแกรมสะสมแต้ม และจัดโปรโมชันกระตุ้นยอดในเดือนที่ต่ำ เช่น "
            f"{MONTHS[int(low_row['month'])]}"
        ),
    )

    with st.expander("📋 ดูตารางตัวเลขรายเดือน"):
        tbl = monthly.pivot(index="month", columns="ช่องทาง", values="ยอดขาย").fillna(0)
        tbl.index = [MONTHS[i] + " 2025" for i in tbl.index]
        tbl["รวม"] = tbl.sum(axis=1)
        st.dataframe(tbl.style.format("{:,.2f}"), **W)


def bar_h(data, cat, val, title, fmt=",.0f", color="#1d4ed8"):
    return (
        alt.Chart(data)
        .mark_bar(cornerRadius=7, color=color)
        .encode(
            y=alt.Y(f"{cat}:N", sort="-x", title=None),
            x=alt.X(f"{val}:Q", title=title, axis=alt.Axis(format=fmt)),
            tooltip=[cat, alt.Tooltip(f"{val}:Q", format=",.2f")],
        )
        .properties(height=max(140, 48 * len(data)))
    )


def page_products(d, m_from, m_to):
    g = (
        d.groupby("หมวด", as_index=False)
        .agg(
            ยอดขาย=("Total_Amount", "sum"),
            **{"จำนวนชิ้น": ("Units_Sold", "sum")},
            รายการ=("Transaction_ID", "count"),
            ราคาเฉลี่ยต่อหน่วย=("Price_Per_Unit", "mean"),
        )
        .sort_values("ยอดขาย", ascending=False)
    )
    g["สัดส่วนยอดขาย (%)"] = g["ยอดขาย"] / g["ยอดขาย"].sum() * 100
    k1, k2, k3, k4 = st.columns(4)
    kpi(k1, "หมวดที่ขายดีที่สุด", g.iloc[0]["หมวด"])
    kpi(k2, "ยอดขายของหมวดนั้น (บาท)", fmt2(g.iloc[0]["ยอดขาย"]))
    kpi(k3, "จำนวนชิ้นที่ขายรวม (ชิ้น)", fmt0(g["จำนวนชิ้น"].sum()))
    kpi(k4, "ยอดขายเฉลี่ยต่อชิ้น (บาท)", fmt2(g["ยอดขาย"].sum() / g["จำนวนชิ้น"].sum()))

    st.subheader("ยอดขายแยกตามหมวดสินค้า")
    st.altair_chart(bar_h(g, "หมวด", "ยอดขาย", "ยอดขาย (บาท)"), **W)

    first, last = g.iloc[0], g.iloc[-1]
    if len(g) > 1:
        what = (
            f"หมวด <b>{first['หมวด']}</b> ยอดขายสูงสุด {fmt2(first['ยอดขาย'])} บาท ({first['สัดส่วนยอดขาย (%)']:.1f}% ของทั้งหมด) "
            f"ส่วนหมวดต่ำสุดคือ <b>{last['หมวด']}</b> ({last['สัดส่วนยอดขาย (%)']:.1f}%)"
        )
        so = (
            f"หมวดอันดับ 1 ห่างจากอันดับสุดท้าย {first['ยอดขาย'] / last['ยอดขาย']:.2f} เท่า "
            "หมวดหลักจึงมีผลต่อรายได้รวมมากที่สุด หากขาดสต็อกจะกระทบทันที"
        )
        now = (
            f"บริหารสต็อกหมวด {first['หมวด']} ให้เพียงพอ และทำโปรโมชันขายพ่วง (Cross-selling) "
            f"กับหมวด {last['หมวด']} เพื่อเพิ่มยอดต่อบิล"
        )
    else:
        what = f"ตัวกรองปัจจุบันเหลือเพียงหมวด <b>{first['หมวด']}</b> ยอดขาย {fmt2(first['ยอดขาย'])} บาท"
        so = "เมื่อดูหมวดเดียว ให้เปรียบเทียบด้วยแนวโน้มรายเดือนและภูมิภาคแทน"
        now = "เลือก 'ทุกหมวดสินค้า' เพื่อเปรียบเทียบระหว่างหมวด"
    insight_box(what, so, now)

    st.dataframe(
        g.style.format({
            "ยอดขาย": "{:,.2f}", "จำนวนชิ้น": "{:,.0f}", "รายการ": "{:,.0f}",
            "ราคาเฉลี่ยต่อหน่วย": "{:,.2f}", "สัดส่วนยอดขาย (%)": "{:.1f}",
        }),
        hide_index=True, **W,
    )

    st.subheader("แนวโน้มยอดขายรายเดือน แยกตามหมวดสินค้า")
    t = d.groupby(["month", "หมวด"], as_index=False)["Total_Amount"].sum().rename(columns={"Total_Amount": "ยอดขาย"})
    t["เดือน"] = t["month"].map(lambda i: SHORT[i])
    st.altair_chart(
        alt.Chart(t)
        .mark_line(point=True, strokeWidth=3)
        .encode(
            x=alt.X("เดือน:N", sort=SHORT[m_from: m_to + 1], axis=alt.Axis(labelAngle=0)),
            y=alt.Y("ยอดขาย:Q", axis=alt.Axis(format=",.0f"), title="ยอดขาย (บาท)"),
            color=alt.Color("หมวด:N"),
            tooltip=["เดือน", "หมวด", alt.Tooltip("ยอดขาย:Q", format=",.2f")],
        )
        .properties(height=360),
        **W,
    )

    st.subheader("เจาะลึกรายหมวดสินค้า")
    pick = st.selectbox("เลือกหมวดสินค้าที่ต้องการดูรายละเอียด", list(g["หมวด"]), key="drill_cat")
    dd = d[d["หมวด"] == pick]
    a, b = st.columns(2)
    with a:
        st.markdown(f"**{pick}: ยอดขายแยกตามภูมิภาค**")
        r = dd.groupby("ภูมิภาค", as_index=False)["Total_Amount"].sum().rename(columns={"Total_Amount": "ยอดขาย"})
        st.altair_chart(bar_h(r, "ภูมิภาค", "ยอดขาย", "ยอดขาย (บาท)", color="#0f7a3d"), **W)
    with b:
        st.markdown(f"**{pick}: ยอดขายแยกตามช่องทาง**")
        r = dd.groupby("ช่องทาง", as_index=False)["Total_Amount"].sum().rename(columns={"Total_Amount": "ยอดขาย"})
        st.altair_chart(bar_h(r, "ช่องทาง", "ยอดขาย", "ยอดขาย (บาท)", color="#c2410c"), **W)


def page_customers(d, df_all):
    per = (
        d.groupby("Customer_ID")
        .agg(ครั้งที่ซื้อ=("Transaction_ID", "count"), ยอดซื้อรวม=("Total_Amount", "sum"),
             ซื้อล่าสุด=("Transaction_Date", "max"))
        .reset_index()
    )
    n = len(per)
    repeat_rate = (per["ครั้งที่ซื้อ"] >= 2).mean() * 100
    k1, k2, k3, k4 = st.columns(4)
    kpi(k1, "ลูกค้าที่ซื้อ ไม่นับซ้ำ (คน)", fmt0(n))
    kpi(k2, "ซื้อเฉลี่ยต่อคน (ครั้ง)", fmt2(per["ครั้งที่ซื้อ"].mean()))
    kpi(k3, "ยอดซื้อเฉลี่ยต่อคน (บาท)", fmt2(per["ยอดซื้อรวม"].mean()))
    kpi(k4, "ลูกค้าซื้อซ้ำ (ตั้งแต่ 2 ครั้ง)", f"{repeat_rate:.1f}%")

    a, b = st.columns(2)
    with a:
        st.subheader("วิธีชำระเงิน")
        p = d.groupby("วิธีชำระเงิน", as_index=False).agg(
            ยอดขาย=("Total_Amount", "sum"), รายการ=("Transaction_ID", "count")
        )
        st.altair_chart(bar_h(p, "วิธีชำระเงิน", "รายการ", "จำนวนรายการ", color="#6d28d9"), **W)
    with b:
        st.subheader("การกระจายความถี่ในการซื้อ")
        fq = per.groupby("ครั้งที่ซื้อ", as_index=False).agg(**{"จำนวนลูกค้า": ("Customer_ID", "count")})
        st.altair_chart(
            alt.Chart(fq)
            .mark_bar(cornerRadius=7, color="#1d4ed8")
            .encode(
                x=alt.X("ครั้งที่ซื้อ:O", title="จำนวนครั้งที่ซื้อ", axis=alt.Axis(labelAngle=0)),
                y=alt.Y("จำนวนลูกค้า:Q", title="จำนวนลูกค้า (คน)"),
                tooltip=["ครั้งที่ซื้อ", "จำนวนลูกค้า"],
            )
            .properties(height=260),
            **W,
        )

    top_pay = p.sort_values("รายการ", ascending=False).iloc[0]
    one_time = 100 - repeat_rate
    insight_box(
        what_text=(
            f"ลูกค้าซื้อซ้ำ <b>{repeat_rate:.1f}%</b> (ซื้อครั้งเดียว {one_time:.1f}%) ซื้อเฉลี่ย {fmt2(per['ครั้งที่ซื้อ'].mean())} ครั้ง/คน "
            f"และวิธีชำระเงินยอดนิยมคือ <b>{top_pay['วิธีชำระเงิน']}</b> ({fmt0(top_pay['รายการ'])} รายการ)"
        ),
        so_what_text="อัตราซื้อซ้ำสะท้อนความภักดีของลูกค้า ซึ่งมีต้นทุนต่ำกว่าการหาลูกค้าใหม่ ส่วนวิธีชำระเงินที่นิยมบอกว่าควรจัดสิทธิพิเศษผ่านช่องทางใด",
        now_what_text=(
            f"ส่งคูปองกระตุ้นลูกค้าที่ซื้อครั้งเดียวให้กลับมาซื้อซ้ำ และทำโปรโมชันคู่กับการชำระผ่าน {top_pay['วิธีชำระเงิน']}"
        ),
    )

    st.subheader("ลูกค้าที่มียอดซื้อสูงสุด 10 อันดับ")
    top = per.sort_values("ยอดซื้อรวม", ascending=False).head(10).copy()
    top["ซื้อล่าสุด"] = top["ซื้อล่าสุด"].dt.strftime("%Y-%m-%d")
    st.dataframe(top.style.format({"ยอดซื้อรวม": "{:,.2f}"}), hide_index=True, **W)

    st.subheader("ข้อมูลโปรไฟล์ลูกค้า (เพศ / อายุ / ระดับสมาชิก)")
    cust = load_customers(str(CUST_FILE)) if CUST_FILE.exists() else None
    if cust is None:
        st.info(
            f"ℹ️ ไม่พบไฟล์ {CUST_FILE.name} ในโฟลเดอร์เดียวกับ app.py จึงยังไม่แสดงส่วนนี้ "
            "(วางไฟล์ customers_data.csv เพื่อเปิดใช้งานกราฟโปรไฟล์ลูกค้า)"
        )
        return
    m = per.merge(cust, on="Customer_ID", how="left")
    match = m["Customer_ID"].isin(cust["Customer_ID"]).mean() * 100
    st.caption(f"อัตราการจับคู่รหัสลูกค้ากับตารางลูกค้า (FK Match): {match:.1f}%")
    cols = st.columns(3)
    if "Membership_Tier" in m:
        t = (
            m.dropna(subset=["Membership_Tier"])
            .groupby("Membership_Tier", as_index=False)
            .agg(ยอดซื้อเฉลี่ยต่อคน=("ยอดซื้อรวม", "mean"), **{"จำนวนลูกค้า": ("Customer_ID", "count")})
            .rename(columns={"Membership_Tier": "ระดับสมาชิก"})
        )
        cols[0].markdown("**ยอดซื้อเฉลี่ยต่อคน ตามระดับสมาชิก**")
        cols[0].altair_chart(bar_h(t, "ระดับสมาชิก", "ยอดซื้อเฉลี่ยต่อคน", "บาท/คน", color="#b45309"), **W)
    if "Gender" in m:
        t = (
            m.dropna(subset=["Gender"])
            .assign(เพศ=lambda x: x["Gender"].map(GENDER_TH).fillna(x["Gender"]))
            .groupby("เพศ", as_index=False)
            .agg(ยอดซื้อเฉลี่ยต่อคน=("ยอดซื้อรวม", "mean"), **{"จำนวนลูกค้า": ("Customer_ID", "count")})
        )
        cols[1].markdown("**ยอดซื้อเฉลี่ยต่อคน ตามเพศ**")
        cols[1].altair_chart(bar_h(t, "เพศ", "ยอดซื้อเฉลี่ยต่อคน", "บาท/คน", color="#be185d"), **W)
    if "Age" in m:
        age = pd.to_numeric(m["Age"], errors="coerce")
        m["ช่วงอายุ"] = pd.cut(age, [0, 24, 34, 44, 54, 200], labels=["ต่ำกว่า 25", "25-34", "35-44", "45-54", "55 ขึ้นไป"])
        t = (
            m.dropna(subset=["ช่วงอายุ"])
            .groupby("ช่วงอายุ", as_index=False, observed=True)
            .agg(ยอดซื้อเฉลี่ยต่อคน=("ยอดซื้อรวม", "mean"), **{"จำนวนลูกค้า": ("Customer_ID", "count")})
        )
        t["ช่วงอายุ"] = t["ช่วงอายุ"].astype(str)
        cols[2].markdown("**ยอดซื้อเฉลี่ยต่อคน ตามช่วงอายุ**")
        cols[2].altair_chart(
            alt.Chart(t)
            .mark_bar(cornerRadius=7, color="#0f766e")
            .encode(
                x=alt.X("ช่วงอายุ:N", sort=["ต่ำกว่า 25", "25-34", "35-44", "45-54", "55 ขึ้นไป"], axis=alt.Axis(labelAngle=0)),
                y=alt.Y("ยอดซื้อเฉลี่ยต่อคน:Q", title="บาท/คน"),
                tooltip=["ช่วงอายุ", alt.Tooltip("ยอดซื้อเฉลี่ยต่อคน:Q", format=",.2f"), "จำนวนลูกค้า"],
            )
            .properties(height=260),
            **W,
        )


def page_eda(d):
    st.subheader("ตารางสถิติเชิงพรรณนา (Descriptive Statistics)")
    st.caption("แสดงค่าเฉลี่ย (Mean), มัธยฐาน (Median), ส่วนเบี่ยงเบนมาตรฐาน (SD) และค่าพิสัยระหว่างควอไทล์ (IQR)")

    col_map = {
        "Total_Amount": "ยอดขายรวมต่อรายการ (บาท)",
        "Units_Sold": "จำนวนชิ้นที่ซื้อ (ชิ้น)",
        "Price_Per_Unit": "ราคาขายต่อหน่วย (บาท)",
    }
    stats_rows = []
    for col, th_name in col_map.items():
        s = pd.to_numeric(d[col], errors="coerce").dropna()
        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        stats_rows.append({
            "ตัวแปรเชิงปริมาณ": th_name, "จำนวน (N)": len(s), "ค่าเฉลี่ย (Mean)": s.mean(),
            "มัธยฐาน (Median)": s.median(), "SD": s.std(), "Q1 (25%)": q1, "Q3 (75%)": q3,
            "IQR (Q3-Q1)": q3 - q1, "ต่ำสุด (Min)": s.min(), "สูงสุด (Max)": s.max(),
        })
    stats_df = pd.DataFrame(stats_rows)
    st.dataframe(
        stats_df.style.format({
            "จำนวน (N)": "{:,.0f}", "ค่าเฉลี่ย (Mean)": "{:,.2f}", "มัธยฐาน (Median)": "{:,.2f}",
            "SD": "{:,.2f}", "Q1 (25%)": "{:,.2f}", "Q3 (75%)": "{:,.2f}", "IQR (Q3-Q1)": "{:,.2f}",
            "ต่ำสุด (Min)": "{:,.2f}", "สูงสุด (Max)": "{:,.2f}",
        }),
        hide_index=True, **W,
    )

    corr_units = d["Units_Sold"].corr(d["Total_Amount"])
    corr_price = d["Price_Per_Unit"].corr(d["Total_Amount"])
    corr_units = 0.0 if pd.isna(corr_units) else corr_units
    corr_price = 0.0 if pd.isna(corr_price) else corr_price

    k1, k2, k3 = st.columns(3)
    kpi(k1, "สหสัมพันธ์: จำนวนชิ้น vs ยอดขาย", f"{corr_units:.3f}")
    kpi(k2, "สหสัมพันธ์: ราคาต่อหน่วย vs ยอดขาย", f"{corr_price:.3f}")
    kpi(k3, "IQR ยอดขายต่อรายการ (บาท)", fmt2(stats_rows[0]["IQR (Q3-Q1)"]))

    st.subheader("การกระจายตัวของยอดขายและความสัมพันธ์ระหว่างตัวแปร")
    e1, e2 = st.columns(2)
    with e1:
        st.markdown("**การกระจายตัวของยอดขายต่อรายการ (Histogram)**")
        hist = (
            alt.Chart(d[["Total_Amount"]])
            .mark_bar(color="#2563eb", cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("Total_Amount:Q", bin=alt.Bin(maxbins=25), title="ช่วงยอดขายต่อรายการ (บาท)"),
                y=alt.Y("count():Q", title="จำนวนรายการ (ความถี่)"),
                tooltip=[alt.Tooltip("count():Q", title="จำนวนรายการ")],
            )
            .properties(height=320)
        )
        st.altair_chart(hist, **W)

    with e2:
        st.markdown("**ความสัมพันธ์ระหว่างราคาต่อหน่วยกับจำนวนชิ้นที่ซื้อ**")
        sample_d = d[["หมวด", "Price_Per_Unit", "Units_Sold", "Total_Amount"]].sample(min(800, len(d)), random_state=42)
        scatter = (
            alt.Chart(sample_d)
            .mark_circle(size=55, opacity=0.6)
            .encode(
                x=alt.X("Price_Per_Unit:Q", title="ราคาขายต่อหน่วย (บาท)"),
                y=alt.Y("Units_Sold:Q", title="จำนวนชิ้นที่ซื้อ (ชิ้น)"),
                color=alt.Color("หมวด:N", title="หมวดสินค้า"),
                tooltip=["หมวด", "Price_Per_Unit", "Units_Sold", alt.Tooltip("Total_Amount:Q", format=",.2f")],
            )
            .properties(height=320)
        )
        st.altair_chart(scatter, **W)

    # นับ outlier จริงตามเกณฑ์ 1.5×IQR จากข้อมูลที่กรองอยู่ (ไม่ยืนยันลอยๆ ว่าทำความสะอาดแล้ว)
    s0 = stats_rows[0]
    upper = s0["Q3 (75%)"] + 1.5 * s0["IQR (Q3-Q1)"]
    lower = s0["Q1 (25%)"] - 1.5 * s0["IQR (Q3-Q1)"]
    n_out = int(((d["Total_Amount"] > upper) | (d["Total_Amount"] < lower)).sum())
    skew_note = "ค่าเฉลี่ยสูงกว่ามัธยฐาน แสดงว่าการกระจายเบ้ขวา" if s0["ค่าเฉลี่ย (Mean)"] > s0["มัธยฐาน (Median)"] else \
        "ค่าเฉลี่ยใกล้เคียงหรือต่ำกว่ามัธยฐาน การกระจายไม่เบ้ขวา"
    stronger = "จำนวนชิ้นที่ซื้อ" if abs(corr_units) >= abs(corr_price) else "ราคาต่อหน่วย"
    insight_box(
        what_text=(
            f"ยอดขายต่อรายการมีค่าเฉลี่ย {fmt2(s0['ค่าเฉลี่ย (Mean)'])} บาท มัธยฐาน {fmt2(s0['มัธยฐาน (Median)'])} บาท "
            f"(IQR = {fmt2(s0['IQR (Q3-Q1)'])} บาท) พบค่าผิดปกติตามเกณฑ์ 1.5×IQR จำนวน {n_out:,} รายการ "
            f"สหสัมพันธ์กับยอดขาย: จำนวนชิ้น r = {corr_units:.3f}, ราคาต่อหน่วย r = {corr_price:.3f}"
        ),
        so_what_text=f"{skew_note} และ <b>{stronger}</b> สัมพันธ์กับยอดขายต่อรายการมากกว่า จึงเป็นตัวแปรที่ควรให้ความสำคัญก่อน",
        now_what_text=f"ใช้ช่วง IQR ของยอดขายต่อบิลเป็นเกณฑ์ตั้งเป้ายอดขั้นต่ำสำหรับคูปองหรือส่งฟรี และตรวจสอบรายการผิดปกติ {n_out:,} รายการก่อนนำไปสรุปผล",
    )


# ======================= แถบสถานะ + หน้าที่เลือก =======================
status_slot = st.container()

if st.session_state.page == "ภาพรวมยอดขาย":
    page_overview(d, m_from, m_to)
elif st.session_state.page == "รายละเอียดสินค้า":
    page_products(d, m_from, m_to)
elif st.session_state.page == "พฤติกรรมลูกค้า":
    page_customers(d, df)
else:
    page_eda(d)

# วัดเวลาหลังเรนเดอร์จริง แล้วค่อยเติมลงแถบสถานะด้านบน
elapsed_ms = (time.perf_counter() - t_start) * 1000
with status_slot:
    sc1, sc2, sc3 = st.columns([2.6, 1.2, 1.2])
    sc1.success(f"✅ แสดงข้อมูลแล้ว: {fmt0(len(d))} / {fmt0(len(df))} รายการ (ประมวลผล {elapsed_ms:.0f} ms)")
    sc2.download_button(
        "⬇️ ส่งออก CSV", make_csv(fkey),
        file_name="filtered_transactions.csv", mime="text/csv", **W,
    )
    # Excel สร้างเมื่อกดปุ่มเท่านั้น (ช้าเมื่อข้อมูลเยอะ) แล้วจดจำไว้จนกว่าจะเปลี่ยนตัวกรอง
    if st.session_state.get("xlsx_key") == fkey and st.session_state.get("xlsx_bytes"):
        sc3.download_button(
            "📊 ดาวน์โหลด Excel (.xlsx)", st.session_state["xlsx_bytes"], file_name="SmartRetail_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **W,
        )
    elif sc3.button("📊 เตรียมไฟล์ Excel", **W):
        with st.spinner("⏳ กำลังสร้างไฟล์ Excel..."):
            xb = to_excel_bytes(d)
        if xb is None:
            st.warning("ติดตั้ง openpyxl เพื่อส่งออก Excel (pip install openpyxl)")
        else:
            st.session_state["xlsx_key"], st.session_state["xlsx_bytes"] = fkey, xb
            st.rerun()
