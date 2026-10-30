import streamlit as st
import pandas as pd
import datetime
import os
import re
import tempfile
import shutil
import uuid
from openpyxl import load_workbook, Workbook

st.set_page_config(
    page_title="LubaNur Order & Stock System",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === صلاحيات المستخدمين: مدير واحد + مستخدمون للعرض فقط ===
def get_secret(name, default=""):
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = os.environ.get(name, default)
    return str(value or default).strip()

ADMIN_PASSWORD = get_secret("ADMIN_PASSWORD", get_secret("APP_PASSWORD", ""))
VIEWER_PASSWORD = get_secret("VIEWER_PASSWORD", "")
if not ADMIN_PASSWORD or not VIEWER_PASSWORD:
    st.error("يرجى إعداد ADMIN_PASSWORD و VIEWER_PASSWORD في Streamlit Secrets.")
    st.code('ADMIN_PASSWORD = "كلمة_مرور_المدير"\nVIEWER_PASSWORD = "كلمة_مرور_المشاهد"\nGITHUB_TOKEN = "توكن_خاص"\nGITHUB_REPO = "elsaifeldiaa-cpu/lubanur"\nGITHUB_BRANCH = "main"\nGITHUB_EXCEL_PATH = "system_database.xlsx"', language="toml")
    st.stop()

if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if st.session_state.user_role is None:
    # شاشة دخول مستقلة؛ استخدام form يجعل الإرسال واضحاً ويعمل بزر Enter أيضاً.
    st.markdown("""
        <style>
        /* إخفاء عناصر التطبيق الجانبية أثناء تسجيل الدخول فقط */
        [data-testid="stSidebar"] {display: none;}
        [data-testid="stHeader"] {background: transparent;}
        .stApp {
            background: linear-gradient(135deg, #f4f1e8 0%, #e7eee9 52%, #f8f6ef 100%);
        }
        .block-container {
            max-width: 760px;
            padding-top: 7vh;
            padding-bottom: 5vh;
        }
        .login-shell {
            direction: rtl;
            text-align: center;
            background: rgba(255,255,255,.96);
            border: 1px solid #e5e6dc;
            border-radius: 24px;
            padding: 34px 34px 26px 34px;
            box-shadow: 0 18px 55px rgba(32, 65, 49, .12);
            max-width: 520px;
            margin: 0 auto 18px auto;
        }
        .login-mark {
            width: 76px; height: 76px; line-height: 76px;
            margin: 0 auto 14px auto; border-radius: 22px;
            background: #214b3a; color: #fff; font-size: 36px;
        }
        .login-title { color: #214b3a; font-size: 30px; font-weight: 800; margin: 0; }
        .login-subtitle { color: #6e786f; font-size: 15px; margin-top: 10px; }
        .login-foot { color: #849087; text-align: center; font-size: 12px; margin-top: 14px; }
        div[data-testid="stForm"] {
            background: rgba(255,255,255,.96);
            border: 1px solid #e5e6dc;
            border-top: 0; border-radius: 0 0 24px 24px;
            padding: 0 34px 28px 34px;
            max-width: 520px; margin: -18px auto 0 auto;
            box-shadow: 0 18px 55px rgba(32, 65, 49, .12);
        }
        div[data-testid="stForm"] label { direction: rtl; text-align: right; }
        div[data-testid="stForm"] input {
            text-align: right; direction: rtl; border-radius: 12px;
            min-height: 48px;
        }
        div[data-testid="stForm"] button[kind="primary"] {
            width: 100%; min-height: 48px; border-radius: 12px;
            background: #214b3a; border: 1px solid #214b3a;
            color: white; font-weight: 700; font-size: 17px;
        }
        div[data-testid="stForm"] button[kind="primary"]:hover {
            background: #173a2c; border-color: #173a2c;
        }
        @media (max-width: 600px) {
            .block-container {padding: 4vh 1rem 2vh 1rem;}
            .login-shell {padding: 26px 18px 24px 18px;}
            div[data-testid="stForm"] {padding: 0 18px 22px 18px;}
