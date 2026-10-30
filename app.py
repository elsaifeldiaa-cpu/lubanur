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

if st.session_state.user_role is None:
    # Modern responsive login screen; role is explicitly selected before password validation.
    st.markdown("""
    <style>
      .stApp {
        background: radial-gradient(circle at top right, #f4efe2 0%, #f7f8fb 42%, #eef1f6 100%);
      }
      [data-testid="stMainBlockContainer"] {
        padding-top: 5vh;
        max-width: 1100px;
      }
      .lubanur-login-title {
        text-align: center;
        color: #17243a;
        font-size: clamp(26px, 3.2vw, 38px);
        font-weight: 800;
        margin: 8px 0 4px 0;
      }
      .lubanur-login-subtitle {
        text-align: center;
        color: #7a6844;
        font-size: 14px;
        margin: 0 0 22px 0;
      }
      [data-testid="stForm"] {
        background: rgba(255,255,255,.96);
        border: 1px solid #e6dfcf;
        border-radius: 20px;
        padding: 24px;
        box-shadow: 0 16px 42px rgba(25,36,55,.10);
      }
      [data-testid="stForm"] button[kind="primaryFormSubmit"] {
        background: linear-gradient(100deg, #17243a, #263d5b);
        color: #fff;
        border: 1px solid #c6a45a;
        border-radius: 12px;
        min-height: 48px;
        font-weight: 700;
      }
      [data-testid="stForm"] input {
        border-radius: 10px;
      }
      @media (max-width: 700px) {
        [data-testid="stMainBlockContainer"] { padding: 18px 14px 30px 14px; }
        [data-testid="stForm"] { padding: 16px; }
      }
    </style>
    """, unsafe_allow_html=True)

    login_left, login_center, login_right = st.columns([1, 1.15, 1], gap="large")
    with login_center:
        st.markdown("<div style='text-align:center;font-size:46px;margin-bottom:4px'>🌿</div>", unsafe_allow_html=True)
        st.markdown("<div class='lubanur-login-title'>لبانور</div>", unsafe_allow_html=True)
        st.markdown("<div class='lubanur-login-subtitle'>نظام الإدارة والمبيعات والمخزون</div>", unsafe_allow_html=True)
        with st.form("lubanur_login_form", clear_on_submit=False, border=True):
            selected_login_role = st.radio(
                "نوع المستخدم",
                options=["مدير النظام", "مشاهد"],
                horizontal=True,
                key="login_role_choice"
            )
            pwd = st.text_input("كلمة المرور", type="password", key="login_pass",
                                placeholder="أدخل كلمة المرور")
            login_submitted = st.form_submit_button("دخول إلى النظام", use_container_width=True)

        if login_submitted:
            expected_password = ADMIN_PASSWORD if selected_login_role == "مدير النظام" else VIEWER_PASSWORD
            if not pwd:
                st.warning("يرجى إدخال كلمة المرور أولاً.")
            elif pwd == expected_password:
                st.session_state.user_role = "admin" if selected_login_role == "مدير النظام" else "viewer"
                st.session_state.authenticated = True
                st.session_state.login_event_pending = (
                    "تسجيل دخول المدير" if selected_login_role == "مدير النظام"
                    else "تسجيل دخول مستخدم العرض"
                )
                st.rerun()
            else:
                st.error("كلمة المرور غير صحيحة لنوع المستخدم المحدد.")
    st.stop()

IS_ADMIN = st.session_state.user_role == "admin"

# === تخصيص CSS لمظهر عصري وواضح (Elite UI) ===
st.markdown("""
    <style>
    .main {
        direction: rtl;
        text-align: right;
    }
    .stSidebar {
        background-color: #f8f9fa;
        padding-top: 10px;
    }
    .sidebar-title {
        font-size: 16px;
        font-weight: bold;
        color: #2c3e50;
        margin-bottom: 10px;
        text-align: center;
        border-bottom: 2px solid #007bff;
        padding-bottom: 8px;
    }
    .version-tag {
        font-size: 12px;
        color: #666;
        text-align: left;
        padding: 5px;
    }
    </style>
""", unsafe_allow_html=True)

DB_FILE = "system_database.xlsx"
ORIGINAL_EXCEL = "System.xlsx"

# GitHub workbook synchronization. Configure these values in Streamlit Secrets.
GITHUB_TOKEN = get_secret("GITHUB_TOKEN", "")
GITHUB_REPO = get_secret("GITHUB_REPO", "elsaifeldiaa-cpu/lubanur")
GITHUB_BRANCH = get_secret("GITHUB_BRANCH", "main")
GITHUB_EXCEL_PATH = get_secret("GITHUB_EXCEL_PATH", "system_database.xlsx").lstrip("/")

import json, base64, urllib.request, urllib.error

def github_api(method, path, payload=None):
    if not GITHUB_TOKEN:
        raise RuntimeError("GITHUB_TOKEN غير مضبوط في Streamlit Secrets")
    url = "https://api.github.com/repos/" + GITHUB_REPO + "/" + path.lstrip("/")
    body = None if payload is None else json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=body, method=method, headers={
        "Authorization": "Bearer " + GITHUB_TOKEN,
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "Content-Type": "application/json",
        "User-Agent": "LubaNur-Streamlit-App"
    })
    try:
        with urllib.request.urlopen(req, timeout=30) as response:
            raw = response.read()
            return json.loads(raw.decode("utf-8")) if raw else {}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"GitHub API ({exc.code}): {detail[:500]}") from exc

def github_download_workbook():
    """Pull the repository's workbook before app initialization."""
    if not GITHUB_TOKEN:
        return False
    try:
        meta = github_api("GET", f"contents/{GITHUB_EXCEL_PATH}?ref={GITHUB_BRANCH}")
        content = base64.b64decode(meta["content"])
        tmp = DB_FILE + ".github.tmp"
        with open(tmp, "wb") as f:
            f.write(content)
        _read_workbook_safely(tmp)
        os.replace(tmp, DB_FILE)
        return True
    except Exception as exc:
        # Do not silently replace a valid local workbook with an empty one.
        if os.path.exists(DB_FILE) and os.path.getsize(DB_FILE) > 0:
            st.warning(f"تعذر جلب أحدث Excel من GitHub؛ سيتم فتح النسخة المحلية: {exc}")
            return False
        raise RuntimeError(f"تعذر تنزيل ملف Excel من GitHub: {exc}") from exc

def github_upload_workbook():
    """Commit the current workbook to GitHub using its current SHA."""
    if not GITHUB_TOKEN:
        raise RuntimeError("لم تتم المزامنة: أضف GITHUB_TOKEN إلى Streamlit Secrets")
    meta = github_api("GET", f"contents/{GITHUB_EXCEL_PATH}?ref={GITHUB_BRANCH}")
    with open(DB_FILE, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("ascii")
    github_api("PUT", f"contents/{GITHUB_EXCEL_PATH}", {
        "message": "Update LubaNur database after transaction",
        "content": encoded,
        "sha": meta["sha"],
        "branch": GITHUB_BRANCH
    })

EXACT_ITEMS = [
    "لبانور نسيم البحر غسول 500 مل",
    "لبانور مانجو بابايا غسول 500مل",
    "لبانور اللبان غسول لليدين 500مل",
    "لبانور عود غسول لليدين 500مل",
    "لبانور نسيم الزهور غسول لليدين 500مل",
    "لبانور شغف غسول 500مل",
    "لبانور غسول لليدين 4 لتر",
    "لبانور لافندر شاور جيل 1000 مل",
    "لبانور شغف شاور جيل 1000 مل",
    "لبانور مانجو بابايا شاور جيل 1000 مل",
    "لبانور فراولة شاور جيل 1000 مل",
    "لبانور عود شاور جيل 1000 مل",
    "لبانور عطور فرنسية شاور جيل 1000 مل",
    "لبانور اللبان شاور جيل 1000 مل",
    "لبانور شامبو لبان 1000 مل",
    "لبانور شامبو شجرة السدر و النعناع 1000 مل",
    "لبانور شامبو بالروزماري 1000 مل",
    "لبانور شامبو جوز الهند 1000 مل",
    "لبانور بلسم لبان 1000 مل",
    "لبانور بلسم روزماري 1000 مل",
    "لبانور بخور ملكي معطر 500 مل",
    "لبانور شغف معطر 500 مل",
    "لبانور لمسه ملكيه معطر 500 مل",
    "لبانور عبق العود معطر 500 مل",
    "لبانور خوخ معطر 500 مل",
    "لبانور فراولة معطر 500 مل",
    "لبانور لافندر معطر 500 مل",
    "لبانور عبق اللبان معطر 500 مل"
]

DEFAULT_CARGO_COMPANIES = ["وصل لي", "نول", "جيناكم", "دليلي", "بدون شحن"]

AUDIT_SHEET = "Activity_Log"
AUDIT_COLUMNS = ["Timestamp", "User_Role", "Session_ID", "Action", "Details", "Connection_Status"]

if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())[:12]


def append_activity_log(path, action, details="", connection_status="غير محدد"):
    """Append a durable audit row to the workbook without changing business sheets."""
    try:
        if os.path.exists(path) and os.path.getsize(path) > 0:
            wb = load_workbook(path)
        else:
            wb = Workbook()
        if AUDIT_SHEET in wb.sheetnames:
            ws = wb[AUDIT_SHEET]
        else:
            ws = wb.create_sheet(AUDIT_SHEET)
            ws.append(AUDIT_COLUMNS)
        if ws.max_row == 0:
            ws.append(AUDIT_COLUMNS)
        ws.append([
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            st.session_state.get("user_role", "unknown"),
            st.session_state.get("session_id", ""),
            str(action)[:150], str(details)[:1000], str(connection_status)[:100]
        ])
        wb.save(path)
        wb.close()
    except Exception as exc:
        # Audit failures must be visible; do not hide them behind a successful business save.
        st.warning(f"تعذر تسجيل حدث في سجل الحركة: {exc}")


def read_activity_log(path=DB_FILE):
    try:
        if os.path.exists(path):
            xls = pd.ExcelFile(path, engine="openpyxl")
            if AUDIT_SHEET in xls.sheet_names:
                return pd.read_excel(xls, sheet_name=AUDIT_SHEET)
    except Exception as exc:
        st.warning(f"تعذر قراءة سجل الحركة: {exc}")
    return pd.DataFrame(columns=AUDIT_COLUMNS)


def normalize_orders_df(orders):
    if orders.empty:
        return pd.DataFrame(columns=["Order_ID", "Date", "Customer", "Phone", "Region", "Items_Detail", "Products_Total", "Shipping_Price", "Total", "Order_Payment", "Shipping_Payment", "Cargo_Company", "Delivery_Type", "Delivery_Status", "Collection_Status", "Bags", "Cargo_Barcode"])
    
    orders.columns = [str(c).strip() for c in orders.columns]
    
    rename_map = {
        'رقم الطلب': 'Order_ID', 'order_id': 'Order_ID', 'Order_ID': 'Order_ID',
        'التاريخ': 'Date', 'date': 'Date', 'Date': 'Date',
        'العميل': 'Customer', 'customer': 'Customer', 'Customer': 'Customer', 'اسم العميل': 'Customer',
        'الهاتف': 'Phone', 'phone': 'Phone', 'Phone': 'Phone', 'رقم الهاتف': 'Phone',
        'المنطقة': 'Region', 'region': 'Region', 'Region': 'Region', 'المنطقة / الولاية': 'Region',
        'تفاصيل المنتجات': 'Items_Detail', 'items_detail': 'Items_Detail', 'Items_Detail': 'Items_Detail',
        'قيمة المنتجات': 'Products_Total', 'products_total': 'Products_Total', 'Products_Total': 'Products_Total',
        'الشحن': 'Shipping_Price', 'shipping_price': 'Shipping_Price', 'Shipping_Price': 'Shipping_Price', 'سعر الشحن': 'Shipping_Price',
        'الإجمالي الكلي': 'Total', 'total': 'Total', 'Total': 'Total',
        'حالة دفع الطلب': 'Order_Payment', 'order_payment': 'Order_Payment', 'Order_Payment': 'Order_Payment',
        'حالة دفع الشحن': 'Shipping_Payment', 'shipping_payment': 'Shipping_Payment', 'Shipping_Payment': 'Shipping_Payment',
        'شركة الشحن': 'Cargo_Company', 'cargo_company': 'Cargo_Company', 'Cargo_Company': 'Cargo_Company',
        'نوع التسليم': 'Delivery_Type', 'delivery_type': 'Delivery_Type', 'Delivery_Type': 'Delivery_Type',
        'حالة التوصيل': 'Delivery_Status', 'delivery_status': 'Delivery_Status', 'Delivery_Status': 'Delivery_Status',
        'حالة التحصيل': 'Collection_Status', 'collection_status': 'Collection_Status', 'Collection_Status': 'Collection_Status',
        'الأكياس': 'Bags', 'bags': 'Bags', 'Bags': 'Bags', 'عدد الأكياس': 'Bags',
        'باركود شركة الشحن': 'Cargo_Barcode', 'cargo_barcode': 'Cargo_Barcode', 'Cargo_Barcode': 'Cargo_Barcode'
    }
    orders = orders.rename(columns=rename_map)
    orders.columns = [str(c).strip() for c in orders.columns]
    
    expected_defaults = {
        "Order_ID": 0, "Date": "غير محدد", "Customer": "غير محدد", "Phone": "غير محدد", 
        "Region": "غير محدد", "Items_Detail": "غير محدد", "Products_Total": 0.0, 
        "Shipping_Price": 0.0, "Total": 0.0, "Order_Payment": "لم يتم الدفع", 
        "Shipping_Payment": "لم يتم دفع الشحن", "Cargo_Company": "وصل لي", 
        "Delivery_Type": "من المكتب (Pickup)", "Delivery_Status": "قيد التحضير", 
        "Collection_Status": "قيد التحصيل", "Bags": 1, "Cargo_Barcode": ""
    }
    
    for col, default_val in expected_defaults.items():
        if col not in orders.columns:
            orders[col] = default_val
        else:
            if col in ["Customer", "Phone", "Region", "Cargo_Company", "Items_Detail", "Order_Payment", "Shipping_Payment", "Delivery_Type", "Delivery_Status", "Collection_Status", "Date", "Cargo_Barcode"]:
                orders[col] = orders[col].fillna(default_val).astype(str)
                orders[col] = orders[col].replace(["nan", "None", "", "NaN"], default_val)
            elif col in ["Products_Total", "Shipping_Price", "Total", "Bags", "Order_ID"]:
                orders[col] = pd.to_numeric(orders[col], errors='coerce').fillna(default_val)
                if col in ["Bags", "Order_ID"]:
                    orders[col] = orders[col].astype(int)
                    
    orders['Total'] = orders['Products_Total'] + orders['Shipping_Price']
    # Product collection is determined ONLY by invoice/product payment, never shipping payment.
    orders['Collection_Status'] = orders['Order_Payment'].astype(str).map(
        lambda value: 'مدفوع بالكامل' if value.strip() == 'تم الدفع' else 'قيد التحصيل'
    )
    return orders

def parse_items_to_list(items_str):
    if not items_str or str(items_str).lower() == "nan":
        return []
    parts = str(items_str).split(" | ")
    parsed_items = []
    for part in parts:
        match = re.search(r'(.+?)\s*\(الكمية:\s*(\d+)\)', part)
        if match:
            item_name = match.group(1).strip()
            qty = int(match.group(2))
            parsed_items.append({"Item": item_name, "Qty": qty})
        else:
            item_name = part.strip()
            if item_name:
                parsed_items.append({"Item": item_name, "Qty": 1})
    return parsed_items

def item_quantities(items):
    quantities = {}
    for item in items or []:
        name = str(item.get("Item", "")).strip()
        try:
            qty = int(item.get("Qty", 0))
        except (TypeError, ValueError):
            qty = 0
        if name and qty > 0:
            quantities[name] = quantities.get(name, 0) + qty
    return quantities

def adjust_stock_for_order(stock_df, old_items, new_items):
    """Apply only the quantity delta between old and new invoice items.
    Returns (updated_stock, error_message); no changes are applied on insufficient stock.
    """
    updated = stock_df.copy(deep=True)
    old_q = item_quantities(old_items)
    new_q = item_quantities(new_items)
    all_names = set(old_q) | set(new_q)
    for name in all_names:
        if name not in updated["Item_Name"].astype(str).values:
            if new_q.get(name, 0) > old_q.get(name, 0):
                return stock_df, f"الصنف «{name}» غير موجود في جدول المخزون."
            continue
        idx = updated.index[updated["Item_Name"].astype(str) == name][0]
        qty_value = pd.to_numeric(updated.loc[idx, "Qty"], errors="coerce")
        current_qty = int(qty_value) if pd.notna(qty_value) else 0
        delta = old_q.get(name, 0) - new_q.get(name, 0)
        # delta > 0 means items were removed from the invoice and stock must be returned.
        if delta < 0 and current_qty < -delta:
            return stock_df, f"الكمية المتاحة من «{name}» غير كافية. المتاح: {current_qty}، المطلوب إضافته للفاتورة: {-delta}."
        updated.loc[idx, "Qty"] = current_qty + delta
    return updated, None

def _read_workbook_safely(path):
    """Read and validate the required workbook sheets without modifying the source file."""
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        raise RuntimeError(f"ملف قاعدة البيانات غير موجود أو فارغ: {path}")
    xls = pd.ExcelFile(path, engine="openpyxl")
    required = {"Orders", "Stock", "Expenses", "Cargo_Companies"}
    missing = required.difference(xls.sheet_names)
    if missing:
        raise RuntimeError("ملف قاعدة البيانات لا يحتوي على الأوراق المطلوبة: " + ", ".join(sorted(missing)))
    orders = normalize_orders_df(pd.read_excel(xls, sheet_name="Orders"))
    stock = pd.read_excel(xls, sheet_name="Stock")
    if not {"Item_Name", "Qty"}.issubset(stock.columns):
        raise RuntimeError("ورقة Stock لا تحتوي على الأعمدة Item_Name و Qty. لم يتم تعديل الملف الأصلي.")
    expenses = pd.read_excel(xls, sheet_name="Expenses")
    cargo_df = pd.read_excel(xls, sheet_name="Cargo_Companies")
    if "Cargo_Name" not in cargo_df.columns:
        raise RuntimeError("ورقة Cargo_Companies لا تحتوي على العمود Cargo_Name. لم يتم تعديل الملف الأصلي.")
    return orders, stock, expenses, cargo_df


def save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list, product_costs_df=None):
    """Save locally, back up, then synchronize the committed workbook to GitHub."""
    if not IS_ADMIN:
        raise PermissionError("حساب العرض فقط؛ لا يملك صلاحية تعديل البيانات.")
    db_dir = os.path.dirname(os.path.abspath(DB_FILE))
    os.makedirs(db_dir, exist_ok=True)
    backup_dir = os.path.join(db_dir, "database_backups")
    cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
    if product_costs_df is None:
        product_costs_df = load_product_costs(DB_FILE)

    # Never overwrite an unreadable source workbook. A bad read must be investigated first.
    if os.path.exists(DB_FILE) and os.path.getsize(DB_FILE) > 0:
        try:
            old_orders, _, _, _ = _read_workbook_safely(DB_FILE)
        except Exception as exc:
            raise RuntimeError(f"تم إيقاف الحفظ لحماية البيانات؛ تعذرت قراءة قاعدة البيانات الحالية: {exc}") from exc
        old_count = len(old_orders)
        new_count = len(orders_df)
        # Individual invoice deletion is supported; a sudden drop by more than one invoice is blocked.
        if old_count - new_count > 1:
            raise RuntimeError(
                f"تم إيقاف الحفظ لحماية الفواتير: عدد الفواتير الحالية {old_count} والجديدة {new_count}. "
                "الانخفاض أكبر من فاتورة واحدة. لم يتم استبدال قاعدة البيانات. راجع البيانات أو استعد نسخة احتياطية صحيحة."
            )
        os.makedirs(backup_dir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_path = os.path.join(backup_dir, f"system_database_{stamp}.xlsx")
        shutil.copy2(DB_FILE, backup_path)

    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=db_dir)
    os.close(fd)
    try:
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            orders_df.to_excel(writer, sheet_name="Orders", index=False)
            stock_df.to_excel(writer, sheet_name="Stock", index=False)
            expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
            cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
            product_costs_df.to_excel(writer, sheet_name="Product_Costs", index=False)
            existing_log = read_activity_log(DB_FILE)
            if existing_log.empty:
                existing_log = pd.DataFrame(columns=AUDIT_COLUMNS)
            new_log = pd.DataFrame([{
                "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "User_Role": st.session_state.get("user_role", "unknown"),
                "Session_ID": st.session_state.get("session_id", ""),
                "Action": "حفظ معاملة / تحديث قاعدة البيانات",
                "Details": f"عدد الفواتير: {len(orders_df)} | عدد أصناف المخزون: {len(stock_df)} | المصروفات: {len(expenses_df)}",
                "Connection_Status": "جارٍ مزامنة GitHub"
            }])
            pd.concat([existing_log, new_log], ignore_index=True).to_excel(writer, sheet_name=AUDIT_SHEET, index=False)
        # Verify the temporary workbook before replacing the active database.
        _read_workbook_safely(temp_path)
        os.replace(temp_path, DB_FILE)
        try:
            github_upload_workbook()
            append_activity_log(DB_FILE, "نجاح مزامنة GitHub", "تم رفع ملف Excel بعد الحفظ", "متصل - تمت المزامنة")
        except Exception as sync_exc:
            append_activity_log(DB_FILE, "فشل مزامنة GitHub", str(sync_exc), "غير متصل أو رفض GitHub الطلب")
            raise RuntimeError(f"تم حفظ البيانات محلياً، لكن فشلت مزامنتها مع GitHub: {sync_exc}") from sync_exc
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
    st.cache_data.clear()


def init_db():
    """Initialize only a genuinely new database; never replace an existing broken database."""
    marker_file = ".system_database_initialized"
    if os.path.exists(DB_FILE) and os.path.getsize(DB_FILE) > 0:
        # Validate existing data; any error stops startup without overwriting the source.
        _read_workbook_safely(DB_FILE)
        with open(marker_file, "w", encoding="utf-8") as marker:
            marker.write("initialized\n")
        return

    # If the database existed previously, its disappearance must not trigger an empty replacement.
    if os.path.exists(marker_file):
        raise RuntimeError(
            "ملف system_database.xlsx مفقود أو فارغ رغم أن النظام سبق تهيئته. "
            "تم إيقاف التشغيل لمنع إنشاء قاعدة بيانات جديدة وفقدان الفواتير. أعد الملف من نسخة احتياطية."
        )

    if os.path.exists(ORIGINAL_EXCEL) and os.path.getsize(ORIGINAL_EXCEL) > 0:
        try:
            xls_orig = pd.ExcelFile(ORIGINAL_EXCEL, engine="openpyxl")
            stock = pd.read_excel(xls_orig, sheet_name="Stock +").groupby("Item Name")["Qty"].sum().reset_index()
            stock.columns = ["Item_Name", "Qty"]
        except Exception as exc:
            raise RuntimeError(
                f"تعذرت قراءة ملف التأسيس {ORIGINAL_EXCEL}: {exc}. لم يتم إنشاء قاعدة بيانات فارغة."
            ) from exc
    else:
        stock = pd.DataFrame({"Item_Name": EXACT_ITEMS, "Qty": [50] * len(EXACT_ITEMS)})

    orders = pd.DataFrame(columns=["Order_ID", "Date", "Customer", "Phone", "Region", "Items_Detail", "Products_Total", "Shipping_Price", "Total", "Order_Payment", "Shipping_Payment", "Cargo_Company", "Delivery_Type", "Delivery_Status", "Collection_Status", "Bags", "Cargo_Barcode"])
    expenses = pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])
    cargo_df = pd.DataFrame({"Cargo_Name": DEFAULT_CARGO_COMPANIES})
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(os.path.abspath(DB_FILE)))
    os.close(fd)
    try:
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            stock.to_excel(writer, sheet_name="Stock", index=False)
            expenses.to_excel(writer, sheet_name="Expenses", index=False)
            cargo_df.to_excel(writer, sheet_name="Cargo_Companies", index=False)
        _read_workbook_safely(temp_path)
        os.replace(temp_path, DB_FILE)
        with open(marker_file, "w", encoding="utf-8") as marker:
            marker.write("initialized\n")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


try:
    GITHUB_DATA_LOADED = github_download_workbook()
    init_db()
except Exception as db_init_error:
    st.error("تعذر فتح قاعدة البيانات بأمان. لم يتم إنشاء ملف بديل أو الكتابة فوق البيانات الموجودة.")
    st.exception(db_init_error)
    st.stop()


def load_product_costs(path=DB_FILE):
    """Read private product unit costs from their dedicated workbook sheet."""
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
        if "Product_Costs" in xls.sheet_names:
            costs = pd.read_excel(xls, sheet_name="Product_Costs")
        else:
            costs = pd.DataFrame(columns=["Item_Name", "Unit_Cost"])
    except Exception:
        costs = pd.DataFrame(columns=["Item_Name", "Unit_Cost"])
    for col in ["Item_Name", "Unit_Cost"]:
        if col not in costs.columns:
            costs[col] = "" if col == "Item_Name" else 0.0
    costs = costs[["Item_Name", "Unit_Cost"]].copy()
    costs["Item_Name"] = costs["Item_Name"].astype(str).str.strip()
    costs["Unit_Cost"] = pd.to_numeric(costs["Unit_Cost"], errors="coerce").fillna(0.0)
    return costs[costs["Item_Name"] != ""].drop_duplicates("Item_Name", keep="last").reset_index(drop=True)


def calculate_profitability(orders_df, costs_df, expenses_df):
    cost_map = dict(zip(costs_df["Item_Name"].astype(str), pd.to_numeric(costs_df["Unit_Cost"], errors="coerce").fillna(0.0)))
    qty_by_item = {}
    for _, order in orders_df.iterrows():
        for item in parse_items_to_list(order.get("Items_Detail", "")):
            name = str(item.get("Item", "")).strip()
            qty = int(item.get("Qty", 0) or 0)
            if name and qty > 0:
                qty_by_item[name] = qty_by_item.get(name, 0) + qty
    rows = []
    total_cogs = 0.0
    for name, qty in sorted(qty_by_item.items()):
        unit_cost = float(cost_map.get(name, 0.0))
        subtotal = unit_cost * qty
        total_cogs += subtotal
        rows.append({"المنتج": name, "الكمية المباعة": qty, "تكلفة الوحدة (ر.ع)": unit_cost, "إجمالي التكلفة (ر.ع)": subtotal, "حالة التكلفة": "مدخلة" if name in cost_map and unit_cost > 0 else "لم تُدخل"})
    sales = pd.to_numeric(orders_df.get("Products_Total", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    expenses = pd.to_numeric(expenses_df.get("Amount", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    return pd.DataFrame(rows), float(sales), float(total_cogs), float(expenses)


def load_data():
    st.cache_data.clear()
    try:
        orders, stock, expenses, cargo_df = _read_workbook_safely(DB_FILE)
    except Exception as exc:
        st.error("تعذر تحميل قاعدة البيانات. أوقف النظام الحفظ لحماية الفواتير الحالية.")
        st.exception(exc)
        st.stop()

    cargo_list = cargo_df["Cargo_Name"].dropna().astype(str).tolist() if not cargo_df.empty else DEFAULT_CARGO_COMPANIES
    if "وصل لي" in cargo_list:
        cargo_list.remove("وصل لي")
        cargo_list.insert(0, "وصل لي")
    else:
        cargo_list.insert(0, "وصل لي")
    return orders, stock, expenses, cargo_list

def render_invoice_snippet(order_row):
    raw_items = str(order_row['Items_Detail'])
    formatted_items_text = raw_items.replace(" | ", "<br>- ")
    if formatted_items_text and not formatted_items_text.startswith("-"):
        formatted_items_text = "- " + formatted_items_text

    barcode_str = f"<b>باركود الشحن:</b> {order_row.get('Cargo_Barcode', 'غير متوفر')}" if order_row.get('Cargo_Barcode') else ""

    receipt_html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="utf-8">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <style>
            body {{ font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 10px; }}
            .receipt-container {{ border: 2px dashed #2c3e50; border-radius: 12px; padding: 20px; background-color: #ffffff; direction: rtl; text-align: right; box-shadow: 0 4px 10px rgba(0,0,0,0.1); max-width: 420px; margin: 0 auto; }}
            h3 {{ text-align: center; color: #2c3e50; margin-bottom: 5px; }}
            .tagline {{ text-align: center; color: #666; font-size: 12px; margin-top: 0px; border-bottom: 1px dashed #ddd; padding-bottom: 10px; }}
            .info {{ font-size: 14px; color: #333; line-height: 1.6; }}
            .items {{ font-size: 13px; color: #333; background: #f9f9f9; padding: 10px; border-radius: 6px; line-height: 1.6; margin: 10px 0; }}
            .total {{ font-size: 16px; color: #007bff; font-weight: bold; text-align: center; margin-top: 10px; }}
            .btn-container {{ display: flex; gap: 8px; justify-content: center; margin-top: 15px; flex-wrap: wrap; max-width: 420px; margin-left: auto; margin-right: auto; }}
            .btn {{ padding: 10px 14px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; font-size: 13px; color: white; }}
            .btn-copy {{ background-color: #007bff; }}
            .btn-download {{ background-color: #ff9800; }}
            .btn-print {{ background-color: #28a745; }}
            .toast {{ text-align: center; color: green; font-weight: bold; margin-top: 8px; font-size: 13px; }}
        </style>
    </head>
    <body>
        <div id="receipt-capture" class="receipt-container">
            <h3>=== لبانور - LubaNur ===</h3>
            <div class="tagline">السيف والضياء للاستثمار الدولي</div>
            <div class="info">
                <p><b>رقم الفاتورة:</b> #{order_row['Order_ID']} &nbsp;|&nbsp; <b>التاريخ:</b> <span style="direction: ltr; display: inline-block;">{order_row['Date']}</span></p>
                <p><b>العميل:</b> {order_row['Customer']} &nbsp;|&nbsp; <b>الهاتف:</b> <span style="direction: ltr; display: inline-block;">{order_row['Phone']}</span></p>
                <p><b>المنطقة:</b> {order_row['Region']}</p>
                <p><b>نوع التسليم:</b> {order_row.get('Delivery_Type', 'غير محدد')} &nbsp;|&nbsp; <b>شركة الشحن:</b> {order_row.get('Cargo_Company', 'غير محدد')}</p>
                <p><b>حالة دفع الطلب:</b> {order_row.get('Order_Payment', 'لم يتم الدفع')} &nbsp;|&nbsp; <b>دفع الشحن:</b> {order_row.get('Shipping_Payment', 'لم يتم دفع الشحن')}</p>
                <p>{barcode_str}</p>
            </div>
            <hr style="border: 0.5px dashed #ccc;">
            <strong>المنتجات:</strong>
            <div class="items">{formatted_items_text}</div>
            <hr style="border: 0.5px dashed #ccc;">
            <div class="info">
                <p><b>قيمة المنتجات:</b> {order_row['Products_Total']} ر.ع</p>
                <p><b>سعر الشحن:</b> {order_row['Shipping_Price']} ر.ع</p>
                <div class="total">الإجمالي الكلي: {order_row['Total']} ر.ع</div>
            </div>
            <p style="text-align: center; font-weight: bold; color: #2c3e50; margin-top: 15px; border-top: 1px dashed #ddd; padding-top: 10px;">شكراً لتعاملكم معنا!</p>
        </div>
        
        <div class="btn-container">
            <button class="btn btn-copy" onclick="copyImage()">📋 نسخ الصورة</button>
            <button class="btn btn-download" onclick="downloadImage()">📥 تحميل كصورة</button>
            <button class="btn btn-print" onclick="window.print()">🖨️ طباعة / حفظ PDF</button>
        </div>
        <div id="toast-msg" class="toast"></div>

        <script>
            function copyImage() {{
                const el = document.getElementById('receipt-capture');
                html2canvas(el, {{ scale: 2, useCORS: true, scrollY: -window.scrollY }}).then(canvas => {{
                    canvas.toBlob(blob => {{
                        try {{
                            navigator.clipboard.write([
                                new ClipboardItem({{ 'image/png': blob }})
                            ]).then(() => {{
                                document.getElementById('toast-msg').innerText = '✅ تم نسخ صورة الفاتورة بنجاح!';
                            }}).catch(err => {{
                                alert('يرجى استخدام زر "تحميل كصورة" لأجهزة أبل والأيباد.');
                            }});
                        }} catch(e) {{
                            alert('يرجى استخدام زر "تحميل كصورة" لأجهزة أبل والأيباد.');
                        }}
                    }});
                }});
            }}
            function downloadImage() {{
                const el = document.getElementById('receipt-capture');
                html2canvas(el, {{ scale: 2, useCORS: true, scrollY: -window.scrollY }}).then(canvas => {{
                    const link = document.createElement('a');
                    link.download = 'Invoice_#{order_row["Order_ID"]}.png';
                    link.href = canvas.toDataURL('image/png');
                    link.click();
                    document.getElementById('toast-msg').innerText = '✅ تم تحميل الفاتورة بنجاح!';
                }});
            }}
        </script>
    </body>
    </html>
    """
    return receipt_html

orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
if st.session_state.get("login_event_pending"):
    append_activity_log(DB_FILE, st.session_state.pop("login_event_pending"), "تم فتح جلسة جديدة", "تم تحميل البيانات")
if not st.session_state.get("startup_logged"):
    append_activity_log(DB_FILE, "بدء جلسة التطبيق", "تم تشغيل التطبيق وتحميل ملف البيانات", "تم تحميل أحدث نسخة من GitHub" if GITHUB_DATA_LOADED else "استخدمت النسخة المحلية؛ لم يثبت اتصال GitHub")
    st.session_state.startup_logged = True

st.sidebar.markdown('<p class="sidebar-title">🌟 لبانور - لوحة التحكم الإدارية</p>', unsafe_allow_html=True)
st.sidebar.markdown('<p class="version-tag">📌 الإصدار المحدث: v6:10pm 28-10-2026</p>', unsafe_allow_html=True)
st.sidebar.info("👑 المدير" if IS_ADMIN else "👁️ وضع العرض فقط")

# All action buttons are disabled for viewers; navigation remains available via selectbox.
_original_button = st.button
def _permission_button(*args, **kwargs):
    if not IS_ADMIN:
        kwargs["disabled"] = True
    return _original_button(*args, **kwargs)
st.button = _permission_button
_original_form_submit = st.form_submit_button
def _permission_form_submit(*args, **kwargs):
    if not IS_ADMIN:
        kwargs["disabled"] = True
    return _original_form_submit(*args, **kwargs)
st.form_submit_button = _permission_form_submit


if st.sidebar.button("🔄 إعادة تحميل البيانات من ملف النظام"):
    st.cache_data.clear()
    st.rerun()
st.sidebar.caption("تتم مزامنة ملف Excel مع GitHub بعد كل عملية حفظ ناجحة.")
if IS_ADMIN and st.sidebar.button("☁️ مزامنة Excel الآن مع GitHub"):
    try:
        github_upload_workbook()
        append_activity_log(DB_FILE, "مزامنة يدوية ناجحة", "ضغط المستخدم زر المزامنة اليدوية", "متصل - تمت المزامنة")
        st.sidebar.success("تمت مزامنة Excel مع GitHub بنجاح")
    except Exception as sync_error:
        append_activity_log(DB_FILE, "فشل المزامنة اليدوية", str(sync_error), "غير متصل أو رفض GitHub الطلب")
        st.sidebar.error(f"فشلت المزامنة: {sync_error}")
if IS_ADMIN and st.sidebar.button("🔎 اختبار اتصال GitHub"):
    try:
        github_api("GET", f"contents/{GITHUB_EXCEL_PATH}?ref={GITHUB_BRANCH}")
        append_activity_log(DB_FILE, "اختبار اتصال GitHub ناجح", f"تم الوصول إلى {GITHUB_REPO}/{GITHUB_EXCEL_PATH}", "متصل")
        st.sidebar.success("الاتصال بـ GitHub يعمل")
    except Exception as connection_error:
        append_activity_log(DB_FILE, "فشل اختبار اتصال GitHub", str(connection_error), "غير متصل أو صلاحيات غير كافية")
        st.sidebar.error(f"تعذر الاتصال بـ GitHub: {connection_error}")

menu_options = {
    "📊 اللوحة الرئيسية": "اللوحة الرئيسية",
    "📝 تسجيل طلب جديد": "تسجيل طلب جديد",
    "🚚 تتبع الشحنات والتحصيل": "تتبع الشحنات والتحصيل",
    "📋 ورقة تجهيز الطلبيات (Packing List)": "ورقة تجهيز الطلبيات",
    "📦 إدارة المخزون والجرد": "إدارة المخزون والجرد",
    "🚚 إدارة شركات الشحن": "إدارة شركات الشحن",
    "📑 تقرير الشحن (شركات الشحن)": "تقرير الشحن",
    "📅 تقرير الطلبات للفترة مع الأصناف": "تقرير الطلبات للفترة مع الأصناف",
    "📊 التعديل الجماعي الشامل (شيت إكسيل)": "التعديل الجماعي الشامل",
    "✏️ تعديل وحذف الفواتير": "تعديل وحذف الفواتير",
    "🖨️ فاتورة العميل": "فاتورة العميل",
    "💸 إدارة المصروفات": "إدارة المصروفات",
    "📈 التقرير المالي التفصيلي": "التقرير المالي التفصيلي",
    "💹 الربحية والأرباح": "الربحية والأرباح",
    "🧾 سجل حركة المستخدم والاتصال": "سجل حركة المستخدم والاتصال",
    "🚨 طلب توريد المخزون الحرِج": "طلب توريد المخزون الحرِج"
}

selected_label = st.sidebar.selectbox("🧭 الانتقال السريع بين الأقسام", list(menu_options.keys()))
menu = menu_options[selected_label]

if st.sidebar.button("⬅️ رجوع"):
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("💾 النسخ الاحتياطي واسترجاع البيانات")

if os.path.exists(DB_FILE):
    with open(DB_FILE, "rb") as f:
        excel_bytes = f.read()
    current_date_str = datetime.date.today().strftime("%Y-%m-%d")
    st.sidebar.download_button(
        label="📥 تحميل نسخة احتياطية (Backup)",
        data=excel_bytes,
        file_name=f"LubaNur_Backup_{current_date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

uploaded_backup = st.sidebar.file_uploader("📤 استرجاع بيانات من نسخة احتياطية (.xlsx)", type=["xlsx"]) if IS_ADMIN else None
if uploaded_backup is not None:
    try:
        with open(DB_FILE, "wb") as f:
            f.write(uploaded_backup.getbuffer())
        st.sidebar.success("✅ تم استرجاع البيانات بنجاح!")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"خطأ: {e}")

if st.sidebar.button("🔒 تسجيل الخروج (Logout)"):
    append_activity_log(DB_FILE, "تسجيل خروج", "ضغط المستخدم زر تسجيل الخروج", "جلسة التطبيق نشطة")
    st.session_state.authenticated = False
    st.session_state.user_role = None
    st.rerun()

st.title("📦 نظام إدارة الطلبيات والمخزون والمصروفات - لبانور / السيف والضياء")
st.markdown("---")

def get_cargo_summary_df(orders_df):
    """ملخص شركة الشحن على أساس قيمة المنتجات فقط؛ رسوم الشحن لا تدخل حسابات الشركة."""
    if orders_df.empty or 'Cargo_Company' not in orders_df.columns:
        return pd.DataFrame()

    df = orders_df.copy()
    for col in ['Products_Total', 'Shipping_Price']:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
    if 'Order_Payment' not in df.columns:
        df['Order_Payment'] = 'لم يتم الدفع'
    if 'Shipping_Payment' not in df.columns:
        df['Shipping_Payment'] = 'لم يتم دفع الشحن'

    summary_list = []
    for company, group in df.groupby('Cargo_Company', dropna=False):
        is_collected = group['Order_Payment'].astype(str).eq('تم الدفع')
        products_sales = float(group['Products_Total'].sum())
        collected_products = float(group.loc[is_collected, 'Products_Total'].sum())
        uncollected_products = float(group.loc[~is_collected, 'Products_Total'].sum())
        # رسوم الشحن معلومات تشغيلية تخص شركة الشحن، وليست إيراداً أو تحصيلاً للشركة.
        shipping_collected = float(group.loc[group['Shipping_Payment'].astype(str).eq('تم دفع الشحن'), 'Shipping_Price'].sum())
        shipping_pending = float(group.loc[~group['Shipping_Payment'].astype(str).eq('تم دفع الشحن'), 'Shipping_Price'].sum())
        summary_list.append({
            'شركة الشحن': str(company) if pd.notna(company) else 'غير محدد',
            'عدد الطلبات': len(group),
            'مبيعات المنتجات (ر.ع)': products_sales,
            'المحصل من مبيعات المنتجات (ر.ع)': collected_products,
            'المبيعات غير المحصلة (ر.ع)': uncollected_products,
            'رسوم الشحن المحصلة للشركة (معلومة فقط)': shipping_collected,
            'رسوم الشحن غير المحصلة للشركة (معلومة فقط)': shipping_pending,
        })
    return pd.DataFrame(summary_list)

if menu == "اللوحة الرئيسية":
    st.header("📊 الملخص التشغيلي والمالي الشامل")
    orders_df, stock_df, expenses_df, _ = load_data()
    
    total_products_sales = pd.to_numeric(orders_df['Products_Total'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    total_shipping = pd.to_numeric(orders_df['Shipping_Price'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    total_expenses = pd.to_numeric(expenses_df['Amount'], errors='coerce').fillna(0).sum() if not expenses_df.empty else 0.0

    paid_mask = orders_df['Order_Payment'].astype(str).eq('تم الدفع') if not orders_df.empty else pd.Series(dtype=bool)
    collected_products_total = pd.to_numeric(orders_df.loc[paid_mask, 'Products_Total'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    pending_products_total = pd.to_numeric(orders_df.loc[~paid_mask, 'Products_Total'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    # رسوم الشحن تُدفع مباشرة لشركة الشحن؛ لا تُضاف إلى إيرادات أو تحصيلات أو خزينة لبانور.
    total_collected_cash = collected_products_total
    actual_treasury = total_collected_cash - total_expenses
    shipping_paid = pd.to_numeric(orders_df.loc[orders_df['Shipping_Payment'].astype(str).eq('تم دفع الشحن'), 'Shipping_Price'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    shipping_unpaid = total_shipping - shipping_paid

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("إجمالي مبيعات المنتجات (بدون الشحن)", f"{total_products_sales:.3f} ر.ع")
    c2.metric("المحصل من مبيعات المنتجات فقط", f"{total_collected_cash:.3f} ر.ع")
    c3.metric("إجمالي المصروفات المسجلة", f"{total_expenses:.3f} ر.ع")
    c4.metric("💰 صافي التدفق النقدي (بدون الشحن)", f"{actual_treasury:.3f} ر.ع")

    st.markdown("---")
    sc1, sc2 = st.columns(2)
    sc1.metric("⏳ المبيعات غير المحصلة لدى شركات الشحن", f"{pending_products_total:.3f} ر.ع")
    sc2.metric("🚚 رسوم الشحن الخاصة بشركات الشحن (خارج الحسابات)", f"مدفوع: {shipping_paid:.3f} | متبقٍ: {shipping_unpaid:.3f} ر.ع", f"الإجمالي: {total_shipping:.3f} ر.ع — لا يدخل في حسابات لبانور")

    st.markdown("---")
    st.subheader("📋 ملخص مبيعات ومستحقات المنتجات لدى شركات الشحن")
    cargo_summary = get_cargo_summary_df(orders_df)
    if not cargo_summary.empty:
        st.dataframe(cargo_summary, use_container_width=True)

elif menu == "تسجيل طلب جديد":
    st.header("📝 إنشاء فاتورة جديدة (إضافة عدة منتجات في نفس الطلب)")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()

    if 'cart' not in st.session_state:
        st.session_state.cart = []

    next_order_id = int(orders_df['Order_ID'].max() + 1) if not orders_df.empty and not pd.isna(orders_df['Order_ID'].max()) else 1
    st.info(f"🔢 **رقم الطلب الجديد (تلقائي):** #{next_order_id}")

    cust_name = st.text_input("اسم العميل (Customer Name)", key="new_cust_name")
    
    col_inf1, col_inf2, col_inf3 = st.columns(3)
    with col_inf1:
        phone = st.text_input("رقم الهاتف (Phone Number)", key="new_phone")
        region = st.text_input("المنطقة / الولاية (Region)", key="new_region")
    with col_inf2:
        order_date = st.date_input("تاريخ الطلب (Date)", datetime.date.today(), key="new_date")
        cargo_company = st.selectbox("شركة الشحن (Cargo Company)", cargo_companies_list, key="new_cargo")
        delivery_type = st.selectbox("نوع التسليم (Delivery Type)", ["من المكتب (Pickup)", "إلى المنزل (Home Delivery)", "بدون"], key="new_dtype")
    with col_inf3:
        shipping_price = st.number_input("سعر الشحن (Shipping Price OMR)", min_value=0.0, value=1.0, key="new_ship_price")
        order_payment = st.selectbox("حالة دفع الطلب (Order Payment Status)", ["لم يتم الدفع", "تم الدفع"], key="new_ord_pay")
        shipping_payment = st.selectbox("حالة دفع قيمة الشحن (Shipping Payment Status)", ["لم يتم دفع الشحن", "تم دفع الشحن"], key="new_shp_pay")
        bags = st.number_input("عدد الأكياس (Bags)", min_value=0, value=1, key="new_bags")

    st.markdown("---")
    st.subheader("2️⃣ إضافة منتجات الفاتورة (مع عرض المخزون الحالي)")
    
    with st.expander("📦 عرض رصيد المخزون الحالي لكل الأصناف"):
        st.dataframe(stock_df, use_container_width=True)

    col_item1, col_item2, col_item3 = st.columns([3, 1, 1])
    with col_item1:
        sel_item = st.selectbox("اختر المنتج", stock_df['Item_Name'].tolist(), key="cart_sel_item")
    with col_item2:
        sel_qty = st.number_input("الكمية", min_value=1, value=1, key="cart_sel_qty")
    with col_item3:
        st.write("")
        st.write("")
        if st.button("➕ إضافة للسلة"):
            st.session_state.cart.append({"Item": sel_item, "Qty": sel_qty})
            st.success("تمت الإضافة!")
            st.rerun()

    if st.session_state.cart:
        st.markdown("🛒 **محتويات السلة الحالية:**")
        cart_df = pd.DataFrame(st.session_state.cart)
        for idx, row in cart_df.iterrows():
            c_c1, c_c2, c_c3 = st.columns([3, 1, 1])
            c_c1.write(f"• {row['Item']}")
            c_c2.write(f"الكمية: {row['Qty']}")
            if c_c3.button("🗑️ حذف", key=f"del_cart_{idx}"):
                st.session_state.cart.pop(idx)
                st.rerun()

        products_total = st.number_input("إجمالي قيمة المنتجات (Products Total OMR)", min_value=0.0, value=5.0, key="new_prod_total")
        
        if st.button("✅ حفظ واعتماد الطلب نهائياً وخصم المخزون"):
            if not cust_name:
                st.error("الرجاء إدخال اسم العميل.")
            elif not st.session_state.cart:
                st.error("الرجاء إضافة صنف واحد على الأقل إلى الطلب.")
            else:
                required = item_quantities(st.session_state.cart)
                stock_issues = []
                for item_name, needed_qty in required.items():
                    rows = stock_df.loc[stock_df["Item_Name"].astype(str) == item_name, "Qty"]
                    available_value = pd.to_numeric(rows.iloc[0], errors="coerce") if not rows.empty else 0
                    available_qty = int(available_value) if pd.notna(available_value) else 0
                    if rows.empty or available_qty < needed_qty:
                        stock_issues.append(f"{item_name}: المتاح {available_qty}، المطلوب {needed_qty}")
                if stock_issues:
                    st.error("لا يمكن حفظ الطلب بسبب نقص المخزون: " + "؛ ".join(stock_issues))
                else:
                    items_summary_list = [f"{r['Item']} (الكمية: {r['Qty']})" for r in st.session_state.cart]
                    items_detail_str = " | ".join(items_summary_list)
                    grand_total = products_total + shipping_price
                    # Product collection is independent from shipping fee payment.
                    coll_status = "مدفوع بالكامل" if order_payment == "تم الدفع" else "قيد التحصيل"
                
                    new_row = {
                        "Order_ID": next_order_id, "Date": str(order_date), "Customer": cust_name, "Phone": phone, "Region": region,
                        "Items_Detail": items_detail_str, "Products_Total": products_total, "Shipping_Price": shipping_price, "Total": grand_total,
                        "Order_Payment": order_payment, "Shipping_Payment": shipping_payment, "Cargo_Company": cargo_company, "Delivery_Type": delivery_type,
                        "Delivery_Status": "قيد التحضير", "Collection_Status": coll_status, "Bags": bags, "Cargo_Barcode": ""
                    }
                
                    for r in st.session_state.cart:
                        it_name = r['Item']
                        it_qty = r['Qty']
                        if it_name in stock_df['Item_Name'].values:
                            idx_st = stock_df[stock_df['Item_Name'] == it_name].index[0]
                            stock_df.loc[idx_st, 'Qty'] = int(stock_df.loc[idx_st, 'Qty']) - int(it_qty)

                    orders_df = pd.concat([orders_df, pd.DataFrame([new_row])], ignore_index=True)
                
                    cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                    save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                
                    st.session_state.cart = []
                    st.session_state.just_created_order = new_row
                    st.cache_data.clear()
                    st.success(f"🎉 تم حفظ الطلب رقم #{next_order_id} بنجاح!")

    if "just_created_order" in st.session_state and st.session_state.just_created_order:
        st.markdown("---")
        st.subheader("🖨️ فاتورة الطلب الجديد (جاهزة للطباعة والنسخ والتحميل):")
        created_row = st.session_state.just_created_order
        st.components.v1.html(render_invoice_snippet(created_row), height=740)

elif menu == "تتبع الشحنات والتحصيل":
    st.header("🚚 تقرير تتبع الشحنات والتحصيل ومتابعة التوصيل")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    
    if not orders_df.empty:
        st.markdown("### 🔎 خيارات البحث والفلترة المتقدمة:")
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            search_by = st.selectbox("البحث والتصفية حسب:", ["عرض كافة الشحنات", "التاريخ المحدد", "شركة الشحن", "رقم الهاتف / العميل", "كود الطلبية (رقم الفاتورة)", "باركود شركة الشحن"])
        with col_f2:
            status_filter = st.selectbox("فلترة حسب حالة التوصيل:", ["الكل", "قيد التحضير", "جاهز للتسليم", "تم التسليم لشركة الشحن", "جاري التوصيل", "تم تسليم العميل", "لم يتم التوصيل", "مرتجع"])
        with col_f3:
            collection_filter = st.selectbox("فلترة حسب حالة التحصيل المالي:", ["الكل", "قيد التحصيل", "مدفوع بالكامل"])
            
        filtered_df = orders_df.copy()
        
        if search_by == "التاريخ المحدد":
            filter_date = st.date_input("اختر التاريخ المطلوب", datetime.date.today())
            filtered_df['Date_Parsed'] = pd.to_datetime(filtered_df['Date'], errors='coerce').dt.date
            filtered_df = filtered_df[filtered_df['Date_Parsed'] == filter_date]
            
        elif search_by == "شركة الشحن":
            sel_cargo = st.selectbox("اختر شركة الشحن", cargo_companies_list)
            filtered_df = filtered_df[filtered_df['Cargo_Company'] == sel_cargo]
            
            use_cargo_date = st.checkbox("📅 تصفية شركة الشحن بتاريخ محدد (اختياري)", value=False)
            if use_cargo_date:
                cargo_date = st.date_input("اختر التاريخ لشركة الشحن", datetime.date.today())
                filtered_df['Date_Parsed'] = pd.to_datetime(filtered_df['Date'], errors='coerce').dt.date
                filtered_df = filtered_df[filtered_df['Date_Parsed'] == cargo_date]
                
        elif search_by == "رقم الهاتف / العميل":
            phone_q = st.text_input("أدخل رقم الهاتف أو اسم العميل للبحث:")
            if phone_q:
                q_clean = phone_q.strip().lower()
                filtered_df = filtered_df[
                    filtered_df['Phone'].astype(str).str.lower().str.contains(q_clean) | 
                    filtered_df['Customer'].astype(str).str.lower().str.contains(q_clean)
                ]
        elif search_by == "كود الطلبية (رقم الفاتورة)":
            id_q = st.number_input("أدخل رقم الطلب (Order ID):", min_value=0, value=0)
            if id_q > 0:
                filtered_df = filtered_df[filtered_df['Order_ID'] == int(id_q)]
        elif search_by == "باركود شركة الشحن":
            barcode_q = st.text_input("أدخل باركود الشحن:")
            if barcode_q:
                filtered_df = filtered_df[filtered_df['Cargo_Barcode'].astype(str).str.contains(barcode_q.strip(), na=False)]
                
        if status_filter != "الكل":
            filtered_df = filtered_df[filtered_df['Delivery_Status'] == status_filter]
        if collection_filter != "الكل":
            # The collection state is based on product/invoice payment only.
            filtered_df = filtered_df[filtered_df['Order_Payment'].map(lambda v: 'مدفوع بالكامل' if str(v).strip() == 'تم الدفع' else 'قيد التحصيل') == collection_filter]
            
        filtered_df = filtered_df.copy()
        filtered_df['Collection_Status'] = filtered_df['Order_Payment'].map(lambda v: 'مدفوع بالكامل' if str(v).strip() == 'تم الدفع' else 'قيد التحصيل')
        total_due_amt = filtered_df[filtered_df['Order_Payment'] != 'تم الدفع']['Products_Total'].sum()
        total_undelivered = len(filtered_df[filtered_df['Delivery_Status'] != 'تم تسليم العميل'])
        
        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        m1.metric("عدد الشحنات الظاهرة", len(filtered_df))
        m2.metric("إجمالي المبالغ المستحقة (منتجات فقط)", f"{total_due_amt:.3f} ر.ع")
        m3.metric("الطلبيات التي لم تسلم بعد", total_undelivered)
        
        st.markdown("---")
        st.subheader("📋 جدول الشحنات والطلبيات المطابقة:")
        
        if not filtered_df.empty:
            column_labels = {
                "رقم الفاتورة": "Order_ID", "رقم ID": "Order_ID", "التاريخ": "Date",
                "اسم العميل": "Customer", "رقم هاتف العميل": "Phone", "المنطقة": "Region",
                "شركة الشحن": "Cargo_Company", "رقم الباركود": "Cargo_Barcode", "عدد الأكياس": "Bags",
                "حالة الدفع": "Order_Payment", "حالة الشحن والتوصيل": "Delivery_Status",
                "حالة التحصيل": "Collection_Status", "قيمة المنتجات": "Products_Total",
                "قيمة الشحن": "Shipping_Price", "الإجمالي شامل الشحن": "Total"
            }
            default_visible = ["رقم الفاتورة", "رقم ID", "رقم الباركود", "حالة الدفع", "حالة الشحن والتوصيل", "قيمة المنتجات", "رقم هاتف العميل"]
            selected_labels = st.multiselect(
                "⚙️ تخصيص أعمدة جدول الشحنات (اختر ما تريد عرضه)",
                options=list(column_labels.keys()), default=default_visible,
                key="tracking_visible_columns"
            )
            if not selected_labels:
                selected_labels = ["رقم الفاتورة"]
            display_df = pd.DataFrame(index=filtered_df.index)
            for label in selected_labels:
                source_col = column_labels[label]
                display_df[label] = filtered_df[source_col]
            event = st.dataframe(
                display_df,
                use_container_width=True,
                selection_mode="single-row",
                on_select="rerun",
                key="orders_grid_select"
            )
            
            selected_rows = event.selection.rows if hasattr(event, 'selection') else []
            
            # تحديد الطلبية المختارة إما عن طريق النقر في الجدول أو القائمة المنسدلة
            st.markdown("---")
            st.subheader("✏️ قسم التعديل السريع لحالة الطلبية (التوصيل، الدفع، والشحن):")
            
            order_options_map = {row['Order_ID']: f"طلب #{row['Order_ID']} - العميل: {row['Customer']} (هاتف: {row['Phone']})" for idx, row in filtered_df.iterrows()}
            
            default_index = 0
            if selected_rows and len(selected_rows) > 0:
                clicked_row_idx = selected_rows[0]
                if clicked_row_idx < len(filtered_df):
                    clicked_order_id = filtered_df.iloc[clicked_row_idx]['Order_ID']
                    if clicked_order_id in list(order_options_map.keys()):
                        default_index = list(order_options_map.keys()).index(clicked_order_id)

            chosen_order_id = st.selectbox("اختر رقم الطلبية للتعديل الفوري:", list(order_options_map.keys()), index=default_index, format_func=lambda x: order_options_map[x])
            
            target_order_row = orders_df[orders_df['Order_ID'] == chosen_order_id].iloc[0]
            orig_idx = orders_df[orders_df['Order_ID'] == chosen_order_id].index[0]
            
            with st.form("quick_update_order_form"):
                st.info(f"📌 الطلب المختار: #{chosen_order_id} | العميل: {target_order_row['Customer']} | قيمة المنتجات: {target_order_row['Products_Total']} ر.ع")
                
                q_col1, q_col2, q_col3 = st.columns(3)
                
                delivery_status_options = ["قيد التحضير", "جاهز للتسليم", "تم التسليم لشركة الشحن", "جاري التوصيل", "تم تسليم العميل", "لم يتم التوصيل", "مرتجع"]
                cur_d_stat = str(target_order_row.get('Delivery_Status', 'قيد التحضير'))
                d_i = delivery_status_options.index(cur_d_stat) if cur_d_stat in delivery_status_options else 0
                new_d_stat = q_col1.selectbox("حالة التوصيل", delivery_status_options, index=d_i)
                
                order_pay_options = ["لم يتم الدفع", "تم الدفع"]
                cur_o_pay = str(target_order_row.get('Order_Payment', 'لم يتم الدفع'))
                o_i = order_pay_options.index(cur_o_pay) if cur_o_pay in order_pay_options else 0
                new_o_pay = q_col2.selectbox("دفع الطلب", order_pay_options, index=o_i)
                
                shipping_pay_options = ["لم يتم دفع الشحن", "تم دفع الشحن"]
                cur_s_pay = str(target_order_row.get('Shipping_Payment', 'لم يتم دفع الشحن'))
                s_i = shipping_pay_options.index(cur_s_pay) if cur_s_pay in shipping_pay_options else 0
                new_s_pay = q_col3.selectbox("دفع الشحن", shipping_pay_options, index=s_i)
                
                st.caption("حالة التحصيل تُحدد تلقائياً من حالة دفع المنتجات، ولا تتأثر بحالة دفع الشحن.")
                
                new_barcode_val = st.text_input("تعديل باركود شركة الشحن", value=str(target_order_row.get('Cargo_Barcode', '')))
                
                if st.form_submit_button("💾 حفظ وتحديث حالة الطلبية فوراً"):
                    orders_df.loc[orig_idx, 'Delivery_Status'] = new_d_stat
                    orders_df.loc[orig_idx, 'Order_Payment'] = new_o_pay
                    orders_df.loc[orig_idx, 'Shipping_Payment'] = new_s_pay
                    orders_df.loc[orig_idx, 'Collection_Status'] = 'مدفوع بالكامل' if new_o_pay == 'تم الدفع' else 'قيد التحصيل'
                    orders_df.loc[orig_idx, 'Cargo_Barcode'] = new_barcode_val
                    
                    cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                    save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                    
                    st.cache_data.clear()
                    st.success(f"✅ تم تحديث الطلب رقم #{chosen_order_id} بنجاح!")
                    st.rerun()
        else:
            st.warning("لا توجد طلبيات مطابقة لمعايير البحث المحددة.")
    else:
        st.info("لا توجد طلبيات مسجلة في النظام بعد.")

elif menu == "ورقة تجهيز الطلبيات":
    st.header("📋 ورقة تجهيز الطلبيات التفاعلية (Packing & Picking List)")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    
    if not orders_df.empty:
        st.markdown("📅 **اختر تاريخ الطلبات لتجهيزها:**")
        pick_date = st.date_input("تاريخ التجهيز", datetime.date.today(), key="packing_date")
        
        pick_df = orders_df.copy()
        pick_df['Date_Parsed'] = pd.to_datetime(pick_df['Date'], errors='coerce').dt.date
        
        unprocessed_orders = pick_df[(pick_df['Date_Parsed'] == pick_date) & (~pick_df['Delivery_Status'].isin(["جاهز للتسليم", "تم التسليم لشركة الشحن", "تم تسليم العميل"]))]
        
        if not unprocessed_orders.empty:
            st.info(f"الطلبيات المتبقية قيد التجهيز في تاريخ {pick_date}: {len(unprocessed_orders)}")
            
            for idx, order in unprocessed_orders.iterrows():
                order_id = order['Order_ID']
                cust = order['Customer']
                phone = order['Phone']
                region = order['Region']
                current_d_status = order['Delivery_Status']
                order_date_str = str(order['Date'])
                total_amount = order['Total']
                delivery_type = order.get('Delivery_Type', 'غير محدد')
                current_barcode = str(order.get('Cargo_Barcode', ''))
                
                with st.expander(f"📦 طلب رقم #{order_id} | العميل: {cust} | المبلغ: {total_amount} ر.ع | الحالة: {current_d_status}"):
                    
                    col_info_d1, col_info_d2 = st.columns(2)
                    with col_info_d1:
                        st.markdown(f"👤 **اسم العميل:** {cust}")
                        st.markdown(f"📞 **الهاتف:** <span style='direction: ltr; display: inline-block;'>{phone}</span>", unsafe_allow_html=True)
                        st.markdown(f"📍 **المنطقة:** {region}")
                    with col_info_d2:
                        st.markdown(f"📅 **تاريخ الطلب:** <span style='direction: ltr; display: inline-block;'>{order_date_str}</span>", unsafe_allow_html=True)
                        st.markdown(f"🚚 **مكان التسليم (نوع التسليم):** {delivery_type}")
                        st.markdown(f"💰 **المبلغ الإجمالي:** {total_amount} ر.ع (المنتجات: {order['Products_Total']} + الشحن: {order['Shipping_Price']})")
                    
                    st.markdown(f"🏷️ **عدد الأكياس:** {order['Bags']} &nbsp;|&nbsp; 📦 **شركة الشحن:** {order['Cargo_Company']}")
                    
                    st.markdown("---")
                    st.markdown("🔍 **إدخال أو مسح باركود شركة الشحن:**")
                    with st.form(key=f"barcode_form_{order_id}"):
                        new_bc_input = st.text_input("باركود الشحن (Barcode)", value=current_barcode, key=f"bc_field_{order_id}")
                        if st.form_submit_button("💾 حفظ الباركود"):
                            ord_idx_target = orders_df[orders_df['Order_ID'] == order_id].index[0]
                            orders_df.loc[ord_idx_target, 'Cargo_Barcode'] = new_bc_input
                            
                            cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                            save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                            st.cache_data.clear()
                            st.success(f"✅ تم حفظ الباركود للطلب #{order_id} بنجاح!")
                            st.rerun()

                    st.markdown("---")
                    st.write("✅ **أصناف الطلبية (ضع علامة صح عند تجهيز كل صنف):**")
                    
                    parsed_items = parse_items_to_list(order['Items_Detail'])
                    
                    packed_all = True
                    for i, item_dict in enumerate(parsed_items):
                        iname = item_dict['Item']
                        iqty = item_dict['Qty']
                        
                        col_chk1, col_chk2 = st.columns([4, 1])
                        with col_chk1:
                            is_checked = st.checkbox(f"{iname} - الكمية المطلوبة: {iqty}", value=False, key=f"pack_{order_id}_{i}")
                        with col_chk2:
                            if is_checked:
                                st.markdown("✅ **تم تجهيزه**")
                            else:
                                packed_all = False
                                st.markdown("⏳ **قيد التجهيز**")
                    
                    st.markdown("---")
                    if packed_all:
                        st.success("🎉 تم تجهيز كافة الأصناف بالكامل!")
                        if st.button(f"🚀 اعتماد تجهيز الطلب #{order_id} وتحويله لـ (جاهز للتسليم)", key=f"btn_ready_{order_id}"):
                            idx_ord = orders_df[orders_df['Order_ID'] == order_id].index[0]
                            orders_df.loc[idx_ord, 'Delivery_Status'] = "جاهز للتسليم"
                            
                            cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                            save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                            
                            st.cache_data.clear()
                            st.success(f"✅ تم تحويل الطلب #{order_id} إلى (جاهز للتسليم) وإخفاؤه من الشاشة!")
                            st.rerun()
                    else:
                        st.info("💡 يرجى وضع علامة صح على كافة الأصناف للتحويل الفوري للطلب.")
        else:
            st.success("🎉 تم تجهيز كافة طلبات هذا اليوم بنجاح أو لا توجد طلبات معلقة!")
    else:
        st.info("لا توجد طلبيات مسجلة في النظام بعد.")

elif menu == "إدارة المخزون والجرد":
    st.header("📦 إدارة المخزون، تعزيز الأصناف وتسوية الجرد")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    
    st.subheader("📋 تقرير الأصناف والرصيد الحالي:")
    if not stock_df.empty:
        stock_display = stock_df.copy()
        stock_display['الحالة'] = stock_display['Qty'].apply(lambda x: "🚨 حرِج (أقل من 5)" if x < 5 else "✅ متوفر")
        st.dataframe(stock_display, use_container_width=True)
    
    st.markdown("---")
    tab1, tab2 = st.tabs(["➕ تعزيز / سحب كميات المنتجات", "📊 تسوية المخزون بعد الجرد (العجز والزيادة)"])
    
    with tab1:
        st.subheader("تعزيز أو سحب أصناف يدوياً")
        with st.form("restock_form"):
            r_item = st.selectbox("اختر المنتج", stock_df['Item_Name'].tolist(), key="rst_item")
            r_action = st.selectbox("نوع العملية", ["إضافة إلى المخزون (تعزيز)", "سحب من المخزون"])
            r_qty = st.number_input("الكمية", min_value=1, value=1)
            
            if st.form_submit_button("💾 تنفيذ وتحديث المخزون"):
                idx_st = stock_df[stock_df['Item_Name'] == r_item].index[0]
                cur_q = int(stock_df.loc[idx_st, 'Qty'])
                if r_action == "إضافة إلى المخزون (تعزيز)":
                    stock_df.loc[idx_st, 'Qty'] = cur_q + r_qty
                else:
                    stock_df.loc[idx_st, 'Qty'] = max(0, cur_q - r_qty)
                
                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                
                st.cache_data.clear()
                st.success(f"✅ تم تحديث مخزون المنتج ({r_item}) بنجاح!")
                st.rerun()
                
    with tab2:
        st.subheader("تسوية المخزون بعد الجرد الفعلي")
        if not stock_df.empty:
            audit_df = stock_df.copy()
            audit_df['الكمية الفعلية بالجرد'] = audit_df['Qty']
            
            edited_audit = st.data_editor(audit_df[['Item_Name', 'Qty', 'الكمية الفعلية بالجرد']], use_container_width=True, key="audit_editor")
            
            if st.button("⚖️ اعتماد تسوية الجرد وتحديث الأرصدة"):
                for idx, row in edited_audit.iterrows():
                    it_name = row['Item_Name']
                    actual_q = int(row['الكمية الفعلية بالجرد'])
                    stock_df.loc[stock_df['Item_Name'] == it_name, 'Qty'] = actual_q
                
                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                
                st.cache_data.clear()
                st.success("🎉 تمت تسوية المخزون واعتماد الجرد بنجاح!")
                st.rerun()

elif menu == "إدارة شركات الشحن":
    st.header("🚚 إدارة وإضافة شركات الشحن")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    
    st.subheader("الشركات الحالية في النظام:")
    st.write(cargo_companies_list)
    
    with st.form("add_cargo_form"):
        new_cargo_name = st.text_input("اسم شركة الشحن الجديدة")
        if st.form_submit_button("➕ إضافة شركة الشحن"):
            if new_cargo_name and new_cargo_name not in cargo_companies_list:
                cargo_companies_list.append(new_cargo_name.strip())
                
                if "وصل لي" in cargo_companies_list:
                    cargo_companies_list.remove("وصل لي")
                    cargo_companies_list.insert(0, "وصل لي")
                
                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                
                st.cache_data.clear()
                st.success(f"✅ تمت إضافة شركة الشحن '{new_cargo_name}' بنجاح!")
                st.rerun()
            else:
                st.error("الاسم فارغ أو موجود مسبقاً.")

elif menu == "تقرير الشحن":
    st.header("📑 تقرير تسليم الشحنات لشركات الشحن")
    orders_df, _, _, cargo_companies_list = load_data()
    
    if not orders_df.empty:
        col_rep_c1, col_rep_c2, col_rep_c3 = st.columns(3)
        with col_rep_c1:
            selected_cargo = st.selectbox("اختر شركة الشحن", cargo_companies_list, key="rep_cargo_sel")
        with col_rep_c2:
            use_date_filter = st.checkbox("تصفية حسب التاريخ المحدد (اختياري)", value=False)
        with col_rep_c3:
            if use_date_filter:
                report_date = st.date_input("اختر التاريخ", datetime.date.today(), key="rep_date_sel")
            else:
                report_date = None

        filtered_cargo_df = orders_df[orders_df['Cargo_Company'] == selected_cargo].copy()
        
        if use_date_filter and report_date:
            filtered_cargo_df['Date_Parsed'] = pd.to_datetime(filtered_cargo_df['Date'], errors='coerce').dt.date
            filtered_cargo_df = filtered_cargo_df[filtered_cargo_df['Date_Parsed'] == report_date]

        if not filtered_cargo_df.empty:
            report_display_df = pd.DataFrame({
                "رقم الفاتورة": filtered_cargo_df['Order_ID'],
                "التاريخ": filtered_cargo_df['Date'],
                "اسم العميل": filtered_cargo_df['Customer'],
                "رقم الهاتف": filtered_cargo_df['Phone'],
                "مكان التوصيل (المنطقة)": filtered_cargo_df['Region'],
                "عدد الأكياس": filtered_cargo_df['Bags'],
                "حالة تسليم العميل": filtered_cargo_df['Delivery_Status'],
                "قيمة مبيعات المنتجات (ر.ع)": filtered_cargo_df['Products_Total'],
                "حالة تحصيل المنتجات": filtered_cargo_df['Order_Payment'],
                "المحصل من المنتجات (ر.ع)": filtered_cargo_df.apply(lambda r: r['Products_Total'] if str(r['Order_Payment']) == 'تم الدفع' else 0, axis=1),
                "المتبقي من المنتجات (ر.ع)": filtered_cargo_df.apply(lambda r: 0 if str(r['Order_Payment']) == 'تم الدفع' else r['Products_Total'], axis=1),
                "رسوم الشحن (للشركة مباشرة - خارج إيرادات لبانور)": filtered_cargo_df['Shipping_Price'],
                "حالة دفع الشحن": filtered_cargo_df['Shipping_Payment'],
                "باركود الشحن": filtered_cargo_df['Cargo_Barcode'],
            })

            st.markdown(f"### 📋 بيان شحنات شركة: **{selected_cargo}** (إجمالي الشحنات: {len(report_display_df)})")
            st.caption("حالة تسليم العميل موضحة لكل فاتورة. رسوم الشحن تخص شركة الشحن فقط، وتُعرض للمطابقة التشغيلية ولا تدخل في إيرادات لبانور أو تحصيل المنتجات.")
            # Widths are adjustable per column in the expander, including on touch screens.
            report_widths = {}
            width_labels = {"ضيق": "small", "متوسط": "medium", "واسع": "large"}
            with st.expander("⚙️ التحكم في عرض أعمدة تقرير الشحن (مناسب للـ iPad)"):
                st.caption("اختر عرض كل عمود، ثم أغلق القائمة لعرض التقرير. يمكن التمرير أفقياً لبقية الأعمدة.")
                report_col_names = list(report_display_df.columns)
                for pair_start in range(0, len(report_col_names), 2):
                    width_cols = st.columns(2)
                    for slot, col_name in enumerate(report_col_names[pair_start:pair_start + 2]):
                        with width_cols[slot]:
                            width_choice = st.selectbox(
                                f"عرض: {col_name}",
                                options=list(width_labels.keys()),
                                index=1,
                                key=f"cargo_report_width_{pair_start}_{slot}"
                            )
                            report_widths[col_name] = width_labels[width_choice]
            report_column_config = {
                col_name: st.column_config.Column(width=report_widths.get(col_name, "medium"))
                for col_name in report_display_df.columns
            }
            st.dataframe(
                report_display_df,
                use_container_width=True,
                hide_index=True,
                column_config=report_column_config
            )
            products_sales = pd.to_numeric(filtered_cargo_df['Products_Total'], errors='coerce').fillna(0).sum()
            products_collected = pd.to_numeric(filtered_cargo_df.loc[filtered_cargo_df['Order_Payment'].astype(str).eq('تم الدفع'), 'Products_Total'], errors='coerce').fillna(0).sum()
            products_pending = products_sales - products_collected
            st.markdown("#### ملخص حساب المنتجات لدى شركة الشحن")
            k1, k2, k3 = st.columns(3)
            k1.metric("مبيعات المنتجات", f"{products_sales:.3f} ر.ع")
            k2.metric("المحصل", f"{products_collected:.3f} ر.ع")
            k3.metric("المبيعات غير المحصلة / المتبقي", f"{products_pending:.3f} ر.ع")
            st.caption("رسوم الشحن معروضة للمراجعة التشغيلية فقط؛ فهي تخص شركة الشحن ولا تدخل في إيرادات لبانور أو تحصيل المنتجات. إجمالي الفاتورة قد يعرضها للعميل كبند منفصل، لكن التقارير المالية تحتسب قيمة المنتجات فقط.")
        else:
            st.warning("لا توجد شحنات مطابقة لخيارات البحث المحددة.")
    else:
        st.info("لا توجد طلبيات مسجلة في النظام بعد.")

elif menu == "تقرير الطلبات للفترة مع الأصناف":
    st.header("📅 تقرير الطلبات للفترة مع الأصناف")
    orders_df, _, _, _ = load_data()
    if not orders_df.empty:
        col_dt1, col_dt2 = st.columns(2)
        with col_dt1:
            start_date = st.date_input("من تاريخ", datetime.date.today() - datetime.timedelta(days=30), key="rep_start_date")
        with col_dt2:
            end_date = st.date_input("إلى تاريخ", datetime.date.today(), key="rep_end_date")
            
        filtered_period_df = orders_df.copy()
        filtered_period_df['Date_Parsed'] = pd.to_datetime(filtered_period_df['Date'], errors='coerce').dt.date
        
        mask = (filtered_period_df['Date_Parsed'] >= start_date) & (filtered_period_df['Date_Parsed'] <= end_date)
        period_result = filtered_period_df[mask]
        
        st.info(f"عدد الطلبات في الفترة المحددة: {len(period_result)}")
        st.dataframe(period_result[['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Company', 'Items_Detail', 'Products_Total', 'Shipping_Price', 'Total']], use_container_width=True)
    else:
        st.info("لا توجد طلبيات مسجلة في النظام بعد.")

elif menu == "التعديل الجماعي الشامل":
    st.header("📊 التعديل الجماعي الفوري للطلبات (شيت إكسيل تفاعلي)")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    if not orders_df.empty:
        edited_orders_df = st.data_editor(orders_df, num_rows="dynamic", use_container_width=True, key="batch_order_editor")
        if st.button("💾 حفظ جميع التعديلات دفعة واحدة"):
            try:
                original_items = orders_df.set_index("Order_ID")["Items_Detail"].astype(str).to_dict()
                edited_items = edited_orders_df.set_index("Order_ID")["Items_Detail"].astype(str).to_dict()
                changed_item_orders = [oid for oid, value in edited_items.items() if oid in original_items and value != original_items[oid]]
                if changed_item_orders:
                    st.error("لم يتم الحفظ: تعديل أصناف الفاتورة من الجدول الجماعي قد يسبب اختلاف المخزون. استخدم قسم «تعديل وحذف الفواتير» لتعديل الأصناف مع تحديث المخزون تلقائيًا.")
                    st.stop()
                edited_orders_df['Total'] = pd.to_numeric(edited_orders_df['Products_Total'], errors='coerce').fillna(0) + pd.to_numeric(edited_orders_df['Shipping_Price'], errors='coerce').fillna(0)
                save_all_data(edited_orders_df, stock_df, expenses_df, cargo_companies_list)
                st.success("🎉 تم حفظ جميع التعديلات وتحديث قاعدة البيانات بنجاح!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ حدث خطأ: {e}")
    else:
        st.info("لا توجد طلبيات.")

elif menu == "فاتورة العميل":
    st.header("🖨️ فاتورة العميل (معاينة وطباعة ونسخ وتصدير)")
    orders_df, _, _, _ = load_data()
    
    if not orders_df.empty:
        st.markdown("🔍 **خيارات البحث المرنة للوصول للفاتورة:**")
        
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            search_invoice_id = st.number_input("البحث برقم الفاتورة (Order ID):", min_value=0, value=0)
        with col_s2:
            search_phone_cust = st.text_input("البحث برقم الهاتف أو اسم العميل:")
        with col_s3:
            search_barcode = st.text_input("البحث بباركود الشحن:")
            
        use_inv_date = st.checkbox("📅 تصفية حسب التاريخ (اختياري)", value=False)
        
        filtered_inv_df = orders_df.copy()
        
        if use_inv_date:
            receipt_filter_date = st.date_input("اختر التاريخ", datetime.date.today(), key="rec_filter_date")
            filtered_inv_df['Date_Parsed'] = pd.to_datetime(filtered_inv_df['Date'], errors='coerce').dt.date
            filtered_inv_df = filtered_inv_df[filtered_inv_df['Date_Parsed'] == receipt_filter_date]
            
        if search_invoice_id > 0:
            filtered_inv_df = filtered_inv_df[filtered_inv_df['Order_ID'] == int(search_invoice_id)]
        if search_phone_cust:
            q_str = search_phone_cust.strip().lower()
            filtered_inv_df = filtered_inv_df[
                filtered_inv_df['Phone'].astype(str).str.lower().str.contains(q_str) | 
                filtered_inv_df['Customer'].astype(str).str.lower().str.contains(q_str)
            ]
        if search_barcode:
            filtered_inv_df = filtered_inv_df[filtered_inv_df['Cargo_Barcode'].astype(str).str.contains(search_barcode.strip(), na=False)]

        if not filtered_inv_df.empty:
            order_options = filtered_inv_df['Order_ID'].tolist()
            selected_order_id = st.selectbox("اختر الفاتورة المطلوبة من النتائج:", order_options, format_func=lambda x: f"فاتورة #{x} - العميل: {filtered_inv_df.loc[filtered_inv_df['Order_ID']==x, 'Customer'].values[0]} (هاتف: {filtered_inv_df.loc[filtered_inv_df['Order_ID']==x, 'Phone'].values[0]})")
            
            order_row = orders_df[orders_df['Order_ID'] == selected_order_id].iloc[0]
            st.components.v1.html(render_invoice_snippet(order_row), height=740)
        else:
            st.warning("لا توجد فواتير مطابقة لمعايير البحث المحددة.")
    else:
        st.info("لا توجد فواتير مسجلة في النظام بعد.")

elif menu == "تعديل وحذف الفواتير":
    st.header("✏️ تعديل وحذف الفواتير مع إدارة الأصناف بسهولة")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    
    if not orders_df.empty:
        edit_order_id = st.selectbox("اختر رقم الطلب للتعديل", orders_df['Order_ID'].tolist())
        target_row = orders_df[orders_df['Order_ID'] == edit_order_id].iloc[0]
        
        st.info(f"📌 **تعديل الفاتورة رقم #{edit_order_id} للعميل: {target_row['Customer']}**")
        
        with st.form("edit_form"):
            e_cust = st.text_input("اسم العميل", value=str(target_row['Customer']))
            e_phone = st.text_input("الهاتف", value=str(target_row['Phone']))
            e_region = st.text_input("المنطقة / الولاية", value=str(target_row.get('Region', '')))
            
            cur_cargo = str(target_row.get('Cargo_Company', cargo_companies_list[0]))
            cargo_idx = cargo_companies_list.index(cur_cargo) if cur_cargo in cargo_companies_list else 0
            e_cargo = st.selectbox("شركة الشحن", cargo_companies_list, index=cargo_idx)
            
            e_bags = st.number_input("عدد الأكياس", min_value=0, value=int(target_row.get('Bags', 1)))
            e_prod_tot = st.number_input("قيمة المنتجات (OMR)", value=float(target_row['Products_Total']))
            e_ship_p = st.number_input("سعر الشحن (OMR)", value=float(target_row['Shipping_Price']))
            e_barcode = st.text_input("باركود شركة الشحن", value=str(target_row.get('Cargo_Barcode', '')))
            
            cur_ord_pay = str(target_row.get('Order_Payment', 'لم يتم الدفع'))
            cur_shp_pay = str(target_row.get('Shipping_Payment', 'لم يتم دفع الشحن'))
            
            ord_pay_list = ["لم يتم الدفع", "تم الدفع"]
            o_idx = ord_pay_list.index(cur_ord_pay) if cur_ord_pay in ord_pay_list else 0
            e_ord_pay = st.selectbox("حالة دفع الطلب", ord_pay_list, index=o_idx)
            
            shp_pay_list = ["لم يتم دفع الشحن", "تم دفع الشحن"]
            s_idx = shp_pay_list.index(cur_shp_pay) if cur_shp_pay in shp_pay_list else 0
            e_shp_pay = st.selectbox("حالة دفع قيمة الشحن", shp_pay_list, index=s_idx)
            
            submitted_main_edit = st.form_submit_button("💾 حفظ البيانات الأساسية للفاتورة")
            if submitted_main_edit:
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Customer'] = e_cust
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Phone'] = e_phone
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Region'] = e_region
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Cargo_Company'] = e_cargo
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Bags'] = e_bags
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Products_Total'] = e_prod_tot
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Shipping_Price'] = e_ship_p
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Total'] = e_prod_tot + e_ship_p
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Cargo_Barcode'] = e_barcode
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Order_Payment'] = e_ord_pay
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Shipping_Payment'] = e_shp_pay
                
                # Collection of product value depends only on Order_Payment; shipping payment is separate.
                new_coll_status = "مدفوع بالكامل" if e_ord_pay == "تم الدفع" else "قيد التحصيل"
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Collection_Status'] = new_coll_status

                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                st.cache_data.clear()
                st.success("✅ تم تحديث بيانات الفاتورة بنجاح!")
                st.rerun()

        st.markdown("---")
        st.subheader("🛒 إدارة أصناف هذه الفاتورة:")
        
        current_parsed_items = parse_items_to_list(target_row['Items_Detail'])
        
        if current_parsed_items:
            st.markdown("**الأصناف الحالية في الفاتورة:**")
            for idx, item_dict in enumerate(current_parsed_items):
                col_item_info1, col_item_info2 = st.columns([5, 1])
                with col_item_info1:
                    st.write(f"• {item_dict['Item']} &nbsp;|&nbsp; **الكمية:** {item_dict['Qty']}")
                with col_item_info2:
                    if st.button("🗑️ حذف", key=f"del_item_inv_{edit_order_id}_{idx}"):
                        new_items = current_parsed_items.copy()
                        new_items.pop(idx)
                        adjusted_stock, stock_error = adjust_stock_for_order(stock_df, current_parsed_items, new_items)
                        if stock_error:
                            st.error(stock_error)
                            st.stop()
                        stock_df = adjusted_stock
                        current_parsed_items = new_items
                        new_items_str = " | ".join([f"{x['Item']} (الكمية: {x['Qty']})" for x in current_parsed_items]) if current_parsed_items else "لا توجد أصناف"
                        orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Items_Detail'] = new_items_str
                        
                        cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                        save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                        st.cache_data.clear()
                        st.success("✅ تم حذف الصنف بنجاح!")
                        st.rerun()

        st.markdown("---")
        st.markdown("**➕ إضافة منتج جديد للفاتورة:**")
        with st.form("add_item_to_inv_form"):
            add_sel_item = st.selectbox("اختر المنتج", EXACT_ITEMS, key=f"add_item_sel_{edit_order_id}")
            add_sel_qty = st.number_input("الكمية", min_value=1, value=1, key=f"add_item_qty_{edit_order_id}")
            
            if st.form_submit_button("➕ إضافة المنتج إلى الفاتورة"):
                new_items = current_parsed_items.copy()
                new_items.append({"Item": add_sel_item, "Qty": int(add_sel_qty)})
                adjusted_stock, stock_error = adjust_stock_for_order(stock_df, current_parsed_items, new_items)
                if stock_error:
                    st.error(stock_error)
                    st.stop()
                stock_df = adjusted_stock
                current_parsed_items = new_items
                new_items_str = " | ".join([f"{x['Item']} (الكمية: {x['Qty']})" for x in current_parsed_items])
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Items_Detail'] = new_items_str
                
                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
                st.cache_data.clear()
                st.success("✅ تمت إضافة المنتج بنجاح!")
                st.rerun()

        st.markdown("---")
        if st.button("🗑️ حذف الفاتورة بالكامل نهائياً"):
            old_items = parse_items_to_list(target_row["Items_Detail"])
            adjusted_stock, stock_error = adjust_stock_for_order(stock_df, old_items, [])
            if stock_error:
                st.error(stock_error)
                st.stop()
            stock_df = adjusted_stock
            orders_df = orders_df[orders_df['Order_ID'] != edit_order_id]
            cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
            save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
            st.cache_data.clear()
            st.success("🗑️ تم حذف الفاتورة بنجاح!")
            st.rerun()
    else:
        st.info("لا توجد طلبيات مسجلة في النظام بعد.")

elif menu == "إدارة المصروفات":
    st.header("💸 إدارة المصروفات")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    with st.form("add_expense_form"):
        exp_date = st.date_input("التاريخ", datetime.date.today())
        exp_category = st.selectbox("البند", ["إعلانات ميتا (Meta Ads)", "شراء أصناف / مواد خام (Purchases)", "سحب مبالغ مالية / مسحوبات شخصية (Withdrawals)", "مصروفات أخرى (Other Expenses)"])
        exp_amount = st.number_input("المبلغ (OMR)", min_value=0.0, value=10.0)
        exp_desc = st.text_input("الوصف")
        if st.form_submit_button("💾 حفظ المصروف"):
            new_exp_id = int(expenses_df['Expense_ID'].max() + 1) if not expenses_df.empty and not pd.isna(expenses_df['Expense_ID'].max()) else 1
            new_row = {"Expense_ID": new_exp_id, "Date": str(exp_date), "Category": exp_category, "Amount": exp_amount, "Description": exp_desc}
            expenses_df = pd.concat([expenses_df, pd.DataFrame([new_row])], ignore_index=True)
            cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
            save_all_data(orders_df, stock_df, expenses_df, cargo_companies_list)
            st.cache_data.clear()
            st.success("✅ تم حفظ المصروف بنجاح!")
            st.rerun()
    if not expenses_df.empty:
        st.dataframe(expenses_df, use_container_width=True)

elif menu == "التقرير المالي التفصيلي":
    st.header("📈 التقرير المالي التفصيلي والشامل")
    orders_df, stock_df, expenses_df, _ = load_data()
    
    total_products_value = pd.to_numeric(orders_df['Products_Total'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    paid_mask = orders_df['Order_Payment'].astype(str).eq('تم الدفع') if not orders_df.empty else pd.Series(dtype=bool)
    collected_products_total = pd.to_numeric(orders_df.loc[paid_mask, 'Products_Total'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    pending_products_total = pd.to_numeric(orders_df.loc[~paid_mask, 'Products_Total'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    total_shipping = pd.to_numeric(orders_df['Shipping_Price'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    shipping_paid = pd.to_numeric(orders_df.loc[orders_df['Shipping_Payment'].astype(str).eq('تم دفع الشحن'), 'Shipping_Price'], errors='coerce').fillna(0).sum() if not orders_df.empty else 0.0
    shipping_pending = total_shipping - shipping_paid
    total_expenses = pd.to_numeric(expenses_df['Amount'], errors='coerce').fillna(0).sum() if not expenses_df.empty else 0.0
    # التحصيل المالي للشركة يخص المنتجات فقط؛ رسوم الشحن مستبعدة لأنها تذهب مباشرة للناقل.
    total_collected_cash = collected_products_total
    actual_treasury = total_collected_cash - total_expenses

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("إجمالي مبيعات المنتجات", f"{total_products_value:.3f} ر.ع")
    m2.metric("المحصل من مبيعات المنتجات", f"{total_collected_cash:.3f} ر.ع")
    m3.metric("المبيعات غير المحصلة", f"{pending_products_total:.3f} ر.ع")
    m4.metric("💰 صافي التدفق النقدي (بدون الشحن)", f"{actual_treasury:.3f} ر.ع")
    st.caption(f"تنبيه محاسبي: رسوم الشحن ليست إيراداً أو مصروفاً أو تحصيلاً خاصاً بلبانور؛ هي مبالغ تخص شركة الشحن وتُدفع لها مباشرة. وتظهر هنا للمطابقة فقط — مدفوع لشركات الشحن: {shipping_paid:.3f} ر.ع | متبقٍ لشركات الشحن: {shipping_pending:.3f} ر.ع | إجمالي رسوم الشحن: {total_shipping:.3f} ر.ع. لا تدخل هذه المبالغ في المبيعات أو التحصيل أو صافي التدفق النقدي.")
    
    st.markdown("---")
    cargo_summary = get_cargo_summary_df(orders_df)
    if not cargo_summary.empty:
        st.dataframe(cargo_summary, use_container_width=True)

elif menu == "سجل حركة المستخدم والاتصال":
    st.header("🧾 سجل حركة المستخدم والاتصال")
    if not IS_ADMIN:
        st.error("هذا التقرير متاح للمدير فقط.")
    else:
        audit_df = read_activity_log(DB_FILE)
        if audit_df.empty:
            st.info("لا توجد حركات مسجلة بعد. سيبدأ النظام تسجيل العمليات المهمة والمزامنة من الآن.")
        else:
            if "Timestamp" in audit_df.columns:
                audit_df["Timestamp"] = pd.to_datetime(audit_df["Timestamp"], errors="coerce")
                date_range = st.date_input("فلترة السجل بالتاريخ", value=(datetime.date.today() - datetime.timedelta(days=30), datetime.date.today()))
                if isinstance(date_range, tuple) and len(date_range) == 2:
                    audit_df = audit_df[(audit_df["Timestamp"].dt.date >= date_range[0]) & (audit_df["Timestamp"].dt.date <= date_range[1])]
            if "Connection_Status" in audit_df.columns:
                connection_filter = st.selectbox("حالة الاتصال", ["الكل", "متصل - تمت المزامنة", "غير متصل أو رفض GitHub الطلب", "جارٍ مزامنة GitHub", "تم تحميل البيانات"])
                if connection_filter != "الكل":
                    audit_df = audit_df[audit_df["Connection_Status"].astype(str) == connection_filter]
            st.metric("عدد الأحداث المسجلة", len(audit_df))
            st.dataframe(audit_df.sort_values("Timestamp", ascending=False) if "Timestamp" in audit_df.columns else audit_df, use_container_width=True, hide_index=True)
            st.download_button("📥 تحميل سجل الحركة CSV", audit_df.to_csv(index=False).encode("utf-8-sig"), file_name="LubaNur_Activity_Log.csv", mime="text/csv")
        st.caption("يسجل النظام عمليات الدخول والخروج، وحفظ المعاملات، ومحاولات مزامنة GitHub ونتائجها. انقطاع اتصال المتصفح نفسه لا يمكن رصده لحظياً من الخادم؛ يسجل النظام فشل الاتصال عند محاولة المزامنة أو الحفظ.")

elif menu == "الربحية والأرباح":
    st.header("💹 قسم الربحية والأرباح")
    st.info("تكاليف شراء المنتجات تظهر في هذا القسم فقط، ولا تظهر في الفواتير أو لوحة المبيعات أو تقارير شركات الشحن.")
    orders_df, stock_df, expenses_df, _ = load_data()
    costs_df = load_product_costs(DB_FILE)
    stock_names = stock_df["Item_Name"].dropna().astype(str).str.strip().drop_duplicates().tolist()
    cost_map = dict(zip(costs_df["Item_Name"], costs_df["Unit_Cost"]))
    cost_editor_df = pd.DataFrame({"Item_Name": stock_names, "Unit_Cost": [float(cost_map.get(n, 0.0)) for n in stock_names]})
    st.subheader("إدخال تكلفة شراء كل منتج")
    st.caption("أدخل تكلفة الوحدة الواحدة لكل منتج بالريال العماني. تُحفظ التكاليف في ورقة مستقلة داخل Excel باسم Product_Costs.")
    edited_costs = st.data_editor(cost_editor_df, use_container_width=True, hide_index=True, disabled=["Item_Name"], column_config={"Item_Name": st.column_config.TextColumn("اسم المنتج"), "Unit_Cost": st.column_config.NumberColumn("تكلفة الوحدة (ر.ع)", min_value=0.0, step=0.001, format="%.3f")}, key="profitability_cost_editor")
    if IS_ADMIN and st.button("💾 حفظ تكاليف المنتجات", key="save_product_costs"):
        try:
            clean_costs = edited_costs[["Item_Name", "Unit_Cost"]].copy()
            clean_costs["Unit_Cost"] = pd.to_numeric(clean_costs["Unit_Cost"], errors="coerce").fillna(0.0).clip(lower=0)
            save_all_data(orders_df, stock_df, expenses_df, load_data()[3], product_costs_df=clean_costs)
            st.cache_data.clear()
            st.success("تم حفظ تكاليف المنتجات ومزامنتها مع ملف Excel.")
            st.rerun()
        except Exception as cost_error:
            st.error(f"تعذر حفظ تكاليف المنتجات: {cost_error}")

    detail_df, sales_value, cogs_value, expenses_value = calculate_profitability(orders_df, costs_df, expenses_df)
    gross_profit = sales_value - cogs_value
    net_profit = gross_profit - expenses_value
    missing_costs = detail_df[detail_df["حالة التكلفة"] == "لم تُدخل"] if not detail_df.empty else pd.DataFrame()
    st.markdown("---")
    st.subheader("ملخص الأرباح")
    p1, p2, p3 = st.columns(3)
    p1.metric("مبيعات المنتجات (بدون الشحن)", f"{sales_value:.3f} ر.ع")
    p2.metric("تكلفة المنتجات المباعة", f"{cogs_value:.3f} ر.ع")
    p3.metric("مجمل الربح قبل المصروفات", f"{gross_profit:.3f} ر.ع")
    p4, p5 = st.columns(2)
    p4.metric("إجمالي المصروفات المخصومة", f"{expenses_value:.3f} ر.ع")
    p5.metric("صافي الربح بعد خصم المصروفات", f"{net_profit:.3f} ر.ع")
    st.subheader("تفاصيل تكلفة المنتجات المباعة")
    if not detail_df.empty:
        st.dataframe(detail_df, use_container_width=True, hide_index=True)
    if not missing_costs.empty:
        st.warning("هناك منتجات مباعة لم تُدخل تكلفة وحدتها بعد. لذلك صافي الربح الحالي غير نهائي، والتكلفة المحتسبة لهذه المنتجات تساوي صفراً حتى إدخالها.")
        st.dataframe(missing_costs[["المنتج", "الكمية المباعة"]], use_container_width=True, hide_index=True)
    st.caption("طريقة الحساب: مجمل الربح = مبيعات المنتجات − تكلفة الوحدات المباعة. صافي الربح = مجمل الربح − جميع المصروفات المسجلة. رسوم الشحن مستبعدة لأنها تخص شركة الشحن. التكاليف التي تدخلها الآن تُطبق على الكميات المباعة في الطلبات التاريخية أيضاً؛ إذا تغيرت تكلفة الشراء بمرور الوقت فستحتاج لاحقاً إلى سجل تكلفة بتاريخ سريان.")

elif menu == "طلب توريد المخزون الحرِج":
    st.header("🚨 طلب توريد المخزون الحرِج")
    _, stock_df, _, _ = load_data()
    critical_stock = stock_df[stock_df['Qty'] < 5].copy()
    if not critical_stock.empty:
        st.dataframe(critical_stock, use_container_width=True)
    else:
        st.success("🎉 ممتاز! جميع المخزونات في مستويات آمنة.")
