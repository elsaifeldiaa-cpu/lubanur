import streamlit as st
import pandas as pd
import datetime
import os
from io import BytesIO

st.set_page_config(
    page_title="LubaNur Order & Stock System",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === 🔒 نظام حماية كلمة المرور (Password Protection) ===
PASSWORD_CORRECT = "142025"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
        <style>
        .login-container {
            max-width: 400px;
            margin: 100px auto;
            padding: 30px;
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            direction: rtl;
            text-align: right;
            font-family: Tahoma, sans-serif;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<h2 style='text-align: center; color: #2c3e50;'>🔐 تسجيل الدخول لنظام لبانور</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>الرجاء إدخال كلمة المرور للوصول إلى لوحة التحكم</p>", unsafe_allow_html=True)
    
    entered_password = st.text_input("كلمة المرور (Password)", type="password", key="login_pass")
    
    if st.button("دخول (Login)", use_container_width=True):
        if entered_password == PASSWORD_CORRECT:
            st.session_state.authenticated = True
            st.success("✅ تم تسجيل الدخول بنجاح!")
            st.rerun()
        else:
            st.error("❌ كلمة المرور غير صحيحة! حاول مرة أخرى.")
    st.stop()

# === تخصيص CSS إضافي لمظهر راقي وعصري (Elite UI) ===
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
        "Shipping_Price": 0.0, "Total": 0.0, "Order_Payment": "تم الدفع", 
        "Shipping_Payment": "تم دفع الشحن", "Cargo_Company": "بدون شحن", 
        "Delivery_Type": "غير محدد", "Delivery_Status": "قيد التحضير", 
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
    return orders

def init_db():
    if not os.path.exists(DB_FILE):
        if os.path.exists(ORIGINAL_EXCEL):
            try:
                xls_orig = pd.ExcelFile(ORIGINAL_EXCEL)
                stock = pd.read_excel(xls_orig, sheet_name="Stock +").groupby("Item Name")["Qty"].sum().reset_index()
                stock.columns = ["Item_Name", "Qty"]
            except Exception:
                stock = pd.DataFrame({"Item_Name": EXACT_ITEMS, "Qty": [50] * len(EXACT_ITEMS)})
        else:
            stock = pd.DataFrame({"Item_Name": EXACT_ITEMS, "Qty": [50] * len(EXACT_ITEMS)})
            
        orders = pd.DataFrame(columns=["Order_ID", "Date", "Customer", "Phone", "Region", "Items_Detail", "Products_Total", "Shipping_Price", "Total", "Order_Payment", "Shipping_Payment", "Cargo_Company", "Delivery_Type", "Delivery_Status", "Collection_Status", "Bags", "Cargo_Barcode"])
        expenses = pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])
        
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            stock.to_excel(writer, sheet_name="Stock", index=False)
            expenses.to_excel(writer, sheet_name="Expenses", index=False)
    else:
        xls = pd.ExcelFile(DB_FILE)
        orders = pd.read_excel(xls, sheet_name="Orders")
        orders = normalize_orders_df(orders)
        
        stock = pd.read_excel(xls, sheet_name="Stock") if "Stock" in xls.sheet_names else pd.DataFrame(columns=["Item_Name", "Qty"])
        expenses = pd.read_excel(xls, sheet_name="Expenses") if "Expenses" in xls.sheet_names else pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])

        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            stock.to_excel(writer, sheet_name="Stock", index=False)
            expenses.to_excel(writer, sheet_name="Expenses", index=False)

init_db()

def load_data():
    xls = pd.ExcelFile(DB_FILE)
    orders = pd.read_excel(xls, sheet_name="Orders")
    orders = normalize_orders_df(orders)
    
    stock = pd.read_excel(xls, sheet_name="Stock")
    expenses = pd.read_excel(xls, sheet_name="Expenses") if "Expenses" in xls.sheet_names else pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])
    return orders, stock, expenses

orders_df, stock_df, expenses_df = load_data()

st.sidebar.markdown('<p class="sidebar-title">🌟 لبانور - لوحة التحكم الإدارية</p>', unsafe_allow_html=True)
st.sidebar.markdown('<p class="version-tag">📌 الإصدار: v1:10am 27-10-2026</p>', unsafe_allow_html=True)

menu_options = {
    "📊 اللوحة الرئيسية (Dashboard & Financials)": "اللوحة الرئيسية (Dashboard & Financials)",
    "📝 تسجيل طلب جديد": "تسجيل طلب جديد (فاتورة متعددة الأصناف)",
    "⏳ الطلبات قيد التحضير": "الطلبات التي لم يتم تحضيرها (Pending Preparation)",
    "🔍 تتبع الطلبات الشامل": "تتبع الطلبات الشامل (بحث بالهاتف أو الباركود)",
    "🚚 تتبع التوصيل والشحن": "تتبع التوصيل والشحن",
    "📑 تقرير الشحن (شركات الشحن)": "تقرير الشحن لشركة الشحن (تقرير الاستلام)",
    "📅 تقرير الطلبات للفترة مع الأصناف": "تقرير الطلبات للفترة مع الأصناف",
    "💵 متابعة التحصيل المالي": "متابعة التحصيل المالي",
    "✏️ تعديل وحذف الفواتير": "تعديل وحذف الفواتير (Edit & Delete)",
    "🖨️ فاتورة العميل (Compact)": "فاتورة العميل (Compact Receipt)",
    "💸 إدارة المصروفات": "إدارة المصروفات (Expenses)",
    "📈 التقرير المالي التفصيلي": "التقرير المالي التفصيلي (Financial Report)",
    "📦 إدارة المخزون": "إدارة المخزون (Stock)",
    "🚨 طلب توريد المخزون الحرِج": "طلب توريد المخزون الحرِج"
}

selected_label = st.sidebar.selectbox("🧭 الانتقال السريع بين الأقسام", list(menu_options.keys()))
menu = menu_options[selected_label]

# === زر الرجوع (Back Button) في القائمة الجانبية أو في الأعلى ===
if st.sidebar.button("⬅️ رجوع للوحة الرئيسية"):
    st.rerun()

# === 💾 إدارة النسخ الاحتياطي واسترجاع البيانات في القائمة الجانبية ===
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
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="انقر لتحميل نسخة كاملة ومحدثة من قاعدة البيانات على جهازك أو الأيباد"
    )

uploaded_backup = st.sidebar.file_uploader("📤 استرجاع بيانات من نسخة احتياطية (.xlsx)", type=["xlsx"])
if uploaded_backup is not None:
    try:
        with open(DB_FILE, "wb") as f:
            f.write(uploaded_backup.getbuffer())
        st.sidebar.success("✅ تم استرجاع البيانات بنجاح! جاري التحديث...")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"خطأ في استرجاع الملف: {e}")

if st.sidebar.button("🔒 تسجيل الخروج (Logout)"):
    st.session_state.authenticated = False
    st.rerun()

st.title("📦 نظام إدارة الطلبيات والمخزون والمصروفات - لبانور / السيف والضياء")
st.markdown("---")

def get_cargo_summary_df(orders_df):
    if orders_df.empty or 'Cargo_Company' not in orders_df.columns:
        return pd.DataFrame()
    
    cargo_groups = orders_df.groupby('Cargo_Company')
    summary_list = []
    
    for company, group in cargo_groups:
        total_orders = len(group)
        products_total = group['Products_Total'].sum()
        shipping_total = group['Shipping_Price'].sum()
        grand_total = group['Total'].sum()
        
        collected_amount = group[group['Collection_Status'] == 'مدفوع بالكامل']['Products_Total'].sum()
        pending_amount = group[group['Collection_Status'] == 'قيد التحصيل']['Products_Total'].sum()
        
        summary_list.append({
            'شركة الشحن': company,
            'عدد الطلبات': total_orders,
            'مستحقات المنتجات (ر.ع)': products_total,
            'ما تم تحصيله (ر.ع)': collected_amount,
            'المستحق المتبقي لدى الشركة (ر.ع)': pending_amount,
            'قيمة الشحن (ر.ع)': shipping_total,
            'الإجمالي الكلي (ر.ع)': grand_total
        })
        
    return pd.DataFrame(summary_list)

if menu == "اللوحة الرئيسية (Dashboard & Financials)":
    st.header("📊 الملخص التشغيلي والمالي الشامل")
    col_ref1, _ = st.columns([1, 4])
    with col_ref1:
        if st.button("🔄 تحديث البيانات (Refresh)", key="ref_dash"):
            st.rerun()
            
    orders_df, stock_df, expenses_df = load_data()
    
    total_products_sales = orders_df['Products_Total'].sum() if not orders_df.empty and 'Products_Total' in orders_df.columns else 0.0
    total_shipping = orders_df['Shipping_Price'].sum() if not orders_df.empty and 'Shipping_Price' in orders_df.columns else 0.0
    total_expenses = expenses_df['Amount'].sum() if not expenses_df.empty and 'Amount' in expenses_df.columns else 0.0
    net_treasury = total_products_sales - total_expenses

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("إجمالي مبيعات المنتجات (إيراداتنا)", f"{total_products_sales:.3f} ر.ع")
    col2.metric("إجمالي قيمة الشحن (لشركة الشحن)", f"{total_shipping:.3f} ر.ع")
    col3.metric("إجمالي المصروفات", f"{total_expenses:.3f} ر.ع")
    col4.metric("صافي الخزينة", f"{net_treasury:.3f} ر.ع", delta="بعد خصم المصروفات من إيرادات المنتجات")
    
    st.markdown("---")
    st.subheader("📋 ملخص مبيعات ومستحقات المنتجات لدى شركات الشحن (ما تم تحصيله والمستحق المتبقي)")
    cargo_summary = get_cargo_summary_df(orders_df)
    if not cargo_summary.empty:
        st.dataframe(cargo_summary, use_container_width=True)
    else:
        st.info("لا توجد طلبيات مسجلة بعد.")

elif menu == "تسجيل طلب جديد (فاتورة متعددة الأصناف)":
    st.header("📝 إنشاء فاتورة جديدة (إضافة عدة منتجات في نفس الطلب)")
    
    col_ref_new, _ = st.columns([1, 4])
    with col_ref_new:
        if st.button("🔄 تحديث البيانات (Refresh)", key="ref_new_ord"):
            st.rerun()

    if 'cart' not in st.session_state:
        st.session_state.cart = []

    cust_name = st.text_input("اسم العميل (Customer Name)", key="new_cust_name")
    
    col_inf1, col_inf2, col_inf3 = st.columns(3)
    with col_inf1:
        phone = st.text_input("رقم الهاتف (Phone Number)", key="new_phone")
        region = st.text_input("المنطقة / الولاية (Region)", key="new_region")
    with col_inf2:
        order_date = st.date_input("تاريخ الطلب (Date)", datetime.date.today(), key="new_date")
        cargo_company = st.selectbox("شركة الشحن (Cargo Company)", ["بدون شحن", "نول", "جيناكم", "دليلي", "وصل لي"], key="new_cargo")
        delivery_type = st.selectbox("نوع التسليم (Delivery Type)", ["من المكتب (Pickup)", "إلى المنزل (Home Delivery)"], key="new_dtype")
    with col_inf3:
        shipping_price = st.number_input("سعر الشحن (Shipping Price OMR)", min_value=0.0, value=1.0, key="new_ship_price")
        order_payment = st.selectbox("حالة دفع الطلب (Order Payment Status)", ["تم الدفع", "لم يتم الدفع (COD)"], key="new_ord_pay")
        shipping_payment = st.selectbox("حالة دفع قيمة الشحن (Shipping Payment Status)", ["تم دفع الشحن", "لم يتم دفع الشحن"], key="new_shp_pay")
        bags = st.number_input("عدد الأكياس (Bags)", min_value=0, value=1, key="new_bags")

    st.markdown("---")
    st.subheader("2️⃣ إضافة منتجات الفاتورة (حدد المنتج والكمية فقط)")
    
    sel_item = st.selectbox("اختر المنتج (Select Product)", stock_df['Item_Name'].tolist(), key="cart_sel_item")
    sel_qty = st.number_input("الكمية المطلوبة (Qty)", min_value=1, value=1, key="cart_sel_qty")

    if st.button("➕ إضافة هذا المنتج إلى الفاتورة", key="btn_add_prod"):
        avail_stock = stock_df.loc[stock_df['Item_Name'] == sel_item, 'Qty'].values[0]
        if avail_stock >= sel_qty:
            st.session_state.cart.append({"Item": sel_item, "Qty": sel_qty})
            st.success(f"تم بنجاح إضافة: {sel_item} (الكمية: {sel_qty}) إلى الفاتورة!")
        else:
            st.error(f"⚠️ المخزون غير كافٍ! المتوفر من هذا المنتج هو {avail_stock} فقط.")

    if st.session_state.cart:
        st.markdown("---")
        st.subheader("🛒 محتويات المنتجات في الفاتورة الحالية:")
        cart_df = pd.DataFrame(st.session_state.cart)
        st.dataframe(cart_df, use_container_width=True)
        
        st.markdown("---")
        products_total = st.number_input("إجمالي قيمة المنتجات (Products Total OMR)", min_value=0.0, value=5.0, key="new_prod_total")
        grand_total = products_total + shipping_price
        st.info(f"💡 الإجمالي الكلي للفاتورة: **{grand_total} ريال عماني**")
        
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("🗑️ تفريغ السلة وإلغاء الطلب", key="btn_clear_cart"):
                st.session_state.cart = []
                st.rerun()
        with col_b2:
            if st.button("✅ حفظ واعتماد الطلب نهائياً وخصم المخزون", key="btn_save_ord"):
                if not cust_name:
                    st.error("الرجاء إدخال اسم العميل أولاً.")
                else:
                    orders_df, stock_df, expenses_df = load_data()
                    stock_ok = True
                    for row in st.session_state.cart:
                        it_n, it_q = row['Item'], row['Qty']
                        stk = stock_df.loc[stock_df['Item_Name'] == it_n, 'Qty'].values[0]
                        if stk < it_q:
                            st.error(f"⚠️ المخزون غير كافٍ للمنتج {it_n} (المتوفر: {stk})")
                            stock_ok = False
                            break
                    
                    if stock_ok:
                        items_summary_list = []
                        for row in st.session_state.cart:
                            it_n, it_q = row['Item'], row['Qty']
                            stock_df.loc[stock_df['Item_Name'] == it_n, 'Qty'] -= it_q
                            items_summary_list.append(f"{it_n} (الكمية: {it_q})")
                        
                        items_detail_str = " | ".join(items_summary_list)
                        new_id = int(orders_df['Order_ID'].max() + 1) if not orders_df.empty and 'Order_ID' in orders_df.columns and not pd.isna(orders_df['Order_ID'].max()) else 1
                        
                        coll_status = "مدفوع بالكامل" if (order_payment == "تم الدفع" and shipping_payment == "تم دفع الشحن") else "قيد التحصيل"
                        
                        new_order_row = {
                            "Order_ID": new_id, "Date": str(order_date), "Customer": cust_name, "Phone": phone, "Region": region,
                            "Items_Detail": items_detail_str, "Products_Total": products_total, "Shipping_Price": shipping_price, "Total": grand_total,
                            "Order_Payment": order_payment, "Shipping_Payment": shipping_payment, "Cargo_Company": cargo_company, "Delivery_Type": delivery_type,
                            "Delivery_Status": "قيد التحضير", "Collection_Status": coll_status, "Bags": bags, "Cargo_Barcode": ""
                        }
                        
                        orders_df = pd.concat([orders_df, pd.DataFrame([new_order_row])], ignore_index=True)
                        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                            orders_df.to_excel(writer, sheet_name="Orders", index=False)
                            stock_df.to_excel(writer, sheet_name="Stock", index=False)
                            expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                        
                        st.session_state.cart = []
                        st.cache_data.clear()
                        st.success(f"🎉 تم حفظ الطلب رقم #{new_id} بنجاح!")
                        st.rerun()

elif menu == "الطلبات التي لم يتم تحضيرها (Pending Preparation)":
    st.header("⏳ متابعة الطلبات التي لم يتم تحضيرها بعد (قيد التحضير)")
    orders_df, stock_df, expenses_df = load_data()
    if not orders_df.empty and 'Delivery_Status' in orders_df.columns:
        pending_df = orders_df[orders_df['Delivery_Status'] == 'قيد التحضير']
        if not pending_df.empty:
            st.dataframe(pending_df[['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Company', 'Items_Detail', 'Bags']], use_container_width=True)
            up_pend_id = st.selectbox("اختر رقم الطلب لتحويله", pending_df['Order_ID'].tolist())
            
            new_pend_status = st.selectbox("الحالة الجديدة", ["قيد التحضير", "تم التسليم لشركة الشحن", "جاري التوصيل", "تم تسليم العميل", "مرتجع"])
            
            cargo_barcode_input = ""
            if new_pend_status == "تم التسليم لشركة الشحن":
                cargo_barcode_input = st.text_input("📦 أدخل باركود / رقم بوليصة شركة الشحن (Tracking Barcode)")
                
            if st.button("💾 تحديث حالة الطلب"):
                orders_df.loc[orders_df['Order_ID'] == up_pend_id, 'Delivery_Status'] = new_pend_status
                if new_pend_status == "تم التسليم لشركة الشحن" and cargo_barcode_input:
                    orders_df.loc[orders_df['Order_ID'] == up_pend_id, 'Cargo_Barcode'] = cargo_barcode_input
                    
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                st.success("✅ تم تحديث حالة الطلب والباركود بنجاح!")
                st.rerun()
        else:
            st.success("🎉 لا توجد طلبيات معلقة قيد التحضير!")

elif menu == "🔍 تتبع الطلبات الشامل":
    st.header("🔍 نظام تتبع الطلبات الشامل")
    orders_df, _, _ = load_data()
    
    if not orders_df.empty:
        st.markdown("### 🎯 اختر نوع البحث المطلوب:")
        search_type = st.selectbox(
            "البحث بواسطة:",
            ["عرض كافة الطلبات", "رقم الطلب (Order ID)", "رقم هاتف العميل (Phone)", "باركود شركة الشحن (Cargo Barcode)"]
        )
        
        search_query = st.text_input("أدخل نص أو رقم البحث:")
        
        filtered_results = orders_df.copy()
        
        if search_type == "رقم الطلب (Order ID)" and search_query:
            filtered_results = orders_df[orders_df['Order_ID'].astype(str).str.contains(search_query.strip(), na=False)]
        elif search_type == "رقم هاتف العميل (Phone)" and search_query:
            filtered_results = orders_df[orders_df['Phone'].astype(str).str.contains(search_query.strip(), na=False)]
        elif search_type == "باركود شركة الشحن (Cargo Barcode)" and search_query:
            filtered_results = orders_df[orders_df['Cargo_Barcode'].astype(str).str.contains(search_query.strip(), na=False)]
        elif search_type != "عرض كافة الطلبات" and search_query:
            # البحث الشامل في حالة عدم التحديد الدقيق
            q = search_query.strip()
            filtered_results = orders_df[
                orders_df['Phone'].astype(str).str.contains(q, na=False) |
                orders_df['Order_ID'].astype(str).str.contains(q, na=False) |
                orders_df['Cargo_Barcode'].astype(str).str.contains(q, na=False) |
                orders_df['Customer'].astype(str).str.contains(q, na=False)
            ]
            
        st.markdown(f"**النتائج المطابقة ({len(filtered_results)} طلب):**")
        st.dataframe(filtered_results[['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Company', 'Cargo_Barcode', 'Delivery_Status', 'Collection_Status', 'Total']], use_container_width=True)
    else:
        st.info("لا توجد طلبيات مسجلة في النظام بعد.")

elif menu == "إدارة المصروفات (Expenses)":
    st.header("💸 إدارة المصروفات")
    orders_df, stock_df, expenses_df = load_data()
    with st.form("add_expense_form"):
        exp_date = st.date_input("التاريخ", datetime.date.today())
        exp_category = st.selectbox("البند", ["إعلانات ميتا (Meta Ads)", "شراء أصناف / مواد خام (Purchases)", "سحب مبالغ مالية / مسحوبات شخصية (Withdrawals)", "مصروفات أخرى (Other Expenses)"])
        exp_amount = st.number_input("المبلغ (OMR)", min_value=0.0, value=10.0)
        exp_desc = st.text_input("الوصف")
        if st.form_submit_button("💾 حفظ المصروف"):
            new_exp_id = int(expenses_df['Expense_ID'].max() + 1) if not expenses_df.empty and not pd.isna(expenses_df['Expense_ID'].max()) else 1
            new_row = {"Expense_ID": new_exp_id, "Date": str(exp_date), "Category": exp_category, "Amount": exp_amount, "Description": exp_desc}
            expenses_df = pd.concat([expenses_df, pd.DataFrame([new_row])], ignore_index=True)
            with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                orders_df.to_excel(writer, sheet_name="Orders", index=False)
                stock_df.to_excel(writer, sheet_name="Stock", index=False)
                expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
            st.success("✅ تم حفظ المصروف بنجاح!")
            st.rerun()
    if not expenses_df.empty:
        st.dataframe(expenses_df, use_container_width=True)

elif menu == "التقرير المالي التفصيلي (Financial Report)":
    st.header("📈 التقرير المالي التفصيلي والشامل")
    orders_df, stock_df, expenses_df = load_data()
    total_products_value = orders_df['Products_Total'].sum() if not orders_df.empty else 0.0
    total_shipping = orders_df['Shipping_Price'].sum() if not orders_df.empty else 0.0
    total_expenses = expenses_df['Amount'].sum() if not orders_df.empty else 0.0
    net_treasury = total_products_value - total_expenses

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("إجمالي مبيعات المنتجات", f"{total_products_value:.3f} ر.ع")
    m2.metric("إجمالي قيمة الشحن", f"{total_shipping:.3f} ر.ع")
    m3.metric("المصروفات", f"{total_expenses:.3f} ر.ع")
    m4.metric("صافي الخزينة", f"{net_treasury:.3f} ر.ع")
    st.info(f"💰 **صافي السيولة بالخزينة (بعد خصم المصروفات من إيرادات المنتجات):** {net_treasury:.3f} ريال عماني")
    st.markdown("---")
    cargo_summary = get_cargo_summary_df(orders_df)
    if not cargo_summary.empty:
        st.dataframe(cargo_summary, use_container_width=True)

elif menu == "فاتورة العميل (Compact Receipt)":
    st.header("🖨️ إصدار فاتورة العميل الحرارية")
    orders_df, _, _ = load_data()
    if not orders_df.empty:
        selected_order_id = st.selectbox("اختر رقم الطلب", orders_df['Order_ID'].tolist())
        order_row = orders_df[orders_df['Order_ID'] == selected_order_id].iloc[0]
        
        raw_items = str(order_row['Items_Detail'])
        formatted_items_text = raw_items.replace(" | ", "<br>- ")
        if formatted_items_text and not formatted_items_text.startswith("-"):
            formatted_items_text = "- " + formatted_items_text

        barcode_display = f"<p><b>باركود شركة الشحن:</b> {order_row.get('Cargo_Barcode', 'غير متوفر')}</p>" if order_row.get('Cargo_Barcode') else ""

        receipt_html = f"""
        <!DOCTYPE html>
        <html dir="rtl" lang="ar">
        <head>
            <meta charset="utf-8">
            <title>فاتورة #{order_row['Order_ID']} - لبانور</title>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <style>
                body {{ font-family: Tahoma, sans-serif; padding: 20px; direction: rtl; background-color: #f4f6f9; }}
                .receipt-box {{ max-width: 420px; margin: 0 auto; border: 2px dashed #2c3e50; border-radius: 12px; padding: 20px; background-color: #ffffff; box-shadow: 0 4px 10px rgba(0,0,0,0.08); }}
                h3 {{ text-align: center; color: #2c3e50; margin-bottom: 5px; }}
                .tagline {{ text-align: center; color: #666; font-size: 12px; margin-top: 0px; border-bottom: 1px dashed #ddd; padding-bottom: 10px; }}
                .info {{ font-size: 14px; color: #333; line-height: 1.6; }}
                .items {{ font-size: 13px; color: #333; background: #f9f9f9; padding: 10px; border-radius: 6px; line-height: 1.6; margin: 10px 0; }}
                .total {{ font-size: 16px; color: #007bff; font-weight: bold; text-align: center; margin-top: 10px; }}
                .btn-container {{ display: flex; gap: 10px; justify-content: center; margin-top: 15px; flex-wrap: wrap; }}
                .btn {{ padding: 10px 18px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; text-decoration: none; font-size: 14px; }}
                .btn-print {{ background-color: #28a745; color: white; }}
                .btn-img {{ background-color: #007bff; color: white; }}
                @media print {{ .no-print {{ display: none !important; }} body {{ background-color: white; }} .receipt-box {{ border: 1px solid #000; box-shadow: none; }} }}
            </style>
        </head>
        <body>
            <div id="receipt-capture" class="receipt-box">
                <h3>=== لبانور - LubaNur ===</h3>
                <div class="tagline">السيف والضياء للاستثمار الدولي</div>
                <div class="info">
                    <p><b>رقم الفاتورة:</b> #{order_row['Order_ID']} &nbsp;|&nbsp; <b>التاريخ:</b> {order_row['Date']}</p>
                    <p><b>العميل:</b> {order_row['Customer']} &nbsp;|&nbsp; <b>الهاتف:</b> {order_row['Phone']}</p>
                    <p><b>المنطقة:</b> {order_row['Region']}</p>
                    {barcode_display}
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
            
            <div class="btn-container" style="max-width: 420px; margin: 15px auto;">
                <button class="btn btn-print" onclick="window.print()">🖨️ طباعة / PDF</button>
                <button class="btn btn-img" onclick="downloadPNG()">📷 تحميل الفاتورة كصورة (PNG)</button>
            </div>

            <script>
                function downloadPNG() {{
                    const element = document.getElementById('receipt-capture');
                    html2canvas(element, {{ scale: 2, useCORS: true }}).then(canvas => {{
                        const link = document.createElement('a');
                        link.download = 'LubaNur_Invoice_#{order_row["Order_ID"]}.png';
                        link.href = canvas.toDataURL('image/png');
                        link.click();
                    }});
                }}
            </script>
        </body>
        </html>
        """
        
        st.download_button(
            label="📥 اضغط هنا لفتح واجهة الفاتورة (للطباعة أو حفظ صورة PNG)",
            data=receipt_html.encode('utf-8'),
            file_name=f"Invoice_{order_row['Order_ID']}_{order_row['Customer']}.html",
            mime="text/html",
            use_container_width=True
        )

elif menu == "تعديل وحذف الفواتير (Edit & Delete)":
    st.header("✏️ تعديل وحذف الفواتير")
    orders_df, stock_df, expenses_df = load_data()
    if not orders_df.empty:
        edit_order_id = st.selectbox("اختر رقم الطلب للتعديل", orders_df['Order_ID'].tolist())
        target_row = orders_df[orders_df['Order_ID'] == edit_order_id].iloc[0]
        with st.form("edit_form"):
            e_cust = st.text_input("اسم العميل", value=str(target_row['Customer']))
            e_phone = st.text_input("الهاتف", value=str(target_row['Phone']))
            e_region = st.text_input("المنطقة / الولاية", value=str(target_row.get('Region', '')))
            e_bags = st.number_input("عدد الأكياس", min_value=0, value=int(target_row.get('Bags', 1)))
            e_prod_tot = st.number_input("قيمة المنتجات", value=float(target_row['Products_Total']))
            e_ship_p = st.number_input("سعر الشحن", value=float(target_row['Shipping_Price']))
            e_barcode = st.text_input("باركود شركة الشحن", value=str(target_row.get('Cargo_Barcode', '')))
            
            cur_ord_pay = str(target_row.get('Order_Payment', 'تم الدفع'))
            cur_shp_pay = str(target_row.get('Shipping_Payment', 'تم دفع الشحن'))
            
            e_ord_pay = st.selectbox("حالة دفع الطلب", ["تم الدفع", "لم يتم الدفع (COD)"], index=0 if cur_ord_pay == "تم الدفع" else 1)
            e_shp_pay = st.selectbox("حالة دفع قيمة الشحن", ["تم دفع الشحن", "لم يتم دفع الشحن"], index=0 if cur_shp_pay == "تم دفع الشحن" else 1)
            
            if st.form_submit_button("💾 حفظ التعديلات"):
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Customer'] = e_cust
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Phone'] = e_phone
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Region'] = e_region
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Bags'] = e_bags
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Products_Total'] = e_prod_tot
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Shipping_Price'] = e_ship_p
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Total'] = e_prod_tot + e_ship_p
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Cargo_Barcode'] = e_barcode
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Order_Payment'] = e_ord_pay
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Shipping_Payment'] = e_shp_pay
                
                new_coll_status = "مدفوع بالكامل" if (e_ord_pay == "تم الدفع" and e_shp_pay == "تم دفع الشحن") else "قيد التحصيل"
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Collection_Status'] = new_coll_status

                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                st.success("✅ تم التعديل بنجاح!")
                st.rerun()
                
        if st.button("🗑️ حذف الفاتورة بالكامل"):
            orders_df = orders_df[orders_df['Order_ID'] != edit_order_id]
            with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                orders_df.to_excel(writer, sheet_name="Orders", index=False)
                stock_df.to_excel(writer, sheet_name="Stock", index=False)
                expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
            st.success("🗑️ تم حذف الفاتورة بالكامل!")
            st.rerun()

elif menu == "إدارة المخزون (Stock)":
    st.header("📦 إدارة المخزون الحالي")
    _, stock_df, _ = load_data()
    st.dataframe(stock_df, use_container_width=True)
    up_item = st.selectbox("اختر المنتج", stock_df['Item_Name'].tolist())
    add_qty = st.number_input("الكمية المضافة (أو بالسالب)", value=10)
    if st.button("💾 تحديث المخزون"):
        orders_df, stock_df, expenses_df = load_data()
        current_q = stock_df.loc[stock_df['Item_Name'] == up_item, 'Qty'].values[0]
        stock_df.loc[stock_df['Item_Name'] == up_item, 'Qty'] = max(0, current_q + add_qty)
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders_df.to_excel(writer, sheet_name="Orders", index=False)
            stock_df.to_excel(writer, sheet_name="Stock", index=False)
            expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
        st.success("✅ تم تحديث المخزون بنجاح!")
        st.rerun()

elif menu == "طلب توريد المخزون الحرِج":
    st.header("🚨 طلب توريد المخزون الحرِج (أقل من 5 وحدات)")
    _, stock_df, _ = load_data()
    
    critical_stock = stock_df[stock_df['Qty'] < 5].copy()
    
    if not critical_stock.empty:
        st.warning(f"⚠️ يوجد عدد {len(critical_stock)} صنف بحاجة ماسة لإعادة التوريد! (أقل من 5 قطع)")
        
        def highlight_critical(s):
            return ['background-color: #ffcccc; color: #990000; font-weight: bold;' if s['Qty'] < 5 else '' for _ in s]
        
        st.dataframe(critical_stock.style.apply(highlight_critical, axis=1), use_container_width=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            critical_stock.to_excel(writer, sheet_name="Restock Order", index=False)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 تحميل تقرير طلب التوريد (Excel)",
            data=processed_data,
            file_name="LubaNur_Restock_Order.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("🎉 ممتاز! جميع المخزونات في مستويات آمنة (لا توجد أصناف أقل من 5 وحدات).")

elif menu == "تتبع التوصيل والشحن":
    st.header("🚚 متابعة التوصيل والشحن")
    orders_df, _, _ = load_data()
    if not orders_df.empty:
        st.dataframe(orders_df[['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Company', 'Cargo_Barcode', 'Delivery_Type', 'Delivery_Status']], use_container_width=True)

elif menu == "تقرير الشحن لشركة الشحن (تقرير الاستلام)":
    st.header("📑 تقرير تسليم الشحنات لشركات الشحن")
    orders_df, _, _ = load_data()
    
    if not orders_df.empty:
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            selected_cargo = st.selectbox("اختر شركة الشحن", orders_df['Cargo_Company'].unique().tolist())
        with col_c2:
            filter_date_option = st.checkbox("فلترة حسب التاريخ؟", value=False)
            
        filtered_df = orders_df[orders_df['Cargo_Company'] == selected_cargo].copy()
        
        if filter_date_option:
            rep_date = st.date_input("اختر تاريخ التقرير", datetime.date.today())
            filtered_df = filtered_df[filtered_df['Date'] == str(rep_date)]
            
        st.subheader(f"الطلبات الخاصة بـ: {selected_cargo}")
        
        display_df = filtered_df[['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Barcode', 'Products_Total', 'Shipping_Price', 'Total', 'Bags']].copy()
        display_df.columns = ['رقم الطلب', 'التاريخ', 'العميل', 'الهاتف', 'المنطقة', 'باركود الشحن', 'إجمالي المنتجات', 'سعر الشحن', 'الإجمالي', 'الأكياس']
        
        st.dataframe(display_df, use_container_width=True)
        
        if not filtered_df.empty:
            st.markdown("---")
            st.subheader("📥 تحميل وتصدير التقرير للطباعة")
            
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                excel_output = BytesIO()
                with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                    display_df.to_excel(writer, sheet_name="Cargo_Report", index=False)
                excel_data = excel_output.getvalue()
                
                st.download_button(
                    label="📊 تحميل التقرير (ملف Excel)",
                    data=excel_data,
                    file_name=f"LubaNur_Cargo_{selected_cargo}_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col_d2:
                html_table = display_df.to_html(index=False, classes='styled-table')
                html_page = f"""
                <!DOCTYPE html>
                <html dir="rtl" lang="ar">
                <head>
                    <meta charset="utf-8">
                    <title>تقرير شحن - {selected_cargo}</title>
                    <style>
                        body {{ font-family: Tahoma, sans-serif; padding: 20px; direction: rtl; }}
                        h2 {{ color: #2c3e50; text-align: center; margin-bottom: 20px; }}
                        .styled-table {{ width: 100%; border-collapse: collapse; margin: 25px 0; font-size: 15px; text-align: right; }}
                        .styled-table th {{ background-color: #007bff; color: white; padding: 12px 15px; }}
                        .styled-table td {{ padding: 10px 15px; border-bottom: 1px solid #dddddd; }}
                        .styled-table tr:nth-of-type(even) {{ background-color: #f3f3f3; }}
                        @media print {{ .no-print {{ display: none; }} }}
                    </style>
                </head>
                <body>
                    <h2>📦 تقرير تسليم شحنات شركة ({selected_cargo}) - {datetime.date.today()}</h2>
                    <button class="no-print" onclick="window.print()" style="padding: 10px 20px; font-size: 16px; margin-bottom: 15px; cursor: pointer; background-color: #28a745; color: white; border: none; border-radius: 5px;">🖨️ طباعة التقرير / حفظ PDF</button>
                    {html_table}
                </body>
                </html>
                """
                
                st.download_button(
                    label="🖨️ فتح التقرير للطباعة / حفظ كـ PDF",
                    data=html_page.encode('utf-8'),
                    file_name=f"Report_{selected_cargo}_{datetime.date.today()}.html",
                    mime="text/html",
                    use_container_width=True
                )
        else:
            st.info("لا توجد طلبيات مطابقة لهذه الفلاتر.")

elif menu == "تقرير الطلبات للفترة مع الأصناف":
    st.header("📅 تقرير الطلبات للفترة التفصيلي مع محتويات الأصناف")
    orders_df, _, _ = load_data()
    
    if not orders_df.empty:
        col_fd1, col_fd2 = st.columns(2)
        with col_fd1:
            start_date = st.date_input("من تاريخ", datetime.date.today() - datetime.timedelta(days=30))
        with col_fd2:
            end_date = st.date_input("إلى تاريخ", datetime.date.today())
            
        orders_df['Date_parsed'] = pd.to_datetime(orders_df['Date'], errors='coerce')
        mask = (orders_df['Date_parsed'].dt.date >= start_date) & (orders_df['Date_parsed'].dt.date <= end_date)
        period_df = orders_df[mask].copy()
        
        st.subheader(f"الطلبات المسجلة من {start_date} إلى {end_date}")
        if not period_df.empty:
            display_period = period_df[['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Barcode', 'Items_Detail', 'Products_Total', 'Shipping_Price', 'Total', 'Cargo_Company', 'Collection_Status']].copy()
            st.dataframe(display_period, use_container_width=True)
            
            excel_output = BytesIO()
            with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                display_period.to_excel(writer, sheet_name="Period_Report", index=False)
            excel_data = excel_output.getvalue()
            
            st.download_button(
                label="📊 تحميل تقرير الفترة (ملف Excel)",
                data=excel_data,
                file_name=f"LubaNur_Period_Report_{start_date}_to_{end_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("لا توجد طلبيات مسجلة ضمن هذه الفترة المحددة.")

elif menu == "متابعة التحصيل المالي":
    st.header("💵 متابعة التحصيل المالي")
    orders_df, stock_df, expenses_df = load_data()
    if not orders_df.empty:
        st.dataframe(orders_df[['Order_ID', 'Date', 'Customer', 'Phone', 'Cargo_Company', 'Cargo_Barcode', 'Products_Total', 'Collection_Status']], use_container_width=True)
        pending_coll = orders_df[orders_df['Collection_Status'] == 'قيد التحصيل']
        if not pending_coll.empty:
            sel_coll_id = st.selectbox("اختر رقم الطلب لتحصيله", pending_coll['Order_ID'].tolist())
            if st.button("💾 تحويل إلى 'مدفوع بالكامل'"):
                orders_df.loc[orders_df['Order_ID'] == sel_coll_id, 'Collection_Status'] = "مدفوع بالكامل"
                orders_df.loc[orders_df['Order_ID'] == sel_coll_id, 'Order_Payment'] = "تم الدفع"
                orders_df.loc[orders_df['Order_ID'] == sel_coll_id, 'Shipping_Payment'] = "تم دفع الشحن"
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                st.success("✅ تم تحديث التحصيل بنجاح!")
                st.rerun()
