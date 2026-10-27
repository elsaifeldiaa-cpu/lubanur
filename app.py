import streamlit as st
import pandas as pd
import datetime
import os
import re
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

DEFAULT_CARGO_COMPANIES = ["وصل لي", "نول", "جيناكم", "دليلي", "بدون شحن"]

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
    return orders

def parse_items_to_df(items_str):
    if not items_str or str(items_str).lower() == "nan":
        return pd.DataFrame(columns=["اسم المنتج", "الكمية"])
    parts = str(items_str).split(" | ")
    rows = []
    for part in parts:
        match = re.search(r'(.+?)\s*\(الكمية:\s*(\d+)\)', part)
        if match:
            item_name = match.group(1).strip()
            qty = int(match.group(2))
            rows.append({"اسم المنتج": item_name, "الكمية": qty})
        else:
            rows.append({"اسم المنتج": part.strip(), "الكمية": 1})
    return pd.DataFrame(rows)

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
        cargo_df = pd.DataFrame({"Cargo_Name": DEFAULT_CARGO_COMPANIES})
        
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            stock.to_excel(writer, sheet_name="Stock", index=False)
            expenses.to_excel(writer, sheet_name="Expenses", index=False)
            cargo_df.to_excel(writer, sheet_name="Cargo_Companies", index=False)
    else:
        xls = pd.ExcelFile(DB_FILE)
        orders = pd.read_excel(xls, sheet_name="Orders")
        orders = normalize_orders_df(orders)
        
        stock = pd.read_excel(xls, sheet_name="Stock") if "Stock" in xls.sheet_names else pd.DataFrame(columns=["Item_Name", "Qty"])
        expenses = pd.read_excel(xls, sheet_name="Expenses") if "Expenses" in xls.sheet_names else pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])
        cargo_df = pd.read_excel(xls, sheet_name="Cargo_Companies") if "Cargo_Companies" in xls.sheet_names else pd.DataFrame({"Cargo_Name": DEFAULT_CARGO_COMPANIES})

        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            stock.to_excel(writer, sheet_name="Stock", index=False)
            expenses.to_excel(writer, sheet_name="Expenses", index=False)
            cargo_df.to_excel(writer, sheet_name="Cargo_Companies", index=False)

init_db()

def load_data():
    st.cache_data.clear()
    xls = pd.ExcelFile(DB_FILE)
    orders = pd.read_excel(xls, sheet_name="Orders")
    orders = normalize_orders_df(orders)
    
    stock = pd.read_excel(xls, sheet_name="Stock")
    expenses = pd.read_excel(xls, sheet_name="Expenses") if "Expenses" in xls.sheet_names else pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])
    cargo_df = pd.read_excel(xls, sheet_name="Cargo_Companies") if "Cargo_Companies" in xls.sheet_names else pd.DataFrame({"Cargo_Name": DEFAULT_CARGO_COMPANIES})
    
    cargo_list = cargo_df['Cargo_Name'].dropna().astype(str).tolist() if not cargo_df.empty else DEFAULT_CARGO_COMPANIES
    if "وصل لي" in cargo_list:
        cargo_list.remove("وصل لي")
        cargo_list.insert(0, "وصل لي")
    elif "وصل لي" not in cargo_list:
        cargo_list.insert(0, "وصل لي")
        
    return orders, stock, expenses, cargo_list

orders_df, stock_df, expenses_df, cargo_companies_list = load_data()

st.sidebar.markdown('<p class="sidebar-title">🌟 لبانور - لوحة التحكم الإدارية</p>', unsafe_allow_html=True)
st.sidebar.markdown('<p class="version-tag">📌 الإصدار المحدث المصحح: v3:00pm 27-10-2026</p>', unsafe_allow_html=True)

if st.sidebar.button("🔄 تحديث البيانات مع السيرفر"):
    st.cache_data.clear()
    st.success("✅ تم تحديث ومزامنة البيانات مع السيرفر بنجاح!")
    st.rerun()

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

uploaded_backup = st.sidebar.file_uploader("📤 استرجاع بيانات من نسخة احتياطية (.xlsx)", type=["xlsx"])
if uploaded_backup is not None:
    try:
        with open(DB_FILE, "wb") as f:
            f.write(uploaded_backup.getbuffer())
        st.sidebar.success("✅ تم استرجاع البيانات بنجاح!")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"خطأ: {e}")

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

if menu == "اللوحة الرئيسية":
    st.header("📊 الملخص التشغيلي والمالي الشامل")
    orders_df, stock_df, expenses_df, _ = load_data()
    
    total_products_sales = orders_df['Products_Total'].sum() if not orders_df.empty and 'Products_Total' in orders_df.columns else 0.0
    total_shipping = orders_df['Shipping_Price'].sum() if not orders_df.empty and 'Shipping_Price' in orders_df.columns else 0.0
    total_expenses = expenses_df['Amount'].sum() if not expenses_df.empty and 'Amount' in expenses_df.columns else 0.0
    net_treasury = total_products_sales - total_expenses

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("إجمالي مبيعات المنتجات (إيراداتنا)", f"{total_products_sales:.3f} ر.ع")
    col2.metric("إجمالي قيمة الشحن (لشركة الشحن)", f"{total_shipping:.3f} ر.ع")
    col3.metric("إجمالي المصروفات", f"{total_expenses:.3f} ر.ع")
    col4.metric("صافي الخزينة", f"{net_treasury:.3f} ر.ع")
    
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
            else:
                items_summary_list = [f"{r['Item']} (الكمية: {r['Qty']})" for r in st.session_state.cart]
                items_detail_str = " | ".join(items_summary_list)
                grand_total = products_total + shipping_price
                coll_status = "مدفوع بالكامل" if (order_payment == "تم الدفع" and shipping_payment == "تم دفع الشحن") else "قيد التحصيل"
                
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
                        stock_df.loc[idx_st, 'Qty'] = max(0, int(stock_df.loc[idx_st, 'Qty']) - it_qty)

                orders_df = pd.concat([orders_df, pd.DataFrame([new_row])], ignore_index=True)
                
                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                    cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
                
                st.session_state.cart = []
                st.cache_data.clear()
                st.success(f"🎉 تم حفظ الطلب رقم #{next_order_id} بنجاح وتم خصم المخزون تلقائياً!")
                st.session_state.last_saved_order = next_order_id
                st.rerun()

    if 'last_saved_order' in st.session_state:
        saved_id = st.session_state.last_saved_order
        if saved_id in orders_df['Order_ID'].values:
            s_row = orders_df[orders_df['Order_ID'] == saved_id].iloc[0]
            st.markdown("---")
            st.subheader(f"🖨️ معاينة الفاتورة المصدرة للتو (طلب #{saved_id}):")
            
            raw_items = str(s_row['Items_Detail'])
            formatted_items_text = raw_items.replace(" | ", "<br>- ")
            if formatted_items_text and not formatted_items_text.startswith("-"):
                formatted_items_text = "- " + formatted_items_text
            barcode_str = f"<b>باركود الشحن:</b> {s_row.get('Cargo_Barcode', 'غير متوفر')}" if s_row.get('Cargo_Barcode') else ""

            receipt_html_snippet = f"""
            <!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8">
                <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                <style>
                    body {{ font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 10px; }}
                    .receipt-container {{ max-height: 750px; overflow-y: auto; border: 2px dashed #2c3e50; border-radius: 12px; padding: 20px; background-color: #ffffff; direction: rtl; text-align: right; box-shadow: 0 4px 10px rgba(0,0,0,0.1); max-width: 420px; margin: 0 auto; }}
                    h3 {{ text-align: center; color: #2c3e50; margin-bottom: 5px; }}
                    .tagline {{ text-align: center; color: #666; font-size: 12px; margin-top: 0px; border-bottom: 1px dashed #ddd; padding-bottom: 10px; }}
                    .info {{ font-size: 14px; color: #333; line-height: 1.6; }}
                    .items {{ font-size: 13px; color: #333; background: #f9f9f9; padding: 10px; border-radius: 6px; line-height: 1.6; margin: 10px 0; }}
                    .total {{ font-size: 16px; color: #007bff; font-weight: bold; text-align: center; margin-top: 10px; }}
                    .btn-container {{ display: flex; gap: 8px; justify-content: center; margin-top: 15px; flex-wrap: wrap; max-width: 420px; margin-left: auto; margin-right: auto; }}
                    .btn {{ padding: 10px 14px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; text-decoration: none; font-size: 13px; color: white; }}
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
                        <p><b>رقم الفاتورة:</b> #{s_row['Order_ID']} &nbsp;|&nbsp; <b>التاريخ:</b> <span style="direction: ltr; display: inline-block;">{s_row['Date']}</span></p>
                        <p><b>العميل:</b> {s_row['Customer']} &nbsp;|&nbsp; <b>الهاتف:</b> <span style="direction: ltr; display: inline-block;">{s_row['Phone']}</span></p>
                        <p><b>المنطقة:</b> {s_row['Region']}</p>
                        <p><b>نوع التسليم:</b> {s_row.get('Delivery_Type', 'غير محدد')} &nbsp;|&nbsp; <b>شركة الشحن:</b> {s_row.get('Cargo_Company', 'غير محدد')}</p>
                        <p><b>حالة دفع الطلب:</b> {s_row.get('Order_Payment', 'لم يتم الدفع')} &nbsp;|&nbsp; <b>دفع الشحن:</b> {s_row.get('Shipping_Payment', 'لم يتم دفع الشحن')}</p>
                        <p>{barcode_str}</p>
                    </div>
                    <hr style="border: 0.5px dashed #ccc;">
                    <strong>المنتجات:</strong>
                    <div class="items">{formatted_items_text}</div>
                    <hr style="border: 0.5px dashed #ccc;">
                    <div class="info">
                        <p><b>قيمة المنتجات:</b> {s_row['Products_Total']} ر.ع</p>
                        <p><b>سعر الشحن:</b> {s_row['Shipping_Price']} ر.ع</p>
                        <div class="total">الإجمالي الكلي: {s_row['Total']} ر.ع</div>
                    </div>
                    <p style="text-align: center; font-weight: bold; color: #2c3e50; margin-top: 15px; border-top: 1px dashed #ddd; padding-top: 10px;">شكراً لتعاملكم معنا!</p>
                </div>
                
                <div class="btn-container">
                    <button class="btn btn-copy" onclick="copyImage()">📋 نسخ الصورة</button>
                    <button class="btn btn-download" onclick="downloadImage()">📥 تحميل كصورة</button>
                    <button class="btn btn-print" onclick="window.print()">🖨️ طباعة</button>
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
                                        document.getElementById('toast-msg').innerText = '✅ تم النسخ بنجاح!';
                                    }}).catch(err => {{
                                        alert('استخدم زر "تحميل كصورة" أدناه للأجهزة اللوحية والأيباد.');
                                    }});
                                }} catch(e) {{
                                    alert('استخدم زر "تحميل كصورة" أدناه للأجهزة اللوحية والأيباد.');
                                }});
                            }});
                        }});
                    }}
                    function downloadImage() {{
                        const el = document.getElementById('receipt-capture');
                        html2canvas(el, {{ scale: 2, useCORS: true, scrollY: -window.scrollY }}).then(canvas => {{
                            const link = document.createElement('a');
                            link.download = 'Invoice_#{s_row["Order_ID"]}.png';
                            link.href = canvas.toDataURL('image/png');
                            link.click();
                            document.getElementById('toast-msg').innerText = '✅ تم تحميل صورة الفاتورة لجهازك!';
                        }});
                    }}
                </script>
            </body>
            </html>
            """
            st.components.v1.html(receipt_html_snippet, height=740)

elif menu == "تتبع الشحنات والتحصيل":
    st.header("🚚 تقرير تتبع الشحنات ومتابعة التحصيل المالي")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    
    if not orders_df.empty:
        col_t1, col_t2, col_t3 = st.columns(3)
        with col_t1:
            sel_cargo_track = st.selectbox("اختر شركة الشحن", ["الكل (جميع الشركات)"] + cargo_companies_list)
        with col_t2:
            sel_delivery_status = st.selectbox("حالة التوصيل", ["الكل", "قيد التحضير", "تم التسليم لشركة الشحن", "جاري التوصيل", "تم تسليم العميل", "لم يتم التوصيل", "مرتجع"])
        with col_t3:
            sel_collection_status = st.selectbox("حالة التحصيل المالي", ["الكل", "قيد التحصيل", "مدفوع بالكامل"])
            
        track_df = orders_df.copy()
        
        if sel_cargo_track != "الكل (جميع الشركات)":
            track_df = track_df[track_df['Cargo_Company'] == sel_cargo_track]
        if sel_delivery_status != "الكل":
            track_df = track_df[track_df['Delivery_Status'] == sel_delivery_status]
        if sel_collection_status != "الكل":
            track_df = track_df[track_df['Collection_Status'] == sel_collection_status]
            
        total_products_due = track_df[track_df['Collection_Status'] == 'قيد التحصيل']['Products_Total'].sum()
        total_undelivered = len(track_df[track_df['Delivery_Status'] != 'تم تسليم العميل'])
        
        m_c1, m_c2, m_c3 = st.columns(3)
        m_c1.metric("عدد الشحنات المعروضة", len(track_df))
        m_c2.metric("إجمالي المبالغ المستحقة (قيد التحصيل)", f"{total_products_due:.3f} ر.ع")
        m_c3.metric("الطلبيات التي لم تسلم بعد", total_undelivered)
        
        st.markdown("---")
        st.subheader("📋 تفاصيل الطلبيات والشحنات:")
        if not track_df.empty:
            display_cols = ['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Company', 'Cargo_Barcode', 'Delivery_Status', 'Collection_Status', 'Products_Total', 'Shipping_Price', 'Total', 'Bags']
            st.dataframe(track_df[display_cols], use_container_width=True)
        else:
            st.info("لا توجد طلبيات مطابقة لشروط البحث المحددة.")
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
        date_orders = pick_df[pick_df['Date_Parsed'] == pick_date]
        
        if not date_orders.empty:
            st.info(f"إجمالي الطلبيات في تاريخ {pick_date}: {len(date_orders)}")
            
            for idx, order in date_orders.iterrows():
                order_id = order['Order_ID']
                cust = order['Customer']
                phone = order['Phone']
                region = order['Region']
                current_d_status = order['Delivery_Status']
                
                with st.expander(f"📦 طلب رقم #{order_id} - العميل: {cust} ({region}) | الحالة الحالية: {current_d_status}"):
                    st.write(f"📞 **الهاتف:** {phone} &nbsp;|&nbsp; 🏷️ **عدد الأكياس:** {order['Bags']} &nbsp;|&nbsp; 🚚 **شركة الشحن:** {order['Cargo_Company']}")
                    st.markdown("---")
                    st.write("✅ **أصناف الطلبية (قم بوضع علامة صح عند تجهيز كل صنف):**")
                    
                    items_df = parse_items_to_df(order['Items_Detail'])
                    
                    packed_all = True
                    for i, item_row in items_df.iterrows():
                        iname = item_row['اسم المنتج']
                        iqty = item_row['الكمية']
                        
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
                        st.success("🎉 تم تجهيز جميع أصناف هذه الطلبية بالكامل!")
                        if current_d_status == "قيد التحضير":
                            if st.button(f"🚀 تحويل حالة الطلب #{order_id} إلى جاري استلام شركة الشحن", key=f"btn_ready_{order_id}"):
                                idx_ord = orders_df[orders_df['Order_ID'] == order_id].index[0]
                                orders_df.loc[idx_ord, 'Delivery_Status'] = "تم التسليم لشركة الشحن"
                                
                                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                                    cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
                                
                                st.cache_data.clear()
                                st.success(f"✅ تم تحديث حالة الطلب #{order_id} بنجاح!")
                                st.rerun()
                    else:
                        st.info("💡 يرجى تجهيز كافة الأصناف ووضع علامة صح عليها لتحديث حالة الطلبية.")
        else:
            st.warning("لا توجد طلبيات مسجلة في التاريخ المحدد.")
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
            r_note = st.text_input("سبب العملية / ملاحظات")
            
            if st.form_submit_button("💾 تنفيذ وتحديث المخزون"):
                idx_st = stock_df[stock_df['Item_Name'] == r_item].index[0]
                cur_q = int(stock_df.loc[idx_st, 'Qty'])
                if r_action == "إضافة إلى المخزون (تعزيز)":
                    stock_df.loc[idx_st, 'Qty'] = cur_q + r_qty
                else:
                    stock_df.loc[idx_st, 'Qty'] = max(0, cur_q - r_qty)
                
                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                    cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
                
                st.cache_data.clear()
                st.success(f"✅ تم تحديث مخزون المنتج ({r_item}) بنجاح!")
                st.rerun()
                
    with tab2:
        st.subheader("تسوية المخزون بعد الجرد الفعلي (حساب العجز والزيادة)")
        st.markdown("💡 قم بإدخال الكمية الفعلية الموجودة في المخزون بعد عملية الجرد وسيقوم النظام بحساب الفارق (العجز أو الزيادة) وتحديث الرصيد.")
        
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
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                    cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
                
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
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                    cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
                
                st.cache_data.clear()
                st.success(f"✅ تمت إضافة شركة الشحن '{new_cargo_name}' بنجاح!")
                st.rerun()
            else:
                st.error("الاسم فارغ أو موجود مسبقاً.")

elif menu == "تقرير الشحن":
    st.header("📑 تقرير تسليم الشحنات لشركات الشحن (مع خيارات التصدير والطباعة)")
    orders_df, _, _, cargo_companies_list = load_data()
    
    if not orders_df.empty:
        col_rep_c1, col_rep_c2, col_rep_c3 = st.columns(3)
        with col_rep_c1:
            selected_cargo = st.selectbox("اختر شركة الشحن", cargo_companies_list, key="rep_cargo_sel")
        with col_rep_c2:
            use_date_filter = st.checkbox("تصفية حسب التاريخ المحدد", value=False)
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
                "الطلب مدفوع؟": filtered_cargo_df['Order_Payment'],
                "الشحن مدفوع؟": filtered_cargo_df['Shipping_Payment'],
                "باركود الشحن": filtered_cargo_df['Cargo_Barcode'],
                "إجمالي المبلغ (ر.ع)": filtered_cargo_df['Total']
            })

            st.markdown(f"### 📋 بيان شحنات شركة: **{selected_cargo}** (إجمالي الشحنات: {len(report_display_df)})")
            st.dataframe(report_display_df, use_container_width=True)

            html_report_card = f"""
            <!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8">
                <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                <style>
                    body {{ font-family: Tahoma, sans-serif; background: #fff; direction: rtl; padding: 20px; color: #333; }}
                    .report-box {{ background: #fff; padding: 20px; border-radius: 8px; border: 1px solid #ddd; max-width: 900px; margin: 0 auto; }}
                    h2 {{ text-align: center; color: #2c3e50; margin-bottom: 5px; }}
                    .sub-header {{ text-align: center; color: #666; font-size: 13px; margin-bottom: 20px; border-bottom: 2px solid #007bff; padding-bottom: 10px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }}
                    th, td {{ border: 1px solid #bdc3c7; padding: 8px; text-align: center; }}
                    th {{ background-color: #2c3e50; color: white; }}
                    tr:nth-child(even) {{ background-color: #f9f9f9; }}
                    .actions {{ display: flex; gap: 10px; justify-content: center; margin-top: 20px; }}
                    .btn {{ padding: 10px 18px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; color: white; font-size: 13px; }}
                    .btn-print {{ background-color: #28a745; }}
                    .btn-download {{ background-color: #ff9800; }}
                    .toast {{ text-align: center; color: green; font-weight: bold; margin-top: 10px; font-size: 13px; }}
                </style>
            </head>
            <body>
                <div id="capture-area" class="report-box">
                    <h2>تقرير تسليم الشحنات - شركة {selected_cargo}</h2>
                    <div class="sub-header">لبانور / السيف والضياء للاستثمار الدولي &nbsp;|&nbsp; التاريخ: {datetime.date.today()}</div>
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>رقم الفاتورة</th>
                                <th>التاريخ</th>
                                <th>اسم العميل</th>
                                <th>الهاتف</th>
                                <th>المنطقة</th>
                                <th>الأكياس</th>
                                <th>دفع الطلب</th>
                                <th>دفع الشحن</th>
                                <th>باركود الشحن</th>
                                <th>الإجمالي</th>
                            </tr>
                        </thead>
                        <tbody>
            """
            
            for i, (_, r) in enumerate(filtered_cargo_df.iterrows(), 1):
                html_report_card += f"""
                            <tr>
                                <td>{i}</td>
                                <td>#{r['Order_ID']}</td>
                                <td>{r['Date']}</td>
                                <td>{r['Customer']}</td>
                                <td>{r['Phone']}</td>
                                <td>{r['Region']}</td>
                                <td>{r['Bags']}</td>
                                <td>{r['Order_Payment']}</td>
                                <td>{r['Shipping_Payment']}</td>
                                <td>{r.get('Cargo_Barcode', '')}</td>
                                <td>{r['Total']} ر.ع</td>
                            </tr>
                """
            
            html_report_card += f"""
                        </tbody>
                    </table>
                </div>
                
                <div class="actions">
                    <button class="btn btn-print" onclick="window.print()">🖨️ طباعة أو تصدير PDF</button>
                    <button class="btn btn-download" onclick="downloadReportImage()">📥 تحميل التقرير كصورة</button>
                </div>
                <div id="toast" class="toast"></div>

                <script>
                    function downloadReportImage() {{
                        const el = document.getElementById('capture-area');
                        html2canvas(el, {{ scale: 2, useCORS: true }}).then(canvas => {{
                            const link = document.createElement('a');
                            link.download = 'Cargo_Report_{selected_cargo}.png';
                            link.href = canvas.toDataURL('image/png');
                            link.click();
                            document.getElementById('toast').innerText = '✅ تم تحميل صورة التقرير بنجاح!';
                        }});
                    }}
                </script>
            </body>
            </html>
            """
            st.components.v1.html(html_report_card, height=600)
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
        st.dataframe(period_result[['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Company', 'Items_Detail', 'Total']], use_container_width=True)
    else:
        st.info("لا توجد طلبيات مسجلة في النظام بعد.")

elif menu == "التعديل الجماعي الشامل":
    st.header("📊 التعديل الجماعي الفوري للطلبات (شيت إكسيل تفاعلي)")
    orders_df, stock_df, expenses_df, cargo_companies_list = load_data()
    if not orders_df.empty:
        edited_orders_df = st.data_editor(orders_df, num_rows="dynamic", use_container_width=True, key="batch_order_editor")
        if st.button("💾 حفظ جميع التعديلات دفعة واحدة"):
            try:
                edited_orders_df['Total'] = pd.to_numeric(edited_orders_df['Products_Total'], errors='coerce').fillna(0) + pd.to_numeric(edited_orders_df['Shipping_Price'], errors='coerce').fillna(0)
                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    edited_orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                    cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
                st.cache_data.clear()
                st.success("🎉 تم حفظ جميع التعديلات وتحديث قاعدة البيانات بنجاح!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ حدث خطأ: {e}")
    else:
        st.info("لا توجد طلبيات.")

elif menu == "فاتورة العميل":
    st.header("🖨️ فاتورة العميل (معاينة فورية وطباعة ونسخ كصورة لصق في واتساب)")
    orders_df, _, _, _ = load_data()
    if not orders_df.empty:
        st.markdown("📅 **تصفية الفواتير حسب التاريخ:**")
        receipt_filter_date = st.date_input("اختر تاريخ الفاتورة للبحث", datetime.date.today(), key="rec_filter_date")
        
        filtered_rec_df = orders_df.copy()
        filtered_rec_df['Date_Parsed'] = pd.to_datetime(filtered_rec_df['Date'], errors='coerce').dt.date
        date_matched_orders = filtered_rec_df[filtered_rec_df['Date_Parsed'] == receipt_filter_date]
        
        if not date_matched_orders.empty:
            order_options = date_matched_orders['Order_ID'].tolist()
            selected_order_id = st.selectbox("اختر الطلب / العميل لهذا التاريخ", order_options, format_func=lambda x: f"طلب #{x} - {date_matched_orders.loc[date_matched_orders['Order_ID']==x, 'Customer'].values[0]}")
            order_row = orders_df[orders_df['Order_ID'] == selected_order_id].iloc[0]
            
            raw_items = str(order_row['Items_Detail'])
            formatted_items_text = raw_items.replace(" | ", "<br>- ")
            if formatted_items_text and not formatted_items_text.startswith("-"):
                formatted_items_text = "- " + formatted_items_text

            barcode_str = f"<b>باركود الشحن:</b> {order_row.get('Cargo_Barcode', 'غير متوفر')}" if order_row.get('Cargo_Barcode') else ""

            receipt_html_snippet = f"""
            <!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8">
                <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                <style>
                    body {{ font-family: Tahoma, sans-serif; background: #fff; margin: 0; padding: 10px; }}
                    .receipt-container {{ max-height: 750px; overflow-y: auto; border: 2px dashed #2c3e50; border-radius: 12px; padding: 20px; background-color: #ffffff; direction: rtl; text-align: right; box-shadow: 0 4px 10px rgba(0,0,0,0.1); max-width: 420px; margin: 0 auto; }}
                    h3 {{ text-align: center; color: #2c3e50; margin-bottom: 5px; }}
                    .tagline {{ text-align: center; color: #666; font-size: 12px; margin-top: 0px; border-bottom: 1px dashed #ddd; padding-bottom: 10px; }}
                    .info {{ font-size: 14px; color: #333; line-height: 1.6; }}
                    .items {{ font-size: 13px; color: #333; background: #f9f9f9; padding: 10px; border-radius: 6px; line-height: 1.6; margin: 10px 0; }}
                    .total {{ font-size: 16px; color: #007bff; font-weight: bold; text-align: center; margin-top: 10px; }}
                    .btn-container {{ display: flex; gap: 8px; justify-content: center; margin-top: 15px; flex-wrap: wrap; max-width: 420px; margin-left: auto; margin-right: auto; }}
                    .btn {{ padding: 10px 14px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; text-decoration: none; font-size: 13px; color: white; }}
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
                    <button class="btn btn-print" onclick="window.print()">🖨️ طباعة</button>
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
                                        document.getElementById('toast-msg').innerText = '✅ تم النسخ بنجاح!';
                                    }}).catch(err => {{
                                        alert('استخدم زر "تحميل كصورة" أدناه للأجهزة اللوحية والأيباد.');
                                    }});
                                }} catch(e) {{
                                    alert('استخدم زر "تحميل كصورة" أدناه للأجهزة اللوحية والأيباد.');
                                }});
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
                            document.getElementById('toast-msg').innerText = '✅ تم تحميل صورة الفاتورة لجهازك!';
                        }});
                    }}
                </script>
            </body>
            </html>
            """
            st.components.v1.html(receipt_html_snippet, height=740)
        else:
            st.warning("لا توجد فواتير مسجلة في تاريخ اليوم المحدد.")

elif menu == "تعديل وحذف الفواتير":
    st.header("✏️ تعديل وحذف الفواتير مع جدول الأصناف الاحترافي")
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
            e_prod_tot = st.number_input("قيمة المنتجات", value=float(target_row['Products_Total']))
            e_ship_p = st.number_input("سعر الشحن", value=float(target_row['Shipping_Price']))
            e_barcode = st.text_input("باركود شركة الشحن", value=str(target_row.get('Cargo_Barcode', '')))
            
            cur_ord_pay = str(target_row.get('Order_Payment', 'لم يتم الدفع'))
            cur_shp_pay = str(target_row.get('Shipping_Payment', 'لم يتم دفع الشحن'))
            
            ord_pay_list = ["لم يتم الدفع", "تم الدفع"]
            o_idx = ord_pay_list.index(cur_ord_pay) if cur_ord_pay in ord_pay_list else 0
            e_ord_pay = st.selectbox("حالة دفع الطلب", ord_pay_list, index=o_idx)
            
            shp_pay_list = ["لم يتم دفع الشحن", "تم دفع الشحن"]
            s_idx = shp_pay_list.index(cur_shp_pay) if cur_shp_pay in shp_pay_list else 0
            e_shp_pay = st.selectbox("حالة دفع قيمة الشحن", shp_pay_list, index=s_idx)
            
            st.markdown("---")
            st.markdown("📦 **جدول أصناف الفاتورة (إضافة، تعديل، أو حذف أصناف):**")
            
            current_items_df = parse_items_to_df(target_row['Items_Detail'])
            edited_items_df = st.data_editor(current_items_df, num_rows="dynamic", use_container_width=True, key=f"editor_items_{edit_order_id}")
            
            if st.form_submit_button("💾 حفظ تعديلات الفاتورة والأصناف"):
                updated_items_list = []
                for _, row_it in edited_items_df.iterrows():
                    iname = str(row_it.get('اسم المنتج', '')).strip()
                    iqty = int(row_it.get('الكمية', 1))
                    if iname:
                        updated_items_list.append(f"{iname} (الكمية: {iqty})")
                
                new_items_detail_str = " | ".join(updated_items_list)
                
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
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Items_Detail'] = new_items_detail_str
                
                new_coll_status = "مدفوع بالكامل" if (e_ord_pay == "تم الدفع" and e_shp_pay == "تم دفع الشحن") else "قيد التحصيل"
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Collection_Status'] = new_coll_status

                cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                    cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
                st.cache_data.clear()
                st.success("✅ تم تعديل الفاتورة والأصناف وحفظ البيانات بنجاح!")
                st.rerun()
                
        if st.button("🗑️ حذف الفاتورة بالكامل"):
            orders_df = orders_df[orders_df['Order_ID'] != edit_order_id]
            cargo_df_save = pd.DataFrame({"Cargo_Name": cargo_companies_list})
            with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                orders_df.to_excel(writer, sheet_name="Orders", index=False)
                stock_df.to_excel(writer, sheet_name="Stock", index=False)
                expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
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
            with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                orders_df.to_excel(writer, sheet_name="Orders", index=False)
                stock_df.to_excel(writer, sheet_name="Stock", index=False)
                expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                cargo_df_save.to_excel(writer, sheet_name="Cargo_Companies", index=False)
            st.cache_data.clear()
            st.success("✅ تم حفظ المصروف بنجاح!")
            st.rerun()
    if not expenses_df.empty:
        st.dataframe(expenses_df, use_container_width=True)

elif menu == "التقرير المالي التفصيلي":
    st.header("📈 التقرير المالي التفصيلي والشامل")
    orders_df, stock_df, expenses_df, _ = load_data()
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

elif menu == "طلب توريد المخزون الحرِج":
    st.header("🚨 طلب توريد المخزون الحرِج")
    _, stock_df, _, _ = load_data()
    critical_stock = stock_df[stock_df['Qty'] < 5].copy()
    if not critical_stock.empty:
        st.dataframe(critical_stock, use_container_width=True)
    else:
        st.success("🎉 ممتاز! جميع المخزونات في مستويات آمنة.")
