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
            
        orders = pd.DataFrame(columns=["Order_ID", "Date", "Customer", "Phone", "Region", "Items_Detail", "Products_Total", "Shipping_Price", "Total", "Order_Payment", "Shipping_Payment", "Cargo_Company", "Delivery_Type", "Delivery_Status", "Collection_Status", "Bags"])
        expenses = pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])
        
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            stock.to_excel(writer, sheet_name="Stock", index=False)
            expenses.to_excel(writer, sheet_name="Expenses", index=False)
    else:
        xls = pd.ExcelFile(DB_FILE)
        orders = pd.read_excel(xls, sheet_name="Orders")
        stock = pd.read_excel(xls, sheet_name="Stock") if "Stock" in xls.sheet_names else pd.DataFrame(columns=["Item_Name", "Qty"])
        expenses = pd.read_excel(xls, sheet_name="Expenses") if "Expenses" in xls.sheet_names else pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])
        
        expected_cols = ["Order_ID", "Date", "Customer", "Phone", "Region", "Items_Detail", "Products_Total", "Shipping_Price", "Total", "Order_Payment", "Shipping_Payment", "Cargo_Company", "Delivery_Type", "Delivery_Status", "Collection_Status", "Bags"]
        for col in expected_cols:
            if col not in orders.columns:
                orders[col] = "غير محدد" if col in ["Items_Detail", "Order_Payment", "Shipping_Payment", "Cargo_Company", "Delivery_Type", "Delivery_Status", "Collection_Status"] else 0.0
        
        expected_exp_cols = ["Expense_ID", "Date", "Category", "Amount", "Description"]
        for col in expected_exp_cols:
            if col not in expenses.columns:
                expenses[col] = "" if col != "Amount" else 0.0

        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            stock.to_excel(writer, sheet_name="Stock", index=False)
            expenses.to_excel(writer, sheet_name="Expenses", index=False)

init_db()

def load_data():
    xls = pd.ExcelFile(DB_FILE)
    orders = pd.read_excel(xls, sheet_name="Orders")
    stock = pd.read_excel(xls, sheet_name="Stock")
    expenses = pd.read_excel(xls, sheet_name="Expenses") if "Expenses" in xls.sheet_names else pd.DataFrame(columns=["Expense_ID", "Date", "Category", "Amount", "Description"])
    return orders, stock, expenses

orders_df, stock_df, expenses_df = load_data()

st.sidebar.markdown('<p class="sidebar-title">🌟 لبانور - لوحة التحكم الإدارية</p>', unsafe_allow_html=True)

menu_options = {
    "📊 اللوحة الرئيسية (Dashboard & Financials)": "اللوحة الرئيسية (Dashboard & Financials)",
    "📝 تسجيل طلب جديد": "تسجيل طلب جديد (فاتورة متعددة الأصناف)",
    "⏳ الطلبات قيد التحضير": "الطلبات التي لم يتم تحضيرها (Pending Preparation)",
    "🚚 تتبع التوصيل والشحن": "تتبع التوصيل والشحن",
    "📑 تقرير الشحن (شركات الشحن)": "تقرير الشحن لشركة الشحن (تقرير الاستلام)",
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

# === 💾 قسم النسخ الاحتياطي في القائمة الجانبية ===
st.sidebar.markdown("---")
st.sidebar.subheader("💾 النسخ الاحتياطي للبيانات")
if os.path.exists(DB_FILE):
    with open(DB_FILE, "rb") as f:
        excel_bytes = f.read()
    current_date_str = datetime.date.today().strftime("%Y-%m-%d")
    st.sidebar.download_button(
        label="📥 تحميل نسخة احتياطية (Backup)",
        data=excel_bytes,
        file_name=f"LubaNur_Backup_{current_date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="اضغط هنا لتنزيل نسخة كاملة من قاعدة البيانات المحدثة على الأيباد أو جهازك"
    )

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
    
    total_sales = orders_df['Total'].sum() if not orders_df.empty and 'Total' in orders_df.columns else 0.0
    total_products_sales = orders_df['Products_Total'].sum() if not orders_df.empty and 'Products_Total' in orders_df.columns else total_sales
    total_shipping = orders_df['Shipping_Price'].sum() if not orders_df.empty and 'Shipping_Price' in orders_df.columns else 0.0
    total_expenses = expenses_df['Amount'].sum() if not expenses_df.empty and 'Amount' in expenses_df.columns else 0.0
    net_treasury = total_products_sales - total_expenses

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("إجمالي مبيعات المنتجات", f"{total_products_sales:.3f} ر.ع")
    col2.metric("إجمالي قيمة الشحن", f"{total_shipping:.3f} ر.ع")
    col3.metric("إجمالي المصروفات", f"{total_expenses:.3f} ر.ع")
    col4.metric("الخزينة (صافي السيولة)", f"{net_treasury:.3f} ر.ع", delta="بعد خصم المصروفات")
    
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
        order_payment = st.selectbox("حالة دفع الاوردر (Order Payment Status)", ["تم الدفع", "لم يتم الدفع (COD)"], key="new_ord_pay")
        shipping_payment = st.selectbox("حالة دفع الشحن (Shipping Payment Status)", ["تم دفع الشحن", "لم يتم دفع الشحن"], key="new_shp_pay")
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
                        
                        new_order_row = {
                            "Order_ID": new_id, "Date": str(order_date), "Customer": cust_name, "Phone": phone, "Region": region,
                            "Items_Detail": items_detail_str, "Products_Total": products_total, "Shipping_Price": shipping_price, "Total": grand_total,
                            "Order_Payment": order_payment, "Shipping_Payment": shipping_payment, "Cargo_Company": cargo_company, "Delivery_Type": delivery_type,
                            "Delivery_Status": "قيد التحضير", "Collection_Status": "قيد التحصيل" if (order_payment == "لم يتم الدفع (COD)" or shipping_payment == "لم يتم دفع الشحن") else "مدفوع بالكامل", "Bags": bags
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
            new_pend_status = st.selectbox("الحالة الجديدة", ["تم التسليم لشركة الشحن", "جاري التوصيل", "تم تسليم العميل", "مرتجع"])
            if st.button("💾 تحديث حالة الطلب"):
                orders_df.loc[orders_df['Order_ID'] == up_pend_id, 'Delivery_Status'] = new_pend_status
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    orders_df.to_excel(writer, sheet_name="Orders", index=False)
                    stock_df.to_excel(writer, sheet_name="Stock", index=False)
                    expenses_df.to_excel(writer, sheet_name="Expenses", index=False)
                st.success("✅ تم التحديث بنجاح!")
                st.rerun()
        else:
            st.success("🎉 لا توجد طلبيات معلقة!")

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
    total_sales = orders_df['Total'].sum() if not orders_df.empty else 0.0
    total_products_value = orders_df['Products_Total'].sum() if not orders_df.empty else 0.0
    total_shipping = orders_df['Shipping_Price'].sum() if not orders_df.empty else 0.0
    total_expenses = expenses_df['Amount'].sum() if not expenses_df.empty else 0.0
    net_treasury = total_products_value - total_expenses

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("إجمالي المبيعات", f"{total_sales:.3f} ر.ع")
    m2.metric("قيمة المنتجات", f"{total_products_value:.3f} ر.ع")
    m3.metric("قيمة الشحن", f"{total_shipping:.3f} ر.ع")
    m4.metric("المصروفات", f"{total_expenses:.3f} ر.ع")
    st.info(f"💰 **صافي الخزينة:** {net_treasury:.3f} ريال عماني")
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
        st.code(f"""
=== لبانور - LubaNur ===
رقم الفاتورة: #{order_row['Order_ID']} | التاريخ: {order_row['Date']}
العميل: {order_row['Customer']} | الهاتف: {order_row['Phone']}
المنطقة: {order_row['Region']}
-----------------------------------
المنتجات:
{order_row['Items_Detail']}
-----------------------------------
قيمة المنتجات: {order_row['Products_Total']} ر.ع
سعر الشحن: {order_row['Shipping_Price']} ر.ع
الإجمالي الكلي: {order_row['Total']} ر.ع
شكراً لتعاملكم معنا!
        """)

elif menu == "تعديل وحذف الفواتير (Edit & Delete)":
    st.header("✏️ تعديل وحذف الفواتير")
    orders_df, stock_df, expenses_df = load_data()
    if not orders_df.empty:
        edit_order_id = st.selectbox("اختر رقم الطلب للتعديل", orders_df['Order_ID'].tolist())
        target_row = orders_df[orders_df['Order_ID'] == edit_order_id].iloc[0]
        with st.form("edit_form"):
            e_cust = st.text_input("اسم العميل", value=str(target_row['Customer']))
            e_phone = st.text_input("الهاتف", value=str(target_row['Phone']))
            e_prod_tot = st.number_input("قيمة المنتجات", value=float(target_row['Products_Total']))
            e_ship_p = st.number_input("سعر الشحن", value=float(target_row['Shipping_Price']))
            if st.form_submit_button("💾 حفظ التعديلات"):
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Customer'] = e_cust
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Phone'] = e_phone
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Products_Total'] = e_prod_tot
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Shipping_Price'] = e_ship_p
                orders_df.loc[orders_df['Order_ID'] == edit_order_id, 'Total'] = e_prod_tot + e_ship_p
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
            st.success("🗑️ تم حذف الفاتورة!")
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
        st.success("✅ تم تحديث المخزون!")
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
        st.dataframe(orders_df[['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Cargo_Company', 'Delivery_Type', 'Delivery_Status']], use_container_width=True)

elif menu == "تقرير الشحن لشركة الشحن (تقرير الاستلام)":
    st.header("📑 تقرير تسليم الشحنات لشركات الشحن")
    orders_df, _, _ = load_data()
    if not orders_df.empty:
        selected_cargo = st.selectbox("اختر شركة الشحن", orders_df['Cargo_Company'].unique().tolist())
        st.dataframe(orders_df[orders_df['Cargo_Company'] == selected_cargo][['Order_ID', 'Date', 'Customer', 'Phone', 'Region', 'Products_Total', 'Shipping_Price', 'Total', 'Bags']], use_container_width=True)

elif menu == "متابعة التحصيل المالي":
    st.header("💵 متابعة التحصيل المالي")
    orders_df, stock_df, expenses_df = load_data()
    if not orders_df.empty:
        st.dataframe(orders_df[['Order_ID', 'Date', 'Customer', 'Phone', 'Cargo_Company', 'Products_Total', 'Collection_Status']], use_container_width=True)
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