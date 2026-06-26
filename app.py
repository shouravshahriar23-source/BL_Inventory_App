import streamlit as st
import pandas as pd
import os
import plotly.express as px
from io import BytesIO

# ফোল্ডার কনফিগারেশন
FOLDER_PATH = "uploaded_files"
if not os.path.exists(FOLDER_PATH):
    os.makedirs(FOLDER_PATH)

st.set_page_config(page_title="Banglalink DD House Inventory & Product Analyzer", layout="wide")

# --- ১. মাল্টি-ইউজার লগইন সিস্টেম ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["user_role"] = "Viewer"  # ডিফল্ট রোল ভিউয়ার

# ইউজারনেম এবং পাসওয়ার্ডের তালিকা (এখানে ইচ্ছামতো ইউজার বাড়ানো যাবে)
user_credentials = {
    "Shourav": {"password": "admin123", "role": "Admin"},   # এডমিন (ফাইল আপলোড করতে পারবে)
    "Viewer": {"password": "mymtan13", "role": "Viewer"}   # স্টাফ/অন্যান্যরা (শুধু দেখতে পারবে)
}

if not st.session_state["logged_in"]:
    st.markdown("<br><br><h2 style='text-align: center; color: #FF6600;'>🔐 Banglalink Inventory System Login</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<div style='background-color: #f9f9f9; padding: 25px; border-radius: 10px; border: 1px solid #ddd;'>", unsafe_allow_html=True)
        username = st.text_input("ইউজারনেম (Username):", key="user_input")
        password = st.text_input("পাসওয়ার্ড (Password):", type="password", key="pass_input")
        login_btn = st.button("লগইন করুন", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        if login_btn:
            if username in user_credentials and password == user_credentials[username]["password"]:
                st.session_state["logged_in"] = True
                st.session_state["user_role"] = user_credentials[username]["role"]  # রোল সেভ হবে (Admin নাকি Viewer)
                st.rerun()
            else:
                st.error("❌ ভুল ইউজারনেম অথবা পাসওয়ার্ড! আবার চেষ্টা করুন।")
    st.stop()

# --- লগইন সফল হলে নিচের মূল কোডটি রান হবে ---

st.markdown("<h1 style='text-align: center; color: #FF6600;'>📶 Banglalink DD House MYMTAN13</h1>", unsafe_allow_html=True)

# ফাইল লোড করা (শুধুমাত্র .xlsx ফাইল প্রসেস করা হবে)
files = [f for f in os.listdir(FOLDER_PATH) if f.endswith(".xlsx")]
files.sort(reverse=True)

mode = st.sidebar.radio("মোড সিলেক্ট করুন:", ["দৈনিক রিপোর্ট", "মাসিক সামারি"])


# --- ২. শুধুমাত্র Admin (Shourav) এর জন্য ফাইল আপলোডার প্রোটেকশন ---
if st.session_state["user_role"] == "Admin":
    st.sidebar.markdown("---")
    st.sidebar.markdown("<b style='color: #FF6600;'>🛠️ অ্যাডমিন প্যানেল (ফাইল আপলোড)</b>", unsafe_allow_html=True)
    uploaded_file = st.sidebar.file_uploader("📤 নতুন xlsx ফাইল আপলোড করুন", type=["xlsx"])

    if uploaded_file is not None:
        file_path = os.path.join(FOLDER_PATH, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success(f"✅ {uploaded_file.name} successfully upload hoyeche!")
        st.rerun()


# Helper function: Excel data clean & dynamic row extraction
def get_clean_section(df, start_keyword, num_rows=12, cols_range=(0, 4)):
    idx = df[df.iloc[:, 0].astype(str).str.contains(start_keyword, na=False)].index
    if not idx.empty:
        start_idx = idx[0] + 2
        sub_df = df.iloc[start_idx:start_idx+num_rows, cols_range[0]:cols_range[1]].copy()
        sub_df = sub_df.dropna(subset=[sub_df.columns[0]])
        
        # প্রোডাক্ট বা কর্মীর তালিকায় হেডライン বা অপ্রয়োজনীয় রো বাদ দেওয়ার ফিল্টার
        invalid_keywords = [
            "Total", "সর্বমোট", "প্রোডাক্টের নাম", "Product Name", "Product", 
            "মোট প্রোডাক্ট", "Total Product Profit", "কার কাছে বর্তমানে", 
            "SR Wise", "Breakdown", "মার্কেট হোল্ড সামারি", "Overall Sales"
        ]
        pattern = "|".join(invalid_keywords)
        sub_df = sub_df[~sub_df.iloc[:, 0].astype(str).str.contains(pattern, case=False, na=False)]
        return sub_df
    return pd.DataFrame()

# এক্সেল কনভার্ট করার হেল্পার ফাংশন
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Summary_Report')
    return output.getvalue()


if not files:
    st.info("💡 অনুগ্রহ করে `uploaded_files` ফোল্ডারে আপনার এক্সেল (.xlsx) ফাইলগুলো রাখুন।")

# --- দৈনিক রিপোর্ট লজিক ---
elif mode == "দৈনিক রিপোর্ট":
    selected_file = st.sidebar.selectbox("📅 ফাইল সিলেক্ট করুন:", files)
    if selected_file:
        file_path = os.path.join(FOLDER_PATH, selected_file)
        
        # দৈনিক রিপোর্টের উপরে তারিখ প্রদর্শন
        clean_date = selected_file.replace("Final Daily House  Inventory Master Tracker ", "").replace(".xlsx", "").strip()
        st.markdown(f"<h3 style='text-align: center; color: #444444;'>📅 রিপোর্টের তারিখ: {clean_date}</h3>", unsafe_allow_html=True)
        st.markdown("---")
        
        def read_sheet(sheet_name):
            return pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        tab1, tab2, tab3 = st.tabs(["📊 ড্যাশবোর্ড", "📝 সেলস পার্সন এন্ট্রি", "💰 লাভ ও অ্যাকাউন্টস"])

        with tab1:
            try:
                df_dash = read_sheet('Dashboard Summary')
                
                # সেকশন ১
                st.subheader("১. হাউজের মূল স্টক")
                df_stock = get_clean_section(df_dash, "হাউজের মূল স্টক", 12, (0, 4))
                df_stock.columns = ['প্রোডাক্ট', 'ওপেনিং', 'SR ইস্যু', 'বর্তমান স্টক']
                st.dataframe(df_stock, use_container_width=True, hide_index=True)
                
                # সেকশন ২
                st.subheader("২. প্রোডাক্ট অনুযায়ী মোট বিক্রি ও মার্কেট হোল্ড")
                df_sales = get_clean_section(df_dash, "মোট বিক্রি ও মার্কেট হোল্ড", 12, (0, 4))
                df_sales.columns = ['প্রোডাক্ট', 'মোট বিক্রি', 'মার্কেট হোল্ড', 'SR থেকে ফেরত']
                for col in ['মোট বিক্রি', 'মার্কেট হোল্ড', 'SR থেকে ফেরত']:
                    df_sales[col] = pd.to_numeric(df_sales[col], errors='coerce').fillna(0).astype(int)
                st.dataframe(df_sales, use_container_width=True, hide_index=True)
                
                # সেকশন ৩
                st.subheader("৩. কার কাছে বর্তমানে কি প্রোডাক্ট রয়েছে (SR Wise Closing Stock)")
                idx_sr = df_dash[df_dash.iloc[:, 0].astype(str).str.contains("কার কাছে বর্তমানে", na=False)].index
                if not idx_sr.empty:
                    start_sr = idx_sr[0] + 1
                    sr_cols = df_dash.iloc[start_sr].dropna().tolist()
                    
                    df_sr_stock = df_dash.iloc[start_sr+1 : start_sr+15, 0:len(sr_cols)].copy()
                    df_sr_stock.columns = sr_cols
                    df_sr_stock = df_sr_stock.dropna(subset=[sr_cols[0]])
                    
                    df_sr_stock = df_sr_stock[~df_sr_stock.iloc[:, 0].astype(str).str.contains("কার কাছে|SR-এর নাম", na=False)]
                    st.dataframe(df_sr_stock, use_container_width=True, hide_index=True)
                
                # চার্ট
                fig = px.bar(df_sales, x='প্রোডাক্ট', y='মোট বিক্রি', text='মোট বিক্রি', title="আজকের প্রোডাক্ট বিক্রির চার্ট")
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.error(f"ড্যাশবোর্ড লোড করতে সমস্যা হয়েছে: {e}")

        with tab2:
            try:
                df_entry = read_sheet('Daily Data Entry')
                mask = df_entry.iloc[:, 0].astype(str).str.contains("RSO|Supervisor|BP", case=False, na=False)
                all_indices = df_entry[mask].index.tolist()
                
                person_names = [str(df_entry.iloc[i, 0]).strip() for i in all_indices]
                selected_person = st.selectbox("সেলস পার্সন সিলেক্ট করুন:", person_names)
                
                idx = all_indices[person_names.index(selected_person)]
                person_data = df_entry.iloc[idx+2:idx+12, 0:7].copy()
                person_data.columns = ['প্রোডাক্ট', 'Opening', 'Issue', 'Total', 'Sales', 'Return', 'Closing']
                person_data = person_data.dropna(subset=['প্রোডাক্ট'])
                
                invalid_keywords = ["প্রোডাক্টের নাম", "Product Name", "Product"]
                person_data = person_data[~person_data['প্রোডাক্ট'].astype(str).str.contains("|".join(invalid_keywords), na=False)]
                
                numeric_cols = ['Opening', 'Issue', 'Total', 'Sales', 'Return', 'Closing']
                for col in numeric_cols:
                    person_data[col] = pd.to_numeric(person_data[col], errors='coerce').fillna(0).astype(int)
                
                st.dataframe(person_data, use_container_width=True, hide_index=True)
                
                if person_data['Sales'].sum() > 0:
                    fig2 = px.pie(person_data, values='Sales', names='প্রোডাক্ট', title=f"{selected_person}-এর সেলস ডিস্ট্রিবিউশন")
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.warning("⚠️ এই কর্মীর আজ কোনো বিক্রি নেই।")
            except Exception as e:
                st.error(f"সেলস পার্সন ডেটা লোড করতে সমস্যা হয়েছে: {e}")

        with tab3:
            try:
                df_profit = read_sheet('Daily Profit & Accounts')
                df_profit_clean = df_profit.iloc[3:14, 0:8].copy()
                df_profit_clean.columns = ['প্রোডাক্ট', 'বিক্রি সংখ্যা', 'কেনা দাম', 'বিক্রি দাম', 'লেস', 'মোট কেনা', 'মোট বিক্রি', 'লাভ']
                df_profit_clean = df_profit_clean.dropna(subset=['প্রোডাক্ট'])
                
                invalid_keywords = ["মোট প্রোডাক্ট প্রফিট", "Product Name", "Product", "Total"]
                df_profit_clean = df_profit_clean[~df_profit_clean['প্রোডাক্ট'].astype(str).str.contains("|".join(invalid_keywords), na=False)]
                
                for col in ['বিক্রি সংখ্যা', 'মোট কেনা', 'মোট বিক্রি', 'লাভ']:
                    df_profit_clean[col] = pd.to_numeric(df_profit_clean[col], errors='coerce').fillna(0).astype(float)
                
                st.subheader("💰 আর্থিক সামারি (লাভ ও লস)")
                st.dataframe(df_profit_clean, use_container_width=True, hide_index=True)
                
                fig3 = px.bar(df_profit_clean, x='প্রোডাক্ট', y=['মোট কেনা', 'মোট বিক্রি', 'লাভ'], 
                              title="প্রোডাক্ট অনুযায়ী লাভ ও খরচের তুলনামূলক গ্রাফ", barmode='group')
                st.plotly_chart(fig3, use_container_width=True)

                # দৈনিক রিপোর্টের লাভ-ক্ষতি এক্সেল ডাউনলোড বাটন
                daily_excel_data = convert_df_to_excel(df_profit_clean)
                st.download_button(
                    label="📥 আজকের লাভ-ক্ষতির রিপোর্ট ডাউনলোড করুন (Excel)",
                    data=daily_excel_data,
                    file_name=f"Daily_Profit_Report_{clean_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"আর্থিক খতিয়ান লোড করতে সমস্যা হয়েছে: {e}")

# --- মাসিক সামারি লজিক ---
else:
    st.subheader("📅 মাসিক সামারি রিপোর্ট")
    selected_month = st.sidebar.text_input("মাস লিখুন (যেমন: 26-06):", "26-06")
    
    if selected_month:
        monthly_files = [f for f in files if selected_month in f]
        
        if monthly_files:
            st.markdown(f"<h3 style='text-align: center; color: #444444;'>📅 মাসিক সামারি রিপোর্ট: {selected_month}</h3>", unsafe_allow_html=True)
            st.markdown("---")

            st.success(f"📊 {selected_month} মাসের মোট {len(monthly_files)} টি ফাইল পাওয়া গেছে।")
            
            m_sales_list = []
            m_person_list = []
            m_profit_list = []
            
            for f in monthly_files:
                try:
                    file_path = os.path.join(FOLDER_PATH, f)
                    
                    df_d = pd.read_excel(file_path, sheet_name='Dashboard Summary', header=None)
                    df_e = pd.read_excel(file_path, sheet_name='Daily Data Entry', header=None)
                    df_p = pd.read_excel(file_path, sheet_name='Daily Profit & Accounts', header=None)
                    
                    # ১. সেলস সামারি ডেটা
                    sl_df = get_clean_section(df_d, "মোট বিক্রি ও মার্কেট হোল্ড", 12, (0, 2))
                    if not sl_df.empty:
                        sl_df.columns = ['প্রোডাক্ট', 'মোট বিক্রি']
                        m_sales_list.append(sl_df)
                        
                    # ২. পার্সন ওয়াইজ ডেটা
                    mask = df_e.iloc[:, 0].astype(str).str.contains("RSO|Supervisor|BP", case=False, na=False)
                    all_indices = df_e[mask].index.tolist()
                    for idx in all_indices:
                        p_name = str(df_e.iloc[idx, 0]).strip()
                        p_data = df_e.iloc[idx+2:idx+12, 0:7].copy()
                        p_data.columns = ['প্রোডাক্ট', 'Opening', 'Issue', 'Total', 'Sales', 'Return', 'Closing']
                        p_data = p_data.dropna(subset=['প্রোডাক্ট'])
                        
                        p_data = p_data[~p_data['প্রোডাক্ট'].astype(str).str.contains("প্রোডাক্টের নাম|Product Name|Product", na=False)]
                        p_data['Name'] = p_name
                        m_person_list.append(p_data)
                        
                    # ৩. প্রফিট ডেটা
                    p_df = df_p.iloc[3:14, 0:8].copy()
                    p_df.columns = ['প্রোডাক্ট', 'বিক্রি সংখ্যা', 'কেনা দাম', 'বিক্রি দাম', 'লেস', 'মোট কেনা', 'মোট বিক্রি', 'লাভ']
                    p_df = p_df.dropna(subset=['প্রোডাক্ট'])
                    p_df = p_df[~p_df['প্রোডাক্ট'].astype(str).str.contains("মোট প্রোডাক্ট প্রফিট|Product Name|Product|Total", na=False)]
                    m_profit_list.append(p_df)
                    
                except Exception as e:
                    st.error(f"ফাইল {f} প্রসেস করতে সমস্যা: {e}")
                    
            mtab1, mtab2, mtab3 = st.tabs(["📊 মাসিক মোট বিক্রি", "👤 কর্মীভিত্তিক মাসিক পারফরম্যান্স", "💰 মাসিক লাভ-ক্ষতি"])
            
            # Tab 1
            with mtab1:
                if m_sales_list:
                    all_sales_df = pd.concat(m_sales_list)
                    all_sales_df['মোট বিক্রি'] = pd.to_numeric(all_sales_df['মোট বিক্রি'], errors='coerce').fillna(0).astype(int)
                    monthly_sales_sum = all_sales_df.groupby('প্রোডাক্ট')['মোট বিক্রি'].sum().reset_index()
                    
                    st.subheader("📈 এই মাসের সর্বমোট প্রোডাক্ট ভিত্তিক বিক্রি")
                    st.dataframe(monthly_sales_sum, use_container_width=True, hide_index=True)
                    
                    fig_month = px.bar(monthly_sales_sum, x='প্রোডাক্ট', y='মোট বিক্রি', text='মোট বিক্রি', title="মাসিক মোট বিক্রির সামারি")
                    st.plotly_chart(fig_month, use_container_width=True)

            # Tab 2
            with mtab2:
                if m_person_list:
                    all_people_df = pd.concat(m_person_list)
                    cols_to_sum = ['Opening', 'Issue', 'Total', 'Sales', 'Return', 'Closing']
                    for col in cols_to_sum:
                        all_people_df[col] = pd.to_numeric(all_people_df[col], errors='coerce').fillna(0).astype(int)
                        
                    unique_staff = all_people_df['Name'].unique()
                    selected_staff = st.selectbox("👤 মাসিক ডেটা দেখতে কর্মী সিলেক্ট করুন:", unique_staff)
                    
                    staff_monthly_df = all_people_df[all_people_df['Name'] == selected_staff]
                    staff_summary = staff_monthly_df.groupby('প্রোডাক্ট')[['Issue', 'Sales', 'Return']].sum().reset_index()
                    
                    st.subheader(f"📋 {selected_staff} - এর এই মাসের মোট পারফরম্যান্স খতিয়ান")
                    st.dataframe(staff_summary.rename(columns={'Issue':'মোট ইস্যু', 'Sales':'মোট বিক্রি', 'Return':'মোট ফেরত'}), use_container_width=True, hide_index=True)
                    
                    fig_staff = px.bar(staff_summary, x='প্রোডাক্ট', y='Sales', text='Sales', title=f"{selected_staff} এর মাসিক মোট বিক্রির গ্রাফ")
                    st.plotly_chart(fig_staff, use_container_width=True)

            # Tab 3
            with mtab3:
                if m_profit_list:
                    all_profit_df = pd.concat(m_profit_list)
                    for col in ['বিক্রি সংখ্যা', 'মোট কেনা', 'মোট বিক্রি', 'লাভ']:
                        all_profit_df[col] = pd.to_numeric(all_profit_df[col], errors='coerce').fillna(0).astype(float)
                        
                    monthly_profit_sum = all_profit_df.groupby('প্রোডাক্ট')[['বিক্রি সংখ্যা', 'মোট কেনা', 'মোট বিক্রি', 'লাভ']].sum().reset_index()
                    
                    st.subheader("💰 এই মাসের সর্বমোট লাভ-ক্ষতি ও রেভিনিউ হিসাব")
                    st.dataframe(monthly_profit_sum, use_container_width=True, hide_index=True)
                    
                    total_m_profit = monthly_profit_sum['লাভ'].sum()
                    st.metric(label=f"💵 {selected_month} মাসের মোট নেট প্রফিট", value=f"{total_m_profit:,.2f} টাকা")
                    
                    fig_m_profit = px.bar(monthly_profit_sum, x='প্রোডাক্ট', y='লাভ', text='লাভ', title="মাসিক নিট লাভ (প্রোডাক্ট অনুযায়ী)")
                    st.plotly_chart(fig_m_profit, use_container_width=True)

                    # মাসিক সামারি রিপোর্টের এক্সেল ডাউনলোড বাটন
                    monthly_excel_data = convert_df_to_excel(monthly_profit_sum)
                    st.download_button(
                        label="📥 পুরো মাসের রিপোর্ট ডাউনলোড করুন (Excel)",
                        data=monthly_excel_data,
                        file_name=f"Monthly_Report_{selected_month}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
        else:
            st.warning(f"⚠️ '{selected_month}' ফরম্যাটের কোনো ফাইল পাওয়া যায়নি।")


# স্ট্রিমলিটের ডিফল্ট ফুটার হাইড করার কোড
hide_default_style = """
            <style>
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_default_style, unsafe_allow_html=True)

# একদম নিচে আলাদা লাইনে ডেভেলপারের নাম দেখানোর কোড
st.markdown("---")
st.markdown(
    "<p style='text-align: right; color: #555555; font-size: 14px; font-weight: 500;'> @ Shourav Shahriar</p>", 
    unsafe_allow_html=True
)
