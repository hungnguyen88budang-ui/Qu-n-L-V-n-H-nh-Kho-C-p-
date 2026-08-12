import streamlit as st
import sqlite3
import pandas as pd
import os
from datetime import datetime
import base64

# --- 1. CẤU HÌNH TRANG & NHÚNG ICON CHO ĐIỆN THOẠI ---
def get_base64_of_bin_file(bin_file):
    try:
        with open(bin_file, 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except:
        return ''

icon_base64 = get_base64_of_bin_file('icon.png')
icon_link = f'data:image/png;base64,{icon_base64}' if icon_base64 else ''

st.set_page_config(
    page_title="Quản Lý Kho Cấp Đông",
    page_icon="icon.png" if os.path.exists("icon.png") else "❄️",
    layout="wide"
)

if icon_link:
    st.markdown(
        f"""
        <head>
            <link rel="apple-touch-icon" sizes="192x192" href="{icon_link}">
            <link rel="apple-touch-icon-precomposed" href="{icon_link}">
            <link rel="icon" type="image/png" sizes="192x192" href="{icon_link}">
            <link rel="shortcut icon" href="{icon_link}">
        </head>
        """,
        unsafe_allow_html=True
    )

# --- 2. KHỞI TẠO CƠ SỞ DỮ LIỆU ---
conn = sqlite3.connect('kho_cap_dong.db', check_same_thread=False)
c = conn.cursor()

# Bảng người dùng & phân quyền
c.execute('''
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT,
        fullname TEXT,
        role TEXT,
        can_report INTEGER DEFAULT 1,
        can_view_history INTEGER DEFAULT 1,
        can_edit INTEGER DEFAULT 0,
        can_delete INTEGER DEFAULT 0
    )
''')

# Bảng kho hàng mặc định
c.execute('''
    CREATE TABLE IF NOT EXISTS inventory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        type TEXT,
        item_name TEXT,
        quantity REAL,
        unit TEXT,
        created_by TEXT,
        note TEXT
    )
''')

# Tạo tài khoản Admin mặc định
c.execute("SELECT * FROM users WHERE username = 'admin'")
if not c.fetchone():
    c.execute('''
        INSERT INTO users (username, password, fullname, role, can_report, can_view_history, can_edit, can_delete)
        VALUES ('admin', '123456', 'Mr Hưng (Admin)', 'admin', 1, 1, 1, 1)
    ''')
conn.commit()

# --- 3. ĐĂNG NHẬP HỆ THỐNG ---
if 'user' not in st.session_state:
    st.session_state['user'] = None

if st.session_state['user'] is None:
    st.title("❄️ ĐĂNG NHẬP HỆ THỐNG KHO CẤP ĐÔNG")
    with st.form("login_form"):
        username = st.text_input("Tên đăng nhập").strip()
        password = st.text_input("Mật khẩu", type="password").strip()
        submit = st.form_submit_button("Đăng Nhập")
        
        if submit:
            c.execute("SELECT username, fullname, role, can_report, can_view_history, can_edit, can_delete FROM users WHERE username=? AND password=?", (username, password))
            row = c.fetchone()
            if row:
                st.session_state['user'] = {
                    'username': row[0],
                    'fullname': row[1],
                    'role': row[2],
                    'can_report': bool(row[3]),
                    'can_view_history': bool(row[4]),
                    'can_edit': bool(row[5]),
                    'can_delete': bool(row[6])
                }
                st.success(f"Xin chào {row[1]}!")
                st.rerun()
            else:
                st.error("Tên đăng nhập hoặc mật khẩu không chính xác!")
    st.info("💡 Mặc định Admin: Tên đăng nhập: **admin** | Mật khẩu: **123456**")
    st.stop()

# --- 4. BỐ CỤC CHÍNH (GIỮ NGUYÊN BỐ CỤC BAN ĐẦU) ---
user = st.session_state['user']

# Thanh bên Sidebar
st.sidebar.title("❄️ KHO CẤP ĐÔNG")
st.sidebar.write(f"👤 **{user['fullname']}** ({'Admin' if user['role']=='admin' else 'Nhân viên'})")

menu = []
if user['can_report']:
    menu.append("Gửi Báo Cáo / Nhập Xuất")
if user['can_view_history']:
    menu.append("Xem Lịch Sử Báo Cáo")
if user['role'] == 'admin':
    menu.append("Quản Lý Phân Quyền Nhân Viên")

choice = st.sidebar.selectbox("Điều Hướng Chức Năng", menu)

if st.sidebar.button("Đăng Xuất"):
    st.session_state['user'] = None
    st.rerun()

# --- BỐ CỤC BAN ĐẦU 1: GỬI BÁO CÁO ---
if choice == "Gửi Báo Cáo / Nhập Xuất":
    st.title("📝 GỬI BÁO CÁO KHO CẤP ĐÔNG")
    
    with st.form("report_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            report_type = st.selectbox("Loại thao tác", ["Nhập kho", "Xuất kho"])
            item_name = st.text_input("Tên mặt hàng / Lô hàng").strip()
        with col2:
            quantity = st.number_input("Số lượng", min_value=0.1, step=1.0)
            unit = st.selectbox("Đơn vị tính", ["Tấn", "Kg", "Thùng", "Khay", "Bao"])
        
        note = st.text_area("Ghi chú")
        submitted = st.form_submit_button("Gửi Báo Cáo")
        
        if submitted:
            if not item_name:
                st.error("Vui lòng nhập tên mặt hàng!")
            else:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c.execute('''
                    INSERT INTO inventory (date, type, item_name, quantity, unit, created_by, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (now, report_type, item_name, quantity, unit, user['fullname'], note))
                conn.commit()
                st.success("✅ Đã ghi nhận báo cáo thành công!")

# --- BỐ CỤC BAN ĐẦU 2: XEM LỊCH SỬ ---
elif choice == "Xem Lịch Sử Báo Cáo":
    st.title("📜 LỊCH SỬ BÁO CÁO KHO")
    
    df = pd.read_sql_query("SELECT id, date AS 'Thời Gian', type AS 'Loại', item_name AS 'Tên Mặt Hàng', quantity AS 'Số Lượng', unit AS 'Đơn Vị', created_by AS 'Người Báo Cáo', note AS 'Ghi Chú' FROM inventory ORDER BY id DESC", conn)
    
    if df.empty:
        st.info("Chưa có dữ liệu lịch sử.")
    else:
        st.dataframe(df.drop(columns=['id']), use_container_width=True)
        
        if user['can_delete']:
            st.write("---")
            st.subheader("🗑️ Xóa bản ghi (Dành riêng cho Admin)")
            del_id = st.selectbox("Chọn dòng ID cần xóa", df['id'].tolist())
            if st.button("Xóa Dòng Đã Chọn"):
                c.execute("DELETE FROM inventory WHERE id=?", (del_id,))
                conn.commit()
                st.success("Đã xóa bản ghi thành công!")
                st.rerun()

# --- TÍNH NĂNG MỚI BỔ SUNG: QUẢN LÝ PHÂN QUYỀN (CHỈ ADMIN) ---
elif choice == "Quản Lý Phân Quyền Nhân Viên":
    st.title("👥 QUẢN LÝ TÀI KHOẢN & PHÂN QUYỀN")
    
    tab1, tab2 = st.tabs(["➕ Thêm Nhân Viên", "⚙️ Danh Sách & Cắt Quyền"])
    
    with tab1:
        st.subheader("Tạo tài khoản cho nhân viên")
        with st.form("add_user_form", clear_on_submit=True):
            new_username = st.text_input("Tên đăng nhập (viết liền không dấu)").strip().lower()
            new_password = st.text_input("Mật khẩu", type="password").strip()
            new_fullname = st.text_input("Tên nhân viên").strip()
            
            st.write("📌 **Tích chọn cấp quyền:**")
            p_report = st.checkbox("Quyền gửi báo cáo", value=True)
            p_history = st.checkbox("Quyền xem lịch sử", value=True)
            p_delete = st.checkbox("Quyền xóa dữ liệu", value=False)
            
            submit_user = st.form_submit_button("Tạo Tài Khoản")
            if submit_user:
                if not new_username or not new_password or not new_fullname:
                    st.error("Vui lòng điền đủ thông tin!")
                else:
                    try:
                        c.execute('''
                            INSERT INTO users (username, password, fullname, role, can_report, can_view_history, can_delete)
                            VALUES (?, ?, ?, 'staff', ?, ?, ?)
                        ''', (new_username, new_password, new_fullname, int(p_report), int(p_history), int(p_delete)))
                        conn.commit()
                        st.success(f"✅ Đã tạo tài khoản cho {new_fullname}!")
                    except:
                        st.error("Tên đăng nhập đã tồn tại!")

    with tab2:
        st.subheader("Danh sách nhân viên")
        users_df = pd.read_sql_query("SELECT username AS 'Tên ĐN', fullname AS 'Họ Tên', can_report AS 'Quyền Gửi BC', can_view_history AS 'Quyền Xem LS', can_delete AS 'Quyền Xóa' FROM users WHERE role='staff'", conn)
        if users_df.empty:
            st.info("Chưa có tài khoản nhân viên nào.")
        else:
            st.dataframe(users_df, use_container_width=True)
            st.write("---")
            user_to_del = st.selectbox("Chọn tài khoản cần xóa / thu hồi quyền", users_df['Tên ĐN'].tolist())
            if st.button("Xóa / Cắt Quyền Nhân Viên Giờ"):
                c.execute("DELETE FROM users WHERE username=?", (user_to_del,))
                conn.commit()
                st.success("Đã xóa tài khoản nhân viên khỏi hệ thống!")
                st.rerun()