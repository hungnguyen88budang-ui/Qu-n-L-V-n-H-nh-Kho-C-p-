import streamlit as st
import sqlite3
import pandas as pd
import os
from datetime import datetime
import base64

# --- CẤU HÌNH TRANG & ICON ---
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
    page_title="Quản Lý Kho Cấp Đông Mr Hưng",
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

# --- KHỞI TẠO CƠ SỞ DỮ LIỆU ---
conn = sqlite3.connect('kho_cap_dong.db', check_same_thread=False)
c = conn.cursor()

# Bảng người dùng & phân quyền
c.execute('''
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT,
        fullname TEXT,
        role TEXT, -- 'admin' hoặc 'staff'
        can_report INTEGER DEFAULT 1, -- Quyền gửi báo cáo
        can_view_history INTEGER DEFAULT 1, -- Quyền xem lịch sử
        can_edit INTEGER DEFAULT 0, -- Quyền sửa
        can_delete INTEGER DEFAULT 0 -- Quyền xóa
    )
''')

# Bảng kho hàng
c.execute('''
    CREATE TABLE IF NOT EXISTS inventory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        type TEXT, -- 'Nhập' hoặc 'Xuất'
        item_name TEXT,
        quantity REAL,
        unit TEXT,
        created_by TEXT,
        note TEXT
    )
''')

# Tạo tài khoản Admin mặc định nếu chưa có
c.execute("SELECT * FROM users WHERE username = 'admin'")
if not c.fetchone():
    c.execute('''
        INSERT INTO users (username, password, fullname, role, can_report, can_view_history, can_edit, can_delete)
        VALUES ('admin', '123456', 'Mr Hưng (Admin)', 'admin', 1, 1, 1, 1)
    ''')
conn.commit()

# --- XỬ LÝ ĐĂNG NHẬP ---
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
    st.info("💡 Tài khoản mặc định của Anh Hưng: Tên đăng nhập: **admin** | Mật khẩu: **123456**")
    st.stop()

# --- GIAO DIỆN CHÍNH SAU KHU ĐĂNG NHẬP ---
user = st.session_state['user']

# Thanh bên (Sidebar)
st.sidebar.title(f"👤 {user['fullname']}")
if user['role'] == 'admin':
    st.sidebar.caption("👑 Tài khoản Admin (Toàn quyền)")
else:
    st.sidebar.caption("📋 Tài khoản Nhân viên")

# Menu điều hướng tùy theo quyền
menu_options = []
if user['can_report']:
    menu_options.append("📝 Gửi Báo Cáo Nhập/Xuất")
if user['can_view_history']:
    menu_options.append("📜 Lịch Sử Báo Cáo")
if user['role'] == 'admin':
    menu_options.append("👥 Quản Lý Phân Quyền Nhân Viên")

choice = st.sidebar.radio("CHỨC NĂNG", menu_options)

if st.sidebar.button("Đăng Xuất"):
    st.session_state['user'] = None
    st.rerun()

# --- CHỨC NĂNG 1: GỬI BÁO CÁO ---
if choice == "📝 Gửi Báo Cáo Nhập/Xuất":
    st.header("📝 BÁO CÁO NHẬP / XUẤT KHO")
    
    with st.form("report_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            report_type = st.selectbox("Loại báo cáo", ["Nhập kho", "Xuất kho"])
            item_name = st.text_input("Tên mặt hàng / Lô hàng").strip()
        with col2:
            quantity = st.number_input("Số lượng", min_value=0.1, step=1.0)
            unit = st.selectbox("Đơn vị tính", ["Tấn", "Kg", "Thùng", "Khay", "Bao"])
        
        note = st.text_area("Ghi chú thêm (nếu có)")
        submitted = st.form_submit_button("📤 Gửi Báo Cáo")
        
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
                st.success(f"✅ Đã gửi báo cáo {report_type} thành công!")

# --- CHỨC NĂNG 2: XEM LỊCH SỬ BÁO CÁO ---
elif choice == "📜 Lịch Sử Báo Cáo":
    st.header("📜 LỊCH SỬ BÁO CÁO KHO")
    
    df = pd.read_sql_query("SELECT id, date AS 'Thời Gian', type AS 'Loại', item_name AS 'Tên Tệp/Mặt Hàng', quantity AS 'Số Lượng', unit AS 'Đơn Vị', created_by AS 'Người Báo Cáo', note AS 'Ghi Chú' FROM inventory ORDER BY id DESC", conn)
    
    if df.empty:
        st.info("Chưa có lịch sử báo cáo nào.")
    else:
        st.dataframe(df.drop(columns=['id']), use_container_width=True)
        
        # CHỈ ADMIN HOẶC NGƯỜI ĐƯỢC CẤP QUYỀN MỚI THẤY NÚT XÓA
        if user['can_delete']:
            st.subheader("⚠️ Quản lý xóa báo cáo (Quyền Đặc Biệt)")
            delete_id = st.selectbox("Chọn ID dòng cần xóa", df['id'].tolist())
            if st.button("🗑️ Xóa dòng này"):
                c.execute("DELETE FROM inventory WHERE id=?", (delete_id,))
                conn.commit()
                st.success("Đã xóa báo cáo!")
                st.rerun()

# --- CHỨC NĂNG 3: QUẢN LÝ PHÂN QUYỀN (CHỈ ADMIN) ---
elif choice == "👥 Quản Lý Phân Quyền Nhân Viên":
    st.header("👥 QUẢN LÝ TÀI KHOẢN & CẤP QUYỀN NHÂN VIÊN")
    
    tab1, tab2 = st.tabs(["➕ Thêm Nhân Viên Mới", "⚙️ Chỉnh Sửa Quyền Hạn"])
    
    with tab1:
        st.subheader("Tạo tài khoản mới cho nhân viên")
        with st.form("add_user_form", clear_on_submit=True):
            new_username = st.text_input("Tên đăng nhập (viết liền không dấu, ví dụ: nam, tuan)").strip().lower()
            new_password = st.text_input("Mật khẩu", type="password").strip()
            new_fullname = st.text_input("Họ và tên nhân viên (ví dụ: Nguyễn Văn Nam)").strip()
            
            st.write("📌 **Tích chọn cấp quyền cho nhân viên này:**")
            p_report = st.checkbox("Quyền gửi báo cáo Nhập/Xuất", value=True)
            p_history = st.checkbox("Quyền xem lịch sử báo cáo", value=True)
            p_edit = st.checkbox("Quyền chỉnh sửa dữ liệu", value=False)
            p_delete = st.checkbox("Quyền xóa dữ liệu", value=False)
            
            submit_user = st.form_submit_button("Thêm Nhân Viên")
            if submit_user:
                if not new_username or not new_password or not new_fullname:
                    st.error("Vui lòng nhập đầy đủ thông tin!")
                else:
                    try:
                        c.execute('''
                            INSERT INTO users (username, password, fullname, role, can_report, can_view_history, can_edit, can_delete)
                            VALUES (?, ?, ?, 'staff', ?, ?, ?, ?)
                        ''', (new_username, new_password, new_fullname, int(p_report), int(p_history), int(p_edit), int(p_delete)))
                        conn.commit()
                        st.success(f"✅ Đã tạo tài khoản cho nhân viên {new_fullname} thành công!")
                    except:
                        st.error("Tên đăng nhập này đã tồn tại! Vui lòng chọn tên khác.")

    with tab2:
        st.subheader("Danh sách nhân viên & Thay đổi quyền")
        users_df = pd.read_sql_query("SELECT username AS 'Tên ĐN', fullname AS 'Họ Tên', can_report AS 'Gửi BC', can_view_history AS 'Xem LS', can_edit AS 'Sửa', can_delete AS 'Xóa' FROM users WHERE role='staff'", conn)
        if users_df.empty:
            st.info("Chưa có tài khoản nhân viên nào.")
        else:
            st.dataframe(users_df, use_container_width=True)
            
            st.markdown("---")
            st.write("❌ **Xóa tài khoản nhân viên:**")
            user_to_del = st.selectbox("Chọn nhân viên cần xóa", users_df['Tên ĐN'].tolist())
            if st.button("Xóa tài khoản này"):
                c.execute("DELETE FROM users WHERE username=?", (user_to_del,))
                conn.commit()
                st.success("Đã xóa tài khoản nhân viên!")
                st.rerun()