import streamlit as st
import sqlite3
import pandas as pd
import os
from datetime import datetime

# ---------------------------------------------------------
# 1. CẤU HÌNH CƠ SỞ DỮ LIỆU & HỆ THỐNG
# ---------------------------------------------------------
# CẤU HÌNH TRANG & NHÚNG ICON CHO ĐIỆN THOẠI
import base64

# Hàm đọc file icon chuyển thành dạng nhúng trực tiếp
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

# Đọc icon.png từ thư mục
try:
    icon_base64 = get_base64_of_bin_file('icon.png')
    icon_link = f'data:image/png;base64,{icon_base64}'
except:
    icon_link = ''

# CẤU HÌNH TRANG
st.set_page_config(
    page_title="Quản Lý Kho Cấp Đông Mr Hưng",
    page_icon="icon.png",
    layout="wide"
)

# ÉP ĐIỆN THOẠI NHẬN ICON MỚI TRỰC TIẾP
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


# Nhúng manifest để điện thoại nhận Icon ngoài màn hình chính
st.markdown(
    """
    <link rel="manifest" href="https://raw.githubusercontent.com/hungnguyen88budang-ui/Qu-n-L-V-n-H-nh-Kho-C-p-/main/manifest.json">
    <link rel="apple-touch-icon" href="https://raw.githubusercontent.com/hungnguyen88budang-ui/Qu-n-L-V-n-H-nh-Kho-C-p-/main/icon.png">
    """,
    unsafe_allow_html=True
)

UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

conn = sqlite3.connect("kho_cap_dong_v5.db", check_same_thread=False)
cursor = conn.cursor()

# Bảng Báo cáo
cursor.execute('''
CREATE TABLE IF NOT EXISTS bao_cao (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_bao_cao TEXT UNIQUE,
    thoi_gian TEXT,
    ten_kho TEXT,
    loai_bao_cao TEXT,
    nhiet_do REAL,
    do_am REAL,
    san_luong REAL,
    trang_thai_may TEXT,
    noi_dung TEXT,
    duong_dan_anh TEXT,
    nguoi_bao_cao TEXT
)
''')

# Bảng Cài đặt Kho & Nhân viên
cursor.execute('CREATE TABLE IF NOT EXISTS danh_muc_kho (ten_kho TEXT UNIQUE)')
cursor.execute('CREATE TABLE IF NOT EXISTS danh_muc_nv (ten_nv TEXT UNIQUE)')

# Bảng Quản lý Tài khoản & Phân quyền
cursor.execute('''
CREATE TABLE IF NOT EXISTS tai_khoan (
    username TEXT PRIMARY KEY,
    password TEXT,
    ho_ten TEXT,
    vai_tro TEXT,
    trang_thai TEXT
)
''')

# Dữ liệu danh mục ban đầu
cursor.execute('INSERT OR IGNORE INTO danh_muc_kho VALUES ("Kho Cấp Đông 1"), ("Kho Cấp Đông 2"), ("Phòng Máy Nén")')
cursor.execute('INSERT OR IGNORE INTO danh_muc_nv VALUES ("Nguyễn Văn A"), ("Trần Văn B"), ("Lê Văn C")')

# Khởi tạo Tài khoản mặc định
cursor.execute('INSERT OR IGNORE INTO tai_khoan VALUES ("admin", "admin123", "Quản Trị Viên Hùng", "admin", "hoat_dong")')
cursor.execute('INSERT OR IGNORE INTO tai_khoan VALUES ("nv01", "123", "Nguyễn Văn A", "nhanvien", "hoat_dong")')
cursor.execute('INSERT OR IGNORE INTO tai_khoan VALUES ("xem01", "123", "Cấp Quản Lý Xem", "viewer", "hoat_dong")')
conn.commit()

# ---------------------------------------------------------
# 2. XỬ LÝ ĐĂNG NHẬP & PHÂN QUYỀN
# ---------------------------------------------------------
if "user_info" not in st.session_state:
    st.session_state["user_info"] = None

st.sidebar.image("https://img.icons8.com/color/96/000000/cold-storage.png", width=70)
st.sidebar.title("🔐 ĐĂNG NHẬP HỆ THỐNG")

with st.sidebar:
    if st.session_state["user_info"] is None:
        u_input = st.text_input("Tài khoản")
        p_input = st.text_input("Mật khẩu", type="password")
        if st.button("Đăng nhập"):
            user_query = cursor.execute("SELECT username, password, ho_ten, vai_tro, trang_thai FROM tai_khoan WHERE username = ?", (u_input,)).fetchone()
            if user_query:
                if user_query[4] == "bi_khoa":
                    st.error("❌ Tài khoản này đã bị Admin thu hồi/khóa quyền!")
                elif user_query[1] == p_input:
                    st.session_state["user_info"] = {
                        "username": user_query[0],
                        "ho_ten": user_query[2],
                        "vai_tro": user_query[3]
                    }
                    st.success(f"Xin chào: {user_query[2]}")
                    st.rerun()
                else:
                    st.error("Sai mật khẩu!")
            else:
                st.error("Tài khoản không tồn tại!")
    else:
        u_info = st.session_state["user_info"]
        st.success(f"👤 **{u_info['ho_ten']}**")
        st.caption(f"Quyền: **{u_info['vai_tro'].upper()}**")
        if st.button("Đăng xuất"):
            st.session_state["user_info"] = None
            st.rerun()

# ---------------------------------------------------------
# 3. GIAO DIỆN CHÍNH
# ---------------------------------------------------------
st.title("❄️ QUẢN LÝ & BÁO CÁO VẬN HÀNH KHO CẤP ĐÔNG")

if st.session_state["user_info"] is None:
    st.info("👈 Vui lòng đăng nhập tài khoản ở menu bên trái để sử dụng hệ thống.")
else:
    current_user = st.session_state["user_info"]
    is_admin = (current_user["vai_tro"] == "admin")
    
    ds_kho = [row[0] for row in cursor.execute("SELECT ten_kho FROM danh_muc_kho").fetchall()]
    ds_nv = [row[0] for row in cursor.execute("SELECT ten_nv FROM danh_muc_nv").fetchall()]

    tabs_list = ["📊 Xem Tất Cả Báo Cáo", "📝 Lập Báo Cáo Mới"]
    if is_admin:
        tabs_list.append("⚙️ Admin - Cài Đặt Hệ Thống & Chỉnh Sửa Báo Cáo")
        tabs_list.append("👥 Admin - Quản Lý & Thu Hồi Tài Khoản")

    tabs = st.tabs(tabs_list)

    # TAB 1: XEM BÁO CÁO (KHÓA CHỈNH SỬA VỚI NHÂN VIÊN)
    with tabs[0]:
        st.subheader("👀 Nhật Ký Báo Cáo Kho (Số liệu đã gửi không thể chỉnh sửa)")
        df = pd.read_sql_query("SELECT * FROM bao_cao ORDER BY id DESC", conn)
        
        if not df.empty:
            excel_data = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Tải Báo Cáo File Excel", data=excel_data, file_name="Bao_Cao_Kho.csv", mime="text/csv")
            
            st.markdown("---")
            for idx, row in df.iterrows():
                with st.expander(f"📌 Mã BC: {row['ma_bao_cao']} | Kho: {row['ten_kho']} | Ngày: {row['thoi_gian']} (Lập bởi: {row['nguoi_bao_cao']})"):
                    col_l, col_r = st.columns([2, 1])
                    with col_l:
                        st.write(f"**Loại báo cáo:** {row['loai_bao_cao']} | **Trạng thái:** {row['trang_thai_may']}")
                        st.write(f"**Thông số:** Nhiệt độ: `{row['nhiet_do']}°C` | Độ ẩm: `{row['do_am']}%` | Sản lượng: `{row['san_luong']} Tấn`")
                        st.write(f"**Nội dung ghi chú:**\n{row['noi_dung']}")
                        st.caption("🔒 *Dữ liệu này đã được lưu cố định vào hệ thống.*")
                    with col_r:
                        if row['duong_dan_anh'] and os.path.exists(row['duong_dan_anh']):
                            st.image(row['duong_dan_anh'], caption="Ảnh thực tế", use_column_width=True)
                        else:
                            st.write("*(Không đính kèm ảnh)*")
        else:
            st.info("Chưa có báo cáo nào.")

    # TAB 2: LẬP BÁO CÁO MỚI (CÓ Ô CHỤP / TẢI ẢNH)
    with tabs[1]:
        if current_user["vai_tro"] in ["nhanvien", "admin"]:
            st.subheader("📝 Lập Báo Cáo Vận Hành Mới")
            with st.form("form_nhap", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                with c1:
                    ma_bc = st.text_input("Mã Báo Cáo Auto", value=f"BC-{datetime.now().strftime('%Y%m%d-%H%M%S')}")
                    ten_kho = st.selectbox("Chọn Kho Cấp Đông", ds_kho if ds_kho else ["Chưa có kho"])
                    loai_bc = st.selectbox("Loại Báo Cáo", ["Định kỳ ca trực", "Bảo trì / Báo sự cố", "Nhập / Xuất hàng"])
                with c2:
                    nhiet_do = st.number_input("Nhiệt độ (°C)", value=-18.0, step=0.1)
                    do_am = st.number_input("Độ ẩm (%)", value=85.0, step=0.5)
                    san_luong = st.number_input("Sản lượng lưu kho (Tấn)", value=0.0, step=0.1)
                with c3:
                    trang_thai = st.selectbox("Trạng thái máy", ["Bình thường", "Cảnh báo nhẹ", "Sự cố - Cần sửa chữa"])
                    nguoi_lap = st.text_input("Tên Người Lập", value=current_user["ho_ten"], disabled=True)

                noi_dung = st.text_area("Nội dung chi tiết / Ghi chú sự cố")
                uploaded_file = st.file_uploader("📸 Chụp / Tải ảnh đính kèm", type=["jpg", "png", "jpeg"])
                
                if st.form_submit_button("🚀 GỬI BÁO CÁO VỀ HỆ THỐNG"):
                    img_path = ""
                    if uploaded_file:
                        img_path = os.path.join(UPLOAD_DIR, f"{ma_bc}_{uploaded_file.name}")
                        with open(img_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                    
                    cursor.execute('''
                    INSERT INTO bao_cao (ma_bao_cao, thoi_gian, ten_kho, loai_bao_cao, nhiet_do, do_am, san_luong, trang_thai_may, noi_dung, duong_dan_anh, nguoi_bao_cao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (ma_bc, datetime.now().strftime("%d/%m/%Y %H:%M:%S"), ten_kho, loai_bc, nhiet_do, do_am, san_luong, trang_thai, noi_dung, img_path, nguoi_lap))
                    conn.commit()
                    st.success("✅ Đã gửi báo cáo thành công! Dữ liệu đã được khóa tự động.")
        else:
            st.warning("Tài khoản của bạn chỉ có quyền XEM, không có quyền lập báo cáo.")

    # TAB 3: ADMIN CÀI ĐẶT
    if is_admin:
        with tabs[2]:
            st.subheader("⚙️ Quản Lý Danh Mục Hệ Thống")
            col_k, col_n = st.columns(2)
            with col_k:
                st.markdown("**🏢 Thêm Tên Kho Mới**")
                new_kho = st.text_input("Tên kho mới")
                if st.button("➕ Thêm Kho"):
                    if new_kho:
                        cursor.execute("INSERT OR IGNORE INTO danh_muc_kho VALUES (?)", (new_kho,))
                        conn.commit()
                        st.success(f"Đã thêm: {new_kho}")
                        st.rerun()
                st.write("Danh sách kho hiện có:", ds_kho)

            with col_n:
                st.markdown("**👷 Thêm Tên Nhân Viên Mới**")
                new_nv = st.text_input("Tên nhân viên mới")
                if st.button("➕ Thêm Nhân Viên"):
                    if new_nv:
                        cursor.execute("INSERT OR IGNORE INTO danh_muc_nv VALUES (?)", (new_nv,))
                        conn.commit()
                        st.success(f"Đã thêm: {new_nv}")
                        st.rerun()
                st.write("Danh sách nhân viên hiện có:", ds_nv)

            st.markdown("---")
            st.subheader("🛠️ Sửa Hoặc Xóa Báo Cáo Nhập Sai (Đặc Quyền Admin)")
            if not df.empty:
                id_sel = st.selectbox("Chọn ID Báo Cáo Cần Xóa / Sửa", df['id'].tolist())
                if st.button("❌ XÓA BÁO CÁO NÀY"):
                    cursor.execute("DELETE FROM bao_cao WHERE id = ?", (id_sel,))
                    conn.commit()
                    st.success(f"Đã xóa báo cáo ID {id_sel}")
                    st.rerun()

    # TAB 4: ADMIN QUẢN LÝ TÀI KHOẢN & THU HỒI
    if is_admin:
        with tabs[3]:
            st.subheader("👥 Cấp Tài Khoản Mới")
            with st.form("form_tao_tk"):
                c_u, c_p, c_n, c_r = st.columns(4)
                new_user = c_u.text_input("Tên Đăng Nhập")
                new_pass = c_p.text_input("Mật Khẩu")
                new_fullname = c_n.text_input("Họ Và Tên")
                new_role = c_r.selectbox("Phân Quyền", ["nhanvien", "viewer", "admin"])
                
                if st.form_submit_button("➕ TẠO TÀI KHOẢN MỚI"):
                    if new_user and new_pass and new_fullname:
                        try:
                            cursor.execute("INSERT INTO tai_khoan VALUES (?, ?, ?, ?, 'hoat_dong')", (new_user, new_pass, new_fullname, new_role))
                            conn.commit()
                            st.success(f"Đã tạo tài khoản cho {new_fullname} ({new_role})")
                            st.rerun()
                        except:
                            st.error("Tên đăng nhập này đã tồn tại!")
                    else:
                        st.warning("Vui lòng điền đủ thông tin.")

            st.markdown("---")
            st.subheader("🔒 Thu Hồi Quyền / Khóa Tài Khoản Nhân Viên")
            df_users = pd.read_sql_query("SELECT username, ho_ten, vai_tro, trang_thai FROM tai_khoan", conn)
            st.dataframe(df_users, use_container_width=True)

            col_usr, col_act = st.columns(2)
            usr_target = col_usr.selectbox("Chọn tài khoản cần thao tác", df_users['username'].tolist())
            
            if usr_target != "admin":
                if col_act.button("🚫 KHÓA / THU HỒI QUYỀN TÀI KHOẢN NÀY"):
                    cursor.execute("UPDATE tai_khoan SET trang_thai = 'bi_khoa' WHERE username = ?", (usr_target,))
                    conn.commit()
                    st.success(f"Đã khóa tài khoản {usr_target}!")
                    st.rerun()
                    
                if col_act.button("✅ MỞ KHÓA TÀI KHOẢN"):
                    cursor.execute("UPDATE tai_khoan SET trang_thai = 'hoat_dong' WHERE username = ?", (usr_target,))
                    conn.commit()
                    st.success(f"Đã kích hoạt lại {usr_target}.")
                    st.rerun()
            else:
                col_act.warning("Tài khoản Admin gốc không thể bị khóa.")