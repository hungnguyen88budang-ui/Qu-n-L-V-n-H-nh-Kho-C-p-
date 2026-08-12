import streamlit as st
import sqlite3
import pandas as pd
import os
from datetime import datetime
import base64
import zipfile
import io

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG & ICON
# ---------------------------------------------------------
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
    page_title="Hệ Thống Báo Cáo Vận Hành Kho & Thiết Bị", 
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

# ---------------------------------------------------------
# 2. KHỞI TẠO CƠ SỞ DỮ LIỆU SẠCH (DATABASE V100 MỚI TÍNH)
# ---------------------------------------------------------
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# Sử dụng tên database hoàn toàn mới để ép Streamlit tạo mới sạch sẽ 100%
conn = sqlite3.connect("kho_system_v100.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute('''
CREATE TABLE IF NOT EXISTS bao_cao_tong_hop (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_ca_truc TEXT UNIQUE,
    thoi_gian TEXT,
    ca_truc TEXT,
    nguoi_bao_cao TEXT,
    ghi_chu_chung TEXT
)
''')

cursor.execute('''
CREATE TABLE IF NOT EXISTS chi_tiet_kho (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ma_ca_truc TEXT,
    ten_kho TEXT,
    nhiet_do REAL,
    do_am REAL,
    san_luong REAL,
    trang_thai_may TEXT,
    duong_dan_anh TEXT
)
''')

# Khởi tạo danh mục kho
cursor.execute('CREATE TABLE IF NOT EXISTS danh_muc_kho (id INTEGER PRIMARY KEY AUTOINCREMENT, ten_kho TEXT UNIQUE)')
count_kho = cursor.execute('SELECT COUNT(*) FROM danh_muc_kho').fetchone()[0]
if count_kho == 0:
    danh_sach_kho_ban_dau = [f"Cụm Kho Số {i}" for i in range(1, 8)] + ["Phòng Máy Nén", "Trạm Biến Áp"]
    for kho in danh_sach_kho_ban_dau:
        cursor.execute('INSERT OR IGNORE INTO danh_muc_kho (ten_kho) VALUES (?)', (kho,))

# Khởi tạo danh mục trạng thái
cursor.execute('CREATE TABLE IF NOT EXISTS danh_muc_trang_thai (id INTEGER PRIMARY KEY AUTOINCREMENT, ten_trang_thai TEXT UNIQUE)')
count_tt = cursor.execute('SELECT COUNT(*) FROM danh_muc_trang_thai').fetchone()[0]
if count_tt == 0:
    ds_tt_ban_dau = ["Bình thường", "Cảnh báo nhẹ", "Sự cố - Cần sửa chữa", "Bảo trì định kỳ"]
    for tt in ds_tt_ban_dau:
        cursor.execute('INSERT OR IGNORE INTO danh_muc_trang_thai (ten_trang_thai) VALUES (?)', (tt,))

# Tạo bảng Tài khoản chuẩn 5 cột ngay từ đầu
cursor.execute('''
CREATE TABLE IF NOT EXISTS tai_khoan (
    username TEXT PRIMARY KEY,
    password TEXT,
    ho_ten TEXT,
    vai_tro TEXT,
    trang_thai TEXT
)
''')

cursor.execute('INSERT OR IGNORE INTO tai_khoan (username, password, ho_ten, vai_tro, trang_thai) VALUES ("admin", "admin123", "Quản Trị Viên Hùng", "admin", "hoat_dong")')
cursor.execute('INSERT OR IGNORE INTO tai_khoan (username, password, ho_ten, vai_tro, trang_thai) VALUES ("nv01", "123", "Nguyễn Văn A", "nhanvien", "hoat_dong")')
cursor.execute('INSERT OR IGNORE INTO tai_khoan (username, password, ho_ten, vai_tro, trang_thai) VALUES ("xem01", "123", "Ban Giám Đốc", "viewer", "hoat_dong")')
conn.commit()

# ---------------------------------------------------------
# 3. ĐĂNG NHẬP & XỬ LÝ PHÂN QUYỀN
# ---------------------------------------------------------
if "user_info" not in st.session_state:
    st.session_state["user_info"] = None

st.sidebar.image("https://img.icons8.com/color/96/000000/cold-storage.png", width=70)
st.sidebar.title("🔐 ĐĂNG NHẬP")

with st.sidebar:
    if st.session_state["user_info"] is None:
        u_input = st.text_input("Tài khoản").strip()
        p_input = st.text_input("Mật khẩu", type="password").strip()
        if st.button("Đăng nhập"):
            user_query = cursor.execute("SELECT username, password, ho_ten, vai_tro, trang_thai FROM tai_khoan WHERE username = ?", (u_input,)).fetchone()
            if user_query:
                if user_query[4] == "bi_khoa":
                    st.error("❌ Tài khoản này đã bị Admin khóa!")
                elif user_query[1] == p_input:
                    st.session_state["user_info"] = {"username": user_query[0], "ho_ten": user_query[2], "vai_tro": user_query[3]}
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
# 4. GIAO DIỆN CHÍNH
# ---------------------------------------------------------
st.title("❄️ QUẢN LÝ & BÁO CÁO VẬN HÀNH KHO / THIẾT BỊ")

if st.session_state["user_info"] is None:
    st.info("👈 Vui lòng đăng nhập tài khoản ở menu bên trái.")
else:
    current_user = st.session_state["user_info"]
    is_admin = (current_user["vai_tro"] == "admin")
    can_report = current_user["vai_tro"] in ["admin", "nhanvien"]

    ds_kho = [row[0] for row in cursor.execute("SELECT ten_kho FROM danh_muc_kho ORDER BY id ASC").fetchall()]
    ds_tt = [row[0] for row in cursor.execute("SELECT ten_trang_thai FROM danh_muc_trang_thai ORDER BY id ASC").fetchall()]

    tabs_list = ["📊 Xem & Tải Báo Cáo", "📝 Lập Báo Cáo Ca Trực"]
    if is_admin:
        tabs_list.append("⚙️ Admin - Quản Lý Kho & Trạng Thái")
        tabs_list.append("👥 Admin - Quản Lý Tài Khoản")

    tabs = st.tabs(tabs_list)

    # TAB 1: XEM BÁO CÁO
    with tabs[0]:
        st.subheader("📊 Nhật Ký Báo Cáo Ca Trực")
        df_ca = pd.read_sql_query("SELECT * FROM bao_cao_tong_hop ORDER BY id DESC", conn)

        if not df_ca.empty:
            for idx, row_ca in df_ca.iterrows():
                ma_ca = row_ca['ma_ca_truc']
                with st.expander(f"📌 Mã Ca: {ma_ca} | Ngày: {row_ca['thoi_gian']} | Ca: {row_ca['ca_truc']} | Người báo cáo: {row_ca['nguoi_bao_cao']}"):
                    st.write(f"**Ghi chú chung:** {row_ca['ghi_chu_chung']}")
                    
                    df_chitiet = pd.read_sql_query("SELECT ten_kho AS 'Tên Kho / Thiết Bị', nhiet_do AS 'Nhiệt Độ (°C)', do_am AS 'Độ Ẩm (%)', san_luong AS 'Sản Lượng (Tấn)', trang_thai_may AS 'Trạng Thái' FROM chi_tiet_kho WHERE ma_ca_truc = ?", conn, params=(ma_ca,))
                    col_dl1, col_dl2 = st.columns(2)
                    
                    csv_data = df_chitiet.to_csv(index=False).encode('utf-8-sig')
                    col_dl1.download_button("📥 Tải Báo Cáo Excel Ca Này", data=csv_data, file_name=f"BaoCao_{ma_ca}.csv", mime="text/csv", key=f"csv_{ma_ca}")

                    list_anh = cursor.execute("SELECT ten_kho, duong_dan_anh FROM chi_tiet_kho WHERE ma_ca_truc = ? AND duong_dan_anh != ''", (ma_ca,)).fetchall()
                    if list_anh:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            for item in list_anh:
                                if item[1] and os.path.exists(item[1]):
                                    ext = os.path.splitext(item[1])[1]
                                    zip_file.write(item[1], arcname=f"{item[0]}{ext}")
                        col_dl2.download_button("📸 Tải Trọn Bộ Ảnh (.ZIP)", data=zip_buffer.getvalue(), file_name=f"Anh_{ma_ca}.zip", mime="application/zip", key=f"zip_{ma_ca}")

                    st.markdown("---")
                    st.write("📷 **BÁO CÁO CHI TIẾT TỪNG KHO / THIẾT BỊ:**")
                    
                    df_full = cursor.execute("SELECT ten_kho, nhiet_do, do_am, san_luong, trang_thai_may, duong_dan_anh FROM chi_tiet_kho WHERE ma_ca_truc = ?", (ma_ca,)).fetchall()
                    cols = st.columns(3)
                    for idx_item, item in enumerate(df_full):
                        with cols[idx_item % 3]:
                            with st.container(border=True):
                                st.markdown(f"#### 🏭 {item[0]}")
                                st.write(f"• Nhiệt độ: **{item[1]} °C**")
                                st.write(f"• Độ ẩm: **{item[2]} %**")
                                st.write(f"• Sản lượng: **{item[3]} Tấn**")
                                st.write(f"• Trạng thái: **{item[4]}**")
                                st.markdown("**Hình ảnh thực tế:**")
                                if item[5] and os.path.exists(item[5]):
                                    st.image(item[5], use_container_width=True)
                                else:
                                    st.info("Không có ảnh")
        else:
            st.info("Chưa có báo cáo ca trực nào.")

    # TAB 2: LẬP BÁO CÁO
    with tabs[1]:
        if can_report:
            st.subheader("📝 Lập Báo Cáo Ca Trực Mới")
            if not ds_kho:
                st.warning("⚠️ Hiện tại chưa có kho/thiết bị nào trong hệ thống. Vui lòng nhờ Admin thêm kho vào danh mục.")
            else:
                with st.form("form_nhap_ca", clear_on_submit=True):
                    c1, c2, c3 = st.columns(3)
                    ma_ca = c1.text_input("Mã Ca Trực", value=f"CA-{datetime.now().strftime('%Y%m%d-%H%M%S')}")
                    ca_truc = c2.selectbox("Ca Trực", ["Ca Sáng (06h - 14h)", "Ca Chiều (14h - 22h)", "Ca Đêm (22h - 06h)"])
                    nguoi_lap = c3.text_input("Người Báo Cáo", value=current_user["ho_ten"], disabled=True)
                    ghi_chu_chung = st.text_area("Ghi chú chung ca trực")

                    st.markdown("---")
                    st.write(f"📋 **NHẬP BÁO CÁO CHO {len(ds_kho)} KHO / THIẾT BỊ HIỆN CÓ:**")

                    kho_inputs = {}
                    for kho in ds_kho:
                        with st.container(border=True):
                            st.markdown(f"### ❄️ {kho}")
                            col_l, col_r = st.columns([3, 2])
                            
                            with col_l:
                                col_a, col_b, col_c = st.columns(3)
                                n_do = col_a.number_input(f"Nhiệt độ (°C)", value=-18.0, step=0.1, key=f"nd_{kho}")
                                d_am = col_b.number_input(f"Độ ẩm (%)", value=85.0, step=0.5, key=f"da_{kho}")
                                s_luong = col_c.number_input(f"Sản lượng (Tấn)", value=0.0, step=0.1, key=f"sl_{kho}")
                                t_thai = st.selectbox(f"Trạng thái vận hành", ds_tt if ds_tt else ["Bình thường"], key=f"tt_{kho}")

                            with col_r:
                                file_img = st.file_uploader(f"📸 Chụp / Tải ảnh riêng cho [{kho}]", type=["jpg", "png", "jpeg"], key=f"img_{kho}")

                            kho_inputs[kho] = {
                                "nhiet_do": n_do, "do_am": d_am, "san_luong": s_luong,
                                "trang_thai": t_thai, "file_img": file_img
                            }

                    if st.form_submit_button("🚀 GỬI TOÀN BỘ BÁO CÁO CA TRỰC"):
                        now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        cursor.execute("INSERT INTO bao_cao_tong_hop (ma_ca_truc, thoi_gian, ca_truc, nguoi_bao_cao, ghi_chu_chung) VALUES (?, ?, ?, ?, ?)",
                                       (ma_ca, now_str, ca_truc, nguoi_lap, ghi_chu_chung))
                        
                        for kho, data in kho_inputs.items():
                            img_path = ""
                            if data["file_img"]:
                                img_path = os.path.join(UPLOAD_DIR, f"{ma_ca}_{kho}_{data['file_img'].name}")
                                with open(img_path, "wb") as f:
                                    f.write(data["file_img"].getbuffer())

                            cursor.execute('''
                                INSERT INTO chi_tiet_kho (ma_ca_truc, ten_kho, nhiet_do, do_am, san_luong, trang_thai_may, duong_dan_anh)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                            ''', (ma_ca, kho, data["nhiet_do"], data["do_am"], data["san_luong"], data["trang_thai"], img_path))

                        conn.commit()
                        st.success("✅ Đã lưu toàn bộ báo cáo!")
                        st.rerun()
        else:
            st.warning("🔒 Tài khoản của bạn là quyền VIEWER (Chỉ xem), không được phép lập báo cáo.")

    # TAB 3: ADMIN QUẢN LÝ KHO & TRẠNG THÁI
    if is_admin:
        with tabs[2]:
            st.subheader("⚙️ Quản Lý Danh Mục Kho & Trạng Thái Vận Hành")
            
            st.markdown("### 🏬 1. Tùy Chỉnh Kho / Thiết Bị")
            c_k1, c_k2, c_k3 = st.columns(3)
            
            with c_k1:
                with st.container(border=True):
                    st.markdown("##### ➕ Thêm Kho Mới")
                    ten_kho_moi = st.text_input("Nhập tên kho mới", key="add_kho_input")
                    if st.button("Thêm Vào Danh Mục", key="btn_add_kho"):
                        if ten_kho_moi.strip():
                            try:
                                cursor.execute("INSERT INTO danh_muc_kho (ten_kho) VALUES (?)", (ten_kho_moi.strip(),))
                                conn.commit()
                                st.success(f"Đã thêm kho: {ten_kho_moi}")
                                st.rerun()
                            except:
                                st.error("Tên kho này đã tồn tại!")

            with c_k2:
                with st.container(border=True):
                    st.markdown("##### ✏️ Đổi Tên Kho")
                    kho_doi = st.selectbox("Chọn kho cần đổi tên", ds_kho if ds_kho else ["Chưa có"], key="sel_doi_kho")
                    ten_kho_renamed = st.text_input("Tên mới", value=kho_doi if ds_kho else "", key="txt_ren_kho")
                    if st.button("Cập Nhật Tên Kho", key="btn_ren_kho"):
                        if ds_kho and ten_kho_renamed.strip() and ten_kho_renamed != kho_doi:
                            try:
                                cursor.execute("UPDATE danh_muc_kho SET ten_kho = ? WHERE ten_kho = ?", (ten_kho_renamed.strip(), kho_doi))
                                conn.commit()
                                st.success("Đã đổi tên kho thành công!")
                                st.rerun()
                            except:
                                st.error("Tên mới trùng với kho khác!")

            with c_k3:
                with st.container(border=True):
                    st.markdown("##### 🗑️ XÓA HẲN KHO")
                    kho_xoa = st.selectbox("Chọn kho cần XÓA MẤT LUÔN", ds_kho if ds_kho else ["Chưa có"], key="sel_xoa_kho")
                    if st.button("❌ XÓA VĨNH VIỄN KHO NÀY", key="btn_del_kho"):
                        if ds_kho:
                            cursor.execute("DELETE FROM danh_muc_kho WHERE ten_kho = ?", (kho_xoa,))
                            conn.commit()
                            st.success(f"Đã xóa vĩnh viễn kho '{kho_xoa}'!")
                            st.rerun()

            st.markdown("---")
            st.markdown("### 🔴 2. Tùy Chỉnh Danh Mục Trạng Thái Vận Hành")
            c_t1, c_t2 = st.columns(2)
            
            with c_t1:
                with st.container(border=True):
                    st.markdown("##### ➕ Thêm Trạng Thái Mới")
                    tt_moi = st.text_input("Ví dụ: Đang sửa chữa máy nén 2", key="add_tt_input")
                    if st.button("Thêm Trạng Thái", key="btn_add_tt"):
                        if tt_moi.strip():
                            try:
                                cursor.execute("INSERT INTO danh_muc_trang_thai (ten_trang_thai) VALUES (?)", (tt_moi.strip(),))
                                conn.commit()
                                st.success("Đã thêm trạng thái mới!")
                                st.rerun()
                            except:
                                st.error("Trạng thái này đã có!")

            with c_t2:
                with st.container(border=True):
                    st.markdown("##### 🗑️ Xóa Trạng Thái")
                    tt_xoa = st.selectbox("Chọn trạng thái cần xóa", ds_tt if ds_tt else ["Chưa có"], key="sel_xoa_tt")
                    if st.button("❌ Xóa Trạng Thái Này", key="btn_del_tt"):
                        if ds_tt:
                            cursor.execute("DELETE FROM danh_muc_trang_thai WHERE ten_trang_thai = ?", (tt_xoa,))
                            conn.commit()
                            st.success(f"Đã xóa trạng thái '{tt_xoa}'!")
                            st.rerun()

    # TAB 4: ADMIN QUẢN LÝ & XÓA VĨNH VIỄN TÀI KHOẢN
    if is_admin:
        with tabs[3]:
            st.subheader("👥 Cấp Tài Khoản & Xóa Nhân Viên")
            
            with st.form("form_tao_tk"):
                st.markdown("##### ➕ Tạo tài khoản mới")
                c_u, c_p, c_n, c_r = st.columns(4)
                new_u = c_u.text_input("Tên Đăng Nhập")
                new_p = c_p.text_input("Mật Khẩu")
                new_n = c_n.text_input("Họ Và Tên")
                new_r = c_r.selectbox("Phân Quyền", ["nhanvien", "viewer", "admin"])
                
                if st.form_submit_button("TẠO TÀI KHOẢN"):
                    if new_u and new_p and new_n:
                        try:
                            cursor.execute("INSERT INTO tai_khoan (username, password, ho_ten, vai_tro, trang_thai) VALUES (?, ?, ?, ?, 'hoat_dong')", (new_u, new_p, new_n, new_r))
                            conn.commit()
                            st.success(f"Đã tạo tài khoản cho {new_n} ({new_r})!")
                            st.rerun()
                        except:
                            st.error("Tên đăng nhập này đã có người sử dụng!")

            st.markdown("---")
            st.subheader("📋 Danh Sách Tài Khoản Trong Hệ Thống")
            df_users = pd.read_sql_query("SELECT username AS 'Tên ĐN', ho_ten AS 'Họ Tên', vai_tro AS 'Quyền', trang_thai AS 'Trạng Thái' FROM tai_khoan", conn)
            st.dataframe(df_users, use_container_width=True)
            
            st.markdown("##### 🛠️ Thao Tác Với Tài Khoản Nhân Viên")
            col_usr, col_b1, col_b2 = st.columns([2, 1.5, 1.5])
            usr_target = col_usr.selectbox("Chọn tài khoản", df_users['Tên ĐN'].tolist())
            
            if usr_target == "admin":
                st.info("💡 Tài khoản 'admin' mặc định không thể xóa.")
            else:
                if col_b1.button("🗑️ XÓA VĨNH VIỄN TÀI KHOẢN"):
                    cursor.execute("DELETE FROM tai_khoan WHERE username = ?", (usr_target,))
                    conn.commit()
                    st.success(f"Đã XÓA MẤT LUÔN tài khoản {usr_target} khỏi hệ thống!")
                    st.rerun()

                if col_b2.button("🔒 Khóa / Mở Khóa"):
                    curr_st = cursor.execute("SELECT trang_thai FROM tai_khoan WHERE username = ?", (usr_target,)).fetchone()[0]
                    new_st = "bi_khoa" if curr_st == "hoat_dong" else "hoat_dong"
                    cursor.execute("UPDATE tai_khoan SET trang_thai = ? WHERE username = ?", (new_st, usr_target))
                    conn.commit()
                    st.success(f"Đã đổi trạng thái tài khoản {usr_target} thành: {new_st}")
                    st.rerun()