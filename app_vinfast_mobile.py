import os
import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime
from streamlit_javascript import st_javascript

# Tích hợp Google Sheets Realtime
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False

# ==============================================================================
# 1. CẤU HÌNH TRANG STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="VinFast Hưng Thịnh Phát",
    page_icon="🚗",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .stApp { background-color: #121212; color: #ffffff; }
    .main-title { text-align: center; color: #ffffff; font-weight: bold; font-size: 22px; margin-bottom: 20px; background-color: #1e1e1e; padding: 15px; border-radius: 12px; border: 1px solid #333333; }
    .main-title a { color: #ffffff !important; text-decoration: none !important; }
    .main-title a:hover { color: #00d26a !important; }
    .sub-title { color: #ffc107; font-weight: bold; font-size: 18px; text-align: center; margin-bottom: 15px; }
    
    div.stButton > button {
        width: 100% !important;
        height: 55px !important;
        font-weight: bold !important;
        font-size: 16px !important;
        border-radius: 12px !important;
        margin-bottom: 12px !important;
        background: linear-gradient(135deg, #1e293b, #0f172a) !important;
        color: #ffffff !important;
        border: 2px solid #00d26a !important;
        box-shadow: 0px 4px 10px rgba(0, 210, 106, 0.2) !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stButton > button:hover, div.stButton > button:active {
        background: #00d26a !important; color: #000000 !important; border-color: #00d26a !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. KHỞI TẠO SESSION & LẤY ĐỊNH DANH THIẾT BỊ (DEVICE ID)
# ==============================================================================
if "user_role" not in st.session_state:
    st.session_state.user_role = None  # "VFHTP" hoặc "NHAN_VIEN"
if "user_id" not in st.session_state:
    st.session_state.user_id = ""
if "page" not in st.session_state:
    st.session_state.page = "login"

# Lấy thông tin User Agent của thiết bị để tạo ID duy nhất
device_signature = st_javascript("navigator.userAgent")

def set_page(page_name):
    st.session_state.page = page_name

# ==============================================================================
# 3. HÀM XÁC THỰC DỮ LIỆU & BẢO MẬT THIẾT BỊ
# ==============================================================================
def connect_gsheet():
    if "gcp_service_account" in st.secrets and HAS_GSPREAD:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)
        return client.open("ThongKe_Vinfast_Showroom")
    return None

def verify_login(role, user_code, password, device_id):
    sheet = connect_gsheet()
    if not sheet:
        return False, "Chưa kết nối được với Server Google Sheets!"
    
    ws_name = "QuanLy" if role == "VFHTP" else "NhanVien"
    ws = sheet.worksheet(ws_name)
    records = ws.get_all_records()
    
    col_code = "Mã VFHTP" if role == "VFHTP" else "Mã NV"
    
    for idx, row in enumerate(records, start=2):
        if str(row[col_code]).strip() == user_code.strip():
            if str(row["Mật Khẩu"]).strip() != password.strip():
                return False, "Mật khẩu không chính xác!"
            
            if role == "NHAN_VIEN" and str(row.get("Trạng Thái", "")).strip() != "Đã duyệt":
                return False, "Tài khoản Nhân viên chưa được cấp Mã VFHTP duyệt!"
            
            current_device = str(row.get("Device_ID", "")).strip()
            if current_device and current_device != str(device_id).strip():
                return False, "⛔ Tài khoản này đã đăng nhập trên thiết bị khác! Mỗi tài khoản chỉ dùng trên 1 thiết bị."
            
            # Cập nhật Device ID nếu chưa có
            if not current_device:
                ws.update_cell(idx, list(row.keys()).index("Device_ID") + 1, str(device_id))
                
            return True, "Thành công"
            
    return False, f"Không tìm thấy {col_code} trên hệ thống!"

def luu_du_lieu_realtime(loai, hang_du_lieu):
    file_name = f"ThongKe_Showroom_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    wb = openpyxl.load_workbook(file_name) if os.path.exists(file_name) else openpyxl.Workbook()
    if "Sheet" in wb.sheetnames: wb.remove(wb["Sheet"])
    if "Khách Đến" not in wb.sheetnames:
        wb.create_sheet("Khách Đến").append(["Thời gian", "Loại khách", "Mã NV", "Họ tên KH", "Dòng xe", "Mục đích sử dụng", "Màu sắc", "Nhu cầu vay", "SDT"])
    if "Khách Về" not in wb.sheetnames:
        wb.create_sheet("Khách Về").append(["Thời gian", "Trạng thái cọc", "Mã NV", "Họ tên", "SDT", "CCCD", "Xe đã xem", "Tiền cọc"])
    ws = wb["Khách Đến" if loai == "KHÁCH ĐẾN" else "Khách Về"]
    ws.append(hang_du_lieu)
    wb.save(file_name)

    sheet = connect_gsheet()
    if sheet:
        try:
            worksheet = sheet.worksheet("Khách Đến" if loai == "KHÁCH ĐẾN" else "Khách Về")
            worksheet.append_row(hang_du_lieu)
        except Exception as e:
            st.warning(f"Lỗi đồng bộ Realtime: {e}")

# Dữ liệu xe VinFast
data_vinfast = {
    "Dòng xe": ["VF 2", "VF 3", "VF 5 Plus", "VF 6", "VF 7", "VF 8", "VF 9"],
    "Giá niêm yết (VND)": [188000000, 302000000, 529000000, 689000000, 789000000, 1069000000, 1499000000],
    "Quãng đường tối đa": ["~120 - 150 km", "~210 km", "~326 km", "~399 km", "~450 km", "~471 km", "~626 km"]
}
df_vinfast = pd.DataFrame(data_vinfast)
cac_dong_xe = sorted(list(set(data_vinfast["Dòng xe"])))
cac_mau_xe = ["Trắng", "Đen", "Xám", "Bạc", "Xanh", "Đỏ"]

# ==============================================================================
# 4. MÀN HÌNH ĐĂNG NHẬP
# ==============================================================================
if st.session_state.page == "login":
    st.markdown("""
        <div class="main-title">
            🚘 <a href="https://vinfasthungthinhphat.vn" target="_blank">VINFAST HƯNG THỊNH PHÁT</a><br>
            <span style="font-size: 14px; font-weight: normal; color: #00d26a;">HỆ THỐNG QUẢN LÝ SHOWROOM</span>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="sub-title">🔐 ĐĂNG NHẬP HỆ THỐNG</div>', unsafe_allow_html=True)
    
    loai_tk = st.radio("Cấp độ đăng nhập:", ["Quản lý (Mã VFHTP)", "Nhân viên (Mã NV)"], horizontal=True)
    
    with st.form("form_login"):
        if loai_tk == "Quản lý (Mã VFHTP)":
            user_code = st.text_input("Nhập Mã VFHTP: *")
            password = st.text_input("Nhập Mật khẩu (1): *", type="password")
        else:
            user_code = st.text_input("Nhập Mã Nhân Viên (NV): *")
            password = st.text_input("Nhập Mật khẩu Nhân viên: *", type="password")
            
        btn_submit = st.form_submit_button("🔓 ĐĂNG NHẬP", use_container_width=True, type="primary")
        
        if btn_submit:
            role = "VFHTP" if loai_tk == "Quản lý (Mã VFHTP)" else "NHAN_VIEN"
            if not user_code.strip() or not password.strip():
                st.error("⚠️ Vui lòng điền đầy đủ Mã và Mật khẩu!")
            else:
                success, msg = verify_login(role, user_code, password, device_signature)
                if success:
                    st.session_state.user_role = role
                    st.session_state.user_id = user_code.strip()
                    set_page("home")
                    st.rerun()
                else:
                    st.error(f"❌ Dang nhập thất bại: {msg}")

# ==============================================================================
# 5. MÀN HÌNH CHÍNH (HOME)
# ==============================================================================
elif st.session_state.page == "home":
    st.markdown(f"""
        <div class="main-title">
            🚘 <a href="https://vinfasthungthinhphat.vn" target="_blank">VINFAST HƯNG THỊNH PHÁT</a><br>
            <span style="font-size: 14px; color: #00d26a;">Xin chào: {st.session_state.user_id} ({'Quản lý' if st.session_state.user_role == 'VFHTP' else 'Nhân viên'})</span>
        </div>
    """, unsafe_allow_html=True)

    if st.button("1. 🚗 KHÁCH ĐẾN SHOWROOM", use_container_width=True, type="primary"):
        set_page("khach_den")
        st.rerun()

    if st.button("2. 🚶 KHÁCH RỜI SHOWROOM", use_container_width=True):
        set_page("khach_ve")
        st.rerun()

    if st.button("3. 📋 TRA CỨU BẢNG GIÁ / THÔNG SỐ XE", use_container_width=True):
        set_page("tra_cuu")
        st.rerun()

    # CHỈ MÃ VFHTP MỚI HIỂN THỊ NÚT BÁO CÁO REALTIME
    if st.session_state.user_role == "VFHTP":
        if st.button("4. 📈 BÁO CÁO REALTIME (ĐẶC QUYỀN VFHTP)", use_container_width=True):
            set_page("bao_cao")
            st.rerun()
            
    if st.button("🚪 ĐĂNG XUẤT", use_container_width=True):
        st.session_state.user_role = None
        st.session_state.user_id = ""
        set_page("login")
        st.rerun()

# ==============================================================================
# 6. KHÂU KHÁCH ĐẾN SHOWROOM
# ==============================================================================
elif st.session_state.page == "khach_den":
    st.markdown('<div class="sub-title">📥 THÔNG TIN KHÁCH ĐẾN SHOWROOM</div>', unsafe_allow_html=True)

    co_hen = st.radio("Khách hàng đã có hẹn trước chưa?", ["Chưa có hẹn (Vãng lai)", "Đã có hẹn trước"], horizontal=True)
    is_hen = (co_hen == "Đã có hẹn trước")

    xe_chon = st.selectbox("Dòng xe bạn quan tâm: *", cac_dong_xe)

    st.markdown(f"**📊 Chi tiết thông số so sánh dòng xe {xe_chon}:**")
    df_sub = df_vinfast[df_vinfast["Dòng xe"] == xe_chon]
    st.dataframe(df_sub, use_container_width=True, hide_index=True)
    st.write("---")

    with st.form("form_khach_den", clear_on_submit=True):
        # Tự động gán Mã NV nếu tài khoản đăng nhập là Nhân viên
        if st.session_state.user_role == "NHAN_VIEN":
            st.info(f"👤 Mã nhân viên ghi nhận: **{st.session_state.user_id}**")
            ma_nv = st.session_state.user_id
        else:
            ma_nv = st.text_input("Mã nhân viên tư vấn/hẹn trước: *", value=st.session_state.user_id if is_hen else "")
            
        ho_ten = st.text_input("Họ và tên khách hàng: *")
        muc_dich = st.selectbox("Mục đích sử dụng xe: *", ["Chạy gia đình", "Chạy dịch vụ"])
        phan_khuc = st.text_input("Phân khúc xe yêu cầu: *", placeholder="Ví dụ: SUV hạng A, C...")
        mau_sac = st.selectbox("Màu sắc lựa chọn: *", cac_mau_xe)
        vay_ngan_hang = st.radio("Cần hỗ trợ vay ngân hàng?", ["Không", "Có"], horizontal=True)
        sdt = st.text_input("Số điện thoại liên hệ: *")

        submitted = st.form_submit_button("💾 HOÀN TẤT & LƯU THÔNG TIN", use_container_width=True, type="primary")

        if submitted:
            if not ma_nv.strip() or not ho_ten.strip() or not phan_khuc.strip() or not sdt.strip():
                st.error("⚠️ Vui lòng nhập đầy đủ các thông tin bắt buộc (*)")
            else:
                data_row = [
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Đã hẹn" if is_hen else "Vãng lai",
                    ma_nv.strip(),
                    ho_ten.strip(),
                    f"{xe_chon} ({phan_khuc.strip()})",
                    muc_dich, mau_sac, vay_ngan_hang, sdt.strip()
                ]
                luu_du_lieu_realtime("KHÁCH ĐẾN", data_row)
                st.success("✅ Đã lưu và đồng bộ thông tin khách đến thành công!")

    st.write("---")
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()

# ==============================================================================
# 7. KHÂU KHÁCH RỜI SHOWROOM
# ==============================================================================
elif st.session_state.page == "khach_ve":
    st.markdown('<div class="sub-title">📤 THÔNG TIN KHÁCH RỜI SHOWROOM</div>', unsafe_allow_html=True)

    da_coc = st.radio("Khách hàng đã đặt cọc xe chưa?", ["Chưa cọc (Chưa đặt)", "Đã đặt cọc xe"], horizontal=True)
    is_coc = (da_coc == "Đã đặt cọc xe")

    ds_xe_da_xem = st.multiselect("Dòng xe khách đã xem: *", cac_dong_xe, default=[cac_dong_xe[0]])

    if ds_xe_da_xem:
        st.markdown(f"**📊 Chi tiết thông số so sánh dòng xe đã xem:**")
        df_sub_ve = df_vinfast[df_vinfast["Dòng xe"].isin(ds_xe_da_xem)]
        st.dataframe(df_sub_ve, use_container_width=True, hide_index=True)
    st.write("---")

    with st.form("form_khach_ve", clear_on_submit=True):
        if st.session_state.user_role == "NHAN_VIEN":
            st.info(f"👤 Mã nhân viên tư vấn: **{st.session_state.user_id}**")
            ma_nv = st.session_state.user_id
        else:
            ma_nv = st.text_input("Mã NV tư vấn: *")
            
        ho_ten = st.text_input("Họ tên khách hàng: *")
        sdt = st.text_input("Số điện thoại: *")
        cccd = st.text_input("Số CCCD: *")
        tien_coc = st.text_input("Số tiền cọc (VNĐ): *") if is_coc else "-"

        submitted = st.form_submit_button("💾 HOÀN TẤT & LƯU THÔNG TIN", use_container_width=True, type="primary")

        if submitted:
            if not ma_nv.strip() or not ho_ten.strip() or not sdt.strip() or not cccd.strip() or not ds_xe_da_xem or (is_coc and not tien_coc.strip()):
                st.error("⚠️ Vui lòng điền đầy đủ các thông tin bắt buộc (*)")
            else:
                data_row = [
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Đã đặt cọc" if is_coc else "Chưa cọc",
                    ma_nv.strip(), ho_ten.strip(), sdt.strip(), cccd.strip(),
                    ", ".join(ds_xe_da_xem), tien_coc.strip()
                ]
                luu_du_lieu_realtime("KHÁCH VỀ", data_row)
                st.success("✅ Đã lưu thông tin khách rời showroom thành công!")

    st.write("---")
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()

# ==============================================================================
# 8. MÀN HÌNH BÁO CÁO REALTIME (ĐẶC QUYỀN MÃ VFHTP)
# ==============================================================================
elif st.session_state.page == "bao_cao":
    if st.session_state.user_role != "VFHTP":
        st.error("⛔ Bạn không có quyền truy cập màn hình báo cáo!")
        if st.button("🏠 QUAY LẠI", use_container_width=True):
            set_page("home")
            st.rerun()
    else:
        st.markdown('<div class="sub-title">📈 BÁO CÁO REALTIME (ĐẶC QUYỀN MÃ VFHTP)</div>', unsafe_allow_html=True)
        
        tab1, tab2 = st.tabs(["📥 Danh sách Khách Đến", "📤 Danh sách Khách Về"])
        sheet = connect_gsheet()
        
        if sheet:
            with tab1:
                df_den = pd.DataFrame(sheet.worksheet("Khách Đến").get_all_records())
                st.write(f"**Tổng số lượt khách đến:** {len(df_den)}")
                st.dataframe(df_den, use_container_width=True, hide_index=True)
            with tab2:
                df_ve = pd.DataFrame(sheet.worksheet("Khách Về").get_all_records())
                st.write(f"**Tổng số lượt khách về:** {len(df_ve)}")
                st.dataframe(df_ve, use_container_width=True, hide_index=True)
        else:
            st.error("Chưa thể kết nối tới Google Sheets.")

        st.write("---")
        if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
            set_page("home")
            st.rerun()
# ==============================================================================
# HÀM CẬP NHẬT TRẠNG THÁI / MẬT KHẨU NHÂN VIÊN (DÀNH CHO VFHTP)
# ==============================================================================
def cap_nhat_nhan_vien(ma_nv, mat_khau_moi=None, trang_thai_moi=None, reset_device=False):
    sheet = connect_gsheet()
    if not sheet:
        return False, "Chưa kết nối được với Google Sheets!"
    
    ws = sheet.worksheet("NhanVien")
    records = ws.get_all_records()
    
    for idx, row in enumerate(records, start=2): # Bắt đầu từ dòng 2 (sau tiêu đề)
        if str(row.get("Mã NV", "")).strip() == ma_nv.strip():
            if mat_khau_moi:
                ws.update_cell(idx, list(row.keys()).index("Mật Khẩu") + 1, mat_khau_moi.strip())
            if trang_thai_moi:
                ws.update_cell(idx, list(row.keys()).index("Trạng Thái") + 1, trang_thai_moi.strip())
            if reset_device:
                ws.update_cell(idx, list(row.keys()).index("Device_ID") + 1, "")
            return True, f"Đã cập nhật thành công cho Mã NV: {ma_nv}"
            
    return False, f"Không tìm thấy Mã NV: {ma_nv}"

