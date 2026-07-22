Dưới đây là toàn bộ mã nguồn app_vinfast_mobile.py đã được cập nhật bổ sung đầy đủ:

Bổ sung các thông số chi tiết:

Loại trần xe: Trần kính toàn cảnh / Trần thép.

Tính năng nổi trội cùng phân khúc: ADAS (Lái xe thông minh), Smart Services, trợ lý ảo ViVi, camera 360, v.v.

So sánh xe cùng loại / Đối thủ cùng phân khúc: So sánh trực tiếp với các dòng xe xăng/điện cùng tầm giá.

Cập nhật giao diện hiển thị:

Thêm tab "Đặc điểm nổi bật & So sánh" khi khách chọn xe (ở cả phần Khách đến, Khách rời và Tra cứu).

Giữ nguyên toàn bộ cấu hình ảnh .jpg chuẩn theo thư mục hinhanh trên Repository vfhungthinh.

Python
import os
import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime

# Tích hợp Google Sheets Realtime
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False

# ==============================================================================
# 1. CẤU HÌNH TRANG & MÀN HÌNH NỀN (BACKGROUND)
# ==============================================================================
st.set_page_config(
    page_title="VinFast Hưng Thịnh Phát",
    page_icon="🚗",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Đường dẫn ảnh nền (.jpg)
URL_ANH_NEN = "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/anh_nen_tong_the.jpg"

st.markdown(f"""
    <style>
    /* Cấu hình màn hình nền có lớp phủ tối để nổi bật chữ */
    .stApp {{
        background: linear-gradient(rgba(18, 18, 18, 0.85), rgba(18, 18, 18, 0.85)), 
                    url('{URL_ANH_NEN}');
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
        color: #ffffff;
    }}
    
    .main-title {{ 
        text-align: center; 
        color: #ffffff; 
        font-weight: bold; 
        font-size: 22px; 
        margin-bottom: 20px; 
        background-color: rgba(30, 30, 30, 0.9); 
        padding: 15px; 
        border-radius: 12px; 
        border: 1px solid #333333; 
    }}
    .main-title a {{ color: #ffffff !important; text-decoration: none !important; }}
    .main-title a:hover {{ color: #00d26a !important; }}
    .sub-title {{ color: #ffc107; font-weight: bold; font-size: 18px; text-align: center; margin-bottom: 15px; }}
    
    div.stButton > button {{
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
    }}
    div.stButton > button:hover, div.stButton > button:active {{
        background: #00d26a !important; color: #000000 !important; border-color: #00d26a !important;
    }}
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. KHỞI TẠO SESSION & DỮ LIỆU CÁC DÒNG XE & HÌNH ẢNH & THÔNG SỐ SO SÁNH
# ==============================================================================
if "page" not in st.session_state:
    st.session_state.page = "home"

def set_page(page_name):
    st.session_state.page = page_name

# URL hình ảnh .jpg cho từng dòng xe
HINH_ANH_XE = {
    "VF 2": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%202.jpg",
    "VF 3": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%203.jpg",
    "VF 5 Plus": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%205%20Plus.jpg",
    "VF 6": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%206.jpg",
    "VF 7": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%207.jpg",
    "VF 8": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%208.jpg",
    "VF 9": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%209.jpg"
}

def hien_thi_anh_xe(dong_xe):
    """Hàm hiển thị hình ảnh xe dạng .jpg"""
    file_local_jpg_exact = f"hinhanh/{dong_xe}.jpg"
    file_local_jpg_slug = f"hinhanh/{dong_xe.lower().replace(' ', '')}.jpg"
    
    if os.path.exists(file_local_jpg_exact):
        st.image(file_local_jpg_exact, caption=f"Xe VinFast {dong_xe}", use_container_width=True)
    elif os.path.exists(file_local_jpg_slug):
        st.image(file_local_jpg_slug, caption=f"Xe VinFast {dong_xe}", use_container_width=True)
    elif dong_xe in HINH_ANH_XE:
        st.image(HINH_ANH_XE[dong_xe], caption=f"Xe VinFast {dong_xe}", use_container_width=True)

# Dữ liệu xe VinFast chi tiết (Bổ sung Trần xe, Tính năng nổi bật, So sánh)
data_vinfast = {
    "Dòng xe": ["VF 2", "VF 3", "VF 5 Plus", "VF 6", "VF 7", "VF 8", "VF 9"],
    "Phân khúc": ["Mini Car", "Mini SUV", "SUV Hạng A", "SUV Hạng B", "SUV Hạng C", "SUV Hạng D", "SUV Hạng E"],
    "Giá niêm yết (VND)": [188000000, 302000000, 529000000, 689000000, 789000000, 1069000000, 1499000000],
    "Quãng đường/sạc": ["~120 - 150 km", "~210 km", "~326 km", "~399 km", "~450 km", "~471 km", "~626 km"],
    "Trần xe": [
        "Trần thép", 
        "Trần thép", 
        "Trần thép", 
        "Trần thép (Bản Eco) / Kính (Bản Plus)", 
        "Trần thép (Bản Eco) / Kính toàn cảnh (Bản Plus)", 
        "Trần thép (Bản Eco) / Kính toàn cảnh (Bản Plus)", 
        "Trần kính toàn cảnh chống tia UV"
    ],
    "Tính năng nổi trội": [
        "Nhỏ gọn tối ưu đô thị, chi phí vận hành siêu rẻ",
        "Thiết kế SUV vuông vức cá tính, khoảng sáng gầm cao, màn hình 10 inch",
        "6 túi khí, cảnh báo điểm mù, luồng giao thông đến khi mở cửa, giá cực tốt",
        "Trợ lý ảo ViVi, ADAS cấp độ 2, chip xử lý hiện đại, nội thất da cao cấp",
        "Thiết kế phi thuyền cá tính, công suất tới 349 mã lực, HUD hắt kính, ADAS đầy đủ",
        "Động cơ 402 mã lực, dẫn động 4 bánh AWD, sưởi/thông gió hàng ghế trước",
        "Nội thất Thương gia VIP, ghế massage, màn hình 15.6 inch, hệ thống treo khí nén"
    ],
    "So sánh cùng phân khúc": [
        "Vượt trội Wuling Mini EV về quãng đường đi và độ an toàn khung gầm",
        "Rộng rãi và nhiều trang bị công nghệ hơn các dòng xe xăng hạng A giá rẻ",
        "Chi phí lăn bánh rẻ hơn Raize, Sonet; công nghệ an toàn ADAS vượt trội",
        "Mạnh mẽ hơn Seltos, Creta; chi phí nhiên liệu chỉ bằng 1/3 xe xăng",
        "Tăng tốc nhanh hơn CX-5, CR-V, Tucson; công nghệ thông minh áp đảo",
        "Động cơ mạnh tương đương Porsche Macan, chi phí bảo dưỡng cực thấp hơn SantaFe, Sorento",
        "Đẳng cấp tương đương Lexus RX, BMW X5 nhưng giá chỉ bằng 1/3"
    ]
}

df_vinfast = pd.DataFrame(data_vinfast)
cac_dong_xe = sorted(list(set(data_vinfast["Dòng xe"])))
cac_mau_xe = ["Trắng", "Đen", "Xám", "Bạc", "Xanh", "Đỏ"]

def hien_thi_thong_tin_so_sanh(dong_xe):
    """Hàm hiển thị bảng so sánh & tính năng nổi bật của xe"""
    row = df_vinfast[df_vinfast["Dòng xe"] == dong_xe].iloc[0]
    
    col1, col2 = st.columns(2)
    with col1:
        st.info(f"🧱 **Loại trần xe:** {row['Trạng thái trần xe'] if 'Trạng thái trần xe' in row else row['Trần xe']}")
        st.success(f"🔋 **Quãng đường:** {row['Quãng đường/sạc']}")
    with col2:
        st.warning(f"🏷️ **Phân khúc:** {row['Phân khúc']}")
        st.error(f"💰 **Giá niêm yết:** {row['Giá niêm yết (VND)']:,.0f} VNĐ")
        
    st.markdown("---")
    st.markdown(f"✨ **Tính năng nổi trội cùng phân khúc:**\n- {row['Tính năng nổi trội']}")
    st.markdown(f"⚔️ **So sánh đối thủ cùng tầm giá:**\n- {row['So sánh cùng phân khúc']}")

# ==============================================================================
# 3. HÀM KẾT NỐI VÀ LƯU DỮ LIỆU
# ==============================================================================
@st.cache_resource(ttl=600)
def connect_gsheet():
    """Hàm kết nối Google Sheet qua ID chính xác"""
    if "gcp_service_account" in st.secrets and HAS_GSPREAD:
        try:
            scope = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            creds_dict = dict(st.secrets["gcp_service_account"])
            if "private_key" in creds_dict:
                creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            
            creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
            client = gspread.authorize(creds)
            
            SPREADSHEET_ID = "1wo7v0XEEdFABLMWHg80hchf4nwgnPoej9fzCTxOEs4o"
            return client.open_by_key(SPREADSHEET_ID)
        except Exception as e:
            st.error(f"⚠️ Lỗi kết nối Google Sheets: {e}")
            return None
    return None

def luu_du_lieu_realtime(loai, hang_du_lieu):
    # 1. Lưu file Excel cục bộ dự phòng
    try:
        file_name = f"ThongKe_Showroom_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        if os.path.exists(file_name):
            wb = openpyxl.load_workbook(file_name)
        else:
            wb = openpyxl.Workbook()
            if "Sheet" in wb.sheetnames:
                wb.remove("Sheet")
        
        ws_title = "Khách Đến" if loai == "KHÁCH ĐẾN" else "Khách Về"
        if ws_title not in wb.sheetnames:
            ws = wb.create_sheet(ws_title)
            if loai == "KHÁCH ĐẾN":
                ws.append(["Thời gian", "Loại khách", "Mã NV", "Họ tên KH", "Dòng xe", "Mục đích sử dụng", "Màu sắc", "Nhu cầu vay", "SDT"])
            else:
                ws.append(["Thời gian", "Trạng thái cọc", "Mã NV", "Họ tên", "SDT", "CCCD", "Xe đã xem", "Tiền cọc"])
        else:
            ws = wb[ws_title]
            
        ws.append(hang_du_lieu)
        wb.save(file_name)
    except Exception as e:
        print(f"Lỗi ghi Excel cục bộ: {e}")

    # 2. Đẩy dữ liệu Realtime lên Google Sheets
    sheet = connect_gsheet()
    if sheet:
        try:
            ws_target = "Khách Đến" if loai == "KHÁCH ĐẾN" else "Khách Về"
            worksheet = sheet.worksheet(ws_target)
            worksheet.append_row(hang_du_lieu)
        except Exception as e:
            st.warning(f"Lỗi đồng bộ Realtime: {e}")

# ==============================================================================
# 4. MÀN HÌNH CHÍNH (HOME)
# ==============================================================================
if st.session_state.page == "home":
    st.markdown("""
        <div class="main-title">
            🚘 <a href="https://vinfasthungthinhphat.vn" target="_blank">VINFAST HƯNG THỊNH PHÁT</a><br>
            <span style="font-size: 14px; color: #00d26a;">HỆ THỐNG QUẢN LÝ SHOWROOM</span>
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

# ==============================================================================
# 5. KHÂU KHÁCH ĐẾN SHOWROOM
# ==============================================================================
elif st.session_state.page == "khach_den":
    st.markdown('<div class="sub-title">📥 THÔNG TIN KHÁCH ĐẾN SHOWROOM</div>', unsafe_allow_html=True)

    co_hen = st.radio("Khách hàng đã có hẹn trước chưa?", ["Chưa có hẹn (Vãng lai)", "Đã có hẹn trước"], horizontal=True)
    is_hen = (co_hen == "Đã có hẹn trước")

    xe_chon = st.selectbox("Dòng xe bạn quan tâm: *", cac_dong_xe)

    # Hiển thị hình ảnh xe chọn
    hien_thi_anh_xe(xe_chon)

    # Hiển thị phân tích nổi bật & So sánh đối thủ
    with st.expander(f"🔍 Xem So sánh & Tính năng nổi bật của {xe_chon}", expanded=True):
        hien_thi_thong_tin_so_sanh(xe_chon)

    st.write("---")

    with st.form("form_khach_den", clear_on_submit=True):
        ma_nv = st.text_input("Mã nhân viên tư vấn/hẹn trước: *")
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
# 6. KHÂU KHÁCH RỜI SHOWROOM
# ==============================================================================
elif st.session_state.page == "khach_ve":
    st.markdown('<div class="sub-title">📤 THÔNG TIN KHÁCH RỜI SHOWROOM</div>', unsafe_allow_html=True)

    da_coc = st.radio("Khách hàng đã đặt cọc xe chưa?", ["Chưa cọc (Chưa đặt)", "Đã đặt cọc xe"], horizontal=True)
    is_coc = (da_coc == "Đã đặt cọc xe")

    ds_xe_da_xem = st.multiselect("Dòng xe khách đã xem: *", cac_dong_xe, default=[cac_dong_xe[0]])

    if ds_xe_da_xem:
        # Hiển thị ảnh các xe được chọn
        cols = st.columns(len(ds_xe_da_xem))
        for idx, xe in enumerate(ds_xe_da_xem):
            with cols[idx]:
                hien_thi_anh_xe(xe)
                with st.expander(f"ℹ️ Tính năng & So sánh {xe}"):
                    hien_thi_thong_tin_so_sanh(xe)

    st.write("---")

    with st.form("form_khach_ve", clear_on_submit=True):
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
# 7. MÀN HÌNH TRA CỨU BẢNG GIÁ / THÔNG SỐ XE & XEM HÌNH ẢNH / SO SÁNH
# ==============================================================================
elif st.session_state.page == "tra_cuu":
    st.markdown('<div class="sub-title">📋 TRA CỨU BẢNG GIÁ & THÔNG SỐ CÁC DÒNG XE</div>', unsafe_allow_html=True)
    
    xe_xem_anh = st.selectbox("Chọn dòng xe để tra cứu chi tiết:", cac_dong_xe)
    hien_thi_anh_xe(xe_xem_anh)
    
    st.markdown(f"### ⚡ Chi tiết xe VinFast {xe_xem_anh}")
    hien_thi_thong_tin_so_sanh(xe_xem_anh)
    
    st.write("---")
    st.markdown("**📊 Bảng tổng hợp so sánh tất cả các dòng xe:**")
    df_display = df_vinfast.copy()
    df_display["Giá niêm yết (VND)"] = df_display["Giá niêm yết (VND)"].apply(lambda x: f"{x:,.0f} VNĐ")
    st.dataframe(df_display, use_container_width=True, hide_index=True)
    
    st.write("---")
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()
