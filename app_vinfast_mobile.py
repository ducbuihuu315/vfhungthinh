import os
import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime

# ==============================================================================
# 1. CẤU HÌNH TRANG STREAMLIT TỐI ƯU CHO ĐIỆN THOẠI
# ==============================================================================
st.set_page_config(
    page_title="VinFast Hưng Thịnh Phát",
    page_icon="🚗",
    layout="centered", # Tối ưu hiển thị dọc cho điện thoại
    initial_sidebar_state="collapsed"
)

# Custom CSS giao diện đậm chất VinFast & nút bấm lớn dễ nhấn trên điện thoại
st.markdown("""
    <style>
    .stApp {
        background-color: #1a1a1a;
        color: #ffffff;
    }
    .main-title {
        text-align: center;
        color: #ffffff;
        font-weight: bold;
        font-size: 22px;
        margin-bottom: 20px;
        background-color: #000000;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #333333;
    }
    .sub-title {
        color: #ffc107;
        font-weight: bold;
        font-size: 18px;
        text-align: center;
        margin-bottom: 15px;
    }
    div.stButton > button {
        width: 100%;
        height: 50px;
        font-weight: bold;
        font-size: 16px;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. KHỞI TẠO DỮ LIỆU XE VINFAST
# ==============================================================================
THU_MUC_ANH = "hinhanh"
ANH_SO_SANH_PATH = "so_sanh_vinfast.png"

data_vinfast = {
    "Phân khúc": [
        "Micro-car", "Mini SUV", "Mini SUV", "SUV Cỡ A", "SUV Cỡ B", "SUV Cỡ B",
        "SUV Cỡ C", "SUV Cỡ C", "SUV Cỡ C", "SUV Cỡ D", "SUV Cỡ D", "SUV Cỡ E", "SUV Cỡ E",
        "Hatchback Dịch vụ", "Sedan/Hatchback B", "MPV Dịch vụ", "MPV Cao cấp", "Xe tải Van", "Xe tải Van"
    ],
    "Dòng xe": [
        "VF 2", "VF 3", "VF 3", "VF 5 Plus", "VF 6", "VF 6",
        "VF 7", "VF 7", "VF 7", "VF 8", "VF 8", "VF 9", "VF 9",
        "Minio Green", "Herio Green", "Limo Green", "VF MPV 7", "EC Van", "EC Van"
    ],
    "Phiên bản": [
        "Tiêu chuẩn", "Eco", "Plus", "Tiêu chuẩn", "Eco", "Plus",
        "Eco", "Plus Standard", "Plus Premium", "Eco (CATL)", "Plus (CATL)", "Eco (CATL)", "Plus (CATL)",
        "Tiêu chuẩn", "Tiêu chuẩn", "7 chỗ", "Tiêu chuẩn", "Tiêu chuẩn", "Nâng cấp"
    ],
    "Giá niêm yết (VND)": [
        188000000, 302000000, 315000000, 529000000, 689000000, 745000000,
        789000000, 919000000, 999000000, 1069000000, 1199000000, 1499000000, 1699000000,
        269000000, 499000000, 749000000, 819000000, 1746000000, 285000000
    ],
    "Hệ thống dẫn động": [
        "1 cầu (RWD)", "1 cầu (RWD)", "1 cầu (RWD)", "1 cầu (FWD)", "1 cầu (FWD)", "1 cầu (FWD)",
        "1 cầu (FWD)", "1 cầu (FWD)", "2 cầu (AWD)", "1 cầu (FWD)", "2 cầu (AWD)", "2 cầu (AWD)", "2 cầu (AWD)",
        "1 cầu (FWD)", "1 cầu (FWD)", "1 cầu (FWD)", "1 cầu (FWD)", "1 cầu (RWD)", "1 cầu (RWD)"
    ],
    "Loại trần xe": [
        "Trần thép", "Trần thép", "Trần thép", "Trần thép", "Trần thép", "Trần thép",
        "Trần thép", "Trần thép", "Trần kính toàn cảnh", "Trần thép", "Trần thép (Tùy chọn kính)", "Trần thép", "Trần kính toàn cảnh",
        "Trần thép", "Trần thép", "Trần thép", "Trần thép", "Trần thép", "Trần thép"
    ],
    "Quãng đường tối đa": [
        "~120 - 150 km", "~210 km (NEDC)", "~210 km (NEDC)", "~326 km (NEDC)", "~399 km (WLTP)", "~381 km (WLTP)",
        "~450 km (WLTP)", "~431 km (WLTP)", "~431 km (WLTP)", "~471 km (WLTP)", "~450 km (WLTP)", "~626 km (WLTP)", "~602 km (WLTP)",
        "~160 - 180 km", "~300 - 320 km", "~400 km", "~420 km", "~120 - 140 km", "~140 km"
    ],
    "Tính năng nổi bật": [
        "Bán kính quay đầu cực nhỏ, hỗ trợ sạc nhanh tại nhà qua điện dân dụng gia đình.",
        "Thiết kế vuông vức cá tính, màn hình giải trí 10 inch, cần số tích hợp sau vô lăng.",
        "Mâm đúc thể thao hợp kim, màu sơn phối 2 tông thời trang cá nhân hóa cao.",
        "Cảnh báo điểm mù, cảnh báo phương tiện cắt ngang phía sau, trang bị sẵn 6 túi khí an toàn.",
        "Màn hình cảm ứng hướng về người lái, ghế bọc nỉ pha da cao cấp, ga tự động Cruise Control.",
        "Hệ thống hỗ trợ lái thông minh ADAS cấp độ 2 nâng cao, mâm xe lớn thể thao 19 inch.",
        "Thiết kế phong cách phi thuyền tương lai, tay nắm cửa ẩn cơ học mượt mà, mâm 19 inch.",
        "Động cơ nâng cấp lên 201 hp, cốp sau đóng mở điện tiện lợi, nội thất bọc da mịn.",
        "Hệ dẫn động AWD công suất 349 hp mạnh mẽ, màn hình thông tin trên kính lái HUD, trần kính.",
        "Màn hình trung tâm siêu lớn 15.6 inch, hỗ trợ trợ lý ảo thông minh tiếng Việt đa vùng miền.",
        "Ghế da thật cao cấp chỉnh điện đa hướng tích hợp chức năng sấy ấm và thông gió làm mát.",
        "Không gian nội thất SUV cỡ đại 7 chỗ ngồi thực tế rộng rãi, hệ thống loa cao cấp.",
        "Hàng ghế thứ 2 kiểu VIP cơ trưởng, tích hợp massage/sưởi/thông gió, trần kính Panorama.",
        "Chi phí sạc điện siêu rẻ, kết nối app quản lý taxi công nghệ phục vụ doanh thu.",
        "Thiết kế Sedan trường dáng lịch sự, cốp sau rộng để được nhiều vali hành lý của khách.",
        "Hàng ghế khoang khách bọc da êm ái biệt lập, bệ tỳ tay lớn, hỗ trợ cổng sạc nhanh.",
        "Cấu hình MPV gia đình thực dụng, sàn phẳng tối ưu chỗ để chân, cửa gió điều hòa độc lập.",
        "Khoang sau hoán cải hoàn toàn phẳng, vách ngăn cabin kiên cố, được lưu thông phố 24/7.",
        "Cửa lùa trượt mượt mà hai bên hông xe, nâng cấp điều hòa làm lạnh nhanh và trợ lực vô lăng lái."
    ]
}

df_vinfast = pd.DataFrame(data_vinfast)
cac_dong_xe = sorted(list(set(data_vinfast["Dòng xe"])))
cac_mau_xe = ["Trắng (Brahminy White)", "Đen (Jet Black)", "Xám (Neptune Grey)", "Bạc (Desat Silver)", "Xanh (Aurora Blue)", "Đỏ (Mystic Red)"]

# ==============================================================================
# 3. QUẢN LÝ ĐIỀU HƯỚNG & XỬ LÝ DỮ LIỆU EXCEL
# ==============================================================================
if "page" not in st.session_state:
    st.session_state.page = "home"

def set_page(page_name):
    st.session_state.page = page_name

def ghi_excel(loai, hang_du_lieu):
    file_name = f"ThongKe_Showroom_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    wb = openpyxl.load_workbook(file_name) if os.path.exists(file_name) else openpyxl.Workbook()
    
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    if "Khách Đến" not in wb.sheetnames:
        wb.create_sheet("Khách Đến").append(["Thời gian", "Loại khách", "Mã NV", "Họ tên KH", "Dòng xe", "Mục đích sử dụng", "Màu sắc", "Nhu cầu vay", "SDT"])
    if "Khách Về" not in wb.sheetnames:
        wb.create_sheet("Khách Về").append(["Thời gian", "Trạng thái cọc", "Mã NV", "Họ tên", "SDT", "CCCD", "Tiền cọc"])

    ws = wb["Khách Đến" if loai == "KHÁCH ĐẾN" else "Khách Về"]
    ws.append(hang_du_lieu)
    wb.save(file_name)

# ==============================================================================
# 4. CÁC MÀN HÌNH GIAO DIỆN STREAMLIT
# ==============================================================================

# ------------------------------------------------------------------------------
# MÀN HÌNH CHÍNH (HOME)
# ------------------------------------------------------------------------------
if st.session_state.page == "home":
    st.markdown("""
        <div class="main-title">
            🚘 VINFAST HƯNG THỊNH PHÁT<br>
            <span style="font-size: 14px; font-weight: normal; color: #00d26a;">HÀNH TRÌNH XANH - TƯƠNG LAI XANH</span>
        </div>
    """, unsafe_allow_html=True)

    # Hiển thị ảnh nền trang chủ nếu có
    anh_nen = os.path.join(THU_MUC_ANH, "anh_nen_tong_the.jpg")
    if os.path.exists(anh_nen):
        st.image(anh_nen, use_container_width=True)

    st.write(" ")
    
    if st.button("1. 🚗 KHÁCH ĐẾN SHOWROOM", use_container_width=True, type="primary"):
        set_page("khach_den")
        st.rerun()

    if st.button("2. 🚶 KHÁCH RỜI SHOWROOM", use_container_width=True):
        set_page("khach_ve")
        st.rerun()

    if st.button("3. 📋 TRA CỨU BẢNG GIÁ / THÔNG SỐ XE", use_container_width=True):
        set_page("tra_cuu")
        st.rerun()

# ------------------------------------------------------------------------------
# 1. KHÂU KHÁCH ĐẾN SHOWROOM
# ------------------------------------------------------------------------------
elif st.session_state.page == "khach_den":
    st.markdown("<div class="sub-title">📥 THÔNG TIN KHÁCH ĐẾN SHOWROOM</div>", unsafe_allow_html=True)

    co_hen = st.radio("Khách hàng đã có hẹn trước chưa?", ["Chưa có hẹn (Vãng lai)", "Đã có hẹn trước"], horizontal=True)
    is_hen = (co_hen == "Đã có hẹn trước")

    with st.form("form_khach_den", clear_on_submit=True):
        ma_nv = ""
        if is_hen:
            ma_nv = st.text_input("Mã nhân viên hẹn trước: *")
            
        ho_ten = st.text_input("Họ và tên khách hàng: *")
        xe_chon = st.selectbox("Dòng xe bạn quan tâm: *", cac_dong_xe)
        muc_dich = st.selectbox("Mục đích sử dụng xe: *", ["Chạy gia đình", "Chạy dịch vụ"])
        phan_khuc = st.text_input("Phân khúc xe yêu cầu: *", placeholder="Ví dụ: SUV hạng A, C...")
        mau_sac = st.selectbox("Màu sắc lựa chọn: *", cac_mau_xe)
        vay_ngan_hang = st.radio("Cần hỗ trợ vay ngân hàng?", ["Không", "Có"], horizontal=True)
        sdt = st.text_input("Số điện thoại liên hệ: *")

        submitted = st.form_submit_button("💾 HOÀN TẤT & LƯU THÔNG TIN", use_container_width=True, type="primary")

        if submitted:
            if (is_hen and not ma_nv.strip()) or not ho_ten.strip() or not phan_khuc.strip() or not sdt.strip():
                st.error("⚠️ Vui lòng nhập đầy đủ các thông tin bắt buộc (*)")
            else:
                data_row = [
                    datetime.now().strftime("%H:%M:%S"),
                    "Đã hẹn" if is_hen else "Vãng lai",
                    ma_nv.strip() if is_hen else "Khách vãng lai",
                    ho_ten.strip(),
                    f"{xe_chon} ({phan_khuc.strip()})",
                    muc_dich,
                    mau_sac,
                    vay_ngan_hang,
                    sdt.strip()
                ]
                ghi_excel("KHÁCH ĐẾN", data_row)
                st.success("✅ Lưu thông tin khách đến thành công!")

    # Xem so sánh thông số nhanh về xe được chọn
    with st.expander(f"🔍 Xem chi tiết thông số so sánh dòng xe {xe_chon}"):
        df_sub = df_vinfast[df_vinfast["Dòng xe"] == xe_chon]
        st.dataframe(df_sub[["Phiên bản", "Giá niêm yết (VND)", "Hệ thống dẫn động", "Loại trần xe", "Quãng đường tối đa"]], use_container_width=True)

    st.write("---")
    # Nút quay lại màn hình chính
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()

# ------------------------------------------------------------------------------
# 2. KHÂU KHÁCH RỜI SHOWROOM
# ------------------------------------------------------------------------------
elif st.session_state.page == "khach_ve":
    st.markdown("<div class="sub-title">📤 THÔNG TIN KHÁCH RỜI SHOWROOM</div>", unsafe_allow_html=True)

    da_coc = st.radio("Khách hàng đã đặt cọc xe chưa?", ["Chưa cọc (Chưa đặt)", "Đã đặt cọc xe"], horizontal=True)
    is_coc = (da_coc == "Đã đặt cọc xe")

    with st.form("form_khach_ve", clear_on_submit=True):
        ma_nv = st.text_input("Mã NV tư vấn: *")
        ho_ten = st.text_input("Họ tên khách hàng: *")
        sdt = st.text_input("Số điện thoại: *")
        cccd = st.text_input("Số CCCD: *")
        
        tien_coc = "-"
        if is_coc:
            tien_coc = st.text_input("Số tiền cọc (VNĐ): *")

        submitted = st.form_submit_button("💾 HOÀN TẤT & LƯU THÔNG TIN", use_container_width=True, type="primary")

        if submitted:
            if not ma_nv.strip() or not ho_ten.strip() or not sdt.strip() or not cccd.strip() or (is_coc and not tien_coc.strip()):
                st.error("⚠️ Vui lòng điền đầy đủ các thông tin bắt buộc (*)")
            else:
                data_row = [
                    datetime.now().strftime("%H:%M:%S"),
                    "Đã đặt cọc" if is_coc else "Chưa cọc",
                    ma_nv.strip(),
                    ho_ten.strip(),
                    sdt.strip(),
                    cccd.strip(),
                    tien_coc.strip()
                ]
                ghi_excel("KHÁCH VỀ", data_row)
                st.success("✅ Lưu thông tin khách rời showroom thành công!")

    st.write("---")
    # Nút quay lại màn hình chính
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()

# ------------------------------------------------------------------------------
# 3. TRA CỨU BẢNG GIÁ & THÔNG SỐ XE
# ------------------------------------------------------------------------------
elif st.session_state.page == "tra_cuu":
    st.markdown("<div class="sub-title">📊 TRA CỨU BẢNG GIÁ & THÔNG SỐ XE</div>", unsafe_allow_html=True)

    ds_loc = ["-- Tất cả các dòng xe --"] + cac_dong_xe
    dong_xe_chon = st.selectbox("Chọn dòng xe cần tra cứu:", ds_loc)

    if dong_xe_chon == "-- Tất cả các dòng xe --":
        df_display = df_vinfast
    else:
        df_display = df_vinfast[df_vinfast["Dòng xe"] == dong_xe_chon]

    # Bảng hiển thị thông số gọn trên mobile
    st.dataframe(
        df_display[["Dòng xe", "Phiên bản", "Giá niêm yết (VND)", "Quãng đường tối đa", "Hệ thống dẫn động"]],
        use_container_width=True,
        hide_index=True
    )

    st.write("---")
    st.subheader("🖼️ Hình ảnh & Tính năng chi tiết")

    xe_chi_tiet = st.selectbox("Chọn xe cụ thể để xem ảnh & tính năng:", df_display["Dòng xe"].unique())
    
    row_info = df_vinfast[df_vinfast["Dòng xe"] == xe_chi_tiet].iloc[0]

    # Hiển thị hình ảnh xe
    png_path = os.path.join(THU_MUC_ANH, f"{xe_chi_tiet}.png")
    jpg_path = os.path.join(THU_MUC_ANH, f"{xe_chi_tiet}.jpg")

    if os.path.exists(png_path):
        st.image(png_path, caption=f"Hình ảnh VinFast {xe_chi_tiet}", use_container_width=True)
    elif os.path.exists(jpg_path):
        st.image(jpg_path, caption=f"Hình ảnh VinFast {xe_chi_tiet}", use_container_width=True)
    else:
        st.info(f"ℹ️ Chưa có hình ảnh mẫu `{xe_chi_tiet}` trong thư mục `{THU_MUC_ANH}/`")

    # Thông tin chi tiết
    st.write(f"**🚗 Dòng xe:** {row_info['Dòng xe']} ({row_info['Phiên bản']})")
    st.write(f"**💰 Giá niêm yết:** {row_info['Giá niêm yết (VND)']:,} VNĐ")
    st.write(f"**⚡ Quãng đường tối đa:** {row_info['Quãng đường tối đa']}")
    st.write(f"**⚙️ Hệ thống dẫn động:** {row_info['Hệ thống dẫn động']}")
    st.write(f"**🛋️ Loại trần xe:** {row_info['Loại trần xe']}")
    st.write(f"**🌟 Tính năng nổi bật:** {row_info['Tính năng nổi bật']}")

    st.write("---")
    # Nút quay lại màn hình chính
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()