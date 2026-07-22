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
        background-color: #121212;
        color: #ffffff;
    }
    .main-title {
@@ -27,9 +27,9 @@
        font-weight: bold;
        font-size: 22px;
        margin-bottom: 20px;
        background-color: #000000;
        background-color: #1e1e1e;
        padding: 15px;
        border-radius: 10px;
        border-radius: 12px;
        border: 1px solid #333333;
    }
    .sub-title {
@@ -39,17 +39,33 @@
        text-align: center;
        margin-bottom: 15px;
    }
    
    /* ========================================================================== */
    /* CẤU HÌNH NÚT BẤM NỔI BẬT DỄ NHÌN TRÊN MÀN HÌNH ĐIỆN THOẠI                  */
    /* ========================================================================== */
    div.stButton > button {
        width: 100%;
        height: 50px;
        font-weight: bold;
        font-size: 16px;
        border-radius: 8px;
        margin-bottom: 10px;
        width: 100% !important;
        height: 55px !important;
        font-weight: bold !important;
        font-size: 16px !important;
        border-radius: 12px !important;
        margin-bottom: 12px !important;
        background: linear-gradient(135deg, #1e293b, #0f172a) !important; /* Nền dải màu hiện đại */
        color: #ffffff !important; /* Mẫu chữ màu trắng sáng */
        border: 2px solid #00d26a !important; /* Viền xanh dạ quang siêu nổi */
        box-shadow: 0px 4px 10px rgba(0, 210, 106, 0.2) !important; /* Hiệu ứng phát sáng nhẹ */
        transition: all 0.2s ease-in-out !important;
    }

    /* Hiệu ứng khi nhấn chọn hoặc di chuột vào nút */
    div.stButton > button:hover, div.stButton > button:active, div.stButton > button:focus {
        background: #00d26a !important;
        color: #000000 !important;
        border-color: #00d26a !important;
        box-shadow: 0px 6px 15px rgba(0, 210, 106, 0.5) !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. KHỞI TẠO DỮ LIỆU XE VINFAST
# ==============================================================================
@@ -148,181 +164,153 @@
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
    st.markdown('<div class="sub-title">📥 THÔNG TIN KHÁCH ĐẾN SHOWROOM</div>', unsafe_allow_html=True)

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
    st.markdown('<div class="sub-title">📤 THÔNG TIN KHÁCH RỜI SHOWROOM</div>', unsafe_allow_html=True)

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
    st.markdown('<div class="sub-title">📊 TRA CỨU BẢNG GIÁ & THÔNG SỐ XE</div>', unsafe_allow_html=True)

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
