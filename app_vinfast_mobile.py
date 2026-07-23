import os
import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime
from supabase import create_client, Client

# Tích hợp Google Sheets Realtime
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False
# ==============================================================================
# MÀN HÌNH ĐĂNG NHẬP (LOGIN)
# ==============================================================================
if st.session_state.page == "login" or not st.session_state.user_info:
    st.markdown("""
        <div class="main-title">
            🚘 VINFAST HƯNG THỊNH PHÁT<br>
            <span style="font-size: 14px; color: #00d26a;">ĐĂNG NHẬP HỆ THỐNG</span>
        </div>
    """, unsafe_allow_html=True)

    with st.form("form_login"):
        username = st.text_input("Tên đăng nhập:")
        password = st.text_input("Mật khẩu:", type="password")
        btn_login = st.form_submit_button("🔑 ĐĂNG NHẬP", use_container_width=True, type="primary")

        if btn_login:
            if not username.strip() or not password.strip():
                st.error("⚠️ Vui lòng nhập đầy đủ thông tin!")
            else:
                success, result = xac_thuc_dang_nhap(username.strip(), password.strip())
                if success:
                    st.session_state.user_info = result
                    st.session_state.page = "home"
                    st.success(f"Xin chào {result['ho_ten']}!")
                    st.rerun()
                else:
                    st.error(result)

# ...Tạo Màn Hình Đăng Nhập & Cập Nhật Điều Hướng Trang Chính==============================================================================
# MÀN HÌNH CHÍNH (HOME)
# ==============================================================================
elif st.session_state.page == "home":
    user = st.session_state.user_info
    chuc_vu_text = "Giám Đốc Showroom" if user["chuc_vu"] == "giam_doc" else "Nhân Viên Showroom"

    st.markdown(f"""
        <div class="main-title">
            🚘 VINFAST HƯNG THỊNH PHÁT<br>
            <span style="font-size: 14px; color: #00d26a;">Xin chào: {user['ho_ten']} ({chuc_vu_text})</span>
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

    # PHÂN QUYỀN: Chỉ Giám đốc mới thấy mục Báo cáo
    if user["chuc_vu"] == "giam_doc":
        if st.button("4. 📊 BÁO CÁO THỐNG KÊ THEO NGÀY", use_container_width=True):
            set_page("bao_cao")
            st.rerun()

    st.write("---")
    if st.button("🚪 ĐĂNG XUẤT", use_container_width=True):
        dang_xuat()


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
if "user_info" not in st.session_state:
    st.session_state.user_info = None  # Lưu thông tin người dùng sau khi đăng nhập

if "page" not in st.session_state:
    st.session_state.page = "login"   # Mặc định chưa đăng nhập sẽ ở trang 'login'
    
# ....2 hàm trên là Cập nhật Mã nguồn

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
    "VF 9": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%209.jpg",
    "Minio Green": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/Minio%20Green.jpg",
    "Herio Green": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/Herio%20Green.jpg",
    "Limo Green": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/Limo%20Green.jpg",
    "VF MPV 7": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/VF%20MPV%207.jpg",
    "EC Van": "https://raw.githubusercontent.com/ducbuihuu315/vfhungthinh/main/hinhanh/EC%20Van.jpg"
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

# Dữ liệu đầy đủ 19 phiên bản dòng xe VinFast
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
        269000000, 499000000, 749000000, 819000000, 174600000, 285000000
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
cac_mau_xe = ["Trắng", "Đen", "Xám", "Bạc", "Xanh", "Đỏ"]

def hien_thi_thong_tin_so_sanh(dong_xe):
    """Hàm hiển thị danh sách các phiên bản & tính năng nổi bật của xe được chọn"""
    df_sub = df_vinfast[df_vinfast["Dòng xe"] == dong_xe]
    
    for idx, row in df_sub.iterrows():
        st.markdown(f"#### 📌 Phiên bản: **{row['Phiên bản']}**")
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"🧱 **Loại trần:** {row['Loại trần xe']}")
            st.success(f"🔋 **Quãng đường:** {row['Quãng đường tối đa']}")
        with col2:
            st.warning(f"⚙️ **Dẫn động:** {row['Hệ thống dẫn động']}")
            st.error(f"💰 **Giá niêm yết:** {row['Giá niêm yết (VND)']:,.0f} VNĐ")
            
        st.markdown(f"✨ **Tính năng nổi bật:** {row['Tính năng nổi bật']}")
        st.markdown("---")

# ==============================================================================
# 3. HÀM KẾT NỐI VÀ LƯU DỮ LIỆU CSDL SUPABASE
# ==============================================================================
@st.cache_resource
def init_supabase() -> Client:
    """Khởi tạo kết nối với CSDL Supabase"""
    try:
        url = st.secrets["supabase"]["SUPABASE_URL"]
        key = st.secrets["supabase"]["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception as e:
        st.error(f"⚠️ Lỗi cấu hình Supabase Secrets: {e}")
        return None

supabase_client = init_supabase()

def luu_du_lieu_realtime(loai, hang_du_lieu):
    """Hàm ghi dữ liệu trực tiếp vào Supabase & thu thập IP / Thiết bị"""
    # 1. Lấy thông tin IP và Trình duyệt/Thiết bị
    headers = st.context.headers
    user_ip = headers.get("X-Forwarded-For", "N/A").split(',')[0].strip()
    user_agent = headers.get("User-Agent", "N/A")

    if not supabase_client:
        st.error("⚠️ Không thể kết nối tới CSDL Supabase!")
        return

    try:
        if loai == "KHÁCH ĐẾN":
            data_to_insert = {
                "thoi_gian": hang_du_lieu[0],
                "loai_khach": hang_du_lieu[1],
                "ma_nv": hang_du_lieu[2],
                "ho_ten_kh": hang_du_lieu[3],
                "dong_xe": hang_du_lieu[4],
                "muc_dich": hang_du_lieu[5],
                "mau_sac": hang_du_lieu[6],
                "nhu_cau_vay": hang_du_lieu[7],
                "sdt": hang_du_lieu[8],
                "ip_may": user_ip,
                "user_agent": user_agent
            }
            supabase_client.table("khach_den").insert(data_to_insert).execute()
        else:
            data_to_insert = {
                "thoi_gian": hang_du_lieu[0],
                "trang_thai_coc": hang_du_lieu[1],
                "ma_nv": hang_du_lieu[2],
                "ho_ten": hang_du_lieu[3],
                "sdt": hang_du_lieu[4],
                "cccd": hang_du_lieu[5],
                "xe_da_xem": hang_du_lieu[6],
                "tien_coc": hang_du_lieu[7],
                "ip_may": user_ip,
                "user_agent": user_agent
            }
            supabase_client.table("khach_ve").insert(data_to_insert).execute()

    except Exception as e:
        st.error(f"⚠️ Lỗi khi đồng bộ dữ liệu: {e}")
        
# ==============================================================================
# HÀM XỬ LÝ ĐĂNG NHẬP & GIỚI HẠN THIẾT BỊ
# ==============================================================================
def xac_thuc_dang_nhap(username, password):
    """Hàm xác thực đăng nhập và khóa tài khoản theo thiết bị"""
    headers = st.context.headers
    # Lấy thông tin đặc xưng trình duyệt/máy làm ID thiết bị (device_id)
    device_id = headers.get("User-Agent", "N/A") + "_" + headers.get("X-Forwarded-For", "N/A").split(',')[0].strip()

    if not supabase_client:
        return False, "⚠️ Không kết nối được CSDL!"

    try:
        # Kiểm tra tên đăng nhập & mật khẩu
        res = supabase_client.table("tai_khoan") \
            .select("*") \
            .eq("ten_dang_nhap", username) \
            .eq("mat_khau", password) \
            .execute()

        if not res.data:
            return False, "❌ Sai tên đăng nhập hoặc mật khẩu!"

        user = res.data[0]
        current_device = user.get("device_id")

        # Kiểm tra giới hạn thiết bị
        if current_device and current_device != device_id:
            return False, "🚫 Tài khoản này đang được đăng nhập trên một thiết bị khác!"

        # Cập nhật device_id cho thiết bị hiện tại
        supabase_client.table("tai_khoan") \
            .update({"device_id": device_id}) \
            .eq("ten_dang_nhap", username) \
            .execute()

        return True, user
    except Exception as e:
        return False, f"⚠️ Lỗi xác thực: {e}"


def dang_xuat():
    """Đăng xuất và xóa đăng ký thiết bị"""
    if st.session_state.user_info and supabase_client:
        try:
            supabase_client.table("tai_khoan") \
                .update({"device_id": None}) \
                .eq("ten_dang_nhap", st.session_state.user_info["ten_dang_nhap"]) \
                .execute()
        except Exception:
            pass
    st.session_state.user_info = None
    st.session_state.page = "login"
    st.rerun()

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
    if st.button("4. 📊 BÁO CÁO THỐNG KÊ THEO NGÀY", use_container_width=True):
        set_page("bao_cao")
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

    # Hiển thị phân tích nổi bật & các phiên bản
    with st.expander(f"🔍 Xem Chi tiết các phiên bản & Tính năng của {xe_chon}", expanded=True):
        hien_thi_thong_tin_so_sanh(xe_chon)

    with st.form("form_khach_den", clear_on_submit=True):
        ma_nv = st.text_input("Mã nhân viên tư vấn/hẹn trước: *")
        ho_ten = st.text_input("Họ và tên khách hàng: *")
        muc_dich = st.selectbox("Mục đích sử dụng xe: *", ["Chạy gia đình", "Chạy dịch vụ", "Vận tải / Kinh doanh"])
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
                with st.expander(f"ℹ️ Các phiên bản {xe}"):
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
    
    st.markdown(f"### ⚡ Thông tin chi tiết xe VinFast {xe_xem_anh}")
    hien_thi_thong_tin_so_sanh(xe_xem_anh)
    
    st.write("---")
    st.markdown("**📊 Bảng tổng hợp tất cả 19 phiên bản các dòng xe:**")
    df_display = df_vinfast.copy()
    df_display["Giá niêm yết (VND)"] = df_display["Giá niêm yết (VND)"].apply(lambda x: f"{x:,.0f} VNĐ")
    st.dataframe(df_display, use_container_width=True, hide_index=True)
    
    st.write("---")
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()
# ==============================================================================
# 8. MÀN HÌNH BÁO CÁO THỐNG KÊ
# ==============================================================================
elif st.session_state.page == "bao_cao":
    st.markdown('<div class="sub-title">📊 BÁO CÁO THỐNG KÊ THEO NGÀY</div>', unsafe_allow_html=True)

    # Chọn ngày muốn xem báo cáo
    ngay_chon = st.date_input("Chọn ngày xem báo cáo:", datetime.now())
    str_ngay = ngay_chon.strftime("%Y-%m-%d")

    if st.button("🔍 TRUY XUẤT DỮ LIỆU", type="primary", use_container_width=True):
        if supabase_client:
            try:
                # 1. Truy vấn dữ liệu Khách Đến trong ngày
                res_den = supabase_client.table("khach_den") \
                    .select("*") \
                    .gte("thoi_gian", f"{str_ngay} 00:00:00") \
                    .lte("thoi_gian", f"{str_ngay} 23:59:59") \
                    .execute()
                data_den = res_den.data

                # 2. Truy vấn dữ liệu Khách Về trong ngày
                res_ve = supabase_client.table("khach_ve") \
                    .select("*") \
                    .gte("thoi_gian", f"{str_ngay} 00:00:00") \
                    .lte("thoi_gian", f"{str_ngay} 23:59:59") \
                    .execute()
                data_ve = res_ve.data

                # 3. Hiển thị các chỉ số tổng quan (KPIs)
                col1, col2, col3 = st.columns(3)
                col1.metric("📥 Tổng Khách Đến", f"{len(data_den)} lượt")
                col2.metric("📤 Tổng Khách Về", f"{len(data_ve)} lượt")
                
                # Tính số lượng cọc
                so_luong_coc = sum(1 for item in data_ve if item.get("trang_thai_coc") == "Đã đặt cọc")
                col3.metric("💰 Số Lượng Cọc", f"{so_luong_coc} hợp đồng")

                st.write("---")

                # 4. Tìm dòng xe được quan tâm nhiều nhất
                ds_xe = []
                for item in data_den:
                    if item.get("dong_xe"):
                        # Lấy tên dòng xe chính
                        xe_name = item["dong_xe"].split("(")[0].strip()
                        ds_xe.append(xe_name)
                
                for item in data_ve:
                    if item.get("xe_da_xem"):
                        for xe in item["xe_da_xem"].split(","):
                            ds_xe.append(xe.strip())

                if ds_xe:
                    df_xe = pd.Series(ds_xe).value_counts().reset_index()
                    df_xe.columns = ["Dòng xe", "Số lượt quan tâm"]
                    
                    st.markdown(f"🏆 **Dòng xe hot nhất ngày {str_ngay}:** `{df_xe.iloc[0]['Dòng xe']}` ({df_xe.iloc[0]['Số lượt quan tâm']} lượt)")
                    st.dataframe(df_xe, use_container_width=True, hide_index=True)
                else:
                    st.info("Chưa có thông tin xe được chọn trong ngày này.")

            except Exception as e:
                st.error(f"Lỗi khi tải báo cáo: {e}")
        else:
            st.error("Chưa kết nối CSDL Supabase.")

    st.write("---")
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()
