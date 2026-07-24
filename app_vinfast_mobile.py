import os
import uuid
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from supabase import create_client, Client

# ==============================================================================
# 1. CẤU HÌNH TRANG & MÀN HÌNH NỀN (NHẬN DIỆN THƯƠNG HIỆU VINFAST)
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
    /* Cấu hình nền ứng dụng với tông màu VinFast Blue (#005baa) và tối mờ */
    .stApp {{
        background: linear-gradient(rgba(10, 25, 47, 0.88), rgba(10, 25, 47, 0.88)), 
                    url('{URL_ANH_NEN}');
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
        color: #ffffff;
    }}
    
    /* Tiêu đề chính - Viền Chrome & Màu VinFast Blue */
    .main-title {{ 
        text-align: center; 
        color: #ffffff; 
        font-weight: bold; 
        font-size: 22px; 
        margin-bottom: 20px; 
        background: linear-gradient(135deg, #005baa, #003b73); 
        padding: 15px; 
        border-radius: 12px; 
        border: 1px solid #d1d5db; 
        box-shadow: 0px 4px 12px rgba(0, 91, 170, 0.4);
    }}
    .main-title a {{ color: #ffffff !important; text-decoration: none !important; }}
    .main-title a:hover {{ color: #00a651 !important; }}
    .sub-title {{ color: #00a651; font-weight: bold; font-size: 18px; text-align: center; margin-bottom: 15px; text-transform: uppercase; }}
    
    /* Cấu hình Nút bấm (Buttons) - Chuẩn VinFast Blue & Viền Chrome */
    div.stButton > button {{
        width: 100% !important;
        height: 52px !important;
        font-weight: bold !important;
        font-size: 15px !important;
        border-radius: 10px !important;
        margin-bottom: 10px !important;
        background: linear-gradient(135deg, #005baa, #004080) !important;
        color: #ffffff !important;
        border: 1px solid #d1d5db !important;
        box-shadow: 0px 3px 8px rgba(0, 91, 170, 0.3) !important;
        transition: all 0.25s ease-in-out !important;
    }}
    div.stButton > button:hover, div.stButton > button:active {{
        background: #00a651 !important; 
        color: #ffffff !important; 
        border-color: #ffffff !important;
        box-shadow: 0px 4px 12px rgba(0, 166, 81, 0.5) !important;
    }}
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. KHỞI TẠO CSDL SUPABASE & TRUY XUẤT DỮ LIỆU
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

@st.cache_data(ttl=300)
def load_danh_muc_xe():
    """Tải danh mục xe động từ bảng 'danh_muc_xe' trên Supabase"""
    if not supabase_client:
        return pd.DataFrame()
    try:
        res = supabase_client.table("danh_muc_xe").select("*").order("id").execute()
        if res.data:
            df = pd.DataFrame(res.data)
            # Đổi tên cột khớp với giao diện hiển thị
            rename_map = {
                "phan_khuc": "Phân khúc",
                "dong_xe": "Dòng xe",
                "phien_ban": "Phiên bản",
                "gia_niem_yet": "Giá niêm yết (VND)",
                "he_thong_dan_dong": "Hệ thống dẫn động",
                "loai_tran_xe": "Loại trần xe",
                "quang_duong_toi_da": "Quãng đường tối đa",
                "tinh_nang_noi_bat": "Tính năng nổi bật"
            }
            df = df.rename(columns=rename_map)
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"⚠️ Lỗi tải bảng danh mục xe từ Supabase: {e}")
        return pd.DataFrame()

# Tải DataFrame danh mục xe động từ CSDL
df_vinfast = load_danh_muc_xe()

def luu_du_lieu_realtime(loai, hang_du_lieu):
    """Hàm ghi dữ liệu trực tiếp vào Supabase & thu thập IP / Thiết bị"""
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

def xac_thuc_dang_nhap(username, password):
    """Xác thực đăng nhập & Cấp Session ID mới"""
    if not supabase_client:
        return False, "⚠️ Không kết nối được CSDL!"

    try:
        res = supabase_client.table("tai_khoan").select("*").eq("ten_dang_nhap", username).eq("mat_khau", password).execute()

        if not res.data:
            return False, "❌ Sai tên đăng nhập hoặc mật khẩu!"

        user = res.data[0]

        if user.get("bi_khoa", False):
            return False, "🚫 Tài khoản của bạn đã bị Ban Giám Đốc khóa!"

        new_session_id = str(uuid.uuid4())
        supabase_client.table("tai_khoan").update({
            "session_id": new_session_id,
            "last_active": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }).eq("ten_dang_nhap", username).execute()

        user["session_id"] = new_session_id
        return True, user
    except Exception as e:
        return False, f"⚠️ Lỗi đăng nhập: {e}"

def check_single_device_session():
    """Kiểm tra liên tục phiên đăng nhập đơn thiết bị"""
    if st.session_state.get("logged_in") and st.session_state.get("user_info"):
        user_info = st.session_state.user_info
        try:
            res = supabase_client.table("tai_khoan").select("session_id, bi_khoa").eq("ten_dang_nhap", user_info["ten_dang_nhap"]).execute()
            if res.data:
                current_db_session = res.data[0].get("session_id")
                is_locked = res.data[0].get("bi_khoa", False)

                if is_locked:
                    st.session_state.logged_in = False
                    st.session_state.user_info = None
                    st.error("🚫 Tài khoản của bạn đã bị khóa bởi Ban Giám Đốc!")
                    st.rerun()

                if current_db_session and current_db_session != user_info.get("session_id"):
                    st.session_state.logged_in = False
                    st.session_state.user_info = None
                    st.warning("⚠️ Tài khoản đã được đăng nhập ở một thiết bị khác! Bạn đã bị tự động đăng xuất.")
                    st.rerun()
        except Exception:
            pass

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
    st.session_state.logged_in = False
    st.session_state.page = "login"
    st.rerun()

def xoa_thong_bao_het_han():
    """Tự động xóa các thông báo chung đã tạo quá 7 ngày"""
    if supabase_client:
        try:
            tu_ngay = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
            supabase_client.table("thong_bao_chung").delete().lt("thoi_gian", tu_ngay).execute()
        except Exception:
            pass

# ==============================================================================
# 3. KHỞI TẠO SESSION & DỮ LIỆU DÒNG XE / HÌNH ẢNH
# ==============================================================================
if "user_info" not in st.session_state:
    st.session_state.user_info = None

if "page" not in st.session_state:
    st.session_state.page = "login"

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

# Tự động dọn dẹp thông báo cũ quá 7 ngày
xoa_thong_bao_het_han()

# Lấy danh sách dòng xe động từ bảng danh_muc_xe
if not df_vinfast.empty and "Dòng xe" in df_vinfast.columns:
    cac_dong_xe = sorted(list(df_vinfast["Dòng xe"].unique()))
else:
    cac_dong_xe = ["VF 3", "VF 5 Plus", "VF 6", "VF 7", "VF 8", "VF 9"]

cac_mau_xe = ["Trắng", "Đen", "Xám", "Bạc", "Xanh VinFast", "Đỏ"]

def hien_thi_thong_tin_so_sanh(dong_xe):
    """Hàm hiển thị thông tin các phiên bản xe lấy từ Supabase"""
    if df_vinfast.empty:
        st.warning("Đang tải dữ liệu danh mục xe...")
        return

    df_sub = df_vinfast[df_vinfast["Dòng xe"] == dong_xe]
    
    for idx, row in df_sub.iterrows():
        st.markdown(f"#### 📌 Phiên bản: **{row.get('Phiên bản', 'Tiêu chuẩn')}**")
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"🧱 **Loại trần:** {row.get('Loại trần xe', 'N/A')}")
            st.success(f"🔋 **Quãng đường:** {row.get('Quãng đường tối đa', 'N/A')}")
        with col2:
            st.warning(f"⚙️ **Dẫn động:** {row.get('Hệ thống dẫn động', 'N/A')}")
            gia_val = row.get("Giá niêm yết (VND)", 0)
            try:
                gia_fmt = f"{float(gia_val):,.0f} VNĐ"
            except Exception:
                gia_fmt = f"{gia_val} VNĐ"
            st.error(f"💰 **Giá niêm yết:** {gia_fmt}")
            
        st.markdown(f"✨ **Tính năng nổi bật:** {row.get('Tính năng nổi bật', 'N/A')}")
        st.markdown("---")

check_single_device_session()

# ==============================================================================
# 4. KHAI BÁO THÔNG TIN LẦN ĐẦU CHO NHÂN VIÊN
# ==============================================================================
if st.session_state.get("logged_in") and st.session_state.get("user_info"):
    user = st.session_state.user_info
    chuc_vu_val = user.get("chuc_vu", "")
    chuc_vu_text = "Giám Đốc Showroom" if chuc_vu_val == "giam_doc" else (chuc_vu_val if chuc_vu_val else "Nhân Viên Showroom")
    ho_ten_nhap = user.get("ho_ten", "")
    chuc_vu_nhap = chuc_vu_val

    if user.get("lan_dang_nhap_dau", False) and user.get("chuc_vu") != "giam_doc":
        st.warning("⚠️ Dành cho lần đăng nhập đầu tiên: Vui lòng khai báo Chức vụ / Bộ phận của bạn.")
        chuc_vu_nhap = st.selectbox("Chọn Chức vụ / Bộ phận:", ["Cố vấn dịch vụ", "Kỹ thuật viên", "Tư vấn bán hàng", "Lễ tân", "Bảo vệ", "Tài chính / Kế toán", "Khác"])
        ho_ten_nhap = st.text_input("Họ và tên nhân viên:", value=user.get("ho_ten", ""))

        if st.button("💾 CẬP NHẬT THÔNG TIN", type="primary", use_container_width=True):
            if not ho_ten_nhap.strip():
                st.error("⚠️ Vui lòng nhập họ và tên!")
            else:
                try:
                    supabase_client.table("tai_khoan").update({
                        "ho_ten": ho_ten_nhap.strip(),
                        "chuc_vu": chuc_vu_nhap,
                        "lan_dang_nhap_dau": False
                    }).eq("ten_dang_nhap", user["ten_dang_nhap"]).execute()
                    
                    st.session_state.user_info["ho_ten"] = ho_ten_nhap.strip()
                    st.session_state.user_info["chuc_vu"] = chuc_vu_nhap
                    st.session_state.user_info["lan_dang_nhap_dau"] = False
                    st.success("✅ Cập nhật thông tin thành công!")
                    st.rerun()
                except Exception as e:
                    st.error(f"⚠️ Lỗi cập nhật: {e}")

# ==============================================================================
# 5. ĐIỀU HƯỚNG MÀN HÌNH (ROUTING)
# ==============================================================================

# ------------------------------------------------------------------------------
# MÀN HÌNH ĐĂNG NHẬP (LOGIN)
# ------------------------------------------------------------------------------
if not st.session_state.get("logged_in") or st.session_state.page == "login":
    st.markdown("""
        <div class="main-title">
            🚘 VINFAST HƯNG THỊNH PHÁT<br>
            <span style="font-size: 14px; color: #00a651;">ĐĂNG NHẬP HỆ THỐNG</span>
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
                    st.session_state.logged_in = True
                    st.session_state.user_info = result
                    st.session_state.page = "home"
                    st.success(f"Xin chào {result['ho_ten']}!")
                    st.rerun()
                else:
                    st.error(result)

# ------------------------------------------------------------------------------
# MÀN HÌNH CHÍNH (HOME)
# ------------------------------------------------------------------------------
elif st.session_state.page == "home":
    user = st.session_state.user_info
    chuc_vu_val = user.get("chuc_vu", "")
    chuc_vu_text = "Giám Đốc Showroom" if chuc_vu_val == "giam_doc" else (chuc_vu_val if chuc_vu_val else "Nhân Viên Showroom")

    st.markdown(f"""
        <div class="main-title">
            🚘 <a href="https://vinfasthungthinhphat.vn" target="_blank">VINFAST HƯNG THỊNH PHÁT</a><br>
            <span style="font-size: 14px; color: #ffffff;">Xin chào: <strong>{user.get('ho_ten', '')}</strong> ({chuc_vu_text})</span>
        </div>
    """, unsafe_allow_html=True)

    # 1. HIỂN THỊ THÔNG BÁO TỪ BAN GIÁM ĐỐC
    if supabase_client:
        try:
            res_tb = supabase_client.table("thong_bao_chung").select("*").order("id", desc=True).limit(5).execute()
            ds_tb = res_tb.data or []
            if ds_tb:
                st.markdown("### 📢 BAN GIÁM ĐỐC THÔNG BÁO")
                for tb in ds_tb:
                    st.info(f"📌 **{tb['tieu_de']}** ({tb.get('thoi_gian', '')[:16]})\n\n{tb['noi_dung']}")
        except Exception:
            pass

    # 2. HIỂN THỊ PHẢN HỒI CỦA BGĐ CHO Ý KIẾN CỦA TÀI KHOẢN NÀY
    if supabase_client and user.get("ten_dang_nhap"):
        try:
            res_yk = supabase_client.table("y_kien_dong_gop") \
                .select("*") \
                .eq("nguoi_gui", user.get("ho_ten", "")) \
                .not_.is_("phan_hoi", "null") \
                .order("id", desc=True).execute()
            ds_phan_hoi = res_yk.data or []
            if ds_phan_hoi:
                st.markdown("### 💬 PHẢN HỒI TỪ BAN GIÁM ĐỐC")
                for yk in ds_phan_hoi:
                    with st.expander(f"📩 Ý kiến: '{yk['noi_dung'][:30]}...' - BGĐ đã trả lời"):
                        st.write(f"**Nội dung gửi:** {yk['noi_dung']}")
                        st.success(f"💬 **Trả lời từ BGĐ ({yk.get('thoi_gian_phan_hoi', '')[:16]}):**\n\n{yk['phan_hoi']}")
        except Exception:
            pass

    st.write("---")

    # DANH SÁCH NÚT CHỨC NĂNG
    if st.button("1. 🚗 KHÁCH ĐẾN SHOWROOM", use_container_width=True, type="primary"):
        set_page("khach_den")
        st.rerun()

    if st.button("2. 🚶 KHÁCH RỜI SHOWROOM", use_container_width=True):
        set_page("khach_ve")
        st.rerun()

    if st.button("3. 📋 TRA CỨU BẢNG GIÁ / THÔNG SỐ XE", use_container_width=True):
        set_page("tra_cuu")
        st.rerun()

    if st.button("💡 ĐỀ XUẤT CỦA NHÂN VIÊN", use_container_width=True):
        set_page("gui_y_kien")
        st.rerun()

    # PHÂN QUYỀN BAN GIÁM ĐỐC
    if user.get("chuc_vu") == "giam_doc":
        if st.button("4. 📊 BÁO CÁO THỐNG KÊ THEO NGÀY", key="btn_home_bao_cao", use_container_width=True):
            set_page("bao_cao")
            st.rerun()

        if st.button("5. 📢 BAN GIÁM ĐỐC THÔNG BÁO", key="btn_home_bgd_thong_bao", use_container_width=True, type="primary"):
            set_page("bgd_thong_bao")
            st.rerun()

        if st.button("6. 💬 PHẢN HỒI Ý KIẾN KHÁCH HÀNG & NHÂN VIÊN", key="btn_home_quan_ly_yk", use_container_width=True):
            set_page("quan_ly_y_kien")
            st.rerun()

    st.write("---")
    if st.button("🚪 ĐĂNG XUẤT", key="btn_logout_home", use_container_width=True):
        dang_xuat()

# ------------------------------------------------------------------------------
# KHÂU KHÁCH ĐẾN SHOWROOM
# ------------------------------------------------------------------------------
elif st.session_state.page == "khach_den":
    st.markdown('<div class="sub-title">📥 THÔNG TIN KHÁCH ĐẾN SHOWROOM</div>', unsafe_allow_html=True)

    co_hen = st.radio("Khách hàng đã có hẹn trước chưa?", ["Chưa có hẹn (Vãng lai)", "Đã có hẹn trước"], horizontal=True)
    is_hen = (co_hen == "Đã có hẹn trước")

    xe_chon = st.selectbox("Dòng xe bạn quan tâm: *", cac_dong_xe)

    hien_thi_anh_xe(xe_chon)

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

# ------------------------------------------------------------------------------
# KHÂU KHÁCH RỜI SHOWROOM
# ------------------------------------------------------------------------------
elif st.session_state.page == "khach_ve":
    st.markdown('<div class="sub-title">📤 THÔNG TIN KHÁCH RỜI SHOWROOM</div>', unsafe_allow_html=True)

    da_coc = st.radio("Khách hàng đã đặt cọc xe chưa?", ["Chưa cọc (Chưa đặt)", "Đã đặt cọc xe"], horizontal=True)
    is_coc = (da_coc == "Đã đặt cọc xe")

    ds_xe_da_xem = st.multiselect("Dòng xe khách đã xem: *", cac_dong_xe, default=[cac_dong_xe[0]] if cac_dong_xe else [])

    if ds_xe_da_xem:
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

# ------------------------------------------------------------------------------
# MÀN HÌNH TRA CỨU BẢNG GIÁ / THÔNG SỐ XE
# ------------------------------------------------------------------------------
elif st.session_state.page == "tra_cuu":
    st.markdown('<div class="sub-title">📋 TRA CỨU BẢNG GIÁ & THÔNG SỐ CÁC DÒNG XE</div>', unsafe_allow_html=True)
    
    xe_xem_anh = st.selectbox("Chọn dòng xe để tra cứu chi tiết:", cac_dong_xe)
    hien_thi_anh_xe(xe_xem_anh)
    
    st.markdown(f"### ⚡ Thông tin chi tiết xe VinFast {xe_xem_anh}")
    hien_thi_thong_tin_so_sanh(xe_xem_anh)
    
    st.write("---")
    st.markdown("**📊 Bảng tổng hợp các dòng xe từ CSDL Supabase:**")
    if not df_vinfast.empty:
        df_display = df_vinfast.copy()
        if "Giá niêm yết (VND)" in df_display.columns:
            df_display["Giá niêm yết (VND)"] = df_display["Giá niêm yết (VND)"].apply(
                lambda x: f"{float(x):,.0f} VNĐ" if pd.notnull(x) and str(x).replace('.','',1).isdigit() else f"{x} VNĐ"
            )
        st.dataframe(df_display, use_container_width=True, hide_index=True)
    else:
        st.info("Chưa có dữ liệu danh mục xe.")

    st.write("---")
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()

# ------------------------------------------------------------------------------
# MÀN HÌNH BÁO CÁO THỐNG KÊ
# ------------------------------------------------------------------------------
elif st.session_state.page == "bao_cao":
    st.markdown('<div class="sub-title">📊 BÁO CÁO THỐNG KÊ THEO NGÀY</div>', unsafe_allow_html=True)

    ngay_chon = st.date_input("Chọn ngày xem báo cáo:", datetime.now())
    str_ngay = ngay_chon.strftime("%Y-%m-%d")

    if st.button("🔍 TRUY XUẤT DỮ LIỆU", type="primary", use_container_width=True):
        if supabase_client:
            try:
                res_den = supabase_client.table("khach_den") \
                    .select("*") \
                    .gte("thoi_gian", f"{str_ngay} 00:00:00") \
                    .lte("thoi_gian", f"{str_ngay} 23:59:59") \
                    .execute()
                data_den = res_den.data

                res_ve = supabase_client.table("khach_ve") \
                    .select("*") \
                    .gte("thoi_gian", f"{str_ngay} 00:00:00") \
                    .lte("thoi_gian", f"{str_ngay} 23:59:59") \
                    .execute()
                data_ve = res_ve.data

                col1, col2, col3 = st.columns(3)
                col1.metric("📥 Tổng Khách Đến", f"{len(data_den)} lượt")
                col2.metric("📤 Tổng Khách Về", f"{len(data_ve)} lượt")
                
                so_luong_coc = sum(1 for item in data_ve if item.get("trang_thai_coc") == "Đã đặt cọc")
                col3.metric("💰 Số Lượng Cọc", f"{so_luong_coc} hợp đồng")

                st.write("---")

                ds_xe = []
                for item in data_den:
                    if item.get("dong_xe"):
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

# ------------------------------------------------------------------------------
# MÀN HÌNH GỬI Ý KIẾN ĐÓNG GÓP
# ------------------------------------------------------------------------------
elif st.session_state.page == "gui_y_kien":
    st.markdown('<div class="sub-title">💬 GỬI Ý KIẾN ĐÓNG GÓP</div>', unsafe_allow_html=True)
    st.info("Ý kiến của bạn sẽ được gửi trực tiếp và bảo mật tới Ban Giám Đốc.")

    loai_y_kien = st.radio("Bạn là:", ["Khách hàng", "Nhân viên"], horizontal=True)
    
    with st.form("form_y_kien", clear_on_submit=True):
        if loai_y_kien == "Khách hàng":
            ten_nguoi_gui = st.text_input("Họ tên / Số điện thoại (Không bắt buộc):", value="Khách hàng")
        else:
            user_nv = st.session_state.user_info.get("ho_ten", "Nhân viên") if st.session_state.user_info else "Nhân viên"
            ten_nguoi_gui = st.text_input("Tên / Mã nhân viên:", value=user_nv)

        noi_dung = st.text_area("Nội dung đóng góp ý kiến / Góp ý dịch vụ: *", height=120)
        btn_gui = st.form_submit_button("🚀 GỬI Ý KIẾN TỚI BGĐ", use_container_width=True, type="primary")

        if btn_gui:
            if not noi_dung.strip():
                st.error("⚠️ Vui lòng nhập nội dung ý kiến!")
            else:
                try:
                    data_yk = {
                        "thoi_gian": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "nguoi_gui": ten_nguoi_gui.strip(),
                        "loai_y_kien": loai_y_kien,
                        "noi_dung": noi_dung.strip()
                    }
                    supabase_client.table("y_kien_dong_gop").insert(data_yk).execute()
                    st.success("✅ Cảm ơn bạn! Ý kiến đóng góp đã được gửi thành công tới Ban Giám Đốc.")
                except Exception as e:
                    st.error(f"⚠️ Lỗi khi gửi ý kiến: {e}")

    st.write("---")
    if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
        set_page("home")
        st.rerun()

# ------------------------------------------------------------------------------
# MÀN HÌNH QUẢN LÝ & PHẢN HỒI Ý KIẾN (BAN GIÁM ĐỐC)
# ------------------------------------------------------------------------------
elif st.session_state.page == "quan_ly_y_kien":
    if st.session_state.user_info.get("chuc_vu") != "giam_doc":
        st.error("🚫 Quyền truy cập bị từ chối! Trang này chỉ dành cho Ban Giám Đốc.")
        if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH"):
            set_page("home")
            st.rerun()
    else:
        st.markdown('<div class="sub-title">📬 PHẢN HỒI BAN GIÁM ĐỐC TỚI KHÁCH HÀNG & NHÂN VIÊN</div>', unsafe_allow_html=True)

        if supabase_client:
            try:
                res = supabase_client.table("y_kien_dong_gop").select("*").order("id", desc=True).execute()
                ds_y_kien = res.data or []

                if not ds_y_kien:
                    st.info("Hiện chưa có ý kiến đóng góp nào.")
                else:
                    for item in ds_y_kien:
                        stt_phan_hoi = "✅ Đã trả lời" if item.get('phan_hoi') else "🔴 Chờ BGĐ trả lời"
                        
                        with st.expander(f"📩 [{item.get('loai_y_kien', 'Ý kiến')}] {item.get('nguoi_gui', 'Ẩn danh')} - {item.get('thoi_gian', '')[:16]} ({stt_phan_hoi})", expanded=(not item.get('phan_hoi'))):
                            st.write(f"**Nội dung đóng góp:**")
                            st.warning(item.get('noi_dung', ''))

                            if item.get('phan_hoi'):
                                st.success(f"💬 **Ban Giám Đốc đã trả lời ({item.get('thoi_gian_phan_hoi', '')[:16]}):**\n\n{item['phan_hoi']}")
                            
                            st.write("✍️ **Nhập / Cập nhật nội dung trả lời:**")
                            phan_hoi_text = st.text_area("Nội dung trả lời:", value=item.get('phan_hoi', '') or '', key=f"reply_{item['id']}")
                            
                            if st.button("💾 Gửi phản hồi", key=f"btn_reply_{item['id']}", type="primary"):
                                if phan_hoi_text.strip():
                                    supabase_client.table("y_kien_dong_gop").update({
                                        "phan_hoi": phan_hoi_text.strip(),
                                        "thoi_gian_phan_hoi": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    }).eq("id", item['id']).execute()
                                    st.success("✅ Đã lưu và gửi phản hồi thành công!")
                                    st.rerun()
                                else:
                                    st.error("Vui lòng nhập nội dung trả lời trước khi gửi.")
            except Exception as e:
                st.error(f"⚠️ Lỗi tải danh sách ý kiến: {e}")

        st.write("---")
        if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
            set_page("home")
            st.rerun()

# ------------------------------------------------------------------------------
# MÀN HÌNH BAN GIÁM ĐỐC THÔNG BÁO
# ------------------------------------------------------------------------------
elif st.session_state.page == "bgd_thong_bao":
    if st.session_state.user_info.get("chuc_vu") != "giam_doc":
        st.error("🚫 Quyền truy cập bị từ chối! Tính năng này chỉ dành riêng cho Ban Giám Đốc.")
        if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH"):
            set_page("home")
            st.rerun()
    else:
        st.markdown('<div class="sub-title">📢 BAN GIÁM ĐỐC THÔNG BÁO TỚI TOÀN BỘ NHÂN VIÊN</div>', unsafe_allow_html=True)

        with st.form("form_bgd_thong_bao", clear_on_submit=True):
            tieu_de_tb = st.text_input("📌 Tiêu đề thông báo: *")
            noi_dung_tb = st.text_area("📝 Nội dung thông báo: *", height=120)
            btn_phat_tb = st.form_submit_button("🚀 ĐĂNG THÔNG BÁO (TỰ ĐỘNG XÓA SAU 7 NGÀY)", type="primary", use_container_width=True)

            if btn_phat_tb:
                if not tieu_de_tb.strip() or not noi_dung_tb.strip():
                    st.error("⚠️ Vui lòng nhập đầy đủ Tiêu đề và Nội dung!")
                else:
                    try:
                        supabase_client.table("thong_bao_chung").insert({
                            "nguoi_gui": st.session_state.user_info.get("ho_ten", "Ban Giám Đốc"),
                            "tieu_de": tieu_de_tb.strip(),
                            "noi_dung": noi_dung_tb.strip(),
                            "thoi_gian": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }).execute()
                        st.success("✅ Đã đăng thông báo thành công tới toàn bộ nhân viên!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"⚠️ Lỗi khi đăng thông báo: {e}")

        st.write("---")
        st.subheader("📋 Các thông báo đang hiển thị (Tự động lưu trong 7 ngày):")
        if supabase_client:
            try:
                res_tb = supabase_client.table("thong_bao_chung").select("*").order("id", desc=True).execute()
                ds_tb = res_tb.data or []
                if not ds_tb:
                    st.info("Hiện không có thông báo nào.")
                else:
                    for tb in ds_tb:
                        with st.expander(f"📌 {tb['tieu_de']} ({tb.get('thoi_gian', '')[:16]})"):
                            st.write(tb['noi_dung'])
                            if st.button("🗑️ Xóa thông báo này", key=f"del_tb_{tb['id']}"):
                                supabase_client.table("thong_bao_chung").delete().eq("id", tb['id']).execute()
                                st.success("Đã xóa thông báo!")
                                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Lỗi tải thông báo: {e}")

        st.write("---")
        if st.button("🏠 QUAY LẠI MÀN HÌNH CHÍNH", use_container_width=True):
            set_page("home")
            st.rerun()
