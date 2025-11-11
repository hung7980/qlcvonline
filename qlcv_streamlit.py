# qlcv_streamlit_v12.py
# -*- coding: utf-8 -*-

import os, io, re, unicodedata, time
from datetime import datetime, date
from typing import List, Tuple, Optional

import streamlit as st
import pandas as pd

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import gspread

# =========================
# ======== CONFIG =========
# =========================

APP_TITLE = "QUẢN LÝ CÔNG VIỆC"

DEN_HEADERS = [
    "Cơ quan ban hành","Số văn bản","Ngày ban hành","Trích yếu",
    "Thời hạn xử lý","Nơi nhận","Người xử lý",
    "Ngày báo cáo (;)","Báo cáo cho","Danh sách file (;)","Ngày nhập"
]
DI_HEADERS = [
    "Số văn bản đến (tham chiếu)","Số văn bản đi","Ngày ban hành",
    "Trích yếu văn bản đi","Nơi nhận","Ghi chú","Danh sách file (;)","Phòng/Ban","Ngày nhập"
]
TAP_HEADERS = [
    "Số văn bản","Trích yếu tập huấn","Ngày & giờ tập huấn","Địa điểm tập huấn",
    "Người dự tập huấn","Danh sách file (;)","Ngày nhập"
]

DATE_FMT = "%d/%m/%Y"
DT_FMT   = "%d/%m/%Y %H:%M"

# =========================
# ====== AUTH/CLIENT ======
# =========================

@st.cache_resource(show_spinner=False)
def get_google_clients():
    # Lấy JSON Service Account từ secrets
    sa_info = st.secrets["gcp_service_account"]
    if isinstance(sa_info, str):
        import json
        sa_info = json.loads(sa_info)

    creds = service_account.Credentials.from_service_account_info(
        sa_info,
        scopes=[
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    # gspread (Sheets)
    gc = gspread.authorize(creds)
    # Drive
    drive = build("drive", "v3", credentials=creds)
    return gc, drive

def get_ids():
    return (st.secrets["VB_DEN_FOLDER"],
            st.secrets["VB_DI_FOLDER"],
            st.secrets["TAP_FOLDER"],
            st.secrets["SHEET_DEN_ID"],
            st.secrets["SHEET_DI_ID"],
            st.secrets["SHEET_TAP_ID"])

# =========================
# ======= HELPERS =========
# =========================

def norm(s:str) -> str:
    s = unicodedata.normalize("NFD", str(s).lower().strip())
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+"," ",s)

def contains(hay, kw)->bool:
    if not kw: return True
    return norm(kw) in norm(hay or "")

def parse_dt(s)->Optional[datetime]:
    if isinstance(s, datetime): return s
    s = (s or "").strip()
    for fmt in (DT_FMT, DATE_FMT, "%Y-%m-%d", "%d/%m/%Y %H:%M:%S"):
        try: return datetime.strptime(s, fmt)
        except: pass
    try: return datetime.strptime(s.split()[0], DATE_FMT)
    except: return None

def mm_yyyy_match(dt: Optional[datetime], my: str)->bool:
    if not my: return True
    if not dt: return False
    m = str(dt.month).zfill(2)
    y = str(dt.year)
    if "/" in my:
        try:
            mo, yr = [x.strip() for x in my.split("/")]
            return mo.zfill(2)==m and yr==y
        except:
            return False
    return my.isdigit() and my==y

@st.cache_data(show_spinner=False)
def read_sheet(gc, sheet_id: str, headers: List[str]) -> pd.DataFrame:
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1
    vals = ws.get_all_values()
    if not vals:
        ws.append_row(headers)
        vals = [headers]
    # Chuẩn hóa header
    if vals[0] != headers:
        # Nếu khác, cố gắng map theo tên
        cur_cols = vals[0]
        # tạo df cũ
        df_old = pd.DataFrame(vals[1:], columns=cur_cols)
        # align
        df = pd.DataFrame(columns=headers)
        for h in headers:
            if h in df_old.columns:
                df[h] = df_old[h]
            else:
                df[h] = ""
        return df
    else:
        df = pd.DataFrame(vals[1:], columns=headers)
        return df

def write_sheet(gc, sheet_id: str, headers: List[str], df: pd.DataFrame):
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1
    data = [headers] + df.fillna("").astype(str).values.tolist()
    ws.clear()
    ws.update("A1", data)

def upload_to_drive(drive, folder_id: str, filename: str, content: bytes, mimetype: str = "application/octet-stream") -> Tuple[str, str]:
    file_metadata = {
        "name": filename,
        "parents": [folder_id]
    }
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=mimetype, resumable=False)
    f = drive.files().create(body=file_metadata, media_body=media, fields="id,webViewLink,webContentLink").execute()
    return f["id"], f.get("webViewLink", "")

def build_file_links(folder_id: str, files_semi: str) -> List[Tuple[str, str]]:
    # Chỉ tạo link xem theo tên file: hiển thị dạng tìm “file name in folder” trên Drive (không chắc 1-1)
    # Tốt nhất là lưu fileId theo từng file; ở đây đơn giản: hiển thị tên để click tải (nếu có id).
    links = []
    for raw in [x.strip() for x in (files_semi or "").split(";") if x.strip()]:
        # Không biết id => hiển thị tên thôi
        links.append((raw, ""))  # name, link (nếu có)
    return links

def add_files_and_return_list(drive, folder_id: str, up_files) -> str:
    names = []
    for f in up_files:
        name = f.name
        content = f.read()
        # đoán mimetype
        mt = "application/octet-stream"
        if name.lower().endswith(".pdf"): mt="application/pdf"
        elif name.lower().endswith(".docx"): mt="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif name.lower().endswith(".doc"): mt="application/msword"
        elif name.lower().endswith(".xlsx"): mt="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif name.lower().endswith(".xls"): mt="application/vnd.ms-excel"
        elif name.lower().endswith(".pptx"): mt="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        elif name.lower().endswith(".png"): mt="image/png"
        elif name.lower().endswith(".jpg") or name.lower().endswith(".jpeg"): mt="image/jpeg"
        _, _view = upload_to_drive(drive, folder_id, name, content, mt)
        names.append(name)
    return ";".join(names)

def now_str(fmt=DT_FMT): return datetime.now().strftime(fmt)

# =========================
# ======= LOGIN UI ========
# =========================

def do_login():
    st.title(APP_TITLE)
    st.caption("Đăng nhập để sử dụng hệ thống")

    u = st.text_input("Tài khoản", "")
    p = st.text_input("Mật khẩu", "", type="password")
    if st.button("Đăng nhập", use_container_width=True):
        users = st.secrets.get("users", {})
        if u in users and str(users[u]) == p:
            st.session_state["user"] = u
            st.rerun()
        else:
            st.error("Sai tài khoản hoặc mật khẩu.")

def topbar(gc, drive):
    st.sidebar.header("Điều hướng")
    page = st.sidebar.radio("Trang", [
        "Dashboard", "1. Văn bản đến", "2. Văn bản đi", "3. Tập huấn / Giấy mời", "Tìm kiếm", "Cấu hình"
    ])
    st.sidebar.divider()
    if st.sidebar.button("Đăng xuất"):
        st.session_state.clear(); st.rerun()
    return page

# =========================
# ======= PAGES ===========
# =========================

def page_dashboard(gc, drive):
    den_id, di_id, tap_id, SID_DEN, SID_DI, SID_TAP = get_ids()
    st.header("📊 Dashboard")

    df_den = read_sheet(gc, SID_DEN, DEN_HEADERS)
    df_di  = read_sheet(gc, SID_DI,  DI_HEADERS)
    df_tap = read_sheet(gc, SID_TAP, TAP_HEADERS)

    # Thống kê nhanh
    today = datetime.now().strftime(DATE_FMT)
    need_xl = 0
    for _, r in df_den.iterrows():
        han = parse_dt(r["Thời hạn xử lý"])
        if han and han.strftime(DATE_FMT) == today:
            need_xl += 1
    need_bc = 0
    for _, r in df_den.iterrows():
        mocs = str(r["Ngày báo cáo (;)"] or "").split(";")
        for m in mocs:
            dt = parse_dt(m.strip())
            if dt and dt.strftime(DATE_FMT) == today:
                need_bc += 1; break

    c1, c2, c3 = st.columns(3)
    c1.metric("Văn bản đến", len(df_den))
    c2.metric("Văn bản đi",  len(df_di))
    c3.metric("Tập huấn/Giấy mời", len(df_tap))
    d1, d2 = st.columns(2)
    d1.metric("Cần XỬ LÝ hôm nay", need_xl)
    d2.metric("Cần BÁO CÁO hôm nay", need_bc)

    st.divider()
    st.subheader("Danh sách mới nhất")
    tab1, tab2, tab3 = st.tabs(["Đến", "Đi", "Tập huấn"])
    with tab1:
        if len(df_den):
            st.dataframe(df_den.tail(50), use_container_width=True, height=320)
        else:
            st.info("Chưa có dữ liệu.")
    with tab2:
        if len(df_di):
            st.dataframe(df_di.tail(50), use_container_width=True, height=320)
        else:
            st.info("Chưa có dữ liệu.")
    with tab3:
        if len(df_tap):
            st.dataframe(df_tap.tail(50), use_container_width=True, height=320)
        else:
            st.info("Chưa có dữ liệu.")

def page_den(gc, drive):
    den_id, di_id, tap_id, SID_DEN, *_ = get_ids()
    st.header("1) Văn bản đến")

    df = read_sheet(gc, SID_DEN, DEN_HEADERS)

    with st.form("den_form", clear_on_submit=True):
        coquan = st.text_input("Cơ quan ban hành")
        sovb   = st.text_input("Số văn bản *")
        ngay   = st.date_input("Ngày ban hành", value=date.today())
        trich  = st.text_area("Trích yếu")
        han    = st.text_input("Thời hạn xử lý (dd/mm/yyyy hh:mm)")
        noinhan= st.text_input("Nơi nhận")
        ngxl   = st.text_input("Người xử lý")
        ngaybc = st.text_input("Ngày báo cáo (ngăn cách ;)")
        bccho  = st.text_input("Báo cáo cho")
        files  = st.file_uploader("Chọn tệp đính kèm (nhiều)", accept_multiple_files=True)
        sub = st.form_submit_button("💾 Ghi mới")

    if sub:
        if not sovb.strip():
            st.error("Số văn bản bắt buộc."); st.stop()
        if (df["Số văn bản"].astype(str).str.strip() == sovb.strip()).any():
            st.error("Trùng Số văn bản."); st.stop()

        files_text = ""
        if files:
            files_text = add_files_and_return_list(drive, den_id, files)

        new = pd.DataFrame([{
            "Cơ quan ban hành": coquan,
            "Số văn bản": sovb,
            "Ngày ban hành": ngay.strftime(DATE_FMT),
            "Trích yếu": trich,
            "Thời hạn xử lý": han,
            "Nơi nhận": noinhan,
            "Người xử lý": ngxl,
            "Ngày báo cáo (;)": ngaybc,
            "Báo cáo cho": bccho,
            "Danh sách file (;)": files_text,
            "Ngày nhập": now_str()
        }])
        out = pd.concat([df, new], ignore_index=True)
        write_sheet(gc, SID_DEN, DEN_HEADERS, out)
        st.success("Đã ghi văn bản đến.")

    st.divider()
    st.subheader("Danh sách / Cập nhật nhanh")

    # Bộ lọc
    kw = st.text_input("Từ khóa (Số VB / Trích yếu)")
    my = st.text_input("Tháng/Năm (mm/yyyy | yyyy)")
    if kw or my:
        _df = df.copy()
        mask = _df.apply(lambda r:
                         (contains(r["Số văn bản"], kw) or contains(r["Trích yếu"], kw)) and
                         mm_yyyy_match(parse_dt(r["Ngày ban hành"]), my),
                         axis=1)
        _df = _df[mask]
    else:
        _df = df.copy()

    st.dataframe(_df, use_container_width=True, height=360)

def page_di(gc, drive):
    den_id, di_id, tap_id, _, SID_DI, _ = get_ids()
    st.header("2) Văn bản đi")

    df = read_sheet(gc, SID_DI, DI_HEADERS)
    df_den = read_sheet(gc, st.secrets["SHEET_DEN_ID"], DEN_HEADERS)

    with st.form("di_form", clear_on_submit=True):
        col1, col2, col3 = st.columns([1,1,1])
        with col1:
            so_ref = st.text_input("Số VB đến (tham chiếu)")
        with col2:
            so_di  = st.text_input("Số văn bản đi *")
        with col3:
            ngay   = st.date_input("Ngày ban hành", value=date.today())

        trich  = st.text_area("Trích yếu văn bản đi")
        c1, c2, c3 = st.columns(3)
        with c1: noinhan = st.text_input("Nơi nhận")
        with c2: ghichu  = st.text_input("Ghi chú")
        with c3: phong   = st.text_input("Phòng/Ban")

        files  = st.file_uploader("Chọn tệp đính kèm (nhiều)", accept_multiple_files=True)
        sub = st.form_submit_button("💾 Ghi mới")

    if sub:
        if not so_di.strip():
            st.error("Số văn bản đi bắt buộc."); st.stop()
        if (df["Số văn bản đi"].astype(str).str.strip() == so_di.strip()).any():
            st.error("Trùng Số văn bản đi."); st.stop()

        files_text = ""
        if files:
            files_text = add_files_and_return_list(drive, di_id, files)

        new = pd.DataFrame([{
            "Số văn bản đến (tham chiếu)": so_ref,
            "Số văn bản đi": so_di,
            "Ngày ban hành": ngay.strftime(DATE_FMT),
            "Trích yếu văn bản đi": trich,
            "Nơi nhận": noinhan,
            "Ghi chú": ghichu,
            "Danh sách file (;)": files_text,
            "Phòng/Ban": phong,
            "Ngày nhập": now_str()
        }])
        out = pd.concat([df, new], ignore_index=True)
        write_sheet(gc, SID_DI, DI_HEADERS, out)
        st.success("Đã ghi văn bản đi.")

    st.divider()
    st.subheader("Danh sách / Cập nhật nhanh")
    kw = st.text_input("Từ khóa (Số VB đi / Trích yếu)")
    my = st.text_input("Tháng/Năm (mm/yyyy | yyyy)")
    if kw or my:
        _df = df.copy()
        mask = _df.apply(lambda r:
                         (contains(r["Số văn bản đi"], kw) or contains(r["Trích yếu văn bản đi"], kw)) and
                         mm_yyyy_match(parse_dt(r["Ngày ban hành"]), my),
                         axis=1)
        _df = _df[mask]
    else:
        _df = df.copy()

    st.dataframe(_df, use_container_width=True, height=360)

    st.info("Tip: Muốn chọn tham chiếu từ VB đến, lọc trên sheet **Vb_Den** rồi copy Số VB vào ô 'Số VB đến (tham chiếu)'. (Có thể nâng cấp thành hộp chọn trực tiếp)")

def page_tap(gc, drive):
    den_id, di_id, tap_id, *_ = get_ids()
    st.header("3) Tập huấn / Giấy mời")

    df = read_sheet(gc, st.secrets["SHEET_TAP_ID"], TAP_HEADERS)

    with st.form("tap_form", clear_on_submit=True):
        so  = st.text_input("Số văn bản *")
        tr  = st.text_area("Trích yếu tập huấn")
        c1, c2 = st.columns(2)
        with c1:
            ngaygio = st.text_input("Ngày & giờ tập huấn (dd/mm/yyyy hh:mm)")
        with c2:
            diadiem = st.text_input("Địa điểm tập huấn")
        c3, c4 = st.columns(2)
        with c3:
            nguoidu = st.text_input("Người dự tập huấn")
        with c4:
            files   = st.file_uploader("Chọn tệp đính kèm (nhiều)", accept_multiple_files=True)

        sub = st.form_submit_button("💾 Ghi mới")

    if sub:
        if not so.strip():
            st.error("Số văn bản bắt buộc."); st.stop()
        if (df["Số văn bản"].astype(str).str.strip() == so.strip()).any():
            st.error("Trùng Số văn bản."); st.stop()

        files_text = ""
        if files:
            files_text = add_files_and_return_list(drive, tap_id, files)

        new = pd.DataFrame([{
            "Số văn bản": so,
            "Trích yếu tập huấn": tr,
            "Ngày & giờ tập huấn": ngaygio,
            "Địa điểm tập huấn": diadiem,
            "Người dự tập huấn": nguoidu,
            "Danh sách file (;)": files_text,
            "Ngày nhập": now_str()
        }])
        out = pd.concat([df, new], ignore_index=True)
        write_sheet(gc, st.secrets["SHEET_TAP_ID"], TAP_HEADERS, out)
        st.success("Đã ghi Tập huấn/Giấy mời.")

    st.divider()
    st.subheader("Danh sách / Cập nhật nhanh")
    kw = st.text_input("Từ khóa (Số VB / Trích yếu)")
    my = st.text_input("Tháng/Năm (mm/yyyy | yyyy)")
    if kw or my:
        _df = df.copy()
        mask = _df.apply(lambda r:
                         (contains(r["Số văn bản"], kw) or contains(r["Trích yếu tập huấn"], kw)) and
                         mm_yyyy_match(parse_dt(r["Ngày & giờ tập huấn"]), my),
                         axis=1)
        _df = _df[mask]
    else:
        _df = df.copy()

    st.dataframe(_df, use_container_width=True, height=360)

def page_search(gc, drive):
    st.header("🔎 Tìm kiếm")
    den_id, di_id, tap_id, SID_DEN, SID_DI, SID_TAP = get_ids()

    c1, c2, c3 = st.columns([2,1,1])
    with c1: kw = st.text_input("Từ khóa (Trích yếu / Số VB)")
    with c2: my = st.text_input("Tháng/Năm (mm/yyyy | yyyy)")
    with c3: loai = st.selectbox("Loại", ["ĐẾN","ĐI","TẬP HUẤN"])

    if st.button("Tìm"):
        if loai == "ĐẾN":
            df = read_sheet(gc, SID_DEN, DEN_HEADERS)
            df = df[df.apply(lambda r: (contains(r["Số văn bản"], kw) or contains(r["Trích yếu"], kw)) and mm_yyyy_match(parse_dt(r["Ngày ban hành"]), my), axis=1)]
            if len(df):
                st.dataframe(df, use_container_width=True, height=360)
            else:
                st.info("Không có kết quả.")
        elif loai == "ĐI":
            df = read_sheet(gc, SID_DI, DI_HEADERS)
            df = df[df.apply(lambda r: (contains(r["Số văn bản đi"], kw) or contains(r["Trích yếu văn bản đi"], kw)) and mm_yyyy_match(parse_dt(r["Ngày ban hành"]), my), axis=1)]
            if len(df):
                st.dataframe(df, use_container_width=True, height=360)
            else:
                st.info("Không có kết quả.")
        else:
            df = read_sheet(gc, SID_TAP, TAP_HEADERS)
            df = df[df.apply(lambda r: (contains(r["Số văn bản"], kw) or contains(r["Trích yếu tập huấn"], kw)) and mm_yyyy_match(parse_dt(r["Ngày & giờ tập huấn"]), my), axis=1)]
            if len(df):
                st.dataframe(df, use_container_width=True, height=360)
            else:
                st.info("Không có kết quả.")

    st.caption("Double-click mở tệp: trong phiên bản SA chuẩn, nên lưu kèm fileId để tạo link mở chính xác từng tệp.")

def page_config(gc, drive):
    st.header("⚙️ Cấu hình (read-only từ secrets)")
    den_id, di_id, tap_id, SID_DEN, SID_DI, SID_TAP = get_ids()
    st.code(f"VB_DEN_FOLDER = {den_id}")
    st.code(f"VB_DI_FOLDER  = {di_id}")
    st.code(f"TAP_FOLDER    = {tap_id}")
    st.code(f"SHEET_DEN_ID  = {SID_DEN}")
    st.code(f"SHEET_DI_ID   = {SID_DI}")
    st.code(f"SHEET_TAP_ID  = {SID_TAP}")

    # quick connectivity test
    if st.button("🔌 Test kết nối Sheets & Drive"):
        try:
            _ = read_sheet(gc, SID_DEN, DEN_HEADERS)
            _ = read_sheet(gc, SID_DI, DI_HEADERS)
            _ = read_sheet(gc, SID_TAP, TAP_HEADERS)
            # thử list file trong folder đến (tối đa 1)
            res = drive.files().list(q=f"'{den_id}' in parents and trashed=false", pageSize=1, fields="files(id,name)").execute()
            st.success("Kết nối OK.")
            if res.get("files"):
                st.write("Ví dụ file trong VB_DEN_FOLDER:", res["files"][0])
        except Exception as e:
            st.error(f"Lỗi kết nối: {e}")

# =========================
# ========= MAIN ==========
# =========================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown("""
        <style>
        .stMarkdown, .stTextInput, .stTextArea, .stSelectbox, .stDateInput, .stDataFrame { font-size: 16px; }
        </style>
    """, unsafe_allow_html=True)

    try:
        gc, drive = get_google_clients()
    except Exception as e:
        st.error(f"Lỗi khởi tạo Google clients: {e}")
        st.stop()

    user = st.session_state.get("user")
    if not user:
        do_login()
        return

    st.sidebar.success(f"Xin chào, {user}")
    page = topbar(gc, drive)

    if page == "Dashboard":
        page_dashboard(gc, drive)
    elif page == "1. Văn bản đến":
        page_den(gc, drive)
    elif page == "2. Văn bản đi":
        page_di(gc, drive)
    elif page == "3. Tập huấn / Giấy mời":
        page_tap(gc, drive)
    elif page == "Tìm kiếm":
        page_search(gc, drive)
    else:
        page_config(gc, drive)

if __name__ == "__main__":
    main()
