"""
app.py  —  FILE CHÍNH ĐỂ CHẠY APP
=================================
Chạy bằng lệnh:   streamlit run app.py

App gồm 2 trang (chọn ở thanh bên trái):
  • "Tạo hóa đơn": upload file ezcloud -> lọc checkout theo ngày -> bảng hóa đơn -> xuất file.
  • "Cài đặt": chỉnh thuế suất, công ty chuyên gia, list phòng, đồ uống khác.

Lưu ý kỹ thuật cho người mới:
  - Streamlit chạy lại (rerun) TOÀN BỘ file này mỗi khi bạn bấm/sửa gì đó.
    Vì vậy dữ liệu "đang làm" được giữ trong st.session_state (bộ nhớ tạm của phiên).
  - Bảng hóa đơn được lưu ở st.session_state.bang dưới dạng 1 danh sách các dict.
  - Khi tắt/refresh trình duyệt, bảng đang làm sẽ mất (đúng yêu cầu: chỉ giữ trong phiên).
    Còn Cài đặt + 'số lượng đã xuất' của đồ uống thì lưu trong file hoadon.db (vĩnh viễn).
"""

import datetime
import math

import pandas as pd
import streamlit as st

import database as db
import excel_reader as er
import bangsoat_reader as bsr
import invoice_logic as L
import export_excel as ex
import settings_ui

# st_aggrid là thư viện vẽ bảng lưới giống Excel (header gộp, dropdown, checkbox...)
from st_aggrid import AgGrid, GridUpdateMode, DataReturnMode, JsCode


# =========================================================================
# CÁC HÀM HỖ TRỢ
# =========================================================================

def lay_settings():
    """Gom các thông số từ DB thành 1 dict cho tiện dùng."""
    return {
        "chuyen_gia_companies": db.get_chuyen_gia_companies(),
        "fallback": db.get_gia_co_ban_fallback(),
        "hinh_thuc_thanh_toan": db.get_hinh_thuc_thanh_toan(),
        "thue_suat": db.get_thue_suat(),
    }


def danh_sach_cot(so_nhom):
    """Trả về thứ tự các cột của bảng hóa đơn (gồm cả cột ẩn để tính toán)."""
    cols = list(L.COT_AN) + list(L.COT_THONG_TIN)
    for g in range(1, so_nhom + 1):
        for ht in L.HAU_TO_NHOM:
            cols.append(L.key_nhom(g, ht))
    cols += list(L.COT_TONG)
    return cols


def ten_nhom(g):
    """Tên hiển thị của nhóm hàng hóa g."""
    if g == 1:
        return "HÀNG HÓA, DỊCH VỤ 1 - Tiền phòng"
    if g == 2:
        return "HÀNG HÓA, DỊCH VỤ 2 - Dasani"
    if g == 3:
        return "HÀNG HÓA, DỊCH VỤ 3 - Đồ uống khác"
    return f"HÀNG HÓA, DỊCH VỤ {g}"


def tao_grid_options(so_nhom, ten_do_uong):
    """
    Tạo cấu hình cho bảng lưới AgGrid:
      - Cột ẩn (id, ngày, giá TB...) -> hide.
      - Cột thông tin -> sửa được.
      - 'Loại hóa đơn' -> dropdown 3 giá trị.
      - Mỗi nhóm hàng hóa -> 1 header gộp với 5 cột con. 'Thành tiền' không sửa được.
      - Nhóm 3 (Đồ uống khác) -> ô 'Tên' là dropdown chọn từ danh sách đồ uống.
      - Cột tổng -> không sửa được.
    """
    column_defs = []

    # Bộ định dạng HIỂN THỊ cho cột tiền: thêm dấu '.' phân tách hàng nghìn.
    # CHỈ đổi cách hiển thị, KHÔNG đổi giá trị thật -> file output vẫn là số trơn (1200000).
    # Ví dụ: 1200000 -> "1.200.000".
    dinh_dang_tien = JsCode(r"""
    function(params) {
        var v = params.value;
        if (v === null || v === undefined || v === '') { return ''; }
        var n = Number(v);
        if (isNaN(n)) { return v; }
        n = Math.round(n);
        return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    }
    """)

    # Các cột sẽ được GHIM (freeze) bên trái: scroll ngang vẫn luôn nhìn thấy.
    cot_ghim = {"Mã đặt phòng", "Phòng", "Họ tên người mua hàng", "Loại hóa đơn"}
    # Các cột căn lề TRÁI (mọi cột khác mặc định căn GIỮA).
    # Lưu ý: chỉ "Tên" của nhóm 1 (HH1_Tên) căn trái, còn Tên nhóm 2/3 vẫn căn giữa.
    cot_can_trai = {"Họ tên người mua hàng", "Loại hóa đơn", "HH1_Tên"}

    # Cột ẩn
    for col in L.COT_AN:
        column_defs.append({"field": col, "hide": True})

    # Cột thông tin (cột đầu tiên có thêm checkbox để chọn dòng)
    for i, col in enumerate(L.COT_THONG_TIN):
        d = {"field": col, "headerName": col, "editable": True, "minWidth": 130}
        if i == 0:
            d["checkboxSelection"] = True
            d["headerCheckboxSelection"] = True
            d["minWidth"] = 150
        if col == "Loại hóa đơn":
            d["cellEditor"] = "agSelectCellEditor"
            d["cellEditorParams"] = {"values": L.DANH_SACH_LOAI}
            d["minWidth"] = 240
        if col in cot_ghim:
            d["pinned"] = "left"  # ghim cột bên trái
        if col in cot_can_trai:
            d["cellStyle"] = {"textAlign": "left"}  # cột này căn trái
        column_defs.append(d)

    # Các nhóm hàng hóa (header gộp)
    nhan_con = {
        "Tên": "Tên", "Đơn vị": "Đơn vị tính", "Số lượng": "Số lượng",
        "Đơn giá": "Đơn giá", "Thành tiền": "Thành tiền",
    }
    for g in range(1, so_nhom + 1):
        children = []
        for ht in L.HAU_TO_NHOM:
            field = L.key_nhom(g, ht)
            con = {
                "field": field,
                "headerName": nhan_con[ht],
                "editable": (ht != "Thành tiền"),  # Thành tiền tự tính, không cho sửa
                "minWidth": 110,
            }
            if ht == "Tên":
                con["minWidth"] = 220
                # Nhóm 3 (Đồ uống khác): ô Tên là dropdown
                if g == 3:
                    con["cellEditor"] = "agSelectCellEditor"
                    con["cellEditorParams"] = {"values": [""] + ten_do_uong}
            if ht in ("Đơn giá", "Thành tiền"):
                con["valueFormatter"] = dinh_dang_tien  # hiển thị có dấu '.'
            if field in cot_can_trai:
                con["cellStyle"] = {"textAlign": "left"}  # chỉ HH1_Tên căn trái
            children.append(con)
        # headerClass giúp căn giữa chữ tiêu đề nhóm (CSS định nghĩa ở chỗ gọi AgGrid)
        column_defs.append(
            {"headerName": ten_nhom(g), "headerClass": "nhom-can-giua", "children": children}
        )

    # Cột tổng (chỉ hiển thị)
    for col in L.COT_TONG:
        column_defs.append(
            {"field": col, "headerName": col, "editable": False, "minWidth": 120,
             "valueFormatter": dinh_dang_tien}  # hiển thị có dấu '.'
        )

    grid_options = {
        "columnDefs": column_defs,
        # sortable=True: cho phép bấm vào header để sắp xếp. cellStyle: căn giữa mặc định.
        "defaultColDef": {
            "resizable": True,
            "sortable": True,
            "filter": False,
            "cellStyle": {"textAlign": "center"},
        },
        "rowSelection": "multiple",
        "suppressRowClickSelection": True,  # click ô để sửa, không vô tình chọn dòng
        "rowHeight": 38,          # dòng cao hơn cho dễ đọc
        "headerHeight": 36,
        "groupHeaderHeight": 36,
    }
    return grid_options


def xu_ly_sau_khi_sua(edited_records, so_nhom, settings, room_map, drinks):
    """
    Xử lý dữ liệu trả về từ bảng sau khi user sửa:
      1) Nếu user ĐỔI 'Loại hóa đơn' của 1 dòng -> điền lại mặc định các nhóm cho dòng đó.
      2) Tự điền vài mặc định nhỏ (Đơn vị 'Chai', Đơn giá Dasani 0).
      3) Nhóm 3 ở chế độ 'phải theo thực tế': nếu Tên là 1 đồ uống đã biết -> Đơn giá = giá bán.
      4) Tính lại toàn bộ tiền.
    Trả về True nếu có thay đổi cấu trúc cần làm mới bảng (vd vừa reset theo loại hóa đơn).
    """
    gia_do_uong = {d["ten"]: d["gia_ban"] for d in drinks}
    can_lam_moi = False
    bang_moi = []

    for row in edited_records:
        rid = row.get("_row_id")

        # (1) Phát hiện đổi loại hóa đơn
        loai_truoc = st.session_state.prev_loai.get(rid)
        if loai_truoc is not None and row.get("Loại hóa đơn") != loai_truoc:
            L.dien_mac_dinh_cac_nhom(row, room_map, settings["fallback"], drinks)
            can_lam_moi = True
        st.session_state.prev_loai[rid] = row.get("Loại hóa đơn")

        # (2) Mặc định nhỏ cho nhóm 2 & 3: có Tên thì Đơn vị = 'Chai'
        for g in (2, 3):
            if not L.la_rong(row.get(L.key_nhom(g, "Tên"))) and L.la_rong(row.get(L.key_nhom(g, "Đơn vị"))):
                row[L.key_nhom(g, "Đơn vị")] = "Chai"
        # Dasani: có Tên mà chưa có Đơn giá -> 0
        if not L.la_rong(row.get("HH2_Tên")) and L.la_rong(row.get("HH2_Đơn giá")):
            row["HH2_Đơn giá"] = 0

        # (3) Nhóm 3, chế độ 'phải theo thực tế': lấy đơn giá theo giá bán đồ uống
        if row.get("Loại hóa đơn") == L.LOAI_CHUYEN_GIA_PHAI:
            ten3 = str(row.get("HH3_Tên") or "").strip()
            if ten3 in gia_do_uong:
                row["HH3_Đơn giá"] = gia_do_uong[ten3]

        # (4) Tính lại tiền
        L.cap_nhat_tat_ca_tien(row, so_nhom, settings["thue_suat"])
        bang_moi.append(row)

    st.session_state.bang = bang_moi
    return can_lam_moi


def lam_sach_nan(df):
    """Đổi mọi ô NaN (rỗng của pandas) thành chuỗi rỗng để khỏi hiện 'nan' trên bảng."""
    return df.where(pd.notna(df), "")


# =========================================================================
# HỘP THOẠI (DIALOG)
# =========================================================================

@st.dialog("Gộp các dòng đã chọn thành 1 dòng mới")
def hop_thoai_gop(selected_ids, so_nhom, settings):
    st.write(f"Bạn đang gộp **{len(selected_ids)} dòng**. "
             "Hãy nhập thông tin cho DÒNG MỚI bên dưới rồi bấm Xác nhận. "
             "Các dòng cũ đã chọn sẽ bị xóa.")

    row = {}
    # Thông tin cơ bản
    row["Mã đặt phòng"] = st.text_input("Mã đặt phòng", key="gop_madp")
    row["Phòng"] = st.text_input("Phòng", key="gop_phong")
    row["Họ tên người mua hàng"] = st.text_input("Họ tên người mua hàng", key="gop_hoten")
    row["Loại hóa đơn"] = st.selectbox("Loại hóa đơn", L.DANH_SACH_LOAI, key="gop_loai")
    row["Tên đơn vị mua hàng"] = st.text_input("Tên đơn vị mua hàng", key="gop_donvi")
    row["Mã số thuế"] = st.text_input("Mã số thuế", key="gop_mst")
    row["Địa chỉ"] = st.text_input("Địa chỉ", key="gop_diachi")
    row["Hình thức thanh toán"] = st.text_input(
        "Hình thức thanh toán", value=settings["hinh_thuc_thanh_toan"], key="gop_httt"
    )

    # Các nhóm hàng hóa (nhập tay)
    for g in range(1, so_nhom + 1):
        st.markdown(f"**{ten_nhom(g)}**")
        c1, c2, c3, c4 = st.columns(4)
        row[L.key_nhom(g, "Tên")] = c1.text_input("Tên", key=f"gop_{g}_ten")
        row[L.key_nhom(g, "Đơn vị")] = c2.text_input("Đơn vị", key=f"gop_{g}_dv")
        row[L.key_nhom(g, "Số lượng")] = c3.text_input("Số lượng", key=f"gop_{g}_sl")
        row[L.key_nhom(g, "Đơn giá")] = c4.text_input("Đơn giá", key=f"gop_{g}_dg")
        row[L.key_nhom(g, "Thành tiền")] = ""

    if st.button("✅ Xác nhận gộp", type="primary"):
        import uuid
        # Bổ sung các cột ẩn cho dòng mới
        row["_row_id"] = str(uuid.uuid4())
        row["Công ty"] = ""
        row["Ngày đến"] = ""
        row["Ngày đi"] = ""
        row["Số đêm"] = L.to_int(row.get("HH1_Số lượng"))
        row["Giá trung bình"] = 0
        L.cap_nhat_tat_ca_tien(row, so_nhom, settings["thue_suat"])

        # Xóa các dòng đã chọn, thêm dòng mới vào cuối
        st.session_state.bang = [
            r for r in st.session_state.bang if r.get("_row_id") not in selected_ids
        ]
        st.session_state.bang.append(row)
        st.session_state.prev_loai[row["_row_id"]] = row["Loại hóa đơn"]
        st.session_state.grid_key += 1  # làm mới bảng
        st.rerun()


@st.dialog("Xác nhận tính số lượng 'Đồ uống khác' đã xuất")
def hop_thoai_xuat(so_nhom, settings, drinks):
    st.write("Bạn có muốn **cộng dồn số lượng 'Đồ uống khác'** đã xuất trong danh sách "
             "này vào bảng Đồ uống khác không?")
    st.caption("• Đồng ý: file được tạo VÀ cộng dồn số lượng.\n\n"
               "• Không: file vẫn được tạo nhưng KHÔNG đổi số lượng đã xuất.")

    ngay_str = st.session_state.ngay_hoa_don.strftime("%d/%m/%Y")
    ten_do_uong = {d["ten"] for d in drinks}

    c1, c2 = st.columns(2)
    dong_y = c1.button("✅ Đồng ý (có cộng dồn)", type="primary")
    khong = c2.button("➡️ Không (chỉ tạo file)")

    if dong_y or khong:
        # Nếu đồng ý: cộng dồn số lượng cho các đồ uống ở NHÓM 3 mà tên khớp danh sách
        if dong_y:
            for row in st.session_state.bang:
                if L.nhom_co_du_lieu(row, 3):
                    ten3 = str(row.get("HH3_Tên") or "").strip()
                    sl3 = L.to_int(row.get("HH3_Số lượng"))
                    if ten3 in ten_do_uong and sl3 > 0:
                        db.add_da_xuat_by_name(ten3, sl3)

        # Tạo file output (dạng bytes) và lưu vào phiên để hiện nút tải về
        data = ex.xuat_file_bytes(
            st.session_state.bang, so_nhom, ngay_str, settings["thue_suat"]
        )
        st.session_state.file_bytes = data
        st.session_state.file_name = f"DanhSachHoaDon_{ngay_str.replace('/', '_')}.xlsx"
        st.session_state.da_cong_don = dong_y
        st.rerun()


# =========================================================================
# TRANG: TẠO HÓA ĐƠN
# =========================================================================

def trang_tao_hoa_don():
    st.header("🧾 Tạo hóa đơn")

    settings = lay_settings()
    room_map, room_err = db.build_room_price_map()
    drinks = db.get_drinks()
    ten_do_uong = [d["ten"] for d in drinks]

    if room_err:
        for e in room_err:
            st.warning("⚠️ Cài đặt phòng: " + e + " (vào trang Cài đặt để sửa)")

    # ----- Bước 1: chọn file + ngày -----
    c1, c2 = st.columns([2, 1])
    file = c1.file_uploader("Chọn file doanh thu chi tiết (ezcloud, .xlsx)", type=["xlsx"])
    ngay_loc = c2.date_input("Ngày checkout của hóa đơn", value=datetime.date.today())

    if st.button("📥 Đọc file & tạo bảng hóa đơn", type="primary"):
        if file is None:
            st.error("Bạn chưa chọn file.")
        else:
            bookings, canh_bao = er.doc_va_loc_file(file, ngay_loc)
            for w in canh_bao:
                st.warning(w)
            if not bookings:
                st.info("Không tìm thấy phòng nào CHECKOUT đúng ngày đã chọn.")
            # Tạo bảng hóa đơn từ các booking
            bang = [L.tao_dong_tu_booking(b, settings, room_map, drinks) for b in bookings]
            for d in bang:
                L.cap_nhat_tat_ca_tien(d, 3, settings["thue_suat"])
            st.session_state.bang = bang
            st.session_state.so_nhom = 3
            st.session_state.prev_loai = {d["_row_id"]: d["Loại hóa đơn"] for d in bang}
            st.session_state.grid_key += 1
            st.session_state.ngay_hoa_don = ngay_loc
            # Xóa file cũ (nếu có) để tránh nhầm
            st.session_state.pop("file_bytes", None)
            st.success(f"Đã tạo bảng với {len(bang)} dòng. Ngày hóa đơn = {ngay_loc.strftime('%d/%m/%Y')}.")

    # Nếu chưa có bảng thì dừng ở đây
    if not st.session_state.get("bang"):
        st.info("Hãy chọn file và bấm 'Đọc file & tạo bảng hóa đơn' để bắt đầu.")
        return

    so_nhom = st.session_state.so_nhom

    # ----- Bước 2: hiển thị bảng lưới để sửa -----
    st.markdown("### Bảng hóa đơn (sửa trực tiếp trên ô)")
    st.caption("Nhấp đúp vào 1 ô để sửa. Tick checkbox ở cột đầu để chọn dòng (xóa/gộp). "
               "Cột 'Thành tiền' và các cột Tổng được tính tự động.")

    # Chuẩn bị dữ liệu dạng bảng (DataFrame) theo đúng thứ tự cột
    df = pd.DataFrame(st.session_state.bang)
    df = df.reindex(columns=danh_sach_cot(so_nhom), fill_value="")
    df = lam_sach_nan(df)

    grid_options = tao_grid_options(so_nhom, ten_do_uong)

    # CSS cho bảng:
    #  - Căn GIỮA chữ ở mọi header (cả header nhóm "HÀNG HÓA DỊCH VỤ..." lẫn header cột).
    #  - Phóng to cỡ chữ cho dễ đọc.
    css_bang = {
        # header cột con
        ".ag-header-cell-label": {"justify-content": "center"},
        ".ag-header-cell-text": {"text-align": "center", "font-size": "15px"},
        # header nhóm (dòng "HÀNG HÓA, DỊCH VỤ ...")
        ".ag-header-group-cell-label": {"justify-content": "center"},
        ".ag-header-group-text": {"text-align": "center", "width": "100%",
                                  "font-size": "15px", "font-weight": "700"},
        # nội dung các ô
        ".ag-cell": {"font-size": "15px"},
    }

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.VALUE_CHANGED | GridUpdateMode.SELECTION_CHANGED,
        data_return_mode=DataReturnMode.AS_INPUT,
        reload_data=True,            # nạp lại dữ liệu từ df mỗi lần chạy (để thấy số tự tính)
        allow_unsafe_jscode=True,    # cho phép valueFormatter (định dạng dấu '.')
        fit_columns_on_grid_load=False,
        custom_css=css_bang,         # căn giữa header + phóng to chữ
        height=750,                  # bảng cao gần full màn hình
        key=f"grid_{st.session_state.grid_key}",
    )

    # Lấy dữ liệu đã sửa và xử lý lại
    edited_df = lam_sach_nan(pd.DataFrame(grid_response["data"]))
    edited_records = edited_df.to_dict("records")
    can_lam_moi = xu_ly_sau_khi_sua(edited_records, so_nhom, settings, room_map, drinks)

    # Lấy danh sách _row_id đang được chọn (checkbox)
    selected = grid_response.get("selected_rows")
    if selected is None:
        selected_ids = []
    elif isinstance(selected, list):
        selected_ids = [r.get("_row_id") for r in selected]
    else:  # phiên bản mới của st_aggrid trả về DataFrame
        selected_ids = selected["_row_id"].tolist() if "_row_id" in selected.columns else []

    # ----- Bước 2b: các nút thao tác -----
    b1, b2, b3, b4, b5 = st.columns(5)

    if b1.button("➕ Thêm dòng trống"):
        them_dong_trong(so_nhom, settings)
        st.rerun()

    if b2.button("🗑️ Xóa dòng đã chọn"):
        if selected_ids:
            st.session_state.bang = [
                r for r in st.session_state.bang if r.get("_row_id") not in selected_ids
            ]
            st.session_state.grid_key += 1
            st.rerun()
        else:
            st.warning("Chưa chọn dòng nào để xóa.")

    if b3.button("🔀 Gộp dòng đã chọn"):
        if len(selected_ids) >= 2:
            hop_thoai_gop(selected_ids, so_nhom, settings)
        else:
            st.warning("Cần chọn ít nhất 2 dòng để gộp.")

    if b4.button("➕ Thêm tab HÀNG HÓA DỊCH VỤ"):
        them_tab(so_nhom)
        st.rerun()

    if b5.button("🔄 Tính lại bảng"):
        st.session_state.grid_key += 1
        st.rerun()

    # Làm mới bảng nếu vừa reset theo loại hóa đơn
    if can_lam_moi:
        st.session_state.grid_key += 1
        st.rerun()

    # ----- Bước 3: kiểm tra & xuất file -----
    st.markdown("### Xuất file")
    loi = []
    for i, row in enumerate(st.session_state.bang, start=1):
        loi += L.kiem_tra_dong(row, so_nhom, i)

    if loi:
        st.error("Còn lỗi cần sửa trước khi xuất file:")
        for e in loi:
            st.write("• " + e)
    else:
        if st.button("📤 Xuất file danh sách hóa đơn", type="primary"):
            hop_thoai_xuat(so_nhom, settings, drinks)

    # Nếu file đã được tạo (sau khi xác nhận ở hộp thoại) -> hiện nút tải về
    if st.session_state.get("file_bytes"):
        if st.session_state.get("da_cong_don"):
            st.success("Đã tạo file VÀ cộng dồn số lượng Đồ uống khác.")
        else:
            st.success("Đã tạo file (không cộng dồn số lượng).")
        st.download_button(
            "⬇️ Tải file Excel về máy",
            data=st.session_state.file_bytes,
            file_name=st.session_state.file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def them_dong_trong(so_nhom, settings):
    """Thêm 1 dòng trống (nhập tay) vào cuối bảng."""
    import uuid
    row = {c: "" for c in danh_sach_cot(so_nhom)}
    row["_row_id"] = str(uuid.uuid4())
    row["Loại hóa đơn"] = L.LOAI_CO_BAN
    row["Hình thức thanh toán"] = settings["hinh_thuc_thanh_toan"]
    row["Số đêm"] = ""
    row["Giá trung bình"] = 0
    L.cap_nhat_tat_ca_tien(row, so_nhom, settings["thue_suat"])
    st.session_state.bang.append(row)
    st.session_state.prev_loai[row["_row_id"]] = row["Loại hóa đơn"]
    st.session_state.grid_key += 1


def them_tab(so_nhom):
    """Thêm 1 nhóm hàng hóa mới (tab) với các ô rỗng cho mọi dòng."""
    g = so_nhom + 1
    for row in st.session_state.bang:
        for ht in L.HAU_TO_NHOM:
            row[L.key_nhom(g, ht)] = ""
    st.session_state.so_nhom = g
    st.session_state.grid_key += 1


# =========================================================================
# =========================================================================
# TÍNH NĂNG MỚI: TẠO HÓA ĐƠN TỪ "BẢNG SOÁT HÓA ĐƠN"
# (Trang riêng, bảng riêng, tách biệt hoàn toàn với tính năng ezcloud ở trên)
# =========================================================================
# =========================================================================

def danh_sach_cot_bs(so_nhom):
    """Thứ tự cột của bảng Bảng soát (có thêm cột Email, không có các cột ẩn của ezcloud)."""
    cols = list(L.COT_AN_BS) + list(L.COT_THONG_TIN_BS)
    for g in range(1, so_nhom + 1):
        for ht in L.HAU_TO_NHOM:
            cols.append(L.key_nhom(g, ht))
    cols += list(L.COT_TONG)
    return cols


def tao_grid_options_bs(so_nhom):
    """
    Cấu hình bảng lưới cho tính năng Bảng soát. Khác bảng ezcloud:
      - 'Loại hóa đơn' chỉ 2 giá trị.
      - Có cột 'Email'.
      - Nhóm 1 (Tiền phòng): 'Đơn giá' KHÓA (tự tính = Thành tiền/Số lượng),
        'Thành tiền' CHO SỬA tay.
      - Nhóm 3 (Đồ uống khác): 'Tên' là chữ tự do (không dropdown).
      - Tiền hiển thị có dấu '.', số lẻ giữ 2 chữ số thập phân (vd 92.592,66).
    """
    column_defs = []

    # Định dạng tiền (có hỗ trợ số thập phân): nguyên -> "1.234.567"; lẻ -> "1.234.567,89"
    dinh_dang_tien_le = JsCode(r"""
    function(params){
        var v = params.value;
        if (v === null || v === undefined || v === '') { return ''; }
        var n = Number(v); if (isNaN(n)) { return v; }
        var neg = n < 0; n = Math.abs(n);
        var hasDec = (Math.round(n*100) % 100) !== 0;
        var s;
        if (hasDec) {
            s = n.toFixed(2);
            var p = s.split('.');
            p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
            s = p[0] + ',' + p[1];
        } else {
            s = Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
        }
        return (neg ? '-' : '') + s;
    }
    """)

    cot_ghim = {"Mã đặt phòng", "Phòng", "Họ tên người mua hàng", "CCCD", "PASSPORT", "Loại hóa đơn"}
    cot_can_trai = {"Họ tên người mua hàng", "Loại hóa đơn", "HH1_Tên"}

    # Cột ẩn
    for col in L.COT_AN_BS:
        column_defs.append({"field": col, "hide": True})

    # Cột thông tin
    for i, col in enumerate(L.COT_THONG_TIN_BS):
        d = {"field": col, "headerName": col, "editable": True, "minWidth": 130}
        if i == 0:
            d["checkboxSelection"] = True
            d["headerCheckboxSelection"] = True
            d["minWidth"] = 150
        if col == "Loại hóa đơn":
            d["cellEditor"] = "agSelectCellEditor"
            d["cellEditorParams"] = {"values": L.DANH_SACH_LOAI_BS}  # CHỈ 2 giá trị
            d["minWidth"] = 200
        if col in cot_ghim:
            d["pinned"] = "left"
        if col in cot_can_trai:
            d["cellStyle"] = {"textAlign": "left"}
        column_defs.append(d)

    # Các nhóm hàng hóa
    nhan_con = {
        "Tên": "Tên", "Đơn vị": "Đơn vị tính", "Số lượng": "Số lượng",
        "Đơn giá": "Đơn giá", "Thành tiền": "Thành tiền",
    }
    for g in range(1, so_nhom + 1):
        children = []
        for ht in L.HAU_TO_NHOM:
            field = L.key_nhom(g, ht)
            # Nhóm 1: Đơn giá khóa (tự tính), Thành tiền cho sửa.
            # Nhóm khác: Thành tiền khóa (tự tính = Đơn giá * Số lượng).
            if g == 1:
                editable = ht in ("Tên", "Đơn vị", "Số lượng", "Thành tiền")
            else:
                editable = ht != "Thành tiền"
            con = {
                "field": field,
                "headerName": nhan_con[ht],
                "editable": editable,
                "minWidth": 110,
            }
            if ht == "Tên":
                con["minWidth"] = 220  # Bảng soát: Tên là chữ tự do (không dropdown)
            if ht in ("Đơn giá", "Thành tiền"):
                con["valueFormatter"] = dinh_dang_tien_le
            if field in cot_can_trai:
                con["cellStyle"] = {"textAlign": "left"}
            children.append(con)
        column_defs.append(
            {"headerName": ten_nhom(g), "headerClass": "nhom-can-giua", "children": children}
        )

    # Cột tổng
    for col in L.COT_TONG:
        column_defs.append(
            {"field": col, "headerName": col, "editable": False, "minWidth": 120,
             "valueFormatter": dinh_dang_tien_le}
        )

    return {
        "columnDefs": column_defs,
        "defaultColDef": {
            "resizable": True, "sortable": True, "filter": False,
            "cellStyle": {"textAlign": "center"},
        },
        "rowSelection": "multiple",
        "suppressRowClickSelection": True,
        "rowHeight": 38,
        "headerHeight": 36,
        "groupHeaderHeight": 36,
    }


def xu_ly_sau_khi_sua_bs(edited_records, so_nhom, thue_suat):
    """
    Xử lý sau khi user sửa bảng Bảng soát:
      - Nếu đổi 'Loại hóa đơn' HOẶC 'Số lượng' tab Tiền phòng -> cập nhật lại chữ Tên HH1.
        (Đổi loại hóa đơn CHỈ đổi chữ Tên HH1, không đụng Dasani/Đồ uống.)
      - Tính lại: Đơn giá HH1 = Thành tiền/Số lượng (2 chữ số thập phân);
        Thành tiền HH2/HH3 = Đơn giá*Số lượng; Tổng tiền hàng.
    """
    can_lam_moi = False
    bang_moi = []
    for row in edited_records:
        rid = row.get("_row_id")
        loai = row.get("Loại hóa đơn")
        so_dem = L.to_int(row.get("HH1_Số lượng"))
        khoa = (loai, so_dem)
        truoc = st.session_state.prev_loai_bs.get(rid)
        if truoc is not None and truoc != khoa:
            row["HH1_Tên"] = L.ten_tien_phong_bangsoat(loai, row.get("Ngày hóa đơn"), so_dem)
            can_lam_moi = True
        st.session_state.prev_loai_bs[rid] = khoa

        # Chuẩn hóa Mã số thuế (thêm '0' đầu nếu thiếu) để hiển thị đúng ngay trên bảng
        mst_moi = L.chuan_hoa_mst(row.get("Mã số thuế"))
        if mst_moi != (row.get("Mã số thuế") or ""):
            row["Mã số thuế"] = mst_moi
            can_lam_moi = True

        # Chuẩn hóa CCCD (thêm '0' đầu nếu là dãy số 11 chữ số) để hiển thị đúng trên bảng
        cccd_moi = L.chuan_hoa_cccd(row.get("CCCD"))
        if cccd_moi != (row.get("CCCD") or ""):
            row["CCCD"] = cccd_moi
            can_lam_moi = True

        L.cap_nhat_tien_bangsoat(row, so_nhom, thue_suat)
        bang_moi.append(row)
    st.session_state.bang_bs = bang_moi
    return can_lam_moi


@st.dialog("Gộp các dòng đã chọn thành 1 dòng mới")
def hop_thoai_gop_bs(selected_ids, so_nhom, httt):
    st.write(f"Bạn đang gộp **{len(selected_ids)} dòng**. Nhập thông tin DÒNG MỚI rồi bấm Xác nhận.")
    row = {}
    row["Mã đặt phòng"] = st.text_input("Mã đặt phòng", key="gopbs_madp")
    row["Phòng"] = st.text_input("Phòng", key="gopbs_phong")
    row["Họ tên người mua hàng"] = st.text_input("Họ tên người mua hàng", key="gopbs_hoten")
    row["CCCD"] = st.text_input("CCCD", key="gopbs_cccd")
    row["PASSPORT"] = st.text_input("PASSPORT", key="gopbs_passport")
    row["Loại hóa đơn"] = st.selectbox("Loại hóa đơn", L.DANH_SACH_LOAI_BS, key="gopbs_loai")
    row["Tên đơn vị mua hàng"] = st.text_input("Tên đơn vị mua hàng", key="gopbs_donvi")
    row["Mã số thuế"] = st.text_input("Mã số thuế", key="gopbs_mst")
    row["Địa chỉ"] = st.text_input("Địa chỉ", key="gopbs_diachi")
    row["Email"] = st.text_input("Email", key="gopbs_email")
    row["Hình thức thanh toán"] = st.text_input("Hình thức thanh toán", value=httt, key="gopbs_httt")

    for g in range(1, so_nhom + 1):
        st.markdown(f"**{ten_nhom(g)}**")
        c1, c2, c3, c4 = st.columns(4)
        row[L.key_nhom(g, "Tên")] = c1.text_input("Tên", key=f"gopbs_{g}_ten")
        row[L.key_nhom(g, "Đơn vị")] = c2.text_input("Đơn vị", key=f"gopbs_{g}_dv")
        row[L.key_nhom(g, "Số lượng")] = c3.text_input("Số lượng", key=f"gopbs_{g}_sl")
        # Nhóm 1: nhập Thành tiền (Đơn giá tự tính). Nhóm khác: nhập Đơn giá.
        if g == 1:
            row[L.key_nhom(g, "Thành tiền")] = c4.text_input("Thành tiền", key=f"gopbs_{g}_tt")
            row[L.key_nhom(g, "Đơn giá")] = ""
        else:
            row[L.key_nhom(g, "Đơn giá")] = c4.text_input("Đơn giá", key=f"gopbs_{g}_dg")
            row[L.key_nhom(g, "Thành tiền")] = ""

    if st.button("✅ Xác nhận gộp", type="primary"):
        import uuid
        row["_row_id"] = str(uuid.uuid4())
        row["Ngày hóa đơn"] = st.session_state.ngay_hoa_don_bs.strftime("%d/%m/%Y")
        L.cap_nhat_tien_bangsoat(row, so_nhom, db.get_thue_suat())
        st.session_state.bang_bs = [
            r for r in st.session_state.bang_bs if r.get("_row_id") not in selected_ids
        ]
        st.session_state.bang_bs.append(row)
        st.session_state.prev_loai_bs[row["_row_id"]] = (
            row["Loại hóa đơn"], L.to_int(row.get("HH1_Số lượng"))
        )
        st.session_state.grid_key_bs += 1
        st.rerun()


def them_dong_trong_bs(so_nhom, httt):
    """Thêm 1 dòng trống vào bảng Bảng soát."""
    import uuid
    row = {c: "" for c in danh_sach_cot_bs(so_nhom)}
    row["_row_id"] = str(uuid.uuid4())
    row["Ngày hóa đơn"] = st.session_state.ngay_hoa_don_bs.strftime("%d/%m/%Y")
    row["Loại hóa đơn"] = L.LOAI_BS_CO_BAN
    row["Hình thức thanh toán"] = httt
    L.cap_nhat_tien_bangsoat(row, so_nhom, db.get_thue_suat())
    st.session_state.bang_bs.append(row)
    st.session_state.prev_loai_bs[row["_row_id"]] = (row["Loại hóa đơn"], 0)
    st.session_state.grid_key_bs += 1


def them_tab_bs(so_nhom):
    """Thêm 1 nhóm hàng hóa mới cho bảng Bảng soát."""
    g = so_nhom + 1
    for row in st.session_state.bang_bs:
        for ht in L.HAU_TO_NHOM:
            row[L.key_nhom(g, ht)] = ""
    st.session_state.so_nhom_bs = g
    st.session_state.grid_key_bs += 1


def trang_tao_hoa_don_bangsoat():
    st.header("🧾 Tạo hóa đơn từ Bảng soát hóa đơn")

    # Thuế suất CỐ ĐỊNH 8% (không lấy từ trang Cài đặt nữa vì đã ẩn Cài đặt).
    thue = 8
    httt = db.get_hinh_thuc_thanh_toan()  # hình thức thanh toán mặc định (TM/CK) lấy sẵn từ DB

    # Bước 1: chọn file + ngày
    c1, c2 = st.columns([2, 1])
    file = c1.file_uploader("Chọn file 'Bảng soát hóa đơn' (.xlsx)", type=["xlsx"], key="bs_upload")
    ngay = c2.date_input("Ngày checkout của hóa đơn", value=datetime.date.today(), key="bs_ngay")

    # Luôn hiển thị nút tải file mẫu để user tải về điền dữ liệu
    st.download_button(
        "⬇️ Tải file mẫu để điền dữ liệu",
        data=bsr.tao_file_mau_bytes(),
        file_name="Bang_soat_hoa_don_MAU.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="bs_taimau",
    )
    st.caption("Tải file mẫu, điền dữ liệu theo đúng các cột rồi tải lên ở ô bên trên.")

    if st.button("📥 Đọc file & tạo bảng hóa đơn", type="primary", key="bs_doc"):
        if file is None:
            st.error("Bạn chưa chọn file.")
        else:
            data, canh_bao = bsr.doc_file_bangsoat(file)
            for w in canh_bao:
                st.warning(w)
            ngay_str = ngay.strftime("%d/%m/%Y")
            bang = [L.tao_dong_tu_bangsoat(rec, ngay_str, httt) for rec in data]
            for d in bang:
                L.cap_nhat_tien_bangsoat(d, 3, thue)
            st.session_state.bang_bs = bang
            st.session_state.so_nhom_bs = 3
            st.session_state.prev_loai_bs = {
                d["_row_id"]: (d["Loại hóa đơn"], L.to_int(d.get("HH1_Số lượng"))) for d in bang
            }
            st.session_state.grid_key_bs += 1
            st.session_state.ngay_hoa_don_bs = ngay
            st.session_state.pop("file_bytes_bs", None)
            if bang:
                st.success(f"Đã tạo bảng với {len(bang)} dòng. Ngày hóa đơn = {ngay_str}.")
            else:
                st.info("Không có dòng nào được tạo từ file.")

    if not st.session_state.get("bang_bs"):
        st.info("Hãy chọn file 'Bảng soát hóa đơn' và bấm nút trên để bắt đầu.")
        return

    so_nhom = st.session_state.so_nhom_bs

    # Bước 2: bảng lưới
    st.markdown("### Bảng hóa đơn (sửa trực tiếp trên ô)")
    st.caption("Nhấp đúp để sửa. Tab Tiền phòng: 'Thành tiền' sửa được, 'Đơn giá' tự tính "
               "= Thành tiền/Số lượng (2 chữ số thập phân). Tick checkbox để xóa/gộp.")

    df = pd.DataFrame(st.session_state.bang_bs)
    df = df.reindex(columns=danh_sach_cot_bs(so_nhom), fill_value="")
    df = lam_sach_nan(df)
    # Ép mọi ô về chuỗi để tránh lỗi tuần tự hóa Arrow khi một cột vừa có số vừa có ô
    # trống (vd cột "Số lượng" có dòng = 4, dòng khác để trống). Dữ liệu thật trong
    # session vẫn giữ nguyên kiểu số; đây chỉ là bản CHUỖI để hiển thị trên bảng.
    df = df.astype(str).replace({"None": "", "nan": "", "NaN": "", "<NA>": ""})

    css_bang = {
        ".ag-header-cell-label": {"justify-content": "center"},
        ".ag-header-cell-text": {"text-align": "center", "font-size": "15px"},
        ".ag-header-group-cell-label": {"justify-content": "center"},
        ".ag-header-group-text": {"text-align": "center", "width": "100%",
                                  "font-size": "15px", "font-weight": "700"},
        ".ag-cell": {"font-size": "15px"},
    }

    grid_response = AgGrid(
        df,
        gridOptions=tao_grid_options_bs(so_nhom),
        update_on=["cellValueChanged", "selectionChanged"],
        data_return_mode=DataReturnMode.AS_INPUT,
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=False,
        custom_css=css_bang,
        height=750,
        key=f"grid_bs_{st.session_state.grid_key_bs}",
    )

    edited_df = lam_sach_nan(pd.DataFrame(grid_response["data"]))
    can_lam_moi = xu_ly_sau_khi_sua_bs(edited_df.to_dict("records"), so_nhom, thue)

    selected = grid_response.get("selected_rows")
    if selected is None:
        selected_ids = []
    elif isinstance(selected, list):
        selected_ids = [r.get("_row_id") for r in selected]
    else:
        selected_ids = selected["_row_id"].tolist() if "_row_id" in selected.columns else []

    b1, b2, b3, b4, b5 = st.columns(5)
    if b1.button("➕ Thêm dòng trống", key="bs_them"):
        them_dong_trong_bs(so_nhom, httt)
        st.rerun()
    if b2.button("🗑️ Xóa dòng đã chọn", key="bs_xoa"):
        if selected_ids:
            st.session_state.bang_bs = [
                r for r in st.session_state.bang_bs if r.get("_row_id") not in selected_ids
            ]
            st.session_state.grid_key_bs += 1
            st.rerun()
        else:
            st.warning("Chưa chọn dòng nào để xóa.")
    if b3.button("🔀 Gộp dòng đã chọn", key="bs_gop"):
        if len(selected_ids) >= 2:
            hop_thoai_gop_bs(selected_ids, so_nhom, httt)
        else:
            st.warning("Cần chọn ít nhất 2 dòng để gộp.")
    if b4.button("➕ Thêm tab HÀNG HÓA DỊCH VỤ", key="bs_themtab"):
        them_tab_bs(so_nhom)
        st.rerun()
    if b5.button("🔄 Tính lại bảng", key="bs_tinhlai"):
        st.session_state.grid_key_bs += 1
        st.rerun()

    if can_lam_moi:
        st.session_state.grid_key_bs += 1
        st.rerun()

    # Bước 3: kiểm tra & xuất file (KHÔNG popup, KHÔNG cộng dồn số lượng)
    st.markdown("### Xuất file")
    loi = []
    for i, row in enumerate(st.session_state.bang_bs, start=1):
        loi += L.kiem_tra_dong(row, so_nhom, i)
        # Kiểm tra Mã số thuế: phải để trống hoặc đúng 10/13 chữ số (đếm bỏ gạch ngang)
        if not L.mst_hop_le(row.get("Mã số thuế")):
            loi.append(
                f"Dòng {i} (phòng {row.get('Phòng', '')}): Mã số thuế "
                f"'{row.get('Mã số thuế', '')}' không hợp lệ — phải có 10 hoặc 13 chữ số. "
                f"Vui lòng sửa lại trong file hoặc ngay trên bảng."
            )
        # Kiểm tra CCCD: phải để trống hoặc đúng 12 chữ số
        if not L.cccd_hop_le(row.get("CCCD")):
            loi.append(
                f"Dòng {i} (phòng {row.get('Phòng', '')}): CCCD "
                f"'{row.get('CCCD', '')}' không hợp lệ — phải có đúng 12 chữ số. "
                f"Vui lòng sửa lại trong file hoặc ngay trên bảng."
            )

    if loi:
        st.error("Còn lỗi cần sửa trước khi xuất file:")
        for e in loi:
            st.write("• " + e)
    else:
        if st.button("📤 Xuất file danh sách hóa đơn", type="primary", key="bs_xuat"):
            ngay_str = st.session_state.ngay_hoa_don_bs.strftime("%d/%m/%Y")
            # Thêm tiền tố "Đồ uống " cho cột Tên đồ uống khác khi xuất file (app vẫn giữ nguyên)
            bang_xuat = L.chuan_bi_xuat_bangsoat(st.session_state.bang_bs, so_nhom)
            data = ex.xuat_file_bytes(bang_xuat, so_nhom, ngay_str, thue,
                                      tron_tong=True, don_gia_chua_thue=True)
            st.session_state.file_bytes_bs = data
            st.session_state.file_name_bs = f"DanhSachHoaDon_BangSoat_{ngay_str.replace('/', '_')}.xlsx"
            st.success("Đã tạo file.")

    if st.session_state.get("file_bytes_bs"):
        st.download_button(
            "⬇️ Tải file Excel về máy",
            data=st.session_state.file_bytes_bs,
            file_name=st.session_state.file_name_bs,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="bs_download",
        )


# =========================================================================
# HÀM MAIN
# =========================================================================

def main():
    st.set_page_config(page_title="Tạo hóa đơn điện tử", layout="wide")
    db.init_db()  # tạo DB + nạp mặc định nếu cần (gọi nhiều lần vẫn an toàn)

    # Khởi tạo các biến phiên lần đầu
    if "grid_key" not in st.session_state:
        st.session_state.grid_key = 0
    if "so_nhom" not in st.session_state:
        st.session_state.so_nhom = 3
    if "prev_loai" not in st.session_state:
        st.session_state.prev_loai = {}
    # Biến phiên RIÊNG cho tính năng Bảng soát (tách biệt với tính năng ezcloud)
    if "grid_key_bs" not in st.session_state:
        st.session_state.grid_key_bs = 0
    if "so_nhom_bs" not in st.session_state:
        st.session_state.so_nhom_bs = 3
    if "prev_loai_bs" not in st.session_state:
        st.session_state.prev_loai_bs = {}

    # Thanh điều hướng bên trái
    st.sidebar.caption("App tạo hóa đơn điện tử hàng ngày")

    # ---------------------------------------------------------------------
    # HIỆN TẠI: chỉ dùng tính năng "Tạo hóa đơn từ Bảng soát".
    # Đã ẨN 2 tab "Tạo hóa đơn (ezcloud)" và "Cài đặt" khỏi sidebar
    # (CHỈ ẩn giao diện, code phía sau vẫn giữ để sau này dùng lại).
    #
    # >>> Muốn bật lại 2 tab cũ: bỏ comment đoạn dưới đây và xóa dòng
    #     "trang_tao_hoa_don_bangsoat()" ở cuối hàm.
    #
    # trang = st.sidebar.radio(
    #     "Chọn trang",
    #     ["Tạo hóa đơn (ezcloud)", "Tạo hóa đơn từ Bảng soát", "Cài đặt"],
    # )
    # if trang == "Tạo hóa đơn (ezcloud)":
    #     trang_tao_hoa_don()
    # elif trang == "Tạo hóa đơn từ Bảng soát":
    #     trang_tao_hoa_don_bangsoat()
    # else:
    #     settings_ui.hien_thi_cai_dat()
    # ---------------------------------------------------------------------

    trang_tao_hoa_don_bangsoat()


if __name__ == "__main__":
    main()
