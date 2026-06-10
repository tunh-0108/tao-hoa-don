"""
invoice_logic.py
----------------
Toàn bộ "luật tính toán" của bảng hóa đơn nằm ở đây (tách khỏi giao diện cho dễ debug):
  - Quyết định Loại hóa đơn mặc định theo Công ty.
  - Tạo 1 dòng hóa đơn ban đầu từ 1 booking (điền sẵn các giá trị mặc định).
  - Điền lại các tab HÀNG HÓA DỊCH VỤ khi user đổi "Loại hóa đơn".
  - Tính Thành tiền, các cột tổng.
  - Kiểm tra hợp lệ (validate) trước khi xuất file.

LƯU Ý VỀ "TAB" / "NHÓM HÀNG HÓA DỊCH VỤ":
  Mỗi nhóm có 5 ô: Tên / Đơn vị / Số lượng / Đơn giá / Thành tiền.
  Trong dữ liệu, các ô này được đặt tên là: HH{g}_Tên, HH{g}_Đơn vị, HH{g}_Số lượng, HH{g}_Đơn giá, HH{g}_Thành tiền
  với g = 1 (Tiền phòng), 2 (Dasani), 3 (Đồ uống khác), 4+ (tab tự thêm).
"""

import math
import random
import uuid
from datetime import datetime, timedelta

# ---- 3 giá trị của "Loại hóa đơn" (đặt thành hằng số để tránh gõ sai chính tả) ----
LOAI_CO_BAN = "Lưu trú cơ bản"
LOAI_CHUYEN_GIA_KO = "Lưu trú chuyên gia - minibar ko cần theo thực tế"
LOAI_CHUYEN_GIA_PHAI = "Lưu trú chuyên gia - minibar phải theo thực tế"
DANH_SACH_LOAI = [LOAI_CO_BAN, LOAI_CHUYEN_GIA_KO, LOAI_CHUYEN_GIA_PHAI]

# Tên các ô cố định (không thuộc nhóm hàng hóa) hiển thị trên bảng
COT_THONG_TIN = [
    "Mã đặt phòng",
    "Phòng",
    "Họ tên người mua hàng",
    "Loại hóa đơn",
    "Tên đơn vị mua hàng",
    "Mã số thuế",
    "Địa chỉ",
    "Hình thức thanh toán",
]

# Các ô ẩn (không hiện trên bảng nhưng cần để tính toán lại khi đổi loại hóa đơn)
COT_AN = ["_row_id", "Công ty", "Ngày đến", "Ngày đi", "Số đêm", "Giá trung bình"]

# Các cột tổng (chỉ hiển thị trên bảng, không xuất ra file).
# (Đã bỏ "Tiền thuế" và "Tổng thanh toán" khỏi bảng theo yêu cầu; chỉ giữ "Tổng tiền hàng".)
COT_TONG = ["Tổng tiền hàng"]

# 5 hậu tố của 1 nhóm hàng hóa
HAU_TO_NHOM = ["Tên", "Đơn vị", "Số lượng", "Đơn giá", "Thành tiền"]


# =========================================================================
# CÁC HÀM TIỆN ÍCH NHỎ
# =========================================================================

def la_rong(value):
    """True nếu ô được coi là RỖNG (None hoặc chuỗi trắng)."""
    return value is None or str(value).strip() == ""


def to_int(value):
    """
    Đổi 1 giá trị bất kỳ về số nguyên. Rỗng/không hợp lệ -> 0.
    Chấp nhận cả "1.0", " 2 ", 3.0, ...
    """
    if la_rong(value):
        return 0
    try:
        return int(round(float(str(value).strip())))
    except (TypeError, ValueError):
        return 0


def lam_tron_len(value):
    """Làm tròn LÊN thành số nguyên (theo yêu cầu)."""
    if la_rong(value):
        return 0
    try:
        return int(math.ceil(float(value)))
    except (TypeError, ValueError):
        return 0


def lam_tron_thuong(value):
    """
    Làm tròn THƯỜNG (về số nguyên gần nhất). Quy ước 0.5 thì làm tròn LÊN.
    Ví dụ: 257201.85 -> 257202 ; 100.4 -> 100 ; 100.5 -> 101.
    """
    if la_rong(value):
        return 0
    try:
        return int(math.floor(float(value) + 0.5))
    except (TypeError, ValueError):
        return 0


def format_ngay(dt):
    """Đổi datetime -> chuỗi 'dd/mm/yyyy' (chỉ lấy ngày/tháng/năm). None -> ''."""
    if dt is None:
        return ""
    return dt.strftime("%d/%m/%Y")


def key_nhom(g, hau_to):
    """Tạo tên ô của nhóm g. Ví dụ key_nhom(1, 'Tên') -> 'HH1_Tên'."""
    return f"HH{g}_{hau_to}"


# =========================================================================
# QUYẾT ĐỊNH LOẠI HÓA ĐƠN MẶC ĐỊNH
# =========================================================================

def loai_hoa_don_mac_dinh(cong_ty, danh_sach_chuyen_gia):
    """
    Công ty nằm trong danh sách chuyên gia (booking, ctrip, + setting) -> chuyên gia (ko cần theo thực tế).
    Ngược lại (kể cả công ty rỗng) -> Lưu trú cơ bản.
    """
    if (cong_ty or "").strip().lower() in danh_sach_chuyen_gia:
        return LOAI_CHUYEN_GIA_KO
    return LOAI_CO_BAN


def don_gia_co_ban(ten_phong, room_price_map, fallback):
    """
    Tra đơn giá cơ bản theo tên phòng:
      - Nếu phòng nằm trong 1 list (room_price_map) -> dùng giá của list đó.
      - Nếu không -> dùng fallback (mặc định 277778).
    """
    ten_phong = str(ten_phong or "").strip()
    return room_price_map.get(ten_phong, fallback)


def chon_do_uong_thap_nhat(drinks):
    """
    Chọn 1 đồ uống có 'da_xuat' (số lượng đã xuất) THẤP NHẤT.
    Nếu nhiều loại bằng nhau -> chọn ngẫu nhiên 1 trong số đó.
    Trả về dict đồ uống, hoặc None nếu danh sách rỗng.
    """
    if not drinks:
        return None
    min_da_xuat = min(d["da_xuat"] for d in drinks)
    ung_vien = [d for d in drinks if d["da_xuat"] == min_da_xuat]
    return random.choice(ung_vien)


# =========================================================================
# TẠO 1 DÒNG HÓA ĐƠN BAN ĐẦU
# =========================================================================

def tao_dong_tu_booking(booking, settings, room_price_map, drinks):
    """
    Tạo 1 dòng hóa đơn (dict) từ 1 booking đã lọc ở bước đọc file.

    settings: dict gồm
        'chuyen_gia_companies', 'fallback', 'hinh_thuc_thanh_toan'
    """
    loai = loai_hoa_don_mac_dinh(booking["Công ty"], settings["chuyen_gia_companies"])

    dong = {
        # ----- ô ẩn để tính toán -----
        "_row_id": str(uuid.uuid4()),  # id duy nhất cho mỗi dòng (dùng khi xóa/gộp)
        "Công ty": booking["Công ty"],
        "Ngày đến": format_ngay(booking["Ngày đến"]),
        "Ngày đi": format_ngay(booking["Ngày đi"]),
        "Số đêm": to_int(booking["Số đêm"]),
        "Giá trung bình": booking["Giá trung bình"],
        # ----- ô thông tin -----
        "Mã đặt phòng": booking["Mã đặt phòng"],
        "Phòng": booking["Tên phòng"],
        "Họ tên người mua hàng": booking["Tên khách"],
        "Loại hóa đơn": loai,
        "Tên đơn vị mua hàng": "",
        "Mã số thuế": "",
        "Địa chỉ": "",
        "Hình thức thanh toán": settings["hinh_thuc_thanh_toan"],
    }

    # Tạo sẵn 3 nhóm hàng hóa với ô rỗng, rồi điền mặc định theo loại hóa đơn
    for g in (1, 2, 3):
        for hau_to in HAU_TO_NHOM:
            dong[key_nhom(g, hau_to)] = ""

    dien_mac_dinh_cac_nhom(dong, room_price_map, settings["fallback"], drinks)
    return dong


def dien_mac_dinh_cac_nhom(dong, room_price_map, fallback, drinks):
    """
    Điền giá trị mặc định cho NHÓM 1, 2, 3 dựa theo 'Loại hóa đơn' của dòng.
    Hàm này được gọi: (1) khi tạo dòng mới, (2) khi user ĐỔI loại hóa đơn.
    -> Đổi loại hóa đơn sẽ reset lại các tab về mặc định của loại mới.
    """
    loai = dong.get("Loại hóa đơn")
    nd = dong.get("Ngày đến", "")
    ndi = dong.get("Ngày đi", "")
    so_dem = to_int(dong.get("Số đêm"))

    # ---------------- NHÓM 1: TIỀN PHÒNG ----------------
    if loai in (LOAI_CHUYEN_GIA_KO, LOAI_CHUYEN_GIA_PHAI):
        dong["HH1_Tên"] = f"Dịch vụ lưu trú chuyên gia từ ngày {nd} đến ngày {ndi}"
        dong["HH1_Đơn giá"] = lam_tron_len(dong.get("Giá trung bình"))
    else:  # Lưu trú cơ bản
        dong["HH1_Tên"] = f"Dịch vụ lưu trú cơ bản từ ngày {nd} đến ngày {ndi}"
        dong["HH1_Đơn giá"] = don_gia_co_ban(dong.get("Phòng"), room_price_map, fallback)
    dong["HH1_Đơn vị"] = "Đêm"
    dong["HH1_Số lượng"] = so_dem

    # ---------------- NHÓM 2: DASANI ----------------
    if loai in (LOAI_CHUYEN_GIA_KO, LOAI_CO_BAN):
        dong["HH2_Tên"] = "Nước uống Dasani"
        dong["HH2_Đơn vị"] = "Chai"
        dong["HH2_Số lượng"] = 2
        dong["HH2_Đơn giá"] = 0
    else:  # phải theo thực tế -> để trống cho user tự nhập
        dong["HH2_Tên"] = ""
        dong["HH2_Đơn vị"] = ""
        dong["HH2_Số lượng"] = ""
        dong["HH2_Đơn giá"] = ""

    # ---------------- NHÓM 3: ĐỒ UỐNG KHÁC ----------------
    if loai == LOAI_CHUYEN_GIA_KO:
        do_uong = chon_do_uong_thap_nhat(drinks)
        if do_uong is not None:
            dong["HH3_Tên"] = do_uong["ten"]
            dong["HH3_Đơn vị"] = "Chai"
            dong["HH3_Số lượng"] = random.choice([1, 2])  # random 1 hoặc 2
            dong["HH3_Đơn giá"] = 0
        else:
            # Không có đồ uống nào trong setting -> để trống
            dong["HH3_Tên"] = ""
            dong["HH3_Đơn vị"] = ""
            dong["HH3_Số lượng"] = ""
            dong["HH3_Đơn giá"] = ""
    else:
        # Cơ bản hoặc phải-theo-thực-tế -> để trống (phải-theo-thực-tế user tự chọn)
        dong["HH3_Tên"] = ""
        dong["HH3_Đơn vị"] = ""
        dong["HH3_Số lượng"] = ""
        dong["HH3_Đơn giá"] = ""

    # Tính lại Thành tiền cho 3 nhóm vừa điền
    for g in (1, 2, 3):
        cap_nhat_thanh_tien_nhom(dong, g)


# =========================================================================
# TÍNH THÀNH TIỀN & CÁC CỘT TỔNG
# =========================================================================

def cap_nhat_thanh_tien_nhom(dong, g):
    """Thành tiền của nhóm g = Đơn giá * Số lượng. Nếu nhóm rỗng -> để trống."""
    don_gia = dong.get(key_nhom(g, "Đơn giá"))
    so_luong = dong.get(key_nhom(g, "Số lượng"))
    if la_rong(don_gia) and la_rong(so_luong):
        dong[key_nhom(g, "Thành tiền")] = ""
    else:
        dong[key_nhom(g, "Thành tiền")] = to_int(don_gia) * to_int(so_luong)


def cap_nhat_tat_ca_tien(dong, so_nhom, thue_suat):
    """
    Tính lại toàn bộ tiền cho 1 dòng:
      - Thành tiền từng nhóm.
      - Tổng tiền hàng = cộng Thành tiền các nhóm.
      - Tiền thuế = làm tròn lên (Tổng tiền hàng * thuế suất / 100).
      - Tổng thanh toán = Tổng tiền hàng + Tiền thuế.
    """
    tong_tien_hang = 0
    for g in range(1, so_nhom + 1):
        cap_nhat_thanh_tien_nhom(dong, g)
        tt = dong.get(key_nhom(g, "Thành tiền"))
        if not la_rong(tt):
            tong_tien_hang += to_int(tt)

    tien_thue = lam_tron_len(tong_tien_hang * thue_suat / 100)
    dong["Tổng tiền hàng"] = tong_tien_hang
    dong["Tiền thuế"] = tien_thue
    dong["Tổng thanh toán"] = tong_tien_hang + tien_thue


# =========================================================================
# KIỂM TRA HỢP LỆ (VALIDATE)
# =========================================================================

def kiem_tra_dong(dong, so_nhom, so_thu_tu):
    """
    Kiểm tra 1 dòng. Quy tắc: với MỖI nhóm hàng hóa, nếu 1 trong các ô
    Tên/Đơn vị/Số lượng/Đơn giá khác rỗng thì TẤT CẢ phải khác rỗng.
    Trả về danh sách thông báo lỗi (rỗng nếu hợp lệ).
    """
    loi = []
    for g in range(1, so_nhom + 1):
        cac_o = {
            ht: dong.get(key_nhom(g, ht))
            for ht in ["Tên", "Đơn vị", "Số lượng", "Đơn giá"]
        }
        co_o_dien = any(not la_rong(v) for v in cac_o.values())
        if co_o_dien:
            o_thieu = [ht for ht, v in cac_o.items() if la_rong(v)]
            if o_thieu:
                loi.append(
                    f"Dòng {so_thu_tu} - Nhóm HÀNG HÓA DỊCH VỤ {g}: "
                    f"còn thiếu ô {', '.join(o_thieu)}."
                )
    return loi


def nhom_co_du_lieu(dong, g):
    """True nếu nhóm g có ít nhất 1 ô (Tên/Đơn vị/Số lượng/Đơn giá) khác rỗng."""
    return any(
        not la_rong(dong.get(key_nhom(g, ht)))
        for ht in ["Tên", "Đơn vị", "Số lượng", "Đơn giá"]
    )


# =========================================================================
# =========================================================================
# PHẦN DÀNH RIÊNG CHO TÍNH NĂNG "BẢNG SOÁT HÓA ĐƠN"
# (Tách riêng hoàn toàn, KHÔNG dùng chung logic điền mặc định với tính năng ezcloud)
# =========================================================================
# =========================================================================

# Loại hóa đơn cho tính năng Bảng soát: CHỈ 2 giá trị.
LOAI_BS_CHUYEN_GIA = "Lưu trú chuyên gia"
LOAI_BS_CO_BAN = "Lưu trú cơ bản"
DANH_SACH_LOAI_BS = [LOAI_BS_CHUYEN_GIA, LOAI_BS_CO_BAN]

# Cột thông tin của bảng Bảng soát: GIỐNG bảng cũ nhưng THÊM cột "Email" cạnh "Địa chỉ".
COT_THONG_TIN_BS = [
    "Mã đặt phòng",
    "Phòng",
    "Họ tên người mua hàng",
    "CCCD/PASSPORT",
    "Loại hóa đơn",
    "Tên đơn vị mua hàng",
    "Mã số thuế",
    "Địa chỉ",
    "Email",
    "Hình thức thanh toán",
]

# Cột ẩn của bảng Bảng soát: lưu id dòng và ngày hóa đơn (để tính Tên tab Tiền phòng).
COT_AN_BS = ["_row_id", "Ngày hóa đơn"]


def so_hoac_rong(value):
    """
    Dùng cho các ô SỐ lấy từ file Bảng soát: nếu file để TRỐNG thì giữ TRỐNG (''),
    có giá trị thì đổi sang số nguyên. Tránh việc ô trống bị biến thành 0.
    """
    if la_rong(value):
        return ""
    return to_int(value)


def doc_tien(value):
    """
    Đọc số TIỀN từ file, chấp nhận nhiều cách viết:
      "300.000 đ"  -> 300000
      "300.000"    -> 300000
      "300000"     -> 300000
      300000 / 300000.0 (số) -> 300000
    Cách làm: bỏ chữ tiền tệ (đ/đồng/vnđ), bỏ khoảng trắng, bỏ dấu chấm và dấu phẩy
    (coi là dấu phân cách nghìn vì tiền phòng là số nguyên đồng).
    File để TRỐNG thì trả về '' (giữ ô trống).
    """
    if la_rong(value):
        return ""
    # Nếu đã là số sẵn (Excel đọc ra dạng số)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(round(value))
    s = str(value).strip().lower()
    for tu in ["đồng", "vnđ", "vnd", "đ"]:
        s = s.replace(tu, "")
    # Bỏ khoảng trắng (kể cả khoảng trắng đặc biệt) và dấu phân cách nghìn
    s = s.replace(" ", "").replace("\u00a0", "")
    s = s.replace(".", "").replace(",", "")
    if s == "":
        return ""
    try:
        return int(s)
    except ValueError:
        # Phòng trường hợp còn ký tự lạ: chỉ giữ các chữ số
        chu_so = "".join(ch for ch in s if ch.isdigit())
        return int(chu_so) if chu_so else ""


def truncate_2_so_le(value):
    """
    Giữ tối đa 2 chữ số thập phân, KHÔNG làm tròn (cắt bớt - truncate).
    Ví dụ: 92592.6666 -> 92592.66 ; 277778 -> 277778.0
    """
    if la_rong(value):
        return 0.0
    try:
        return math.floor(float(value) * 100) / 100.0
    except (TypeError, ValueError):
        return 0.0


def ten_tien_phong_bangsoat(loai, ngay_hoa_don_str, so_dem):
    """
    Tạo chữ cho ô Tên ở tab "HÀNG HÓA, DỊCH VỤ 1 - Tiền phòng" của bảng Bảng soát.
    "từ ngày {ngày hóa đơn - số đêm} đến ngày {ngày hóa đơn}" (chỉ ngày/tháng/năm).
    Việc trừ số đêm là trừ theo ngày lịch thật (có tính tháng/năm).
    """
    loai_text = "chuyên gia" if loai == LOAI_BS_CHUYEN_GIA else "cơ bản"
    try:
        ngay_den = datetime.strptime(str(ngay_hoa_don_str), "%d/%m/%Y").date()
        ngay_di = ngay_den - timedelta(days=to_int(so_dem))
        tu = ngay_di.strftime("%d/%m/%Y")
        den = ngay_den.strftime("%d/%m/%Y")
    except (TypeError, ValueError):
        tu, den = "", str(ngay_hoa_don_str)
    return f"Dịch vụ lưu trú {loai_text} từ ngày {tu} đến ngày {den}"


def tao_dong_tu_bangsoat(rec, ngay_hoa_don_str, hinh_thuc_thanh_toan):
    """
    Tạo 1 dòng hóa đơn (dict) cho bảng Bảng soát từ 1 dòng dữ liệu file (rec).

    rec chứa các khóa (do bangsoat_reader trả về):
      Phòng, Họ tên người mua hàng, CCCD/PASSPORT, Loại hóa đơn, Số đêm, Số lượng Dasani,
      Tên đồ uống khác, Số lượng đồ uống khác, Tiền phòng,
      Tên đơn vị mua hàng, Mã số thuế, Địa chỉ, Email
    ngay_hoa_don_str: chuỗi 'dd/mm/yyyy' = ngày checkout của hóa đơn (chung cho mọi dòng).

    Quy tắc: cột nào trong file để TRỐNG thì ô tương ứng trên app cũng để TRỐNG.
    """
    dong = {
        "_row_id": str(uuid.uuid4()),
        "Ngày hóa đơn": ngay_hoa_don_str,
        "Mã đặt phòng": "",  # file Bảng soát không có mã đặt phòng
        "Phòng": rec.get("Phòng", ""),
        "Họ tên người mua hàng": rec.get("Họ tên người mua hàng", ""),
        "CCCD/PASSPORT": rec.get("CCCD/PASSPORT", ""),
        "Loại hóa đơn": rec.get("Loại hóa đơn", LOAI_BS_CO_BAN),
        "Tên đơn vị mua hàng": rec.get("Tên đơn vị mua hàng", ""),
        "Mã số thuế": rec.get("Mã số thuế", ""),
        "Địa chỉ": rec.get("Địa chỉ", ""),
        "Email": rec.get("Email", ""),
        "Hình thức thanh toán": hinh_thuc_thanh_toan,
    }

    # Khởi tạo 3 nhóm hàng hóa với ô rỗng
    for g in (1, 2, 3):
        for hau_to in HAU_TO_NHOM:
            dong[key_nhom(g, hau_to)] = ""

    # ----- HH1: Tiền phòng -----
    # Số đêm và Tiền phòng: file trống thì để trống (không thành 0).
    so_dem = so_hoac_rong(rec.get("Số đêm"))
    dong["HH1_Tên"] = ten_tien_phong_bangsoat(dong["Loại hóa đơn"], ngay_hoa_don_str, so_dem)
    dong["HH1_Đơn vị"] = "Đêm"
    dong["HH1_Số lượng"] = so_dem
    dong["HH1_Thành tiền"] = doc_tien(rec.get("Tiền phòng"))  # đổ thẳng từ file (đọc được "300.000 đ")

    # ----- HH2: Dasani (chỉ điền nếu file có Số lượng Dasani) -----
    if not la_rong(rec.get("Số lượng Dasani")):
        dong["HH2_Tên"] = "Nước uống Dasani"
        dong["HH2_Đơn vị"] = "Chai"
        dong["HH2_Số lượng"] = so_hoac_rong(rec.get("Số lượng Dasani"))
        dong["HH2_Đơn giá"] = 0

    # ----- HH3: Đồ uống khác (chỉ điền nếu file có Tên đồ uống khác) -----
    if not la_rong(rec.get("Tên đồ uống khác")):
        dong["HH3_Tên"] = str(rec.get("Tên đồ uống khác")).strip()
        dong["HH3_Đơn vị"] = "Lon"
        dong["HH3_Số lượng"] = so_hoac_rong(rec.get("Số lượng đồ uống khác"))
        dong["HH3_Đơn giá"] = 0  # luôn = 0 để user tự nhập, bất kể setting giá

    # Tính các con số phụ thuộc
    cap_nhat_tien_bangsoat(dong, 3, 0)  # thuế suất không cần ở bước tạo
    return dong


def cap_nhat_tien_bangsoat(dong, so_nhom, thue_suat):
    """
    Tính lại tiền cho 1 dòng bảng Bảng soát. KHÁC với tính năng ezcloud:
      - HH1 (Tiền phòng): "Thành tiền" là số NHẬP TAY (đổ từ file).
        Đơn giá = Thành tiền / Số lượng, giữ 2 chữ số thập phân (KHÔNG làm tròn).
      - Các nhóm khác (HH2, HH3, tab tự thêm): Thành tiền = Đơn giá * Số lượng (như cũ).
      - Tổng tiền hàng = cộng Thành tiền tất cả nhóm.
    """
    # HH1: tính Đơn giá = Thành tiền / Số lượng (2 chữ số thập phân, không làm tròn).
    # Nếu thiếu Thành tiền HOẶC Số lượng thì Đơn giá để TRỐNG (không đặt 0).
    sl1 = to_int(dong.get("HH1_Số lượng"))
    tt1 = dong.get("HH1_Thành tiền")
    if not la_rong(tt1) and sl1 > 0:
        dong["HH1_Đơn giá"] = truncate_2_so_le(float(to_int(tt1)) / sl1)
    else:
        dong["HH1_Đơn giá"] = ""

    # Các nhóm 2..so_nhom: Thành tiền = Đơn giá * Số lượng
    for g in range(2, so_nhom + 1):
        cap_nhat_thanh_tien_nhom(dong, g)

    # Tổng tiền hàng
    tong = 0
    for g in range(1, so_nhom + 1):
        tt = dong.get(key_nhom(g, "Thành tiền"))
        if not la_rong(tt):
            tong += to_int(tt)
    dong["Tổng tiền hàng"] = tong


def chuan_bi_xuat_bangsoat(bang):
    """
    Trả về BẢN SAO của bảng để XUẤT FILE (không làm đổi dữ liệu hiển thị trên app).
    Việc duy nhất: thêm tiền tố 'Đồ uống ' vào trước ô Tên của nhóm
    'Đồ uống khác' (HH3). Ví dụ: 'Coffee' -> 'Đồ uống Coffee'.
    (Nếu ô trống thì để nguyên; nếu đã có sẵn tiền tố thì không thêm lần nữa.)
    """
    tien_to = "Đồ uống "
    ban_sao = []
    for d in bang:
        d2 = dict(d)  # sao chép nông là đủ vì mỗi ô là giá trị đơn
        ten = d2.get("HH3_Tên")
        if not la_rong(ten):
            ten = str(ten).strip()
            if not ten.startswith(tien_to):
                d2["HH3_Tên"] = tien_to + ten
        ban_sao.append(d2)
    return ban_sao
