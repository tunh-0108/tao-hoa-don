"""
export_excel.py
---------------
Tạo file Excel OUTPUT để import vào VinInvoice.

QUAN TRỌNG: file output được tạo bằng cách NẠP THẲNG file mẫu thật (MauUploadHD_TT78)
làm nền, xóa các dòng ví dụ có sẵn rồi điền dữ liệu vào - KHÔNG tự dựng workbook mới
từ đầu. Lý do: VinInvoice kiểm tra "đúng mẫu hệ thống" trước khi nhận file (báo lỗi
ngay cả khi từng ô dữ liệu hợp lệ), nên file output cần giữ nguyên 100% cấu trúc gốc:
đủ 3 sheet (kể cả 2 sheet danh mục chỉ dùng cho dropdown), named range, data validation,
dòng 1 (mã field) ẩn, dòng 2 (nhãn tiếng Việt) - chỉ khác phần DỮ LIỆU ở sheet chính.

77 cột của sheet chính "Thông tin hoá đơn" (theo file mẫu). Phần lớn cột để trống, chỉ
điền những cột có logic mapping. Mỗi giá trị được ghi dạng CHUỖI (text) để tránh
VinInvoice import bị lệch kiểu dữ liệu.

Quy tắc gộp dòng (quan trọng):
  - Mỗi DÒNG trong bảng hóa đơn = 1 hóa đơn = 1 MaHD (đánh số tăng dần từ 1).
  - Trong 1 dòng, mỗi NHÓM hàng hóa có dữ liệu sẽ tạo ra 1 DÒNG output (cùng MaHD).
  - Các dòng output cùng MaHD chỉ khác nhau ở: TenHangHoa, DonViTinh, SoLuong, DonGia, ThanhTien.
"""

import io
import os
import openpyxl
from openpyxl.styles import Font

import invoice_logic as L

# File mẫu thật của VinInvoice, dùng làm NỀN cho file output (giữ nguyên 3 sheet,
# named range, data validation... để VinInvoice không báo "không đúng mẫu hệ thống").
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "MauUploadHD_TT78(1).xlsx")
TEMPLATE_SHEET = "Thông tin hoá đơn"
DATA_START_ROW = 3  # dòng 1 = mã field (ẩn), dòng 2 = nhãn tiếng Việt, data bắt đầu dòng 3

# 77 tên cột của file output, đúng thứ tự A -> BY (theo file mẫu MauUploadHD_TT78)
OUTPUT_HEADERS = [
    "MaHD", "LoaiHoaDon", "HoaDonLienQuanNgoaiHeThong", "MauSoHoaDonLienQuan",
    "KyHieuHoaDonLienQuan", "SoHoaDonLienQuan", "NgayHoaDonLienQuan", "MSTCLQuan",
    "LDDCTThe", "NgayHoaDon", "MaKhachHang", "TenNguoiMua", "TenDonVi",
    "MDVQHNSach", "MDDKDoanh", "TDDKDoanh", "DCDDKDoanh", "DiaChiKhachHang",
    "CusPhone", "MTinh", "TTinh", "MXa", "TXa", "MaSoThue", "CCCD", "SHChieu",
    "SDDTCNNgoai", "MQTNMua", "QTNMua", "MailKhachHang", "HinhThucThanhToan",
    "SoTaiKhoan", "TenTaiKhoan", "LoaiTien", "TyGia", "SoChungTu", "ProcessInvNote",
    "TCQDGBTSan", "DCCQDGBTSan", "MSTCQDGBTsan", "MCHang", "TCHang", "DCCHang",
    "LoaiHangHoa", "LoaiHangHoaDacTrung", "MaHangHoa", "TenHangHoa", "SoKhung",
    "SoMay", "TNHieu", "TTMai", "TLTSDKy", "LTSan", "TTTSan", "KLXe",
    "TTLVHCSuat", "TTai", "SCNgoi", "MNSXuat", "NSXuat", "XXu", "SBKSoat",
    "SGCNATKThuat", "SSPKTCLXXuong", "BKSPTVanChuyen", "TNGHang", "DCNGuiHang",
    "MSTNGuiHang", "SDDNguiHang", "GhiChu", "DonViTinh", "SoLuong", "DonGia",
    "ThanhTien", "ThueSuat", "TienThue", "TienBangChu",
]


def _chuoi(value):
    """Đổi giá trị về chuỗi để ghi vào Excel. None -> chuỗi rỗng."""
    if value is None:
        return ""
    return str(value)


def _format_don_gia(value):
    """
    Định dạng Đơn giá khi ghi ra file:
      - Số nguyên -> không có phần thập phân (vd 277778).
      - Có lẻ     -> giữ đúng 2 chữ số thập phân (vd 92592.66).
    (Cần cho tính năng Bảng soát vì Đơn giá HH1 = Thành tiền/Số lượng có thể lẻ.)
    """
    if value is None or str(value).strip() == "":
        return ""
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if f == int(f):
        return str(int(f))
    return f"{f:.2f}"


def tao_cac_dong_output(bang_hoa_don, so_nhom, ngay_hoa_don_str, thue_suat,
                        tron_tong=False, don_gia_chua_thue=False):
    """
    Từ bảng hóa đơn (list dict), tạo ra list các dòng output (mỗi dòng là 1 dict
    theo đúng tên cột OUTPUT_HEADERS).

    Tham số:
      - bang_hoa_don: list dict (mỗi dict là 1 dòng trên bảng hóa đơn)
      - so_nhom: tổng số nhóm hàng hóa hiện có (>= 3)
      - ngay_hoa_don_str: chuỗi 'dd/mm/yyyy' = ngày hôm nay
      - thue_suat: số nguyên (vd 8)
    """
    cac_dong = []
    ma_hd = 0  # sẽ tăng lên 1 cho mỗi dòng hóa đơn CÓ ít nhất 1 nhóm dữ liệu

    for dong in bang_hoa_don:
        # Tìm các nhóm có dữ liệu trong dòng này
        nhom_co_dl = [g for g in range(1, so_nhom + 1) if L.nhom_co_du_lieu(dong, g)]
        if not nhom_co_dl:
            continue  # dòng không có hàng hóa nào -> bỏ qua, không tốn MaHD

        ma_hd += 1

        # Phần thông tin chung (giống nhau cho mọi dòng output cùng MaHD)
        thong_tin_chung = {
            "MaHD": ma_hd,
            "LoaiHoaDon": "0",  # 0 = hóa đơn gốc (không điều chỉnh/thay thế)
            "NgayHoaDon": ngay_hoa_don_str,
            "TenNguoiMua": dong.get("Họ tên người mua hàng", ""),
            "TenDonVi": dong.get("Tên đơn vị mua hàng", ""),
            "DiaChiKhachHang": dong.get("Địa chỉ", ""),
            "MaSoThue": dong.get("Mã số thuế", ""),
            "CCCD": L.chuan_hoa_cccd(dong.get("CCCD", "")),  # cột CCCD (Bảng soát); ezcloud để rỗng
            "SHChieu": dong.get("PASSPORT", ""),  # số hộ chiếu (Bảng soát); ezcloud để rỗng
            "MailKhachHang": dong.get("Email", ""),  # cột Email (tính năng Bảng soát); ezcloud để rỗng
            "HinhThucThanhToan": dong.get("Hình thức thanh toán", ""),
            "LoaiTien": "VND",
            "TyGia": "1",
            "LoaiHangHoa": "1",
            "LoaiHangHoaDacTrung": "",  # dropdown file mới chỉ có 1/2/3, không còn "0" nên để trống
            "ThueSuat": thue_suat,
        }

        # Mỗi nhóm có dữ liệu -> 1 dòng output
        for g in nhom_co_dl:
            # "Thành tiền" trên bảng là giá ĐÃ GỒM thuế.
            # ThanhTien trong file output cần là giá CHƯA thuế:
            #   ThanhTien = Thành tiền / (1 + thuế/100), rồi làm tròn gần nhất.
            thanh_tien_gom_thue = L.to_int(dong.get(L.key_nhom(g, "Thành tiền")))
            gia_chua_thue = thanh_tien_gom_thue / (1 + thue_suat / 100)
            thanh_tien = L.lam_tron_thuong(gia_chua_thue)
            if tron_tong:
                # Lấy Tiền thuế = phần còn lại để ThanhTien + TienThue = đúng số tiền
                # gồm thuế ban đầu (tránh lệch 1 đồng kiểu 300001).
                tien_thue = thanh_tien_gom_thue - thanh_tien
            else:
                # TienThue tính trên giá CHƯA thuế (ThanhTien mới), làm tròn lên.
                tien_thue = L.lam_tron_len(thanh_tien * thue_suat / 100)

            dong_out = dict(thong_tin_chung)  # copy phần chung
            dong_out["TenHangHoa"] = dong.get(L.key_nhom(g, "Tên"), "")
            dong_out["DonViTinh"] = dong.get(L.key_nhom(g, "Đơn vị"), "")
            dong_out["SoLuong"] = dong.get(L.key_nhom(g, "Số lượng"), "")
            so_luong = L.to_int(dong.get(L.key_nhom(g, "Số lượng")))
            if don_gia_chua_thue and so_luong != 0:
                # Đơn giá trong file phải là đơn giá CHƯA thuế = ThanhTien (chưa thuế) / Số lượng.
                # (Trên bảng, Đơn giá là giá ĐÃ gồm thuế nên không dùng trực tiếp được.)
                dong_out["DonGia"] = _format_don_gia(thanh_tien / so_luong)
            else:
                dong_out["DonGia"] = _format_don_gia(dong.get(L.key_nhom(g, "Đơn giá"), ""))
            dong_out["ThanhTien"] = thanh_tien
            dong_out["TienThue"] = tien_thue
            cac_dong.append(dong_out)

    return cac_dong


def xuat_file_bytes(bang_hoa_don, so_nhom, ngay_hoa_don_str, thue_suat,
                    tron_tong=False, don_gia_chua_thue=False):
    """
    Tạo file Excel trong bộ nhớ và trả về dạng bytes (để Streamlit cho tải về).
    tron_tong=True: làm cho ThanhTien + TienThue = đúng số tiền gồm thuế (không lệch 1đ).
    don_gia_chua_thue=True: ghi Đơn giá trong file = ThanhTien (chưa thuế) / Số lượng.
    """
    cac_dong = tao_cac_dong_output(
        bang_hoa_don, so_nhom, ngay_hoa_don_str, thue_suat,
        tron_tong=tron_tong, don_gia_chua_thue=don_gia_chua_thue,
    )

    # Nạp THẲNG file mẫu thật làm nền (giữ nguyên đủ 3 sheet, named range, data
    # validation, dòng 1 ẩn, dòng 2 nhãn...) thay vì tự dựng workbook trắng.
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(
            f"Không tìm thấy file mẫu VinInvoice tại: {TEMPLATE_PATH}. "
            "File này bắt buộc phải có trong thư mục dự án để xuất đúng định dạng."
        )
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb[TEMPLATE_SHEET]

    font = Font(name="Arial")

    # Xóa các dòng dữ liệu VÍ DỤ có sẵn trong file mẫu (không được lẫn vào file thật),
    # chỉ giữ lại dòng 1 (mã field, ẩn) và dòng 2 (nhãn tiếng Việt).
    so_dong_vi_du = ws.max_row - (DATA_START_ROW - 1)
    if so_dong_vi_du > 0:
        ws.delete_rows(DATA_START_ROW, so_dong_vi_du)

    # Ghi dữ liệu thật từ dòng DATA_START_ROW
    for row_idx, dong_out in enumerate(cac_dong, start=DATA_START_ROW):
        for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
            gia_tri = _chuoi(dong_out.get(header, ""))
            c = ws.cell(row=row_idx, column=col_idx, value=gia_tri)
            c.font = font
            # Ép kiểu hiển thị là TEXT để các số không bị Excel tự định dạng
            c.number_format = "@"

    # Lưu vào bộ nhớ thay vì ghi ra ổ đĩa
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
