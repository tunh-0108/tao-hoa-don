# App tạo hóa đơn điện tử hàng ngày (iRest Signature)

App đọc file doanh thu chi tiết của **ezcloud**, lọc các phòng **CHECKOUT** trong
ngày đã chọn, dựng **bảng hóa đơn** cho bạn chỉnh sửa, rồi xuất ra file Excel **41 cột**
đúng định dạng import của **VinInvoice** (theo file mẫu `MauUploadHDMTT`).

---

## 1. Cài đặt (chỉ làm 1 lần)

Cần có **Python 3.9+**. Mở terminal trong thư mục này và chạy:

```bash
pip install -r requirements.txt
```

## 2. Chạy app

```bash
streamlit run app.py
```

Trình duyệt sẽ tự mở (thường ở địa chỉ http://localhost:8501).

## 3. Cách dùng

**Trang "Tạo hóa đơn":**
1. Bấm **Chọn file** và chọn file doanh thu chi tiết (.xlsx) từ ezcloud.
2. Chọn **Ngày checkout cần lọc** (mặc định = hôm nay).
3. Bấm **📥 Đọc file & tạo bảng hóa đơn**.
4. Sửa trực tiếp trên bảng:
   - Nhấp đúp vào ô để sửa. Đổi **Loại hóa đơn** (dropdown) sẽ tự điền lại các tab.
   - Cột **Thành tiền** và các cột **Tổng** tự tính, không sửa được.
   - Tick **checkbox** ở cột đầu để chọn dòng → **Xóa** hoặc **Gộp**.
   - **➕ Thêm dòng trống**, **➕ Thêm tab HÀNG HÓA DỊCH VỤ** nếu cần.
   - **🔄 Tính lại bảng**: bấm nếu thấy số tổng chưa cập nhật kịp.
5. Bấm **📤 Xuất file danh sách hóa đơn** → chọn có/không cộng dồn "số lượng đã xuất"
   của Đồ uống khác → **⬇️ Tải file Excel về máy**.

**Trang "Cài đặt":** chỉnh thuế suất, đơn giá fallback, hình thức thanh toán, danh sách
công ty → chuyên gia, các list phòng → đơn giá, và bảng Đồ uống khác.

---

## 4. Cấu trúc code (để dễ sửa về sau)

| File | Nhiệm vụ |
|------|----------|
| `app.py` | File chính: điều hướng + trang Tạo hóa đơn + bảng lưới + xuất file |
| `settings_ui.py` | Giao diện trang Cài đặt |
| `database.py` | Cơ sở dữ liệu SQLite (settings, công ty, list phòng, đồ uống) |
| `excel_reader.py` | Đọc & lọc file ezcloud |
| `invoice_logic.py` | Toàn bộ "luật tính toán" (mặc định, validate, tính tiền) |
| `export_excel.py` | Tạo file output 41 cột |
| `hoadon.db` | File dữ liệu (tự sinh khi chạy lần đầu) — **đừng xóa nếu muốn giữ cài đặt** |

Dữ liệu **Cài đặt** và **số lượng đã xuất** được lưu vĩnh viễn trong `hoadon.db`.
Riêng **bảng hóa đơn đang làm** chỉ giữ trong phiên, sẽ mất khi tải lại trang (đúng yêu cầu).

---

## 5. Ghi chú về giao diện bảng

Bảng dùng thư viện **streamlit-aggrid** để có header gộp (các nhóm HÀNG HÓA DỊCH VỤ),
dropdown và checkbox giống ảnh mẫu. Thay cho "icon sửa/xóa khi rê chuột từng ô",
bảng cho phép **sửa trực tiếp** (nhấp đúp ô) và **xóa/gộp theo dòng** qua checkbox —
đây là cách thao tác lưới tiêu chuẩn, nhanh và ít lỗi hơn. Nếu sau này bạn muốn đúng
kiểu icon từng ô, có thể bổ sung, nhưng sẽ cần thêm mã JavaScript tùy biến.

## 6. Đưa lên GitHub / deploy nội bộ (về sau)

Chỉ cần đẩy cả thư mục này lên GitHub (nên thêm `hoadon.db` và `__pycache__/` vào
`.gitignore` nếu không muốn đưa dữ liệu lên). Có thể deploy bằng Streamlit Community
Cloud hoặc chạy `streamlit run app.py` trên máy chủ nội bộ.
