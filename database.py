"""
database.py
-----------
Tất cả những gì liên quan đến CƠ SỞ DỮ LIỆU (SQLite) nằm ở file này.

Trong app có 4 nhóm dữ liệu cần LƯU VĨNH VIỄN (không mất khi tắt app):
  1) settings        : các thông số chung (thuế suất, giá cơ bản fallback, hình thức thanh toán mặc định)
  2) chuyen_gia_company : danh sách "Công ty" sẽ mặc định ra loại "Lưu trú chuyên gia"
  3) room_price + room_price_member : các "list phòng -> đơn giá cơ bản"
  4) do_uong_khac    : bảng Đồ uống khác (tên, giá bán, số lượng đã xuất)

Cách dùng: ở đầu app gọi init_db() một lần. Sau đó dùng các hàm get_* / add_* / ...
Mỗi hàm tự mở và đóng kết nối nên rất dễ hiểu, không cần lo quản lý connection.
"""

import sqlite3
import os

# File DB sẽ được tạo ngay cạnh các file code (cùng thư mục).
DB_PATH = os.path.join(os.path.dirname(__file__), "hoadon.db")


def _connect():
    """Mở 1 kết nối tới file SQLite. row_factory giúp đọc dữ liệu ra dạng dict cho dễ."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """
    Tạo các bảng nếu chưa có, và nạp sẵn (seed) một vài giá trị mặc định cho lần chạy đầu tiên.
    Hàm này gọi nhiều lần cũng không sao (đã dùng IF NOT EXISTS).
    """
    conn = _connect()
    cur = conn.cursor()

    # 1) Bảng settings dạng key - value (mỗi dòng là 1 thông số)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key   TEXT PRIMARY KEY,
            value TEXT
        )
    """)

    # 2) Danh sách công ty -> loại "chuyên gia"
    cur.execute("""
        CREATE TABLE IF NOT EXISTS chuyen_gia_company (
            company TEXT PRIMARY KEY
        )
    """)

    # 3) Các "list phòng" và đơn giá tương ứng
    cur.execute("""
        CREATE TABLE IF NOT EXISTS room_price (
            id    INTEGER PRIMARY KEY AUTOINCREMENT,
            name  TEXT,
            price INTEGER
        )
    """)
    # Mỗi phòng thuộc về 1 list (qua list_id)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS room_price_member (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            list_id INTEGER,
            room    TEXT
        )
    """)

    # 4) Bảng Đồ uống khác
    cur.execute("""
        CREATE TABLE IF NOT EXISTS do_uong_khac (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            ten     TEXT,
            gia_ban INTEGER,
            da_xuat INTEGER DEFAULT 0
        )
    """)

    conn.commit()

    # ----- Seed giá trị mặc định (chỉ nạp nếu bảng đang rỗng) -----

    # settings mặc định
    defaults = {
        "thue_suat": "8",               # thuế suất %, dùng để tính TienThue
        "gia_co_ban_fallback": "277778",  # đơn giá cơ bản cho phòng không nằm trong list nào
        "hinh_thuc_thanh_toan": "TM/CK",  # hình thức thanh toán mặc định
    }
    for k, v in defaults.items():
        cur.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v)
        )

    # Công ty -> chuyên gia mặc định: booking, ctrip
    if cur.execute("SELECT COUNT(*) FROM chuyen_gia_company").fetchone()[0] == 0:
        for c in ["booking", "ctrip"]:
            cur.execute(
                "INSERT INTO chuyen_gia_company (company) VALUES (?)", (c,)
            )

    # List phòng mặc định: phòng 801 -> 370000 (370370). Các phòng khác dùng fallback 277778.
    if cur.execute("SELECT COUNT(*) FROM room_price").fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO room_price (name, price) VALUES (?, ?)",
            ("List 370370", 370370),
        )
        list_id = cur.lastrowid
        cur.execute(
            "INSERT INTO room_price_member (list_id, room) VALUES (?, ?)",
            (list_id, "801"),
        )

    # Vài đồ uống mẫu để tính năng "Đồ uống khác" chạy được ngay. Bạn có thể sửa/xóa sau.
    if cur.execute("SELECT COUNT(*) FROM do_uong_khac").fetchone()[0] == 0:
        for ten, gia in [("Coca Cola", 15000), ("Pepsi", 15000), ("Red Bull", 20000)]:
            cur.execute(
                "INSERT INTO do_uong_khac (ten, gia_ban, da_xuat) VALUES (?, ?, 0)",
                (ten, gia),
            )

    conn.commit()
    conn.close()


# =========================================================================
# NHÓM 1: SETTINGS (key - value)
# =========================================================================

def get_setting(key, default=None):
    """Lấy 1 thông số theo key. Trả về chuỗi (string)."""
    conn = _connect()
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key, value):
    """Lưu/ghi đè 1 thông số."""
    conn = _connect()
    conn.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, str(value)),
    )
    conn.commit()
    conn.close()


def get_thue_suat():
    """Tiện ích: lấy thuế suất dạng số (int). Mặc định 8."""
    try:
        return int(float(get_setting("thue_suat", "8")))
    except (TypeError, ValueError):
        return 8


def get_gia_co_ban_fallback():
    """Tiện ích: lấy đơn giá cơ bản fallback dạng số (int). Mặc định 277778."""
    try:
        return int(float(get_setting("gia_co_ban_fallback", "277778")))
    except (TypeError, ValueError):
        return 277778


def get_hinh_thuc_thanh_toan():
    return get_setting("hinh_thuc_thanh_toan", "TM/CK")


# =========================================================================
# NHÓM 2: CÔNG TY -> CHUYÊN GIA
# =========================================================================

def get_chuyen_gia_companies():
    """Trả về danh sách tên công ty (đã lowercase) sẽ mặc định ra loại chuyên gia."""
    conn = _connect()
    rows = conn.execute("SELECT company FROM chuyen_gia_company").fetchall()
    conn.close()
    return [r["company"].strip().lower() for r in rows]


def add_chuyen_gia_company(company):
    company = (company or "").strip().lower()
    if not company:
        return
    conn = _connect()
    conn.execute(
        "INSERT OR IGNORE INTO chuyen_gia_company (company) VALUES (?)", (company,)
    )
    conn.commit()
    conn.close()


def remove_chuyen_gia_company(company):
    conn = _connect()
    conn.execute(
        "DELETE FROM chuyen_gia_company WHERE company = ?",
        ((company or "").strip().lower(),),
    )
    conn.commit()
    conn.close()


# =========================================================================
# NHÓM 3: LIST PHÒNG -> ĐƠN GIÁ CƠ BẢN
# =========================================================================

def get_room_price_lists():
    """
    Trả về danh sách các list phòng, mỗi list là 1 dict:
    {"id": ..., "name": ..., "price": ..., "rooms": ["201", "301", ...]}
    """
    conn = _connect()
    lists = conn.execute("SELECT * FROM room_price ORDER BY id").fetchall()
    result = []
    for lst in lists:
        members = conn.execute(
            "SELECT room FROM room_price_member WHERE list_id = ? ORDER BY room",
            (lst["id"],),
        ).fetchall()
        result.append(
            {
                "id": lst["id"],
                "name": lst["name"],
                "price": lst["price"],
                "rooms": [m["room"] for m in members],
            }
        )
    conn.close()
    return result


def add_room_price_list(name, price):
    conn = _connect()
    conn.execute(
        "INSERT INTO room_price (name, price) VALUES (?, ?)", (name, int(price))
    )
    conn.commit()
    conn.close()


def update_room_price_list(list_id, name, price):
    conn = _connect()
    conn.execute(
        "UPDATE room_price SET name = ?, price = ? WHERE id = ?",
        (name, int(price), list_id),
    )
    conn.commit()
    conn.close()


def delete_room_price_list(list_id):
    conn = _connect()
    conn.execute("DELETE FROM room_price WHERE id = ?", (list_id,))
    conn.execute("DELETE FROM room_price_member WHERE list_id = ?", (list_id,))
    conn.commit()
    conn.close()


def set_rooms_for_list(list_id, rooms):
    """
    Ghi đè toàn bộ danh sách phòng cho 1 list.
    rooms: danh sách chuỗi tên phòng, ví dụ ["201", "301"].
    """
    conn = _connect()
    conn.execute("DELETE FROM room_price_member WHERE list_id = ?", (list_id,))
    for r in rooms:
        r = str(r).strip()
        if r:
            conn.execute(
                "INSERT INTO room_price_member (list_id, room) VALUES (?, ?)",
                (list_id, r),
            )
    conn.commit()
    conn.close()


def build_room_price_map():
    """
    Gộp tất cả list phòng lại thành 1 dict: { "tên phòng": đơn_giá }.
    Đồng thời PHÁT HIỆN LỖI: nếu 1 phòng nằm trong từ 2 list trở lên.

    Trả về (room_map, errors):
      - room_map: dict phòng -> giá
      - errors: danh sách chuỗi mô tả phòng bị trùng (rỗng nếu không có lỗi)
    """
    lists = get_room_price_lists()
    room_map = {}
    seen = {}  # phòng -> tên list đầu tiên gặp
    duplicates = {}  # phòng -> danh sách tên list bị trùng

    for lst in lists:
        for room in lst["rooms"]:
            room = str(room).strip()
            if room in seen:
                duplicates.setdefault(room, [seen[room]])
                duplicates[room].append(lst["name"])
            else:
                seen[room] = lst["name"]
                room_map[room] = lst["price"]

    errors = [
        f"Phòng '{room}' đang nằm trong nhiều list: {', '.join(names)}"
        for room, names in duplicates.items()
    ]
    return room_map, errors


# =========================================================================
# NHÓM 4: ĐỒ UỐNG KHÁC
# =========================================================================

def get_drinks():
    """Trả về danh sách đồ uống: [{"id", "ten", "gia_ban", "da_xuat"}, ...]."""
    conn = _connect()
    rows = conn.execute("SELECT * FROM do_uong_khac ORDER BY id").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_drink(ten, gia_ban, da_xuat=0):
    conn = _connect()
    conn.execute(
        "INSERT INTO do_uong_khac (ten, gia_ban, da_xuat) VALUES (?, ?, ?)",
        (ten, int(gia_ban), int(da_xuat)),
    )
    conn.commit()
    conn.close()


def update_drink(drink_id, ten, gia_ban, da_xuat):
    conn = _connect()
    conn.execute(
        "UPDATE do_uong_khac SET ten = ?, gia_ban = ?, da_xuat = ? WHERE id = ?",
        (ten, int(gia_ban), int(da_xuat), drink_id),
    )
    conn.commit()
    conn.close()


def delete_drink(drink_id):
    conn = _connect()
    conn.execute("DELETE FROM do_uong_khac WHERE id = ?", (drink_id,))
    conn.commit()
    conn.close()


def add_da_xuat_by_name(ten, quantity):
    """
    Cộng thêm 'quantity' vào cột 'da_xuat' của đồ uống có tên = ten.
    Dùng khi user xác nhận tính số lượng Đồ uống khác đã xuất.
    """
    conn = _connect()
    conn.execute(
        "UPDATE do_uong_khac SET da_xuat = da_xuat + ? WHERE ten = ?",
        (int(quantity), ten),
    )
    conn.commit()
    conn.close()
