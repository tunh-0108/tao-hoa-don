"""
settings_ui.py
--------------
Vẽ trang "Cài đặt" của app. Tách riêng cho gọn, app.py sẽ gọi hàm hien_thi_cai_dat().

Trang này cho phép chỉnh:
  1) Thuế suất, đơn giá cơ bản fallback, hình thức thanh toán mặc định.
  2) Danh sách Công ty -> mặc định ra loại "Lưu trú chuyên gia".
  3) Các "list phòng -> đơn giá cơ bản" (báo lỗi nếu 1 phòng nằm ở 2 list).
  4) Bảng Đồ uống khác (tên, giá bán, số lượng đã xuất).
"""

import streamlit as st
import database as db


def hien_thi_cai_dat():
    st.header("⚙️ Cài đặt")

    # =====================================================================
    # 1) THÔNG SỐ CHUNG
    # =====================================================================
    st.subheader("1. Thông số chung")
    col1, col2, col3 = st.columns(3)
    with col1:
        thue = st.number_input(
            "Thuế suất (%)", min_value=0, max_value=100,
            value=db.get_thue_suat(), step=1,
        )
    with col2:
        fallback = st.number_input(
            "Đơn giá cơ bản fallback (phòng không thuộc list nào)",
            min_value=0, value=db.get_gia_co_ban_fallback(), step=1,
        )
    with col3:
        httt = st.text_input(
            "Hình thức thanh toán mặc định", value=db.get_hinh_thuc_thanh_toan()
        )
    if st.button("💾 Lưu thông số chung"):
        db.set_setting("thue_suat", int(thue))
        db.set_setting("gia_co_ban_fallback", int(fallback))
        db.set_setting("hinh_thuc_thanh_toan", httt)
        st.success("Đã lưu thông số chung.")

    st.divider()

    # =====================================================================
    # 2) CÔNG TY -> CHUYÊN GIA
    # =====================================================================
    st.subheader("2. Công ty mặc định là 'Lưu trú chuyên gia'")
    st.caption("Các 'Công ty' trong danh sách này sẽ mặc định ra loại "
               "'Lưu trú chuyên gia - minibar ko cần theo thực tế'.")

    cong_ty_list = db.get_chuyen_gia_companies()
    for cty in cong_ty_list:
        c1, c2 = st.columns([4, 1])
        c1.write(f"• {cty}")
        if c2.button("Xóa", key=f"del_cty_{cty}"):
            db.remove_chuyen_gia_company(cty)
            st.rerun()

    cty_moi = st.text_input("Thêm công ty mới", key="them_cty")
    if st.button("➕ Thêm công ty"):
        if cty_moi.strip():
            db.add_chuyen_gia_company(cty_moi)
            st.rerun()

    st.divider()

    # =====================================================================
    # 3) LIST PHÒNG -> ĐƠN GIÁ CƠ BẢN
    # =====================================================================
    st.subheader("3. List phòng & đơn giá cơ bản")
    st.caption("Mỗi list gồm 1 đơn giá và danh sách phòng. Phòng KHÔNG thuộc list "
               "nào sẽ dùng 'đơn giá fallback' ở mục 1.")

    # Báo lỗi nếu phòng trùng giữa các list
    _, loi_trung = db.build_room_price_map()
    if loi_trung:
        for e in loi_trung:
            st.error("⚠️ " + e)

    for lst in db.get_room_price_lists():
        with st.expander(f"List: {lst['name']}  —  đơn giá {lst['price']:,}", expanded=False):
            ten = st.text_input("Tên list", value=lst["name"], key=f"ten_{lst['id']}")
            gia = st.number_input(
                "Đơn giá", min_value=0, value=int(lst["price"]), step=1,
                key=f"gia_{lst['id']}",
            )
            phong_text = st.text_area(
                "Danh sách phòng (mỗi phòng cách nhau bởi dấu phẩy hoặc xuống dòng)",
                value=", ".join(lst["rooms"]),
                key=f"phong_{lst['id']}",
            )
            cc1, cc2 = st.columns(2)
            if cc1.button("💾 Lưu list này", key=f"luu_list_{lst['id']}"):
                db.update_room_price_list(lst["id"], ten, gia)
                # Tách chuỗi phòng theo dấu phẩy và xuống dòng
                rooms = [r.strip() for r in phong_text.replace("\n", ",").split(",") if r.strip()]
                db.set_rooms_for_list(lst["id"], rooms)
                st.success("Đã lưu list.")
                st.rerun()
            if cc2.button("🗑️ Xóa list này", key=f"xoa_list_{lst['id']}"):
                db.delete_room_price_list(lst["id"])
                st.rerun()

    st.write("**Thêm list mới:**")
    nc1, nc2 = st.columns(2)
    ten_moi = nc1.text_input("Tên list mới", key="ten_list_moi")
    gia_moi = nc2.number_input("Đơn giá list mới", min_value=0, value=277778, step=1, key="gia_list_moi")
    if st.button("➕ Thêm list phòng"):
        if ten_moi.strip():
            db.add_room_price_list(ten_moi, gia_moi)
            st.rerun()

    st.divider()

    # =====================================================================
    # 4) BẢNG ĐỒ UỐNG KHÁC
    # =====================================================================
    st.subheader("4. Bảng 'Đồ uống khác'")
    st.caption("'Số lượng đã xuất' sẽ tự cộng dồn mỗi khi bạn xuất file và bấm "
               "'Đồng ý' tính số lượng. Bạn cũng có thể sửa tay tại đây.")

    for d in db.get_drinks():
        c1, c2, c3, c4 = st.columns([3, 2, 2, 1])
        ten = c1.text_input("Tên", value=d["ten"], key=f"d_ten_{d['id']}", label_visibility="collapsed")
        gia = c2.number_input("Giá bán", min_value=0, value=int(d["gia_ban"]), step=1,
                              key=f"d_gia_{d['id']}", label_visibility="collapsed")
        dx = c3.number_input("Đã xuất", min_value=0, value=int(d["da_xuat"]), step=1,
                             key=f"d_dx_{d['id']}", label_visibility="collapsed")
        with c4:
            if st.button("💾", key=f"d_luu_{d['id']}", help="Lưu dòng này"):
                db.update_drink(d["id"], ten, gia, dx)
                st.rerun()
            if st.button("🗑️", key=f"d_xoa_{d['id']}", help="Xóa dòng này"):
                db.delete_drink(d["id"])
                st.rerun()

    st.write("**Thêm đồ uống mới:**")
    a1, a2 = st.columns(2)
    ten_du = a1.text_input("Tên đồ uống mới", key="du_ten_moi")
    gia_du = a2.number_input("Giá bán", min_value=0, value=0, step=1, key="du_gia_moi")
    if st.button("➕ Thêm đồ uống"):
        if ten_du.strip():
            db.add_drink(ten_du, gia_du, 0)
            st.rerun()
