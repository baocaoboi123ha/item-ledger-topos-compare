# -*- coding: utf-8 -*-
"""Web UI — so sánh Item Ledger Entries vs Topos (tham khảo MGMTNA Compare By Date 9.0)."""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

import streamlit as st

_APP_DIR = Path(__file__).resolve().parent
if str(_APP_DIR) not in sys.path:
    sys.path.insert(0, str(_APP_DIR))

import engine

XLSX = ["xlsx"]

st.set_page_config(
    page_title="ILE vs Topos Compare",
    page_icon="📑",
    layout="centered",
)

st.title("ILE vs Topos Compare")
st.caption(
    f"**Engine {engine.ENGINE_VERSION}** — Không đăng nhập. "
    "File chỉ xử lý tạm trên máy chủ rồi xóa — không lưu dữ liệu."
)

st.markdown(
    """
**Key đối chiếu (gom Quantity theo từng key):**
1. **Ngày** — Item Ledger `Posting Date` ↔ Topos `Ngay`
2. **Location / Tên Cửa Hàng** — `Location Code` ↔ `Ten Cua Hang`
3. **Type** — cột `Type` (ST, CP, …)
4. **Item No / Mã Nội Bộ** — `Item No.` ↔ `Ma Noi Bo`

**Topos (Phương Thức):**
- **HĐ Trả** khớp **HĐ Bán** (Ma Hoa Don Goc, cùng SL) → không cộng bán/trả (net 0 tại ngày bán).
- **HĐ Bán** cùng **Mã chứng từ** → net theo mã CT (vd. +N và −N cùng CT = 0) rồi gom theo ngày/key.

**Cột kết quả:** Ngày, Location, Type, Item No/Mã Nội Bộ,
Quantity Item Ledger, Quantity Topos, **Lệch = Topos − Item Ledger**,
**Unit cost** (Cost Amount ÷ Qty trên ILE theo mã M; không có thì 0).
Chỉ xuất các key **lệch** (không in dòng khớp).

File tải về **gộp file gốc + kết quả**: sheet nguyên bản giữ nguyên thứ tự ở trước,
sheet `Ket qua lech` và `TongHop` ở sau.
"""
)

tab_one, tab_two = st.tabs(["Một file (2 sheet)", "Hai file riêng"])


def _write_upload(uploaded, folder: Path) -> Path:
    dest = folder / Path(uploaded.name).name
    dest.write_bytes(uploaded.getvalue())
    return dest


def _run(ile_file, topos_file):
    work = Path(tempfile.mkdtemp(prefix="ile-topos-"))
    try:
        ile_path = _write_upload(ile_file, work)
        topos_path = _write_upload(topos_file, work) if topos_file is not None else None
        output = work / (ile_path.stem + engine.OUTPUT_SUFFIX)
        logs: list[str] = []

        def log(msg: str) -> None:
            logs.append(str(msg))

        result = engine.compare_workbook(ile_path, output, log, topos_path)
        payload = Path(result["OutputPath"]).read_bytes()
        name = Path(result["OutputPath"]).name
        summary = (
            f"Xong {ile_file.name}: {result['MismatchRows']} key lệch "
            f"(cả hai bên {result['QtyDiffBoth']}, "
            f"chỉ Item Ledger {result['OnlyIle']}, chỉ Topos {result['OnlyTopos']})."
        )
        return payload, name, "\n".join(logs), summary, result
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _show_result(payload, name, logs, summary, result, dl_key: str):
    st.success(summary)
    st.download_button(
        "Tải file kết quả",
        data=payload,
        file_name=name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=dl_key,
    )
    preview = result.get("PreviewRows") or []
    if preview:
        st.subheader("Xem nhanh các dòng lệch")
        st.caption(
            f"Hiện tối đa {engine.WEB_PREVIEW_ROWS} dòng. Đầy đủ trong file tải về "
            f"(sheet `{engine.RESULT_SHEET}`)."
        )
        st.dataframe(preview, use_container_width=True, hide_index=True)
    with st.expander("Nhật ký xử lý"):
        st.code(logs or "(trống)", language="text")


with tab_one:
    st.write(
        "Một `.xlsx` có sheet **Item Ledger Entries** và **Topos** "
        "(ví dụ `Item Ledger Entries.xlsx`)."
    )
    data_one = st.file_uploader(
        "File dữ liệu (.xlsx)",
        type=XLSX,
        key="one_file",
    )
    if st.button("Chạy đối chiếu", type="primary", key="run_one"):
        if data_one is None:
            st.warning("Hãy chọn file có đủ 2 sheet Item Ledger và Topos.")
        else:
            try:
                with st.spinner("Đang gom key và đối chiếu…"):
                    payload, name, logs, summary, result = _run(data_one, None)
                _show_result(payload, name, logs, summary, result, "dl_one")
            except engine.CompareError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Lỗi: {exc}")

with tab_two:
    st.write("Tải riêng file Item Ledger và file Topos nếu chúng không nằm cùng workbook.")
    ile_file = st.file_uploader(
        "1. File Item Ledger Entries (.xlsx)",
        type=XLSX,
        key="ile_file",
    )
    topos_file = st.file_uploader(
        "2. File Topos (.xlsx)",
        type=XLSX,
        key="topos_file",
    )
    if st.button("Chạy đối chiếu", type="primary", key="run_two"):
        if ile_file is None:
            st.warning("Hãy chọn file Item Ledger.")
        elif topos_file is None:
            st.warning("Hãy chọn file Topos.")
        else:
            try:
                with st.spinner("Đang gom key và đối chiếu…"):
                    payload, name, logs, summary, result = _run(ile_file, topos_file)
                _show_result(payload, name, logs, summary, result, "dl_two")
            except engine.CompareError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Lỗi: {exc}")
