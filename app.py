"""Streamlit application for automating AMIS downloads and document creation."""

from __future__ import annotations

import os
import tempfile
from datetime import datetime

import streamlit as st

# Dùng import tuyệt đối vì app.py và amis.py ở cùng thư mục
import amis


def main() -> None:
    st.set_page_config(page_title="AMIS Automation", layout="centered")
    st.title("AMIS Auto-Downloader")

    st.write(
        """
        Công cụ này tự động đăng nhập AMIS và tải **Phiếu TTTT - Nhà đất** kèm ảnh.
        Bạn **chỉ cần nhập Record ID** (ví dụ: 80002) và upload ảnh chữ ký.
        Phần còn lại công cụ sẽ xử lý tự động.
        """
    )

    # ===== Thông tin đăng nhập =====
    username = st.text_input("AMIS username")
    password = st.text_input("AMIS password", type="password")

    # ===== Chỉ nhập Record ID =====
    record_id = st.text_input("Record ID", placeholder="Ví dụ: 80002")

    # ===== Ảnh chữ ký =====
    signature_file = st.file_uploader(
        "Signature image (PNG/JPG)", type=["png", "jpg", "jpeg"]
    )

    # Tùy chọn headless (khuyến nghị khi chạy trên Cloud/Streamlit)
    headless = st.checkbox("Run headless", value=True)

    if st.button("Run automation"):
        # Kiểm tra dữ liệu đầu vào
        if not username or not password or not record_id or not signature_file:
            st.error("Vui lòng nhập username, password, Record ID và upload ảnh chữ ký.")
            return

        # Tạo thư mục tạm làm việc
        with tempfile.TemporaryDirectory() as tmpdir:
            downloads_dir = os.path.join(tmpdir, "downloads")
            os.makedirs(downloads_dir, exist_ok=True)

            # Lưu chữ ký tạm
            sig_path = os.path.join(tmpdir, "signature.png")
            sig_bytes = signature_file.read()
            with open(sig_path, "wb") as f:
                f.write(sig_bytes)

            st.info("Đang đăng nhập AMIS và tải dữ liệu…")
            try:
                # Gọi tool: chỉ truyền record_id, bỏ qua bất kỳ thao tác tìm kiếm thủ công trên UI
                template_path, images = amis.run_automation(
                    username=username,
                    password=password,
                    download_dir=downloads_dir,
                    headless=headless,
                    record_id=record_id,     # chỉ cần Record ID
                    # Nếu amis.py hỗ trợ execution_id thì sẽ tự bỏ qua bước tìm kiếm
                    # (không cần sửa gì thêm ở đây)
                )
            except Exception as e:
                st.exception(e)
                return

            st.success("Đã tải Word template và ảnh. Đang ghép tài liệu…")

            # Đặt tên file đầu ra
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(tmpdir, f"Phieu_TTTT_{record_id}_{timestamp}.docx")

            try:
                amis.fill_document(
                    template_path=template_path,
                    images=images,
                    signature_path=sig_path,
                    output_path=output_path,
                )
            except Exception as e:
                st.exception(e)
                return

            # Xuất file cho người dùng tải
            with open(output_path, "rb") as f:
                final_docx = f.read()

            st.success("Hoàn tất!")
            st.download_button(
                label="Tải Word đã hoàn thiện",
                data=final_docx,
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )


if __name__ == "__main__":
    main()
