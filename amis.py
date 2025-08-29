"""Helper module for AMIS automation and document manipulation."""

import os
from typing import List, Tuple

from PIL import Image
from docx import Document
from docx.shared import Inches


def run_automation(
    username: str,
    password: str,
    record_id: str,
    download_dir: str,
    headless: bool = True,
) -> Tuple[str, List[str]]:
    """
    Dummy automation function.

    In thực tế bạn sẽ dùng Selenium:
    - Đăng nhập AMIS bằng username/password
    - Nhập record_id vào ô tìm kiếm
    - Mở chi tiết và tải file Word (Phiếu TTTT - Nhà đất)
    - Tải ảnh tài sản, ảnh sổ đỏ...

    Ở đây để đơn giản, hàm giả định bạn đã có sẵn file Word và ảnh trong thư mục download_dir.

    Returns:
        template_path: đường dẫn file Word đã tải
        images: list đường dẫn ảnh đã tải
    """
    # TODO: Thay phần này bằng selenium automation
    template_path = os.path.join(download_dir, "template.docx")
    # giả lập copy từ 1 file sẵn có
    with open(template_path, "wb") as f:
        f.write(b"")  # placeholder

    images = []  # ở đây bạn sẽ thêm danh sách ảnh đã tải
    return template_path, images


def fill_document(
    template_path: str,
    images: List[str],
    signature_path: str,
    output_path: str,
) -> None:
    """
    Mở file Word template, chèn ảnh và chữ ký.

    Args:
        template_path: file Word mẫu
        images: danh sách ảnh (tài sản, sổ đỏ...)
        signature_path: file chữ ký (png/jpg)
        output_path: file Word hoàn chỉnh để lưu
    """
    doc = Document(template_path)

    # Chèn ảnh vào cuối phụ lục
    doc.add_heading("Phụ lục ảnh", level=2)
    for img_path in images:
        if os.path.exists(img_path):
            doc.add_picture(img_path, width=Inches(3))
            doc.add_paragraph(" ")

    # Chèn chữ ký ở cuối
    if os.path.exists(signature_path):
        doc.add_page_break()
        doc.add_heading("Chữ ký cán bộ khảo sát", level=2)
        doc.add_picture(signature_path, width=Inches(2))

    doc.save(output_path)
