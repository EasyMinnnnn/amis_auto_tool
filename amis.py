"""Helper module for AMIS automation and document manipulation.

Trong bản này:
- run_automation() CHƯA đăng nhập AMIS thật. Thay vào đó tạo 1 file .docx mẫu
  hợp lệ để tránh lỗi khi demo trên Streamlit Cloud.
- Khi bạn sẵn sàng, thay phần “TODO: Selenium” bằng code tự động truy cập AMIS,
  tìm ID, tải file Word + ảnh về thư mục download_dir, rồi return đường dẫn thực.
"""

from __future__ import annotations

import os
from typing import List, Tuple

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
    TẠM THỜI: tạo file DOCX mẫu hợp lệ để tránh lỗi PackageNotFoundError.
    TODO (bạn sẽ thay sau): dùng Selenium để đăng nhập AMIS, tìm theo record_id,
    tải file Word (Phiếu TTTT - Nhà đất) và các ảnh về download_dir.

    Returns:
        template_path: đường dẫn file Word (hiện là file mẫu được tạo)
        images: danh sách đường dẫn ảnh đã tải (tạm để rỗng)
    """
    os.makedirs(download_dir, exist_ok=True)

    # Tạo file DOCX mẫu hợp lệ
    template_path = os.path.join(download_dir, "template.docx")
    doc = Document()
    doc.add_heading("PHIẾU THU THẬP THÔNG TIN VỀ BẤT ĐỘNG SẢN (MẪU)", level=1)
    doc.add_paragraph(f"Mã Tài sản (demo): {record_id}")
    doc.add_paragraph("Phần nội dung sẽ được thay bằng file tải từ AMIS khi bạn triển khai Selenium.")
    doc.add_page_break()
    doc.add_heading("Phụ lục Ảnh TSSS", level=2)
    doc.add_paragraph("Các ảnh sẽ được chèn bên dưới…")
    doc.save(template_path)

    # Tạm thời chưa có ảnh nào (khi nối Selenium thì điền list ảnh thật)
    images: List[str] = []

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
        template_path: file Word mẫu (hợp lệ .docx)
        images: danh sách ảnh (tài sản, sổ đỏ…) – có thể để rỗng
        signature_path: file chữ ký (png/jpg)
        output_path: file Word hoàn chỉnh để lưu
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    doc = Document(template_path)

    # Chèn ảnh (nếu có) vào cuối phụ lục
    if images:
        doc.add_paragraph("")  # cách dòng
        for img_path in images:
            if os.path.exists(img_path):
                try:
                    doc.add_picture(img_path, width=Inches(3))
                    doc.add_paragraph("")  # cách dòng giữa ảnh
                except Exception as e:
                    # Không để app gãy nếu 1 ảnh lỗi
                    doc.add_paragraph(f"[Không chèn được ảnh: {img_path}] ({e})")

    # Chèn chữ ký ở cuối (nếu có)
    if signature_path and os.path.exists(signature_path):
        doc.add_page_break()
        doc.add_heading("Chữ ký cán bộ khảo sát", level=2)
        try:
            doc.add_picture(signature_path, width=Inches(2))
        except Exception as e:
            doc.add_paragraph(f"[Không chèn được chữ ký: {e}]")
    else:
        doc.add_page_break()
        doc.add_heading("Chữ ký cán bộ khảo sát", level=2)
        doc.add_paragraph("[Chưa có chữ ký]")

    # Lưu file hoàn chỉnh
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(output_path)
