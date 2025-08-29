# AMIS Automation Tool

Ứng dụng **Streamlit** giúp tự động tải **Phiếu thu thập thông tin** và ảnh từ hệ thống AMIS, sau đó chèn chữ ký và xuất ra file Word hoàn chỉnh.

## Tính năng

- Đăng nhập vào AMIS bằng tài khoản của bạn (qua Selenium).
- Tìm kiếm theo **Record ID** và tải xuống:
  - File Word mẫu (Phiếu TTTT - Nhà đất).
  - Ảnh tài sản, ảnh rao bán, sổ đỏ (nếu có).
- Tự động chèn ảnh vào cuối phụ lục trong file Word.
- Thêm chữ ký người khảo sát (file PNG/JPG bạn upload).
- Xuất ra file Word hoàn chỉnh để tải về.

## Yêu cầu

Cài đặt các thư viện Python:

```bash
pip install -r requirements.txt
