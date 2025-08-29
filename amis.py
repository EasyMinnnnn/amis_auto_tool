"""Helper module for AMIS automation and document manipulation.

- run_automation(): đăng nhập AMIS, tìm ID, tải file Word và ảnh về download_dir.
- fill_document(): chèn ảnh đúng ô trong bảng “Phụ lục Ảnh TSSS” và chữ ký.
"""

import os, time
from typing import List, Tuple
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from docx import Document
from docx.shared import Inches

# ===================== Selenium: login + download =====================

def _make_driver(download_dir: str) -> webdriver.Chrome:
    """Tạo Chrome headless với thư mục tải cụ thể (dùng trên Streamlit Cloud)."""
    os.makedirs(download_dir, exist_ok=True)
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    # Nếu môi trường có Chromium theo path khác, chỉnh lại dòng dưới:
    opts.binary_location = "/usr/bin/chromium"

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": True,
    }
    opts.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=opts)
    driver.set_window_size(1440, 900)
    return driver


def _wait_clickable(wait: WebDriverWait, by: By, value: str):
    return wait.until(EC.element_to_be_clickable((by, value)))


def _open_global_search(driver: webdriver.Chrome, wait: WebDriverWait):
    """
    Mở Global Search ở top-nav nếu đang đóng, sau đó trả về phần tử ô tìm kiếm.
    Thử nhiều selector để tránh phụ thuộc UI thay đổi (textarea / input).
    """
    # Luôn về khung gốc (top_nav không nằm trong iframe nội dung)
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # 1) Nếu textarea/input đã có sẵn thì trả về luôn
    quick_selectors = [
        (By.CSS_SELECTOR, "#top_nav textarea"),
        (By.CSS_SELECTOR, "textarea[placeholder*='Tìm kiếm']"),
        (By.XPATH, "//textarea[contains(@placeholder,'Tìm kiếm')]"),
        (By.XPATH, "//input[contains(@placeholder,'Tìm kiếm')]"),
    ]
    for by, sel in quick_selectors:
        els = driver.find_elements(by, sel)
        if els:
            return els[0]

    # 2) Thử bấm icon mở Global Search (kính lúp)
    #   Thử nhiều khả năng tên lớp/aria-label khác nhau
    openers = [
        (By.CSS_SELECTOR, "#top_nav .global-search-wrap .icon, #top_nav .global-search-wrap button"),
        (By.CSS_SELECTOR, "#top_nav [aria-label*='Tìm kiếm']"),
        (By.CSS_SELECTOR, "#top_nav .global-search-wrap svg"),
        (By.XPATH, "//*[@id='top_nav']//*[self::button or self::div]//*[name()='svg' or contains(.,'Tìm kiếm')]/ancestor::*[self::button or self::div][1]"),
    ]
    clicked = False
    for by, sel in openers:
        try:
            btn = _wait_clickable(wait, by, sel)
            btn.click()
            clicked = True
            break
        except Exception:
            continue

    # 3) Một số hệ thống bật Global Search bằng phím tắt ('/' hoặc Ctrl+K)
    if not clicked:
        try:
            body = driver.find_element(By.TAG_NAME, "body")
            body.send_keys("/")   # thử '/'
            time.sleep(0.5)
            body.send_keys(Keys.CONTROL, "k")  # thử Ctrl+K
            time.sleep(0.5)
        except Exception:
            pass

    # 4) Chờ ô tìm kiếm xuất hiện rồi trả về
    return wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//textarea[contains(@placeholder,'Tìm kiếm')] | //input[contains(@placeholder,'Tìm kiếm')]"
        ))
    )


def run_automation(
    username: str,
    password: str,
    record_id: str,
    download_dir: str,
    headless: bool = True,
) -> Tuple[str, List[str]]:
    """
    Đăng nhập AMIS, tìm record_id, tải file Word và ảnh về.
    Trả về: (đường dẫn file Word, danh sách ảnh).
    """
    driver = _make_driver(download_dir)
    wait = WebDriverWait(driver, 25)

    try:
        # 1) Login AMIS
        driver.get("https://amisapp.misa.vn/")

        # Nhập Username
        username_input = wait.until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    "#box-login-right > div > div > div.login-form-basic-container > "
                    "div > div.login-form-inputs.login-class > "
                    "div.wrap-input.username-wrap.validate-input > input",
                )
            )
        )
        username_input.send_keys(username)

        # Nhập Password
        password_input = driver.find_element(
            By.CSS_SELECTOR,
            "#box-login-right > div > div > div.login-form-basic-container > "
            "div > div.login-form-inputs.login-class > "
            "div.wrap-input.pass-wrap.validate-input > input",
        )
        password_input.send_keys(password)

        # Click nút Đăng nhập
        login_btn = driver.find_element(
            By.CSS_SELECTOR,
            "#box-login-right > div > div > div.login-form-basic-container > "
            "div > div.login-form-btn-container.login-class > button",
        )
        login_btn.click()

        # Chờ tới khi login thành công (URL chứa amisapp.misa.vn)
        wait.until(EC.url_contains("amisapp.misa.vn"))
        time.sleep(2)

        # 2) Vào trang quy trình (execute/1) rồi mở Global Search để tìm record_id
        driver.get("https://amisapp.misa.vn/process/execute/1")

        # Đợi top_nav có mặt (đảm bảo header đã render)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#top_nav")))
        time.sleep(0.5)

        search = _open_global_search(driver, wait)
        # Đảm bảo có thể thao tác
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[self::textarea or self::input][contains(@placeholder,'Tìm kiếm')]")))
        except Exception:
            pass

        # Nhập & tìm
        try:
            search.clear()
        except Exception:
            # một số textarea không hỗ trợ clear(); dùng tổ hợp phím
            search.send_keys(Keys.CONTROL, "a")
            search.send_keys(Keys.BACK_SPACE)
        search.send_keys(record_id)
        search.send_keys(Keys.ENTER)

        # Chờ bảng kết quả rồi click vào dòng đầu
        first_row = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "table tbody tr"))
        )
        first_row.click()
        time.sleep(1.5)

        # 3) Xem trước mẫu in -> chọn Phiếu TTTT - Nhà đất -> Tải xuống
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(.,'Xem trước mẫu in')]")
            )
        ).click()
        time.sleep(1)

        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(.,'Phiếu TTTT - Nhà đất')]")
            )
        ).click()
        time.sleep(1)

        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Tải xuống')]"))
        ).click()

        template_path = _wait_for_docx(download_dir, timeout=60)

        # 4) Tải ảnh tài sản/rao bán (tùy chỉnh selector khi biết đúng thumbnail)
        images = _download_images_from_detail(driver, download_dir)

        return template_path, images

    finally:
        driver.quit()


def _wait_for_docx(folder: str, timeout: int = 60) -> str:
    for _ in range(timeout):
        for f in os.listdir(folder):
            if f.lower().endswith(".docx"):
                return os.path.join(folder, f)
        time.sleep(1)
    raise FileNotFoundError("Không thấy file .docx sau khi tải")


def _download_images_from_detail(driver, download_dir: str) -> List[str]:
    """Ví dụ: lấy src của ảnh thumbnail và tải về bằng requests.
    TODO: thay selector 'img' bằng selector thumbnail cụ thể của AMIS (nếu có).
    """
    images: List[str] = []
    thumbs = driver.find_elements(By.CSS_SELECTOR, "img")
    for i, t in enumerate(thumbs[:10], start=1):
        try:
            src = t.get_attribute("src")
            if src and src.startswith("http"):
                r = requests.get(src, timeout=20)
                img_path = os.path.join(download_dir, f"image_{i}.jpg")
                with open(img_path, "wb") as f:
                    f.write(r.content)
                images.append(img_path)
        except Exception:
            pass
    return images


# ===================== Xử lý Word =====================

def fill_document(
    template_path: str, images: List[str], signature_path: str, output_path: str
) -> None:
    """Mở file Word template, chèn ảnh đúng ô trong bảng 'Phụ lục Ảnh TSSS' và chữ ký."""
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    doc = Document(template_path)

    # Map ảnh theo nhãn trong bảng
    slot_map = {
        "Thông tin rao bán/sổ đỏ": images[0:1],
        "Mặt trước tài sản": images[1:2],
        "Tổng thể tài sản": images[2:3],
        "Đường phía trước tài sản": images[3:5],  # 2 ảnh
        "Ảnh khác": images[5:7],  # 2 ảnh
    }

    # Tìm bảng chứa chữ "Phụ lục Ảnh TSSS"
    def _table_has_phu_luc(tbl):
        text = "\n".join(cell.text for row in tbl.rows for cell in row.cells)
        return "Phụ lục" in text and "Ảnh TSSS" in text

    target_table = None
    for tbl in doc.tables:
        if _table_has_phu_luc(tbl):
            target_table = tbl
            break

    if target_table:
        for row in target_table.rows:
            for ci, cell in enumerate(row.cells):
                label = cell.text.strip()
                if label in slot_map and ci + 1 < len(row.cells):
                    dest = row.cells[ci + 1]
                    for p in dest.paragraphs:
                        if p.text:
                            p.text = ""
                    for pth in slot_map[label]:
                        if os.path.exists(pth):
                            dest.paragraphs[0].add_run().add_picture(
                                pth, width=Inches(2.2)
                            )
                            dest.add_paragraph("")
    else:
        # fallback: chèn ảnh cuối tài liệu
        doc.add_heading("Phụ lục Ảnh TSSS", level=2)
        for pth in images:
            if os.path.exists(pth):
                doc.add_picture(pth, width=Inches(3))
                doc.add_paragraph("")

    # Chèn chữ ký
    doc.add_page_break()
    doc.add_heading("Chữ ký cán bộ khảo sát", level=2)
    if signature_path and os.path.exists(signature_path):
        try:
            doc.add_picture(signature_path, width=Inches(2))
        except Exception as e:
            doc.add_paragraph(f"[Không chèn được chữ ký: {e}]")

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
