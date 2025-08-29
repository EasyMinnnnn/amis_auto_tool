"""Helper module for AMIS automation and document manipulation.

- run_automation(): đăng nhập AMIS, tìm ID, tải file Word và ảnh về download_dir.
- fill_document(): chèn ảnh đúng ô trong bảng “Phụ lục Ảnh TSSS” và chữ ký.
"""

import os, time
from typing import List, Tuple, Optional, Callable
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from docx import Document
from docx.shared import Inches


# ===================== Selenium helpers =====================

def _make_driver(download_dir: str, headless: bool = True) -> webdriver.Chrome:
    """Tạo Chrome với thư mục tải cụ thể (hợp với Streamlit Cloud)."""
    os.makedirs(download_dir, exist_ok=True)
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    # Nếu Chromium ở path khác, chỉnh lại:
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


def _click_if_appears(driver: webdriver.Chrome, selectors) -> bool:
    """Thử click một trong các selector (im lặng nếu không có)."""
    for by, sel in selectors:
        try:
            el = driver.find_element(by, sel)
            el.click()
            time.sleep(0.2)
            return True
        except Exception:
            pass
    return False


def _handle_post_login_popups(driver: webdriver.Chrome):
    """Đóng popups/onboarding/2FA prompt che UI nếu có."""
    _click_if_appears(driver, [
        (By.XPATH, "//button[contains(.,'Bỏ qua') or contains(.,'Tiếp tục làm việc')]"),
        (By.XPATH, "//button[contains(.,'Đóng') or contains(.,'Bỏ qua, tiếp tục làm việc')]"),
        (By.CSS_SELECTOR, "[aria-label='Close'],[data-dismiss],.close"),
    ])


def _switch_to_last_tab(driver: webdriver.Chrome):
    """Chuyển sang tab cuối nếu site mở tab mới."""
    try:
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
    except Exception:
        pass
    try:
        driver.switch_to.default_content()
    except Exception:
        pass


def _for_each_frame(driver: webdriver.Chrome, fn: Callable[[], Optional[object]]) -> Optional[object]:
    """
    Chạy fn() ở current context; nếu không ra, duyệt đệ quy mọi iframe cho tới khi có kết quả.
    Trả về kết quả đầu tiên fn() trả về khác None.
    """
    # 1) thử ở current context
    try:
        res = fn()
        if res is not None:
            return res
    except Exception:
        pass

    # 2) duyệt mọi iframe
    frames = driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
    for i in range(len(frames)):
        try:
            driver.switch_to.frame(i)
            res = _for_each_frame(driver, fn)
            driver.switch_to.parent_frame()
            if res is not None:
                return res
        except Exception:
            try:
                driver.switch_to.parent_frame()
            except Exception:
                pass
    return None


def find_in_any_frame(driver: webdriver.Chrome, candidates: List[tuple]) -> Optional[object]:
    """Tìm phần tử khớp một trong các (By, selector) ở mọi frame."""
    def finder():
        for by, sel in candidates:
            try:
                el = driver.find_element(by, sel)
                return el
            except NoSuchElementException:
                continue
        return None

    # Luôn bắt đầu từ default_content
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    return _for_each_frame(driver, finder)


def wait_in_any_frame(driver: webdriver.Chrome, candidates: List[tuple], timeout: int = 30) -> object:
    """Chờ tới khi tìm thấy phần tử trong bất kỳ frame nào (polling)."""
    end = time.time() + timeout
    last_exc = None
    while time.time() < end:
        el = find_in_any_frame(driver, candidates)
        if el is not None:
            return el
        time.sleep(0.5)
    raise TimeoutException("Timeout waiting element in any frame")


# ===================== Main automation =====================

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
    driver = _make_driver(download_dir, headless=headless)
    wait = WebDriverWait(driver, 40)

    try:
        # 1) Login AMIS
        driver.get("https://amisapp.misa.vn/")

        # Username
        username_input = wait.until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR,
                "#box-login-right > div > div > div.login-form-basic-container > "
                "div > div.login-form-inputs.login-class > "
                "div.wrap-input.username-wrap.validate-input > input",
            ))
        )
        username_input.send_keys(username)

        # Password
        password_input = driver.find_element(
            By.CSS_SELECTOR,
            "#box-login-right > div > div > div.login-form-basic-container > "
            "div > div.login-form-inputs.login-class > "
            "div.wrap-input.pass-wrap.validate-input > input",
        )
        password_input.send_keys(password)

        # Click Login
        driver.find_element(
            By.CSS_SELECTOR,
            "#box-login-right > div > div > div.login-form-basic-container > "
            "div > div.login-form-btn-container.login-class > button",
        ).click()

        # Đợi domain đúng, xử lý popup
        wait.until(EC.url_contains("amisapp.misa.vn"))
        time.sleep(0.8)
        _handle_post_login_popups(driver)

        # 2) Vào trang quy trình (không chờ URL; chờ UI thực tế)
        driver.get("https://amisapp.misa.vn/process/execute/1")
        time.sleep(0.6)
        _switch_to_last_tab(driver)

        # Nếu có spinner/overlay loading của UI (tùy framework), đợi biến mất (best-effort)
        # Không fail nếu không có.
        try:
            # ví dụ các lớp phổ biến: .el-loading-mask, .loading-mask, .ant-spin
            for css in [".el-loading-mask", ".loading-mask", ".ant-spin"]:
                for _ in range(10):
                    if not driver.find_elements(By.CSS_SELECTOR, css):
                        break
                    time.sleep(0.3)
        except Exception:
            pass

        # 3) Tìm ô tìm kiếm trong MỌI frame (input/textarea, “Tìm kiếm”/“Nhấn Enter”)
        search = wait_in_any_frame(
            driver,
            candidates=[
                (By.CSS_SELECTOR, "input[placeholder*='Tìm kiếm']"),
                (By.CSS_SELECTOR, "input[placeholder*='Nhấn Enter']"),
                (By.XPATH, "//input[contains(@placeholder,'Tìm kiếm') or contains(@placeholder,'Nhấn Enter')]"),
                (By.XPATH, "//textarea[contains(@placeholder,'Tìm kiếm') or contains(@placeholder,'Nhấn Enter')]"),
            ],
            timeout=35,
        )

        # 4) Nhập & tìm
        try:
            search.clear()
        except Exception:
            search.send_keys(Keys.CONTROL, "a")
            search.send_keys(Keys.BACK_SPACE)

        search.send_keys(record_id)
        search.send_keys(Keys.ENTER)

        # 5) Click dòng đầu tiên của bảng kết quả (tìm trong mọi frame)
        first_row = wait_in_any_frame(
            driver,
            candidates=[
                (By.CSS_SELECTOR, "table tbody tr"),
                (By.XPATH, "(//table//tbody//tr)[1]"),
            ],
            timeout=35,
        )
        try:
            first_row.click()
        except Exception:
            driver.execute_script("arguments[0].click();", first_row)
        time.sleep(0.8)

        # 6) Xem trước mẫu in -> chọn Phiếu TTTT - Nhà đất -> Tải xuống
        #    Các nút này cũng có thể nằm trong frame => tìm theo mọi frame
        btn_preview = wait_in_any_frame(
            driver,
            candidates=[(By.XPATH, "//span[contains(.,'Xem trước mẫu in')]")],
            timeout=35,
        )
        try:
            btn_preview.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn_preview)
        time.sleep(0.6)

        item_phieu = wait_in_any_frame(
            driver,
            candidates=[(By.XPATH, "//div[contains(.,'Phiếu TTTT - Nhà đất')]")],
            timeout=35,
        )
        try:
            item_phieu.click()
        except Exception:
            driver.execute_script("arguments[0].click();", item_phieu)
        time.sleep(0.6)

        btn_download = wait_in_any_frame(
            driver,
            candidates=[(By.XPATH, "//button[contains(.,'Tải xuống')]")],
            timeout=35,
        )
        try:
            btn_download.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn_download)

        template_path = _wait_for_docx(download_dir, timeout=60)

        # 7) Tải ảnh (placeholder – có thể cần refine selector)
        images = _download_images_from_detail(driver, download_dir)

        return template_path, images

    finally:
        driver.quit()


# ===================== Download helpers =====================

def _wait_for_docx(folder: str, timeout: int = 60) -> str:
    for _ in range(timeout):
        for f in os.listdir(folder):
            if f.lower().endswith(".docx"):
                return os.path.join(folder, f)
        time.sleep(1)
    raise FileNotFoundError("Không thấy file .docx sau khi tải")


def _download_images_from_detail(driver, download_dir: str) -> List[str]:
    """Ví dụ: lấy src của ảnh thumbnail và tải về bằng requests."""
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


# ===================== Word processing =====================

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
