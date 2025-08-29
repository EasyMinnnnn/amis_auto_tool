"""Helper module for AMIS automation and document manipulation."""

import os, time
from typing import List, Tuple
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from docx import Document
from docx.shared import Inches

# ===================== Selenium base =====================

def _make_driver(download_dir: str, headless: bool = True) -> webdriver.Chrome:
    os.makedirs(download_dir, exist_ok=True)
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    # chỉnh nếu chromium ở path khác
    opts.binary_location = "/usr/bin/chromium"
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": True,
    }
    opts.add_experimental_option("prefs", prefs)
    drv = webdriver.Chrome(options=opts)
    drv.set_window_size(1440, 900)
    return drv

# ===================== MAIN =====================

def run_automation(
    username: str,
    password: str,
    record_id: str,
    download_dir: str,
    headless: bool = True,
) -> Tuple[str, List[str]]:
    """
    Đăng nhập AMIS, nhập ID vào ô 'Tìm kiếm thông minh với AI', mở chi tiết, tải file Word, lấy ảnh.
    Trả về: (đường dẫn file Word, danh sách ảnh).
    """
    driver = _make_driver(download_dir, headless=headless)
    wait = WebDriverWait(driver, 35)

    try:
        # 1) Login
        driver.get("https://amisapp.misa.vn/")

        user_el = wait.until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "#box-login-right .login-form-inputs .username-wrap input",
        )))
        user_el.send_keys(username)

        pw_el = driver.find_element(
            By.CSS_SELECTOR,
            "#box-login-right .login-form-inputs .pass-wrap input"
        )
        pw_el.send_keys(password)

        driver.find_element(
            By.CSS_SELECTOR,
            "#box-login-right .login-form-btn-container button"
        ).click()

        wait.until(EC.url_contains("amisapp.misa.vn"))
        time.sleep(0.8)

        # Đóng popup 2FA/onboarding nếu có (best-effort)
        for by, sel in [
            (By.XPATH, "//button[contains(.,'Bỏ qua')]"),
            (By.XPATH, "//button[contains(.,'Tiếp tục làm việc')]"),
            (By.XPATH, "//button[contains(.,'Đóng')]"),
            (By.CSS_SELECTOR, "[aria-label='Close'],[data-dismiss],.close"),
        ]:
            try:
                driver.find_element(by, sel).click()
                time.sleep(0.2)
            except Exception:
                pass

        # 2) Mở trang Quy trình (Lượt chạy)
        driver.get("https://amisapp.misa.vn/process/execute/1")
        time.sleep(0.8)

        # Đảm bảo đúng tab "Lượt chạy" (nếu tồn tại)
        try:
            tab_runs = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(.,'Lượt chạy')]"))
            )
            driver.execute_script("arguments[0].click();", tab_runs)
            time.sleep(0.4)
        except Exception:
            pass

        # 3) Tìm ô tìm kiếm (textarea.global-search-input hoặc biến thể placeholder)
        candidates = [
            (By.CSS_SELECTOR, "textarea.global-search-input"),
            (By.CSS_SELECTOR, "textarea[placeholder*='Tìm kiếm']"),
            (By.CSS_SELECTOR, "input.global-search-input"),
            (By.XPATH, "//*[self::textarea or self::input][contains(@placeholder,'Tìm kiếm')]"),
        ]

        search_el = None
        for by, sel in candidates:
            try:
                els = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((by, sel))
                )
                # chọn phần tử ĐANG hiển thị
                for e in els:
                    if e.is_displayed():
                        search_el = e
                        break
                if search_el:
                    break
            except Exception:
                continue

        if not search_el:
            raise TimeoutException("Không tìm thấy ô tìm kiếm thông minh (global-search-input).")

        # 4) Bơm record_id và nhấn Enter bằng JavaScript (tránh yêu cầu clickable)
        driver.execute_script("""
const el = arguments[0];
el.focus();
try { el.select && el.select(); } catch(e) {}
el.value = arguments[1];
el.dispatchEvent(new Event('input', {bubbles: true}));
el.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter', code:'Enter', bubbles:true}));
el.dispatchEvent(new KeyboardEvent('keyup',   {key:'Enter', code:'Enter', bubbles:true}));
        """, search_el, record_id)

        # 5) Chờ panel kết quả và click đúng ID
        try:
            result_link = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, f"//a[normalize-space()='{record_id}']"))
            )
        except Exception:
            result_link = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, f"//a[contains(.,'{record_id}')]"))
            )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", result_link)
        driver.execute_script("arguments[0].click();", result_link)
        time.sleep(1.0)

        # 6) Xem trước mẫu in -> chọn Phiếu TTTT - Nhà đất -> Tải xuống
        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[contains(.,'Xem trước mẫu in')]"
        ))).click()
        time.sleep(0.8)

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//div[contains(.,'Phiếu TTTT - Nhà đất')]"
        ))).click()
        time.sleep(0.6)

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[contains(.,'Tải xuống')]"
        ))).click()

        template_path = _wait_for_docx(download_dir, timeout=60)

        # 7) Tải một vài ảnh (có thể chỉnh selector thumbnail sau)
        images = _download_images_from_detail(driver, download_dir)

        return template_path, images

    finally:
        driver.quit()

# ===================== Helpers =====================

def _wait_for_docx(folder: str, timeout: int = 60) -> str:
    for _ in range(timeout):
        for f in os.listdir(folder):
            if f.lower().endswith(".docx"):
                return os.path.join(folder, f)
        time.sleep(1)
    raise FileNotFoundError("Không thấy file .docx sau khi tải")

def _download_images_from_detail(driver, download_dir: str) -> List[str]:
    images: List[str] = []
    thumbs = driver.find_elements(By.CSS_SELECTOR, "img")
    for i, t in enumerate(thumbs[:8], start=1):
        try:
            src = t.get_attribute("src")
            if src and src.startswith("http"):
                r = requests.get(src, timeout=15)
                p = os.path.join(download_dir, f"image_{i}.jpg")
                with open(p, "wb") as f:
                    f.write(r.content)
                images.append(p)
        except Exception:
            pass
    return images

# ===================== Word processing =====================

def fill_document(template_path: str, images: List[str],
                  signature_path: str, output_path: str) -> None:
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    doc = Document(template_path)

    slot_map = {
        "Thông tin rao bán/sổ đỏ": images[0:1],
        "Mặt trước tài sản":       images[1:2],
        "Tổng thể tài sản":        images[2:3],
        "Đường phía trước tài sản": images[3:5],
        "Ảnh khác":                images[5:7],
    }

    def _table_has_phu_luc(tbl):
        text = "\n".join(cell.text for row in tbl.rows for cell in row.cells)
        return "Phụ lục" in text and "Ảnh TSSS" in text

    target_table = None
    for tbl in doc.tables:
        if _table_has_phu_luc(tbl):
            target_table = tbl; break

    if target_table:
        for row in target_table.rows:
            for ci, cell in enumerate(row.cells):
                label = cell.text.strip()
                if label in slot_map and ci+1 < len(row.cells):
                    dest = row.cells[ci+1]
                    for p in dest.paragraphs:
                        if p.text: p.text = ""
                    for pth in slot_map[label]:
                        if os.path.exists(pth):
                            dest.paragraphs[0].add_run().add_picture(pth, width=Inches(2.2))
                            dest.add_paragraph("")
    else:
        doc.add_heading("Phụ lục Ảnh TSSS", level=2)
        for pth in images:
            if os.path.exists(pth):
                doc.add_picture(pth, width=Inches(3))
                doc.add_paragraph("")

    doc.add_page_break()
    doc.add_heading("Chữ ký cán bộ khảo sát", level=2)
    if signature_path and os.path.exists(signature_path):
        try:
            doc.add_picture(signature_path, width=Inches(2))
        except Exception as e:
            doc.add_paragraph(f"[Không chèn được chữ ký: {e}]")
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
