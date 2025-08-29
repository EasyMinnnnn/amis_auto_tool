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
    Đăng nhập AMIS, tìm record_id bằng ô 'Tìm kiếm thông minh với AI', tải file Word và ảnh.
    """
    driver = _make_driver(download_dir, headless=headless)
    wait = WebDriverWait(driver, 30)

    try:
        # 1) Login
        driver.get("https://amisapp.misa.vn/")
        user_el = wait.until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "#box-login-right .login-form-inputs .username-wrap input",
        )))
        user_el.send_keys(username)

        pw_el = driver.find_element(By.CSS_SELECTOR,
            "#box-login-right .login-form-inputs .pass-wrap input")
        pw_el.send_keys(password)

        driver.find_element(By.CSS_SELECTOR,
            "#box-login-right .login-form-btn-container button").click()

        wait.until(EC.url_contains("amisapp.misa.vn"))
        time.sleep(0.8)

        # Có popup "Bỏ qua, tiếp tục làm việc" thì đóng (best-effort)
        for by, sel in [
            (By.XPATH, "//button[contains(.,'Bỏ qua')]"),
            (By.XPATH, "//button[contains(.,'Tiếp tục làm việc')]"),
            (By.XPATH, "//button[contains(.,'Đóng')]"),
        ]:
            try:
                driver.find_element(by, sel).click(); time.sleep(0.2)
            except Exception:
                pass

        # 2) Vào trang Quy trình (Lượt chạy)
        driver.get("https://amisapp.misa.vn/process/execute/1")
        time.sleep(0.8)

        # 3) Gõ ID vào ô tìm kiếm thông minh (textarea.global-search-input) rồi Enter
        search = wait.until(EC.element_to_be_clickable((
            By.CSS_SELECTOR, "textarea.global-search-input"
        )))
        try:
            search.clear()
        except Exception:
            search.send_keys(Keys.CONTROL, "a"); search.send_keys(Keys.BACK_SPACE)

        search.send_keys(record_id)
        search.send_keys(Keys.ENTER)

        # 4) Click kết quả mang đúng ID (panel “Kết quả tìm kiếm…”)
        #    a) ưu tiên link hiển thị chính xác ID
        try:
            result_link = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((
                By.XPATH, f"//a[normalize-space()='{record_id}']"
            )))
        except Exception:
            #    b) fallback: link chứa ID trong text
            result_link = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((
                By.XPATH, f"//a[contains(.,'{record_id}')]"
            )))
        driver.execute_script("arguments[0].click();", result_link)
        time.sleep(1.0)

        # 5) Xem trước mẫu in -> chọn Phiếu TTTT - Nhà đất -> Tải xuống
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

        # 6) Tải một vài ảnh minh họa (tùy chỉnh sau nếu cần)
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
                with open(p, "wb") as f: f.write(r.content)
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
