"""Helper module for AMIS automation and document manipulation.

Phiên bản tối giản:
- KHÔNG dùng tìm kiếm UI
- KHÔNG gọi API GlobalSearch
- Chỉ cần ID trong URL (execution_id). Nếu app cũ truyền record_id, sẽ coi như execution_id.
"""

import os
import time
from typing import List, Tuple, Optional
import requests
import re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
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
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1440,900")
    # opts.binary_location = "/usr/bin/chromium"  # nếu môi trường yêu cầu

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": True,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=opts)


# ===================== Debug helpers =====================

def _dump_debug(driver, out_dir: str, tag: str) -> None:
    """Lưu screenshot + HTML để debug khi fail."""
    try:
        os.makedirs(out_dir, exist_ok=True)
        driver.save_screenshot(os.path.join(out_dir, f"debug_{tag}.png"))
        with open(os.path.join(out_dir, f"debug_{tag}.html"), "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    except Exception:
        pass


# ===================== Click helpers =====================

def _click_text(driver, texts: List[str], timeout: int = 25) -> bool:
    """Click phần tử có text (thử nhiều XPaths)"""
    xpaths = []
    for t in texts:
        xpaths.extend([
            f"//button[normalize-space()='{t}']",
            f"//span[normalize-space()='{t}']/ancestor::button",
            f"//*[self::button or self::a or self::span or self::div][contains(normalize-space(),'{t}')]",
        ])
    end = time.time() + timeout
    while time.time() < end:
        for xp in xpaths:
            try:
                el = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, xp))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                driver.execute_script("arguments[0].click();", el)
                return True
            except Exception:
                pass
        time.sleep(0.3)
    return False


# ===================== MAIN =====================

def run_automation(
    username: str,
    password: str,
    download_dir: str,
    headless: bool = True,
    record_id: Optional[str] = None,       # Tương thích ngược: coi như execution_id
    execution_id: Optional[str] = None,    # Nên dùng: ID trong URL
    status: int = 1,                       # Theo URL chi tiết
    company_code: str = "RH7VZQAQ",        # Tenant/Company code
) -> Tuple[str, List[str]]:
    """
    Đăng nhập AMIS, điều hướng thẳng tới trang chi tiết bằng execution_id (ID trong URL),
    rồi mở 'Xem trước mẫu in' → chọn 'Phiếu TTTT - Nhà đất' → 'Tải xuống'.

    Trả về: (đường dẫn file .docx, danh sách ảnh best-effort).
    """
    # Nếu app cũ truyền record_id, coi như execution_id để không phải sửa app.py
    if not execution_id and record_id:
        execution_id = str(record_id)

    if not execution_id:
        raise ValueError("Cần truyền execution_id (hoặc record_id sẽ được coi như execution_id).")

    driver = _make_driver(download_dir, headless=headless)
    wait = WebDriverWait(driver, 90)

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
        time.sleep(1.0)

        # Đóng các popup nhẹ nếu có (best-effort)
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

        # 2) Vào thẳng trang chi tiết
        detail_url = (
            "https://amisapp.misa.vn/process/execute/1"
            f"?ID={execution_id}&type=1&status={status}&appCode=AMISProcess&companyCode={company_code}"
        )
        driver.get(detail_url)
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#top_nav"))
            )
        except Exception:
            pass
        time.sleep(0.6)

        # 3) Xem trước mẫu in → chọn mẫu → Tải xuống
        if not _click_text(driver, ["Xem trước mẫu in", "Xem trước", "Mẫu in"], timeout=30):
            _dump_debug(driver, download_dir, "cannot_click_preview")
            raise TimeoutException("Không click được 'Xem trước mẫu in'.")

        time.sleep(0.6)

        if not _click_text(driver, ["Phiếu TTTT - Nhà đất", "TTTT - Nhà đất", "Phiếu TTTT"], timeout=30):
            _dump_debug(driver, download_dir, "cannot_pick_template")
            raise TimeoutException("Không chọn được mẫu 'Phiếu TTTT - Nhà đất'.")

        time.sleep(0.5)

        if not _click_text(driver, ["Tải xuống", "Tải về", "Download"], timeout=30):
            _dump_debug(driver, download_dir, "cannot_click_download")
            raise TimeoutException("Không click được nút 'Tải xuống'.")

        # 4) Chờ file .docx xuất hiện
        template_path = _wait_for_docx(download_dir, timeout=120)

        # 5) Lấy ảnh minh hoạ (best-effort)
        images = _download_images_from_detail(driver, download_dir)

        return template_path, images

    finally:
        driver.quit()


# ===================== Helpers =====================

def _wait_for_docx(folder: str, timeout: int = 120) -> str:
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
                # Nếu có tên file trong header thì dùng, không thì đặt mặc định
                cd = r.headers.get("Content-Disposition", "")
                m = re.search(r'filename="?([^"]+)"?', cd) if cd else None
                name = m.group(1) if m else f"image_{i}.jpg"
                p = os.path.join(download_dir, name)
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
        return ("Phụ lục" in text) and ("Ảnh TSSS" in text)

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
