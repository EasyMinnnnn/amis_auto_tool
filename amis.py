"""Helper module for AMIS automation and document manipulation.

Bản này bỏ hẳn bước *tìm kiếm UI*:
- Nếu có `execution_id`: vào thẳng trang chi tiết bằng URL.
- Nếu chỉ có `record_id`: gọi API GlobalSearch để đổi sang `execution_id` (dùng cookie sau đăng nhập).
Sau đó tự động mở **Xem trước mẫu in → Phiếu TTTT - Nhà đất → Tải xuống** với logic click
được làm robust hơn (thử nhiều selector và quét qua iframe nếu có).
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

    drv = webdriver.Chrome(options=opts)
    return drv


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


# ===================== Requests session từ Selenium =====================

def _requests_session_from_driver(driver) -> requests.Session:
    """
    Lấy cookie sau khi đã login bằng Selenium và chuyển sang requests.Session
    để gọi API nội bộ (GlobalSearch/Export...).
    """
    s = requests.Session()
    for c in driver.get_cookies():
        try:
            s.cookies.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path", "/"))
        except Exception:
            pass
    s.headers.update({
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "https://amisapp.misa.vn",
        "Referer": "https://amisapp.misa.vn/process/execute/1",
        "User-Agent": "Mozilla/5.0",
    })
    return s


# ===================== GlobalSearch: record_id -> execution_id =====================

def _resolve_execution_id_via_api(
    session: requests.Session,
    record_id: str,
    tenant_code: str,
    timeout: int = 30,
) -> Tuple[str, int]:
    """
    Dùng API GlobalSearch để đổi Record ID (vd 80002) -> (execution_id, status).
    Không dùng UI.
    """
    base = "https://amisapp.misa.vn"
    url = f"{base}/process/APIS/g2/ProcessAPI/api/AvaBot/GlobalSearch"
    params = {"pageSize": 3, "pageIndex": 0, "dateQuery": ""}

    candidate_bodies = [
        {"Keyword": str(record_id), "companyCode": tenant_code},
        {"keyword": str(record_id), "companyCode": tenant_code},
        {"Text":    str(record_id), "companyCode": tenant_code},
        {"SearchText": str(record_id), "companyCode": tenant_code},
        {"Keyword": str(record_id)},  # fallback nếu body không cần companyCode
    ]

    last_err = None
    for body in candidate_bodies:
        try:
            r = session.post(url, params=params, json=body, timeout=timeout)
            if r.status_code != 200:
                last_err = f"HTTP {r.status_code}"
                continue
            data = r.json()
            items = ((data or {}).get("Data") or {}).get("Items") or []
            for it in items:
                if str(it.get("ProcessExecutionCode")) == str(record_id):
                    exec_id = it.get("ProcessExecutionID")
                    status = it.get("Status", 1)
                    if exec_id:
                        return str(exec_id), int(status)
        except Exception as e:
            last_err = str(e)

    raise RuntimeError(
        f"Không tìm thấy executionId cho record_id={record_id} qua API GlobalSearch. "
        f"Chi tiết cuối: {last_err}"
    )


# ===================== Helpers: tìm/click robust, quét iframe =====================

def _all_frames(driver):
    """Danh sách tất cả frame (None = default)."""
    frames = [None]
    try:
        ifrs = driver.find_elements(By.TAG_NAME, "iframe")
        frames.extend(ifrs)
    except Exception:
        pass
    return frames


def _with_each_frame(driver, func):
    """Chạy func trong default_content và từng iframe; trả về (element, frame) đầu tiên tìm thấy."""
    for fr in _all_frames(driver):
        try:
            driver.switch_to.default_content()
            if fr is not None:
                driver.switch_to.frame(fr)
            el = func()
            if el:
                return el, fr
        except Exception:
            continue
    driver.switch_to.default_content()
    return None, None


def _find_clickable_by_xpaths(driver, xpaths: List[str], timeout: int = 15):
    def finder():
        for xp in xpaths:
            try:
                el = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, xp))
                )
                return el
            except Exception:
                pass
        return None

    end = time.time() + timeout
    while time.time() < end:
        el, fr = _with_each_frame(driver, finder)
        if el:
            return el, fr
        time.sleep(0.3)
    return None, None


def _click_text(driver, texts: List[str], timeout: int = 20) -> bool:
    """Click vào phần tử có text khớp (thử nhiều biến thể) – quét qua mọi iframe."""
    xpaths = []
    for t in texts:
        # Tối ưu hóa normalize-space và contains (xử lý biến thể font/span)
        xpaths.extend([
            f"//button[normalize-space()='{t}']",
            f"//span[normalize-space()='{t}']/ancestor::button",
            f"//span[contains(normalize-space(),'{t}')]/ancestor::button",
            f"//*[self::button or self::a or self::span or self::div][contains(normalize-space(),'{t}')]",
        ])
    el, fr = _find_clickable_by_xpaths(driver, xpaths, timeout=timeout)
    if not el:
        return False
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        try:
            el.click()
            return True
        except Exception:
            return False


# ===================== MAIN =====================

def run_automation(
    username: str,
    password: str,
    download_dir: str,
    headless: bool = True,
    record_id: Optional[str] = None,       # chỉ nhập Record ID cũng được
    execution_id: Optional[str] = None,    # hoặc đưa thẳng execution ID (ID=... trong URL)
    status: Optional[int] = None,          # nếu không có, sẽ lấy từ API search (khi dùng record_id)
    company_code: str = "RH7VZQAQ",
) -> Tuple[str, List[str]]:
    """
    Đăng nhập AMIS, vào thẳng trang chi tiết (theo execution_id hoặc theo record_id->API),
    sau đó mở Xem trước mẫu in -> chọn Phiếu TTTT - Nhà đất -> Tải xuống.
    Trả về: (đường dẫn file Word, danh sách ảnh trên trang).
    """
    if not (execution_id or record_id):
        raise ValueError("Cần truyền execution_id hoặc record_id.")

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

        # 2) Nếu chưa có execution_id thì đổi từ record_id qua API (không dùng UI)
        if not execution_id:
            session = _requests_session_from_driver(driver)
            execution_id, api_status = _resolve_execution_id_via_api(
                session=session,
                record_id=str(record_id),
                tenant_code=company_code,
            )
            if status is None:
                status = api_status
        if status is None:
            status = 1

        # 3) Vào thẳng trang chi tiết
        detail_url = (
            f"https://amisapp.misa.vn/process/execute/1"
            f"?ID={execution_id}&type=1&status={status}&appCode=AMISProcess&companyCode={company_code}"
        )
        driver.get(detail_url)
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#top_nav"))
            )
        except Exception:
            pass
        time.sleep(0.8)

        # 4) Xem trước mẫu in -> chọn Phiếu TTTT - Nhà đất -> Tải xuống
        #    (click robust: thử nhiều selector / iframe)
        # 4.1 Click "Xem trước mẫu in"
        if not _click_text(driver, ["Xem trước mẫu in", "Xem trước", "Mẫu in"], timeout=25):
            _dump_debug(driver, download_dir, "cannot_click_preview")
            raise TimeoutException("Không tìm thấy/không click được nút 'Xem trước mẫu in'.")

        time.sleep(0.8)

        # 4.2 Chọn "Phiếu TTTT - Nhà đất"
        if not _click_text(
            driver,
            ["Phiếu TTTT - Nhà đất", "TTTT - Nhà đất", "Phiếu TTTT"],
            timeout=25,
        ):
            _dump_debug(driver, download_dir, "cannot_pick_template")
            raise TimeoutException("Không chọn được mẫu 'Phiếu TTTT - Nhà đất'.")

        time.sleep(0.6)

        # 4.3 Click "Tải xuống"
        if not _click_text(driver, ["Tải xuống", "Tải về"], timeout=25):
            _dump_debug(driver, download_dir, "cannot_click_download")
            raise TimeoutException("Không click được nút 'Tải xuống'.")

        template_path = _wait_for_docx(download_dir, timeout=120)

        # 5) Tải một số ảnh hiển thị trên trang chi tiết (best-effort)
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
                # Lấy tên từ header nếu có
                cd = r.headers.get("Content-Disposition", "")
                if cd:
                    m = re.search(r'filename="?([^"]+)"?', cd)
                else:
                    m = None
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
