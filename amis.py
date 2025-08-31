"""Helper module for AMIS automation and document manipulation."""

import os
import time
from typing import List, Tuple, Optional
import requests

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
    # opts.binary_location = "/usr/bin/chromium"  # nếu cần

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


# ===================== API helpers (bỏ UI search) =====================

def _requests_session_from_driver(driver) -> requests.Session:
    """
    Lấy cookie sau khi đã login bằng Selenium và chuyển sang requests.Session
    để gọi API nội bộ (GlobalSearch/Export...).
    """
    s = requests.Session()
    # copy cookie từ webdriver sang requests
    for c in driver.get_cookies():
        try:
            s.cookies.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path", "/"))
        except Exception:
            pass
    # header cơ bản
    s.headers.update({
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "https://amisapp.misa.vn",
        "Referer": "https://amisapp.misa.vn/process/execute/1",
        "User-Agent": "Mozilla/5.0",
    })
    return s


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

    # Mỗi tenant có thể khác tên field keyword; thử lần lượt các key phổ biến.
    candidate_bodies = [
        {"Keyword": str(record_id), "companyCode": tenant_code},
        {"keyword": str(record_id), "companyCode": tenant_code},
        {"Text":    str(record_id), "companyCode": tenant_code},
        {"SearchText": str(record_id), "companyCode": tenant_code},
        # Nếu tenant không yêu cầu companyCode trong body, thử không gửi:
        {"Keyword": str(record_id)},
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


# ===================== MAIN =====================

def run_automation(
    username: str,
    password: str,
    download_dir: str,
    headless: bool = True,
    record_id: Optional[str] = None,       # bạn có thể chỉ nhập Record ID
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
    wait = WebDriverWait(driver, 60)

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
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#top_nav"))
            )
        except Exception:
            pass

        # 4) Xem trước mẫu in -> chọn Phiếu TTTT - Nhà đất -> Tải xuống
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

        template_path = _wait_for_docx(download_dir, timeout=90)

        # 5) Tải một số ảnh hiển thị trên trang chi tiết (best-effort)
        images = _download_images_from_detail(driver, download_dir)

        return template_path, images

    finally:
        driver.quit()


# ===================== Helpers =====================

def _wait_for_docx(folder: str, timeout: int = 90) -> str:
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
