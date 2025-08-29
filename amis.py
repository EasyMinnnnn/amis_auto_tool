"""Helper module for AMIS automation and document manipulation."""

import os, time, base64
from typing import List, Tuple, Optional, Callable
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, JavascriptException

from docx import Document
from docx.shared import Inches


# =============== Utilities ===============

def _make_driver(download_dir: str, headless: bool = True) -> webdriver.Chrome:
    os.makedirs(download_dir, exist_ok=True)
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
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

def _safe_click(driver, el):
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)

def _switch_to_last_tab(driver: webdriver.Chrome):
    try:
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
    except Exception:
        pass
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

def _dump_debug(driver, download_dir: str, tag: str):
    try:
        with open(os.path.join(download_dir, f"debug_{tag}.html"), "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    except Exception:
        pass
    try:
        path = os.path.join(download_dir, f"debug_{tag}.png")
        driver.save_screenshot(path)
    except Exception:
        pass
    # Liệt kê iframe src
    try:
        frames = driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
        with open(os.path.join(download_dir, f"frames_{tag}.txt"), "w", encoding="utf-8") as f:
            for i, fr in enumerate(frames, 1):
                try:
                    src = fr.get_attribute("src")
                except Exception:
                    src = ""
                try:
                    name = fr.get_attribute("name")
                except Exception:
                    name = ""
                f.write(f"{i}. name={name} src={src}\n")
    except Exception:
        pass

# =============== Find in iframes & shadow DOM ===============

def _for_each_frame(driver: webdriver.Chrome, fn: Callable[[], Optional[object]]) -> Optional[object]:
    # thử ở current
    try:
        res = fn()
        if res is not None:
            return res
    except Exception:
        pass
    # duyệt iframe
    frames = driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
    for idx in range(len(frames)):
        try:
            driver.switch_to.frame(idx)
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

def _query_shadow_dom(driver, js_selector: str):
    """
    Đi xuyên shadow DOM mở. js_selector là chuỗi JS:
    ví dụ: "return (function(){ const hosts=[...document.querySelectorAll('some-host')]; ...; return el; })();"
    """
    try:
        return driver.execute_script(js_selector)
    except JavascriptException:
        return None

def _find_search_shadow(driver):
    # Thử path shadow dựa trên gợi ý của bạn trước đó (#top_nav ... textarea)
    # và thử cả input với placeholder có 'Tìm kiếm'
    scripts = [
        # theo path cụ thể (nếu global-search dùng shadow host)
        """
        return (function() {
            const pick = (root) => root && (root.querySelector("textarea[placeholder*='Tìm kiếm']") ||
                                            root.querySelector("input[placeholder*='Tìm kiếm']") ||
                                            root.querySelector("input[placeholder*='Nhấn Enter']"));
            const walk = (node) => {
                if (!node) return null;
                let hit = pick(node);
                if (hit) return hit;
                const hosts = node.querySelectorAll("*");
                for (const h of hosts) {
                    if (h.shadowRoot) {
                        let el = pick(h.shadowRoot);
                        if (el) return el;
                        el = walk(h.shadowRoot);
                        if (el) return el;
                    }
                }
                return null;
            };
            return walk(document);
        })();
        """,
    ]
    for js in scripts:
        el = _query_shadow_dom(driver, js)
        if el:
            return el
    return None

def _find_in_any_frame_or_shadow(driver: webdriver.Chrome, candidates: List[tuple]) -> Optional[object]:
    # Bắt đầu ở default
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # 1) thử thẳng DOM thường
    for by, sel in candidates:
        els = driver.find_elements(by, sel)
        if els:
            return els[0]

    # 2) thử shadow DOM mở
    el = _find_search_shadow(driver)
    if el:
        return el

    # 3) thử toàn bộ iframe (bên trong mỗi frame cũng thử DOM thường + shadow)
    def finder():
        # DOM thường
        for by, sel in candidates:
            try:
                el2 = driver.find_element(by, sel)
                return el2
            except NoSuchElementException:
                continue
        # shadow
        return _find_search_shadow(driver)

    return _for_each_frame(driver, finder)

def _wait_in_any_frame_or_shadow(driver: webdriver.Chrome, candidates: List[tuple], timeout: int = 35) -> object:
    end = time.time() + timeout
    while time.time() < end:
        el = _find_in_any_frame_or_shadow(driver, candidates)
        if el is not None:
            return el
        time.sleep(0.5)
    raise TimeoutException("Timeout waiting element (any frame/shadow)")

# =============== App-specific small helpers ===============

def _close_popups_soft(driver):
    for by, sel in [
        (By.XPATH, "//button[contains(.,'Bỏ qua') or contains(.,'Tiếp tục làm việc') or contains(.,'Đóng')]"),
        (By.CSS_SELECTOR, "[aria-label='Close'],[data-dismiss],.close"),
    ]:
        try:
            el = driver.find_element(by, sel)
            _safe_click(driver, el)
            time.sleep(0.2)
        except Exception:
            pass

def _try_open_modules(driver):
    # Thử click các nút/tab hay gặp để vào đúng màn hình quy trình/lượt chạy
    for by, sel in [
        (By.XPATH, "//a[contains(.,'Quy trình')]"),
        (By.XPATH, "//*[contains(.,'Lượt chạy')]"),
        (By.XPATH, "//button//*[name()='svg' or contains(.,'Tìm kiếm')]/ancestor::button[1]"),
    ]:
        try:
            el = driver.find_element(by, sel)
            _safe_click(driver, el)
            time.sleep(0.4)
        except Exception:
            pass

# =============== Main ===============

def run_automation(username: str, password: str, record_id: str,
                   download_dir: str, headless: bool = True) -> Tuple[str, List[str]]:

    driver = _make_driver(download_dir, headless=headless)
    wait = WebDriverWait(driver, 40)

    try:
        # 1) Login
        driver.get("https://amisapp.misa.vn/")
        username_input = wait.until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "#box-login-right > div > div > div.login-form-basic-container > "
            "div > div.login-form-inputs.login-class > "
            "div.wrap-input.username-wrap.validate-input > input",
        )))
        username_input.send_keys(username)

        password_input = driver.find_element(
            By.CSS_SELECTOR,
            "#box-login-right > div > div > div.login-form-basic-container > "
            "div > div.login-form-inputs.login-class > "
            "div.wrap-input.pass-wrap.validate-input > input",
        )
        password_input.send_keys(password)

        driver.find_element(
            By.CSS_SELECTOR,
            "#box-login-right > div > div > div.login-form-basic-container > "
            "div > div.login-form-btn-container.login-class > button",
        ).click()

        wait.until(EC.url_contains("amisapp.misa.vn"))
        time.sleep(0.8)
        _close_popups_soft(driver)

        # 2) Điều hướng vào quy trình / lượt chạy
        driver.get("https://amisapp.misa.vn/process/execute/1")
        time.sleep(0.6)
        _switch_to_last_tab(driver)
        _close_popups_soft(driver)
        _try_open_modules(driver)

        # 3) Tìm ô tìm kiếm (input/textarea/role search) — DOM thường, shadow DOM, hoặc iframe
        search = _wait_in_any_frame_or_shadow(
            driver,
            candidates=[
                (By.CSS_SELECTOR, "input[placeholder*='Tìm kiếm']"),
                (By.CSS_SELECTOR, "input[placeholder*='Nhấn Enter']"),
                (By.CSS_SELECTOR, "input[type='search']"),
                (By.XPATH, "//input[contains(@placeholder,'Tìm kiếm') or contains(@placeholder,'Nhấn Enter')]"),
                (By.XPATH, "//*[@role='search']//input"),
                (By.XPATH, "//textarea[contains(@placeholder,'Tìm kiếm')]"),
            ],
            timeout=40,
        )

        # 4) Nhập record_id & Enter
        try:
            search.clear()
        except Exception:
            search.send_keys(Keys.CONTROL, "a")
            search.send_keys(Keys.BACK_SPACE)
        search.send_keys(record_id)
        search.send_keys(Keys.ENTER)

        # 5) Chọn dòng đầu tiên
        first_row = _wait_in_any_frame_or_shadow(
            driver,
            candidates=[
                (By.CSS_SELECTOR, "table tbody tr"),
                (By.XPATH, "(//table//tbody//tr)[1]"),
            ],
            timeout=40,
        )
        _safe_click(driver, first_row)
        time.sleep(0.8)

        # 6) Xem trước mẫu in -> chọn Phiếu TTTT - Nhà đất -> Tải xuống
        btn_preview = _wait_in_any_frame_or_shadow(
            driver, [(By.XPATH, "//span[contains(.,'Xem trước mẫu in')]")], timeout=40
        )
        _safe_click(driver, btn_preview)
        time.sleep(0.6)

        item_phieu = _wait_in_any_frame_or_shadow(
            driver, [(By.XPATH, "//div[contains(.,'Phiếu TTTT - Nhà đất')]")], timeout=40
        )
        _safe_click(driver, item_phieu)
        time.sleep(0.6)

        btn_download = _wait_in_any_frame_or_shadow(
            driver, [(By.XPATH, "//button[contains(.,'Tải xuống')]")], timeout=40
        )
        _safe_click(driver, btn_download)

        template_path = _wait_for_docx(download_dir, timeout=60)

        # 7) Ảnh (placeholder)
        images = _download_images_from_detail(driver, download_dir)

        return template_path, images

    except Exception as e:
        # Ghi debug trước khi raise để bạn xem được trên Streamlit
        _dump_debug(driver, download_dir, "fail")
        raise
    finally:
        driver.quit()

# =============== Download helpers ===============

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

# =============== Word processing ===============

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
