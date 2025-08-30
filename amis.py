"""Helper module for AMIS automation and document manipulation."""

import os
import time
from typing import List, Tuple
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from docx import Document
from docx.shared import Inches


# ===================== Selenium base =====================

def _make_driver(download_dir: str, headless: bool = True) -> webdriver.Chrome:
    os.makedirs(download_dir, exist_ok=True)
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    # Các flag an toàn phổ biến cho môi trường container/Cloud:
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1440,900")
    # Nếu chắc chắn chromium ở path cụ thể thì mở dòng dưới (Cloud thường KHÔNG cần):
    # opts.binary_location = "/usr/bin/chromium"

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
    """Lưu screenshot + HTML + info frame để debug khi fail."""
    try:
        os.makedirs(out_dir, exist_ok=True)
        png = os.path.join(out_dir, f"debug_{tag}.png")
        html = os.path.join(out_dir, f"debug_{tag}.html")
        info = os.path.join(out_dir, f"debug_{tag}_info.txt")

        driver.save_screenshot(png)
        with open(html, "w", encoding="utf-8") as f:
            f.write(driver.page_source)

        with open(info, "w", encoding="utf-8") as f:
            try:
                url = driver.execute_script("return document.location.href;")
            except Exception:
                url = "(cannot read document.location.href)"
            try:
                title = driver.title
            except Exception:
                title = "(no title)"
            f.write(f"URL: {url}\nTITLE: {title}\n")
            try:
                frames = driver.find_elements(By.TAG_NAME, "iframe")
                f.write(f"IFRAME count: {len(frames)}\n")
            except Exception:
                pass
    except Exception:
        pass


# ===================== MAIN =====================

def run_automation(
    username: str,
    password: str,
    record_id: str,
    download_dir: str,
    headless: bool = True,
) -> Tuple[str, List[str]]:
    """
    Đăng nhập AMIS, nhập ID vào ô 'Tìm kiếm thông minh với AI', mở chi tiết,
    tải file Word, lấy ảnh.
    Trả về: (đường dẫn file Word, danh sách ảnh).
    """
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
        # chờ top nav có mặt (UI chính đã tải)
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#top_nav"))
            )
        except Exception:
            pass
        time.sleep(1.0)

        # 3) Tìm ô tìm kiếm theo đúng thao tác click → gõ → Enter

        CSS_CANDIDATES = [
            # chính xác theo lớp bạn cung cấp:
            "textarea.global-search-input",
            # selector đầy đủ bạn gửi (để bám đúng đường dẫn hiện tại):
            "#top_nav > div.flex-m.w-full > div.global-search-wrap.active-feature-search > div > "
            "div.flex-t.global-search.global-search-root > div.p-t-4.p-l-2.p-b-4.w-full > textarea",
            # fallback thêm:
            "div.global-search-wrap textarea.global-search-input",
            "div.global-search-root textarea.global-search-input",
            "textarea[placeholder*='Tìm kiếm']",
            "textarea[placeholder*='Search']",
        ]
        XPATH_CANDIDATES = [
            # XPath bạn gửi:
            "/html/body/div[2]/div/div[2]/div/div[2]/div[3]/div/div[1]/div[3]/textarea",
            # fallback tổng quát:
            "//textarea[contains(@class,'global-search-input')]",
            "//div[contains(@class,'global-search-wrap')]//textarea",
            "//textarea[contains(@placeholder,'Tìm kiếm')]",
            "//textarea[contains(@placeholder,'Search')]",
        ]

        def _try_open_search_ui():
            # Đảm bảo top header thấy được
            try:
                driver.execute_script("window.scrollTo(0,0);")
            except Exception:
                pass

            # Click các vùng có thể mở hộp search
            for by, sel in [
                (By.CSS_SELECTOR, "#top_nav .global-search-wrap"),
                (By.CSS_SELECTOR, ".global-search-wrap"),
                (By.CSS_SELECTOR, "#top_nav [aria-label*='Tìm kiếm']"),
                (By.CSS_SELECTOR, "#top_nav [title*='Tìm kiếm']"),
                (By.CSS_SELECTOR, "[aria-label*='Search']"),
                (By.CSS_SELECTOR, "[title*='Search']"),
            ]:
                try:
                    el = driver.find_element(by, sel)
                    ActionChains(driver).move_to_element(el).pause(0.05).click(el).perform()
                    time.sleep(0.15)
                    break
                except Exception:
                    continue

            # Gửi phím tắt phổ biến để bật search
            try:
                body = driver.find_element(By.TAG_NAME, "body")
                for combo in [
                    ("/",),  # nhiều app bật search bằng '/'
                    (Keys.CONTROL, "k"),
                    (Keys.CONTROL, "f"),
                ]:
                    try:
                        if len(combo) == 1:
                            body.send_keys(combo[0])
                        else:
                            body.send_keys(*combo)
                        time.sleep(0.12)
                    except Exception:
                        pass
            except Exception:
                pass

        def _find_here() -> object | None:
            # thử CSS trước
            for sel in CSS_CANDIDATES:
                try:
                    el = driver.find_element(By.CSS_SELECTOR, sel)
                    if el.is_displayed() and el.is_enabled():
                        return el
                except Exception:
                    pass
            # thử XPath
            for xp in XPATH_CANDIDATES:
                try:
                    el = driver.find_element(By.XPATH, xp)
                    if el.is_displayed() and el.is_enabled():
                        return el
                except Exception:
                    pass
            return None

        def _find_in_frames(max_depth: int = 6) -> object | None:
            """DFS mọi iframe; tại mỗi ngữ cảnh: mở UI và tìm ô search."""
            _try_open_search_ui()
            found = _find_here()
            if found:
                return found

            if max_depth <= 0:
                return None

            frames = driver.find_elements(By.TAG_NAME, "iframe")
            for f in frames:
                try:
                    driver.switch_to.frame(f)
                    sub = _find_in_frames(max_depth - 1)
                    if sub:
                        return sub
                except Exception:
                    pass
                finally:
                    try:
                        driver.switch_to.parent_frame()
                    except Exception:
                        try:
                            driver.switch_to.default_content()
                        except Exception:
                            pass
            return None

        def _find_search_el(timeout_s=55):
            end = time.time() + timeout_s
            while time.time() < end:
                # luôn về root trước khi quét
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

                el = _find_in_frames(max_depth=6)
                if el:
                    return el

                time.sleep(0.3)
            return None

        search_el = _find_search_el(timeout_s=55)

        if not search_el:
            _dump_debug(driver, download_dir, "no_global_search")
            raise TimeoutException(
                "Không tìm thấy ô tìm kiếm thông minh (global-search-input). "
                "Đã lưu debug screenshot/HTML trong thư mục download."
            )

        # 4) Bơm record_id & nhấn Enter (ưu tiên send_keys; fallback JS)
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", search_el)
            driver.execute_script("arguments[0].focus();", search_el)
            try:
                search_el.clear()
            except Exception:
                search_el.send_keys(Keys.CONTROL, "a")
                search_el.send_keys(Keys.BACKSPACE)

            search_el.send_keys(str(record_id))
            time.sleep(0.1)
            search_el.send_keys(Keys.ENTER)
        except Exception:
            # Fallback JS nếu send_keys không hoạt động
            driver.execute_script("""
                const el = arguments[0], val = arguments[1];
                el.focus();
                try { el.select && el.select(); } catch(e){}
                if (el && el.tagName) {
                  const tag = el.tagName.toLowerCase();
                  if (tag === 'textarea' || tag === 'input') {
                    el.value = val;
                  } else {
                    el.textContent = val;
                  }
                }
                el.dispatchEvent(new Event('input', {bubbles:true}));
                el.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter', code:'Enter', bubbles:true}));
                el.dispatchEvent(new KeyboardEvent('keyup',   {key:'Enter', code:'Enter', bubbles:true}));
            """, search_el, str(record_id))

        # 5) Chờ panel kết quả và click đúng ID
        try:
            result_link = WebDriverWait(driver, 25).until(
                EC.presence_of_element_located((By.XPATH, f"//a[normalize-space()='{record_id}']"))
            )
        except Exception:
            result_link = WebDriverWait(driver, 25).until(
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

        template_path = _wait_for_docx(download_dir, timeout=90)

        # 7) Tải vài ảnh (placeholder; có thể chỉnh selector thumbnail sau)
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
