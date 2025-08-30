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

        # 3) Mở ô tìm kiếm theo đúng thao tác click → gõ → Enter,
        #    ưu tiên dựa vào #top_nav và activeElement thay vì locator cụ thể.

        def _click_search_area_once() -> None:
            # Cuộn về top để chắc chắn vùng top_nav visible
            try:
                driver.execute_script("window.scrollTo(0,0);")
            except Exception:
                pass

            # Các vị trí có khả năng mở hộp search
            click_targets = [
                (By.CSS_SELECTOR, "#top_nav .global-search-wrap"),
                (By.CSS_SELECTOR, ".global-search-wrap"),
                (By.CSS_SELECTOR, "#top_nav [aria-label*='Tìm kiếm']"),
                (By.CSS_SELECTOR, "#top_nav [title*='Tìm kiếm']"),
                (By.CSS_SELECTOR, "[aria-label*='Search']"),
                (By.CSS_SELECTOR, "[title*='Search']"),
            ]
            for by, sel in click_targets:
                try:
                    el = driver.find_element(by, sel)
                    ActionChains(driver).move_to_element(el).pause(0.05).click(el).perform()
                    time.sleep(0.15)
                    break
                except Exception:
                    continue

        def _try_keyboard_shortcuts() -> None:
            try:
                body = driver.find_element(By.TAG_NAME, "body")
                # nhiều app bật search bằng '/', Ctrl+K, Ctrl+F
                for combo in [
                    ("/",),
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

        def _is_typable(el) -> bool:
            try:
                tag = el.tag_name.lower()
            except Exception:
                return False
            if tag in ("textarea", "input"):
                return el.is_enabled() and el.is_displayed()
            # contenteditable
            try:
                ce = el.get_attribute("contenteditable")
                if ce and ce.lower() == "true":
                    return el.is_displayed()
            except Exception:
                pass
            return False

        def _focus_typable_via_js():
            # Tìm phần tử gõ được bằng JS và focus (kể cả trong iframe same-origin)
            js = """
return (function(){
  function isVisible(el){
    if (!el) return false;
    const st = getComputedStyle(el);
    if (!st) return false;
    if (st.display === 'none' || st.visibility === 'hidden' || st.opacity === '0') return false;
    const r = el.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) return false;
    return true;
  }
  function findTypable(doc){
    try{
      const qs = [
        'textarea.global-search-input',
        'textarea[placeholder*="Tìm kiếm"]',
        'textarea[placeholder*="Search"]',
        'input.global-search-input',
        'input[type="search"]',
        '[contenteditable="true"]'
      ];
      for (const sel of qs){
        const list = doc.querySelectorAll(sel);
        for (const el of list){
          if (isVisible(el)) return el;
        }
      }
    }catch(e){}
    return null;
  }
  function crawl(win, depth){
    if (depth > 4) return null;
    const el = findTypable(win.document);
    if (el){ el.focus(); return el; }
    const ifrs = win.document.querySelectorAll('iframe');
    for (const f of ifrs){
      try{
        const child = f.contentWindow;
        const got = crawl(child, depth+1);
        if (got) return got;
      }catch(e){}
    }
    return null;
  }
  return crawl(window, 0);
})();
"""
            try:
                return driver.execute_script(js)
            except Exception:
                return None

        def _get_active() :
            try:
                return driver.switch_to.active_element
            except Exception:
                return None

        def _obtain_search_focus(timeout_s=55):
            end = time.time() + timeout_s
            while time.time() < end:
                # 1) click vùng search
                _click_search_area_once()
                # 2) thử lấy activeElement
                el = _get_active()
                if el and _is_typable(el):
                    return el
                # 3) thử JS tìm phần tử có thể gõ
                el = _focus_typable_via_js()
                if el and _is_typable(el):
                    return el
                # 4) thử phím tắt
                _try_keyboard_shortcuts()
                el = _get_active()
                if el and _is_typable(el):
                    return el
                # 5) Tab một cái để nhảy vào input nếu đang trong overlay
                try:
                    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.TAB)
                    time.sleep(0.08)
                    el = _get_active()
                    if el and _is_typable(el):
                        return el
                except Exception:
                    pass
                time.sleep(0.25)
            return None

        search_el = _obtain_search_focus(timeout_s=55)

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
                if (!el) return;
                el.focus && el.focus();
                const tag = (el.tagName||'').toLowerCase();
                if (tag === 'textarea' || tag === 'input') {
                  el.value = val;
                } else {
                  el.textContent = val;
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
