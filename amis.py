"""Helper module for AMIS automation and document manipulation.

Luồng làm việc (không gọi API/tìm kiếm):
- Đăng nhập -> trang chi tiết -> bấm More (ba chấm) -> 'In mẫu thiết lập'
- Chờ popup #popupexecution -> tick mẫu thứ 3 -> 'Tải xuống mẫu in đã chọn'
"""

import os, time, re, requests
from typing import List, Tuple, Optional

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
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
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": True,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=opts)


# ===================== Utils (gọn) =====================

def _wait_css(driver, css: str, timeout: int = 15):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css))
    )

def _visible_and_clickable(driver, css: str, timeout: int = 15):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css))
    )

def _scroll_into_view_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)

def _hover_then_click(driver, el):
    try:
        ActionChains(driver).move_to_element(el).pause(0.1).click(el).perform()
    except Exception:
        _scroll_into_view_click(driver, el)

def _try_click(driver, css: str, timeout: int = 8) -> bool:
    try:
        el = _visible_and_clickable(driver, css, timeout)
        _hover_then_click(driver, el)
        return True
    except Exception:
        try:
            el = driver.find_element(By.CSS_SELECTOR, css)
            _hover_then_click(driver, el)
            return True
        except Exception:
            return False

def _dump_debug(driver, out_dir: str, tag: str) -> None:
    try:
        os.makedirs(out_dir, exist_ok=True)
        driver.save_screenshot(os.path.join(out_dir, f"debug_{tag}.png"))
        with open(os.path.join(out_dir, f"debug_{tag}.html"), "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    except Exception:
        pass


# ===================== CSS theo bạn cung cấp =====================

# (1) Nút ba chấm (More) – bỏ tiền tố #popupexecution vì lúc đó popup chưa có
MORE_BTN_STRICT = (
    "div.nav.flex.items-center.offset-title-information > "
    "div.d-flex.content-user > "
    "div.d-flex.justify-flexend.wrap-icon-more.m-t-14.more-title-execution > button > div > i"
)
MORE_BTN_RELAX = "div.d-flex.justify-flexend.wrap-icon-more.more-title-execution > button, .wrap-icon-more.more-title-execution > button"

# (2) Item “In mẫu thiết lập” trong popover
IN_MAU_STRICT = (
    "body > div.dx-overlay-wrapper.dx-popup-wrapper.dx-popover-wrapper.popover-action-process."
    "dx-popover-without-title.dx-position-bottom > div > div.dx-popup-content > div > div:nth-child(2)"
)
IN_MAU_RELAX = "div.dx-popover-wrapper.popover-action-process .dx-popup-content > div > div:nth-child(2)"

# (3) Tick checkbox mẫu thứ 3
CHECKBOX_MAU3 = (
    "#popupexecution > div.ms-popup.flex.flex-col > div.ms-popup--content-header > div:nth-child(2) > "
    "div > div > div > div.dx-popup-content > div > div > div.dx-scrollable-wrapper > "
    "div > div.dx-scrollable-content > div.dx-scrollview-content > div > div > "
    "div.list-template.h-full.pos-relative > div.h-full.p-16 > div:nth-child(3) > label > "
    "span.icon-square-check-primary.checkmark"
)

# (4) Nút “Tải xuống mẫu in đã chọn” (ô chữ xanh)
DOWNLOAD_BLUE = (
    "#popupexecution > div.ms-popup.flex.flex-col > div.ms-popup--content-header > div:nth-child(2) > "
    "div > div > div > div.dx-popup-content > div > div > div.dx-scrollable-wrapper > "
    "div > div.dx-scrollable-content > div.dx-scrollview-content > div > div > "
    "div.list-template.h-full.pos-relative > div.h-full.p-16 > div.flex.items-center.justify-between.m-b-8 > "
    "div > div.text-blue"
)

POPOVER_WRAPPER = "body div.dx-popover-wrapper.popover-action-process, body div.dx-popup-wrapper.popover-action-process"
POPUPEXECUTION = "#popupexecution"


# ===================== Open popover & chọn “In mẫu thiết lập” =====================

def _open_print_preview_via_popover(driver, download_dir: str) -> None:
    # Nếu popup preview đã mở thì thôi
    try:
        if driver.find_elements(By.CSS_SELECTOR, POPUPEXECUTION):
            return
    except Exception:
        pass

    # Đưa viewport lên đầu để tránh header che nút
    try:
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(0.2)
    except Exception:
        pass

    # (1) Click nút ba chấm
    clicked_more = _try_click(driver, MORE_BTN_STRICT, timeout=6)
    if not clicked_more:
        clicked_more = _try_click(driver, MORE_BTN_RELAX, timeout=6)
    if not clicked_more:
        _dump_debug(driver, download_dir, "cannot_click_more_button")
        # vẫn tiếp tục thử bắn “In mẫu thiết lập” ở bước dưới

    # Chờ popover hiện (nếu đã click được)
    if clicked_more:
        end = time.time() + 8
        while time.time() < end and not driver.find_elements(By.CSS_SELECTOR, POPOVER_WRAPPER):
            time.sleep(0.2)

    # (2) Click “In mẫu thiết lập”
    clicked_in_mau = _try_click(driver, IN_MAU_STRICT, timeout=5)
    if not clicked_in_mau:
        clicked_in_mau = _try_click(driver, IN_MAU_RELAX, timeout=5)

    if not clicked_in_mau:
        # Fallback JS theo text
        try:
            driver.execute_script("""
const wants = ['In mẫu thiết lập','Xem trước mẫu in','In mẫu','Mẫu in','Xem trước'];
function vis(el){if(!el)return false; const s=getComputedStyle(el);
  if(s.display==='none'||s.visibility==='hidden') return false;
  const r=el.getBoundingClientRect(); return r.width>0 && r.height>0; }
const all=document.querySelectorAll('*');
for(const el of all){ if(!vis(el)) continue; const t=(el.innerText||'').trim(); if(!t) continue;
  for(const w of wants){ if(t.includes(w)){ try{ el.scrollIntoView({block:'center'}); el.click(); return; }catch(e){} } } }
""")
        except Exception:
            pass

    # (3) Chờ popup preview xuất hiện
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, POPUPEXECUTION)))
    except Exception:
        _dump_debug(driver, download_dir, "no_popupexecution_after_in_mau")
        raise TimeoutException("Không mở được 'In mẫu thiết lập'.")


# ===================== Chọn template & tải xuống =====================

def _choose_template_and_download(driver, download_dir: str) -> str:
    _wait_css(driver, POPUPEXECUTION, timeout=20)

    # Tick mẫu thứ 3
    if not _try_click(driver, CHECKBOX_MAU3, timeout=8):
        _dump_debug(driver, download_dir, "cannot_tick_template_3")
        raise TimeoutException("Không tick được mẫu thứ 3 trong popup.")

    time.sleep(0.3)

    # Nút Tải xuống màu xanh
    if not _try_click(driver, DOWNLOAD_BLUE, timeout=8):
        # Fallback: tìm text trong #popupexecution
        try:
            driver.execute_script("""
const root=document.querySelector('#popupexecution'); if(!root) return;
function vis(el){if(!el)return false; const s=getComputedStyle(el);
  if(s.display==='none'||s.visibility==='hidden') return false;
  const r=el.getBoundingClientRect(); return r.width>0 && r.height>0; }
const wants=['Tải xuống','Tải về','Download'];
const all=root.querySelectorAll('*');
for(const el of all){ if(!vis(el)) continue; const t=(el.innerText||'').trim(); if(!t) continue;
  for(const w of wants){ if(t.includes(w)){ try{ el.scrollIntoView({block:'center'}); el.click(); return; }catch(e){} } } }
""")
        except Exception:
            pass

    return _wait_for_docx(download_dir, timeout=120)


# ===================== MAIN =====================

def run_automation(
    username: str,
    password: str,
    download_dir: str,
    headless: bool = True,
    record_id: Optional[str] = None,
    execution_id: Optional[str] = None,
    status: int = 1,
    company_code: str = "RH7VZQAQ",
) -> Tuple[str, List[str]]:
    if not execution_id and record_id:
        execution_id = str(record_id)
    if not execution_id:
        raise ValueError("Cần truyền execution_id (hoặc record_id sẽ được coi là execution_id).")

    driver = _make_driver(download_dir, headless=headless)
    wait = WebDriverWait(driver, 90)

    try:
        # 1) Login
        driver.get("https://amisapp.misa.vn/")
        user_el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#box-login-right .login-form-inputs .username-wrap input")))
        user_el.send_keys(username)
        driver.find_element(By.CSS_SELECTOR, "#box-login-right .login-form-inputs .pass-wrap input").send_keys(password)
        driver.find_element(By.CSS_SELECTOR, "#box-login-right .login-form-btn-container button").click()
        wait.until(EC.url_contains("amisapp.misa.vn"))
        time.sleep(0.6)

        # Đóng vài popup nhẹ (best-effort)
        for by, sel in [
            (By.XPATH, "//button[contains(.,'Bỏ qua')]"),
            (By.XPATH, "//button[contains(.,'Tiếp tục làm việc')]"),
            (By.XPATH, "//button[contains(.,'Đóng')]"),
            (By.CSS_SELECTOR, "[aria-label='Close'],[data-dismiss],.close"),
        ]:
            try:
                driver.find_element(by, sel).click(); time.sleep(0.1)
            except Exception:
                pass

        # 2) Vào thẳng trang chi tiết
        detail_url = (
            "https://amisapp.misa.vn/process/execute/1"
            f"?ID={execution_id}&type=1&status={status}&appCode=AMISProcess&companyCode={company_code}"
        )
        driver.get(detail_url)
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#top_nav")))
        except Exception:
            pass
        time.sleep(0.4)

        # 3) Mở In mẫu thiết lập
        _open_print_preview_via_popover(driver, download_dir)

        # 4) Tick mẫu 3 & Tải xuống
        template_path = _choose_template_and_download(driver, download_dir)

        # 5) Ảnh minh hoạ (best-effort)
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
    for i, t in enumerate(driver.find_elements(By.CSS_SELECTOR, "img")[:8], start=1):
        try:
            src = t.get_attribute("src")
            if src and src.startswith("http"):
                r = requests.get(src, timeout=15)
                cd = r.headers.get("Content-Disposition", "")
                m = re.search(r'filename=\"?([^\"]+)\"?', cd) if cd else None
                name = m.group(1) if m else f"image_{i}.jpg"
                p = os.path.join(download_dir, name)
                with open(p, "wb") as f: f.write(r.content)
                images.append(p)
        except Exception:
            pass
    return images


# ===================== Word processing =====================

def fill_document(template_path: str, images: List[str], signature_path: str, output_path: str) -> None:
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    doc = Document(template_path)
    slot_map = {
        "Thông tin rao bán/sổ đỏ": images[0:1],
        "Mặt trước tài sản": images[1:2],
        "Tổng thể tài sản": images[2:3],
        "Đường phía trước tài sản": images[3:5],
        "Ảnh khác": images[5:7],
    }

    def _table_has_phu_luc(tbl):
        text = "\n".join(cell.text for row in tbl.rows for cell in row.cells)
        return ("Phụ lục" in text) and ("Ảnh TSSS" in text)

    target_table = None
    for tbl in doc.tables:
        if _table_has_phu_luc(tbl):
            target_table = tbl; break

    if target_table:
        for row in target_table.rows:
            for ci, cell in enumerate(row.cells):
                label = cell.text.strip()
                if label in slot_map and ci + 1 < len(row.cells):
                    dest = row.cells[ci + 1]
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
