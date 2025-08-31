"""Helper module for AMIS automation and document manipulation.

Bản tối giản theo yêu cầu:
- KHÔNG tìm kiếm, KHÔNG gọi API.
- Chỉ cần ID trong URL (execution_id). Nếu app cũ truyền record_id, coi như execution_id.
- Truy cập thẳng chi tiết -> mở 'In mẫu thiết lập' -> chọn mẫu -> 'Tải xuống'.
"""

import os
import time
import re
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
    # opts.binary_location = "/usr/bin/chromium"

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
    try:
        os.makedirs(out_dir, exist_ok=True)
        driver.save_screenshot(os.path.join(out_dir, f"debug_{tag}.png"))
        with open(os.path.join(out_dir, f"debug_{tag}.html"), "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    except Exception:
        pass


# ===================== Generic click helpers =====================

def _all_frames(driver):
    frames = [None]
    try:
        frames.extend(driver.find_elements(By.TAG_NAME, "iframe"))
    except Exception:
        pass
    return frames

def _with_each_frame(driver, func):
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

def _click_with_xpaths(driver, xpaths: List[str], timeout: int = 20) -> bool:
    end = time.time() + timeout
    while time.time() < end:
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

        el, _ = _with_each_frame(driver, finder)
        if el:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                driver.execute_script("arguments[0].click();", el)
                return True
            except Exception:
                try:
                    el.click()
                    return True
                except Exception:
                    pass
        time.sleep(0.3)
    return False

def _click_by_texts(driver, texts: List[str], timeout: int = 20, scope_xpath_prefix: str = "") -> bool:
    xps = []
    for t in texts:
        x = t.strip()
        prefix = scope_xpath_prefix or ""
        if prefix and not prefix.strip().startswith("(") and not prefix.strip().startswith("/"):
            # nếu truyền CSS dạng //div[@id='popupexecution']
            prefix = f"{prefix}"
        # nhiều biến thể để bắt đủ kiểu markup
        xps.extend([
            f"{prefix}//button[normalize-space()='{x}']",
            f"{prefix}//span[normalize-space()='{x}']/ancestor::button",
            f"{prefix}//*[self::button or self::a or self::span or self::div][contains(normalize-space(),'{x}')]",
            f"{prefix}//*[@role='button' and (contains(@aria-label,'{x}') or contains(@title,'{x}'))]",
        ])
    return _click_with_xpaths(driver, xps, timeout=timeout)

def _js_click_contains(driver, selector_scope: str, texts: List[str]) -> bool:
    """JS fallback: trong scope CSS, tìm nút có innerText chứa chuỗi và click."""
    js = """
const scopeSel = arguments[0];
const wants = arguments[1];
function visible(el){
  if(!el) return false;
  const st = getComputedStyle(el);
  if(st.display==='none'||st.visibility==='hidden') return false;
  const r=el.getBoundingClientRect(); return r.width>0 && r.height>0;
}
function tryClick(el){
  try{
    el.scrollIntoView({block:'center'}); 
    el.click(); 
    return true;
  }catch(e){}
  return false;
}
const root = scopeSel ? document.querySelector(scopeSel) : document;
if(!root) return false;
const all = root.querySelectorAll('*');
for(const el of all){
  if(!visible(el)) continue;
  const txt = (el.innerText||'').trim();
  if(!txt) continue;
  for(const w of wants){
    if(txt.includes(w)){
      if(tryClick(el)) return true;
      let p=el;
      for(let i=0;i<3;i++){ // thử click ancestor gần nhất
        p=p.parentElement; if(!p) break;
        if(visible(p) && tryClick(p)) return true;
      }
    }
  }
}
return false;
"""
    try:
        return bool(driver.execute_script(js, selector_scope, texts))
    except Exception:
        return False


# ===================== Steps: open preview by "In mẫu thiết lập" =====================

def _open_print_preview_via_popover(driver, download_dir: str) -> None:
    """
    Mở menu 'In mẫu thiết lập' trong popover-action-process rồi mở popup Xem trước (popupexecution).
    Dựa theo selector bạn cung cấp.
    """
    # 0) kéo lên đầu trang cho chắc
    try:
        driver.execute_script("window.scrollTo(0,0);")
    except Exception:
        pass

    # 1) Nếu popup preview đã mở sẵn thì thôi
    try:
        el = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#popupexecution"))
        )
        if el:
            return
    except Exception:
        pass

    # 2) Thử mở popover thao tác (nút ba chấm/thao tác)
    #    Không biết chính xác nút nào, nên thử các khả năng thường gặp.
    opened = False
    # a) nếu popover đã mở thì bỏ qua bước này
    try:
        driver.find_element(
            By.CSS_SELECTOR,
            "body > div.dx-overlay-wrapper.dx-popup-wrapper.dx-popover-wrapper.popover-action-process .dx-popup-content"
        )
        opened = True
    except Exception:
        pass

    if not opened:
        triggers = [
            "//button[contains(@class,'more') or contains(@class,'ms-action') or contains(@class,'icon-more')]",
            "//*[@aria-label='Thao tác' or contains(@aria-label,'Thao tác') or contains(@title,'Thao tác')]",
            "//button[contains(.,'Thao tác') or contains(.,'Tác vụ') or contains(.,'Hành động')]",
            "//div[contains(@class,'more') or contains(@class,'ms-action') or contains(@class,'icon-more')]",
        ]
        _click_with_xpaths(driver, triggers, timeout=6)
        time.sleep(0.3)

    # 3) Chờ popover hiện rồi click mục 'In mẫu thiết lập'
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                "body > div.dx-overlay-wrapper.dx-popup-wrapper.dx-popover-wrapper.popover-action-process .dx-popup-content"
            ))
        )
    except Exception:
        # Fallback: thử JS tìm theo text
        _js_click_contains(driver,
            "body", ["In mẫu thiết lập", "In mẫu thiết lập", "Xem trước mẫu in"]
        )

    # Click chính xác mục trong popover theo CSS bạn gửi
    # body > div.dx-overlay-wrapper.dx-popup-wrapper.dx-popover-wrapper.popover-action-process.dx-popover-without-title.dx-position-bottom > div > div.dx-popup-content > div > div:nth-child(2) > div.m-l-8.p-t-4
    if not _click_with_xpaths(driver, [
        "//div[contains(@class,'popover-action-process')]//div[contains(@class,'dx-popup-content')]//div[contains(@class,'m-l-8') and contains(@class,'p-t-4')]",
    ], timeout=6):
        # Fallback theo text
        if not _click_by_texts(driver, ["In mẫu thiết lập", "Xem trước mẫu in"], timeout=8):
            if not _js_click_contains(driver, None, ["In mẫu thiết lập", "Xem trước mẫu in"]):
                _dump_debug(driver, download_dir, "cannot_open_popover_or_find_preview")
                raise TimeoutException("Không mở được menu 'In mẫu thiết lập'.")

    # 4) Chờ popup preview (#popupexecution)
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#popupexecution"))
        )
    except Exception:
        _dump_debug(driver, download_dir, "no_popupexecution_after_click")
        raise TimeoutException("Không thấy popup Xem trước (popupexecution).")


def _choose_template_and_download(driver, download_dir: str) -> str:
    """
    Trong popup #popupexecution:
      - Chọn 'Phiếu TTTT - Nhà đất' (ưu tiên), nếu không có thì chọn 'Phiếu TTTT - Chung cư/SVP'
      - Bấm 'Tải xuống'
    Trả về tên file .docx (sẽ dò trong thư mục download bằng _wait_for_docx).
    """
    # chờ popup sẵn sàng
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#popupexecution"))
    )

    # Selector bạn cung cấp để chọn item thứ 3 trong list template (thường là 'Nhà đất')
    css_nha_dat = ("#popupexecution > div.ms-popup.flex.flex-col > div.ms-popup--content-header > "
                   "div:nth-child(2) > div > div > div > div.dx-popup-content > div > div > "
                   "div.dx-scrollable-wrapper > div > div.dx-scrollable-content > "
                   "div.dx-scrollview-content > div > div > "
                   "div.list-template.h-full.pos-relative > div.h-full.p-16 > div:nth-child(3) > div")

    clicked_template = False
    # 1) thử click đúng CSS đường dẫn cụ thể
    try:
        el = driver.find_element(By.CSS_SELECTOR, css_nha_dat)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        driver.execute_script("arguments[0].click();", el)
        clicked_template = True
    except Exception:
        pass

    # 2) fallback theo text trong popup
    scope = "//div[@id='popupexecution']"
    if not clicked_template:
        if _click_by_texts(driver, ["Phiếu TTTT - Nhà đất", "TTTT - Nhà đất"], timeout=6, scope_xpath_prefix=scope):
            clicked_template = True

    if not clicked_template:
        # 3) fallback chọn 'Chung cư/SVP' nếu 'Nhà đất' không có
        if not _click_by_texts(driver, ["Phiếu TTTT - Chung cư/SVP", "Chung cư/SVP"], timeout=6, scope_xpath_prefix=scope):
            # 4) Fallback JS tìm theo text
            if not _js_click_contains(driver, "#popupexecution", ["Phiếu TTTT - Nhà đất", "Chung cư/SVP"]):
                _dump_debug(driver, download_dir, "cannot_pick_template")
                raise TimeoutException("Không chọn được mẫu in trong popup.")

    time.sleep(0.4)

    # 5) Click nút Tải xuống trong popup (#popupexecution)
    # CSS bạn cung cấp:
    # #popupexecution ... > div.flex.items-center.justify-between.m-b-8 > div > div.text-blue
    if not _click_with_xpaths(driver, [
        "//div[@id='popupexecution']//div[contains(@class,'text-blue')][contains(.,'Tải xuống') or contains(.,'Tải về') or contains(.,'Download')]",
        "//div[@id='popupexecution']//button[contains(.,'Tải xuống') or contains(.,'Tải về') or contains(.,'Download')]",
    ], timeout=8):
        # Fallback JS theo text
        if not _js_click_contains(driver, "#popupexecution", ["Tải xuống", "Tải về", "Download"]):
            _dump_debug(driver, download_dir, "cannot_click_download_in_popup")
            raise TimeoutException("Không click được 'Tải xuống' trong popup.")

    # File sẽ được tải về thư mục download mặc định
    return _wait_for_docx(os.path.dirname(download_dir) if download_dir.endswith(os.sep) else download_dir, timeout=120)


# ===================== MAIN =====================

def run_automation(
    username: str,
    password: str,
    download_dir: str,
    headless: bool = True,
    record_id: Optional[str] = None,       # Tương thích: coi như execution_id nếu execution_id None
    execution_id: Optional[str] = None,    # Nên dùng: ID trong URL
    status: int = 1,
    company_code: str = "RH7VZQAQ",
) -> Tuple[str, List[str]]:
    """
    Đăng nhập AMIS, truy cập thẳng bằng execution_id rồi:
      In mẫu thiết lập -> chọn mẫu -> Tải xuống.
    Trả về: (đường dẫn .docx, danh sách ảnh best-effort).
    """
    if not execution_id and record_id:
        execution_id = str(record_id)
    if not execution_id:
        raise ValueError("Cần truyền execution_id (hoặc record_id sẽ được coi là execution_id).")

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
        driver.find_element(By.CSS_SELECTOR, "#box-login-right .login-form-inputs .pass-wrap input").send_keys(password)
        driver.find_element(By.CSS_SELECTOR, "#box-login-right .login-form-btn-container button").click()
        wait.until(EC.url_contains("amisapp.misa.vn"))
        time.sleep(0.8)

        # Best-effort: đóng vài popup/onboarding
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
        time.sleep(0.5)

        # 3) Mở "In mẫu thiết lập" (popover) -> hiện popup preview (#popupexecution)
        _open_print_preview_via_popover(driver, download_dir)

        # 4) Chọn mẫu và bấm Tải xuống (trong #popupexecution)
        template_path = _choose_template_and_download(driver, download_dir)

        # 5) Lấy ảnh minh hoạ best-effort
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
