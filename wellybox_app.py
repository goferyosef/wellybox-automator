#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""WellyBox Downloader v4.0"""

import re
import threading
import time
import traceback
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.common.exceptions import (
        NoSuchElementException, StaleElementReferenceException,
        TimeoutException,
    )
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    import keyring
    KEYRING_OK = True
except ImportError:
    KEYRING_OK = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ── Constants ─────────────────────────────────────────────────────────────────
APP_NAME     = "WellyBox Downloader"
KEYRING_SVC  = "WellyBoxApp"
WELLYBOX_URL = "https://app.wellybox.com/"
WAIT_S, WAIT_L = 10, 30

HEBREW_MONTHS = {
    "ינואר":1,"ינו":1,"פברואר":2,"פבר":2,"מרץ":3,"מרס":3,
    "אפריל":4,"אפר":4,"מאי":5,"יוני":6,"יונ":6,
    "יולי":7,"יול":7,"אוגוסט":8,"אוג":8,
    "ספטמבר":9,"ספט":9,"אוקטובר":10,"אוק":10,
    "נובמבר":11,"נוב":11,"דצמבר":12,"דצמ":12,
}
ENGLISH_MONTHS = {
    "january":1,"jan":1,"february":2,"feb":2,"march":3,"mar":3,
    "april":4,"apr":4,"may":5,"june":6,"jun":6,
    "july":7,"jul":7,"august":8,"aug":8,"september":9,"sep":9,
    "october":10,"oct":10,"november":11,"nov":11,"december":12,"dec":12,
}
INVOICE_TYPES = ["חשבונית מס / קבלה", "חשבונית מס/קבלה", "חשבונית"]
RECEIPT_TYPES = ["קבלה"]

# All label strings that appear in the detail panel — never treat these as values
KNOWN_LABELS  = {
    "סוג מסמך", "ספק", "מטבע", "סכום", "תאריך", "מספר", "סטטוס",
    "מקור", "קטגוריה", "הערות", "עסק",
    'תאריך ע"ג המסמך',   # ASCII double-quote variant
    'תאריך ע״ג המסמך',   # Hebrew gershayim variant (U+05F4)
    "תאריך העלאת המסמך", "תאריך הנפקה", "תאריך קבלה",
    "סך הכל", 'סה"כ', 'סה״כ', "מספר חשבונית", "שם העסק",
    "מטבע חשבונית", "מספר מסמך", "פרטים", "קובץ",
    "סטטוס תשלום", "אמצעי תשלום", "סכום לתשלום", "יתרה",
    "תשלום", "הערה", "כתובת", "ח.פ", "ע.מ", "מיקוד",
    "חבר צוות", "מקור", "שולח האימייל",
}

COLOR_GREEN  = "C8E6C9"
COLOR_RED    = "FFCDD2"
COLOR_BLUE   = "BBDEFB"
COLOR_YELLOW = "FFF9C4"
COLOR_GREY   = "EEEEEE"
COLOR_ORANGE = "FFE0B2"


# ── Data ──────────────────────────────────────────────────────────────────────
@dataclass
class CardResult:
    idx:      int
    vendor:   str = ""
    doc_date: str = ""
    doc_type: str = ""
    status:   str = ""   # downloaded_invoice|dup_invoice|downloaded_receipt|dup_receipt|skipped|error
    filename: str = ""
    note:     str = ""


# ── Credentials ───────────────────────────────────────────────────────────────
def load_creds():
    if not KEYRING_OK:
        return None, None
    return (keyring.get_password(KEYRING_SVC, "username"),
            keyring.get_password(KEYRING_SVC, "password"))

def save_creds(u, p):
    if KEYRING_OK:
        keyring.set_password(KEYRING_SVC, "username", u)
        keyring.set_password(KEYRING_SVC, "password", p)


# ── Credentials dialog ────────────────────────────────────────────────────────
class CredsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("פרטי כניסה — WellyBox")
        self.resizable(False, False)
        self.grab_set()
        self.username = ""
        self.password = ""
        saved_u, saved_p = load_creds()
        pad = {"padx": 10, "pady": 5}
        tk.Label(self, text="אימייל:").grid(row=0, column=0, sticky="e", **pad)
        self._u = tk.Entry(self, width=34)
        self._u.grid(row=0, column=1, **pad)
        if saved_u:
            self._u.insert(0, saved_u)
        tk.Label(self, text="סיסמה:").grid(row=1, column=0, sticky="e", **pad)
        self._p = tk.Entry(self, width=34, show="*")
        self._p.grid(row=1, column=1, **pad)
        if saved_p:
            self._p.insert(0, saved_p)
        f = tk.Frame(self)
        f.grid(row=2, column=0, columnspan=2, pady=8)
        tk.Button(f, text="שמור",  width=10, command=self._ok).pack(side="left", padx=6)
        tk.Button(f, text="ביטול", width=10, command=self.destroy).pack(side="left", padx=6)
        self._u.focus_set()
        self.bind("<Return>", lambda _: self._ok())
        self.wait_window()

    def _ok(self):
        self.username = self._u.get().strip()
        self.password = self._p.get()
        self.destroy()


# ── Date helpers ──────────────────────────────────────────────────────────────
def parse_date(text: str) -> Optional[datetime]:
    text = text.strip()
    # dd/mm/yyyy or dd.mm.yyyy  (4-digit year — try first to avoid partial match)
    m = re.search(r'(\d{1,2})[/.](\d{1,2})[/.](\d{4})', text)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except Exception:
            pass
    # dd/mm/yy or dd.mm.yy  (2-digit year)
    m = re.search(r'(\d{1,2})[/.](\d{1,2})[/.](\d{2})(?!\d)', text)
    if m:
        try:
            y = int(m.group(3))
            y = 2000 + y if y < 70 else 1900 + y
            return datetime(y, int(m.group(2)), int(m.group(1)))
        except Exception:
            pass
    # Hebrew: 24 מרץ 2026
    m = re.search(r'(\d{1,2})\s+([א-ת]+)\s+(\d{4})', text)
    if m:
        d, mon, y = int(m.group(1)), m.group(2), int(m.group(3))
        mn = HEBREW_MONTHS.get(mon) or HEBREW_MONTHS.get(mon[:3])
        if mn:
            try:
                return datetime(y, mn, d)
            except Exception:
                pass
    # English: Mar 24th 2026 / Mar 24 2026
    m = re.search(r'([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(\d{4})', text)
    if m:
        mon, d, y = m.group(1).lower(), int(m.group(2)), int(m.group(3))
        mn = ENGLISH_MONTHS.get(mon) or ENGLISH_MONTHS.get(mon[:3])
        if mn:
            try:
                return datetime(y, mn, d)
            except Exception:
                pass
    # English: 24 Mar 2026
    m = re.search(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', text)
    if m:
        d, mon, y = int(m.group(1)), m.group(2).lower(), int(m.group(3))
        mn = ENGLISH_MONTHS.get(mon) or ENGLISH_MONTHS.get(mon[:3])
        if mn:
            try:
                return datetime(y, mn, d)
            except Exception:
                pass
    return None

def fmt_date_il(dt: datetime) -> str:
    return f"{dt.day}.{dt.month}.{dt.year}"

def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', '', s).strip()

def type_matches(doc_type: str, expected_types: list) -> bool:
    """Exact match ignoring spaces and slash variants — prevents 'קבלה' matching 'חשבונית מס / קבלה'."""
    dt = re.sub(r'[\s/]+', '', doc_type).strip()
    for et in expected_types:
        if dt == re.sub(r'[\s/]+', '', et).strip():
            return True
    return False


# ── Bot ───────────────────────────────────────────────────────────────────────
class Bot:
    def __init__(self, days_back: int, output_folder: Path, log_cb, done_cb):
        self.days_back     = days_back
        self.output_folder = output_folder
        self._log_cb       = log_cb
        self._done_cb      = done_cb
        self.driver        = None
        self.stop_event    = threading.Event()
        self.results       = []
        self._lines        = []
        self._cutoff       = (datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                              - timedelta(days=days_back - 1))

    def start(self):
        threading.Thread(target=self._run, daemon=True).start()

    def stop(self):
        self.stop_event.set()

    def _emit(self, msg: str, level: str = "INFO"):
        ts   = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] [{level}] {msg}"
        self._lines.append(line)
        self._log_cb(line, level)

    # ── main flow ─────────────────────────────────────────────────────────────
    def _run(self):
        try:
            u, p = load_creds()
            if not u or not p:
                self._done_cb(need_creds=True)
                return
            self._emit(f"מתחיל — {self.days_back} ימים (מ-{self._cutoff.strftime('%d/%m/%Y')})")
            self._start_browser()
            self._login(u, p)
            if self.stop_event.is_set():
                return
            # Stage 1 — invoices
            self._emit("══════ שלב 1: חשבוניות ══════")
            self._go_to_all_receipts()
            self._apply_filter(check=["חשבונית", "חשבונית מס /קבלה"],
                               uncheck=["קבלה"])
            self._process_stage(self.output_folder / "חשבוניות לקליטה",
                                INVOICE_TYPES, "חשבונית")
            if self.stop_event.is_set():
                return
            # Stage 2 — receipts
            self._emit("══════ שלב 2: קבלות ══════")
            self._go_to_all_receipts()
            self._apply_filter(check=["קבלה"],
                               uncheck=["חשבונית", "חשבונית מס /קבלה"])
            self._process_stage(self.output_folder / "קבלות",
                                RECEIPT_TYPES, "קבלה")
            self._logout()
            self._save_reports()
            self._emit("✓ סיום בהצלחה")
        except Exception as exc:
            self._emit(f"שגיאה כללית: {exc}", "ERROR")
            self._emit(traceback.format_exc(), "ERROR")
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
            self._done_cb(need_creds=False)

    # ── browser ───────────────────────────────────────────────────────────────
    def _start_browser(self):
        self._emit("פותח Chrome…")
        dl_tmp = self.output_folder / "_dl_tmp"
        # Clear any leftover files from previous runs so detection works cleanly
        if dl_tmp.exists():
            for f in dl_tmp.iterdir():
                try:
                    f.unlink()
                except Exception:
                    pass
        dl_tmp.mkdir(parents=True, exist_ok=True)
        self._dl_tmp = dl_tmp

        opts = ChromeOptions()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        opts.add_experimental_option("prefs", {
            "download.default_directory":  str(dl_tmp),
            "download.prompt_for_download": False,
            "download.directory_upgrade":   True,
            "safebrowsing.enabled":         True,
        })
        service = ChromeService(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=opts)
        self.driver.execute_script(
            "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
        )
        # Set download directory via CDP (works for newer Chrome versions)
        dl_tmp_str = str(dl_tmp).replace("\\", "/")
        try:
            self.driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": str(dl_tmp),
                "eventsEnabled": True,
            })
        except Exception:
            try:
                self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                    "behavior": "allow",
                    "downloadPath": str(dl_tmp),
                })
            except Exception:
                pass
        self._emit(f"  תיקיית הורדות: {dl_tmp}")

    def _w(self, t=WAIT_L):
        return WebDriverWait(self.driver, t)

    def _shot(self, name: str):
        try:
            d = self.output_folder / "logs" / "diag"
            d.mkdir(parents=True, exist_ok=True)
            self.driver.save_screenshot(
                str(d / f"{name}_{datetime.now().strftime('%H%M%S')}.png")
            )
        except Exception:
            pass

    # ── login ─────────────────────────────────────────────────────────────────
    def _login(self, username: str, password: str):
        self._emit("מתחבר…")
        self.driver.get(WELLYBOX_URL)
        time.sleep(3)

        # Click כניסה on landing page if present
        try:
            btn = self._w(WAIT_S).until(EC.element_to_be_clickable((
                By.XPATH,
                "//*[self::a or self::button]"
                "[contains(.,'כניסה') or contains(.,'התחברות') or contains(.,'Login')]"
            )))
            btn.click()
            time.sleep(2)
        except TimeoutException:
            pass

        # Email field
        email_el = None
        for xp in ["//input[@type='email']", "//input[@name='email']", "//input[@id='email']"]:
            try:
                email_el = self._w(8).until(EC.element_to_be_clickable((By.XPATH, xp)))
                break
            except TimeoutException:
                continue
        if not email_el:
            raise RuntimeError("שדה אימייל לא נמצא")
        email_el.click()
        time.sleep(0.2)
        email_el.send_keys(Keys.CONTROL + "a")
        email_el.send_keys(username)
        self._emit(f"  אימייל: {username}")

        # Password field (may require submitting email first in 2-step flows)
        pass_el = None
        for attempt in range(2):
            try:
                pass_el = self._w(6).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
                )
                break
            except TimeoutException:
                if attempt == 0:
                    try:
                        self.driver.find_element(By.XPATH, "//button[@type='submit']").click()
                        time.sleep(2)
                    except Exception:
                        pass
        if not pass_el:
            raise RuntimeError("שדה סיסמה לא נמצא")
        pass_el.click()
        time.sleep(0.2)
        pass_el.send_keys(Keys.CONTROL + "a")
        pass_el.send_keys(password)

        # Submit
        submitted = False
        for xp in [
            "//button[@type='submit']",
            "//button[contains(.,'התחברות') or contains(.,'כניסה') or contains(.,'Login')]",
        ]:
            try:
                for b in self.driver.find_elements(By.XPATH, xp):
                    if b.is_displayed() and b.is_enabled():
                        self.driver.execute_script("arguments[0].click();", b)
                        submitted = True
                        break
                if submitted:
                    break
            except Exception:
                pass
        if not submitted:
            pass_el.send_keys(Keys.RETURN)

        # Wait for dashboard
        self._emit("ממתין לכניסה…")
        try:
            self._w(WAIT_L).until(EC.any_of(
                EC.url_contains("dashboard"),
                EC.url_contains("receipts"),
                EC.url_contains("ng2ux"),
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(.,'כל החשבוניות')]")
                ),
            ))
            self._emit("✓ מחובר")
        except TimeoutException:
            self._shot("login_fail")
            raise RuntimeError("כניסה נכשלה — בדוק פרטי כניסה")
        time.sleep(2)

    # ── navigate ──────────────────────────────────────────────────────────────
    def _go_to_all_receipts(self):
        self._emit("עובר ל'כל החשבוניות'…")
        try:
            el = self._w(WAIT_L).until(EC.element_to_be_clickable((
                By.XPATH, "//*[normalize-space(text())='כל החשבוניות']"
            )))
            self.driver.execute_script("arguments[0].click();", el)
            self._emit("  ✓ כל החשבוניות")
        except TimeoutException:
            self._emit("  'כל החשבוניות' לא נמצא", "WARN")
        # wait for skeletons to disappear
        try:
            self._w(60).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiSkeleton-root"))
            )
        except TimeoutException:
            pass
        time.sleep(1.5)
        # Clear any active filters before proceeding
        self._clear_filters()

    def _clear_filters(self):
        try:
            el = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
                By.XPATH, "//*[normalize-space(text())='נקה מסננים']"
            )))
            self.driver.execute_script("arguments[0].click();", el)
            self._emit("  ✓ נקה מסננים")
            time.sleep(1)
        except TimeoutException:
            pass  # no active filters — nothing to clear

    # ── filter ────────────────────────────────────────────────────────────────
    def _apply_filter(self, check: list, uncheck: list):
        self._emit(f"פילטר: {', '.join(check)}")

        # Open "מסננים נוספים"
        if not self._js_click_text("מסננים נוספים", wait=WAIT_L):
            raise RuntimeError("'מסננים נוספים' לא נמצא")
        time.sleep(1)

        # Open "סוג מסמך"
        if not self._js_click_text("סוג מסמך", wait=WAIT_S):
            raise RuntimeError("'סוג מסמך' לא נמצא")
        time.sleep(1)

        # Set checkboxes
        for label in check:
            self._set_checkbox(label, True)
        for label in uncheck:
            self._set_checkbox(label, False)

        time.sleep(0.8)
        self._shot("after_checkboxes")   # verify both boxes are ticked before clicking החל

        # Click "החל"
        if not self._click_any("החל"):
            self._shot("error_no_apply")
            # Log all visible buttons for diagnosis
            btns = self.driver.find_elements(By.XPATH, "//button")
            visible = [b.text.strip() for b in btns if b.is_displayed() and b.text.strip()]
            self._emit(f"  כפתורים נראים: {visible}", "WARN")
            raise RuntimeError("כפתור 'החל' לא נמצא")

        self._emit("  ✓ פילטר הוחל")
        # Wait for filter panel to fully close, then for results to load
        time.sleep(2)
        # Wait until the filter panel / dropdown is gone
        try:
            self._w(10).until_not(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(normalize-space(text()),'סוג מסמך') and @role!='menuitem']"
                               "[ancestor::*[contains(@class,'filter') or contains(@class,'Filter') "
                               "or contains(@class,'drawer') or contains(@class,'Drawer') "
                               "or contains(@class,'popover') or contains(@class,'Popover')]]")
                )
            )
        except TimeoutException:
            pass
        time.sleep(1.5)
        try:
            self._w(30).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiSkeleton-root"))
            )
        except TimeoutException:
            pass
        time.sleep(1.5)

    def _js_click_text(self, text: str, wait: int = WAIT_S) -> bool:
        try:
            el = WebDriverWait(self.driver, wait).until(
                EC.element_to_be_clickable((By.XPATH, f"//*[normalize-space(text())='{text}']"))
            )
            self.driver.execute_script("arguments[0].click();", el)
            self._emit(f"  לחץ: '{text}'")
            return True
        except TimeoutException:
            return False

    def _click_any(self, text: str) -> bool:
        for xp in [
            f"//*[normalize-space(text())='{text}']",
            f"//button[contains(.,'{text}')]",
            f"//*[contains(.,'{text}')][self::button or self::a]",
            f"//*[contains(.,'{text}') and (contains(@class,'btn') or contains(@class,'Btn'))]",
        ]:
            try:
                for el in self.driver.find_elements(By.XPATH, xp):
                    if el.is_displayed() and el.is_enabled():
                        self.driver.execute_script("arguments[0].click();", el)
                        self._emit(f"  לחץ: '{text}'")
                        return True
            except Exception:
                pass
        return False

    def _set_checkbox(self, label: str, want_checked: bool):
        """Find a filter checkbox by exact label text and set its checked state."""
        # Build variants to try (normalise spaces around slash)
        variants = [label]
        if "/" in label:
            variants.append(re.sub(r'\s*/\s*', ' / ', label))
            variants.append(re.sub(r'\s*/\s*', '/', label))

        for lv in variants:
            lv_js = lv.replace("'", "\\'")
            result = self.driver.execute_script(f"""
                var target = '{lv_js}';
                var want   = {'true' if want_checked else 'false'};
                // Walk all visible elements looking for exact text match
                var all = document.querySelectorAll('label, li, span, div');
                for (var i = 0; i < all.length; i++) {{
                    var el = all[i];
                    if (!el.offsetParent) continue;          // skip hidden
                    var txt = el.textContent.trim();
                    if (txt !== target) continue;            // EXACT match only
                    // Find the associated checkbox
                    var cb = el.querySelector('input[type="checkbox"]');
                    if (!cb) {{
                        var p = el.parentElement;
                        for (var k = 0; k < 4 && p; k++) {{
                            cb = p.querySelector('input[type="checkbox"]');
                            if (cb) break;
                            p = p.parentElement;
                        }}
                    }}
                    if (cb) {{
                        if (cb.checked === want) return 'already';
                        cb.click();
                        return 'clicked_cb';
                    }}
                    // No input found — click the label element itself
                    el.click();
                    return 'clicked_el';
                }}
                return null;
            """)

            if result:
                state = "✓" if want_checked else "✗"
                note  = " (כבר במצב הנכון)" if result == "already" else ""
                self._emit(f"  {state} '{lv}'{note}")
                time.sleep(0.4)
                return

        if want_checked:
            self._emit(f"  '{label}' לא נמצא בפילטר", "WARN")

    # ── card processing ───────────────────────────────────────────────────────
    def _scroll_card_list(self) -> bool:
        """Scroll the card list container down to trigger batch loading.
        Returns True if the scroll position actually changed (more content may load)."""
        return self.driver.execute_script("""
            var el = document.querySelector('.wbgvbct_receipts_list');
            if (!el) return false;
            var before = el.scrollTop;
            el.scrollBy(0, el.clientHeight * 0.85);
            return el.scrollTop > before;
        """) or False

    def _card_key(self, card) -> str:
        """Stable dedup key: text content + DOM position to survive identical-text duplicates."""
        try:
            txt = card.text.strip()[:80]
            loc = card.location  # {'x': ..., 'y': ...}
            return f"{txt}|{loc.get('x',0)},{loc.get('y',0)}"
        except Exception:
            return ""

    def _process_stage(self, dest_folder: Path, expected_types: list, stage_name: str):
        dest_folder.mkdir(parents=True, exist_ok=True)

        seen_keys: set = set()   # keys of cards already handled (or skipped)
        global_idx = 0           # monotonic card counter
        out_of_range_streak = 0  # consecutive cards older than cutoff → stop
        no_new_cards_streak = 0  # consecutive scrolls with nothing new → stop
        first_fetch = True

        while True:
            if self.stop_event.is_set():
                break

            # Re-fetch card references every iteration — DOM may have changed
            cards = self._find_cards()
            if first_fetch:
                self._emit(f"נמצאו {len(cards)} כרטיסים בטעינה ראשונה")
                first_fetch = False

            # Find the FIRST card we haven't handled yet
            card = None
            card_key = None
            for c in cards:
                k = self._card_key(c)
                if k not in seen_keys:
                    card = c
                    card_key = k
                    break

            if card is None:
                # All visible cards handled — try scrolling for more
                scrolled = self._scroll_card_list()
                if not scrolled:
                    break
                no_new_cards_streak += 1
                if no_new_cards_streak >= 3:
                    break
                time.sleep(1.5)
                continue

            no_new_cards_streak = 0
            seen_keys.add(card_key)   # mark before processing to prevent re-entry
            global_idx += 1
            card_num = global_idx

            try:
                text = card.text.strip()

                # Log status for diagnosis but do NOT skip based on it.
                # WellyBox marks every AI-processed document "נשמר" on arrival —
                # that is unrelated to whether our app has downloaded it yet.
                _status = "נשמר" if "נשמר" in text else ("חדש" if "חדש" in text else "?")
                _lines  = [l.strip() for l in text.split('\n')
                           if l.strip() and l.strip() not in ('נשמר', 'חדש')]
                _hint   = _lines[0][:40] if _lines else ''
                self._emit(f"  #{card_num}: {_status} ({_hint})")

                card_date = self._date_from_text(text)
                if card_date is None:
                    self._emit(f"  #{card_num}: תאריך לא נקרא — דלג", "WARN")
                    continue

                if card_date < self._cutoff:
                    self._emit(f"  #{card_num}: {card_date.strftime('%d/%m/%Y')} — מחוץ לטווח")
                    out_of_range_streak += 1
                    if out_of_range_streak >= 5:
                        self._emit("5 כרטיסים רצופים מחוץ לטווח — עוצר גלילה")
                        return
                    continue

                out_of_range_streak = 0

                self._emit(f"פותח #{card_num} ({card_date.strftime('%d/%m/%Y')})…")
                result = CardResult(idx=card_num, doc_date=fmt_date_il(card_date))
                self._open_process(card, result, dest_folder, expected_types, stage_name)
                self.results.append(result)

                self._close_card()
                time.sleep(0.5)

            except StaleElementReferenceException:
                # Card element went stale — it will be re-fetched next iteration
                self._emit(f"  #{card_num}: stale element — מרענן", "WARN")
                seen_keys.discard(card_key)   # allow retry with fresh reference
            except Exception as exc:
                self._emit(f"  #{card_num}: שגיאה — {exc}", "ERROR")
                r = CardResult(idx=card_num, status="error", note=str(exc))
                self.results.append(r)
                self._close_card()
                time.sleep(0.5)

        self._emit(f"שלב '{stage_name}' הסתיים")

    def _find_cards(self) -> list:
        return self.driver.execute_script("""
            var STATUS = ['חדש','נשמר'];
            var container = document.querySelector('.wbgvbct_receipts_list') || document.body;
            var seen = new Set(), cards = [];
            var all = container.querySelectorAll('*');
            for (var i = 0; i < all.length; i++) {
                var node = all[i];
                // Strip invisible/zero-width chars before comparing
                var t = (node.innerText || node.textContent || '').trim()
                         .replace(/[\u200b\u200c\u200d\uFEFF\u00a0]/g, '');
                if (!STATUS.includes(t)) continue;
                // Walk up to find a card-sized ancestor
                var anc = node.parentElement;
                for (var j = 0; j < 15; j++) {
                    if (!anc || anc === document.body) break;
                    var w = anc.offsetWidth, h = anc.offsetHeight;
                    if (w > 120 && w < window.innerWidth * 0.7 && h > 80 && h < 750) {
                        var key = Math.round(anc.getBoundingClientRect().top) + '|' + anc.offsetLeft;
                        if (!seen.has(key)) { seen.add(key); cards.push(anc); }
                        break;
                    }
                    anc = anc.parentElement;
                }
            }
            return cards;
        """) or []

    def _date_from_text(self, text: str) -> Optional[datetime]:
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        for line in reversed(lines[-5:]):
            dt = parse_date(line)
            if dt:
                return dt
        for line in lines:
            dt = parse_date(line)
            if dt:
                return dt
        return None

    def _hover_tooltip(self, card) -> tuple:
        """Hover over a card and read the dark popup that shows ספק and date.
        Returns (vendor, date_str) — either may be empty string on failure."""
        from selenium.webdriver.common.action_chains import ActionChains
        vendor = ""
        date_str = ""
        try:
            # Scroll card into view, then hover
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", card)
            time.sleep(0.3)
            ActionChains(self.driver).move_to_element(card).perform()
            time.sleep(1.0)   # wait for dark overlay to appear

            # The popup is a dark overlay inside the card or a MUI portal near it.
            # Try to find any newly-visible element containing "ספק:"
            tooltip_text = self.driver.execute_script("""
                var card = arguments[0];
                function findWithLabel(root) {
                    var els = root.querySelectorAll('*');
                    for (var i = 0; i < els.length; i++) {
                        var el = els[i];
                        if (!el.offsetParent) continue;
                        var txt = (el.innerText || el.textContent || '').trim();
                        if (txt.includes('\u05e1\u05e4\u05e7') && txt.includes('\u05ea\u05d0\u05e8\u05d9\u05da'))
                            return txt;
                    }
                    return null;
                }
                // 1. inside the card itself
                var t = findWithLabel(card);
                if (t) return t;
                // 2. MUI portals / tooltips appended to body
                var portals = document.querySelectorAll(
                    '[class*="tooltip" i],[class*="Tooltip"],[class*="overlay" i],' +
                    '[class*="Overlay"],[class*="popover" i],[class*="Popover"],' +
                    '[role="tooltip"],[data-popper-placement]');
                for (var i = 0; i < portals.length; i++) {
                    var el = portals[i];
                    if (!el.offsetParent) continue;
                    var txt = (el.innerText || el.textContent || '').trim();
                    if (txt.includes('\u05e1\u05e4\u05e7')) return txt;
                }
                return null;
            """, card)

            if tooltip_text:
                self._emit(f"  tooltip raw: {repr(tooltip_text[:120])}")
                # Normalize: insert newline before each field label so the
                # parser works regardless of whether the tooltip rendered with
                # newlines (innerText) or without them (textContent fallback).
                _FIELD = (r'(?:ספק|סך\s*הכל|סה[״"]\s*כ|תאריך[^:\n]{0,30}?'
                          r'|עסק|קטגוריה|מספר[^:\n]{0,20}?'
                          r'|סטטוס|מקור|מטבע|הערה|שולח)\s*:')
                tooltip_norm = re.sub(
                    r'(?<!\n)(?=' + _FIELD + r')', '\n', tooltip_text
                ).strip()
                # Extract vendor — stops at newline (i.e. at the next field)
                m = re.search(r'ספק\s*:?\s*([^\n]+)', tooltip_norm)
                if m:
                    vendor = m.group(1).strip()
                # Extract document date — prefer תאריך ע״ג המסמך / תאריך ע"ג המסמך
                m = re.search(r'תאריך\s+ע[״"]\s*ג\s+המסמך\s*:?\s*([^\n]+)', tooltip_norm)
                if not m:
                    m = re.search(r'תאריך[^\n:]*:\s*([^\n]+)', tooltip_norm)
                if m:
                    date_str = m.group(1).strip()
            else:
                self._emit("  tooltip לא נמצא", "WARN")
        except Exception as e:
            self._emit(f"  שגיאת hover: {e}", "WARN")
        return vendor, date_str

    def _open_process(self, card, result: CardResult, dest_folder: Path,
                      expected_types: list, stage_name: str):
        # ── Step 1: hover over card to read the popup tooltip ─────────────────
        vendor, date_str = self._hover_tooltip(card)
        self._emit(f"  tooltip → ספק: '{vendor}' | תאריך: '{date_str}'")

        # ── Fast pre-check: if file already exists, skip without opening panel ─
        if vendor and (date_str or result.doc_date):
            _dt = parse_date(date_str) if date_str else None
            _fd = fmt_date_il(_dt) if _dt else result.doc_date
            if _fd:
                _bn = f"{safe_name(vendor)} {safe_name(_fd)}"
                if list(dest_folder.glob(f"{_bn}.*")):
                    result.status   = "dup_invoice" if stage_name != "קבלה" else "dup_receipt"
                    result.note     = "קובץ זהה כבר קיים"
                    result.filename = _bn
                    result.vendor   = vendor
                    result.doc_date = _fd
                    self._emit(f"  ← קיים כבר (ללא פתיחת פאנל) — {_bn}")
                    return

        # ── Step 2: click card to open detail panel (for doc_type only) ───────
        try:
            self.driver.execute_script("""
                var r=arguments[0].getBoundingClientRect();
                var el=document.elementFromPoint(r.left+r.width/2, r.top+r.height/2);
                if(el) el.click(); else arguments[0].click();
            """, card)
        except Exception:
            self.driver.execute_script("arguments[0].click();", card)
        time.sleep(2)

        # Get panel
        panel = self._get_panel()
        if not panel:
            self._shot("error_no_panel")
            result.status = "error"
            result.note   = "פאנל לא נפתח"
            return

        lines = [l.strip() for l in panel.text.split('\n') if l.strip()]
        doc_type = self._field_val("סוג מסמך", lines)

        # Fallback: if hover didn't get vendor/date, try panel text
        if not vendor:
            vendor = self._js_field_val(panel, "ספק") or self._field_val("ספק", lines)
        if not date_str:
            date_str = (self._js_field_val(panel, 'תאריך ע״ג המסמך') or
                        self._js_field_val(panel, 'תאריך ע"ג המסמך') or
                        self._field_val('תאריך ע״ג המסמך', lines) or
                        self._field_val('תאריך ע"ג המסמך', lines))

        result.doc_type = doc_type
        result.vendor   = vendor

        if date_str:
            dt = parse_date(date_str)
            if dt:
                result.doc_date = fmt_date_il(dt)

        self._emit(f"  סוג: '{doc_type}' | ספק: '{vendor}' | תאריך: '{date_str}'")

        # Verify doc type — exact match (prevents 'קבלה' matching inside 'חשבונית מס / קבלה')
        if not type_matches(doc_type, expected_types):
            result.status = "skipped"
            result.note   = f"סוג '{doc_type}' — לא {stage_name}"
            self._emit(f"  ← דלג: {result.note}", "WARN")
            return

        # Build filename — safe_name on full string ensures no illegal chars
        fn_vendor = safe_name(vendor) if vendor else "ללא_ספק"
        fn_date   = safe_name(result.doc_date or fmt_date_il(datetime.now()))
        base_name = f"{fn_vendor} {fn_date}"

        # Duplicate check
        if list(dest_folder.glob(f"{base_name}.*")):
            result.status = "dup_invoice" if stage_name != "קבלה" else "dup_receipt"
            result.note   = "קובץ זהה כבר קיים"
            result.filename = base_name
            self._emit(f"  ← דלג: קיים כבר — {base_name}", "WARN")
            return

        # Download
        self._emit(f"  מוריד: {base_name}…")
        downloaded = self._download(dest_folder, base_name)
        if downloaded:
            result.status   = "downloaded_invoice" if stage_name != "קבלה" else "downloaded_receipt"
            result.filename = downloaded.name
            self._emit(f"  ✓ {downloaded.name}")
        else:
            result.status = "error"
            result.note   = "הורדה נכשלה"
            self._emit("  ✗ הורדה נכשלה", "ERROR")

    def _get_panel(self):
        selectors = [
            "[class*='drawer']","[class*='Drawer']",
            "[class*='panel']","[class*='Panel']",
            "[class*='detail']","[class*='Detail']",
            "[class*='sidebar']","[class*='Sidebar']",
            "[role='dialog']","[role='complementary']",
        ]
        for sel in selectors:
            try:
                for el in self.driver.find_elements(By.CSS_SELECTOR, sel):
                    txt = el.text.strip()
                    if el.is_displayed() and len(txt) > 30 and "סוג מסמך" in txt:
                        return el
            except Exception:
                pass
        # fallback: find element containing "סוג מסמך" and walk up
        try:
            label = self.driver.find_element(
                By.XPATH, "//*[contains(text(),'סוג מסמך')]"
            )
            el = label
            for _ in range(8):
                el = el.find_element(By.XPATH, "..")
                w, h = el.size.get('width', 0), el.size.get('height', 0)
                if w > 200 and h > 300:
                    return el
        except Exception:
            pass
        return None

    def _js_field_val(self, panel_el, label: str) -> str:
        """Find a panel field value by label text using JS DOM traversal,
        scoped to the panel element to avoid matching the column-picker sidebar.
        Also checks input[value] for React-controlled fields."""
        variants = [label]
        if '\u05f4' in label:                          # gershayim → also try ASCII "
            variants.append(label.replace('\u05f4', '"'))
        elif '"' in label:                             # ASCII " → also try gershayim
            variants.append(label.replace('"', '\u05f4'))

        for lv in variants:
            result = self.driver.execute_script("""
                var container = arguments[0];
                var lbl = arguments[1];
                var KNOWN = new Set([
                    '\u05e1\u05d5\u05d2 \u05de\u05e1\u05de\u05da','\u05e1\u05e4\u05e7',
                    '\u05de\u05d8\u05d1\u05e2','\u05e1\u05db\u05d5\u05dd','\u05ea\u05d0\u05e8\u05d9\u05da',
                    '\u05de\u05e1\u05e4\u05e8','\u05e1\u05d8\u05d0\u05d8\u05d5\u05e1',
                    '\u05de\u05e7\u05d5\u05e8','\u05e7\u05d8\u05d2\u05d5\u05e8\u05d9\u05d4',
                    '\u05e2\u05e1\u05e7','\u05e1\u05da \u05d4\u05db\u05dc',
                    '\u05e1\u05d8\u05d0\u05d8\u05d5\u05e1 \u05ea\u05e9\u05dc\u05d5\u05dd',
                    '\u05de\u05e1\u05e4\u05e8 \u05de\u05e1\u05de\u05da',
                    '\u05d7\u05d1\u05e8 \u05e6\u05d5\u05d5\u05ea','\u05d4\u05e2\u05e8\u05d4',
                    '\u05ea\u05d0\u05e8\u05d9\u05da \u05e2\u05dc\u05d0\u05ea \u05d4\u05de\u05e1\u05de\u05da'
                ]);
                var all = container.querySelectorAll(
                    'p,span,div,td,dt,label,h1,h2,h3,h4,h5,h6');
                for (var i = 0; i < all.length; i++) {
                    var el = all[i];
                    if (!el.offsetParent) continue;
                    if (el.children.length > 2) continue;
                    var txt = (el.innerText || el.textContent || '').trim();
                    if (txt !== lbl) continue;

                    // Strategy 1: sibling in same parent
                    var par = el.parentElement;
                    if (par) {
                        var ch = Array.prototype.slice.call(par.children);
                        for (var j = 0; j < ch.length; j++) {
                            if (ch[j] === el) continue;
                            var v = (ch[j].innerText || ch[j].textContent || '').trim();
                            if (v && v !== lbl && !KNOWN.has(v)) return v;
                            // check input[value] inside sibling
                            var inp = ch[j].querySelector('input,textarea');
                            if (inp && inp.value) return inp.value;
                        }
                        // Strategy 2: parent's next sibling (2-col grid row)
                        var gp = par.parentElement;
                        if (gp) {
                            var gpch = Array.prototype.slice.call(gp.children);
                            var pi = gpch.indexOf(par);
                            if (pi + 1 < gpch.length) {
                                var v = (gpch[pi+1].innerText || gpch[pi+1].textContent || '').trim();
                                if (v && v !== lbl && !KNOWN.has(v)) return v;
                                var inp = gpch[pi+1].querySelector('input,textarea');
                                if (inp && inp.value) return inp.value;
                            }
                        }
                    }
                    // Strategy 3: next sibling element
                    var ns = el.nextElementSibling;
                    if (ns) {
                        var v = (ns.innerText || ns.textContent || '').trim();
                        if (v && v !== lbl && !KNOWN.has(v)) return v;
                        var inp = ns.querySelector('input,textarea');
                        if (inp && inp.value) return inp.value;
                    }
                }
                return null;
            """, panel_el, lv)
            if result:
                return result.strip()
        return ""

    def _field_val(self, label: str, lines: list) -> str:
        idx = -1
        try:
            idx = lines.index(label)
        except ValueError:
            for i, l in enumerate(lines):
                if l.startswith(label):
                    idx = i
                    break
        if idx == -1:
            return ""
        for val in lines[idx + 1:]:
            if val in KNOWN_LABELS:
                continue
            if val:
                return val
        return ""

    def _close_card(self):
        try:
            self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(0.5)
        except Exception:
            pass
        for xp in [
            "//button[@aria-label='close' or @aria-label='סגור']",
            "//*[@data-testid='CloseIcon']/..",
            "//button[contains(@class,'close') or contains(@class,'Close')]",
        ]:
            try:
                self.driver.find_element(By.XPATH, xp).click()
                time.sleep(0.5)
                return
            except Exception:
                pass

    # ── download ──────────────────────────────────────────────────────────────
    def _download(self, dest_folder: Path, base_name: str) -> Optional[Path]:
        """
        Download file to _dl_tmp, wait for it to complete, then move+rename to dest_folder.
        Using _dl_tmp avoids mid-session CDP path changes which are unreliable.
        """
        tmp = self._dl_tmp
        tmp.mkdir(parents=True, exist_ok=True)

        # Snapshot tmp folder before clicking
        before = set(tmp.iterdir())

        # Find and click the download button — screenshot before to diagnose
        self._shot("before_dl")
        clicked = False

        # Priority 1: direct download buttons (no submenu)
        for xp in [
            "//*[@aria-label='download' or @aria-label='הורד' or @title='הורד' or @title='Download']",
            "//*[@data-testid and (contains(@data-testid,'ownload') or contains(@data-testid,'Download'))]",
            "//*[contains(@class,'download') or contains(@class,'Download')][self::button or self::a]",
        ]:
            try:
                for el in self.driver.find_elements(By.XPATH, xp):
                    if el.is_displayed():
                        self._emit(f"  כפתור הורדה ישיר: '{el.text.strip()[:30]}'")
                        self.driver.execute_script("arguments[0].click();", el)
                        clicked = True
                        break
                if clicked:
                    break
            except Exception:
                pass

        # Priority 2: "הורדה/שליחה" split button → click it; if a submenu appears, also click "הורדה"
        if not clicked:
            try:
                split_btn = self.driver.find_element(
                    By.XPATH, "//button[contains(.,'הורדה') or contains(.,'שליחה')]"
                )
                if split_btn.is_displayed():
                    self._emit(f"  כפתור הורדה: tag=button text='{split_btn.text.strip()[:30]}'")
                    self.driver.execute_script("arguments[0].click();", split_btn)
                    clicked = True   # clicking the button itself counts — download may start now
                    time.sleep(1.5)
                    self._shot("after_split_click")
                    # If a submenu appeared, click "הורדה" within it
                    for xp in [
                        "//*[normalize-space(text())='הורדה']",
                        "//*[normalize-space(text())='הורד']",
                        "//*[contains(.,'הורדה') and not(contains(.,'שליחה'))][self::li or self::a or self::div or self::span]",
                        "//li[contains(.,'הורד')]",
                        "//*[@role='menuitem' and contains(.,'הורד')]",
                        "//*[@role='option' and contains(.,'הורד')]",
                    ]:
                        found_sub = False
                        try:
                            for opt in self.driver.find_elements(By.XPATH, xp):
                                if opt.is_displayed():
                                    self._emit(f"  לחץ submenu: '{opt.text.strip()[:30]}'")
                                    self.driver.execute_script("arguments[0].click();", opt)
                                    found_sub = True
                                    break
                        except Exception:
                            pass
                        if found_sub:
                            break
            except NoSuchElementException:
                pass

        if not clicked:
            self._shot("error_no_dl_btn")
            try:
                panel = self._get_panel()
                if panel:
                    btns = panel.find_elements(By.XPATH, ".//button | .//a")
                    visible = [(b.tag_name, b.text.strip()[:40],
                                b.get_attribute("aria-label"),
                                b.get_attribute("class")) for b in btns if b.is_displayed()]
                    self._emit(f"  כפתורים בפאנל: {visible}", "WARN")
            except Exception:
                pass
            self._emit("  כפתור הורדה לא נמצא", "WARN")
            return None

        import shutil
        click_time = time.time()
        time.sleep(1)
        self._shot("after_dl_click")   # see if a dialog/tab opened

        # Folders to watch: _dl_tmp + system Downloads (fallback)
        watch_dirs = [tmp]
        sys_dl = Path.home() / "Downloads"
        if sys_dl.exists() and sys_dl != tmp:
            watch_dirs.append(sys_dl)

        deadline = time.time() + 45
        while time.time() < deadline:
            if self.stop_event.is_set():
                return None
            time.sleep(1)
            for watch in watch_dirs:
                try:
                    candidates = [
                        f for f in watch.iterdir()
                        if f.is_file()
                        and f.suffix not in (".crdownload", ".tmp", ".part")
                        and f.stat().st_mtime >= click_time - 3
                    ]
                    if candidates:
                        src = max(candidates, key=lambda x: x.stat().st_mtime)
                        # Wait a moment to ensure fully written
                        time.sleep(0.8)
                        dest_folder.mkdir(parents=True, exist_ok=True)
                        dst = dest_folder / f"{base_name}{src.suffix}"
                        if dst.exists():
                            src.unlink()
                            return dst
                        shutil.move(str(src), str(dst))
                        self._emit(f"  ✓ קובץ הועבר: {dst.name}")
                        return dst
                except Exception as e:
                    self._emit(f"  שגיאת העברה ({watch.name}): {e}", "WARN")

        self._emit("  הורדה לא הושלמה תוך 45 שניות", "WARN")
        self._shot("dl_timeout")
        return None

    # ── logout ────────────────────────────────────────────────────────────────
    def _logout(self):
        self._emit("מתנתק…")
        try:
            for xp in [
                "//*[contains(@class,'avatar') or contains(@class,'Avatar') or contains(@class,'user-menu')]",
                "//*[@aria-label='account' or @aria-label='user' or @aria-label='profile']",
            ]:
                els = self.driver.find_elements(By.XPATH, xp)
                for el in els:
                    if el.is_displayed():
                        el.click()
                        time.sleep(1)
                        break
            for xp in ["//*[contains(.,'התנתק') or contains(.,'Logout') or contains(.,'Sign out')]"]:
                for el in self.driver.find_elements(By.XPATH, xp):
                    if el.is_displayed():
                        el.click()
                        self._emit("  ✓ התנתק")
                        return
        except Exception:
            pass
        self._emit("  התנתקות אוטומטית לא הצליחה", "WARN")

    # ── reports ───────────────────────────────────────────────────────────────
    def _save_reports(self):
        rep_dir = self.output_folder / "חשבוניות לקליטה"
        rep_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M")

        # End Report — LibreOffice Writer (.docx)
        doc_path = rep_dir / f"End Report {ts}.docx"
        if DOCX_OK:
            doc = Document()
            # Title
            title = doc.add_heading("WellyBox — End Report", level=1)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Meta
            doc.add_paragraph(f"תאריך הפעלה: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            doc.add_paragraph(f"ימים אחורה: {self.days_back}")
            doc.add_paragraph(f"תיקיית פלט: {self.output_folder}")
            doc.add_paragraph("─" * 60)
            doc.add_paragraph("")
            # Log lines
            doc.add_heading("לוג פעולות", level=2)
            for line in self._lines:
                p = doc.add_paragraph()
                run = p.add_run(line)
                run.font.size = Pt(9)
                run.font.name = "Courier New"
                if "[ERROR]" in line:
                    run.font.color.rgb = RGBColor(0xC6, 0x28, 0x28)
                elif "[WARN]" in line:
                    run.font.color.rgb = RGBColor(0xE6, 0x5C, 0x00)
            doc.save(str(doc_path))
            self._emit(f"End Report → {doc_path.name}")
            # Open in LibreOffice Writer
            import subprocess
            subprocess.Popen([
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                "--writer", str(doc_path)
            ])
        else:
            # Fallback to plain text if python-docx not available
            txt_path = rep_dir / f"End Report {ts}.txt"
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("WellyBox — End Report\n")
                f.write(f"תאריך הפעלה: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n")
                f.write(f"ימים אחורה : {self.days_back}\n")
                f.write("=" * 60 + "\n\n")
                for line in self._lines:
                    f.write(line + "\n")
            self._emit(f"End Report → {txt_path.name}")

        if not OPENPYXL_OK:
            return

        # Excel colored report
        wb  = openpyxl.Workbook()
        ws  = wb.active
        ws.title = "Report"

        # Header row
        headers = ["#", "ספק", "תאריך", "סוג מסמך", "סטטוס", "שם קובץ", "הערה"]
        hdr_fill = PatternFill("solid", fgColor="263238")
        hdr_font = Font(bold=True, color="FFFFFF")
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.fill = hdr_fill
            c.font = hdr_font

        STATUS_COLOR = {
            "downloaded_invoice": COLOR_GREEN,
            "dup_invoice":        COLOR_RED,
            "downloaded_receipt": COLOR_BLUE,
            "dup_receipt":        COLOR_YELLOW,
            "skipped":            COLOR_GREY,
            "error":              COLOR_ORANGE,
        }
        STATUS_HE = {
            "downloaded_invoice": "הורד — חשבוניות",
            "dup_invoice":        "כפול — דלג (חשבוניות)",
            "downloaded_receipt": "הורד — קבלות",
            "dup_receipt":        "כפול — דלג (קבלות)",
            "skipped":            "דלג",
            "error":              "שגיאה",
        }

        for r in self.results:
            row  = ws.max_row + 1
            fill = PatternFill("solid", fgColor=STATUS_COLOR.get(r.status, COLOR_GREY))
            vals = [r.idx, r.vendor, r.doc_date, r.doc_type,
                    STATUS_HE.get(r.status, r.status), r.filename, r.note]
            for col, v in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=v)
                c.fill = fill

        # Legend
        lg = wb.create_sheet("Legend")
        for i, (color, desc) in enumerate([
            (COLOR_GREEN,  "הורד לתיקיית חשבוניות לקליטה"),
            (COLOR_RED,    "קובץ זהה קיים בחשבוניות לקליטה — דלג"),
            (COLOR_BLUE,   "הורד לתיקיית קבלות"),
            (COLOR_YELLOW, "קובץ זהה קיים בקבלות — דלג"),
            (COLOR_GREY,   "דלג (סוג/תאריך לא תואם)"),
            (COLOR_ORANGE, "שגיאה"),
        ], 1):
            lg.cell(row=i, column=1).fill = PatternFill("solid", fgColor=color)
            lg.cell(row=i, column=2, value=desc)

        # Column widths
        for sheet in [ws, lg]:
            for col in sheet.columns:
                w = max((len(str(c.value or "")) for c in col), default=0)
                sheet.column_dimensions[get_column_letter(col[0].column)].width = min(w + 4, 55)

        xls_path = rep_dir / f"Report {ts}.xlsx"
        wb.save(xls_path)
        self._emit(f"Excel Report → {xls_path.name}")


# ── GUI ───────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.resizable(False, False)
        self._bot = None
        self._build_ui()
        # Show ready message — do NOT auto-start anything
        self.after(200, self._show_welcome)

    def _show_welcome(self):
        u, _ = load_creds()
        if u:
            self._append(f"[INFO] מוכן. משתמש: {u}", "INFO")
            self._append("[INFO] הגדר ימים אחורה ולחץ ▶ הפעל", "INFO")
        else:
            self._append("[WARN] פרטי כניסה לא נשמרו — לחץ 🔑 פרטי כניסה", "WARN")
            self._append("[INFO] לאחר מכן הגדר ימים אחורה ולחץ ▶ הפעל", "INFO")

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # Settings frame
        top = tk.LabelFrame(self, text="הגדרות", padx=8, pady=6)
        top.pack(fill="x", padx=10, pady=(10, 0))

        tk.Label(top, text="ימים אחורה:").grid(row=0, column=0, sticky="e", **pad)
        self._days = tk.StringVar(value="3")
        tk.Spinbox(top, from_=1, to=365, textvariable=self._days,
                   width=6, justify="center").grid(row=0, column=1, sticky="w", **pad)

        tk.Label(top, text="תיקיית פלט:").grid(row=1, column=0, sticky="e", **pad)
        self._folder = tk.StringVar(
            value=str(Path.home() / "Desktop")
        )
        tk.Entry(top, textvariable=self._folder, width=44).grid(
            row=1, column=1, sticky="w", **pad
        )
        tk.Button(top, text="…", width=3, command=self._browse).grid(
            row=1, column=2, padx=(0, 8)
        )

        # Buttons
        bf = tk.Frame(self)
        bf.pack(fill="x", padx=10, pady=8)

        self._run_btn = tk.Button(
            bf, text="▶  הפעל", width=14,
            bg="#2e7d32", fg="white", activebackground="#1b5e20",
            font=("Segoe UI", 10, "bold"), command=self._on_run,
        )
        self._run_btn.pack(side="left", padx=(0, 8))

        self._stop_btn = tk.Button(
            bf, text="■  עצור", width=10,
            bg="#c62828", fg="white", state="disabled",
            font=("Segoe UI", 10, "bold"), command=self._on_stop,
        )
        self._stop_btn.pack(side="left")

        tk.Button(bf, text="🔑  פרטי כניסה",
                  command=self._on_creds).pack(side="right")

        # Log
        lf = tk.LabelFrame(self, text="לוג בזמן אמת", padx=6, pady=4)
        lf.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._log = scrolledtext.ScrolledText(
            lf, width=88, height=26, state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
        )
        self._log.pack(fill="both", expand=True)
        self._log.tag_config("ERROR", foreground="#f44336")
        self._log.tag_config("WARN",  foreground="#ff9800")
        self._log.tag_config("INFO",  foreground="#d4d4d4")

        # Status bar
        self._status = tk.StringVar(value="מוכן")
        tk.Label(self, textvariable=self._status, anchor="w",
                 relief="sunken", bd=1).pack(fill="x", side="bottom")

    def _browse(self):
        d = filedialog.askdirectory(title="בחר תיקיית פלט")
        if d:
            self._folder.set(d)

    def _on_run(self):
        if not SELENIUM_OK:
            messagebox.showerror(
                "שגיאה",
                "Selenium לא מותקן.\n\npip install selenium webdriver-manager openpyxl keyring"
            )
            return
        try:
            days = int(self._days.get())
            assert days > 0
        except Exception:
            messagebox.showwarning("קלט שגוי", "נא להזין מספר ימים חיובי")
            return

        u, p = load_creds()
        if not u or not p:
            self._on_creds()
            u, p = load_creds()
            if not u or not p:
                return

        self._clear_log()
        self._set_running(True)
        self._bot = Bot(
            days_back=days,
            output_folder=Path(self._folder.get()),
            log_cb=self._append,
            done_cb=self._on_done,
        )
        self._bot.start()

    def _on_stop(self):
        if self._bot:
            self._bot.stop()
        self._status.set("עוצר…")

    def _on_creds(self):
        dlg = CredsDialog(self)
        if dlg.username and dlg.password:
            save_creds(dlg.username, dlg.password)
            self._append("[INFO] פרטי כניסה נשמרו", "INFO")

    def _on_done(self, need_creds=False):
        self.after(0, self._set_running, False)
        if need_creds:
            self.after(0, self._on_creds)

    def _set_running(self, running: bool):
        self._run_btn.config(state="disabled" if running else "normal")
        self._stop_btn.config(state="normal"  if running else "disabled")
        self._status.set("פועל…" if running else "מוכן")

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _append(self, message: str, level: str = "INFO"):
        def _do():
            self._log.config(state="normal")
            tag = level if level in ("ERROR", "WARN") else "INFO"
            self._log.insert("end", message + "\n", tag)
            self._log.see("end")
            self._log.config(state="disabled")
        self.after(0, _do)


if __name__ == "__main__":
    App().mainloop()
