import os
import re
import difflib
import csv
import uuid
import yaml
import time
from dataclasses import dataclass
from datetime import datetime
from getpass import getpass
from typing import Any, Dict, List, Optional, Set, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

from playwright.sync_api import sync_playwright, expect
from playwright.sync_api import TimeoutError as PWTimeoutError


# =========================
# Ledger
# =========================

LEDGER_HEADERS = [
    "run_id", "timestamp", "ssid", "student_name",
    "action", "details", "screenshot", "trace"
]

def suggest_closest_match(target: str, candidates: list[str], cutoff: float = 0.6) -> str | None:
    """
    Returns the closest fuzzy match from candidates, or None.
    Uses case-insensitive normalized comparison.
    """
    if not target or not candidates:
        return None

    target_norm = norm(target).casefold()
    cand_map = {norm(c).casefold(): c for c in candidates}

    matches = difflib.get_close_matches(
        target_norm,
        list(cand_map.keys()),
        n=1,
        cutoff=cutoff,
    )
    if matches:
        return cand_map[matches[0]]
    return None

def preflight_validate_ui_mappings(page, cfg: Dict[str, Any], log_fn=None, run_id=None):
    """
    Validates that config-mapped UI values actually exist in the WAI dropdowns.
    This prevents typo-induced Playwright flake later.
    """
    def _l(msg):
        if log_fn and run_id:
            _log(log_fn, f"[{run_id}] {msg}")
        else:
            _log(log_fn, msg)

    _l("Preflight: reading UI options for Race dropdown...")
    ui_races = get_combobox_options(page, re.compile(r"Race", re.I))
    ui_races_norm = {norm(x).casefold() for x in ui_races}

    # cfg["race_mapping"] maps raw input -> UI value
    bad = []
    for raw, ui_val in (cfg.get("race_mapping") or {}).items():
        if norm(ui_val).casefold() not in ui_races_norm:
            bad.append((raw, ui_val))

    if bad:
        # show first few to keep message readable
        sample = bad[:10]
        raise RuntimeError(
            "CONFIG_MAPPING_ERROR: Some race_mapping UI values do not exist in WAI Race dropdown.\n"
            f"Examples (raw -> ui_value): {sample}\n"
            f"UI Race options: {ui_races}"
        )

    _l("Preflight: race_mapping UI values all match WAI dropdown options. ✅")

def get_combobox_options(page, label_regex, timeout_ms: int = 8000) -> List[str]:
    combo = page.get_by_label(label_regex)
    expect(combo).to_be_visible(timeout=timeout_ms)
    expect(combo).to_be_enabled(timeout=timeout_ms)

    page.keyboard.press("Escape")
    page.wait_for_timeout(100)
    combo.click(timeout=3000)

    listbox = page.get_by_role("listbox")
    expect(listbox).to_be_visible(timeout=timeout_ms)

    opts = [s.strip() for s in listbox.get_by_role("option").all_inner_texts()]
    opts = [s for s in opts if s]

    page.keyboard.press("Escape")
    try:
        expect(listbox).to_be_hidden(timeout=2000)
    except Exception:
        pass

    return opts

def wait_for_toast_text(page, pattern, appear_timeout_ms: int = 8000, disappear_timeout_ms: int = 8000):
    """
    Waits for a transient toast/snackbar containing `pattern` to appear (and optionally disappear).
    `pattern` can be a compiled regex or string.
    """
    toast = page.get_by_text(pattern).first
    toast.wait_for(state="visible", timeout=appear_timeout_ms)

    # Optional: wait for it to go away (useful to ensure save is truly finished)
    try:
        toast.wait_for(state="hidden", timeout=disappear_timeout_ms)
    except PWTimeoutError:
        pass


def click_save_robust(
    page,
    timeout_ms: int = 15000,
    log_fn=None,
    run_id=None,
    toast_regex: re.Pattern | None = None,
    toast_appear_timeout_ms: int = 8000,
    toast_disappear_timeout_ms: int = 8000,
):
    def _l(msg):
        if log_fn and run_id:
            _log(log_fn, f"[{run_id}] {msg}")
        elif log_fn:
            _log(log_fn, msg)

    # Prefer a visible Save button (and usually the bottom one is the "final" save)
    save_btn = page.locator("button:has-text('Save'):visible").last

    # 1) Ensure no combobox/listbox is still open (ESC is cheap + safe)
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)

    # 2) Wait for button to be visible + enabled
    expect(save_btn).to_be_visible(timeout=timeout_ms)
    expect(save_btn).to_be_enabled(timeout=timeout_ms)

    # 3) Scroll into view (some layouts require this)
    try:
        save_btn.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass

    # 4) Click with retries
    last_err = None
    for attempt in range(1, 4):
        try:
            _l(f"Clicking Save (attempt {attempt}/3)...")
            save_btn.click(timeout=5000)

            # 5) Optional: wait for success toast text (this is the "truth signal")
            if toast_regex is not None:
                toast = page.get_by_text(toast_regex).first
                _l(f"Waiting for toast: {toast_regex.pattern!r}")
                toast.wait_for(state="visible", timeout=toast_appear_timeout_ms)

                # Best-effort: wait for it to go away (ensures UI finished its save cycle)
                try:
                    toast.wait_for(state="hidden", timeout=toast_disappear_timeout_ms)
                except PWTimeoutError:
                    pass

            return

        except Exception as e:
            last_err = e
            _l(f"Save click attempt {attempt} failed: {e}")

            # Close any portal/overlay and try again
            page.keyboard.press("Escape")
            page.wait_for_timeout(250)

            # Re-assert enabled/visible (can re-render)
            try:
                expect(save_btn).to_be_visible(timeout=3000)
                expect(save_btn).to_be_enabled(timeout=3000)
            except Exception:
                pass

    raise RuntimeError(f"Failed to click Save after retries: {last_err}")


def load_ledger_state_xlsx(ledger_path: str) -> Tuple[Set[str], Dict[str, str], Set[str]]:
    """
    Returns:
      completed: SSIDs that are terminal and should be skipped on resume
      last_action_by_ssid: last recorded action for each SSID (for SKIPPED_RESUME details)
      transfer_requested_ssids: SSIDs that have ever had TRANSFER_REQUESTED logged
        (used to emit TRANSFER_PENDING and avoid re-clicking envelope due to UI inconsistency)
    """
    completed: Set[str] = set()
    last_action_by_ssid: Dict[str, str] = {}
    transfer_requested_ssids: Set[str] = set()

    if not os.path.exists(ledger_path):
        return completed, last_action_by_ssid, transfer_requested_ssids

    wb = load_workbook(ledger_path, read_only=True, data_only=True)
    ws = wb.active

    headers = [
        str(c.value).strip() if c.value is not None else ""
        for c in next(ws.iter_rows(min_row=1, max_row=1))
    ]
    idx = {h: i for i, h in enumerate(headers)}

    def get_cell(row, col_name: str) -> str:
        i = idx.get(col_name)
        if i is None:
            return ""
        v = row[i].value
        return str(v).strip() if v is not None else ""

    for r in ws.iter_rows(min_row=2):
        ssid = get_cell(r, "ssid")
        action = get_cell(r, "action")
        if not ssid:
            continue

        if action:
            last_action_by_ssid[ssid] = action

        if action in {"TRANSFER_REQUESTED", "TRANSFER_PENDING"}:
            transfer_requested_ssids.add(ssid)

        # Terminal outcomes (including transfers, per current workflow intent)
        if action in {
            "CREATED",
            "ALREADY_OWNED",
            "SKIPPED_MISSING_SSID",
            "TRANSFER_REQUESTED",
            "TRANSFER_PENDING",
        }:
            completed.add(ssid)


    return completed, last_action_by_ssid, transfer_requested_ssids


def append_ledger_xlsx(path: str, row: Dict[str, str]) -> None:
    ensure_dir(os.path.dirname(path))

    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Ledger"
        ws.append(LEDGER_HEADERS)

        # Freeze header row
        ws.freeze_panes = "A2"

        # Bold header row
        for cell in ws[1]:
            cell.font = Font(bold=True)

        default_widths = {
            "run_id": 10,
            "timestamp": 22,
            "ssid": 12,
            "student_name": 24,
            "action": 28,
            "details": 60,
            "screenshot": 50,
            "trace": 45,
        }
        for i, h in enumerate(LEDGER_HEADERS, start=1):
            ws.column_dimensions[get_column_letter(i)].width = default_widths.get(h, 20)

    ws.append([row.get(h, "") for h in LEDGER_HEADERS])
    wb.save(path)


def select_ethnicity_radio(page, choice_text: str):
    block = page.locator("div").filter(
        has_text=re.compile(r"Ethnicity\s*-\s*Is this student Hispanic or Latino", re.I)
    ).first
    block.get_by_text(choice_text, exact=True).click()


def save_transfer_pending_screenshot(page, run_id: str, ssid: str) -> str:
    ensure_dir("output/screenshots")
    path = os.path.join("output", "screenshots", f"transfer_pending_{run_id}_{ssid}.png")
    page.screenshot(path=path, full_page=False)
    return path


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def get_found_row_owning_org(page, expected_ssid: str, timeout_ms: int = 5000) -> str:
    row = page.locator("tr").filter(has_text=re.compile(re.escape(expected_ssid))).first
    row.wait_for(state="visible", timeout=timeout_ms)

    owning_td = row.locator("td").nth(4)  # 5th td
    owning_td.wait_for(state="visible", timeout=timeout_ms)

    owning_text = " ".join((owning_td.inner_text() or "").split())
    return owning_text


def is_already_owned_by_us(found_owning_org: str, cfg: Dict[str, Any]) -> bool:
    target = norm(cfg["workability"]["owning_org_name"])
    return norm(found_owning_org).lower() == target.lower()


def _log(log_fn, msg: str) -> None:
    if log_fn:
        log_fn(msg)
    else:
        print(msg)


# =========================
# Config + Mapping
# =========================

def norm(s: Any) -> str:
    return " ".join(str(s or "").strip().split())


def load_config(config_path: str) -> Dict[str, Any]:
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    sm = cfg.get("school_mapping", {}) or {}
    normalized = {}
    collisions = {}

    for raw_key, val in sm.items():
        nk = norm(raw_key)
        if nk in normalized and normalized[nk] != val:
            collisions.setdefault(nk, []).append(raw_key)
        normalized[nk] = val

    cfg["school_mapping"] = normalized

    if collisions:
        print("WARNING: school_mapping key collisions after normalization:")
        for nk, raw_keys in collisions.items():
            print(f"  {nk!r} <= {raw_keys}")

    return cfg


def map_grade(raw_grade: Any, cfg: Dict[str, Any]) -> Optional[str]:
    return cfg["grade_mapping"].get(norm(raw_grade))


def map_disability(raw_dis: Any, cfg: Dict[str, Any]) -> Optional[str]:
    return cfg["disability_mapping"].get(norm(raw_dis))


def map_race(raw_race: Any, cfg: Dict[str, Any]) -> str:
    return cfg["race_mapping"].get(norm(raw_race), "Other")


def resolve_school(aeries_site: Any, cfg: Dict[str, Any]) -> str:
    key = norm(aeries_site)
    entry = cfg["school_mapping"].get(key)
    if not entry:
        raise KeyError(f"No school mapping for Aeries site raw={aeries_site!r} normalized={key!r}")
    if not entry.get("active", True):
        raise ValueError(f"Mapped WAI school is inactive/closed: {entry.get('wai_school')!r}")
    return entry["wai_school"]


def normalize_dob_mmddyyyy(raw: Any) -> str:
    s = norm(raw)
    if not s:
        return ""
    parts = s.split("/")
    if len(parts) != 3:
        raise ValueError(f"Invalid DOB format (expected M/D/YYYY): {s!r}")
    m = int(parts[0])
    d = int(parts[1])
    y = int(parts[2])
    return f"{m:02d}/{d:02d}/{y:04d}"


@dataclass
class Student:
    first_name: str
    last_name: str
    ssid: str
    dob: str
    gender_ui: str
    grade_ui: str
    disability_ui: str
    ethnicity_ui: str
    race_ui: str
    aeries_school: str
    wai_school: str


def validate_and_prepare(row: Dict[str, Any], cfg: Dict[str, Any]) -> Student:
    cols = cfg["columns"]

    ssid = norm(row.get(cols["ssid"]))
    if not ssid:
        raise ValueError("Missing SSID")

    gender_raw = norm(row.get(cols["gender"]))
    if gender_raw not in cfg["validation"]["allowed_gender"]:
        raise ValueError(f"Invalid gender {gender_raw!r} (allowed {cfg['validation']['allowed_gender']})")
    gender_ui = cfg["gender_mapping"][gender_raw]

    dob_ui = normalize_dob_mmddyyyy(row.get(cols["dob"]))

    grade_ui = map_grade(row.get(cols["grade"]), cfg)
    if not grade_ui:
        raise ValueError(f"Invalid grade {norm(row.get(cols['grade']))!r} (no mapping)")

    disability_ui = map_disability(row.get(cols["disability"]), cfg)
    if not disability_ui:
        raise ValueError(f"Invalid or missing disability {norm(row.get(cols['disability']))!r} (no WAI option)")

    eth_raw = norm(row.get(cols["ethcd"]))
    eth_map = {"Y": "Yes", "N": "No", "D": "Declined to State"}
    ethnicity_ui = eth_map.get(eth_raw)
    if not ethnicity_ui:
        raise ValueError(f"Invalid or missing ethnicity code {eth_raw!r} (expected Y/N/D)")

    race_ui = map_race(row.get(cols["race1"]), cfg)

    aeries_school = norm(row.get(cols["aeries_school"]))
    wai_school = resolve_school(aeries_school, cfg)

    return Student(
        first_name=norm(row.get(cols["first_name"])),
        last_name=norm(row.get(cols["last_name"])),
        ssid=ssid,
        dob=dob_ui,
        gender_ui=gender_ui,
        grade_ui=grade_ui,
        disability_ui=disability_ui,
        ethnicity_ui=ethnicity_ui,
        race_ui=race_ui,
        aeries_school=aeries_school,
        wai_school=wai_school,
    )


# =========================
# Excel Reader
# =========================

def read_excel_rows(excel_path: str) -> List[Dict[str, Any]]:
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    out: List[Dict[str, Any]] = []
    for r in rows[1:]:
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        out.append({headers[i]: r[i] for i in range(min(len(headers), len(r)))})
    return out


# =========================
# Playwright Actions
# =========================

ZERO_RESULTS_RE = re.compile(r"The following\s+0\s+records match your search criteria", re.I)


def login(page, username: str, password: str, cfg: Dict[str, Any], log_fn=None) -> None:
    def _local(msg: str):
        if log_fn:
            log_fn(msg)
        else:
            print(msg)

    _local("Navigating to login page...")
    page.goto(cfg["workability"]["login_url"])

    user = page.get_by_placeholder("Username")
    pw = page.get_by_placeholder("Password")
    btn = page.get_by_role("button", name=re.compile(r"Log\s*In", re.I))

    expect(user).to_be_visible(timeout=15000)
    expect(pw).to_be_visible(timeout=15000)

    _local("Filling credentials...")
    user.fill(username)
    pw.fill(password)

    _local("Clicking Log In...")
    btn.click()

    student_data_link = page.get_by_role("link", name=re.compile(r"Student\s*Data", re.I))

    def bad_creds_detected() -> bool:
        try:
            if not user.is_visible() or not pw.is_visible():
                return False
            pw_val = pw.input_value()
            if (pw_val or "").strip() != "":
                return False
            user_val = user.input_value()
            if (user_val or "").strip() == "":
                return True
            return True
        except Exception:
            return False

    page.wait_for_timeout(600)

    if bad_creds_detected():
        raise RuntimeError("LOGIN_BAD_CREDS: username/password rejected (login form remained and password cleared).")

    try:
        student_data_link.wait_for(state="visible", timeout=6000)
        _local("Login appears successful.")
        return
    except PWTimeoutError:
        if bad_creds_detected():
            raise RuntimeError("LOGIN_BAD_CREDS: username/password rejected (silent refresh to blank login form).")
        raise RuntimeError("LOGIN_UNKNOWN: login did not reach post-login state and did not match bad-credentials pattern.")


def goto_student_records(page, cfg=None, log_fn=None, run_id=None):
    def _log_local(msg: str):
        if log_fn:
            log_fn(msg)
        else:
            print(msg)

    _log_local("Navigating to Student Records (hover-only, no parent clicks)...")

    submenu_selector = "a[href='/student-data/search-student-records']"
    submenu = page.locator(submenu_selector)

    try:
        if submenu.is_visible():
            _log_local("Fast path: Student Records visible — clicking.")
            submenu.click()
            expect(page.get_by_placeholder("SSID")).to_be_visible(timeout=8000)
            return
    except Exception:
        pass

    parent_anchor = page.locator("a:has-text('Student Data')").first

    hover_attempts = 3
    hover_wait_ms = 350
    for attempt in range(1, hover_attempts + 1):
        try:
            _log_local(f"Hover attempt {attempt}/{hover_attempts} (force=True)...")
            parent_anchor.hover(force=True)
            page.wait_for_timeout(hover_wait_ms)
            if submenu.is_visible():
                _log_local(f"Student Records visible after hover attempt {attempt} — clicking.")
                submenu.click()
                expect(page.get_by_placeholder("SSID")).to_be_visible(timeout=8000)
                return
            _log_local(f"Not visible after hover attempt {attempt}.")
        except Exception as e:
            _log_local(f"Hover attempt {attempt} threw: {e}")

    try:
        _log_local("Dispatching mouseover/mouseenter/mousemove events via JS (no clicks)...")
        page.evaluate(
            """() => {
                const candidates = Array.from(document.querySelectorAll('a,button,div,span'));
                const parent = candidates.find(e => e.textContent && e.textContent.trim().match(/^Student\\s*Data/i));
                if (!parent) return false;
                ['mouseover','mouseenter','mousemove'].forEach(evName => {
                    const ev = new Event(evName, { bubbles: true, cancelable: true });
                    parent.dispatchEvent(ev);
                });
                return true;
            }"""
        )
        page.wait_for_timeout(400)
        if submenu.is_visible():
            _log_local("Student Records visible after dispatch events — clicking.")
            submenu.click()
            expect(page.get_by_placeholder("SSID")).to_be_visible(timeout=8000)
            return
        _log_local("Dispatch events did not reveal Student Records.")
    except Exception as e:
        _log_local(f"Dispatch events attempt failed: {e}")

    try:
        _log_local("Last-resort: removing 'hidden' class from submenu <ul> via JS (no clicking parent).")
        page.evaluate(
            """() => {
                const parent = Array.from(document.querySelectorAll('a')).find(a => a.textContent && a.textContent.trim().startsWith('Student Data'));
                if (!parent) return false;
                let ul = parent.parentElement ? parent.parentElement.querySelector('ul') : null;
                if (!ul) {
                    const sib = parent.nextElementSibling;
                    if (sib && sib.tagName && sib.tagName.toLowerCase() === 'ul') ul = sib;
                }
                if (!ul) return false;
                ul.classList.remove('hidden');
                ul.style.display = 'block';
                ul.style.maxHeight = '500px';
                return true;
            }"""
        )
        page.wait_for_timeout(300)
        if submenu.is_visible():
            _log_local("Student Records visible after DOM tweak — clicking.")
            submenu.click()
            expect(page.get_by_placeholder("SSID")).to_be_visible(timeout=8000)
            return
        _log_local("DOM tweak did not reveal Student Records.")
    except Exception as e:
        _log_local(f"DOM tweak attempt failed: {e}")

    try:
        ensure_dir("output/screenshots")
        debug_path = os.path.join("output", "screenshots", f"student_records_hover_debug_{run_id or 'noid'}.png")
        page.screenshot(path=debug_path, full_page=False)
        _log_local(f"Saved debug screenshot: {debug_path}")
    except Exception as e:
        _log_local(f"Failed to save debug screenshot: {e}")

    raise RuntimeError("Failed to reveal/click Student Records via hover-based strategies. Check screenshot.")


def set_search_project_scope(page, cfg: Dict[str, Any]) -> None:
    target = cfg["workability"]["search_project"]
    page.get_by_label("WAI Project").click()
    page.get_by_label(target, exact=True).click()


def search_by_ssid(page, ssid: str, cfg: Dict[str, Any]) -> None:
    set_search_project_scope(page, cfg)
    page.get_by_placeholder("SSID").fill(ssid)
    page.get_by_role("button", name="Search").click()


def determine_search_outcome(page, timeout_ms: int = 3000) -> str:
    edit_btn = page.get_by_role("button", name="Edit")
    zero_msg = page.locator("p").filter(has_text=ZERO_RESULTS_RE).first

    step_ms = 100
    steps = max(1, timeout_ms // step_ms)

    for _ in range(steps):
        try:
            if edit_btn.is_visible():
                return "FOUND"
        except Exception:
            pass

        try:
            if zero_msg.is_visible():
                return "NOT_FOUND"
        except Exception:
            pass

        page.wait_for_timeout(step_ms)

    raise RuntimeError("Search outcome unclear: neither Edit nor zero-results message appeared within timeout.")


def open_existing_student_edit(page) -> None:
    page.get_by_role("button", name="Edit").first.click()
    expect(page.locator("body")).to_be_visible()


def request_transfer_if_possible(page) -> str:
    transfer_btn = page.get_by_label("Request Student Transfer")
    if transfer_btn.is_visible():
        transfer_btn.click()
        dialog = page.get_by_role("alertdialog")
        expect(dialog.get_by_role("heading", name=re.compile(r"Are you sure", re.I))).to_be_visible(timeout=5000)
        dialog.get_by_role("button", name="Yes").click()
        expect(dialog).to_be_hidden(timeout=5000)
        return "TRANSFER_REQUESTED"
    return "TRANSFER_PENDING"


def select_combobox_option(page, label_regex, option_text: str, timeout_ms: int = 8000):
    """
    Select an option from a Radix/shadcn combobox (robust + deterministic).
    - Verifies the option exists in the UI list before clicking (typo-proof).
    - Re-acquires the listbox each retry to avoid stale/incorrect portals.
    - Closes the portal at the end to prevent eating the next click (Save).
    """

    def get_active_listbox():
        # Radix portals: usually there’s exactly one visible listbox when open.
        lb = page.get_by_role("listbox")
        expect(lb).to_be_visible(timeout=timeout_ms)
        return lb

    combo = page.get_by_label(label_regex)
    expect(combo).to_be_visible(timeout=timeout_ms)
    expect(combo).to_be_enabled(timeout=timeout_ms)

    # Always clear any lingering portal first
    page.keyboard.press("Escape")
    page.wait_for_timeout(100)

    # Open dropdown (retry in case a prior overlay blocks click)
    last_err = None
    for attempt in range(1, 4):
        try:
            combo.click(timeout=3000)
            listbox = get_active_listbox()
            break
        except Exception as e:
            last_err = e
            page.keyboard.press("Escape")
            page.wait_for_timeout(150)
    else:
        raise RuntimeError(f"Failed to open combobox after retries: {last_err}")

    # Pull option texts from the UI (source of truth)
    try:
        available = [s.strip() for s in listbox.get_by_role("option").all_inner_texts()]
        available = [s for s in available if s]
    except Exception:
        available = []

    target_norm = norm(option_text).casefold()
    available_norm = {norm(a).casefold() for a in available}

    # Deterministic mismatch guardrail
    if available and target_norm not in available_norm:
        page.keyboard.press("Escape")

        suggestion = suggest_closest_match(option_text, available)

        hint = f"\nDid you mean: {suggestion!r} ?" if suggestion else ""

        raise RuntimeError(
            f"UI option not found for field "
            f"{label_regex.pattern if hasattr(label_regex,'pattern') else label_regex!r}: "
            f"{option_text!r}\n"
            f"Available options ({len(available)}): {available}"
            f"{hint}"
        )
    
    # Find the option (case-insensitive exact match on the visible label)
    option_re = re.compile(rf"^{re.escape(option_text)}$", re.I)

    last_err = None
    for attempt in range(1, 4):
        try:
            # Re-acquire listbox each attempt (Radix can re-render portals)
            listbox = get_active_listbox()
            option = listbox.get_by_role("option", name=option_re).first

            expect(option).to_be_visible(timeout=timeout_ms)
            option.scroll_into_view_if_needed(timeout=2000)

            # Normal click first
            option.click(timeout=3000)
            break

        except Exception as e:
            last_err = e

            # Close & reopen to reset portal state
            page.keyboard.press("Escape")
            page.wait_for_timeout(150)

            try:
                combo.click(timeout=3000)
                listbox = get_active_listbox()
            except Exception:
                pass

            # Last attempt: sometimes Radix options get finicky; allow force click
            if attempt == 3:
                try:
                    option = page.get_by_role("option", name=option_re).first
                    option.scroll_into_view_if_needed(timeout=2000)
                    option.click(timeout=3000, force=True)
                    break
                except Exception as e2:
                    last_err = e2
    else:
        raise RuntimeError(f"Failed to click combobox option '{option_text}' after retries: {last_err}")

    # Ensure portal closes so it doesn't eat next click (Save)
    page.keyboard.press("Escape")
    try:
        expect(page.get_by_role("listbox")).to_be_hidden(timeout=2000)
    except Exception:
        pass

    # Tiny settle (animations)
    page.wait_for_timeout(100)



def select_radio_by_label_text(page, label_text: str) -> None:
    try:
        page.get_by_role("radio", name=re.compile(rf"^{re.escape(label_text)}$", re.I)).check(timeout=1500)
        return
    except Exception:
        pass

    container = page.locator("div").filter(has_text=re.compile(rf"^{re.escape(label_text)}$", re.I)).first
    for sel in ["#radio-group-item", "[role='radio']", "button[role='radio']"]:
        try:
            container.locator(sel).first.click(timeout=3000)
            return
        except Exception:
            pass

    loose = page.locator(f"text={label_text}").first
    for sel in ["xpath=ancestor::*[1]//*[@role='radio']", "xpath=ancestor::*[2]//*[@role='radio']"]:
        try:
            loose.locator(sel).first.click(timeout=3000)
            return
        except Exception:
            pass

    loose.click(timeout=3000)


def select_gender(page, gender_ui: str) -> None:
    text_map = {"M": "Male", "F": "Female", "Non Binary": "Non-Binary"}
    target_text = text_map.get(gender_ui, gender_ui)

    try:
        page.get_by_role("radio", name=re.compile(rf"^{re.escape(target_text)}$", re.I)).check(timeout=1500)
        return
    except Exception:
        pass

    container = page.locator("div").filter(has_text=re.compile(rf"^{re.escape(target_text)}$", re.I)).first
    for sel in ["#radio-group-item", "[role='radio']", "button[role='radio']"]:
        try:
            container.locator(sel).first.click(timeout=3000)
            return
        except Exception:
            pass

    container.click(timeout=3000)

def preflight_validate_mapping_against_dropdown(
    page,
    mapping: Dict[str, str],
    label_regex,
    label_name: str,
    timeout_ms: int = 8000,
):
    """
    Validates that all mapped UI values exist in the WAI dropdown options.
    Raises a clear error if any UI values are missing (typos, UI change, etc.).
    """
    ui_opts = get_combobox_options(page, label_regex, timeout_ms=timeout_ms)
    ui_norm = {norm(x).casefold() for x in ui_opts}

    bad = []
    for raw, ui_val in (mapping or {}).items():
        if norm(ui_val).casefold() not in ui_norm:
            bad.append((raw, ui_val))

    if bad:
        raise RuntimeError(
            f"CONFIG_MAPPING_ERROR: Some {label_name} mapping UI values do not exist in WAI dropdown.\n"
            f"Examples (raw -> ui_value): {bad[:10]}\n"
            f"UI {label_name} options ({len(ui_opts)}): {ui_opts}"
        )
    
def create_new_student(
    page,
    student: Student,
    cfg: Dict[str, Any],
    log_fn=None,
    run_id=None,
    on_form_expanded=None,
):
    page.get_by_role("button", name="New Student").click()

    page.get_by_placeholder("First Name").fill(student.first_name)
    page.get_by_placeholder("Last Name").fill(student.last_name)
    page.get_by_placeholder("SSID").fill(student.ssid)
    page.get_by_placeholder("MM/DD/YYYY").fill(student.dob)

    if log_fn and run_id:
        _log(log_fn, f"[{run_id}] Selecting gender: {student.gender_ui}")
    select_gender(page, student.gender_ui)

    create_project = cfg["workability"]["create_project"]
    page.get_by_label(re.compile(r"WAI Project", re.I)).click()
    page.get_by_label(create_project, exact=True).click()

    if log_fn and run_id:
        _log(log_fn, f"[{run_id}] Selecting school: {student.wai_school}")

    label = page.get_by_text(re.compile(r"^School of Attendance", re.I)).first

    control_id = None
    try:
        control_id = label.evaluate("el => el.getAttribute && el.getAttribute('for')")
    except Exception:
        control_id = None

    if control_id:
        school_select = page.locator(f"select#{control_id}")
    else:
        school_select = label.locator("xpath=following::select[1]")

    expect(school_select).to_be_visible(timeout=15000)
    expect(school_select).to_be_enabled(timeout=15000)

    for _ in range(30):
        if school_select.locator("option").count() > 1:
            break
        page.wait_for_timeout(500)
    else:
        raise RuntimeError("School dropdown options never populated (still <=1 option after 15s).")

    opt = school_select.locator("option").filter(has_text=re.compile(re.escape(student.wai_school), re.I)).first
    value = opt.get_attribute("value")
    if not value:
        first_few = school_select.locator("option").all_inner_texts()[:10]
        raise RuntimeError(f"Could not find <option value> for school '{student.wai_school}'. First options: {first_few}")

    school_select.select_option(value=value)

    selected_text = school_select.locator("option:checked").inner_text()
    if log_fn and run_id:
        _log(log_fn, f"[{run_id}] School selected (UI): {selected_text}")

    if log_fn and run_id:
        _log(log_fn, f"[{run_id}] Selecting record type: Baseline 2025-26")
    select_radio_by_label_text(page, "Baseline 2025-26")

    click_save_robust(
        page,
        timeout_ms=15000,
        log_fn=log_fn,
        run_id=run_id,
        toast_regex=re.compile(r"baseline.*successfully\s+added", re.I),
    )

    expect(page.get_by_label(re.compile(r"Grade", re.I))).to_be_visible(timeout=10000)

    # One-time preflight hook (run-level state lives in run(), but executes here where Race exists)
    if on_form_expanded:
        on_form_expanded(page)
        
    if log_fn and run_id:
        _log(log_fn, f"[{run_id}] Selecting grade: {student.grade_ui}")
    select_combobox_option(page, re.compile(r"Grade", re.I), str(student.grade_ui))

    if log_fn and run_id:
        _log(log_fn, f"[{run_id}] Selecting disability: {student.disability_ui}")
    select_combobox_option(page, re.compile(r"Disability", re.I), student.disability_ui)

    if log_fn and run_id:
        _log(log_fn, f"[{run_id}] Selecting ethnicity: {student.ethnicity_ui}")
    select_radio_by_label_text(page, student.ethnicity_ui)

    if log_fn and run_id:
        _log(log_fn, f"[{run_id}] Selecting race: {student.race_ui}")
    select_combobox_option(page, re.compile(r"Race", re.I), student.race_ui)

    page.keyboard.press("Escape")
    page.wait_for_timeout(100)

    click_save_robust(
        page,
        timeout_ms=15000,
        log_fn=log_fn,
        run_id=run_id,
        toast_regex=re.compile(r"successfully\s+saved", re.I),
    )


# =========================
# Runner
# =========================

def run(
    excel_path: str,
    config_path: str = "config/config.yaml",
    username: str | None = None,
    password: str | None = None,
    log_fn=None,
    override_headless: bool | None = None,
    stop_event=None,
    user_log_fn=None,
) -> None:

    cfg = load_config(config_path)

    did_preflight = False

    def on_form_expanded(page):
        nonlocal did_preflight
        if did_preflight:
            return
        if not cfg.get("ui", {}).get("preflight_validate", True):
            did_preflight = True
            return

        # Validate Race mapping values against the actual WAI Race dropdown options
        preflight_validate_mapping_against_dropdown(
            page,
            cfg.get("race_mapping", {}),
            re.compile(r"Race", re.I),
            "Race",
            timeout_ms=8000,
        )
        did_preflight = True
        
    # --- UI hooks (closures that see run() parameters safely) ---
    ui_hold_default = float(cfg.get("ui", {}).get("min_status_seconds", 1.5))

    def _user(msg: str, min_hold_sec: float = 0.0):
        if user_log_fn:
            user_log_fn(msg)
        else:
            _log(log_fn, msg)
        hold = float(min_hold_sec or 0.0)
        if hold > 0:
            time.sleep(hold)

    def _should_stop() -> bool:
        return bool(stop_event and stop_event.is_set())

    if "search_project" not in cfg["workability"] or "create_project" not in cfg["workability"]:
        raise KeyError("config.yaml workability must include search_project and create_project.")

    ledger_path = cfg["ledger"]["path"]
    ensure_dir("output/screenshots")
    ensure_dir("output/traces")
    ensure_dir(os.path.dirname(ledger_path))

    run_id = str(uuid.uuid4())[:8]

    username_default = cfg.get("credentials", {}).get("username", "")
    if username is None or not str(username).strip():
        username = username_default or input("WAI Username (email): ").strip()

    if password is None:
        password = getpass("WAI Password: ")

    _log(log_fn, f"[{run_id}] Starting run. Excel: {excel_path}")
    _user("🚀 Starting Workabili-Bot3000 run…", 0.5)

    rows = read_excel_rows(excel_path)
    if not rows:
        _user("⚠️ No rows found in Excel.", 0.5)
        return

    total = len(rows)

    # -------------------------
    # Run summary stats (for UI)
    # -------------------------
    stats = {
        "input_rows": len(rows),
        "created": 0,
        "transfer_requested": 0,
        "transfer_pending": 0,
        "already_owned": 0,
        "skipped": 0,
        "errors": 0,
        "stopped": 0,
    }

    def _inc(key: str, n: int = 1):
        stats[key] = int(stats.get(key, 0)) + n



    resume = bool(cfg.get("run", {}).get("resume", True))
    completed: Set[str] = set()
    last_action_by_ssid: Dict[str, str] = {}
    transfer_requested_ssids: Set[str] = set()

    if resume:
        completed, last_action_by_ssid, transfer_requested_ssids = load_ledger_state_xlsx(ledger_path)

    headless = bool(cfg.get("run", {}).get("headless", False))
    if override_headless is not None:
        headless = bool(override_headless)

    slow_mo_ms = int(cfg.get("run", {}).get("slow_mo_ms", 0))
    trace_on_failure = bool(cfg.get("trace", {}).get("on_failure", True))
    stop_on_error = bool(cfg.get("run", {}).get("stop_on_error", False))

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=slow_mo_ms if slow_mo_ms > 0 else None)
        context = browser.new_context(viewport={"width": 1600, "height": 1400})
        page = context.new_page()

        try:
            login(page, username, password, cfg, log_fn=log_fn)
        except Exception as e:
            _log(log_fn, f"[{run_id}] FATAL LOGIN ERROR -> {e}")

            # Ensure browser shuts down so UI doesn't hang
            try:
                context.close()
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass

            # IMPORTANT: bubble error up to GUI so it shows the BIG modal + shake
            raise

        if _should_stop():
            _user("🛑 Stopped by user.", 0.5)
            context.close()
            browser.close()
            return

        _log(log_fn, f"[{run_id}] Logged in. Opening Student Records...")
        _user("📚 Logged in. Opening Student Records…", 0.5)
        goto_student_records(page, cfg=cfg, log_fn=log_fn, run_id=run_id)

        for i, row in enumerate(rows):
            if _should_stop():
                _user("🛑 Stopped by user.", 0.5)
                _inc("stopped")
                break

            fn = norm(row.get(cfg["columns"]["first_name"]))
            ln = norm(row.get(cfg["columns"]["last_name"]))
            display_name = f"{fn} {ln}".strip()

            _user(f"🔎 {i+1} of {total}: Checking {display_name}…", 0.5)

            try:
                student = validate_and_prepare(row, cfg)
            except ValueError as ve:
                ssid_raw = norm(row.get(cfg["columns"]["ssid"]))

                if "Missing SSID" in str(ve):
                    action = "SKIPPED_MISSING_SSID"
                    details = "Missing SSID in input file (cannot proceed)"
                    _user(f"SKIPPED_MISSING_SSID||display_name={display_name}", ui_hold_default)
                else:
                    action = "NEEDS_INPUT_DATA"
                    details = f"Input data incomplete or invalid: {ve}"
                    _user(f"NEEDS_INPUT_DATA||display_name={display_name}", ui_hold_default)

                _log(log_fn, f"[{run_id}] {action}: {ssid_raw or '(no SSID)'} ({display_name}) -> {ve}")
                _inc("skipped")
                append_ledger_xlsx(ledger_path, {
                    "run_id": run_id,
                    "timestamp": datetime.utcnow().isoformat(),
                    "ssid": ssid_raw,
                    "student_name": display_name,
                    "action": action,
                    "details": details,   # ✅ fixed
                    "screenshot": "",
                    "trace": "",
                })
                continue

            if resume and student.ssid in completed:
                last_action = last_action_by_ssid.get(student.ssid, "UNKNOWN")
                _log(log_fn, f"[{run_id}] SKIPPED_RESUME: {student.ssid} ({display_name}) (already processed: {last_action})")
                _user(f"SKIPPED_RESUME||display_name={display_name}||last_action={last_action}", ui_hold_default)
                _inc("skipped")
                append_ledger_xlsx(ledger_path, {
                    "run_id": run_id,
                    "timestamp": datetime.utcnow().isoformat(),
                    "ssid": student.ssid,
                    "student_name": display_name,
                    "action": "SKIPPED_RESUME",
                    "details": f"Skipped (already processed in prior ledger: {last_action})",
                    "screenshot": "",
                    "trace": "",
                })
                continue

            # If transfer was already requested in a prior run, treat as pending and do NOT re-click UI
            # (You indicated transfers are effectively terminal/completed in your current workflow.)
            if student.ssid in transfer_requested_ssids:
                _log(log_fn, f"[{run_id}] TRANSFER_PENDING: {student.ssid} ({display_name}) (transfer already requested per ledger)")
                _user(f"TRANSFER_PENDING||display_name={display_name}", ui_hold_default)
                _inc("transfer_pending")
                append_ledger_xlsx(ledger_path, {
                    "run_id": run_id,
                    "timestamp": datetime.utcnow().isoformat(),
                    "ssid": student.ssid,
                    "student_name": display_name,
                    "action": "TRANSFER_PENDING",
                    "details": "Transfer already requested in a prior run (ledger source of truth). Skipping repeat request.",
                    "screenshot": "",
                    "trace": "",
                })
                # (Not adding to completed here; your “completed transfer” behavior is already satisfied by this shortcut.)
                continue

            trace_path = ""
            screenshot_path = ""

            try:
                if trace_on_failure:
                    context.tracing.start(screenshots=True, snapshots=True, sources=True)

                _log(log_fn, f"[{run_id}] Processing SSID {student.ssid} ({display_name})")

                if _should_stop():
                    _user("🛑 Stopped by user.", 0.5)
                    break

                goto_student_records(page, cfg=cfg, log_fn=log_fn, run_id=run_id)

                _log(log_fn, f"[{run_id}] Searching SSID {student.ssid} under '{cfg['workability']['search_project']}'...")
                search_by_ssid(page, student.ssid, cfg)

                outcome = determine_search_outcome(page, timeout_ms=3000)

                if outcome == "FOUND":
                    _log(log_fn, f"[{run_id}] FOUND (Edit visible). Checking owning org BEFORE Edit...")

                    owning_org = get_found_row_owning_org(page, student.ssid)
                    _log(log_fn, f"[{run_id}] Owning org for {student.ssid}: {owning_org!r}")

                    if is_already_owned_by_us(owning_org, cfg):
                        _log(log_fn, f"[{run_id}] ALREADY_OWNED: {student.ssid} ({display_name}) -> owned by us, skipping.")
                        _user(f"ALREADY_OWNED||display_name={display_name}", ui_hold_default)

                        screenshot_path = os.path.join("output", "screenshots", f"already_owned_{run_id}_{student.ssid}.png")
                        try:
                            page.screenshot(path=screenshot_path, full_page=False)
                        except Exception:
                            screenshot_path = ""
                        _inc("already_owned")
                        append_ledger_xlsx(ledger_path, {
                            "run_id": run_id,
                            "timestamp": datetime.utcnow().isoformat(),
                            "ssid": student.ssid,
                            "student_name": display_name,
                            "action": "ALREADY_OWNED",
                            "details": f"Student already enrolled in WAI under our org: {owning_org}",
                            "screenshot": screenshot_path,   # ✅ fixed
                            "trace": "",
                        })
                        completed.add(student.ssid)
                        if trace_on_failure:
                            context.tracing.stop()
                        continue

                    _log(log_fn, f"[{run_id}] FOUND (Edit visible, not SBCSS Student). Checking for envelope/transfer status...")
                    _user("📨 Student exists in another org. Checking transfer…", 0.5)

                    open_existing_student_edit(page)
                    action = request_transfer_if_possible(page)

                    details = ""
                    screenshot_path = ""

                    if action == "TRANSFER_REQUESTED":
                        details = "Requested transfer via envelope (Yes)."
                        _user(f"TRANSFER_REQUESTED||display_name={display_name}", ui_hold_default)
                        _inc("transfer_requested")
                        transfer_requested_ssids.add(student.ssid)  # so future runs shortcut
                    else:  # TRANSFER_PENDING
                        details = "Transfer pending (previously requested); waiting for release."
                        _user(f"TRANSFER_PENDING||display_name={display_name}", ui_hold_default)
                        _inc("transfer_pending")
                        try:
                            screenshot_path = save_transfer_pending_screenshot(page, run_id, student.ssid)
                        except Exception as e:
                            _log(log_fn, f"[{run_id}] WARNING: Failed to save transfer-pending screenshot: {e}")

                    _log(log_fn, f"[{run_id}] {action}: {student.ssid} ({display_name})")

                    append_ledger_xlsx(ledger_path, {
                        "run_id": run_id,
                        "timestamp": datetime.utcnow().isoformat(),
                        "ssid": student.ssid,
                        "student_name": display_name,
                        "action": action,
                        "details": details,
                        "screenshot": screenshot_path,
                        "trace": "",
                    })

                    # You asked to leave “transfers are a completed state” behavior as-is:
                    completed.add(student.ssid)

                else:
                    _log(log_fn, f"[{run_id}] NOT FOUND (0 records). Creating under '{cfg['workability']['create_project']}'...")
                    _user("📝 Not found in WAI. Entering now…", 1.0)

                    create_new_student(
                        page,
                        student,
                        cfg,
                        log_fn=log_fn,
                        run_id=run_id,
                        on_form_expanded=on_form_expanded,
                    )

                    _log(log_fn, f"[{run_id}] CREATED: {student.ssid} ({display_name})")
                    _user(f"CREATED||display_name={display_name}", 2.0)
                    _inc("created")
                    append_ledger_xlsx(ledger_path, {
                        "run_id": run_id,
                        "timestamp": datetime.utcnow().isoformat(),
                        "ssid": student.ssid,
                        "student_name": display_name,
                        "action": "CREATED",
                        "details": f"Created. Grade={student.grade_ui}; Disability={student.disability_ui}; Race={student.race_ui}; School={student.wai_school}",
                        "screenshot": "",
                        "trace": "",
                    })
                    completed.add(student.ssid)

                if trace_on_failure:
                    context.tracing.stop()

            except Exception as exc:
                if trace_on_failure:
                    trace_path = os.path.join("output", "traces", f"trace_{run_id}_{student.ssid}.zip")
                    try:
                        context.tracing.stop(path=trace_path)
                    except Exception:
                        trace_path = ""

                try:
                    screenshot_path = os.path.join("output", "screenshots", f"error_{run_id}_{student.ssid}.png")
                    page.screenshot(path=screenshot_path, full_page=True)
                except Exception:
                    screenshot_path = ""

                _log(log_fn, f"[{run_id}] ERROR: {student.ssid} ({display_name}) -> {exc}")
                _user(f"ERROR||display_name={display_name}", ui_hold_default)
                _inc("errors")
                append_ledger_xlsx(ledger_path, {
                    "run_id": run_id,
                    "timestamp": datetime.utcnow().isoformat(),
                    "ssid": student.ssid,
                    "student_name": display_name,
                    "action": "ERROR",
                    "details": str(exc),
                    "screenshot": screenshot_path,
                    "trace": trace_path,
                })

                if stop_on_error:
                    raise
                continue

        _log(log_fn, f"[{run_id}] Run complete. Ledger: {ledger_path}")
        _user("✅ Run complete. Check the ledger for details.", 0.5)

        context.close()
        browser.close()

        print(f"Run complete. Ledger: {ledger_path}")

        return {
            "run_id": run_id,
            "ledger_path": ledger_path,
            **stats,
        }


if __name__ == "__main__":
    excel = input("Path to input Excel file: ").strip('"').strip()
    run(excel_path=excel, config_path="config/config.yaml")
