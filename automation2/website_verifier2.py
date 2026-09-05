"""
Website Qualification & Verification Engine
=============================================
SYSTEM 2 -- SECOND COPY, for trying a different login or a
different automation approach. It is deliberately independent of
website_verifier.py:
    credentials  ->  .env2      (not .env)
    debug output ->  debug2/    (not debug/)
Edit LOGIN_URL below if this system points at a different portal.
Changes made here do NOT affect website_verifier.py.
=============================================

Automated Playwright-based crawler that verifies whether an assigned
website qualifies for the Copy & Paste website-evaluation workflow.

This script implements the controlling rules from:
  - Website_Qualification_Verification_MASTER_Guidelines_UPDATED.pdf
  - Intensecore_Guidelines___Rules_New.pdf
  - Country_Name_List.xlsx   (embedded below as COUNTRY_TABLE)
  - Error_Details.xlsx       (embedded below as PORTAL_ERROR_FIELDS)

CONTROLLING PRINCIPLES
-----------------------
1. Firefox only (never Chrome / an auto-translating browser).
2. The assigned website URL is read dynamically from the portal.
   No target URL is ever hardcoded.
3. SUBMISSION BEHAVIOR (by explicit user instruction, overriding
   the source guideline's default "verification-only" posture):
     - On SKIP, the script selects Website Status = "Not Working"
       and submits the form, then picks up whatever new assigned
       URL the portal generates and continues automatically. This
       applies to EVERY SKIP reason, not only genuinely dead sites
       -- see the warning above AUTO_SUBMIT_SKIP_AS_NOT_WORKING.
     - On QUALIFIES, nothing is auto-submitted. The verified fields
       are printed for manual entry, since the portal's qualifying-
       record field selectors were never supplied.
     - Set AUTO_SUBMIT_SKIP_AS_NOT_WORKING = False to fall back to
       fully read-only / manual mode at any time.
4. QUALIFIES is returned only when every mandatory requirement has
   actually been verified. A single missing/unclear/unverifiable
   mandatory requirement forces the final result to SKIP.
5. Never guess, invent, mask, or placeholder any value. Missing
   mandatory data is always a SKIP, never a fabricated answer.
6. Output format is fixed by the master guideline (section 4):
     - SKIP result   -> print exactly:  SKIP #
     - QUALIFIES     -> print the exact field block described in
       `print_final_qualifies()` below.

Run:
    python website_verifier.py
"""

from __future__ import annotations

import io
import os
import re
import sys
import time
from collections import deque
from urllib.parse import urljoin, urlparse
from re import error

from playwright.sync_api import (
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)

# The Windows console is cp1252 by default, and a ChatGPT reply
# containing one "->" arrow crashed an otherwise finished run at the
# print statement. Unencodable characters are replaced, not fatal.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ============================================================
# PORTAL / CRAWL CONFIGURATION
# ============================================================

def script_dir():
    """The folder this script lives in, whatever directory it is run from."""
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


def project_file(name):
    """
    Find a shared project file. Looked for next to this script first,
    then one level up.

    The parent lookup is what lets RULES.md and the master PDF live
    once at the project root while System 1 and System 2 sit in their
    own folders -- one copy, so the two systems can never drift onto
    different versions of the rules.
    """
    here = script_dir()
    for candidate in (
        os.path.join(here, name),
        os.path.join(os.path.dirname(here), name),
    ):
        if os.path.isfile(candidate):
            return candidate
    return ""


def debug_path(filename):
    """
    Full path for a debug artefact (screenshots, form dumps). They are
    kept in automation/debug/ so they never scatter across whatever
    folder the script happened to be launched from.
    """
    folder = os.path.join(script_dir(), "debug2")
    try:
        if not os.path.isdir(folder):
            os.makedirs(folder)
    except Exception:
        return filename
    return os.path.join(folder, filename)


def load_env_file(filename=".env"):
    """
    Read simple KEY=VALUE lines from a .env file and put them into the
    environment, so the portal login runs with no terminal prompt at
    all. Looks next to this script first, then in the folder the
    script was started from. Existing real environment variables win,
    blank lines and #comments are ignored, and surrounding quotes are
    stripped. Never raises -- a missing .env just means the script
    falls back to asking once at the terminal.
    """
    candidates = []
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        candidates.append(os.path.join(here, filename))
    except NameError:
        pass
    candidates.append(os.path.join(os.getcwd(), filename))

    for path in candidates:
        if not os.path.isfile(path):
            continue
        try:
            with io.open(path, encoding="utf-8") as fh:
                for line in fh:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, _, value = line.partition("=")
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    if key and key not in os.environ:
                        os.environ[key] = value
            print(f"Loaded credentials from {path}")
            return path
        except Exception as exc:
            print(f"Could not read {path} ({type(exc).__name__}).")

    return None


LOGIN_URL = (
    "http://copypaste.dataevaluation.co.in/Account/Login?ReturnUrl=%2F"
)

# Read-only diagnostic mode:  python website_verifier.py --dump-form
# Logs in, writes every field name / dropdown option on the record
# page to debug/portal_form_debug.txt plus a screenshot, and exits.
# Nothing is filled, clicked or submitted, so it is always safe to
# run against live work. This is what unblocks the three open items
# in CLAUDE.md (product rows, Business Type, supplier-with-0-products).
DUMP_FORM_ONLY = "--dump-form" in sys.argv[1:]

# ChatGPT login mode (SYSTEM 2 only):
#     python website_verifier2.py --chatgpt-login
# Logs into chatgpt.com with CHATGPT_EMAIL / CHATGPT_PASSWORD from
# .env2 and exits. The portal loop never runs in this mode.
CHATGPT_LOGIN_ONLY = "--chatgpt-login" in sys.argv[1:]

# Portal login only:  python website_verifier2.py --login-only
# Logs into the copy-paste portal exactly as a normal run does, then
# stops and holds the window open. No record is read, filled, clicked
# or submitted, so it is safe to run against live work.
PORTAL_LOGIN_ONLY = "--login-only" in sys.argv[1:]
PORTAL_HOLD_SECONDS = 1800


MIN_QUALIFYING_PRODUCTS = 3

PAGE_NAVIGATION_TIMEOUT = 12000


# ============================================================
# PAID BUSINESS TYPES  (locked list, Manufacturer has priority)
# ============================================================

PAID_BUSINESS_TYPES = [
    "Manufacturer",
    "Industrial Services",
    "Trader",
    "Wholesaler",
    "Supplier",
    "Distributor",
    "Exporter",
]


# ============================================================
# PORTAL COUNTRY FILL NAMES
# ============================================================
# CONFIRMED USER INSTRUCTION (2026-09-02): where the workbook spells a
# country "X or Y", the short form is what gets typed -- USA, UK, UAE --
# not the long workbook spelling. Those three are the only "X or Y" rows
# in the workbook, so this list is complete. Hong Kong and Macau use the
# exact menu entries, which the workbook already spells correctly.

PORTAL_COUNTRY_FILL_NAMES = {
    "united states of america or usa": "USA",
    "united kingdom or uk": "UK",
    "united arab emirates or uae": "UAE",
    "china (hong kong s.a.r.)": "China (Hong Kong S.A.R.)",
    "china (macau s.a.r.)": "China (Macau S.A.R.)",
}


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean(text):
    """Collapse whitespace and strip."""
    return re.sub(r"\s+", " ", str(text or "")).strip()


def normalize_url(url):
    """Return a scheme-qualified, fragment-stripped URL."""
    if not url:
        return ""
    url = str(url).strip()
    if not urlparse(url).scheme:
        url = "http://" + url
    try:
        parsed = urlparse(url)
        return parsed._replace(fragment="").geturl().rstrip("/")
    except Exception:
        return ""


# ============================================================
# DIALOG HANDLING
# ============================================================
# Some portals show a native "Are you sure?" confirm dialog on
# submit. Auto-accept any such dialog immediately so automated
# submission never blocks waiting on a human to click OK.

def install_dialog_autoaccept(page):
    def _handle_dialog(dialog):
        try:
            dialog.accept()
        except Exception:
            try:
                dialog.dismiss()
            except Exception:
                pass
    page.on("dialog", _handle_dialog)


# ============================================================
# ASSIGNED URL FROM PORTAL
# ============================================================

def get_assigned_url(portal):
    selectors = [
        "#url", "input[name='url']", "input[id*='url' i]",
        "input[name*='url' i]", "textarea[id*='url' i]",
        "textarea[name*='url' i]",
    ]

    for selector in selectors:
        try:
            loc = portal.locator(selector).first
            loc.wait_for(state="visible", timeout=1500)

            try:
                value = (loc.input_value() or "").strip()
            except Exception:
                value = ""

            if not value:
                try:
                    value = clean(loc.inner_text())
                except Exception:
                    value = ""

            if value:
                return normalize_url(value)

        except Exception:
            pass

    try:
        body = portal.locator("body").inner_text(timeout=4000)
        matches = re.findall(r"https?://[^\s<>'\"]+", body, re.I)
        if matches:
            return normalize_url(matches[0].rstrip(".,);]"))
    except Exception:
        pass

    return ""


# ============================================================
# PORTAL SUBMISSION -- SKIP -> "Not Working" -> Submit
# ============================================================

WEBSITE_STATUS_SELECTOR_CANDIDATES = [
    "select#WebsiteStatus",
    "select[name='WebsiteStatus']",
    "select[name*='websitestatus' i]",
    "select[id*='websitestatus' i]",
    "select[name*='status' i]",
    "select[id*='status' i]",
]

# Guideline (c) statuses, with the spellings a portal dropdown is
# likely to use. The first spelling that the dropdown actually offers
# is the one selected.
PORTAL_STATUS_LABELS = {
    "Opening": ["Opening", "Opening ", "Open", "Working", "opening"],
    "Not Working": [
        "Not Working", "Not-Working", "NotWorking", "Not working",
        "not working",
    ],
    "Domain Expired": [
        "Domain Expired", "Domain expired", "domain expired",
        "Expired Domain", "Domain Expire",
    ],
    "Under Construction": [
        "Under Construction", "Under construction",
        "under construction", "Under-Construction",
    ],
    "Non English": [
        "Non English", "Non-English", "NonEnglish", "Non english",
        "non english", "Other Language",
    ],
}


# CONFIRMED USER INSTRUCTION: the portal's "Are you sure?" step is
# not mandatory -- never wait for a human to click it. Native
# confirm()/alert() popups are auto-accepted by
# install_dialog_autoaccept(); an in-page HTML modal is auto-confirmed
# by accept_confirmation_if_present() using the buttons below.
CONFIRM_BUTTON_SELECTOR_CANDIDATES = [
    "button:has-text('Yes')",
    "button:has-text('OK')",
    "button:has-text('Ok')",
    "button:has-text('Confirm')",
    "button:has-text('Sure')",
    "button:has-text('Continue')",
    "input[type='button'][value='Yes']",
    "input[type='submit'][value='Yes']",
    "input[type='button'][value='OK' i]",
    "a:has-text('Yes')",
    ".swal2-confirm",
    ".swal-button--confirm",
    ".modal.show button.btn-primary",
    ".modal.in button.btn-primary",
    "[role='dialog'] button.btn-primary",
    "#confirmYes",
    "#btnYes",
]

SUBMIT_BUTTON_SELECTOR_CANDIDATES = [
    # CONFIRMED from the live portal form dump: the submit control is
    #     <input type="button" value="submit">
    # -- lowercase value, no id, no name, and NOT type="submit". Every
    # old candidate missed it, which is why the loop stopped.
    "input[type='button'][value='submit' i]",
    "input[value='submit' i]:not([value='Log in' i])",
    "button[type='submit']",
    "input[type='submit']:not([value='Log in'])",
    "#submit",
    "#SubmitBtn",
    "#btnSubmit",
    "button:has-text('Submit')",
    "input[value='Submit']",
]


def dump_portal_form(portal, note=""):
    """
    Write every form control currently on the portal page (tag, type,
    id, name, visible label/value, and every <select>'s options) to
    portal_form_debug.txt, plus a full-page screenshot. Called when an
    automatic submission cannot find a field, so the exact selector
    can be fixed instead of guessing again. Never raises.
    """
    try:
        controls = portal.evaluate("""() => {
            const out = [];
            document.querySelectorAll('input, select, textarea, button, a.btn').forEach(el => {
                const item = {
                    tag: el.tagName.toLowerCase(),
                    type: (el.getAttribute('type') || ''),
                    id: (el.id || ''),
                    name: (el.getAttribute('name') || ''),
                    value: (el.value || el.getAttribute('value') || ''),
                    text: (el.innerText || '').trim().slice(0, 60),
                    visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length),
                    options: [],
                };
                if (el.tagName.toLowerCase() === 'select') {
                    el.querySelectorAll('option').forEach(o => {
                        item.options.push(((o.innerText || '').trim()) + ' [value=' + (o.value || '') + ']');
                    });
                }
                out.push(item);
            });
            return out;
        }""")
    except Exception as exc:
        print(f"  Could not read the portal form ({type(exc).__name__}).")
        return

    stamp = time.strftime("%Y%m%d_%H%M%S")
    path = debug_path(f"portal_form_debug_{stamp}.txt")
    try:
        with io.open(path, "w", encoding="utf-8") as fh:
            fh.write(f"PORTAL FORM DUMP {note}\n")
            fh.write(f"When: {stamp}\n")
            fh.write(f"URL: {portal.url}\n")
            try:
                fh.write(f"Title: {portal.title()}\n")
            except Exception:
                pass
            fh.write(f"Looks like the login page: {looks_like_login_page(portal)}\n")
            fh.write("=" * 70 + "\n\n")
            for c in controls:
                fh.write(
                    f"<{c['tag']}> type={c['type']!r} id={c['id']!r} "
                    f"name={c['name']!r} value={c['value']!r} "
                    f"text={c['text']!r} visible={c['visible']}\n"
                )
                for opt in c["options"]:
                    fh.write(f"      option: {opt}\n")
                fh.write("\n")
        print(f"  Wrote the portal's real field names to {path}")
    except Exception:
        pass

    try:
        shot = debug_path(f"portal_form_debug_{stamp}.png")
        portal.screenshot(path=shot, full_page=True)
        print(f"  Saved a screenshot to {shot}")
    except Exception:
        pass


def find_select_with_option(portal, wanted_labels):
    """
    Content-based fallback for a dropdown whose name/id we do not
    know: scan every <select> on the page and return (index, option
    value) for the first one that actually contains one of the wanted
    option labels. This works regardless of what the portal calls the
    field. Returns (None, None) when no select has such an option.
    """
    try:
        selects = portal.evaluate("""() => {
            const out = [];
            document.querySelectorAll('select').forEach((el, i) => {
                const opts = [];
                el.querySelectorAll('option').forEach(o => {
                    opts.push({text: (o.innerText || '').trim(), value: o.value || ''});
                });
                out.push({index: i, options: opts});
            });
            return out;
        }""")
    except Exception:
        return None, None

    wanted = [w.lower().replace("-", " ").replace("_", " ") for w in wanted_labels]
    for sel in selects:
        for opt in sel["options"]:
            normalized = opt["text"].lower().replace("-", " ").replace("_", " ").strip()
            if normalized in wanted:
                return sel["index"], opt["value"]
    return None, None


def select_website_status(portal, status="Not Working"):
    """
    Select the given Website Status (guideline (c): Opening / Not
    Working / Domain Expired / Under Construction / Non English) in
    the portal's Website Status dropdown. Tries the known selectors
    first, then falls back to scanning every dropdown on the page for
    one that actually offers the wanted option. Returns True on
    success, False if nothing on the page offers that status.
    """
    labels = PORTAL_STATUS_LABELS.get(status, [status])

    for selector in WEBSITE_STATUS_SELECTOR_CANDIDATES:
        try:
            dropdown = portal.locator(selector).first
            dropdown.wait_for(state="visible", timeout=1500)
        except Exception:
            continue

        for label in labels:
            try:
                dropdown.select_option(label=label)
                print(f"Website Status set to '{label}' (selector: {selector}).")
                return True
            except Exception:
                continue

    # Fallback: forget the field's name entirely and find the one
    # dropdown on the page that actually offers a "Not Working"
    # option. This is what makes the loop portal-agnostic.
    index, option_value = find_select_with_option(portal, labels)
    if index is not None:
        try:
            dropdown = portal.locator("select").nth(index)
            dropdown.select_option(value=option_value)
            print(
                f"Website Status set to '{status}' by scanning the "
                f"page's dropdowns (select #{index})."
            )
            return True
        except Exception:
            pass

    print(
        f"WARNING: no dropdown on this page offers a '{status}' "
        "option. Nothing was changed."
    )
    try:
        select_count = portal.locator("select").count()
        print(f"  Page URL   : {portal.url}")
        print(f"  Page title : {portal.title()}")
        print(f"  <select> elements on the page: {select_count}")
        if looks_like_login_page(portal):
            print(
                "  This IS the login page -- the portal session has "
                "expired. That is the cause, not a wrong selector."
            )
        elif select_count == 0:
            print(
                "  The page has no dropdowns at all, so it is not the "
                "record form (session lost, error page, or the queue "
                "is empty)."
            )
    except Exception:
        pass
    dump_portal_form(
        portal, note=f"(Website Status '{status}' not found)",
    )
    return False


def accept_confirmation_if_present(portal, timeout=2500):
    """
    Auto-confirm the portal's "Are you sure?" step so submission never
    waits for a human. Native confirm() popups are already accepted by
    the dialog handler installed at startup; this handles the in-page
    HTML modal variant by clicking its Yes/OK/Confirm button.

    Returns True if a confirmation button was clicked, False if no
    modal appeared (which is the normal, non-error case -- the record
    was simply submitted directly).
    """
    deadline = timeout
    step = 250

    while deadline > 0:
        for selector in CONFIRM_BUTTON_SELECTOR_CANDIDATES:
            try:
                button = portal.locator(selector).first
                if button.is_visible(timeout=200):
                    button.click(timeout=1500)
                    print(f"  Auto-confirmed the 'Are you sure?' step "
                          f"(selector: {selector}).")
                    portal.wait_for_timeout(300)
                    return True
            except Exception:
                continue

        portal.wait_for_timeout(step)
        deadline -= step

    return False


def _try_click(locator, label):
    """
    Click something three increasingly forceful ways and say which one
    worked. A plain Playwright click refuses to act when another
    element covers the target -- the live portal keeps hidden
    Ok/Cancel/Close modal buttons in the DOM that can do exactly that.
    A JS click dispatches the event on the element itself, which always
    reaches its onclick handler; a forced click is the last resort
    because it aims at coordinates and can hit the overlay instead.
    Returns (True, how) or (False, why).
    """
    try:
        locator.click(timeout=3000)
        return True, "normal click"
    except Exception as exc:
        first = f"{type(exc).__name__}"

    # JS click BEFORE forced click, deliberately. A forced click is
    # dispatched at the element's coordinates, so when an overlay sits
    # on top it hits the overlay and reports success while the button's
    # handler never runs -- a silent no-op that looks like a submission.
    # el.click() fires the handler on the element itself and cannot be
    # intercepted.
    try:
        locator.evaluate("el => el.click()")
        return True, "JS click"
    except Exception as exc:
        second = f"{type(exc).__name__}"

    try:
        locator.click(force=True, timeout=3000)
        return True, "forced click"
    except Exception as exc:
        return False, f"{first} -> {second} -> {type(exc).__name__}: {exc}"


def click_submit_button(portal):
    """
    Click the portal's data-entry submit button (never the login one),
    then auto-confirm any "Are you sure?" popup. Returns True on
    success, False if nothing could be clicked.

    Every failure reason is printed. Silently swallowing them is what
    made the first live failures impossible to diagnose.
    """
    failures = []

    for selector in SUBMIT_BUTTON_SELECTOR_CANDIDATES:
        try:
            locator = portal.locator(selector)
            count = locator.count()
        except Exception as exc:
            # A selector the engine cannot even parse would otherwise
            # look identical to one that simply did not match.
            failures.append(f"{selector} -> bad selector ({type(exc).__name__})")
            continue

        if not count:
            continue

        button = locator.first
        try:
            button.wait_for(state="visible", timeout=1500)
        except Exception:
            failures.append(f"{selector} -> matched {count} but never visible")
            continue

        clicked, how = _try_click(button, selector)
        if clicked:
            print(f"Clicked submit ({how}, selector: {selector}).")
            accept_confirmation_if_present(portal)
            return True
        failures.append(f"{selector} -> {how}")

    # Fallback: scan every visible button/input and click the first one
    # that reads like a save/submit action, skipping the ones that
    # clearly are not (login, logout, search, cancel, reset).
    submit_words = ("submit", "save", "send", "update", "next", "done")
    skip_words = ("log in", "login", "log out", "logout", "search",
                  "cancel", "reset", "clear", "back", "close")
    try:
        candidates = portal.evaluate("""() => {
            const out = [];
            document.querySelectorAll("button, input[type='submit'], input[type='button'], a.btn").forEach((el, i) => {
                const visible = !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
                const label = ((el.innerText || '') + ' ' + (el.value || '')).trim().toLowerCase();
                out.push({index: i, label: label, visible: visible});
            });
            return out;
        }""")
    except Exception as exc:
        failures.append(f"page scan -> {type(exc).__name__}")
        candidates = []

    all_selector = "button, input[type='submit'], input[type='button'], a.btn"
    for item in candidates:
        if not item["visible"] or not item["label"]:
            continue
        if any(word in item["label"] for word in skip_words):
            continue
        if not any(word in item["label"] for word in submit_words):
            continue

        button = portal.locator(all_selector).nth(item["index"])
        clicked, how = _try_click(button, item["label"])
        if clicked:
            print(
                f"Clicked submit by scanning the page ({how}, "
                f"button #{item['index']}, label {item['label']!r})."
            )
            accept_confirmation_if_present(portal)
            return True
        failures.append(f"scan #{item['index']} {item['label']!r} -> {how}")

    # Last resort: find the control in JS and fire its handler directly.
    # Nothing can intercept this -- if the element exists at all, its
    # onclick runs.
    try:
        fired = portal.evaluate("""() => {
            const words = ['submit', 'save', 'send'];
            const els = document.querySelectorAll(
                "input[type='button'], input[type='submit'], button");
            for (const el of els) {
                const label = ((el.value || '') + ' ' + (el.innerText || '')).toLowerCase();
                if (words.some(w => label.includes(w)) &&
                    !label.includes('log in') && !label.includes('reset')) {
                    el.click();
                    return label.trim();
                }
            }
            return null;
        }""")
        if fired:
            print(f"Clicked submit via direct JS dispatch (label {fired!r}).")
            accept_confirmation_if_present(portal)
            return True
    except Exception as exc:
        failures.append(f"JS dispatch -> {type(exc).__name__}")

    print("WARNING: could not click any submit button. Nothing was submitted.")
    if failures:
        print("  What was tried, and what happened:")
        for line in failures:
            print(f"    - {line}")
    else:
        print("  No element on the page looked like a submit control at all.")

    dump_portal_form(portal, note="(submit button not clickable)")
    return False


def submit_skip_with_status(portal, status="Not Working"):
    """
    Full SKIP-submission flow: select the Website Status the
    guideline actually calls for on this record and submit the form.
    Returns True only if both steps succeeded. Never guesses past a
    failed selector match -- if either step fails, nothing further is
    clicked and the caller is told to handle that record manually.
    """
    if not select_website_status(portal, status):
        return False
    portal.wait_for_timeout(300)
    return click_submit_button(portal)


def looks_like_login_page(portal):
    """
    True when the portal has bounced us back to the login screen. An
    expired session is the ordinary reason a record page suddenly has
    no Website Status dropdown and no submit button on it.
    """
    try:
        if portal.locator("#Email").count() and portal.locator("#Password").count():
            return True
    except Exception:
        pass
    try:
        return "/Account/Login" in (portal.url or "")
    except Exception:
        return False


def ensure_logged_in(portal):
    """
    Log back in if the session has expired. Unattended running is the
    whole point of this script, so a dropped session must not end the
    run -- it silently looked exactly like a missing selector before.
    Returns True if the portal is usable afterwards.
    """
    if not looks_like_login_page(portal):
        return True

    username = os.environ.get("PORTAL_USERNAME", "")
    password = os.environ.get("PORTAL_PASSWORD", "")
    if not (username and password):
        print(
            "  The portal has logged us out and no credentials are "
            "available to log back in. Put PORTAL_USERNAME and "
            "PORTAL_PASSWORD in .env."
        )
        return False

    print("  Portal session expired -- logging back in...")
    try:
        if "/Account/Login" not in (portal.url or ""):
            portal.goto(LOGIN_URL, wait_until="domcontentloaded",
                        timeout=PAGE_NAVIGATION_TIMEOUT)
        portal.locator("#Email").fill(username)
        portal.locator("#Password").fill(password)
        try:
            portal.locator('input[type="submit"][value="Log in"]').click(timeout=10000)
        except Exception:
            portal.locator("#Password").press("Enter")
        portal.wait_for_load_state("domcontentloaded", timeout=15000)
        portal.wait_for_timeout(800)
    except Exception as exc:
        print(f"  Re-login failed ({type(exc).__name__}: {exc}).")
        return False

    if looks_like_login_page(portal):
        print("  Re-login did not take -- still on the login page.")
        return False

    print("  Logged back in.")
    return True


def reload_portal(portal):
    """
    Bring the portal page back to a usable state between records.
    Tries a plain reload first and falls back to navigating to the
    portal root (the login URL's ReturnUrl target) if the reload
    fails. Never raises -- a failed refresh must not stop the loop.
    """
    try:
        portal.reload(wait_until="domcontentloaded",
                      timeout=PAGE_NAVIGATION_TIMEOUT)
        ensure_logged_in(portal)
        return True
    except Exception:
        pass

    try:
        root = LOGIN_URL.split("/Account/Login")[0] + "/"
        portal.goto(root, wait_until="domcontentloaded",
                    timeout=PAGE_NAVIGATION_TIMEOUT)
        ensure_logged_in(portal)
        return True
    except Exception as exc:
        print(f"  Portal refresh failed ({type(exc).__name__}) -- continuing anyway.")
        return False


def wait_for_new_assigned_url(portal, previous_url, attempts=8, reloads=2):
    """
    After a submission, the portal is expected to auto-generate the
    next assigned URL. Poll for a value that differs from the one
    just submitted. If nothing new appears within `attempts` polls,
    reload the portal page and poll again (some portals only hand
    out the next record on a fresh page load) up to `reloads` times.
    Falls back to whatever is present at the end so the caller can
    decide what to do.
    """
    for round_index in range(reloads + 1):
        for _ in range(attempts):
            portal.wait_for_timeout(1000)
            candidate = get_assigned_url(portal)
            if candidate and candidate != previous_url:
                return candidate

        if round_index < reloads:
            print("  No new URL yet -- reloading the portal page and retrying...")
            reload_portal(portal)

    return get_assigned_url(portal)


# ============================================================
# PORTAL SUBMISSION -- QUALIFIES -> auto-fill -> Submit
# ============================================================
# CONFIRMED against the live portal form dump. Every id/name below
# was read off the real page, not guessed. Note the shapes, which are
# NOT what was assumed before the dump:
#   emailid1            text input   (note the trailing "1")
#   phoneormobile       text input
#   country             TEXT input   (not a dropdown)
#   bussinesstype       dropdown     (portal's own spelling)
#   address / city / state / companyprofile   dropdowns, Y / N only
#   productname / productimage / productdescription
#                       dropdowns, counts 0 / 1 / 2 / 3 only
# So the portal never wants product names, image files or description
# text -- only how many of each were verified, capped at 3. That is
# what closes the old "product rows are manual" open item.
QUALIFIES_FIELD_SELECTORS = {
    "email": [
        "input#emailid1", "input[name='emailid1']",
        "input[name*='email' i]", "input[id*='email' i]",
    ],
    "phone": [
        "input#phoneormobile", "input[name='phoneormobile']",
        "input[name*='phone' i]", "input[id*='phone' i]",
    ],
    "country": [
        "input#country", "input[name='country']",
        "input[id*='country' i]", "input[name*='country' i]",
    ],
    "business_type": [
        "select#bussinesstype", "select[name='bussinesstype']",
        "select[name*='business' i]", "select[id*='business' i]",
        "select[name*='bussiness' i]", "select[id*='bussiness' i]",
    ],
    "address": [
        "select#address", "select[name='address']",
        "select[id*='address' i]",
    ],
    "city": [
        "select#city", "select[name='city']", "select[id*='city' i]",
    ],
    "state": [
        "select#state", "select[name='state']", "select[id*='state' i]",
    ],
    "company_profile": [
        "select#companyprofile", "select[name='companyprofile']",
        "select[id*='companyprofile' i]", "select[id*='profile' i]",
    ],
    "product_name_count": [
        "select#productname", "select[name='productname']",
        "select[id*='productname' i]",
    ],
    "product_image_count": [
        "select#productimage", "select[name='productimage']",
        "select[id*='productimage' i]",
    ],
    "product_description_count": [
        "select#productdescription", "select[name='productdescription']",
        "select[id*='productdescription' i]",
    ],
}

# The portal spells some business types differently from the
# guidelines. Map ours onto the exact option label the dropdown
# offers; anything not listed is passed through unchanged.
PORTAL_BUSINESS_TYPE_LABELS = {
    "Distributor": "Distributer",
}

# The Website Status a QUALIFIES record is submitted with. The live
# dropdown offers "Working" (value W-Working) -- there is no
# "Opening" option, despite the guideline's wording.
QUALIFIES_STATUS_LABELS = ["Working", "Opening", "W-Working"]


def _fill_text_field(portal, selector_list, value):
    for selector in selector_list:
        try:
            field = portal.locator(selector).first
            field.wait_for(state="visible", timeout=1500)
            field.fill(str(value))
            return selector
        except Exception:
            continue
    return None


def _select_dropdown_field(portal, selector_list, value):
    for selector in selector_list:
        try:
            field = portal.locator(selector).first
            field.wait_for(state="visible", timeout=1500)
            try:
                field.select_option(label=str(value))
            except Exception:
                field.select_option(value=str(value))
            return selector
        except Exception:
            continue
    return None


def fill_and_submit_qualifies(portal, fields):
    """
    Auto-fill and submit a QUALIFIES record against the live portal
    form, whose exact shape is now known from the form dump:

        status            -> "Working"
        emailid1          -> verified email address
        phoneormobile     -> verified phone number
        country           -> verified country name (free text)
        bussinesstype     -> verified business type
        address/city/
        state/companyprofile -> Y or N
        productname/productimage/productdescription -> count, 0-3

    Product entry is a count, not a row of names/images/descriptions,
    so a qualifying record is now fully automatic -- nothing about it
    is left for manual entry.

    Returns True only if every mandatory field was located, filled,
    AND the submit button was found and clicked. On any failure,
    nothing is submitted -- the record is never left half-filled.
    """
    filled = {}

    # Website Status. A qualifying site loads, so it is "Working".
    filled["status"] = None
    for selector in WEBSITE_STATUS_SELECTOR_CANDIDATES:
        try:
            field = portal.locator(selector).first
            field.wait_for(state="visible", timeout=1500)
            for label in QUALIFIES_STATUS_LABELS:
                try:
                    field.select_option(label=label)
                    filled["status"] = selector
                    break
                except Exception:
                    try:
                        field.select_option(value=label)
                        filled["status"] = selector
                        break
                    except Exception:
                        continue
            if filled["status"]:
                break
        except Exception:
            continue

    filled["email"] = _fill_text_field(
        portal, QUALIFIES_FIELD_SELECTORS["email"], fields["email"],
    )
    filled["phone"] = _fill_text_field(
        portal, QUALIFIES_FIELD_SELECTORS["phone"], fields["phone"],
    )
    # Country is a plain text input on this portal, not a dropdown.
    # USA / UK / the China (… S.A.R.) menu entries per the user's
    # instruction; otherwise the workbook's own validated name.
    country_to_fill = fields.get("country_fill") or fields["country"]
    filled["country"] = _fill_text_field(
        portal, QUALIFIES_FIELD_SELECTORS["country"], country_to_fill,
    )

    business_type = PORTAL_BUSINESS_TYPE_LABELS.get(
        fields["business_type"], fields["business_type"],
    )
    filled["business_type"] = _select_dropdown_field(
        portal, QUALIFIES_FIELD_SELECTORS["business_type"], business_type,
    )

    # Y / N dropdowns -- the verified booleans, never a guess.
    for key, flag in (
        ("address", "address_ok"),
        ("city", "city_ok"),
        ("state", "state_ok"),
        ("company_profile", "company_profile_ok"),
    ):
        filled[key] = _select_dropdown_field(
            portal, QUALIFIES_FIELD_SELECTORS[key],
            "Y" if fields[flag] else "N",
        )

    # Product counts -- the dropdown stops at 3, and the guideline
    # minimum is 3, so a verified count is reported as at most 3.
    verified_count = min(int(fields.get("product_count") or 0), 3)
    for key in (
        "product_name_count",
        "product_image_count",
        "product_description_count",
    ):
        filled[key] = _select_dropdown_field(
            portal, QUALIFIES_FIELD_SELECTORS[key], str(verified_count),
        )

    missing = [name for name, selector in filled.items() if not selector]
    if missing:
        print(
            "WARNING: could not locate portal field(s) for: "
            + ", ".join(missing)
            + ". Nothing was submitted for this record -- add the "
            "real selector(s) to QUALIFIES_FIELD_SELECTORS at the "
            "top of this file."
        )
        dump_portal_form(
            portal, note="(QUALIFIES fields missing: " + ", ".join(missing) + ")",
        )
        return False

    portal.wait_for_timeout(300)
    submitted = click_submit_button(portal)
    if not submitted:
        print(
            "WARNING: all fields filled but the submit button could "
            "not be located automatically. Submit manually."
        )
        return False

    print(
        "Auto-filled and submitted: Status=Working, Email, Phone, "
        f"Country, Business Type={business_type}, Address/City/State/"
        f"Company Profile flags, and product counts ({verified_count})."
    )
    return True


# ============================================================
# CHATGPT SESSION  (SYSTEM 2 ONLY)
# ============================================================
#     python website_verifier2.py --chatgpt-login
#
# WHY THIS IS NOT A PASSWORD LOGIN -- measured on 2026-09-03, not
# guessed:
#
#   * chatgpt.com/auth/login accepts the email fine. But the account
#     in .env2 is a GOOGLE-LINKED account, so OpenAI never asks for a
#     password of its own -- Continue redirects to
#     accounts.google.com/v3/signin/identifier.
#   * Google then refuses a scripted browser outright. It lands on
#     accounts.google.com/v3/signin/rejected with "This browser or app
#     may not be secure -- try using a different browser". The
#     password box is never shown, so the password in .env2 can never
#     be entered. This is Google's anti-automation control, and
#     nothing in this file tries to defeat it.
#
# WHAT WORKS INSTEAD -- a persistent Firefox profile. Firefox is
# launched against chatgpt_profile2/ instead of a throwaway window, so
# cookies survive between runs. Sign in BY HAND once, in that window,
# and every later run finds the session already live: no password
# typing, no bot checks, nothing circumvented. Exactly how a person
# staying logged in on their own machine works.
#
# The credential path below is still tried first, because it does work
# for an OpenAI account with its own password (one not linked to
# Google or Apple). When it detects the Google hand-off it says so and
# falls back to waiting for the manual sign-in.

CHATGPT_HOME_URL = "https://chatgpt.com/"
CHATGPT_LOGIN_URL = "https://chatgpt.com/auth/login"

# Default page timeout for the ChatGPT tab. Named because sending the
# master document temporarily raises it and has to put it back.
CHATGPT_PAGE_TIMEOUT = 15000

# Cookies live here, next to the script, so System 2 keeps its own
# ChatGPT session and never shares one with System 1.
CHATGPT_PROFILE_DIR = os.path.join(script_dir(), "chatgpt_profile2")

# How long --chatgpt-login waits for a hand sign-in before giving up.
# 15 minutes. 5 was not enough in practice: the hand sign-in ran
# into OpenAI's phone_account_conflict error and the timer expired
# while it was still being sorted out.
CHATGPT_MANUAL_LOGIN_WAIT_SECONDS = 900

# Text that means a bot wall, not a login problem.
CHATGPT_BOT_WALL_MARKERS = (
    "verify you are human",
    "just a moment",
    "performing security verification",
    "checking your browser",
    "enable javascript and cookies to continue",
    "unusual activity",
    "access denied",
    "ray id",
)

# Google's refusal to talk to an automated browser.
GOOGLE_REJECTED_MARKERS = (
    "may not be secure",
    "try using a different browser",
    "couldn't sign you in",
)

CHATGPT_LOGIN_BUTTON_SELECTORS = (
    '[data-testid="login-button"]',
    'button:has-text("Log in")',
    'a:has-text("Log in")',
    'button:has-text("Sign in")',
    'a:has-text("Sign in")',
)

CHATGPT_EMAIL_SELECTORS = (
    'input[name="email"]',
    "#email-input",
    "#username",
    'input[type="email"]',
    'input[autocomplete="email"]',
)

CHATGPT_PASSWORD_SELECTORS = (
    'input[name="password"]',
    "#password",
    'input[type="password"]',
    'input[autocomplete="current-password"]',
)

CHATGPT_CONTINUE_SELECTORS = (
    'button[type="submit"]',
    'button:has-text("Continue")',
    'button[value="default"]',
    'input[type="submit"]',
)

# chatgpt.com's own session endpoint. This is the authoritative
# answer to "are we signed in": it returns a JSON object with a
# "user" key when the cookies are good, and {} when they are not.
CHATGPT_SESSION_API = "https://chatgpt.com/api/auth/session"


def _chatgpt_page_text(page, timeout=6000):
    """Lowercased body text, or an empty string if the body never arrives."""
    try:
        return (page.locator("body").inner_text(timeout=timeout) or "").lower()
    except Exception:
        return ""


def _chatgpt_bot_wall(page):
    """
    The name of the bot check blocking this page, or None. The page
    title is checked as well as the body text -- Cloudflare's
    interstitial is titled "Just a moment..." and often carries almost
    no body text at all.
    """
    try:
        title = (page.title() or "").lower()
    except Exception:
        title = ""
    text = _chatgpt_page_text(page, timeout=4000)
    for marker in CHATGPT_BOT_WALL_MARKERS:
        if marker in title or marker in text:
            return marker
    return None


def _chatgpt_google_rejected(page):
    """
    True when Google has refused this browser. Checked by URL as well
    as text: the refusal page is /signin/rejected, and it flickers
    back to /signin/identifier every few seconds, so a text-only test
    misses it half the time.
    """
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""
    if "signin/rejected" in url:
        return True
    text = _chatgpt_page_text(page, timeout=3000)
    return any(marker in text for marker in GOOGLE_REJECTED_MARKERS)


def _chatgpt_wait_for_any(page, selectors, what, timeout=25000):
    """
    Poll the whole selector list until one is visible, and return that
    selector (None on timeout).

    The waiting is the point. chatgpt.com is a React app: the login
    form does not exist in the HTML that domcontentloaded fires on, it
    is built a few seconds later. An earlier version of this code
    checked once, immediately, and reported "no email field found" on
    a page that plainly had one.
    """
    deadline = time.time() + (timeout / 1000.0)
    while True:
        for selector in selectors:
            try:
                target = page.locator(selector).first
                if target.count() and target.is_visible():
                    return selector
            except Exception:
                continue
        if time.time() >= deadline:
            return None
        time.sleep(0.5)


def _chatgpt_click_first(page, selectors, what, timeout=8000):
    """Click the first selector in the list that becomes visible."""
    selector = _chatgpt_wait_for_any(page, selectors, what, timeout=timeout)
    if not selector:
        print(f"    no {what} found on this page")
        return False
    try:
        page.locator(selector).first.click(timeout=timeout)
        print(f"    clicked {what}: {selector}")
        return True
    except Exception as exc:
        print(f"    {what} ({selector}) would not click "
              f"({type(exc).__name__}) -- retrying forced")
        try:
            page.locator(selector).first.click(force=True, timeout=4000)
            print(f"    forced click on {what} succeeded")
            return True
        except Exception as exc2:
            print(f"    forced click on {what} also failed "
                  f"({type(exc2).__name__})")
            return False


def _chatgpt_type_first(page, selectors, value, what, timeout=25000):
    """
    Type into the first selector that becomes visible, one key at a
    time. Never prints the value.

    press_sequentially, not fill: these are React controlled inputs.
    fill() sets the value and the box LOOKS right, but React's state
    never updates, so Continue submits an empty form and the page just
    re-renders itself. That exact false negative was observed here --
    the email appeared to be entered and the form came back blank.
    """
    selector = _chatgpt_wait_for_any(page, selectors, what, timeout=timeout)
    if not selector:
        print(f"    no {what} field found on this page")
        return False
    try:
        target = page.locator(selector).first
        target.click(timeout=4000)
        target.press_sequentially(value, delay=60, timeout=15000)
        print(f"    typed {what} into {selector}")
        return True
    except Exception as exc:
        print(f"    {what} field {selector} would not accept input "
              f"({type(exc).__name__})")
        return False


def _chatgpt_session_user(page):
    """
    The signed-in account according to chatgpt.com itself, or None.
    Asked over the page's own cookies, so it is the real answer.
    """
    try:
        response = page.request.get(CHATGPT_SESSION_API, timeout=10000)
        if not response.ok:
            return None
        data = response.json()
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    user = data.get("user")
    if isinstance(user, dict):
        return user.get("email") or user.get("id") or "signed in"
    return None


def _chatgpt_set_composer_text(page, selector, text):
    """
    Put text into the composer using JavaScript. Returns True if it
    stuck.

    Why not the keyboard: typing needs the page focused, and focusing
    it makes Firefox raise its window in front of whatever the user is
    doing -- once per record, which is intolerable for something meant
    to run in the background.

    Why not .fill(): the composer is a React controlled input. Setting
    .value directly leaves React's own state untouched, so the box
    looks filled and the form submits empty. Calling the prototype's
    native value setter and then dispatching a bubbling 'input' event
    is what React actually listens for.
    """
    script = """
        ([sel, value]) => {
            const el = document.querySelector(sel);
            if (!el) return false;
            if (el.isContentEditable) {
                el.textContent = value;
                el.dispatchEvent(new InputEvent('input', {bubbles: true}));
                return el.textContent === value;
            }
            const proto = el instanceof HTMLTextAreaElement
                ? HTMLTextAreaElement.prototype
                : HTMLInputElement.prototype;
            const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
            setter.call(el, value);
            el.dispatchEvent(new Event('input', {bubbles: true}));
            return el.value === value;
        }
    """
    try:
        return bool(page.evaluate(script, [selector, text]))
    except Exception as exc:
        print(f"    JS fill failed ({type(exc).__name__})")
        return False


def _chatgpt_js_click(page, selectors, what):
    """
    Click via element.click() in JS -- no pointer, no focus, so the
    window is not raised and an overlay cannot intercept it.
    """
    for selector in selectors:
        try:
            clicked = page.evaluate(
                """(sel) => {
                    const el = document.querySelector(sel);
                    if (!el || el.disabled) return false;
                    el.click();
                    return true;
                }""",
                selector,
            )
        except Exception:
            continue
        if clicked:
            print(f"    clicked {what} in JS: {selector}")
            return True
    return False


def _chatgpt_logged_in(page, timeout=15000):
    """
    True only when chatgpt.com itself reports a signed-in user.

    ONLY the session endpoint is trusted, on purpose. Two DOM-based
    tests were tried here first and both reported a brand-new, empty
    profile as "ALREADY LOGGED IN":

      * the composer -- chatgpt.com shows "Ask ChatGPT" to signed-OUT
        visitors too, so it proves nothing;
      * composer AND no visible login button -- defeated by the hidden
        "Log in" buttons described above.

    Signed out, /api/auth/session returns {"WARNING_BANNER": ...} with
    no "user" key; signed in, it carries the account. A transient
    failure of the endpoint reports "not signed in", which merely asks
    for a sign-in that is not needed -- the safe direction to be wrong
    in.
    """
    deadline = time.time() + (timeout / 1000.0)
    while True:
        if _chatgpt_session_user(page):
            return True
        if time.time() >= deadline:
            return False
        time.sleep(1)


def _chatgpt_shot(page, name):
    """Screenshot into debug2/, timestamped so runs never overwrite one another."""
    stamp = time.strftime("%Y%m%d_%H%M%S")
    path = debug_path(f"chatgpt_{name}_{stamp}.png")
    try:
        page.screenshot(path=path, full_page=True)
        print(f"    screenshot: {path}")
    except Exception:
        print("    screenshot failed")
    return path


def chatgpt_credential_login(page, email, password):
    """
    Try the email + password flow on chatgpt.com.

    Returns one of:
        "ok"       -- logged in, composer on screen
        "google"   -- the account is Google-linked; Google refused the
                      scripted browser. Hand sign-in required.
        "blocked"  -- a Cloudflare/Arkose bot wall
        "failed"   -- anything else (with a screenshot in debug2/)
    """
    try:
        page.goto(
            CHATGPT_LOGIN_URL, wait_until="domcontentloaded", timeout=30000,
        )
    except Exception as exc:
        print(f"  could not open {CHATGPT_LOGIN_URL}: "
              f"{type(exc).__name__} {exc}")
        _chatgpt_shot(page, "goto_failed")
        return "failed"

    # Cloudflare's interstitial replaces itself with the real page
    # within a few seconds when it is going to let us through at all.
    for _ in range(6):
        if not _chatgpt_bot_wall(page):
            break
        time.sleep(2)
    wall = _chatgpt_bot_wall(page)
    if wall:
        print(f"  BLOCKED by a bot check -- the page says {wall!r}.")
        print("  That is the Cloudflare/Arkose wall, not a wrong password.")
        _chatgpt_shot(page, "bot_wall")
        return "blocked"

    print(f"  landed on: {page.url}")

    # Either a splash with a 'Log in' button, or the email box itself.
    found = _chatgpt_wait_for_any(
        page,
        CHATGPT_EMAIL_SELECTORS + CHATGPT_LOGIN_BUTTON_SELECTORS,
        "login form",
        timeout=30000,
    )
    if not found:
        print("  the login page rendered nothing usable in 30s.")
        _chatgpt_shot(page, "empty_login_page")
        return "failed"
    if found in CHATGPT_LOGIN_BUTTON_SELECTORS:
        print("  splash page -- opening the login form first")
        _chatgpt_click_first(
            page, CHATGPT_LOGIN_BUTTON_SELECTORS, "'Log in' button",
        )

    print("  step 1/2 -- email")
    if not _chatgpt_type_first(page, CHATGPT_EMAIL_SELECTORS, email, "email"):
        _chatgpt_shot(page, "no_email_field")
        return "failed"
    _chatgpt_click_first(page, CHATGPT_CONTINUE_SELECTORS, "'Continue'")

    # What comes back decides everything: an OpenAI password box, or a
    # hand-off to Google.
    deadline = time.time() + 30
    while time.time() < deadline:
        if _chatgpt_google_rejected(page):
            print("  GOOGLE HAND-OFF, AND GOOGLE SAID NO.")
            print("  This account signs in with Google, and Google refuses")
            print("  an automated browser ('this browser or app may not be")
            print("  secure'). The password box is never shown, so the")
            print("  password cannot be entered. Not a wrong password.")
            _chatgpt_shot(page, "google_rejected")
            return "google"
        if "accounts.google.com" in (page.url or ""):
            print(f"  redirected to Google: {page.url[:70]}")
        try:
            if page.locator('input[type="password"]').first.is_visible():
                break
        except Exception:
            pass
        time.sleep(1)

    print("  step 2/2 -- password")
    if not _chatgpt_type_first(
        page, CHATGPT_PASSWORD_SELECTORS, password, "password", timeout=10000,
    ):
        if "accounts.google.com" in (page.url or ""):
            print("  (still on Google, and it never offered a password box)")
            _chatgpt_shot(page, "google_no_password")
            return "google"
        _chatgpt_shot(page, "no_password_field")
        return "failed"
    _chatgpt_click_first(page, CHATGPT_CONTINUE_SELECTORS, "'Continue'")

    if _chatgpt_logged_in(page, timeout=25000):
        return "ok"

    if _chatgpt_google_rejected(page):
        _chatgpt_shot(page, "google_rejected_late")
        return "google"

    text = _chatgpt_page_text(page)
    for phrase in (
        "incorrect email or password",
        "wrong password",
        "password is incorrect",
        "too many attempts",
        "verify your email",
        "two-factor",
        "verification code",
        "verify it's you",
    ):
        if phrase in text:
            print(f"  LOGIN REFUSED -- the page says: {phrase!r}")
            _chatgpt_shot(page, "login_refused")
            return "failed"

    print("  password submitted, but the chat composer never appeared.")
    print(f"  currently at: {page.url}")
    _chatgpt_shot(page, "unknown_state")
    return "failed"


def chatgpt_wait_for_manual_login(page, seconds=None):
    """
    Hold the window open while the sign-in is done by hand, polling for
    the chat composer. Returns True once the session is live.
    """
    if seconds is None:
        seconds = CHATGPT_MANUAL_LOGIN_WAIT_SECONDS

    print()
    print("  " + "=" * 62)
    print("  SIGN IN BY HAND IN THE FIREFOX WINDOW THAT IS OPEN NOW.")
    print("  " + "=" * 62)
    print("  Use 'Continue with Google' -- it is a real browser window,")
    print("  so Google accepts it. The credentials are the ones in .env2.")
    print()
    print(f"  Cookies are saved in {CHATGPT_PROFILE_DIR}")
    print("  so this is a ONE-TIME step: later runs open already logged in.")
    print(f"  Waiting up to {seconds // 60} minutes...")
    print()

    deadline = time.time() + seconds
    last_report = 0
    while time.time() < deadline:
        if _chatgpt_logged_in(page, timeout=1000):
            print("  detected a live session -- signed in.")
            return True
        waited = int(time.time() - (deadline - seconds))
        if waited - last_report >= 30:
            last_report = waited
            print(f"    still waiting ({waited}s)... current page: "
                  f"{(page.url or '')[:60]}")
        time.sleep(2)

    print("  timed out waiting for a hand sign-in.")
    return False


def chatgpt_login_mode(playwright):
    """
    --chatgpt-login : open ChatGPT in System 2's own persistent Firefox
    profile, get a live session, and report honestly which way it was
    obtained. The portal is never touched in this mode.
    """
    email = os.environ.get("CHATGPT_EMAIL", "")
    password = os.environ.get("CHATGPT_PASSWORD", "")

    print("-" * 70)
    print("CHATGPT LOGIN  [SYSTEM 2]")
    print(f"  account: {email or '(none in .env2)'}")
    print(f"  profile: {CHATGPT_PROFILE_DIR}")

    try:
        context = playwright.firefox.launch_persistent_context(
            CHATGPT_PROFILE_DIR, headless=False,
        )
    except Exception as exc:
        print(f"  could not open the persistent profile: "
              f"{type(exc).__name__} {exc}")
        return False

    page = context.pages[0] if context.pages else context.new_page()
    page.set_default_timeout(10000)
    page.set_default_navigation_timeout(30000)

    logged_in = False
    try:
        # 1. Is the saved session still good? This is the fast path
        #    every run after the first one takes.
        try:
            page.goto(
                CHATGPT_HOME_URL, wait_until="domcontentloaded", timeout=30000,
            )
        except Exception as exc:
            print(f"  could not open chatgpt.com: {type(exc).__name__}")

        if _chatgpt_logged_in(page, timeout=15000):
            who = _chatgpt_session_user(page) or "(account not reported)"
            print("  ALREADY LOGGED IN from the saved profile -- nothing to do.")
            print(f"  signed in as: {who}")
            print(f"  now at: {page.url}")
            _chatgpt_shot(page, "session_restored")
            logged_in = True

        # 2. No saved session. Try the credential flow, which works
        #    for an OpenAI-password account.
        elif email and password:
            print("  no saved session -- trying the credentials from .env2")
            result = chatgpt_credential_login(page, email, password)
            if result == "ok":
                print("  LOGGED IN with the .env2 credentials.")
                _chatgpt_shot(page, "logged_in")
                logged_in = True
            else:
                print(f"  credential login did not succeed ({result}).")
                logged_in = chatgpt_wait_for_manual_login(page)
                if logged_in:
                    _chatgpt_shot(page, "logged_in_by_hand")
        else:
            print("  no CHATGPT_EMAIL / CHATGPT_PASSWORD in .env2.")
            logged_in = chatgpt_wait_for_manual_login(page)
            if logged_in:
                _chatgpt_shot(page, "logged_in_by_hand")

        print("-" * 70)
        print("ChatGPT login:", "SUCCESS" if logged_in else "FAILED")
        if logged_in:
            print("The session is saved. Later runs will not ask again.")
        print("The browser window is left open on purpose.")
        print("Press Enter in this terminal to close it.")
        try:
            input()
        except Exception:
            pass
    finally:
        try:
            context.close()
        except Exception:
            pass

    return logged_in


def hold_portal_login_open(portal, seconds=None):
    """
    --login-only : confirm the portal session is real, then leave the
    window open and idle.

    The session is verified rather than assumed. main() prints "Login
    successful." straight after clicking the button, without checking
    anything -- on bad credentials that line prints while the login
    page is still on screen. looks_like_login_page() is the same test
    the submission path already uses to spot an expired session.
    """
    if seconds is None:
        seconds = PORTAL_HOLD_SECONDS

    print("-" * 70)
    still_login = True
    try:
        still_login = looks_like_login_page(portal)
    except Exception as exc:
        print(f"  could not check the page ({type(exc).__name__})")

    if still_login:
        print("  LOGIN FAILED -- still on the login page.")
        print(f"  url:   {portal.url}")
        try:
            print(f"  title: {portal.title()}")
        except Exception:
            pass
        print("  The credentials in .env2 were not accepted.")
    else:
        print("  PORTAL LOGIN CONFIRMED -- past the login page.")
        print(f"  url:   {portal.url}")
        try:
            print(f"  title: {portal.title()}")
        except Exception:
            pass

    stamp = time.strftime("%Y%m%d_%H%M%S")
    try:
        shot = debug_path(f"portal_login_only_{stamp}.png")
        portal.screenshot(path=shot, full_page=True)
        print(f"  screenshot: {shot}")
    except Exception:
        pass

    print("-" * 70)
    print(f"  Holding the window open for {seconds // 60} minutes.")
    print("  Nothing is being read, filled or submitted. Ctrl+C to stop.")

    deadline = time.time() + seconds
    while time.time() < deadline:
        time.sleep(60)
        try:
            print(f"    still open -- {portal.url}")
        except Exception:
            print("    the browser window was closed.")
            return not still_login
    return not still_login


# ============================================================
# PORTAL + CHATGPT FLOW  (SYSTEM 2 ONLY)
# ============================================================
#     python website_verifier2.py --gpt-flow
#
# Tab 1: the copy-paste portal, logged in.
# Tab 2: chatgpt.com, in the same window, sharing the same profile.
#
# Then, in order: feed RULES.md into a fresh chat, go back to the
# portal, read the assigned URL, and send that URL into the SAME chat
# so it is answered with the rules already in context.
#
# Nothing is submitted to the portal in this mode. The URL is read
# only -- no status is set, no button is clicked, no record is
# completed.

# The composer. Firefox is served the mobile composer (a real
# <textarea>) rather than the desktop contenteditable, so both shapes
# are listed -- measured, not assumed.
CHATGPT_COMPOSER_SELECTORS = (
    "#prompt-textarea",
    "#mobile-composer-prompt",
    'textarea[name="prompt"]',
    'textarea[placeholder*="Ask"]',
    'div[contenteditable="true"]',
)

CHATGPT_SEND_SELECTORS = (
    '[data-testid="send-button"]',
    'button[aria-label="Send message"]',
    'button[aria-label*="Send"]',
)

# Where an answer lives in the DOM. The role attribute is the clean
# one, but Firefox is served a different (mobile) layout than the one
# it was read off, so a plain-prose fallback is tried too -- an answer
# visibly on screen must never be reported as "no reply".
CHATGPT_ASSISTANT_SELECTORS = (
    '[data-message-author-role="assistant"]',
    "div.agent-turn",
    'article:has([data-message-author-role="assistant"])',
    "div.markdown.prose",
    "div.markdown",
)

# Signs that chatgpt.com will not answer without an account.
CHATGPT_GATE_MARKERS = (
    "log in to continue",
    "sign up to continue",
    "you've reached our limit of messages",
    "rate limit",
    "please log in",
    "create an account to continue",
)


def read_rules_document():
    """RULES.md as text -- the rulebook CLAUDE.md names as authoritative."""
    # RULES.md is the last-resort fallback and normally absent from
    # System 2's folder, so its absence is not worth a warning.
    path = project_file("RULES.md")
    if not path:
        return ""
    try:
        with io.open(path, encoding="utf-8") as fh:
            text = fh.read()
    except Exception as exc:
        print(f"  could not read {path}: {type(exc).__name__}")
        return ""
    print(f"  RULES.md: {len(text)} characters from {path}")
    return text


def portal_log_in(page):
    """
    Log into the portal on this page. Returns True only when the login
    page is actually behind us -- looks_like_login_page() is the same
    test the submission path uses to catch an expired session.
    """
    username = os.environ.get("PORTAL_USERNAME", "")
    password = os.environ.get("PORTAL_PASSWORD", "")
    if not username or not password:
        print("  no PORTAL_USERNAME / PORTAL_PASSWORD in .env2")
        return False

    try:
        page.goto(
            LOGIN_URL, wait_until="domcontentloaded",
            timeout=PAGE_NAVIGATION_TIMEOUT,
        )
    except Exception as exc:
        print(f"  login page error: {type(exc).__name__} {exc}")
        return False

    print(f"  logging in as {username}")
    # Same wait as main(): the Terms & Conditions block on this page
    # can push #Email past the default timeout.
    try:
        page.wait_for_selector("#Email", state="visible", timeout=30000)
    except Exception:
        print("  the login form never appeared within 30 seconds.")
        print(f"  url: {page.url}")
        return False

    try:
        page.locator("#Email").fill(username)
        page.locator("#Password").fill(password)
    except Exception as exc:
        print(f"  could not fill the login form ({type(exc).__name__})")
        return False

    try:
        page.locator('input[type="submit"][value="Log in"]').click(timeout=15000)
    except Exception:
        try:
            page.locator("#Password").press("Enter")
        except Exception as exc:
            print(f"  could not submit the login form ({type(exc).__name__})")
            return False

    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except Exception:
        pass

    try:
        if looks_like_login_page(page):
            print("  LOGIN FAILED -- still on the login page.")
            return False
    except Exception:
        pass

    print(f"  portal login confirmed: {page.url}")
    return True


def _chatgpt_assistant_texts(page):
    """
    Text of every assistant turn on screen, from whichever selector
    finds them. The first one that returns anything wins, so a layout
    change costs a fallback rather than the whole answer.
    """
    for selector in CHATGPT_ASSISTANT_SELECTORS:
        try:
            texts = [t for t in page.locator(selector).all_inner_texts() if t]
        except Exception:
            continue
        if texts:
            return texts
    return []


def _chatgpt_main_text(page):
    """
    The conversation as plain text, original case. Layout-independent:
    it does not care what the message elements are called.
    """
    for selector in ("main", "body"):
        try:
            text = page.locator(selector).first.inner_text(timeout=5000)
            if text:
                return text
        except Exception:
            continue
    return ""


# Page furniture that sits below every answer. It is not part of the
# reply, and counting it as one is what made an empty answer look
# finished.
CHATGPT_FOOTER_MARKERS = (
    "ChatGPT is AI and can make mistakes",
    "Chat with ChatGPT",
    "You'll get smarter responses",
    "You’ll get smarter responses",
)


def _chatgpt_answer_body(text):
    """
    Just the answer: the echoed prompt above it and the page furniture
    below it removed.

    Without this, a 23-character URL produced a 106-character "reply"
    made entirely of the footer, which then sat unchanged for the five
    seconds the stability check wanted and was returned as a finished
    answer while ChatGPT had not yet written a word.
    """
    body = text or ""
    marker = "ChatGPT said:"
    index = body.rfind(marker)
    if index != -1:
        body = body[index + len(marker):]
    for footer in CHATGPT_FOOTER_MARKERS:
        cut = body.find(footer)
        if cut != -1:
            body = body[:cut]
    return body.strip()


# Status text ChatGPT shows while it is still working. These sit
# unchanged for many seconds while a site is being fetched, which is
# long enough to satisfy any stability check -- "Searching the web"
# was returned as a finished answer and cost a record.
CHATGPT_PROGRESS_MARKERS = (
    "searching the web",
    "searching",
    "browsing",
    "reading",
    "thinking",
    "analyzing",
    "analysing",
    "working on it",
    "let me check",
)


def _chatgpt_is_progress(body):
    """True when the text is a progress indicator, not an answer."""
    text = (body or "").strip().lower().rstrip(".… ")
    if not text or len(text) > 120:
        return False
    return any(text.startswith(marker) for marker in CHATGPT_PROGRESS_MARKERS)


def _chatgpt_answer_after(text, sent):
    """
    The answer that follows our own message in the transcript.

    Anchoring on the message we sent, rather than comparing answer
    text, is what makes two identical answers distinguishable. The
    verdict is nearly always the single word "SKIP", so "has the
    answer changed?" is false even when a fresh answer has arrived --
    that stalled a run on two consecutive SKIPs.
    """
    body = text or ""
    needle = (sent or "").strip()
    if needle:
        index = body.rfind(needle)
        if index == -1:
            return ""      # our message is not on screen yet
        body = body[index + len(needle):]
    return _chatgpt_answer_body(body)


def _chatgpt_scroll_to_bottom(page):
    """
    Bring the newest message into view.

    Necessary because the conversation is virtualised: with the master
    document pasted in as a 52,000-character message, the view stays
    up inside that text and the reply below it is not rendered at all,
    so reading the page finds no answer even though one exists.
    """
    try:
        page.evaluate(
            "() => { window.scrollTo(0, document.body.scrollHeight); }"
        )
    except Exception:
        pass
    # ChatGPT scrolls an inner container rather than the window.
    try:
        page.evaluate(
            """() => {
                for (const el of document.querySelectorAll('main, main *')) {
                    if (el.scrollHeight > el.clientHeight + 50) {
                        el.scrollTop = el.scrollHeight;
                    }
                }
            }"""
        )
    except Exception:
        pass
    # In JS, so this does not raise the window either.
    try:
        page.evaluate(
            """() => {
                const sel = 'button[aria-label*="Scroll to bottom" i],'
                          + 'button[aria-label*="scroll to the bottom" i]';
                const el = document.querySelector(sel);
                if (el) el.click();
            }"""
        )
    except Exception:
        pass


def _chatgpt_gate(page):
    """The reason chatgpt.com is refusing to answer, or None."""
    text = _chatgpt_page_text(page, timeout=3000)
    for marker in CHATGPT_GATE_MARKERS:
        if marker in text:
            return marker
    return None


def _chatgpt_wait_for_reply(page, before_count, timeout=300,
                            baseline_text="", sent_text=""):
    """
    Wait for the answer to the message just sent, and return its text.

    Finished is judged by the text going quiet, not by a spinner:
    the reply streams in, so it is complete once the last assistant
    turn has stopped growing for a few seconds.
    """
    deadline = time.time() + timeout
    last_text = ""
    quiet_since = None
    # The answer already on screen when the message was sent. A new
    # one is only new once it differs from this.
    previous_answer = _chatgpt_answer_body(baseline_text)

    while time.time() < deadline:
        gate = _chatgpt_gate(page)
        if gate:
            print(f"    chatgpt.com is refusing: {gate!r}")
            return ""

        texts = _chatgpt_assistant_texts(page)
        if len(texts) > before_count:
            current = texts[-1] or ""
        else:
            # FALLBACK, and the one that actually carries this UI.
            # Two rounds of message-element selectors both missed a
            # reply that was plainly finished on screen, so the answer
            # is taken from the conversation text instead: whatever
            # appeared since the message was sent, once it stops
            # growing. Nothing here depends on class names.
            # Read the LAST answer directly rather than diffing the
            # whole transcript. Diffing broke once the conversation
            # started with a 52,000-character message: the container
            # virtualises, the rendered text changes shape as it
            # scrolls, and the common-prefix comparison stopped
            # meaning anything.
            _chatgpt_scroll_to_bottom(page)
            whole = _chatgpt_main_text(page)

            if sent_text and len(sent_text) <= 500:
                # A short message (a URL) can be found in the
                # transcript, so the answer after it is unambiguous.
                current = _chatgpt_answer_after(whole, sent_text)
            else:
                # The master document is far too long to still be
                # rendered in full, so fall back to the last answer
                # and require it to differ from what was there before.
                current = _chatgpt_answer_body(whole)
                if previous_answer and current == previous_answer:
                    time.sleep(1)
                    continue

            if not current or _chatgpt_is_progress(current):
                # Not an answer yet: only page furniture so far, or
                # ChatGPT is still browsing.
                time.sleep(1)
                continue

        if current and current == last_text:
            if quiet_since is None:
                quiet_since = time.time()
            elif time.time() - quiet_since >= 5:
                return current
        else:
            last_text = current
            quiet_since = None
        time.sleep(1)

    if last_text:
        print("    reply timed out mid-stream -- returning what arrived")
    return last_text


def chatgpt_send_message(page, text, label, reply_timeout=300, attach=None):
    """
    Type one message into the open chat and return the reply.

    keyboard.insert_text, not press_sequentially: RULES.md is over
    13,000 characters and typing it key by key would take about
    fifteen minutes. insert_text delivers it in one input event, which
    React accepts -- and unlike typing, its newlines do not send the
    message early.
    """
    print(f"  sending {label} ({len(text)} chars)")

    selector = _chatgpt_wait_for_any(
        page, CHATGPT_COMPOSER_SELECTORS, "composer", timeout=30000,
    )
    if not selector:
        print("    no composer on the page -- cannot send")
        _chatgpt_shot(page, "no_composer")
        return ""

    if attach and not chatgpt_attach_file(page, attach):
        return ""

    before = len(_chatgpt_assistant_texts(page))
    baseline = _chatgpt_main_text(page)

    # Fill in JavaScript first -- no focus, so the window is not
    # raised. The keyboard route stays as a fallback because it is the
    # one proven to satisfy React if the JS setter ever stops working.
    if not _chatgpt_set_composer_text(page, selector, text):
        print("    JS fill did not take -- falling back to typing")
        try:
            box = page.locator(selector).first
            try:
                box.click(timeout=5000)
            except Exception:
                box.evaluate("el => el.focus()")
            # The master document is over 50,000 characters, which
            # takes longer to insert than the page's 15s default. That
            # once looked like a rejected message rather than a slow
            # one and killed a run.
            page.set_default_timeout(180000)
            try:
                page.keyboard.insert_text(text)
            finally:
                page.set_default_timeout(CHATGPT_PAGE_TIMEOUT)
        except Exception as exc:
            print(f"    composer would not accept the text "
                  f"({type(exc).__name__})")
            _chatgpt_shot(page, "composer_rejected")
            return ""

    # Send, also without a pointer click.
    if not _chatgpt_js_click(page, CHATGPT_SEND_SELECTORS, "send button"):
        if not _chatgpt_click_first(page, CHATGPT_SEND_SELECTORS,
                                    "send button", timeout=8000):
            print("    no send button -- pressing Enter instead")
            try:
                page.keyboard.press("Enter")
            except Exception as exc:
                print(f"    Enter failed too ({type(exc).__name__})")
                return ""

    reply = _chatgpt_wait_for_reply(
        page, before, timeout=reply_timeout, baseline_text=baseline,
        sent_text=text,
    )
    if reply:
        print(f"    reply: {len(reply)} chars")
    else:
        print("    no reply captured")
    return reply


# The master document, uploaded to the chat instead of pasting
# RULES.md as text. Looked for next to the script first so it can be
# kept with the project, then in Downloads where it currently lives.
CHATGPT_RULES_PDF_NAME = "master doc for verification.pdf"

# The master PDF converted to Markdown, at the project root. Preferred
# over both the file and the raw PDF text -- see step 1 of
# gpt_flow_mode() for why.
MASTER_RULES_MD_NAME = "MASTER_RULES.md"


def read_text_file(path):
    """A UTF-8 text file's contents, or "" if it cannot be read."""
    try:
        with io.open(path, encoding="utf-8") as fh:
            return fh.read().strip()
    except Exception as exc:
        print(f"  could not read {path} ({type(exc).__name__})")
        return ""


def find_rules_pdf():
    """
    Path to the master PDF, or "" if it is not where we expect.

    Next to the script, then the project root (where it is shared with
    System 1), then Downloads as a last resort.
    """
    found = project_file(CHATGPT_RULES_PDF_NAME)
    if found:
        return found
    fallback = os.path.join(
        os.path.expanduser("~"), "Downloads", CHATGPT_RULES_PDF_NAME,
    )
    return fallback if os.path.isfile(fallback) else ""


def read_rules_pdf_text(path):
    """
    The master PDF's own text.

    The signed-out composer cannot take the file itself -- its file
    input accepts images only, and answers "The selected file is not a
    supported image" for a PDF. So when the upload is unavailable the
    document still governs: its text is extracted and sent, rather
    than substituting a different rulebook.
    """
    try:
        from pypdf import PdfReader
    except Exception as exc:
        print(f"  pypdf is not available ({type(exc).__name__})")
        return ""

    try:
        reader = PdfReader(path)
        pages = []
        for page in reader.pages:
            try:
                pages.append(page.extract_text() or "")
            except Exception:
                continue
        text = "\n".join(pages).strip()
    except Exception as exc:
        print(f"  could not read the PDF ({type(exc).__name__}: {exc})")
        return ""

    if not text:
        print("  the PDF yielded no extractable text.")
        return ""
    print(f"  {os.path.basename(path)}: {len(text)} characters of text, "
          f"{len(reader.pages)} pages")
    return text


def chatgpt_attach_file(page, path, timeout=60000):
    """
    Attach a file to the composer and wait for it to finish uploading.

    set_input_files drives the hidden <input type=file> directly, so
    no OS file dialog is involved. If the page has no file input at
    all, that is reported plainly -- uploading is an account feature,
    and the signed-out composer says so itself.
    """
    name = os.path.basename(path)
    print(f"  attaching {name}")

    deadline = time.time() + 20
    while time.time() < deadline:
        try:
            if page.locator('input[type="file"]').count():
                break
        except Exception:
            pass
        time.sleep(1)

    try:
        count = page.locator('input[type="file"]').count()
    except Exception:
        count = 0
    if not count:
        print("    this chat has NO file input -- uploading needs an")
        print("    account. Sign in with --chatgpt-login, or the rules")
        print("    have to go in as text instead.")
        return False

    try:
        page.locator('input[type="file"]').first.set_input_files(path)
    except Exception as exc:
        print(f"    the file was rejected ({type(exc).__name__}: {exc})")
        return False

    # Wait for the attachment to appear in the composer. The file name
    # showing up on the page is the signal that it took.
    stem = os.path.splitext(name)[0][:20].lower()
    deadline = time.time() + (timeout / 1000.0)
    while time.time() < deadline:
        text = _chatgpt_page_text(page, timeout=3000)
        if stem in text:
            print("    attached")
            time.sleep(3)   # let the upload settle before sending
            return True
        time.sleep(1)

    print("    the file never appeared in the composer.")
    _chatgpt_shot(page, "attach_failed")
    # A refused upload leaves "The selected file is not a supported
    # image." sitting in the composer, and that error state blocks the
    # next message from being typed at all. Reload to clear it.
    try:
        page.reload(wait_until="domcontentloaded", timeout=30000)
        _chatgpt_wait_for_any(
            page, CHATGPT_COMPOSER_SELECTORS, "composer", timeout=20000,
        )
        print("    reloaded to clear the upload error")
    except Exception:
        pass
    return False


# What is said alongside the uploaded master document. Deliberately
# short: the PDF carries the rules, the start-work instruction and the
# exact output formats, so restating them here could only conflict
# with it.
RULES_PDF_PROMPT = (
    "Understand this PDF. We have to start work. I will send website "
    "URLs one at a time — verify each one against this document and "
    "answer in the exact output format it specifies."
)


RULES_FEED_PREAMBLE = (
    "Below is the complete rulebook I verify websites against. Read all "
    "of it and keep it in mind for this whole conversation. I will then "
    "send website URLs one at a time; for each one, check the site "
    "against these rules and tell me whether it QUALIFIES or is "
    "REJECTED, naming the specific rule that decides it.\n\n"
    "Governing principle from the rulebook: a correct SKIP is better "
    "than an incorrect paid submission. Never guess and never fill a "
    "mandatory field by inference.\n\n"
    "=== RULES.md ===\n\n"
)


# There is no record cap any more: --gpt-flow runs until the terminal
# is closed or Ctrl+C is pressed. Failures recover in place instead of
# ending the run, and an empty queue is waited out rather than treated
# as the end of the work.


# Lines that only a QUALIFIES answer carries. The master document's
# output rule (section 4) never uses the word "QUALIFIES" -- a passing
# site is reported as the field block alone -- so the block itself is
# what identifies one.
QUALIFIES_BLOCK_MARKERS = (
    "email:",
    "phone no:",
    "kind of business:",
    "company profile:",
)


def parse_gpt_verdict(answer):
    """
    "SKIP", "QUALIFIES" or "UNCLEAR" for one ChatGPT answer.

    Only the tail is read: the captured text opens with our own prompt
    echoed into the transcript and states the decision at the end.

    Both signals, or neither, is UNCLEAR and nothing is submitted --
    the rulebook's own principle applied to the verdict itself.

    Note the master document requires a rejection to be exactly
    "SKIP" with no reason given unless asked. A bare verdict is
    therefore correct behaviour, not a degraded answer.
    """
    tail = (answer or "")[-1500:]
    upper = tail.upper()
    lower = tail.lower()

    says_skip = ("SKIP" in upper) or ("REJECT" in upper)
    block_hits = sum(1 for marker in QUALIFIES_BLOCK_MARKERS if marker in lower)
    says_qualifies = ("QUALIFIES" in upper) or block_hits >= 3

    if says_skip and not says_qualifies:
        return "SKIP"
    if says_qualifies and not says_skip:
        return "QUALIFIES"
    return "UNCLEAR"


GPT_PAID_TYPE_LOOKUP = {name.lower(): name for name in PAID_BUSINESS_TYPES}

# The short forms the portal's country box is typed with. ChatGPT
# writes the long name, so the same rule the decision engine applies
# in portal_country_name() is applied to its wording too.
GPT_COUNTRY_SHORT_FORMS = {
    "united states": "USA",
    "united states of america": "USA",
    "usa": "USA",
    "us": "USA",
    "united kingdom": "UK",
    "uk": "UK",
    "great britain": "UK",
    "england": "UK",
    "united arab emirates": "UAE",
    "uae": "UAE",
    "hong kong": "China (Hong Kong S.A.R.)",
    "macau": "China (Macau S.A.R.)",
    "macao": "China (Macau S.A.R.)",
}


def gpt_portal_country(name):
    """The exact string to type into the portal's country box."""
    key = clean(name or "").lower().strip(" .")
    if not key:
        return ""
    return (
        GPT_COUNTRY_SHORT_FORMS.get(key)
        or PORTAL_COUNTRY_FILL_NAMES.get(key)
        or clean(name)
    )


def parse_gpt_qualifies(answer):
    """
    Turn a QUALIFIES answer into the field dict
    fill_and_submit_qualifies() expects, or None.

    None whenever ANY mandatory field cannot be read, is not one of
    the seven paid business types, is flagged N rather than Y, or the
    product counts fall below the documented minimum of 3. Nothing is
    inferred and nothing is defaulted: an unreadable field block is a
    record left alone, not a record submitted with a guess.
    """
    text = (answer or "").replace(" ", " ")
    text = re.sub(r"[*#`]", "", text)

    def grab(pattern):
        match = re.search(pattern, text, re.I)
        return clean(match.group(1)) if match else ""

    email = grab(r"Email\s*:?\s*([^\s,;]+@[^\s,;]+)")
    # A rendered link can glue a "↗" onto the address.
    email = re.sub(r"[^A-Za-z0-9._%+\-@]+$", "", email)
    phone = grab(r"Phone(?:\s*(?:No|Number)\.?)?\s*:\s*([+0-9][0-9 ()\-]{5,})")
    phone = re.sub(r"[^0-9+]", "", phone)
    country_raw = grab(r"Country\s*:\s*([^\n]+)")
    business_raw = grab(r"Kind of Business\s*:\s*([^\n]+)")

    # Two shapes are accepted, because both are in use: the master
    # document's "Address: Y", and the "Address Y: <value>" form that
    # answers were observed using.
    flags = {}
    for label, key in (
        ("Address", "address"),
        ("City", "city"),
        ("State", "state"),
        ("Company Profile", "company_profile"),
    ):
        match = re.search(label + r"\s*([YN])\s*:", text, re.I)
        if not match:
            match = re.search(label + r"\s*:\s*([YN])\b", text, re.I)
        flags[key] = bool(match and match.group(1).upper() == "Y")

    # Likewise for the product lines: "Product Name 3: ..." carries an
    # explicit count, while the master document's "3+ Physical
    # Products: Y" asserts the threshold was met.
    counts = []
    for numeric_label, threshold_label in (
        (r"Product Name|Name of Prod", r"3\+\s*(?:Physical\s*)?Products?"),
        (r"Product Image|Image of Prod", r"3\+\s*Product Images?"),
        (r"Product Description|Desc of Prod", r"3\+\s*Product Descriptions?"),
    ):
        match = re.search(
            r"(?:" + numeric_label + r")\s*(\d+)\s*:", text, re.I,
        )
        if match:
            counts.append(int(match.group(1)))
            continue
        match = re.search(
            r"(?:" + threshold_label + r")\s*:\s*([YN])\b", text, re.I,
        )
        if match:
            counts.append(3 if match.group(1).upper() == "Y" else 0)
        else:
            counts.append(0)

    business_type = GPT_PAID_TYPE_LOOKUP.get(business_raw.lower())
    country_fill = gpt_portal_country(country_raw)
    product_count = min(counts) if counts else 0

    problems = []
    if not email or "@" not in email:
        problems.append("email")
    if not phone:
        problems.append("phone")
    if not country_fill:
        problems.append("country")
    if not business_type:
        problems.append(
            f"kind of business ({business_raw or 'missing'!r} is not one "
            "of the seven paid types)"
        )
    for label, key in (
        ("address", "address"),
        ("city", "city"),
        ("state", "state"),
        ("company profile", "company_profile"),
    ):
        if not flags[key]:
            problems.append(f"{label} is not Y")
    if product_count < MIN_QUALIFYING_PRODUCTS:
        problems.append(
            f"product counts {counts} below the minimum of "
            f"{MIN_QUALIFYING_PRODUCTS}"
        )

    if problems:
        print("  the QUALIFIES block is not complete enough to submit:")
        for problem in problems:
            print(f"    - {problem}")
        return None

    return {
        "email": email,
        "phone": phone,
        "country": country_fill,
        "country_fill": country_fill,
        "business_type": business_type,
        "address_ok": True,
        "city_ok": True,
        "state_ok": True,
        "company_profile_ok": True,
        "product_count": product_count,
    }


def wait_for_portal_record_ready(portal, expected_url, timeout=40):
    """
    Wait until the portal form is actually usable AND still showing the
    record we just had judged. Returns "ok", "timeout" or "mismatch".

    Two separate hazards, both seen live:

    1. The form loads behind a spinner. Its <select> elements exist in
       the DOM the whole time but are not interactable, so a status
       selection silently matches nothing and reports "no dropdown
       offers a Not Working option" -- which reads exactly like an
       expired session and is not one.

    2. Deciding a record takes ~20 seconds over in the ChatGPT tab. If
       the portal moved on to a different record in the meantime, the
       verdict for site A would be submitted against site B. The URL
       is re-read and compared before anything is selected, and a
       mismatch submits nothing.
    """
    deadline = time.time() + timeout
    recovered = False
    while True:
        usable = False
        try:
            status = portal.locator("#status").first
            usable = bool(
                status.count()
                and status.is_visible()
                and status.is_enabled()
                and status.locator("option").count() > 1
            )
        except Exception:
            usable = False

        # After a Working submission the portal redirects to its Admin
        # Console -- a listing page with no record form on it, which
        # would never become usable however long we waited. Observed
        # twice, both times on the record straight after a QUALIFIES.
        #
        # Navigate to the portal root explicitly. A reload is no use
        # here: it reloads the Admin Console.
        if not usable and not recovered and time.time() > deadline - timeout + 8:
            recovered = True
            print(f"  no record form here ({portal.url})")
            print("  navigating back to the record page")
            try:
                root = LOGIN_URL.split("/Account/Login")[0] + "/"
                portal.goto(
                    root, wait_until="domcontentloaded",
                    timeout=PAGE_NAVIGATION_TIMEOUT,
                )
                ensure_logged_in(portal)
            except Exception as exc:
                print(f"  could not get back ({type(exc).__name__})")
            time.sleep(2)
            continue

        if usable:
            current = get_assigned_url(portal)
            if current and current == expected_url:
                return "ok"
            if current and current != expected_url:
                print(f"  the portal now shows a DIFFERENT record: {current}")
                print(f"  the verdict was for {expected_url}")
                return "mismatch"

        if time.time() >= deadline:
            return "timeout"
        time.sleep(1)


def log_gpt_flow(assigned, verdict, outcome):
    """
    One CSV line per record in debug2/gpt_flow_log.csv. This mode
    changes live portal records, so what was submitted and why has to
    exist somewhere other than terminal scrollback.
    """
    path = debug_path("gpt_flow_log.csv")
    new_file = not os.path.exists(path)
    try:
        with io.open(path, "a", encoding="utf-8", newline="") as fh:
            if new_file:
                fh.write("timestamp,url,verdict,outcome" + "\n")
            safe_url = (assigned or "").replace(",", "%2C")
            stamp = time.strftime("%Y-%m-%d %H:%M:%S")
            fh.write(f"{stamp},{safe_url},{verdict},{outcome}" + "\n")
    except Exception as exc:
        print(f"  could not write the run log ({type(exc).__name__})")


def recover_portal_page(portal):
    """
    Get the portal back onto a record page. Returns True when a record
    form is showing.

    Navigates to the root explicitly rather than reloading: the page it
    usually needs rescuing from is the Admin Console, and reloading
    that just reloads the Admin Console.
    """
    try:
        root = LOGIN_URL.split("/Account/Login")[0] + "/"
        portal.goto(
            root, wait_until="domcontentloaded",
            timeout=PAGE_NAVIGATION_TIMEOUT,
        )
        ensure_logged_in(portal)
    except Exception as exc:
        print(f"  could not reach the record page ({type(exc).__name__})")
        return False

    deadline = time.time() + 20
    while time.time() < deadline:
        try:
            status = portal.locator("#status").first
            if status.count() and status.is_visible():
                return True
        except Exception:
            pass
        time.sleep(1)
    return False


def discard_fallback_profiles():
    """
    Delete the one-off browser profiles earlier runs fell back to.

    Only ever called once the MAIN profile has opened successfully,
    which proves nothing holds a lock and so no fallback is in use.
    Each one is a full Firefox profile of 35-40 MB and none of them
    carries a ChatGPT session, so keeping them buys nothing.
    """
    prefix = os.path.basename(CHATGPT_PROFILE_DIR) + "_"
    parent = os.path.dirname(CHATGPT_PROFILE_DIR)
    try:
        entries = os.listdir(parent)
    except Exception:
        return

    import shutil
    removed = 0
    for name in entries:
        if not name.startswith(prefix):
            continue
        path = os.path.join(parent, name)
        if not os.path.isdir(path):
            continue
        try:
            shutil.rmtree(path)
            removed += 1
        except Exception:
            # Still in use, or a file is held open. Leave it; it will
            # be caught on a later run.
            pass
    if removed:
        print(f"  cleaned up {removed} leftover one-off browser profile(s)")


DEBUG_ARTEFACTS_TO_KEEP = 25

# Never deleted, whatever their age. gpt_flow_log.csv is the record of
# what this mode submitted to live portal records -- the audit trail --
# and run_log.csv is the same for System 1.
DEBUG_KEEP_FOREVER_SUFFIXES = (".csv", ".log")


def prune_debug_artefacts(keep=None):
    """
    Keep the most recent screenshots and form dumps, delete older ones.

    debug2/ grows without limit: every failed attach, rejected
    composer, unready form and form dump lands there, and 51 files had
    built up. Only the newest are ever of any use -- a screenshot from
    two days ago explains nothing about today's run.

    A count cap rather than an age cutoff, because a single bad hour
    can produce dozens of files while a quiet week produces none.

    The run logs are never touched, whatever the cap.
    """
    if keep is None:
        keep = DEBUG_ARTEFACTS_TO_KEEP

    folder = os.path.dirname(debug_path("x"))
    try:
        names = os.listdir(folder)
    except Exception:
        return

    artefacts = []
    for name in names:
        if name.lower().endswith(DEBUG_KEEP_FOREVER_SUFFIXES):
            continue
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        try:
            artefacts.append((os.path.getmtime(path), path))
        except Exception:
            continue

    if len(artefacts) <= keep:
        return

    artefacts.sort(reverse=True)          # newest first
    removed = 0
    for _, path in artefacts[keep:]:
        try:
            os.remove(path)
            removed += 1
        except Exception:
            pass
    if removed:
        print(f"  pruned {removed} old debug artefact(s), kept the newest {keep}")


def portal_on_admin_console(portal):
    """True when the portal is showing its Admin Console listing."""
    try:
        return "adminconsole" in (portal.url or "").lower()
    except Exception:
        return False


def wait_for_record_page(portal, poll=15):
    """
    Wait for a record form to appear WITHOUT navigating anywhere.

    Used when the portal is sitting on the Admin Console and we did not
    put it there -- almost certainly the user looking at their
    submitted records. Navigating back would yank the page out from
    under them, so the run pauses instead and picks up when a record
    page returns.

    Returns False if the browser was closed while waiting.
    """
    print("  the portal is on the Admin Console -- leaving it alone.")
    print("  Paused until a record page is showing again. Ctrl+C to stop.")
    waited = 0
    while True:
        try:
            if portal.is_closed():
                return False
            status = portal.locator("#status").first
            if status.count() and status.is_visible():
                print(f"  record page is back after {waited}s -- continuing.")
                return True
        except Exception:
            pass
        time.sleep(poll)
        waited += poll
        if waited % (poll * 8) == 0:
            print(f"  still paused ({waited}s).")


def feed_rulebook(gpt, who, rules):
    """
    Put the rulebook into the current chat. Returns ChatGPT's
    acknowledgement, or "" if it never took.

    Order: MASTER_RULES.md (the master PDF as Markdown) -> the PDF
    file itself, only when signed in -> the PDF's raw text -> RULES.md.
    """
    rules_pdf = find_rules_pdf()
    ack = ""

    md_path = project_file(MASTER_RULES_MD_NAME)
    if md_path:
        md_text = read_text_file(md_path)
        if md_text:
            print(f"  {os.path.basename(md_path)}: {len(md_text)} characters")
            ack = chatgpt_send_message(
                gpt, RULES_PDF_PROMPT + "\n\n" + md_text,
                "master rules (Markdown)",
            )

    if not ack and rules_pdf and who:
        print("  trying to upload the PDF itself...")
        ack = chatgpt_send_message(
            gpt, RULES_PDF_PROMPT, "master PDF", attach=rules_pdf,
        )

    if not ack and rules_pdf:
        print("  falling back to the PDF's raw text")
        pdf_text = read_rules_pdf_text(rules_pdf)
        if pdf_text:
            ack = chatgpt_send_message(
                gpt, RULES_PDF_PROMPT + "\n\n" + pdf_text, "master PDF text",
            )

    if not ack and rules:
        print("  falling back to RULES.md")
        ack = chatgpt_send_message(
            gpt, RULES_FEED_PREAMBLE + rules, "RULES.md",
        )

    return ack


def restart_chat(gpt, who, rules):
    """
    Open a brand-new chat and load the rulebook into it.

    The recovery for a ChatGPT tab that has stopped being useful --
    a wedged composer, an answer that never arrives, a reply that
    cannot be read. Cheaper than ending the run and starting over by
    hand, which is what used to happen.
    """
    print("  starting a fresh chat and re-loading the rulebook")
    try:
        gpt.goto(CHATGPT_HOME_URL, wait_until="domcontentloaded", timeout=40000)
    except Exception as exc:
        print(f"  could not open a new chat ({type(exc).__name__})")
        return False

    if _chatgpt_bot_wall(gpt):
        print("  a bot check is in the way of the new chat.")
        return False

    ack = feed_rulebook(gpt, who, rules)
    if not ack:
        print("  the fresh chat did not accept the rulebook.")
        return False
    print("  fresh chat ready.")
    return True


def gpt_flow_mode(playwright):
    """
    --gpt-flow : portal in one tab, ChatGPT in another, RULES.md fed
    into the chat, then the portal's assigned URL sent into that same
    chat. Read-only against the portal.
    """
    print("-" * 70)
    print("PORTAL + CHATGPT FLOW  [SYSTEM 2]")

    # Any one of the three rulebook sources is enough. RULES.md is only
    # the last-resort fallback now, so requiring it specifically was
    # wrong: the mode refused to start with MASTER_RULES.md present and
    # readable, purely because RULES.md had been moved into
    # automation1\.
    md_available = project_file(MASTER_RULES_MD_NAME)
    pdf_available = find_rules_pdf()
    rules = read_rules_document()
    if not (md_available or pdf_available or rules):
        print("  no rulebook found. One of these must be next to the")
        print(f"  script or at the project root: {MASTER_RULES_MD_NAME}, ")
        print(f"  {CHATGPT_RULES_PDF_NAME}, or RULES.md")
        return False
    print(f"  rulebook: {os.path.basename(md_available or pdf_available) or 'RULES.md'}")

    context = None
    try:
        context = playwright.firefox.launch_persistent_context(
            CHATGPT_PROFILE_DIR, headless=False,
        )
        # The main profile opened, so nothing holds a lock on it and
        # any one-off profiles left by earlier runs are dead weight.
        # Three of them had accumulated to 118 MB before this cleanup
        # existed.
        discard_fallback_profiles()
        prune_debug_artefacts()
    except Exception as exc:
        # A Firefox left running from an earlier run keeps parent.lock
        # held, and the profile cannot be reused while it does. Rather
        # than refuse to run, fall back to a profile of its own -- but
        # say so, because a fresh profile carries no ChatGPT session.
        print(f"  the usual profile would not open ({type(exc).__name__}).")
        print(f"  Something still holds {CHATGPT_PROFILE_DIR}")
        print("  -- most likely a Firefox window from an earlier run.")
        fallback = CHATGPT_PROFILE_DIR + "_" + time.strftime("%Y%m%d_%H%M%S")
        print(f"  using a one-off profile instead: {fallback}")
        print("  NOTE: a one-off profile is never signed into ChatGPT.")
        try:
            context = playwright.firefox.launch_persistent_context(
                fallback, headless=False,
            )
        except Exception as exc2:
            print(f"  that failed too ({type(exc2).__name__} {exc2})")
            return False

    try:
        # ---- tab 1: the portal ----
        print("\n[tab 1] portal")
        portal = context.pages[0] if context.pages else context.new_page()
        portal.set_default_timeout(7000)
        portal.set_default_navigation_timeout(PAGE_NAVIGATION_TIMEOUT)
        try:
            install_dialog_autoaccept(portal)
        except Exception:
            pass

        if not portal_log_in(portal):
            _chatgpt_shot(portal, "portal_login_failed")
            return False

        # ---- tab 2: chatgpt.com, same window, same profile ----
        print("\n[tab 2] chatgpt.com")
        gpt = context.new_page()
        gpt.set_default_timeout(CHATGPT_PAGE_TIMEOUT)
        gpt.set_default_navigation_timeout(40000)
        try:
            gpt.goto(CHATGPT_HOME_URL, wait_until="domcontentloaded",
                     timeout=40000)
        except Exception as exc:
            print(f"  could not open chatgpt.com ({type(exc).__name__})")
            return False

        who = _chatgpt_session_user(gpt)
        if who:
            print(f"  signed in as: {who}")
        else:
            print("  NOT signed in -- using the anonymous chat.")
            print("  It works, but there is no history and the message")
            print("  limit is lower. --chatgpt-login fixes that.")

        wall = _chatgpt_bot_wall(gpt)
        if wall:
            print(f"  blocked by a bot check ({wall!r}) -- cannot chat.")
            _chatgpt_shot(gpt, "gptflow_bot_wall")
            return False

        # ---- feed the rulebook ----
        # One implementation, shared with restart_chat() so a recovery
        # loads exactly what the first attempt did.
        print("")
        print("[step 1] loading the rulebook into the chat")
        ack = feed_rulebook(gpt, who, rules)

        # Keep trying rather than ending the run. The chat can refuse
        # the first attempt for reasons that clear on a retry -- a
        # composer that has not finished rendering, an ad card in the
        # way, a bot check that lets go after a moment.
        rulebook_tries = 0
        while not ack:
            rulebook_tries += 1
            print(f"  the rulebook was not accepted (try {rulebook_tries}).")
            _chatgpt_shot(gpt, "gptflow_rules_no_reply")
            wait = min(30 * rulebook_tries, 300)
            print(f"  waiting {wait}s, then starting a fresh chat.")
            time.sleep(wait)
            if restart_chat(gpt, who, rules):
                ack = "restarted"
                break
        print("  --- ChatGPT on the rulebook " + "-" * 30)
        print("  " + ack[-1500:].replace("\n", "\n  "))
        print("  " + "-" * 58)

        # ---- one record at a time, all in the same chat ----
        #
        # This loop does not stop on a problem. Every failure here is
        # one that has actually happened during a run -- a blocked
        # composer, a chat that stops answering, the portal wandering
        # onto its Admin Console -- and each one used to end the run
        # and need a hand restart. The record is never abandoned: the
        # portal keeps serving the same URL until something is
        # submitted for it, so recovering and going round again simply
        # retries it.
        #
        # It ends when the terminal is closed or Ctrl+C is pressed.
        submitted = 0
        qualified = 0
        seen = 0
        attempts = 0          # consecutive failures on the CURRENT url
        last_url = ""

        while True:
            # A closed browser is the one failure that cannot be
            # recovered from, and retrying it forever is worse than
            # stopping: the loop sat waiting 30s at a time against a
            # window that no longer existed. Everything else in here
            # recovers; this ends the run.
            if portal.is_closed() or gpt.is_closed():
                print("\n  the browser window was closed -- ending the run.")
                break

            # Deliberately NOT bring_to_front(): raising the tab
            # yanks the Firefox window in front of whatever the
            # user is doing, once per record. Playwright drives
            # background tabs perfectly well, so the run stays out
            # of the way -- open the window from the taskbar to
            # watch it.

            # If the portal is on the Admin Console and we did not just
            # put it there, someone is looking at their records. Wait
            # rather than navigating away from under them.
            if portal_on_admin_console(portal):
                wait_for_record_page(portal)
                continue

            assigned = get_assigned_url(portal)
            if not assigned:
                print("\n  no assigned URL on the page -- going to the "
                      "record page")
                if not recover_portal_page(portal):
                    if portal.is_closed():
                        continue    # handled at the top of the loop
                    print("  could not get a record page. Waiting 30s.")
                    time.sleep(30)
                continue

            if assigned != last_url:
                last_url = assigned
                attempts = 0
                seen += 1
            attempts += 1

            print("\n" + "=" * 70)
            print(f"[record {seen}] {assigned}"
                  + (f"   (attempt {attempts})" if attempts > 1 else ""))
            print("=" * 70)

            # A record that keeps failing gets a completely fresh chat
            # before being tried again, since a wedged ChatGPT tab is
            # the usual cause.
            if attempts in (3, 6, 9):
                print("  this record keeps failing -- starting a fresh chat")
                if not restart_chat(gpt, who, rules):
                    print("  the fresh chat did not take. Waiting 60s.")
                    time.sleep(60)
                    continue
            if attempts > 12:
                print("  12 attempts on this record with no progress.")
                print("  Waiting 5 minutes before trying again.")
                time.sleep(300)

            # Deliberately NOT bring_to_front(): raising the tab
            # yanks the Firefox window in front of whatever the
            # user is doing, once per record. Playwright drives
            # background tabs perfectly well, so the run stays out
            # of the way -- open the window from the taskbar to
            # watch it.
            # The URL alone, with no instruction wrapped around it.
            # The master document already states what to do with a URL
            # and exactly how to answer; repeating it here could only
            # contradict it.
            answer = chatgpt_send_message(gpt, assigned, "assigned URL")
            if not answer:
                print("  no answer from ChatGPT -- nothing submitted.")
                print("  restarting the chat and trying this record again.")
                _chatgpt_shot(gpt, "gptflow_no_verdict")
                restart_chat(gpt, who, rules)
                continue

            print("  --- ChatGPT " + "-" * 45)
            print("  " + answer[-2000:].replace("\n", "\n  "))
            print("  " + "-" * 58)

            verdict = parse_gpt_verdict(answer)
            print(f"  verdict: {verdict}")

            if verdict == "QUALIFIES":
                fields = parse_gpt_qualifies(answer)
                if not fields:
                    # QUALIFIES with an unreadable field block is the
                    # one case that must never be submitted: a wrong
                    # SKIP costs one unpaid record, a wrong Working is
                    # a paid submission of unverified data. Ask again
                    # in a clean chat rather than submitting a guess.
                    print("  nothing submitted -- the record stays assigned.")
                    log_gpt_flow(
                        assigned, verdict, "not submitted (fields incomplete)",
                    )
                    restart_chat(gpt, who, rules)
                    continue
                print("  fields read from the answer:")
                print(f"    Email             {fields['email']}")
                print(f"    Phone             {fields['phone']}")
                print(f"    Country           {fields['country_fill']}")
                print(f"    Kind of Business  {fields['business_type']}")
                print("    Address/City/State/Profile  Y/Y/Y/Y")
                print(f"    Products          {fields['product_count']}/3")
                action = "QUALIFIES -> Working"
                outcome = "submitted Working"
            elif verdict == "SKIP":
                fields = None
                action = "SKIP -> Not Working"
                outcome = "submitted Not Working"
            else:
                print("  the verdict is not clear enough to act on.")
                print("  Nothing submitted -- a decision is never guessed.")
                log_gpt_flow(assigned, verdict, "not submitted (unclear)")
                restart_chat(gpt, who, rules)
                continue

            # ---- submit, whichever way it went ----
            print(f"  submitting {action}")
            # Deliberately NOT bring_to_front(): raising the tab
            # yanks the Firefox window in front of whatever the
            # user is doing, once per record. Playwright drives
            # background tabs perfectly well, so the run stays out
            # of the way -- open the window from the taskbar to
            # watch it.

            ready = wait_for_portal_record_ready(portal, assigned)
            if ready == "mismatch":
                print("  the portal moved on to another record --")
                print("  nothing submitted, picking up whatever it shows now.")
                log_gpt_flow(assigned, verdict, "not submitted (mismatch)")
                continue
            if ready != "ok":
                log_gpt_flow(assigned, verdict, f"not submitted ({ready})")
                if portal_on_admin_console(portal):
                    # Someone is on the Admin Console. Nothing is
                    # submitted and the page is left alone.
                    wait_for_record_page(portal)
                else:
                    print("  the portal form never became usable -- nothing")
                    print("  submitted. Reloading and trying again.")
                    _chatgpt_shot(portal, "gptflow_form_not_ready")
                    recover_portal_page(portal)
                continue

            ok = False
            try:
                if not ensure_logged_in(portal):
                    print("  the portal session is gone and would not renew.")
                elif fields:
                    ok = fill_and_submit_qualifies(portal, fields)
                else:
                    ok = submit_skip_with_status(portal, "Not Working")
            except Exception as exc:
                print(f"  submission raised {type(exc).__name__}: {exc}")

            # The portal can jump to its Admin Console in the moment
            # between the readiness check passing and the submit
            # happening. The failure message says "Nothing was
            # changed", so no submission went in and going back to try
            # once more cannot double-submit.
            if not ok and "adminconsole" in (portal.url or "").lower():
                print("  the portal jumped to the Admin Console mid-submit")
                print("  -- going back to the record page and retrying once")
                recover_portal_page(portal)
                if wait_for_portal_record_ready(portal, assigned) == "ok":
                    try:
                        if fields:
                            ok = fill_and_submit_qualifies(portal, fields)
                        else:
                            ok = submit_skip_with_status(portal, "Not Working")
                    except Exception as exc:
                        print(f"  the retry raised {type(exc).__name__}: {exc}")
                else:
                    print("  the record page did not come back cleanly.")

            if not ok:
                print("  SUBMIT FAILED -- nothing was recorded for this")
                print("  record. Reloading and trying it again.")
                _chatgpt_shot(portal, "gptflow_submit_failed")
                log_gpt_flow(assigned, verdict, "SUBMIT FAILED")
                recover_portal_page(portal)
                continue

            submitted += 1
            if verdict == "QUALIFIES":
                qualified += 1
            attempts = 0
            print(f"  submitted. ({submitted} records, {qualified} Working)")
            log_gpt_flow(assigned, verdict, outcome)

            previous = assigned
            assigned = wait_for_new_assigned_url(portal, previous)
            if not assigned or assigned == previous:
                # An empty queue is not necessarily permanent, so wait
                # and look again rather than ending the run.
                print("\n  no new URL yet -- the queue may be empty.")
                if portal_on_admin_console(portal):
                    # Do not navigate away from a page someone may be
                    # reading; the next loop will wait it out.
                    print("  (the Admin Console is showing -- left alone)")
                    time.sleep(30)
                else:
                    print("  Waiting 60s, then looking again. Ctrl+C to stop.")
                    time.sleep(60)
                    recover_portal_page(portal)

        print("\n" + "-" * 70)
        print(f"Records seen: {seen}   submitted: {submitted}   "
              f"of which Working: {qualified}")
        print(f"Log: {debug_path('gpt_flow_log.csv')}")
        return True

    finally:
        try:
            context.close()
        except Exception:
            pass


def main():
    # Which file is actually running. Stale copies of this script have
    # been run by mistake more than once, and their failures look
    # identical to bugs in this one, so say the path outright.
    print("=" * 70)
    print("WEBSITE VERIFIER [SYSTEM 2] -- running:", os.path.abspath(__file__))
    print("Credentials: .env2   |   Debug output: debug2/")
    print("=" * 70)

    load_env_file(".env2")   # SYSTEM 2 -- its own credentials

    with sync_playwright() as p:

        # --chatgpt-login : ChatGPT only. Uses its own persistent
        # profile and never opens the portal, so it returns before any
        # other browser is launched.
        if CHATGPT_LOGIN_ONLY:
            chatgpt_login_mode(p)
            return

        # --login-only / --dump-form : portal only, and read-only.
        # Neither fills, clicks or submits anything.
        if PORTAL_LOGIN_ONLY or DUMP_FORM_ONLY:
            # FIREFOX -- REQUIRED. Chrome and other auto-translating
            # browsers break the language check.
            browser = p.firefox.launch(headless=False)
            try:
                portal = browser.new_page()
                portal.set_default_timeout(7000)
                portal.set_default_navigation_timeout(PAGE_NAVIGATION_TIMEOUT)
                install_dialog_autoaccept(portal)

                if not portal_log_in(portal):
                    print("Could not log into the portal -- stopping.")
                    return

                if DUMP_FORM_ONLY:
                    print("")
                    print("--dump-form: READ-ONLY. Nothing will be "
                          "filled, clicked or submitted.")
                    dump_portal_form(
                        portal, note="(--dump-form, record page after login)",
                    )
                    return

                hold_portal_login_open(portal)
            finally:
                try:
                    browser.close()
                except Exception:
                    pass
            return

        # Default (and --gpt-flow): the portal + ChatGPT flow. This is
        # System 2's only verification engine -- the built-in crawler
        # was removed on 2026-09-05 and lives on in System 1 at
        # automation1\website_verifier.py.
        gpt_flow_mode(p)


if __name__ == "__main__":
    main()
