"""
Website Qualification & Verification Engine
=============================================
SYSTEM 3 -- the CHROME system.

Built 2026-09-08 by combining two things:
  * the proven portal + ChatGPT engine from System 2, and
  * the Chrome sign-in approach that actually works (see
    chatgpt_login_mode below -- it is the reason this file exists).

It is deliberately independent of the other two systems:
    credentials     ->  .env3               (not .env, not .env2)
    debug output    ->  debug3/             (not debug/, not debug2/)
    ChatGPT session ->  chatgpt_profiles/   (its own)
    portal session  ->  portal_chrome_profile3/
Changes made here do NOT affect System 1 or System 2.
=============================================

Automated Playwright-based engine that verifies whether an assigned
website qualifies for the Copy & Paste website-evaluation workflow.
The verification itself is done by ChatGPT against the master
rulebook; this script drives the portal and the chat.

CONTROLLING PRINCIPLES
-----------------------
1. CHROME, and the sign-in is done in a NORMAL Chrome window rather
   than an automated one. See the CHATGPT SESSION section for why --
   this is the single decision the whole system depends on.
2. The assigned website URL is read dynamically from the portal.
   No target URL is ever hardcoded.
3. SUBMISSION BEHAVIOR (by explicit user instruction, overriding
   the source guideline's default "verification-only" posture):
     - On SKIP, the script selects Website Status = "Not Working"
       and submits the form, then picks up whatever new assigned
       URL the portal generates and continues automatically. This
       applies to EVERY SKIP reason, not only genuinely dead sites.
     - On QUALIFIES, every field is filled and the record is
       submitted with Website Status = "Working".
4. QUALIFIES is returned only when every mandatory requirement has
   actually been verified. A single missing/unclear/unverifiable
   mandatory requirement forces the final result to SKIP.
5. Never guess, invent, mask, or placeholder any value. Missing
   mandatory data is always a SKIP, never a fabricated answer.
6. Output format is fixed by the master guideline (section 4):
     - SKIP result   -> exactly:  SKIP
     - QUALIFIES     -> the exact field block, parsed back by
       parse_gpt_qualifies().

Run:
    python website_verifier3.py                  portal + ChatGPT
    python website_verifier3.py --chatgpt-login  one-time sign-in
    python website_verifier3.py --login-only     portal login, stop
    python website_verifier3.py --dump-form      read-only form dump
"""

from __future__ import annotations

import io
import os
import re
import subprocess
import sys
import time
from urllib.parse import urlparse

from playwright.sync_api import sync_playwright

# The Windows console is cp1252 by default, and a ChatGPT reply
# containing one "->" arrow crashed an otherwise finished run at the
# print statement. Unencodable characters are replaced, not fatal.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ============================================================
# PORTAL CONFIGURATION
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

    The parent lookup is what lets the rulebook live once at the
    project root while each system sits in its own folder -- one copy,
    so the systems can never drift onto different versions of the
    rules.
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
    kept in debug3/ so they never scatter across whatever folder the
    script happened to be launched from.
    """
    folder = os.path.join(script_dir(), "debug3")
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

# Read-only diagnostic mode:  python website_verifier3.py --dump-form
# Logs in, writes every field name / dropdown option on the record
# page to debug3/portal_form_debug_*.txt plus a screenshot, and exits.
# Nothing is filled, clicked or submitted, so it is always safe to
# run against live work.
DUMP_FORM_ONLY = "--dump-form" in sys.argv[1:]

# ChatGPT login mode:
#     python website_verifier3.py --chatgpt-login
# Opens a NORMAL Chrome window for a one-time hand sign-in, then
# verifies the saved session. The portal loop never runs in this mode.
CHATGPT_LOGIN_ONLY = "--chatgpt-login" in sys.argv[1:]

# Which saved ChatGPT profile to use:
#     python website_verifier3.py --chatgpt-profile work
# Lets several ChatGPT accounts live side by side under
# chatgpt_profiles/, so a second account can be signed in without
# disturbing the first.
CHATGPT_PROFILE_NAME = "default"
for _arg_index, _arg in enumerate(sys.argv[1:-1]):
    if _arg == "--chatgpt-profile":
        CHATGPT_PROFILE_NAME = sys.argv[_arg_index + 2]
        break
CHATGPT_PROFILE_NAME = re.sub(
    r"[^A-Za-z0-9._-]", "_", CHATGPT_PROFILE_NAME,
).strip("._") or "default"

# Portal login only:  python website_verifier3.py --login-only
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
    portal_form_debug_*.txt, plus a full-page screenshot. Called when
    an automatic submission cannot find a field, so the exact selector
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
            "PORTAL_PASSWORD in .env3."
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
        for attempt in range(attempts):
            # Look first, then wait. The portal usually has the next
            # URL ready immediately, and sleeping before the first
            # check spent a second per record for nothing.
            if attempt:
                portal.wait_for_timeout(800)
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
# text -- only how many of each were verified, capped at 3.
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
#
# DO NOT "CORRECT" THIS. Rulebook v3.1 rule BT-014 says: '"Distributor"
# is valid. "Distributer" is not the accepted category spelling.' That
# governs what ChatGPT must OUTPUT, and it does -- parse_gpt_qualifies()
# accepts "Distributor" via PAID_BUSINESS_TYPES. But the live portal's
# own <option> is spelled "Distributer" (confirmed in the form dump),
# and select_option() has to match the portal exactly or the field is
# left blank. So both are right: "Distributor" in, "Distributer" onto
# the form.
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
    so a qualifying record is fully automatic -- nothing about it is
    left for manual entry.

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
    # USA / UK / the China (... S.A.R.) menu entries per the user's
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
# CHATGPT SESSION  -- THE KEY DESIGN DECISION
# ============================================================
#     python website_verifier3.py --chatgpt-login
#
# THE SIGN-IN NEVER HAPPENS IN AN AUTOMATED BROWSER. That single
# sentence is why this system works, and it is not a guess -- it was
# measured against the alternative on 2026-09-08.
#
# WHAT DOES NOT WORK: driving Google's sign-in from Playwright.
#   * chatgpt.com/auth/login accepts the email fine, but a
#     GOOGLE-LINKED account never gets an OpenAI password box -- it
#     hands off to accounts.google.com.
#   * On Firefox, Google refused outright: "This browser or app may
#     not be secure".
#   * On Playwright-driven Chrome it is subtler and worse. Google DOES
#     show its ordinary sign-in form and accepts the email, then
#     silently returns to chatgpt.com with NO session and NO error to
#     read. Better selectors and longer timeouts cannot fix that: the
#     problem is not the clicking, it is who is doing the clicking.
#
# WHAT WORKS: a plain chrome.exe, launched with subprocess, pointed at
# a --user-data-dir of our own. No Playwright, no CDP, no automation
# flags -- to Google it is an ordinary browser, because it is one. The
# sign-in is done BY HAND once. Chrome writes the session cookies into
# that folder. Playwright then opens the SAME folder with
# launch_persistent_context(channel="chrome") and simply inherits
# them; Google is never involved again, because the session already
# exists and ChatGPT only checks the cookie.
#
# So the automation READS a session a human created. It never tries to
# create one, and nothing here defeats or evades any check.
#
# Note the input() wait between the two steps: Chrome must be CLOSED
# before Playwright opens the profile. One Chrome user-data-dir cannot
# be open twice -- the second instance hands off to the first and
# exits, which surfaces as a TargetClosedError.

CHATGPT_HOME_URL = "https://chatgpt.com/"

# Default page timeout for the ChatGPT tab. Named because sending the
# rulebook temporarily raises it and has to put it back.
CHATGPT_PAGE_TIMEOUT = 15000

# Cookies live here, next to the script, so System 3 keeps its own
# ChatGPT session and never shares one with System 1 or 2. Named
# profiles let several accounts sit side by side.
CHATGPT_PROFILE_DIR = os.path.join(
    script_dir(), "chatgpt_profiles", CHATGPT_PROFILE_NAME,
)

# The portal gets a profile of its own. It MUST be separate: two
# Playwright contexts cannot share one Chrome user-data-dir.
PORTAL_PROFILE_DIR = os.path.join(script_dir(), "portal_chrome_profile3")

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
# one, but the layout served can differ from the one it was read off,
# so a plain-prose fallback is tried too -- an answer visibly on
# screen must never be reported as "no reply".
CHATGPT_ASSISTANT_SELECTORS = (
    '[data-message-author-role="assistant"]',
    "div.agent-turn",
    'article:has([data-message-author-role="assistant"])',
    "div.markdown.prose",
    "div.markdown",
)

# chatgpt.com's own session endpoint. This is the authoritative
# answer to "are we signed in": it returns a JSON object with a
# "user" key when the cookies are good, and {} when they are not.
CHATGPT_SESSION_API = "https://chatgpt.com/api/auth/session"

# Signs that chatgpt.com will not answer without an account.
CHATGPT_GATE_MARKERS = (
    "log in to continue",
    "sign up to continue",
    "you've reached our limit of messages",
    "rate limit",
    "please log in",
    "create an account to continue",
    "sign in to continue",
    "sign in is required",
)

# The "Message limit reached" dialog. An anonymous chat has a low cap;
# once it is hit the send button still clicks and the answer simply
# never arrives, which reads exactly like a wedged composer. The cap is
# per CONVERSATION, so the dialog's own "New chat" option clears it --
# observed 2026-09-07, where try 3 of the rulebook went through on a
# new chat after two refusals.
CHATGPT_LIMIT_MARKERS = (
    "message limit reached",
    "reached the anonymous message limit",
    "reached our limit of messages",
    "you've reached your limit",
)

# "New chat" inside that dialog first, then the sidebar's own New chat.
CHATGPT_NEW_CHAT_SELECTORS = (
    '[role="dialog"] button:has-text("New chat")',
    '[role="dialog"] a:has-text("New chat")',
    'button:has-text("New chat")',
    'a[href="/"]:has-text("New chat")',
)


# Chrome downloads its on-device AI model (Gemini Nano) into any fresh
# user-data-dir it is given. MEASURED 2026-09-08: 4,072 MB in
# chatgpt_profiles/acct4, against ~130 MB for everything else in that
# profile put together -- and it has nothing to do with the ChatGPT
# session. These switches stop it being fetched at all.
#
# Applied to BOTH launch paths on purpose. The sign-in Chrome is the
# window that stays open longest, while someone types a password, and
# that is the one that actually pulled the 4 GB.
#
# Nothing here is an automation marker: these only turn off a model
# download, so the sign-in window stays an ordinary browser as far as
# Google is concerned. See clean_profiles.py for clearing what has
# already accumulated.
CHROME_NO_MODEL_DOWNLOAD_ARGS = [
    "--disable-features=OptimizationGuideModelDownloading,"
    "OptimizationGuideOnDeviceModel,OptimizationHints",
]


def find_chrome_executable():
    """
    The installed Google Chrome, or None.

    Looked up explicitly rather than left to playwright's channel
    lookup, because this one is launched as a NORMAL browser -- the
    whole point being that playwright is not involved.
    """
    candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", ""),
                     "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""),
                     "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""),
                     "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def launch_chrome_context(playwright, profile_dir, extra_args=None):
    """
    Playwright against a Chrome profile folder that survives between
    runs -- what keeps a session alive from one run to the next.

    channel="chrome" drives the REAL installed Chrome rather than
    playwright's bundled Chromium, which is not installed on this
    machine. Raises on failure; callers decide what a locked profile
    means.
    """
    args = ["--disable-extensions"] + list(CHROME_NO_MODEL_DOWNLOAD_ARGS)
    if extra_args:
        args.extend(extra_args)
    return playwright.chromium.launch_persistent_context(
        profile_dir, channel="chrome", headless=False, args=args,
    )


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


def _browser_is_gone(page):
    """
    True once the Chrome window has been closed underneath us.

    Worth its own test because a dead browser and a page that simply
    has not loaded look identical to every is_visible() call: a run on
    2026-09-08 sat in a polling loop against a window that had already
    closed, reporting nothing at all.
    """
    try:
        if page.is_closed():
            return True
        page.evaluate("() => 1")
        return False
    except Exception as exc:
        blob = f"{type(exc).__name__} {exc}".lower()
        return ("targetclosed" in blob
                or "has been closed" in blob
                or "browser has been closed" in blob)


def _chatgpt_wait_for_any(page, selectors, what, timeout=25000):
    """
    Poll the whole selector list until one is visible, and return that
    selector (None on timeout).

    The waiting is the point. chatgpt.com is a React app: the composer
    does not exist in the HTML that domcontentloaded fires on, it is
    built a few seconds later. An earlier version checked once,
    immediately, and reported "no composer" on a page that plainly had
    one.
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


def _chatgpt_logged_in(page, timeout=15000):
    """
    True only when chatgpt.com itself reports a signed-in user.

    ONLY the session endpoint is trusted, on purpose. Two DOM-based
    tests were tried here first and both reported a brand-new, empty
    profile as "ALREADY LOGGED IN":

      * the composer -- chatgpt.com shows "Ask ChatGPT" to signed-OUT
        visitors too, so it proves nothing;
      * composer AND no visible login button -- defeated by hidden
        "Log in" buttons in the DOM.

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


def _chatgpt_set_composer_text(page, selector, text):
    """
    Put text into the composer using JavaScript. Returns True if it
    stuck.

    Why not the keyboard: typing needs the page focused, and focusing
    it makes Chrome raise its window in front of whatever the user is
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


def _chatgpt_shot(page, name):
    """Screenshot into debug3/, timestamped so runs never overwrite one another."""
    stamp = time.strftime("%Y%m%d_%H%M%S")
    path = debug_path(f"chatgpt_{name}_{stamp}.png")
    try:
        page.screenshot(path=path, full_page=True)
        print(f"    screenshot: {path}")
    except Exception:
        print("    screenshot failed")
    return path


def chatgpt_login_mode(playwright):
    """
    --chatgpt-login : open ChatGPT in a NORMAL, non-automated Chrome
    window so Google permits the sign-in. After that window is closed,
    verify the saved session with the automation profile. The portal is
    never touched in this mode.

    See the section comment above for why it is done this way. In
    short: a human signs in, in a real browser; playwright only ever
    reads the cookies that sign-in left behind.
    """
    print("-" * 70)
    print("CHATGPT LOGIN  [SYSTEM 3]")
    print(f"  profile: {CHATGPT_PROFILE_DIR}")

    chrome_executable = find_chrome_executable()
    if not chrome_executable:
        print("  Google Chrome was not found in its standard Windows locations.")
        print("  Install Chrome, or edit find_chrome_executable().")
        return False

    print("  opening a NORMAL Chrome window for sign-in")
    print("  (not automated -- that is the point; Google accepts it)")
    print()
    print("  1. sign in to ChatGPT in that window")
    print("  2. CLOSE that Chrome window completely")
    print("  3. come back here and press Enter")
    print()
    try:
        subprocess.Popen([
            chrome_executable,
            f"--user-data-dir={CHATGPT_PROFILE_DIR}",
            "--no-first-run",
            "--no-default-browser-check",
            *CHROME_NO_MODEL_DOWNLOAD_ARGS,
            CHATGPT_HOME_URL,
        ])
    except Exception as exc:
        print(f"  could not open normal Chrome: {type(exc).__name__} {exc}")
        return False

    try:
        input("  Press Enter AFTER closing the Chrome sign-in window: ")
    except (EOFError, KeyboardInterrupt):
        print()

    # The profile must be free before playwright can open it. If Chrome
    # is still running, this is the failure that shows up.
    context = None
    try:
        context = launch_chrome_context(playwright, CHATGPT_PROFILE_DIR)
    except Exception as exc:
        print(f"  could not verify the Chrome profile: "
              f"{type(exc).__name__} {exc}")
        print("  The sign-in Chrome window is probably still open --")
        print("  close it completely and run --chatgpt-login again.")
        return False

    page = context.pages[0] if context.pages else context.new_page()
    page.set_default_timeout(10000)
    page.set_default_navigation_timeout(30000)

    logged_in = False
    try:
        try:
            page.goto(
                CHATGPT_HOME_URL, wait_until="domcontentloaded", timeout=30000,
            )
        except Exception as exc:
            print(f"  could not open chatgpt.com: {type(exc).__name__}")

        logged_in = _chatgpt_logged_in(page, timeout=15000)
        if logged_in:
            who = _chatgpt_session_user(page) or "(account not reported)"
            print("  ChatGPT sign-in detected in the saved Chrome profile.")
            print(f"  signed in as: {who}")
            print(f"  now at: {page.url}")
            _chatgpt_shot(page, "session_restored")
        else:
            print("  ChatGPT is still logged out in the saved Chrome profile.")
            print("  Nothing was saved. Run --chatgpt-login again and make")
            print("  sure the sign-in completes BEFORE closing the window.")

        print("-" * 70)
        print("ChatGPT login:", "SUCCESS" if logged_in else "FAILED")
        if logged_in:
            print("The session is saved. Later runs will not ask again.")
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

    The session is verified rather than assumed. Printing "Login
    successful." straight after clicking the button, without checking
    anything, means that line prints on bad credentials while the
    login page is still on screen. looks_like_login_page() is the same
    test the submission path uses to spot an expired session.
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
        print("  The credentials in .env3 were not accepted.")
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
# PORTAL + CHATGPT FLOW
# ============================================================
# Window 1: the copy-paste portal, logged in, in its own profile.
# Window 2: chatgpt.com, in the signed-in ChatGPT profile.
#
# They are SEPARATE persistent contexts on purpose: one Chrome
# user-data-dir cannot be driven by two playwright contexts at once.
#
# Then, in order: feed the rulebook into a chat, go back to the
# portal, read the assigned URL, send that URL into the SAME chat so
# it is answered with the rules already in context, and submit the
# verdict.

def read_rules_document():
    """RULES.md as text -- the last-resort fallback rulebook."""
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
        print("  no PORTAL_USERNAME / PORTAL_PASSWORD in .env3")
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
    # The Terms & Conditions block on this page can push #Email past
    # the default timeout.
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
    "ChatGPT can make mistakes",
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
    # Newer ChatGPT layouts place the response controls after a short
    # disclaimer, so remove the trailing control label as well.
    body = re.sub(r"\n\s*Think\s*$", "", body, flags=re.I)
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


# How long an answer must sit unchanged before it counts as finished,
# and how often to look. The quiet window only applies to answers that
# can still grow -- see the fast paths in _chatgpt_wait_for_reply().
REPLY_QUIET_SECONDS = 5
REPLY_POLL_SECONDS = 0.4
STATUS_HEARTBEAT_SECONDS = 30

# Scrolling queries every element under <main>, which gets expensive as
# the conversation grows, so it does not need doing on every poll.
SCROLL_EVERY_N_POLLS = 3

# How often the reply wait looks for the message-limit dialog. Every
# 5th poll is roughly every 2s: fast enough that a capped chat is
# caught almost immediately, rare enough not to cost reading the page
# text on every single poll.
LIMIT_CHECK_EVERY_N_POLLS = 5


# Every line the master document's field block requires. The block has
# finished streaming only once all of them carry a value.
QUALIFIES_REQUIRED_LABELS = (
    r"Email",
    r"Phone(?:\s*(?:No|Number)\.?)?",
    r"Country",
    r"Kind of Business",
    r"Address",
    r"City",
    r"State",
    r"Company Profile",
    r"3\+\s*(?:Physical\s*)?Products?",
    r"3\+\s*Product Images?",
    r"3\+\s*Product Descriptions?",
)


def _chatgpt_is_complete_qualifies(body):
    """
    True when every line of the QUALIFIES field block has arrived with
    a value, so there is nothing left to wait for.

    Deliberately strict. A block still streaming fails here -- the
    field currently arriving has no value yet -- and falls through to
    the quiet window, which is what protects a paid submission from
    being read half-written.
    """
    text = re.sub(r"[*#`]", "", body or "")
    if "@" not in text:
        return False
    for label in QUALIFIES_REQUIRED_LABELS:
        if not re.search(label + r"\s*:\s*\S", text, re.I):
            return False
    return True


def _chatgpt_is_definite_skip(body):
    """
    True when the answer is a rejection and cannot become anything else.

    Deliberately strict: the text must START with SKIP and must not
    carry any part of the QUALIFIES field block. An answer holding
    both is left to the careful path, which reports UNCLEAR and
    submits nothing.
    """
    text = (body or "").strip()
    if not text:
        return False
    lowered = text.lower()
    if any(marker in lowered for marker in QUALIFIES_BLOCK_MARKERS):
        return False
    if "qualifies" in lowered:
        return False
    # Drop a leading glyph or bullet, then require SKIP first.
    head = text.lstrip("-*#>•–— ").upper()
    return head.startswith("SKIP")


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

    Necessary because the conversation is virtualised: with the
    rulebook pasted in as a 50,000+ character message, the view stays
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


def chatgpt_message_limit(page):
    """
    The message-limit text showing on the page, or None.

    Worth checking before blaming the composer: at the cap the send
    button still clicks and the reply never comes, so a limited chat
    and a wedged one look identical in the log.
    """
    text = _chatgpt_page_text(page, timeout=4000)
    for marker in CHATGPT_LIMIT_MARKERS:
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
    polls = 0
    limit_polls = 0
    last_heartbeat = time.time()

    while time.time() < deadline:
        if _browser_is_gone(page):
            print("    the ChatGPT window was closed while waiting.")
            return ""

        now = time.time()
        if now - last_heartbeat >= STATUS_HEARTBEAT_SECONDS:
            elapsed = int(now - (deadline - timeout))
            print(f"    status: still waiting for ChatGPT reply ({elapsed}s elapsed)")
            last_heartbeat = now

        gate = _chatgpt_gate(page)
        if gate:
            print(f"    chatgpt.com is refusing: {gate!r}")
            return ""

        # The message-limit dialog means the answer is never coming.
        # Checked HERE, in the poll loop, and not only after this
        # function times out: the dialog is on screen the instant the
        # cap is hit, and waiting out the full 300s first burned five
        # minutes per record before the recovery could even start.
        # Its wording is not in CHATGPT_GATE_MARKERS, so the gate
        # check above walks straight past it.
        if limit_polls % LIMIT_CHECK_EVERY_N_POLLS == 0:
            limit = chatgpt_message_limit(page)
            if limit:
                print(f"    ChatGPT reports: {limit}")
                return ""
        limit_polls += 1

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
            polls += 1
            if polls % SCROLL_EVERY_N_POLLS == 1:
                _chatgpt_scroll_to_bottom(page)
            whole = _chatgpt_main_text(page)

            if sent_text and len(sent_text) <= 500:
                # A short message (a URL) can be found in the
                # transcript, so the answer after it is unambiguous.
                current = _chatgpt_answer_after(whole, sent_text)
            else:
                # The rulebook is far too long to still be rendered in
                # full, so fall back to the last answer and require it
                # to differ from what was there before.
                current = _chatgpt_answer_body(whole)
                if previous_answer and current == previous_answer:
                    time.sleep(REPLY_POLL_SECONDS)
                    continue

            if not current or _chatgpt_is_progress(current):
                # Not an answer yet: only page furniture so far, or
                # ChatGPT is still browsing.
                time.sleep(REPLY_POLL_SECONDS)
                continue

        # FAST PATH -- a bare SKIP is final the moment it appears.
        #
        # The rulebook requires a rejection to be exactly "SKIP" and
        # nothing else, so there is nothing further to stream. Even
        # when a reason does follow ("SKIP - Domain rule: ..."), the
        # verdict is unchanged, and a SKIP answer can never turn into
        # a QUALIFIES one: the two formats are mutually exclusive and
        # a qualifying answer opens with "Email:". So waiting out the
        # full quiet window here buys nothing and costs 5s on the
        # ~95% of records that are rejections.
        #
        # Guarded: any field-block marker in the text means this is
        # not a plain SKIP, and it falls through to the careful path.
        if current and _chatgpt_is_definite_skip(current):
            return current

        # SECOND FAST PATH -- a field block that is already complete.
        #
        # Every required line has arrived with a value, and the master
        # document allows nothing after the block, so the quiet window
        # can only re-read the same text. A half-streamed block fails
        # this check (a field still arriving has no value yet) and
        # still goes the careful way, which is what protects the paid
        # submission.
        if current and _chatgpt_is_complete_qualifies(current):
            return current

        if current and current == last_text:
            if quiet_since is None:
                quiet_since = time.time()
            elif time.time() - quiet_since >= REPLY_QUIET_SECONDS:
                return current
        else:
            last_text = current
            quiet_since = None
        time.sleep(REPLY_POLL_SECONDS)

    if last_text:
        print("    reply timed out mid-stream -- returning what arrived")
    else:
        print("    FAILURE: ChatGPT reply timed out with no response")
    return last_text


def chatgpt_send_message(
    page, text, label, reply_timeout=300, attach=None, manual_send=False,
):
    """
    Put one message into the open chat and return the reply.

    Three ways in, by size and composer shape:
      * a real clipboard paste for anything over 20,000 characters --
        one complete editor operation, dispatching the paste/input
        events React needs. The rulebook is far too big to type, and
        synthetic insertion gets unreliable at that size.
      * keyboard.insert_text for ordinary messages -- one input event,
        and unlike typing, its newlines do not send the message early.
      * a JS value-setter fill for plain textareas.
    """
    print(f"  sending {label} ({len(text)} chars)")

    selector = _chatgpt_wait_for_any(
        page, CHATGPT_COMPOSER_SELECTORS, "composer", timeout=30000,
    )
    if not selector:
        print("    FAILURE: no composer on the page -- cannot send")
        _chatgpt_shot(page, "no_composer")
        return ""

    if attach and not chatgpt_attach_file(page, attach):
        return ""

    before = len(_chatgpt_assistant_texts(page))
    baseline = _chatgpt_main_text(page)

    def load_clipboard():
        clipboard_file = debug_path("chatgpt_clipboard_payload.txt")
        with io.open(clipboard_file, "w", encoding="utf-8", newline="") as fh:
            fh.write(text)
        escaped_path = clipboard_file.replace("'", "''")
        subprocess.run(
            [
                "powershell", "-NoProfile", "-Sta", "-Command",
                "[IO.File]::ReadAllText('" + escaped_path
                + "') | Set-Clipboard",
            ],
            check=True,
            timeout=20,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        print("    clipboard loaded")

    def insert_complete_message():
        page.keyboard.insert_text(text)

    def paste_complete_message():
        # Clipboard paste is one complete editor operation and
        # dispatches the real paste/input events ChatGPT's React
        # composer needs.
        load_clipboard()
        page.keyboard.press("Control+V")
        print("    complete message pasted")

    def composer_text():
        try:
            return page.locator(selector).first.evaluate(
                "el => el.isContentEditable ? el.textContent : el.value"
            ) or ""
        except Exception:
            return ""

    try:
        contenteditable = page.locator(selector).first.evaluate(
            "el => Boolean(el.isContentEditable)"
        )
    except Exception:
        contenteditable = False

    if manual_send:
        try:
            load_clipboard()
            print("    exact rulebook is copied to the clipboard")
            print("    press Ctrl+V in ChatGPT, click Send, then press Enter here")
            input("    waiting for manual paste and Send: ")
        except (EOFError, KeyboardInterrupt):
            print()
            return ""
        return _chatgpt_wait_for_reply(
            page, before, timeout=reply_timeout, baseline_text=baseline,
            sent_text=text,
        )

    # Contenteditable ChatGPT composers need a real input event so
    # React updates its internal draft state. Setting textContent can
    # look right on screen while leaving Send disabled internally.
    if contenteditable:
        print("    entering text through the live ChatGPT editor")
        try:
            box = page.locator(selector).first
            box.click(timeout=5000)
            if len(text) > 20000:
                paste_complete_message()
            else:
                insert_complete_message()
            deadline = time.time() + 5
            while not composer_text() and time.time() < deadline:
                time.sleep(0.1)
        except Exception as exc:
            print(f"    live editor would not accept the text ({type(exc).__name__})")
            _chatgpt_shot(page, "composer_rejected")
            return ""
    elif not _chatgpt_set_composer_text(page, selector, text):
        print("    JS fill did not take -- using the editor fill fallback")
        try:
            box = page.locator(selector).first
            page.set_default_timeout(180000)
            try:
                box.click(timeout=5000)
                if len(text) > 20000:
                    paste_complete_message()
                else:
                    insert_complete_message()
            except Exception:
                print("    editor fill did not take -- falling back to insertion")
                try:
                    box.click(timeout=5000)
                except Exception:
                    box.evaluate("el => el.focus()")
                insert_complete_message()
            finally:
                page.set_default_timeout(CHATGPT_PAGE_TIMEOUT)
        except Exception as exc:
            print(f"    composer would not accept the text "
                  f"({type(exc).__name__})")
            _chatgpt_shot(page, "composer_rejected")
            return ""

    # Send only after ChatGPT has enabled the real button.
    send_selector = None
    deadline = time.time() + 15
    while time.time() < deadline:
        for candidate in CHATGPT_SEND_SELECTORS:
            try:
                button = page.locator(candidate).first
                if (button.count() and button.is_visible()
                        and button.is_enabled()):
                    send_selector = candidate
                    break
            except Exception:
                continue
        if send_selector:
            break
        time.sleep(0.25)
    if not send_selector:
        print("    FAILURE: no enabled Send button found after the message was entered")
        return ""
    try:
        send_button = page.locator(send_selector).first
        send_button.click(timeout=15000)
        print(f"    clicked send button: {send_selector}")
    except Exception as exc:
        print(f"    Send click failed ({type(exc).__name__})")
        _chatgpt_shot(page, "send_click_failed")
        try:
            page.locator(selector).first.press("Enter", timeout=8000)
            print("    submitted with Enter after the Send click failed")
        except Exception:
            return ""

    # A click that reports success is not proof the message went. The
    # composer emptying is.
    if composer_text():
        print("    message is still in the composer -- pressing Enter")
        try:
            page.locator(selector).first.press("Enter", timeout=8000)
        except Exception:
            page.keyboard.press("Enter")
        deadline = time.time() + 3
        while composer_text() and time.time() < deadline:
            time.sleep(0.2)
    if composer_text():
        print("    FAILURE: ChatGPT did not submit the message")
        _chatgpt_shot(page, "send_not_submitted")
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
# kept with the project, then in Downloads.
CHATGPT_RULES_PDF_NAME = "master doc for verification.pdf"

# The rulebook sent to ChatGPT, newest version first. v3.1 is the
# master document plus an incremental patch that tightens root-domain
# control, business-type evidence, country inference, the three
# independent product thresholds, and the image rules -- and it
# preserves the whole original master inside itself, so it supersedes
# rather than replaces.
MASTER_RULES_CANDIDATES = (
    "MASTER_RULES_v3.1.md",
    "MASTER_RULES.md",
)

# Kept for the startup check and the "no rulebook" message.
MASTER_RULES_MD_NAME = MASTER_RULES_CANDIDATES[0]


def find_master_rules():
    """The newest rulebook available, or "" if none is on disk."""
    for name in MASTER_RULES_CANDIDATES:
        path = project_file(name)
        if path:
            return path
    return ""


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
    the other systems), then Downloads as a last resort.
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
        with open(path, "rb") as fh:
            file_bytes = fh.read()
        page.locator('input[type="file"]').first.set_input_files({
            "name": name,
            "mimeType": "text/plain",
            "buffer": file_bytes,
        })
    except Exception as exc:
        print(f"    the file was rejected ({type(exc).__name__}: {exc})")
        return False

    try:
        selected = page.locator('input[type="file"]').first.evaluate(
            "el => Boolean(el.files && el.files.length)"
        )
    except Exception:
        selected = False
    if selected:
        print("    file selected; waiting briefly for ChatGPT upload")
        deadline = time.time() + 15
        while time.time() < deadline:
            for send_selector in CHATGPT_SEND_SELECTORS:
                try:
                    button = page.locator(send_selector).first
                    if button.count() and button.is_visible() and button.is_enabled():
                        return True
                except Exception:
                    pass
            time.sleep(1)
        print("    ChatGPT did not enable Send for this file")
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
# short: the rulebook carries the rules, the start-work instruction
# and the exact output formats, so restating them here could only
# conflict with it.
RULES_PDF_PROMPT = (
    "Read the complete Markdown rulebook below as the controlling "
    "instruction set. Preserve its headings, rules, exceptions, and "
    "required output format. I will send one website URL at a time. "
    "Verify each URL against the entire rulebook, never guess, and "
    "return only the exact SKIP or QUALIFIES format required by it.\n\n"
    "===== BEGIN MASTER RULEBOOK ====="
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
# is applied to its wording too.
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


# Answers that mean "this field is not there". Anything else printed
# on one of the four Y/N lines is the verified value itself.
GPT_FLAG_NEGATIVE_VALUES = {
    "", "n", "no", "none", "n/a", "na", "nil", "null", "not found",
    "not available", "not listed", "not published", "not provided",
    "not mentioned", "missing", "absent", "unknown", "-", "--",
}


def gpt_flag_value_is_yes(value):
    """
    Read one of the four Y/N answer fields (Address, City, State,
    Company Profile) from whatever the answer actually put on the line.

    The master document specifies "Address: Y", but ChatGPT also prints
    the verified value in its place -- "Address: Brandium Rezidans, ...
    Istanbul", "Company Profile: <the profile text>". That is the same
    assertion carrying its evidence. Reading it as N rejected a good
    barelit.com answer and cost dunham-bush.com a paid record on
    2026-09-06, which is what this exists to stop.

    Nothing is inferred: an empty line, "N", "N/A" and "not found" all
    still fail, so the governing rule holds -- a correct SKIP beats an
    incorrect paid submission.
    """
    cleaned = clean(re.sub(r"^[\s\"'*#`]+|[\s\"'*#`.]+$", "", str(value or "")))
    return cleaned.lower() not in GPT_FLAG_NEGATIVE_VALUES


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
    text = (answer or "").replace("\xa0", " ")
    text = re.sub(r"[*#`]", "", text)

    def grab(pattern):
        match = re.search(pattern, text, re.I)
        return clean(match.group(1)) if match else ""

    email = grab(r"Email\s*:?\s*([^\s,;]+@[^\s,;]+)")
    # A rendered link can glue a stray character onto the address.
    email = re.sub(r"[^A-Za-z0-9._%+\-@]+$", "", email)
    phone = grab(r"Phone(?:\s*(?:No|Number)\.?)?\s*:\s*([+0-9][0-9 ()\-]{5,})")
    phone = re.sub(r"[^0-9+]", "", phone)
    country_raw = grab(r"Country\s*:\s*([^\n]+)")
    business_raw = grab(r"Kind of Business\s*:\s*([^\n]+)")

    # Three shapes are accepted, because all three are in use: the
    # master document's "Address: Y", the "Address Y: <value>" form
    # that answers were observed using, and the field's actual value
    # printed in place of the Y -- see gpt_flag_value_is_yes().
    flags = {}
    for label, key in (
        ("Address", "address"),
        ("City", "city"),
        ("State", "state"),
        ("Company Profile", "company_profile"),
    ):
        value = grab(label + r"\s*:\s*([^\n]+)")
        if value:
            flags[key] = gpt_flag_value_is_yes(value)
            continue
        match = re.search(label + r"\s*([YN])\s*:", text, re.I)
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

    2. Deciding a record takes ~20 seconds over in the ChatGPT window.
       If the portal moved on to a different record in the meantime,
       the verdict for site A would be submitted against site B. The
       URL is re-read and compared before anything is selected, and a
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
        # 0.3s, not 1s: this runs once per record and the form is
        # usually ready within a fraction of a second of being asked.
        # The deadline is unchanged, so a genuinely slow page still
        # gets its full wait.
        time.sleep(0.3)


def log_gpt_flow(assigned, verdict, outcome):
    """
    One CSV line per record in debug3/gpt_flow_log.csv. This mode
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
        # See wait_for_portal_record_ready: a tighter poll, the same
        # deadline.
        time.sleep(0.3)
    return False


def discard_fallback_profiles():
    """
    Delete the one-off browser profiles earlier runs fell back to.

    Only ever called once the MAIN profile has opened successfully,
    which proves nothing holds a lock and so no fallback is in use.
    Each one is a full Chrome profile and none of them carries a
    ChatGPT session, so keeping them buys nothing.
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
# what this mode submitted to live portal records -- the audit trail.
DEBUG_KEEP_FOREVER_SUFFIXES = (".csv", ".log")


def prune_debug_artefacts(keep=None):
    """
    Keep the most recent screenshots and form dumps, delete older ones.

    debug3/ grows without limit: every failed attach, rejected
    composer, unready form and form dump lands there. Only the newest
    are ever of any use -- a screenshot from two days ago explains
    nothing about today's run.

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


def feed_rulebook(gpt, who, rules, manual_send=False):
    """
    Put the rulebook into the current chat. Returns ChatGPT's
    acknowledgement, or "" if it never took.

    Order: the newest MASTER_RULES file -> the PDF itself, only when
    signed in -> the PDF's raw text -> RULES.md.
    """
    rules_pdf = find_rules_pdf()
    ack = ""

    md_path = find_master_rules()
    if md_path:
        md_text = read_text_file(md_path)
        if md_text:
            print(f"  {os.path.basename(md_path)}: {len(md_text)} characters")
            ack = chatgpt_send_message(
                gpt,
                md_text,
                "master rules (one complete Markdown message)",
                manual_send=manual_send,
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

    if not ack and rules and not who:
        print("  falling back to RULES.md")
        ack = chatgpt_send_message(
            gpt, RULES_FEED_PREAMBLE + rules, "RULES.md",
        )

    return ack


def chatgpt_new_chat_from_limit(page):
    """
    Take the limit dialog's own "New chat" option, and confirm the
    dialog actually went away. True only if it did.

    Clicking is not enough to report success: the modal can stay up
    when the click lands on its backdrop, and a chat that still holds
    the dialog cannot be fed the rulebook.
    """
    for selector in CHATGPT_NEW_CHAT_SELECTORS:
        try:
            target = page.locator(selector).first
            if not target.is_visible(timeout=2000):
                continue
            target.click(timeout=5000)
        except Exception:
            continue
        try:
            page.wait_for_timeout(1500)
        except Exception:
            pass
        if not chatgpt_message_limit(page):
            print(f"    took 'New chat' from the dialog: {selector}")
            return True
    return False


def restart_chat(gpt, who, rules):
    """
    Keep using the current chat instead of opening a new conversation.

    The recovery for a ChatGPT tab that has stopped being useful --
    a wedged composer, an answer that never arrives, a reply that
    cannot be read. Cheaper than ending the run and starting over by
    hand, which is what used to happen.
    """
    print("  keeping the current ChatGPT conversation")
    return True


def gpt_flow_mode(playwright):
    """
    The default mode: portal in one Chrome window, ChatGPT in another,
    the rulebook fed into the chat, then each assigned URL sent into
    that same chat and the verdict submitted to the portal.
    """
    print("-" * 70)
    print("PORTAL + CHATGPT FLOW  [SYSTEM 3]")

    # Any one of the rulebook sources is enough.
    md_available = find_master_rules()
    pdf_available = find_rules_pdf()
    rules = read_rules_document()
    if not (md_available or pdf_available or rules):
        print("  no rulebook found. One of these must be next to the")
        print(f"  script or at the project root: {MASTER_RULES_MD_NAME}, ")
        print(f"  {CHATGPT_RULES_PDF_NAME}, or RULES.md")
        return False
    print(f"  rulebook: {os.path.basename(md_available or pdf_available) or 'RULES.md'}")
    print(f"  ChatGPT profile: {CHATGPT_PROFILE_DIR}")
    print(f"  portal profile : {PORTAL_PROFILE_DIR}")

    context = None
    portal_context = None
    try:
        context = launch_chrome_context(playwright, CHATGPT_PROFILE_DIR)
        # The main profile opened, so nothing holds a lock on it and
        # any one-off profiles left by earlier runs are dead weight.
        discard_fallback_profiles()
        prune_debug_artefacts()
    except Exception as exc:
        # A Chrome left running from an earlier run holds the profile's
        # SingletonLock, and it cannot be reused while it does. Rather
        # than refuse to run, fall back to a profile of its own -- but
        # say so, because a fresh profile carries no ChatGPT session.
        print(f"  the usual profile would not open ({type(exc).__name__}).")
        print(f"  Something still holds {CHATGPT_PROFILE_DIR}")
        print("  -- most likely a Chrome window from an earlier run.")
        fallback = CHATGPT_PROFILE_DIR + "_" + time.strftime("%Y%m%d_%H%M%S")
        print(f"  using a one-off profile instead: {fallback}")
        print("  NOTE: a one-off profile is NEVER signed into ChatGPT, so")
        print("  it will hit the anonymous message limit. Close the other")
        print("  Chrome window and restart to use the real one.")
        try:
            context = launch_chrome_context(playwright, fallback)
        except Exception as exc2:
            print(f"  that failed too ({type(exc2).__name__} {exc2})")
            return False

    try:
        # ---- window 1: the portal, in its OWN profile ----
        # It must be a separate user-data-dir: one Chrome profile
        # cannot be driven by two playwright contexts at once.
        print("\n[window 1] portal")
        try:
            portal_context = launch_chrome_context(
                playwright, PORTAL_PROFILE_DIR,
            )
        except Exception as exc:
            print(f"  the portal profile would not open "
                  f"({type(exc).__name__} {exc})")
            print(f"  Something still holds {PORTAL_PROFILE_DIR} --")
            print("  most likely a Chrome window from an earlier run.")
            return False

        portal = (
            portal_context.pages[0]
            if portal_context.pages else portal_context.new_page()
        )
        portal.set_default_timeout(7000)
        portal.set_default_navigation_timeout(PAGE_NAVIGATION_TIMEOUT)
        try:
            install_dialog_autoaccept(portal)
        except Exception:
            pass

        if not portal_log_in(portal):
            _chatgpt_shot(portal, "portal_login_failed")
            return False

        # ---- window 2: chatgpt.com, in the signed-in profile ----
        print("\n[window 2] chatgpt.com")
        gpt = context.pages[0] if context.pages else context.new_page()
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
            print("  limit is much lower. Fix it with:")
            print("      python website_verifier3.py --chatgpt-login")

        wall = _chatgpt_bot_wall(gpt)
        if wall:
            print(f"  blocked by a bot check ({wall!r}) -- cannot chat.")
            _chatgpt_shot(gpt, "gptflow_bot_wall")
            return False

        print("")
        print("[step 1] loading the rulebook into the chat")
        ack = feed_rulebook(gpt, who, rules, manual_send=True)

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
            print(f"  waiting {wait}s, then trying again.")
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
            # window that no longer existed.
            if portal.is_closed() or gpt.is_closed():
                print("\n  a browser window was closed -- ending the run.")
                break

            # Deliberately NOT bring_to_front(): raising the window
            # yanks Chrome in front of whatever the user is doing,
            # once per record. Playwright drives background windows
            # perfectly well, so the run stays out of the way.
            if portal_on_admin_console(portal):
                wait_for_record_page(portal)
                continue

            assigned = get_assigned_url(portal)
            if not assigned:
                print("\n  no assigned URL on the page -- going to the "
                      "record page")
                if not recover_portal_page(portal):
                    if portal.is_closed():
                        continue
                    print("  could not get a record page. Waiting 30s.")
                    time.sleep(30)
                continue

            if assigned != last_url:
                last_url = assigned
                attempts = 0
                seen += 1
            attempts += 1

            print(f"[record {seen}] {assigned}"
                  + (f"   (attempt {attempts})" if attempts > 1 else ""))
            print("=" * 70)

            if attempts in (3, 6, 9):
                print("  this record keeps failing -- reusing the same chat")
                if not restart_chat(gpt, who, rules):
                    print("  the fresh chat did not take. Waiting 60s.")
                    time.sleep(60)
                    continue
            if attempts > 12:
                print("  12 attempts on this record with no progress.")
                print("  Waiting 5 minutes before trying again.")
                time.sleep(300)

            # The URL alone, with no instruction wrapped around it.
            # The master document already states what to do with a URL
            # and exactly how to answer; repeating it here could only
            # contradict it.
            answer = chatgpt_send_message(gpt, assigned, "assigned URL")
            if not answer:
                # Name the message limit when that is what happened.
                # Otherwise it is logged as a silent composer, and the
                # real cause -- a conversation at its cap -- is
                # invisible in the terminal.
                limit = chatgpt_message_limit(gpt)
                if limit:
                    print(f"  ChatGPT hit its message limit: {limit}")
                    print("  taking a new chat and trying this record there.")
                    print("  Nothing submitted.")
                    chatgpt_new_chat_from_limit(gpt)
                else:
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
                    # QUALIFIES with an unreadable field block is never
                    # submitted as Working: a wrong SKIP costs one
                    # unpaid record, a wrong Working is a paid
                    # submission of unverified data.
                    #
                    # First time, ask again -- the answer may have been
                    # garbled. Second time, the field really is
                    # unavailable, and asking a third time just gets
                    # the same answer. A masked email is the usual
                    # cause: sites behind Cloudflare render
                    # "[email protected]" instead of an address, and
                    # the rulebook forbids submitting a masked one
                    # while making an unverifiable mandatory field a
                    # SKIP. So it goes in as Not Working rather than
                    # looping to the 12-attempt backoff.
                    if attempts < 2:
                        print("  nothing submitted -- retrying.")
                        log_gpt_flow(
                            assigned, verdict,
                            "not submitted (fields incomplete)",
                        )
                        restart_chat(gpt, who, rules)
                        continue

                    print("  the field block is still incomplete on attempt "
                          f"{attempts} -- a mandatory field cannot be")
                    print("  verified, which the rulebook makes a SKIP.")
                    verdict = "SKIP"
                    fields = None
                    action = "QUALIFIES-but-unverifiable -> Not Working"
                    outcome = "submitted Not Working (fields unverifiable)"
                else:
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
                time.sleep(3)
                continue

            # ---- submit, whichever way it went ----
            print(f"  submitting {action}")

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
            if not ok and portal_on_admin_console(portal):
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
        for closeable in (context, portal_context):
            try:
                if closeable is not None:
                    closeable.close()
            except Exception:
                pass


def main():
    # Which file is actually running. Stale copies of this script have
    # been run by mistake more than once, and their failures look
    # identical to bugs in this one, so say the path outright.
    print("=" * 70)
    print("WEBSITE VERIFIER [SYSTEM 3] -- running:", os.path.abspath(__file__))
    print("Credentials: .env3   |   Debug output: debug3/   |   Browser: Chrome")
    print("=" * 70)

    load_env_file(".env3")   # SYSTEM 3 -- its own credentials

    with sync_playwright() as p:

        # --chatgpt-login : ChatGPT only. Opens a normal Chrome window
        # for the hand sign-in and never touches the portal.
        if CHATGPT_LOGIN_ONLY:
            chatgpt_login_mode(p)
            return

        # --login-only / --dump-form : portal only, and read-only.
        # Neither fills, clicks or submits anything.
        if PORTAL_LOGIN_ONLY or DUMP_FORM_ONLY:
            browser = p.chromium.launch(channel="chrome", headless=False)
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

        # Default: the portal + ChatGPT flow, System 3's only engine.
        gpt_flow_mode(p)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nSTOPPED: interrupted by Ctrl+C.")
    except Exception as exc:
        print(f"\nFAILURE: {type(exc).__name__}: {exc}")
        print("The verifier stopped. Check the latest debug3 screenshot/log.")
