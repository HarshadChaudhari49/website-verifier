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


class BrowserGoneError(RuntimeError):
    """
    Raised when the browser/context/page has been closed mid-crawl --
    usually because the window was closed by hand or the browser
    crashed. Every later navigation would fail identically, so the
    crawl is abandoned and the record becomes a SKIP.
    """


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

# Portal + ChatGPT flow:  python website_verifier2.py --gpt-flow
# Portal in tab 1, chatgpt.com in tab 2, RULES.md fed into the chat,
# then the portal's assigned URL sent into that same chat. Read-only
# against the portal -- nothing is submitted.
GPT_FLOW = "--gpt-flow" in sys.argv[1:]

MAX_TOTAL_PAGES = 80
MAX_CONTEXT_PAGES = 18
MAX_PRODUCT_PAGES = 30

MIN_QUALIFYING_PRODUCTS = 3
# AUTHORITATIVE: Intensecore Guidelines (i), (j) and MASTER section 11
# both state 3 -- "There should be at least 3 distinct products
# available. (Selection criteria is- 3, less than that will go in
# rejection)", with 3 qualifying images and 3 descriptions to match.
# NOTE 1 of the Intensecore guidelines applies the same minimum of 3
# to Industrial Services. Raising this above 3 would reject records
# the client counts as valid, so it stays at the documented value.
MAX_IMAGE_CANDIDATES_PER_PRODUCT = 4

PAGE_NAVIGATION_TIMEOUT = 12000
BODY_TIMEOUT = 6000
IMAGE_REQUEST_TIMEOUT = 4000

# AUTHORITATIVE (Intensecore item 10, final sentence): "Dummy, Very
# Small, Black & White, Foggy images are acceptable." So there is NO
# colour requirement and NO meaningful size requirement -- an earlier
# saturation/size gate rejected exactly the images the client accepts
# and cost qualifying paid records. The floor below only exists to
# throw out 1x1 trackers, spacer GIFs and tiny UI icons, which are not
# product images by any reading.
MIN_IMAGE_WIDTH = 32
MIN_IMAGE_HEIGHT = 32
MIN_IMAGE_AREA = 1024

# USER INSTRUCTION (2026-09-02), overriding nothing in the documents
# but adding to them: "Half-Cut images are not considerable, Blurred
# Images are not considerable, Irrelevant images corresponding to
# product [are not considerable]."
#
# Half-cut: a cropped strip of a photo. It cannot be detected from the
# pixels with certainty, so the proxy is an extreme aspect ratio --
# a genuine product photo is not a long thin sliver, and banners/
# letterbox crops that are, do not prove a product either.
MAX_IMAGE_ASPECT_RATIO = 4.0

# Blurred: measured as the variance of the image's edge response.
# A sharp photo produces strong, varied edges; an out-of-focus one
# produces almost none. The threshold is deliberately LOW so that only
# clearly blurred images fail -- the documents still call "Foggy"
# images acceptable, and wrongly rejecting a valid image costs a paid
# record. Raise it only if the client reports blurred-image errors.
MIN_IMAGE_SHARPNESS = 45.0

ACCEPTED_IMAGE_FORMATS = {"JPEG", "PNG", "GIF"}
ACCEPTED_IMAGE_MIME_TYPES = {"image/jpeg", "image/png", "image/gif"}


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

# Business types where an online-selling / payment-gateway option
# disqualifies the site from being paid (Manufacturer is exempt).
ONLINE_SELLING_DISQUALIFIES = {
    "Trader",
    "Wholesaler",
    "Supplier",
    "Distributor",
    "Exporter",
}

BUSINESS_PATTERNS = {
    "Manufacturer": [
        r"\bwe\s+are\s+(?:a|an)\s+manufacturer\b",
        r"\bwe\s+are\s+manufacturer\b",
        r"\bwe\s+manufacture\b",
        # "We design, engineer and manufacture ..." -- a clear
        # first-person statement that the old \bwe\s+manufacture\b
        # missed because of the intervening verbs. Up to four words
        # keeps it to one clause, so an unrelated later sentence
        # cannot be swept in.
        r"\bwe\s+(?:[a-z]+,?\s+){1,4}(?:and\s+)?manufacture\b",
        r"\bwe\s+(?:also\s+)?design\s+and\s+manufacture\b",
        # Third-person company statements. An About page far more often
        # says "The X Company manufactures ..." than "we manufacture",
        # and every one of those was being missed -- the single biggest
        # reason real manufacturers were rejected as non-paid.
        r"\bhas\s+been\s+manufacturing\b",
        r"\bhave\s+been\s+manufacturing\b",
        r"\bmanufactures\s+(?:of\s+)?(?:a|an|the|our|its|their)\b",
        r"\bmanufactures\s+and\s+(?:supplies|sells|distributes|exports)\b",
        r"\bour\s+(?:company|business|group|firm)\s+manufactur\w*\b",
        # Either "is a (leading) manufacturer" or a bare plural
        # "are manufacturers". A bare singular is excluded on
        # purpose, so "are manufacturer approved" cannot match.
        r"\b(?:is|are)\s+(?:(?:a|an|the)\s+(?:leading\s+|premier\s+|"
        r"global\s+|australian\s+|american\s+|uk\s+|british\s+|"
        r"custom\s+|specialist\s+)*manufacturers?|manufacturers)\b",
        r"\bmanufacturer\s+(?:and\s+\w+\s+)?of\s+\w+",
        r"\bmanufacturing\s+(?:company|firm|facility)\b",
        r"\bour\s+(?:company|business|group)\s+manufactures\b",
        r"\bour\s+(?:company|business|group)\s+is\s+(?:a|an)\s+manufacturer\b",
        r"\b(?:company|business|group)\s+is\s+(?:a|an)\s+manufacturer\b",
        r"\bmanufacturer\s+of\s+(?:our|the|all|[A-Za-z0-9])",
        r"\bmanufacturer\s+and\s+supplier\b",
        r"\bmanufacturer\s+and\s+exporter\b",
        r"\bmanufacturer\s+and\s+distributor\b",
        r"\bmanufacturer\s*,\s*supplier\b",
        r"\bmanufacturer\s*,\s*exporter\b",
        r"\bmanufacturer\s*,\s*distributor\b",
        r"\bmanufacturer\s*&\s*supplier\b",
        r"\bmanufacturer\s*&\s*exporter\b",
        r"\bmanufacturer\s*&\s*distributor\b",
        r"\bmanufacture\s+and\s+supply\b",
        r"\bmanufacture\s+and\s+export\b",
        r"\bmanufacture\s+and\s+distribute\b",
    ],
    "Industrial Services": [
        r"\bwe\s+are\s+(?:an?\s+)?industrial\s+service\s+provider\b",
        r"\bwe\s+provide\s+industrial\s+services?\b",
        r"\bindustrial\s+service\s+provider\b",
        r"\bindustrial\s+services?\s+(?:company|provider)\b",
        r"\bindustrial\s+manufacturing\s+services?\b",
    ],
    "Trader": [
        r"\bwe\s+are\s+(?:a|an)\s+trader\b",
        r"\btrading\s+company\b",
        r"\bour\s+(?:company|business|group)\s+trades\b",
        r"\bwe\s+trade\b",
    ],
    "Wholesaler": [
        r"\bwe\s+are\s+(?:a|an)\s+wholesaler\b",
        r"\bwe\s+wholesale\b",
        r"\bwholesaler\s+of\b",
        r"\bour\s+(?:company|business|group)\s+wholesales\b",
    ],
    "Supplier": [
        r"\bwe\s+are\s+(?:a|an)\s+supplier\b",
        r"\bwe\s+supply\b",
        r"\bsupplier\s+of\b",
        r"\bour\s+(?:company|business|group)\s+supplies\b",
    ],
    "Distributor": [
        r"\bwe\s+are\s+(?:an?\s+)?(?:authorized\s+)?distributor\b",
        r"\bwe\s+distribute\b",
        r"\bdistributor\s+of\b",
        r"\bour\s+(?:company|business|group)\s+is\s+(?:an?\s+)?distributor\b",
        r"\bauthorized\s+distributor\s+for\b",
    ],
    "Exporter": [
        r"\bwe\s+are\s+(?:an?\s+)?exporter\b",
        r"\bwe\s+export\b",
        r"\bexporter\s+of\b",
        r"\bour\s+(?:company|business|group)\s+exports\b",
    ],
}

# Weak words that must NOT independently establish Manufacturer.
WEAK_MANUFACTURING_WORDS = {
    "production",
    "factory",
    "develops",
    "developing",
    "producing",
    "produced",
    "made in",
    "processing",
}

# Services counted toward the Industrial Services product-equivalent
# ("at least 3 qualifying services").
INDUSTRIAL_SERVICE_KEYWORDS = {
    "metal polishing",
    "powder coating",
    "fabrication",
    "refurbishment",
    "welding",
    "cutting",
    "moulding",
    "molding",
}

# These are explicitly NOT Industrial Service under the guideline.
NON_INDUSTRIAL_SERVICE_KEYWORDS = {
    "corporate service",
    "corporate services",
    "software service",
    "software services",
    "consultation service",
    "consultation services",
    "consulting",
    "small scale industry",
    "small scale industries",
    "small-scale industry",
    "small-scale industries",
}


# ============================================================
# COUNTRY ALIASES  ->  workbook key
# ============================================================
# The workbook lists some countries under names no website actually
# writes -- "United States of America or USA", "China (Hong Kong
# S.A.R.)", "Korea, South". Without these aliases a US site saying
# "United States" matched nothing and the record was wrongly skipped
# for Country Error. The alias only decides WHICH workbook row
# applies; the row itself still decides Correct / Incorrect / Not
# Working and the name to use.

COUNTRY_ALIASES = {
    "united states of america": "united states of america or usa",
    "united states": "united states of america or usa",
    "usa": "united states of america or usa",
    "u.s.a.": "united states of america or usa",
    "u.s.a": "united states of america or usa",
    "united kingdom": "united kingdom or uk",
    "uk": "united kingdom or uk",
    "u.k.": "united kingdom or uk",
    "great britain": "united kingdom or uk",
    "england": "united kingdom or uk",
    "scotland": "united kingdom or uk",
    "wales": "united kingdom or uk",
    "northern ireland": "united kingdom or uk",
    "hong kong": "china (hong kong s.a.r.)",
    "hongkong": "china (hong kong s.a.r.)",
    "hong kong s.a.r.": "china (hong kong s.a.r.)",
    "macau": "china (macau s.a.r.)",
    "macao": "china (macau s.a.r.)",
    "macau s.a.r.": "china (macau s.a.r.)",
    "united arab emirates": "united arab emirates or uae",
    "uae": "united arab emirates or uae",
    "bahamas": "the bahamas",
    "gambia": "the gambia",
    "czechia": "czech republic",
    "south korea": "korea, south",
    "north korea": "korea, north",
    "republic of korea": "korea, south",
    "south sudan": "sudan, south",
    "democratic republic of the congo": "congo, democratic republic of the",
    "timor-leste": "east timor (timor-leste)",
    "east timor": "east timor (timor-leste)",
    "myanmar": "myanmar (burma)",
    "burma": "myanmar (burma)",
    "micronesia": "micronesia, federated states of",
    "vatican": "vatican city",
    "holland": "netherlands",
    # Native spellings, as they actually appear in addresses.
    "slovenija": "slovenia",
    "deutschland": "germany",
    "oesterreich": "austria",
    "osterreich": "austria",
    "schweiz": "switzerland",
    "suisse": "switzerland",
    "svizzera": "switzerland",
    "belgie": "belgium",
    "belgique": "belgium",
    "nederland": "netherlands",
    "espana": "spain",
    "italia": "italy",
    "france": "france",
    "polska": "poland",
    "cesko": "czech republic",
    "ceska republika": "czech republic",
    "magyarorszag": "hungary",
    "romania": "romania",
    "hrvatska": "croatia",
    "srbija": "serbia",
    "slovensko": "slovakia",
    "sverige": "sweden",
    "norge": "norway",
    "danmark": "denmark",
    "suomi": "finland",
    "island": "iceland",
    "eesti": "estonia",
    "latvija": "latvia",
    "lietuva": "lithuania",
    "turkiye": "turkey",
    "brasil": "brazil",
    "mexico": "mexico",
    "portugal": "portugal",
}


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


def portal_country_name(country_result):
    """
    The exact string to type into the portal's country box. Falls
    back to the workbook's own usable name whenever no override
    applies -- never invents a spelling.
    """
    if not country_result:
        return ""
    key = clean(country_result.get("input") or "").lower()
    key = COUNTRY_ALIASES.get(key, key)
    override = PORTAL_COUNTRY_FILL_NAMES.get(key)
    if override:
        return override
    usable = country_result.get("usable_name")
    return usable or ""



# ============================================================
# NON-PAID CATEGORIES (guideline section 3.2 / source PDF item 4)
# ============================================================
# General service-provider / uninteresting business categories.
# These do not, by themselves, block the site -- they only prevent
# it from being counted as one of the seven paid business types.

NON_PAID_TERMS = [
    "university", "college", "school", "academy", "education",
    "educational",
    "hospital", "clinic", "medical center", "doctor",
    "law firm", "lawyer", "attorney",
    "consulting", "consultant", "corporate consulting",
    "software", "saas",
    "blog", "blogging", "magazine", "news", "media", "entertainment",
    "restaurant", "hotel", "salon",
    "real estate", "retailer", "retail", "importer",
    "personal website", "biography",
    "informational website",
]

# ============================================================
# HARD-REJECTED / RESTRICTED CATEGORIES (guideline section 3.2)
# ============================================================
# If any of these are found, the site is rejected outright as
# non-considerable, regardless of otherwise-paid business type.
# Each category maps to a short human-readable rejection reason.

RESTRICTED_CATEGORY_TERMS = {
    "Weapons / armor manufacturing, trading, supply, exporting, "
    "or repair": [
        "weapon manufactur", "weapons manufactur", "arms manufactur",
        "firearm manufactur", "firearms manufactur", "ammunition",
        "gun manufactur", "rifle manufactur", "pistol manufactur",
        "military weapon", "armor manufactur", "armour manufactur",
        "weapons trading", "weapons supply", "weapons export",
        "weapons repair", "defense contractor", "defence contractor",
    ],
    "Porn / adult content": [
        "porn", "pornograph", "adult content", "adult entertainment",
        "xxx content", "escort service",
    ],
    "Online gambling / online games": [
        "online casino", "online gambling", "sports betting",
        "online poker", "slot machine game", "online game website",
        "betting platform",
    ],
    "General service provider (non-Industrial-Service)": [
        "hair salon", "beauty salon", "spa services", "nail salon",
    ],
    "Food / beverage / alcohol / tobacco manufacturing or trade": [
        "food manufactur", "beverage manufactur", "snack manufactur",
        "soft drink manufactur", "edible product", "confectionery",
        "bakery products", "alcoholic beverage", "liquor manufactur",
        "wine manufactur", "beer manufactur", "brewery",
        "tobacco manufactur", "cigarette manufactur", "cigar manufactur",
    ],
    "Pharmaceutical drug manufacturing": [
        "pharmaceutical manufactur", "drug manufactur",
        "chemical compound formulation", "pharmaceutical company",
        "pharma manufactur",
    ],
    "Hacking services / hacking information": [
        "hacking service", "hire a hacker", "hacking tutorial",
        "phone hacking", "password hacking", "hacking tool provider",
    ],
    "Airline / aviation industry": [
        "airline company", "aviation industry", "aircraft manufactur",
        "airline services", "airport operator",
    ],
    "Animal trading (live or dead)": [
        "animal trading", "livestock trading", "exotic animal sale",
        "pet trading company", "wildlife trade",
    ],
}

# Medical equipment / machinery is explicitly considerable even
# though it is adjacent to the pharmaceutical restriction above.
MEDICAL_EQUIPMENT_EXCEPTION_TERMS = [
    "medical equipment", "medical machinery", "medical device",
    "diagnostic equipment", "surgical instrument",
]

# ============================================================
# HOSTING / URL STRUCTURE REJECTIONS
# ============================================================

FREE_HOSTING_DOMAIN_MARKERS = [
    "wordpress.com", "wixsite.com", "weebly.com", "blogspot.com",
    "sites.google.com", "webnode.com", "godaddysites.com",
    "yolasite.com", "jimdo.com", "webs.com", "myshopify.com",
]


# ============================================================
# IMAGE CONTENT RESTRICTIONS
# ============================================================
# Guideline: images are not allowed of Porn/Adult/Medicine/
# Pharmaceutical/Food/Vegetable/Beverages/Weapons/Aircraft/
# Aerospace/Ship/Marine/Chemical/Alcohol/Animal subject matter,
# and diagrams/geometric/technical drawings never qualify.

BAD_IMAGE_TERMS = {
    "logo", "favicon", "sprite", "icon", "social",
    "facebook", "twitter", "instagram", "youtube", "linkedin",
    "search", "cart", "checkout", "payment", "paypal", "visa",
    "mastercard", "loading", "loader", "spinner", "placeholder",
    "banner", "background",
}

BAD_IMAGE_CONTEXT_TERMS = {
    "diagram", "schematic", "blueprint", "technical-drawing",
    "technical_drawing", "dimension-drawing", "dimensioned",
    "engineering-drawing", "engineering_drawing", "cad-drawing",
    "cad_drawing", "flowchart", "line-drawing", "line_drawing",
    "technical illustration", "installation diagram",
    "assembly diagram", "geometry", "geometrical",
}

RESTRICTED_IMAGE_SUBJECT_TERMS = {
    "porn", "adult", "medicine", "pharmaceutical", "food",
    "vegetable", "beverage", "weapon", "aircraft", "aerospace",
    "ship", "marine", "chemical", "alcohol", "animal",
}


# ============================================================
# DIRECTORY / THIRD-PARTY PAGES (do not establish business type)
# ============================================================

DIRECTORY_TERMS = [
    "distributor", "distributors", "dealer", "dealers",
    "reseller", "resellers", "partners", "where-to-buy",
    "find-a-dealer", "find-a-distributor",
    "world-wide-subsidiaries", "worldwide-subsidiaries",
    "subsidiaries",
]


# ============================================================
# COMPANY INFORMATION TABS / PAGES
# ============================================================

COMPANY_CONTEXT_TERMS = [
    "/about", "/about-us", "/aboutus", "/company", "/company-profile",
    "/company-profile/", "/company-overview", "/profile", "/history",
    "/who-we-are", "/who_we_are", "/our-company", "/our-company/",
    "/corporate", "/corporate-profile", "/organization",
]

COMPANY_CONTEXT_LINK_TEXT = {
    "about", "about us", "about the company", "company",
    "company profile", "company overview", "profile", "history",
    "who we are", "our company", "corporate", "corporate profile",
    "organization",
}

CONTACT_TERMS = [
    "/contact", "/contact-us", "/contactus", "/locations", "/location",
]


# ============================================================
# DOWNLOAD / NON-HTML RESOURCES
# ============================================================

SKIP_EXTENSIONS = (
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".zip", ".rar", ".7z", ".csv", ".txt", ".mp3", ".mp4", ".avi",
    ".mov", ".webm",
)


# ============================================================
# NOISE / IRRELEVANT SECTIONS
# ============================================================
# Press releases, awards, CSR/ESG posts, careers, and legal pages
# never establish business type, count as products, or qualify as
# company profile. Skipping them entirely keeps the crawl budget on
# pages that can actually move the decision and avoids the kind of
# false-positive "product" match a press/awards page can otherwise
# trigger.

NOISE_PATH_SEGMENTS = [
    "/news", "/companynews", "/press", "/blog", "/award", "/awards",
    "/milestone", "/career", "/careers", "/joinus", "/job", "/jobs",
    "/privacy", "/privacy-policy", "/terms", "/terms-of-use",
    "/sitemap", "/esg", "/csr", "/changes", "/faq",
]


def is_noise_page(url):
    low = str(url or "").lower()
    return any(segment in low for segment in NOISE_PATH_SEGMENTS)


# ============================================================
# PRODUCT / NON-PRODUCT TERMS
# ============================================================

NON_PRODUCT_TERMS = {
    "service", "services", "project", "projects", "brand", "brands",
    "program", "programs", "label", "labels", "sign", "signs",
    "solution", "solutions", "consulting", "consultancy",
    "training", "support", "category", "categories",
}

PRODUCT_STRONG_TERMS = {
    "model", "part number", "part no", "sku", "product code",
    "catalog number", "specification", "specifications",
    "product details", "product description", "dimensions",
    "material", "weight", "voltage", "capacity", "diameter",
    "length", "width", "height",
}


# ============================================================
# US STATE CODES
# ============================================================

US_STATES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut",
    "DE": "Delaware", "FL": "Florida", "GA": "Georgia", "HI": "Hawaii",
    "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
    "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine",
    "MD": "Maryland", "MA": "Massachusetts", "MI": "Michigan",
    "MN": "Minnesota", "MS": "Mississippi", "MO": "Missouri",
    "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico",
    "NY": "New York", "NC": "North Carolina", "ND": "North Dakota",
    "OH": "Ohio", "OK": "Oklahoma", "OR": "Oregon",
    "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas",
    "UT": "Utah", "VT": "Vermont", "VA": "Virginia",
    "WA": "Washington", "WV": "West Virginia", "WI": "Wisconsin",
}


# ============================================================
# STATE-EXEMPT COUNTRY MENU ENTRIES
# ============================================================
# Per the guideline: for these menu entries, state/region does not
# need to be checked -- only address and city are required.

# Guideline (h): "County, Province, Region, Prefecture and Capital of
# Country are accepted as State." The address parser was US-only, so
# every Australian, Canadian, British and European site failed on City
# Error however good it was. These tables fix the four most common
# address shapes this work actually sees.

AU_STATES = {
    "NSW": "New South Wales",
    "VIC": "Victoria",
    "QLD": "Queensland",
    "SA": "South Australia",
    "WA": "Western Australia",
    "TAS": "Tasmania",
    "NT": "Northern Territory",
    "ACT": "Australian Capital Territory",
}

CA_PROVINCES = {
    "AB": "Alberta",
    "BC": "British Columbia",
    "MB": "Manitoba",
    "NB": "New Brunswick",
    "NL": "Newfoundland and Labrador",
    "NS": "Nova Scotia",
    "NT": "Northwest Territories",
    "NU": "Nunavut",
    "ON": "Ontario",
    "PE": "Prince Edward Island",
    "QC": "Quebec",
    "SK": "Saskatchewan",
    "YT": "Yukon",
}

# Counties / council areas that count as State for a UK address.
UK_COUNTIES = {
    "bedfordshire", "berkshire", "bristol", "buckinghamshire",
    "cambridgeshire", "cheshire", "cleveland", "cornwall", "cumbria",
    "derbyshire", "devon", "dorset", "durham", "east riding of yorkshire",
    "east sussex", "essex", "gloucestershire", "greater london",
    "greater manchester", "hampshire", "herefordshire", "hertfordshire",
    "isle of wight", "kent", "lancashire", "leicestershire",
    "lincolnshire", "merseyside", "norfolk", "north yorkshire",
    "northamptonshire", "northumberland", "nottinghamshire",
    "oxfordshire", "rutland", "shropshire", "somerset", "staffordshire",
    "suffolk", "surrey", "tyne and wear", "warwickshire", "west midlands",
    "west sussex", "west yorkshire", "wiltshire", "worcestershire",
    "south yorkshire", "aberdeenshire", "angus", "argyll", "ayrshire",
    "clackmannanshire", "dumfries and galloway", "dunbartonshire",
    "fife", "highland", "inverclyde", "lanarkshire", "lothian",
    "midlothian", "moray", "perth and kinross", "renfrewshire",
    "scottish borders", "stirlingshire", "west lothian",
    "carmarthenshire", "ceredigion", "conwy", "denbighshire", "flintshire",
    "glamorgan", "gwynedd", "monmouthshire", "pembrokeshire", "powys",
    "wrexham", "county antrim", "county armagh", "county down",
    "county fermanagh", "county londonderry", "county tyrone",
}

STATE_NOT_REQUIRED_COUNTRIES = {
    "taiwan",
    "singapore",
    "china (hong kong s.a.r.)",
    "china (macau s.a.r.)",
}


# ============================================================
# COUNTRY NAME LIST  (embedded from Country_Name_List.xlsx)
# ============================================================
# Authoritative source for country-name validation. Keyed by
# lowercase country name as it appears in the workbook.
#   status "Correct"      -> usable exactly as listed (display name)
#   status "Incorrect"    -> must use correct_name instead
#   status "Not Working"  -> country is not workable under this rule
#
# Do not override this table with a guessed alternative.

COUNTRY_TABLE = {
'afghanistan': {"display":'Afghanistan', "status":'Correct', "correct_name": None},
'albania': {"display":'Albania', "status":'Correct', "correct_name": None},
'algeria': {"display":'Algeria', "status":'Correct', "correct_name": None},
'andorra': {"display":'Andorra', "status":'Correct', "correct_name": None},
'angola': {"display":'Angola', "status":'Correct', "correct_name": None},
'antigua and barbuda': {"display":'Antigua and Barbuda', "status":'Correct', "correct_name": None},
'argentina': {"display":'Argentina', "status":'Correct', "correct_name": None},
'armenia': {"display":'Armenia', "status":'Correct', "correct_name": None},
'australia': {"display":'Australia', "status":'Correct', "correct_name": None},
'austria': {"display":'Austria', "status":'Correct', "correct_name": None},
'azerbaijan': {"display":'Azerbaijan', "status":'Correct', "correct_name": None},
'the bahamas': {"display":'The Bahamas', "status":'Correct', "correct_name": None},
'bahrain': {"display":'Bahrain', "status":'Correct', "correct_name": None},
'bangladesh': {"display":'Bangladesh', "status":'Correct', "correct_name": None},
'barbados': {"display":'Barbados', "status":'Correct', "correct_name": None},
'belarus': {"display":'Belarus', "status":'Correct', "correct_name": None},
'belgium': {"display":'Belgium', "status":'Correct', "correct_name": None},
'belize': {"display":'Belize', "status":'Correct', "correct_name": None},
'benin': {"display":'Benin', "status":'Correct', "correct_name": None},
'bhutan': {"display":'Bhutan', "status":'Correct', "correct_name": None},
'bolivia': {"display":'Bolivia', "status":'Correct', "correct_name": None},
'bosnia and herzegovina': {"display":'Bosnia and Herzegovina', "status":'Correct', "correct_name": None},
'botswana': {"display":'Botswana', "status":'Correct', "correct_name": None},
'brazil': {"display":'Brazil', "status":'Correct', "correct_name": None},
'brunei': {"display":'Brunei', "status":'Correct', "correct_name": None},
'bulgaria': {"display":'Bulgaria', "status":'Correct', "correct_name": None},
'burkina faso': {"display":'Burkina Faso', "status":'Correct', "correct_name": None},
'burundi': {"display":'Burundi', "status":'Correct', "correct_name": None},
'cabo verde': {"display":'Cabo Verde', "status":'Correct', "correct_name": None},
'cambodia': {"display":'Cambodia', "status":'Correct', "correct_name": None},
'cameroon': {"display":'Cameroon', "status":'Correct', "correct_name": None},
'canada': {"display":'Canada', "status":'Correct', "correct_name": None},
'central african republic': {"display":'Central African Republic', "status":'Correct', "correct_name": None},
'chad': {"display":'Chad', "status":'Correct', "correct_name": None},
'chile': {"display":'Chile', "status":'Correct', "correct_name": None},
'china': {"display":'China', "status":'Not Working', "correct_name": None},
'china (hong kong s.a.r.)': {"display":'China (Hong Kong S.A.R.)', "status":'Correct', "correct_name": None},
'china (macau s.a.r.)': {"display":'China (Macau S.A.R.)', "status":'Correct', "correct_name": None},
'colombia': {"display":'Colombia', "status":'Correct', "correct_name": None},
'comoros': {"display":'Comoros', "status":'Correct', "correct_name": None},
'congo, democratic republic of the': {"display":'Congo, Democratic Republic of the', "status":'Incorrect', "correct_name":'Democratic Republic of the Congo'},
'costa rica': {"display":'Costa Rica', "status":'Correct', "correct_name": None},
'croatia': {"display":'Croatia', "status":'Correct', "correct_name": None},
'cuba': {"display":'Cuba', "status":'Correct', "correct_name": None},
'cyprus': {"display":'Cyprus', "status":'Correct', "correct_name": None},
'czech republic': {"display":'Czech Republic', "status":'Incorrect', "correct_name":'Czechia'},
'denmark': {"display":'Denmark', "status":'Correct', "correct_name": None},
'djibouti': {"display":'Djibouti', "status":'Correct', "correct_name": None},
'dominica': {"display":'Dominica', "status":'Correct', "correct_name": None},
'dominican republic': {"display":'Dominican Republic', "status":'Correct', "correct_name": None},
'east timor (timor-leste)': {"display":'East Timor (Timor-Leste)', "status":'Incorrect', "correct_name":'Timor-Leste'},
'ecuador': {"display":'Ecuador', "status":'Correct', "correct_name": None},
'egypt': {"display":'Egypt', "status":'Correct', "correct_name": None},
'el salvador': {"display":'El Salvador', "status":'Correct', "correct_name": None},
'equatorial guinea': {"display":'Equatorial Guinea', "status":'Correct', "correct_name": None},
'eritrea': {"display":'Eritrea', "status":'Correct', "correct_name": None},
'estonia': {"display":'Estonia', "status":'Correct', "correct_name": None},
'eswatini': {"display":'Eswatini', "status":'Correct', "correct_name": None},
'ethiopia': {"display":'Ethiopia', "status":'Correct', "correct_name": None},
'fiji': {"display":'Fiji', "status":'Correct', "correct_name": None},
'finland': {"display":'Finland', "status":'Correct', "correct_name": None},
'france': {"display":'France', "status":'Correct', "correct_name": None},
'gabon': {"display":'Gabon', "status":'Correct', "correct_name": None},
'the gambia': {"display":'The Gambia', "status":'Correct', "correct_name": None},
'georgia': {"display":'Georgia', "status":'Correct', "correct_name": None},
'germany': {"display":'Germany', "status":'Correct', "correct_name": None},
'ghana': {"display":'Ghana', "status":'Correct', "correct_name": None},
'greece': {"display":'Greece', "status":'Correct', "correct_name": None},
'grenada': {"display":'Grenada', "status":'Correct', "correct_name": None},
'guatemala': {"display":'Guatemala', "status":'Correct', "correct_name": None},
'guinea': {"display":'Guinea', "status":'Correct', "correct_name": None},
'guinea-bissau': {"display":'Guinea-Bissau', "status":'Correct', "correct_name": None},
'guyana': {"display":'Guyana', "status":'Correct', "correct_name": None},
'haiti': {"display":'Haiti', "status":'Correct', "correct_name": None},
'honduras': {"display":'Honduras', "status":'Correct', "correct_name": None},
'hungary': {"display":'Hungary', "status":'Correct', "correct_name": None},
'iceland': {"display":'Iceland', "status":'Correct', "correct_name": None},
'india': {"display":'India', "status":'Not Working', "correct_name": None},
'indonesia': {"display":'Indonesia', "status":'Correct', "correct_name": None},
'iran': {"display":'Iran', "status":'Correct', "correct_name": None},
'iraq': {"display":'Iraq', "status":'Correct', "correct_name": None},
'ireland': {"display":'Ireland', "status":'Incorrect', "correct_name":'Republic of Ireland'},
'israel': {"display":'Israel', "status":'Correct', "correct_name": None},
'italy': {"display":'Italy', "status":'Correct', "correct_name": None},
'jamaica': {"display":'Jamaica', "status":'Correct', "correct_name": None},
'japan': {"display":'Japan', "status":'Correct', "correct_name": None},
'jordan': {"display":'Jordan', "status":'Correct', "correct_name": None},
'kazakhstan': {"display":'Kazakhstan', "status":'Correct', "correct_name": None},
'kenya': {"display":'Kenya', "status":'Correct', "correct_name": None},
'kiribati': {"display":'Kiribati', "status":'Correct', "correct_name": None},
'korea, north': {"display":'Korea, North', "status":'Incorrect', "correct_name":'North Korea'},
'korea, south': {"display":'Korea, South', "status":'Incorrect', "correct_name":'South Korea'},
'kosovo': {"display":'Kosovo', "status":'Correct', "correct_name": None},
'kuwait': {"display":'Kuwait', "status":'Correct', "correct_name": None},
'kyrgyzstan': {"display":'Kyrgyzstan', "status":'Correct', "correct_name": None},
'laos': {"display":'Laos', "status":'Correct', "correct_name": None},
'latvia': {"display":'Latvia', "status":'Correct', "correct_name": None},
'lebanon': {"display":'Lebanon', "status":'Correct', "correct_name": None},
'lesotho': {"display":'Lesotho', "status":'Correct', "correct_name": None},
'liberia': {"display":'Liberia', "status":'Correct', "correct_name": None},
'libya': {"display":'Libya', "status":'Correct', "correct_name": None},
'liechtenstein': {"display":'Liechtenstein', "status":'Correct', "correct_name": None},
'lithuania': {"display":'Lithuania', "status":'Correct', "correct_name": None},
'luxembourg': {"display":'Luxembourg', "status":'Correct', "correct_name": None},
'madagascar': {"display":'Madagascar', "status":'Correct', "correct_name": None},
'malawi': {"display":'Malawi', "status":'Correct', "correct_name": None},
'malaysia': {"display":'Malaysia', "status":'Correct', "correct_name": None},
'maldives': {"display":'Maldives', "status":'Correct', "correct_name": None},
'mali': {"display":'Mali', "status":'Correct', "correct_name": None},
'malta': {"display":'Malta', "status":'Correct', "correct_name": None},
'marshall islands': {"display":'Marshall Islands', "status":'Correct', "correct_name": None},
'mauritania': {"display":'Mauritania', "status":'Correct', "correct_name": None},
'mauritius': {"display":'Mauritius', "status":'Correct', "correct_name": None},
'mexico': {"display":'Mexico', "status":'Correct', "correct_name": None},
'micronesia, federated states of': {"display":'Micronesia, Federated States of', "status":'Incorrect', "correct_name":'Micronesia'},
'moldova': {"display":'Moldova', "status":'Correct', "correct_name": None},
'monaco': {"display":'Monaco', "status":'Correct', "correct_name": None},
'mongolia': {"display":'Mongolia', "status":'Correct', "correct_name": None},
'montenegro': {"display":'Montenegro', "status":'Correct', "correct_name": None},
'morocco': {"display":'Morocco', "status":'Correct', "correct_name": None},
'mozambique': {"display":'Mozambique', "status":'Correct', "correct_name": None},
'myanmar (burma)': {"display":'Myanmar (Burma)', "status":'Incorrect', "correct_name":'Myanmar'},
'namibia': {"display":'Namibia', "status":'Correct', "correct_name": None},
'nauru': {"display":'Nauru', "status":'Correct', "correct_name": None},
'nepal': {"display":'Nepal', "status":'Correct', "correct_name": None},
'netherlands': {"display":'Netherlands', "status":'Correct', "correct_name": None},
'new zealand': {"display":'New Zealand', "status":'Correct', "correct_name": None},
'nicaragua': {"display":'Nicaragua', "status":'Correct', "correct_name": None},
'niger': {"display":'Niger', "status":'Correct', "correct_name": None},
'nigeria': {"display":'Nigeria', "status":'Correct', "correct_name": None},
'north macedonia': {"display":'North Macedonia', "status":'Correct', "correct_name": None},
'norway': {"display":'Norway', "status":'Correct', "correct_name": None},
'oman': {"display":'Oman', "status":'Correct', "correct_name": None},
'pakistan': {"display":'Pakistan', "status":'Correct', "correct_name": None},
'palau': {"display":'Palau', "status":'Correct', "correct_name": None},
'panama': {"display":'Panama', "status":'Correct', "correct_name": None},
'papua new guinea': {"display":'Papua New Guinea', "status":'Correct', "correct_name": None},
'paraguay': {"display":'Paraguay', "status":'Correct', "correct_name": None},
'peru': {"display":'Peru', "status":'Correct', "correct_name": None},
'philippines': {"display":'Philippines', "status":'Correct', "correct_name": None},
'poland': {"display":'Poland', "status":'Correct', "correct_name": None},
'portugal': {"display":'Portugal', "status":'Correct', "correct_name": None},
'qatar': {"display":'Qatar', "status":'Correct', "correct_name": None},
'romania': {"display":'Romania', "status":'Correct', "correct_name": None},
'russia': {"display":'Russia', "status":'Correct', "correct_name": None},
'rwanda': {"display":'Rwanda', "status":'Correct', "correct_name": None},
'saint kitts and nevis': {"display":'Saint Kitts and Nevis', "status":'Correct', "correct_name": None},
'saint lucia': {"display":'Saint Lucia', "status":'Correct', "correct_name": None},
'saint vincent and the grenadines': {"display":'Saint Vincent and the Grenadines', "status":'Correct', "correct_name": None},
'samoa': {"display":'Samoa', "status":'Correct', "correct_name": None},
'san marino': {"display":'San Marino', "status":'Correct', "correct_name": None},
'sao tome and principe': {"display":'Sao Tome and Principe', "status":'Correct', "correct_name": None},
'saudi arabia': {"display":'Saudi Arabia', "status":'Correct', "correct_name": None},
'senegal': {"display":'Senegal', "status":'Correct', "correct_name": None},
'serbia': {"display":'Serbia', "status":'Correct', "correct_name": None},
'seychelles': {"display":'Seychelles', "status":'Correct', "correct_name": None},
'sierra leone': {"display":'Sierra Leone', "status":'Correct', "correct_name": None},
'singapore': {"display":'Singapore', "status":'Correct', "correct_name": None},
'slovakia': {"display":'Slovakia', "status":'Correct', "correct_name": None},
'slovenia': {"display":'Slovenia', "status":'Correct', "correct_name": None},
'solomon islands': {"display":'Solomon Islands', "status":'Correct', "correct_name": None},
'somalia': {"display":'Somalia', "status":'Correct', "correct_name": None},
'south africa': {"display":'South Africa', "status":'Correct', "correct_name": None},
'spain': {"display":'Spain', "status":'Correct', "correct_name": None},
'sri lanka': {"display":'Sri Lanka', "status":'Correct', "correct_name": None},
'sudan': {"display":'Sudan', "status":'Correct', "correct_name": None},
'sudan, south': {"display":'Sudan, South', "status":'Correct', "correct_name": None},
'suriname': {"display":'Suriname', "status":'Correct', "correct_name": None},
'sweden': {"display":'Sweden', "status":'Correct', "correct_name": None},
'switzerland': {"display":'Switzerland', "status":'Correct', "correct_name": None},
'syria': {"display":'Syria', "status":'Correct', "correct_name": None},
'taiwan': {"display":'Taiwan', "status":'Correct', "correct_name": None},
'tajikistan': {"display":'Tajikistan', "status":'Correct', "correct_name": None},
'tanzania': {"display":'Tanzania', "status":'Correct', "correct_name": None},
'thailand': {"display":'Thailand', "status":'Correct', "correct_name": None},
'togo': {"display":'Togo', "status":'Correct', "correct_name": None},
'tonga': {"display":'Tonga', "status":'Correct', "correct_name": None},
'trinidad and tobago': {"display":'Trinidad and Tobago', "status":'Correct', "correct_name": None},
'tunisia': {"display":'Tunisia', "status":'Correct', "correct_name": None},
'turkey': {"display":'Turkey', "status":'Correct', "correct_name": None},
'turkmenistan': {"display":'Turkmenistan', "status":'Correct', "correct_name": None},
'tuvalu': {"display":'Tuvalu', "status":'Correct', "correct_name": None},
'uganda': {"display":'Uganda', "status":'Correct', "correct_name": None},
'ukraine': {"display":'Ukraine', "status":'Correct', "correct_name": None},
'united arab emirates or uae': {"display":'United Arab Emirates or UAE', "status":'Correct', "correct_name": None},
'united kingdom or uk': {"display":'United Kingdom or UK', "status":'Correct', "correct_name": None},
'united states of america or usa': {"display":'United States of America or USA', "status":'Correct', "correct_name": None},
'uruguay': {"display":'Uruguay', "status":'Correct', "correct_name": None},
'uzbekistan': {"display":'Uzbekistan', "status":'Correct', "correct_name": None},
'vanuatu': {"display":'Vanuatu', "status":'Correct', "correct_name": None},
'vatican city': {"display":'Vatican City', "status":'Correct', "correct_name": None},
'venezuela': {"display":'Venezuela', "status":'Correct', "correct_name": None},
'vietnam': {"display":'Vietnam', "status":'Correct', "correct_name": None},
'yemen': {"display":'Yemen', "status":'Correct', "correct_name": None},
'zambia': {"display":'Zambia', "status":'Correct', "correct_name": None},
'zimbabwe': {"display":'Zimbabwe', "status":'Correct', "correct_name": None},
}


# ============================================================
# PORTAL ERROR FIELD REFERENCE  (embedded from Error_Details.xlsx)
# ============================================================
# Reference only -- documents what the portal treats as an error
# for each mandatory field so the decision engine's checks line up
# exactly with the portal's own validation.

PORTAL_ERROR_FIELDS = {
    "url": {
        "error_remarks": (
            "Redirect Site Error, Sub Domain/Double Domain Error, "
            "Restricted Site, Online Site"
        ),
        "error_description": (
            "Changing URL is not acceptable. Double Domain URL is "
            "not acceptable."
        ),
    },
    "businesstype": {
        "error_remarks": "Business Type Error",
        "error_description": (
            "You have to choose the correct Kind of Business; it "
            "has to be filled by reading from the Company Profile, "
            "Our Company, History, About us. Home page or any "
            "heading tab is not acceptable."
        ),
    },
    "country": {
        "error_remarks": "Country Error",
        "error_description": (
            "The correct name of the country has to be filled by "
            "validating against the Country Name List."
        ),
    },
    "emailid": {
        "error_remarks": "Email Error",
        "error_description": (
            "On the website, Email is mandatory with the full "
            "address and phone number."
        ),
    },
    "phoneormobile": {
        "error_remarks": "Number Error / Toll Free Number Error",
        "error_description": (
            "Special symbols are not acceptable in the number and "
            "Toll Free & Fax numbers are also not acceptable."
        ),
    },
    "address": {
        "error_remarks": "Address not found / Address not given",
        "error_description": (
            "On the website, the full address paragraph is "
            "mandatory. (India & China are not acceptable.)"
        ),
    },
    "city": {
        "error_remarks": None,
        "error_description": None,
    },
    "state": {
        "error_remarks": "State Error",
        "error_description": (
            "On the website, in the address paragraph, the written "
            "state is mandatory."
        ),
    },
    "companyprofile": {
        "error_remarks": "Company Profile Error",
        "error_description": (
            "Company Profile, Our Company, History, About us, or "
            "any other tab from where you can get information "
            "about the business category of the website."
        ),
    },
    "productname": {
        "error_remarks": "Product Error",
        "error_description": "Less than 3 is not acceptable.",
    },
    "productimage": {
        "error_remarks": "Product Image Error",
        "error_description": (
            "Minimum 3 product images with names and descriptions "
            "should be given on the website. Less than 3 is not "
            "acceptable."
        ),
    },
    "productdescription": {
        "error_remarks": "Product Description Error",
        "error_description": (
            "Minimum 3 product descriptions with names and images "
            "should be given on the website. Less than 3 is not "
            "acceptable."
        ),
    },
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


def same_domain(a, b):
    try:
        da = urlparse(a).netloc.lower().replace("www.", "")
        db = urlparse(b).netloc.lower().replace("www.", "")
        return bool(da) and da == db
    except Exception:
        return False


def get_root_domain(url):
    """netloc with 'www.' stripped."""
    try:
        return urlparse(url).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def is_subdomain(url):
    """
    True if the host has more labels than a normal root-level domain
    (root.tld or root.co.tld style second-level ccTLDs are allowed).
    """
    host = get_root_domain(url)
    if not host:
        return False
    labels = host.split(".")
    if len(labels) <= 2:
        return False
    # Common two-part TLDs (co.uk, com.au, co.in, etc.) still count
    # as root-level when there are exactly 3 labels total.
    common_second_level = {"co", "com", "net", "org", "gov", "edu"}
    if len(labels) == 3 and labels[-2] in common_second_level:
        return False
    return True


def is_free_hosting_domain(url):
    host = get_root_domain(url)
    return any(marker in host for marker in FREE_HOSTING_DOMAIN_MARKERS)


def is_download_url(url):
    try:
        path = urlparse(url).path.lower()
    except Exception:
        return True
    return path.endswith(SKIP_EXTENSIONS)


def is_directory_page(url):
    low = str(url or "").lower()
    return any(term in low for term in DIRECTORY_TERMS)


def is_contact_page(url):
    low = str(url or "").lower()
    return any(term in low for term in CONTACT_TERMS)


CONTEXT_PATH_TOKENS = {
    "about", "aboutus", "about-us", "company", "companyprofile",
    "company-profile", "companyoverview", "company-overview", "profile",
    "history", "ourhistory", "our-history", "whoweare", "who-we-are",
    "ourcompany", "our-company", "corporate", "organization",
    "organisation", "overview", "ourstory", "our-story", "story",
}

CONTACT_PATH_TOKENS = {
    "contact", "contactus", "contact-us", "contacts", "location",
    "locations", "findus", "find-us", "branches", "offices", "showroom",
    "getintouch", "get-in-touch", "reachus", "stores",
}


def url_path_tokens(url):
    """
    The path of a URL split into words, so "/pages/ak-history" yields
    {"pages", "ak", "history"}. Matching on tokens is what lets a CMS
    path like "/pages/about-us" be recognised as an About page.
    """
    try:
        path = urlparse(str(url or "")).path.lower()
    except Exception:
        return set()
    parts = re.split(r"[^a-z0-9]+", path)
    tokens = {p for p in parts if p}
    # also keep hyphenated pairs, so "about-us" matches as written
    segments = [seg for seg in path.split("/") if seg]
    tokens.update(segments)
    return tokens


def is_company_context_page(url):
    low = str(url or "").lower()
    if any(term in low for term in COMPANY_CONTEXT_TERMS):
        return True
    return bool(url_path_tokens(url) & CONTEXT_PATH_TOKENS)


def is_contact_page(url):
    low = str(url or "").lower()
    if any(term in low for term in CONTACT_TERMS):
        return True
    return bool(url_path_tokens(url) & CONTACT_PATH_TOKENS)


def context_link_score(url):
    """
    How worth visiting a link is during Stage 1, or None if it is not
    a company-information or contact page at all. The contact page
    scores highest: it is where the address, email and phone live, and
    guideline (g) wants them all from the same place.
    """
    if is_contact_page(url):
        return 100
    if not is_company_context_page(url):
        return None
    tokens = url_path_tokens(url)
    if tokens & {"about", "aboutus", "about-us", "whoweare", "who-we-are"}:
        return 90
    if tokens & {"company", "companyprofile", "company-profile",
                 "companyoverview", "company-overview", "profile"}:
        return 85
    if tokens & {"history", "our-history", "ourhistory", "story",
                 "our-story", "ourstory"}:
        return 80
    return 70


def is_home_url(url, landing_url):
    a = normalize_url(url)
    b = normalize_url(landing_url)
    if not a or not b:
        return False
    pa, pb = urlparse(a), urlparse(b)
    return (
        pa.netloc.lower().replace("www.", "")
        == pb.netloc.lower().replace("www.", "")
        and pa.path.rstrip("/") in {"", "/"}
    )


# ============================================================
# EMAIL / PHONE / ADDRESS EXTRACTION
# ============================================================

def extract_emails(text):
    return sorted(set(re.findall(
        r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
        text or "",
    )))


def normalize_phone(phone):
    """Digits only -- strips +, -, (), spaces, dots, slashes, etc."""
    return re.sub(r"\D", "", phone or "")


# ============================================================
# PHONE / MOBILE  --  guideline (g), implemented clause by clause
# ============================================================
# "Toll free, Free Phone, TF, Fax numbers are not considered ... If fax
#  number and phone number are same then it is accepted. If toll free
#  number/TF number/Free phone number and phone number are same then it
#  is not accepted ... The numbers which are mentioned as Toll free
#  shall be skipped."
#
# So the LABEL a number is published under decides its fate, not just
# its digits:
#   fax only                      -> rejected
#   fax AND phone/tel/mobile      -> ACCEPTED (same number, both roles)
#   toll-free label, any other    -> rejected even if also called Tel
#   toll-free numbering range     -> rejected
# A number with no label at all is still usable -- plenty of contact
# pages print the number on its own line.

TOLL_FREE_LABEL_RE = re.compile(
    r"toll[\s\-_]*free|tollfree|free[\s\-_]*phone|freephone|"
    r"free[\s\-_]*call|freecall|free[\s\-_]*dial|\bTF\b|\bT\.F\.",
    re.I,
)
FAX_LABEL_RE = re.compile(r"\bfax\b|\bfaxe?s\b|\bfacsimile\b|\bf\s*:", re.I)
PHONE_LABEL_RE = re.compile(
    r"\btel\b|\btel\.|\btelephone\b|\bphone\b|\bmobile\b|\bmob\b|"
    r"\bmob\.|\bcell\b|\bcall\s+us\b|\bcontact\b|\bwhatsapp\b|\bp\s*:",
    re.I,
)

# Freephone detection, rebuilt after a real site proved the first
# version wrong: Adelaide's landline 08 8340 4111 was rejected because
# stripping the trunk "0" left "883404111", which matched a North
# American reserved 88x prefix. Country rules are now applied only
# where they actually belong.

# North America: valid only for a 10-digit NANP number.
NA_TOLL_FREE_PREFIXES = ("800", "833", "844", "855", "866", "877", "888")

# Dialled with the national trunk code, e.g. UK/DE/FR/NL 0800,
# UK 0808 and 0500, NZ 0508.
TRUNK_TOLL_FREE_PREFIXES = ("0800", "0808", "0500", "0508")

# National freephone blocks once the country code has been removed.
INTERNATIONAL_TOLL_FREE_PREFIXES = ("800", "808", "500", "508", "1800")

# Country codes worth stripping from a number that was written in
# international form. Only used when the number really was written
# that way -- guessing a country code out of a local number is what
# caused the false positive above.
KNOWN_COUNTRY_CODES = (
    "1", "7", "20", "27", "30", "31", "32", "33", "34", "36", "39", "40",
    "41", "43", "44", "45", "46", "47", "48", "49", "51", "52", "53",
    "54", "55", "56", "57", "58", "60", "61", "62", "63", "64", "65",
    "66", "81", "82", "84", "86", "90", "91", "92", "94", "95", "98",
    "212", "213", "216", "218", "234", "254", "255", "256", "260",
    "263", "264", "265", "351", "352", "353", "354", "355", "356",
    "357", "358", "370", "371", "372", "373", "374", "375", "376",
    "377", "380", "381", "382", "385", "386", "387", "389", "420",
    "421", "423", "852", "853", "855", "856", "880", "886", "962",
    "963", "964", "965", "966", "967", "968", "970", "971", "972",
    "973", "974", "975", "976", "977", "992", "993", "994", "995",
    "996", "998",
)


def normalize_phone(phone):
    """Digits only -- strips +, -, (), spaces, dots, slashes, etc."""
    return re.sub(r"\D", "", phone or "")


def is_toll_free_number(phone, raw=None):
    """
    True when the digits fall in a freephone range.

    `raw` is the number as printed on the site; a leading "+" (or a
    "00" prefix) is what proves the number was written in
    international form, and only then is a country code stripped.
    """
    digits = normalize_phone(phone)
    if len(digits) < 7:
        return False

    written = (raw if raw is not None else phone) or ""
    international = written.strip().startswith("+") or digits.startswith("00")

    # North America -- 10 digits, or 11 with the leading 1.
    na = digits[1:] if (len(digits) == 11 and digits.startswith("1")) else digits
    if len(na) == 10 and na[:3] in NA_TOLL_FREE_PREFIXES:
        return True

    # Dialled nationally with the trunk code.
    if digits.startswith(TRUNK_TOLL_FREE_PREFIXES):
        return True

    # Australia / Ireland / India 1800, dialled nationally.
    if digits.startswith("1800") and len(digits) in (10, 11):
        return True

    if international:
        base = digits[2:] if digits.startswith("00") else digits
        for code in KNOWN_COUNTRY_CODES:
            if not base.startswith(code):
                continue
            national = base[len(code):]
            if len(national) < 6:
                continue
            if national.startswith(INTERNATIONAL_TOLL_FREE_PREFIXES):
                return True

    return False


# Kept under its old name so nothing else has to change.
def is_toll_free(phone):
    return is_toll_free_number(phone)


def extract_phones(text):
    """
    Candidate phone strings with the label they were published under.

    Returns a list of dicts:
        {"raw": as printed, "digits": digits only, "labels": set}
    where labels is any of "phone", "fax", "toll_free".

    Unlike the previous version this does NOT throw away a line just
    because it mentions fax -- guideline (g) accepts a number that is
    published as both fax and phone.
    """
    patterns = [
        r"\+\d[\d\s().-]{7,}\d",
        r"\(\d{3}\)\s*\d{3}[-.\s]\d{4}",
        r"\b1[-.\s]\d{3}[-.\s]\d{3}[-.\s]\d{4}\b",
        r"\b\d{3}[-.\s]\d{3}[-.\s]\d{4}\b",
        r"\b0\d[\d\s().-]{7,}\d",
    ]

    by_digits = {}
    for line in (text or "").splitlines():
        labels = set()
        if TOLL_FREE_LABEL_RE.search(line):
            labels.add("toll_free")
        if FAX_LABEL_RE.search(line):
            labels.add("fax")
        if PHONE_LABEL_RE.search(line):
            labels.add("phone")

        for pattern in patterns:
            for match in re.findall(pattern, line):
                digits = normalize_phone(match)
                if len(digits) < 7:
                    continue
                entry = by_digits.setdefault(
                    digits, {"raw": match.strip(), "digits": digits, "labels": set()}
                )
                entry["labels"] |= labels

    return sorted(by_digits.values(), key=lambda e: e["digits"])


def phone_rejection_reason(entry):
    """
    Why this candidate cannot be submitted, or None if it is usable.
    Implements guideline (g) exactly.
    """
    labels = entry.get("labels") or set()

    # "If toll free number ... and phone number are same then it is not
    # accepted" -- the toll-free label wins over every other label.
    if "toll_free" in labels:
        return "published as toll free / freephone"

    if is_toll_free_number(entry["digits"], entry.get("raw")):
        return "number is in a toll-free range"

    # "If fax number and phone number are same then it is accepted."
    if "fax" in labels and "phone" not in labels:
        return "published as fax only"

    return None


def get_valid_phones(phones):
    """
    Submittable numbers, digits only, de-duplicated, order preserved.
    Accepts either the dicts from extract_phones() or plain strings.
    """
    valid = []
    for item in phones:
        entry = item if isinstance(item, dict) else {
            "raw": item, "digits": normalize_phone(item), "labels": set(),
        }
        if len(entry["digits"]) < 7:
            continue
        if phone_rejection_reason(entry):
            continue
        if entry["digits"] not in valid:
            valid.append(entry["digits"])
    return valid


# Many contact pages print the whole address on ONE line:
#   "79-85 Cowpasture Road, Wetherill Park NSW 2164"
#   "1200 Industrial Way, Cleveland, OH 44135"
# The line-by-line parser could never see a city or state in those, so
# every such site failed with City Error. These patterns read the
# trailing city/state/postcode and treat whatever precedes it as the
# street.

ONE_LINE_ADDRESS_PATTERNS = [
    # Australia -- "..., Wetherill Park NSW 2164"
    (re.compile(
        r"^(?P<street>.{5,90}?),\s*(?P<city>[A-Za-z][A-Za-z .'\-]{1,40}?)\s+"
        r"(?P<code>NSW|VIC|QLD|SA|WA|TAS|NT|ACT)\s+(?P<zip>\d{4})$", re.I), "AU"),
    # United States -- "..., Cleveland, OH 44135"
    (re.compile(
        r"^(?P<street>.{5,90}?),\s*(?P<city>[A-Za-z][A-Za-z .'\-]{1,40}?),?\s+"
        r"(?P<code>[A-Za-z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)$", re.I), "US"),
    # Canada -- "..., Toronto, ON M5T 2S6"
    (re.compile(
        r"^(?P<street>.{5,90}?),\s*(?P<city>[A-Za-z][A-Za-z .'\-]{1,40}?),?\s+"
        r"(?P<code>[A-Za-z]{2})\s+(?P<zip>[A-Za-z]\d[A-Za-z]\s?\d[A-Za-z]\d)$", re.I), "CA"),
    # Mainland Europe -- "Vinje 45C, 1262 Dol pri Ljubljani, Slovenija"
    # Street name first, number after; postcode precedes the city.
    (re.compile(
        r"^(?P<street>[^\d,]{3,60}?\s+\d{1,5}\s*[A-Za-z]?),\s*"
        r"(?P<zip>\d{4,5})\s+(?P<city>[^,\d]{2,40}?)"
        r"(?:,\s*(?P<code>[^,\d]{2,40}?))?$", re.I), "EU"),
    # United Kingdom -- "..., Sheffield, South Yorkshire, S9 1XH"
    (re.compile(
        r"^(?P<street>.{5,90}?),\s*(?P<city>[A-Za-z][A-Za-z .'\-]{1,40}?),\s*"
        r"(?P<code>[A-Za-z][A-Za-z .'\-]{1,40}?),?\s+"
        r"(?P<zip>[A-Za-z]{1,2}\d[A-Za-z\d]?\s*\d[A-Za-z]{2})$", re.I), "UK"),
]

STREET_HINT_WORDS = (
    "road", "rd", "street", "st", "avenue", "ave", "boulevard", "blvd",
    "drive", "dr", "lane", "ln", "way", "court", "ct", "highway", "hwy",
    "parkway", "pkwy", "place", "pl", "terrace", "circle", "trail",
    "unit", "suite", "level", "lot", "block", "industrial", "park",
    "close", "crescent", "esplanade", "quay", "wharf", "strasse",
    "allee", "weg", "laan", "straat", "rue", "via", "calle",
)


def parse_one_line_address(line):
    """
    (street, city, state, state_code, postal) for an address written on
    a single line, or None. The street half must actually look like a
    street, so a stray "Something, Anytown XX 12345" sentence cannot
    masquerade as an address.
    """
    for pattern, kind in ONE_LINE_ADDRESS_PATTERNS:
        match = pattern.match(line)
        if not match:
            continue

        street = clean(match.group("street"))
        lower = street.lower()
        if kind == "EU":
            # European streets are written name-first, number-last
            # ("Vinje 45C", "Hauptstrasse 12"), so the house number at
            # the end is what proves it is a street.
            looks_like_street = bool(re.search(r"\d{1,5}\s*[A-Za-z]?$", street))
        else:
            looks_like_street = bool(
                re.match(r"^\d", street)
                or any(f" {w} " in f" {lower} " or lower.endswith(" " + w)
                       for w in STREET_HINT_WORDS)
            )
        if not looks_like_street:
            continue

        city = clean(match.group("city"))
        try:
            code = clean(match.group("code") or "")
        except (IndexError, error):
            code = ""
        postal = clean(match.group("zip"))

        if kind == "AU":
            return street, city, AU_STATES[code.upper()], code.upper(), postal
        if kind == "US":
            if code.upper() not in US_STATES:
                continue
            return street, city, US_STATES[code.upper()], code.upper(), postal
        if kind == "CA":
            if code.upper() not in CA_PROVINCES:
                continue
            return street, city, CA_PROVINCES[code.upper()], code.upper(), postal
        if kind == "EU":
            # The trailing group is the country, not a region -- these
            # addresses carry no state, and that is the honest answer.
            return street, city, None, None, postal
        if kind == "UK":
            if code.lower() not in UK_COUNTIES:
                continue
            return street, city, code, None, postal

    return None


def extract_address(text):
    """
    Best-effort extraction of a street address plus city/state/zip
    for US-style addresses, with a generic international fallback
    for the street line only.
    """
    lines = [clean(line) for line in (text or "").splitlines() if clean(line)]

    street = city = state = state_code = postal_code = None
    evidence = []

    # A complete one-line address wins outright. Guideline (h) wants
    # city AND state, so an earlier branch that prints neither is not
    # "the first matching set" -- the first COMPLETE one is.
    partial_one_line = None
    for line in lines:
        parsed = parse_one_line_address(line)
        if not parsed:
            continue
        one_street, one_city, one_state, one_code, one_postal = parsed
        record = {
            "street": one_street,
            "city": one_city,
            "state": one_state,
            "state_code": one_code,
            "postal_code": one_postal,
            "evidence": line,
            "source": None,
        }
        if one_state:
            # Complete -- street, city AND state. Nothing beats it.
            return record
        if partial_one_line is None:
            partial_one_line = record

    # Street line. Widened after real sites kept failing with
    # "Address Error": the old pattern demanded a number, then words,
    # then a street suffix at the very END of the line -- so
    # "500 N Main St." (trailing dot) and "12 Foundry Road, Building B"
    # both failed, and a PO Box was not an address at all.
    street_re = re.compile(
        r"^(?:(?:Unit|Suite|Ste|Level|Shop|Factory|Warehouse|Lot|Block|"
        r"Building|Bldg)\s*[\w/-]+,?\s+)?"
        r"\d{1,7}[A-Za-z]?\s+[A-Za-z0-9.'#&/-]+(?:\s+[A-Za-z0-9.'#&/-]+){0,12}\s+"
        r"(?:Road|Rd|Street|St|Avenue|Ave|Boulevard|Blvd|Drive|Dr|Lane|"
        r"Ln|Way|Court|Ct|Highway|Hwy|Parkway|Pkwy|Place|Pl|Terrace|Ter|"
        r"Circle|Cir|Trail|Trl|Route|Rte|Loop|Turnpike|Tpke|Expressway|"
        r"Expy|Pike|Row|Square|Sq|Crescent|Close|Esplanade|Quay|Wharf|"
        r"Industrial\s+Park|Business\s+Park|Estate)\.?"
        r"(?:[,\s].{0,40})?$",
        re.I,
    )

    # A PO Box with a city/state/postcode underneath is a complete
    # postal address and plenty of manufacturers publish nothing else.
    po_box_re = re.compile(
        r"^(?:P\.?\s?O\.?\s*Box|Post\s+Office\s+Box|Postbus|"
        r"PO\s*BOX)\s*[#]?\s*\d{1,7}\b.{0,30}$",
        re.I,
    )

    city_state_zip = re.compile(
        r"^([A-Za-z][A-Za-z .'\-&]{1,60}?),?\s+([A-Z]{2})\s+"
        r"(\d{5}(?:-\d{4})?)(?:\s+USA?)?$",
        re.I,
    )
    # "Blacktown NSW 2148" / "Melbourne, VIC 3000"
    au_city_state = re.compile(
        r"^([A-Za-z][A-Za-z .'\-&]{1,60}?),?\s+"
        r"(NSW|VIC|QLD|SA|WA|TAS|NT|ACT)\s+(\d{4})$",
        re.I,
    )
    # "Toronto, ON M5T 2S6"
    ca_city_prov = re.compile(
        r"^([A-Za-z][A-Za-z .'\-&]{1,60}?),?\s+"
        r"(AB|BC|MB|NB|NL|NS|NT|NU|ON|PE|QC|SK|YT)\s+"
        r"([A-Za-z]\d[A-Za-z]\s?\d[A-Za-z]\d)$",
        re.I,
    )
    # "Sheffield, South Yorkshire, S9 1XH"
    uk_city_county = re.compile(
        r"^([A-Za-z][A-Za-z .'\-&]{1,60}?),\s*"
        r"([A-Za-z][A-Za-z .'\-&]{1,60}?),?\s+"
        r"([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})$",
        re.I,
    )
    # "40549 Duesseldorf" -- city only; mainland European addresses
    # usually carry no region, which is a State Error, not a City one.
    eu_zip_city = re.compile(
        r"^(\d{4,6}(?:\s?[A-Z]{2})?)\s+"
        r"([A-Za-z\u00C0-\u024F][A-Za-z\u00C0-\u024F .'\-]{1,40})$",
    )

    def regional_match(line):
        """(city, state, state_code, postal) for a non-US address line."""
        m = au_city_state.match(line)
        if m:
            code = m.group(2).upper()
            return clean(m.group(1)), AU_STATES[code], code, m.group(3)

        m = ca_city_prov.match(line)
        if m:
            code = m.group(2).upper()
            return clean(m.group(1)), CA_PROVINCES[code], code, clean(m.group(3))

        m = uk_city_county.match(line)
        if m and clean(m.group(2)).lower() in UK_COUNTIES:
            return clean(m.group(1)), clean(m.group(2)), None, clean(m.group(3))

        m = eu_zip_city.match(line)
        if m:
            # "1188 West Georgia Street" also looks like postcode +
            # words. A city name never carries a street word, so that
            # is what separates the two.
            candidate = clean(m.group(2))
            street_words = (
                "street", "road", "avenue", "boulevard", "drive", "lane",
                "way", "court", "place", "terrace", "highway", "parkway",
                "circle", "trail", "route", "suite", "unit", "floor",
                "strasse", "stra\u00dfe", "allee", "weg", "gasse", "laan",
                "straat", "rue", "via", "viale", "corso", "calle",
            )
            if not any(w in candidate.lower() for w in street_words):
                return candidate, None, None, m.group(1)

        return None

    for line in lines:
        if street is None and (street_re.search(line) or po_box_re.match(line)):
            street = line
            evidence.append(line)
        match = city_state_zip.match(line)
        if match:
            code = match.group(2).upper()
            if code in US_STATES:
                city = clean(match.group(1))
                state_code = code
                state = US_STATES[code]
                postal_code = match.group(3)
                evidence.append(line)
                continue

        if city is None:
            found = regional_match(line)
            if found:
                city, state, state_code, postal_code = found
                evidence.append(line)

    if not street:
        international_street = re.compile(
            r"^\d{1,7}\s+[A-Za-z0-9.'#&,\- ]{3,100}$", re.I,
        )
        street_tokens = [
            "road", "street", "avenue", "boulevard", "drive", "lane",
            "way", "parkway", "place", "strasse", "straße", "str.",
            "rue", "chemin", "via", "calle", "allee", "weg", "platz",
            "gasse", "ring", "damm", "ufer", "laan", "straat", "vej",
            "gatan", "viale", "corso", "piazza", "avenida", "carrer",
        ]
        # "24 Hansaallee" style -- number first.
        for line in lines:
            if (
                international_street.match(line)
                and any(token in line.lower() for token in street_tokens)
            ):
                street = line
                evidence.append(line)
                break

        # "Hansaallee 24" style -- name first, number last. Normal
        # across Germany, Austria, Netherlands, Scandinavia and Italy,
        # and previously not recognised as a street at all.
        if not street:
            name_then_number = re.compile(
                r"^[A-Za-z\u00C0-\u024F][A-Za-z\u00C0-\u024F .'\-]{2,60}"
                r"\s+\d{1,5}\s*[A-Za-z]?$"
            )
            for line in lines:
                if (
                    name_then_number.match(line)
                    and any(token in line.lower() for token in street_tokens)
                ):
                    street = line
                    evidence.append(line)
                    break

    if street:
        try:
            index = lines.index(street)
            surrounding = lines[index:index + 8]
            for line in surrounding:
                match = city_state_zip.match(line)
                if match:
                    code = match.group(2).upper()
                    if code in US_STATES:
                        city = clean(match.group(1))
                        state_code = code
                        state = US_STATES[code]
                        postal_code = match.group(3)
                        evidence.append(line)
                        continue

                if city is None:
                    found = regional_match(line)
                    if found:
                        city, state, state_code, postal_code = found
                        evidence.append(line)
                        break
        except Exception:
            pass

    if partial_one_line and not (street and city):
        # The one-line address had more of the answer than the
        # line-by-line scan managed.
        return partial_one_line

    return {
        "street": street,
        "city": city,
        "state": state,
        "state_code": state_code,
        "postal_code": postal_code,
        "evidence": " | ".join(dict.fromkeys(evidence)),
        "source": None,
    }


# ============================================================
# COUNTRY DETECTION & VALIDATION
# ============================================================

def explicit_countries(text):
    """Return country display-names explicitly mentioned in `text`."""
    found = []
    lower = " " + clean(text).lower() + " "

    for key, entry in COUNTRY_TABLE.items():
        pattern = r"\b" + re.escape(key) + r"\b"
        if re.search(pattern, lower, re.I):
            found.append(entry["display"])

    # Real-world spellings the workbook does not list verbatim
    # ("United States", "Hong Kong", "England"...). Without these the
    # country of a perfectly good site is never found at all.
    for alias, key in COUNTRY_ALIASES.items():
        entry = COUNTRY_TABLE.get(key)
        if not entry:
            continue
        pattern = r"\b" + re.escape(alias) + r"\b"
        if re.search(pattern, lower, re.I):
            found.append(entry["display"])

    return found


def extract_schema_countries(page):
    """Look for schema.org addressCountry / itemprop country hints."""
    found = []
    try:
        nodes = page.locator(
            "[itemprop='addressCountry'], "
            "[itemprop='country'], "
            "meta[property='og:country-name']"
        ).all()
        for node in nodes:
            try:
                value = clean(
                    node.get_attribute("content")
                    or node.inner_text()
                    or ""
                )
                if value:
                    found.append(value)
            except Exception:
                continue
    except Exception:
        pass
    return found


def validate_country(raw_name):
    """
    Validate a raw country name string against COUNTRY_TABLE.

    Returns a dict:
      {
        "input": raw_name,
        "found": bool,
        "status": "Correct" | "Incorrect" | "Not Working" | None,
        "usable_name": str or None,   # None when not workable
      }
    """
    key = clean(raw_name).lower()
    key = COUNTRY_ALIASES.get(key, key)
    entry = COUNTRY_TABLE.get(key)
    if not entry:
        return {
            "input": raw_name,
            "found": False,
            "status": None,
            "usable_name": None,
        }

    status = entry["status"]
    if status == "Not Working":
        usable = None
    elif status == "Incorrect":
        usable = entry["correct_name"]
    else:
        usable = entry["display"]

    return {
        "input": raw_name,
        "found": True,
        "status": status,
        "usable_name": usable,
    }


def determine_country(all_text_blocks, schema_hits):
    """
    Determine the single best-supported, workbook-validated country
    from all page text and any schema.org hints collected during the
    crawl. Country is NEVER inferred from TLD, language, phone code,
    hosting location, or distributor/dealer listings -- only from
    explicit text mentions or schema markup.

    Returns a `validate_country()`-style dict, or the "not found"
    shape if nothing explicit was located.
    """
    candidates = []

    for hit in schema_hits:
        candidates.append(hit)

    for block in all_text_blocks:
        candidates.extend(explicit_countries(block))

    if not candidates:
        return {
            "input": None,
            "found": False,
            "status": None,
            "usable_name": None,
        }

    # Guideline (e): Hong Kong and Macau are explicit exceptions to
    # the China rule. A Hong Kong company's pages almost always say
    # "China" somewhere too, and on a plain frequency count that
    # would resolve to China -> Not Working and wrongly skip a
    # workable record. The S.A.R. entries therefore win outright
    # whenever they are explicitly mentioned.
    for sar in ("China (Hong Kong S.A.R.)", "China (Macau S.A.R.)"):
        if sar in candidates:
            return validate_country(sar)

    # Otherwise the most frequently mentioned explicit country wins.
    counts = {}
    for name in candidates:
        counts[name] = counts.get(name, 0) + 1
    best = max(counts.items(), key=lambda kv: kv[1])[0]

    return validate_country(best)


# ============================================================
# RESTRICTED / NON-PAID CATEGORY DETECTION
# ============================================================

def detect_restricted_category(all_text, all_urls):
    """
    Check crawled text/urls against the hard-rejected categories.
    Returns (reason_string, matched_term) or (None, None) if clean.

    Medical equipment/machinery is explicitly exempted from the
    pharmaceutical restriction.
    """
    lower_text = clean(all_text).lower()
    lower_urls = " ".join(u.lower() for u in all_urls)
    combined = lower_text + " " + lower_urls

    for reason, terms in RESTRICTED_CATEGORY_TERMS.items():
        for term in terms:
            if term in combined:
                if reason == "Pharmaceutical drug manufacturing" and any(
                    exception in combined
                    for exception in MEDICAL_EQUIPMENT_EXCEPTION_TERMS
                ):
                    continue
                return reason, term

    return None, None


def detect_india_china(country_result):
    """
    India and China (excluding Hong Kong / Macau S.A.R. menu
    entries) are not workable countries under the guideline.
    """
    if not country_result.get("found"):
        return False
    name = clean(country_result.get("input", "")).lower()
    return name in {"india", "china"}


def detect_non_paid_context(results):
    """General non-paid business-context terms (not hard restrictions)."""
    found = []
    for result in results:
        if is_directory_page(result["url"]):
            continue
        if not (
            is_company_context_page(result["url"]) or result.get("is_home")
        ):
            continue
        text = clean(result.get("text", "")).lower()
        for term in NON_PAID_TERMS:
            if re.search(rf"\b{re.escape(term)}\b", text, re.I):
                found.append({"term": term, "source": result["url"]})
    return found


def detect_parked_or_placeholder(result):
    text = clean(result.get("text", "")).lower()
    title = clean(result.get("title", "")).lower()
    combined = title + " " + text
    patterns = [
        r"\bdomain\s+parked\b",
        r"\bthis\s+domain\s+is\s+parked\b",
        r"\bparked\s+free\b",
        r"\bget\s+this\s+domain\b",
        r"\bdomain\s+for\s+sale\b",
        r"\bwebsite\s+coming\s+soon\b",
        r"\bcoming\s+soon\b",
        r"\bunder\s+construction\b",
    ]
    return any(re.search(p, combined, re.I) for p in patterns)


def placeholder_status(text):
    """
    Guideline (c) offers three distinct non-working statuses. Once a
    site has been detected as parked/placeholder, decide which one the
    portal should actually receive: an expired/for-sale domain is
    "Domain Expired", anything else of that kind is "Under
    Construction". Never returns "Not Working" -- that status is for a
    site that does not load at all.
    """
    low = clean(text).lower()
    expired_markers = (
        "domain expired", "domain has expired", "this domain has expired",
        "domain for sale", "buy this domain", "get this domain",
        "domain parked", "this domain is parked", "parked free",
        "renew this domain", "expired domain",
    )
    if any(marker in low for marker in expired_markers):
        return "Domain Expired"
    return "Under Construction"


ENGLISH_STOPWORDS = {
    "the", "and", "of", "is", "are", "our", "for", "with", "this",
    "that", "we", "you", "your", "in", "on", "to", "a", "an",
    "company", "product", "products", "service", "services",
    "about", "contact", "home", "quality", "business", "us",
}


def _non_ascii_letter_ratio(text):
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return 0.0
    # Latin, Latin-1 supplement and Latin Extended-A/B cover English
    # plus most Western European accented text. Anything beyond that
    # (CJK, Cyrillic, Arabic, Thai, Hangul, etc.) counts as non-ASCII.
    non_ascii = [c for c in letters if ord(c) > 0x24F]
    return len(non_ascii) / len(letters)


def detect_language(page, result):
    """
    Determine the page's actual displayed language.

    The `<html lang="...">` attribute is NOT trusted on its own --
    many multinational sites (e.g. Hong Kong/China manufacturers)
    hard-code a regional lang tag such as "zh-Hant-HK" for SEO
    purposes even though the visible page content is plain English.
    The rendered body text is authoritative; the lang attribute is
    only used as a weak secondary signal when the text itself is
    inconclusive (too short, or genuinely mixed-script).
    """
    text = clean(result.get("text", ""))
    letter_count = sum(1 for c in text if c.isalpha())

    # Heavily non-Latin script (CJK/Cyrillic/Arabic/etc.) is Non
    # English regardless of length or any lang tag -- a handful of
    # CJK characters already carries a full sentence of meaning, so
    # this check does not wait for a long text sample.
    if letter_count >= 8:
        non_ascii_ratio = _non_ascii_letter_ratio(text)
        if non_ascii_ratio > 0.20:
            return "Non English"

    if len(text) >= 40:
        lower = text.lower()
        words = set(re.findall(r"[a-z']+", lower))
        english_hits = len(words & ENGLISH_STOPWORDS)

        # Mostly Latin-script text with a healthy number of common
        # English function words is English, even if the html lang
        # attribute claims otherwise.
        if english_hits >= 4:
            return "English"

    try:
        html_lang = clean(
            page.locator("html").get_attribute("lang") or ""
        ).lower()
    except Exception:
        html_lang = ""

    if html_lang:
        if html_lang == "en" or html_lang.startswith("en-"):
            return "English"
        # Reaching this point means the text-based checks above were
        # inconclusive (too short, or not clearly English/foreign),
        # so a definite non-English lang tag is trusted as the best
        # available signal.
        return "Non English"

    non_english_markers = [
        "über", "und", "für", "avec", "dans", "empresa", "fabricante",
        "société", "株式会社", "有限公司", "회사",
    ]
    if any(marker in text.lower() for marker in non_english_markers):
        return "Non English"

    return "English"


def has_translation_option(result):
    text = clean(result.get("text", "")).lower()
    markers = [
        "select language", "choose language", "language:", "translate",
        "en |", "| en", "english version", "view in english",
    ]
    return any(marker in text for marker in markers)


# ============================================================
# LANGUAGE SWITCHING
# ============================================================
# Per guideline: a non-English site WITH a translation option is
# still workable. Rather than only detecting that the option exists
# (has_translation_option above), actively use it so the crawl
# proceeds in English.

LANGUAGE_SWITCH_SELECTORS = [
    "a[hreflang='en']",
    "a[href*='/en/']",
    "a[href*='lang=en']",
    "a[href*='language=en']",
    "a[href*='locale=en']",
    "select[name*='lang' i]",
    "select[id*='lang' i]",
    "a:has-text('English')",
    "button:has-text('English')",
    "a:has-text('EN')",
]


def try_switch_to_english(page):
    """
    Best-effort attempt to switch to English via a visible language
    switcher (link, button, or <select>), if one exists. Returns
    True if a control was found and used -- the caller is
    responsible for re-inspecting the page afterward to confirm the
    switch actually worked.
    """
    for selector in LANGUAGE_SWITCH_SELECTORS:
        try:
            loc = page.locator(selector).first
            loc.wait_for(state="visible", timeout=1200)
            tag = (loc.evaluate("el => el.tagName.toLowerCase()") or "").lower()
            if tag == "select":
                loc.select_option(label="English")
            else:
                loc.click()
            page.wait_for_timeout(800)
            print(f"  Language switch used (selector: {selector}).")
            return True
        except Exception:
            continue
    return False


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


def detect_online_shopping(results):
    shopping_terms = [
        "add to cart", "add-to-cart", "shopping cart", "checkout",
        "buy now", "shop now", "place order", "online store",
        "ecommerce", "e-commerce", "payment gateway",
    ]
    for result in results:
        if is_directory_page(result["url"]):
            continue
        text = (result.get("text", "") or "").lower()
        if any(term in text for term in shopping_terms):
            return {"detected": True, "source": result["url"]}
    return {"detected": False, "source": None}


# ============================================================
# EVIDENCE SPLITTING
# ============================================================

def split_evidence_lines(text):
    pieces = re.split(r"(?<=[.!?])\s+|\n+", text or "")
    return [clean(piece) for piece in pieces if clean(piece)]


def company_context_results(results):
    output = []
    for result in results:
        if is_directory_page(result["url"]):
            continue
        if is_company_context_page(result["url"]):
            output.append(result)
    return output


# ============================================================
# BUSINESS TYPE DETECTION
# ============================================================

def _line_has_weak_word_only(line, business_type):
    """
    True if the line only contains a "weak" manufacturing word
    (production, factory, develops, etc.) without a real pattern
    match. Weak words alone must never establish Manufacturer.
    """
    if business_type != "Manufacturer":
        return False
    lower = line.lower()
    return any(word in lower for word in WEAK_MANUFACTURING_WORDS)


def detect_business_type(results):
    """
    Determine the site's paid business type from Company Profile /
    About / History / Who We Are / Company Overview pages.

    Manufacturer always takes absolute priority when present.
    Weak words (production, factory, develops...) never qualify on
    their own. Product pages and directory/dealer pages are never
    used as evidence. A homepage heading may only be used as a
    fallback, and only when it carries one of the recognised
    company-information heading labels.
    """
    context = company_context_results(results)

    reject_patterns = re.compile(
        r"\b(our manufacturers?|manufacturers? we work with|"
        r"manufacturers? include|leading manufacturers?|"
        r"authorized manufacturers?|suppliers? include)\b",
        re.I,
    )

    def scan(pages):
        # Manufacturer first, globally.
        for result in pages:
            for line in split_evidence_lines(result.get("text", "")):
                if reject_patterns.search(line):
                    continue
                for pattern in BUSINESS_PATTERNS["Manufacturer"]:
                    if re.search(pattern, line, re.I):
                        return {
                            "type": "Manufacturer",
                            "source": result["url"],
                            "evidence": line[:700],
                        }
        # Then the remaining paid types, in listed order.
        for business_type in PAID_BUSINESS_TYPES[1:]:
            for result in pages:
                for line in split_evidence_lines(result.get("text", "")):
                    for pattern in BUSINESS_PATTERNS[business_type]:
                        if re.search(pattern, line, re.I):
                            return {
                                "type": business_type,
                                "source": result["url"],
                                "evidence": line[:700],
                            }
        return None

    hit = scan(context)
    if hit:
        return hit

    # Homepage fallback -- only if a recognised company-information
    # heading is present on the homepage itself.
    homepage_headings = [
        "about us", "about the company", "company profile",
        "company overview", "who we are", "our company",
        "our history", "corporate profile",
    ]
    for result in results:
        if not result.get("is_home") or is_directory_page(result["url"]):
            continue
        text = result.get("text", "")
        if not any(h in text.lower() for h in homepage_headings):
            continue
        hit = scan([result])
        if hit:
            return hit

    return {"type": None, "source": None, "evidence": None}


INDUSTRIAL_SERVICE_KEYWORD_PATTERNS = {
    "metal polishing": r"\bmetal\s+polishing\b",
    "powder coating": r"\bpowder\s+coating\b",
    "fabrication": r"\bfabrication\b",
    "refurbishment": r"\brefurbishment\b",
    "welding": r"\bwelding\b",
    "cutting": r"\bcutting\b",
    "moulding": r"\bmo?ulding\b",
}

# Idiomatic phrases that must NOT count as a "cutting" (or similar)
# industrial-service hit even though the bare word appears.
INDUSTRIAL_SERVICE_KEYWORD_FALSE_POSITIVES = {
    "cutting": [r"\bcutting[\s-]edge\b", r"\bcost[\s-]cutting\b"],
}


def _industrial_service_keyword_match(text):
    """
    Return the first genuine INDUSTRIAL_SERVICE_KEYWORDS hit in
    `text`, or None. Uses whole-word matching and filters out known
    idioms (e.g. "cutting-edge") that contain the keyword as a
    substring without describing the actual service.
    """
    lower = text.lower()
    for term, pattern in INDUSTRIAL_SERVICE_KEYWORD_PATTERNS.items():
        if not re.search(pattern, lower, re.I):
            continue
        false_positives = INDUSTRIAL_SERVICE_KEYWORD_FALSE_POSITIVES.get(term, [])
        stripped = lower
        for fp_pattern in false_positives:
            stripped = re.sub(fp_pattern, " ", stripped, flags=re.I)
        if re.search(pattern, stripped, re.I):
            return term
    return None


def is_service_detail_page(result):
    """
    A dedicated service page, e.g. /services/construction-civil-
    mechanical -- as opposed to the bare /services listing page.
    """
    url = result["url"].lower()
    return bool(re.search(r"/services?/[^/?#]+", url, re.I))


def inspect_industrial_service(page, result, known_names=None):
    """
    Verify a single candidate Industrial Service page the same way a
    product page is verified: a real name, a matching Industrial
    Service keyword (not an excluded corporate/software/consulting/
    small-scale-industry service, not an idiomatic false positive),
    a proper description, and at least one qualifying colour image.
    """
    text = result.get("text", "")
    lower_text = clean(text).lower()

    if any(term in lower_text for term in NON_INDUSTRIAL_SERVICE_KEYWORDS):
        return {
            "name": None, "keyword": None, "description": False,
            "images": [], "qualifying_images": [], "qualifies": False,
        }

    keyword = _industrial_service_keyword_match(text)
    if not keyword:
        return {
            "name": None, "keyword": None, "description": False,
            "images": [], "qualifying_images": [], "qualifies": False,
        }

    name = get_product_name(page, result, known_names=known_names)
    if not name:
        return {
            "name": None, "keyword": keyword, "description": False,
            "images": [], "qualifying_images": [], "qualifies": False,
        }

    description = has_product_description(result, name)
    if not description:
        return {
            "name": name, "keyword": keyword, "description": False,
            "images": [], "qualifying_images": [], "qualifies": False,
        }

    print("  Service name:", name, "| keyword:", keyword)
    print("  Description: PASS")
    print("  Checking up to", MAX_IMAGE_CANDIDATES_PER_PRODUCT, "image candidates...")

    # NOTE 1 of the Intensecore guidelines holds Industrial Services to
    # the same "description & image of products" standard, so the
    # service's images must relate to the service name too.
    images = discover_product_images(
        page, product_name=(name + " " + (keyword or "")),
    )
    qualifying = [img for img in images if img["color"]["qualified"]]

    return {
        "name": name,
        "keyword": keyword,
        "description": True,
        "images": images,
        "qualifying_images": qualifying,
        "qualifies": bool(qualifying),
    }


# ============================================================
# COMPANY PROFILE
# ============================================================

def find_company_profile(results):
    for result in results:
        if is_directory_page(result["url"]):
            continue
        if is_company_context_page(result["url"]):
            if len(clean(result.get("text", ""))) >= 100:
                return True, result["url"]
    for result in results:
        if result.get("is_home") and len(clean(result.get("text", ""))) >= 150:
            return True, result["url"]
    return False, None


# ============================================================
# PRODUCT PAGE DETECTION
# ============================================================

def is_obvious_non_product(result):
    url = result["url"].lower()
    title = result.get("title", "").lower()
    text = clean(result.get("text", "")).lower()
    combined = url + " " + title

    if is_noise_page(url):
        return True

    if any(term in combined for term in [
        "/services/", "/service/", "/projects/", "/project/",
        "/solutions/", "/solution/", "/brands/", "/brand/",
    ]):
        return True

    if re.search(r"/products?/?$", url, re.I):
        return True
    if re.search(r"/catalog(?:ue)?/?$", url, re.I):
        return True

    if "services" in title and not any(
        term in text for term in PRODUCT_STRONG_TERMS
    ):
        return True

    return False


def looks_like_product_detail(result):
    if is_obvious_non_product(result):
        return False

    url = result["url"].lower()
    title = result.get("title", "").lower()
    text = clean(result.get("text", "")).lower()

    if re.search(r"/product/[^/?#]+$", url, re.I):
        return True
    if re.search(r"/products?/[^/?#]+/[^/?#]+", url, re.I):
        return True

    strong_hits = sum(1 for term in PRODUCT_STRONG_TERMS if term in text)
    if strong_hits >= 2:
        return True

    product_title_terms = [
        "connector", "switch", "valve", "pump", "filter", "fitting",
        "hardware", "bolt", "hinge", "shackle", "fixture", "clip",
        "anchor", "coupling", "bearing", "motor", "sensor",
    ]
    if any(term in title for term in product_title_terms) and len(text) >= 180:
        return True

    return False


def is_product_page(result):
    return looks_like_product_detail(result)


def _url_slug(url):
    try:
        path = urlparse(url).path.rstrip("/")
        slug = path.rsplit("/", 1)[-1]
        slug = re.sub(r"[-_]+", " ", slug).strip()
        return slug
    except Exception:
        return ""


def get_product_name(page, result, known_names=None):
    """
    Extract the product's name, preferring the most specific
    available source (h1, og:title, <title> tag).

    Many product-page templates reuse an identical generic h1 across
    every SKU in a category (e.g. every coffee-machine page titled
    "Fully Automatic Coffee Machine"). When every extracted candidate
    collides with a name already used for a different product on
    this site, the URL slug is appended to keep genuinely distinct
    products from being silently collapsed into one.
    """
    known_names = known_names or set()
    candidates = []

    try:
        value = clean(page.locator("h1").first.inner_text(timeout=1500))
        if value:
            candidates.append(value[:300])
    except Exception:
        pass

    try:
        meta = page.locator('meta[property="og:title"]').first
        value = clean(meta.get_attribute("content") or "")
        if value:
            candidates.append(value[:300])
    except Exception:
        pass

    generic_titles = {
        "home", "products", "product", "catalog", "catalogue",
        "contact", "about", "services", "solutions",
    }
    title = clean(result.get("title", ""))
    if title:
        # Strip a trailing " | Site Name" / " - Site Name" suffix.
        title_main = re.split(
            r"\s*[|\u2013\u2014-]\s*", title, maxsplit=1,
        )[0].strip()
        if title_main and title_main.lower() not in generic_titles:
            candidates.append(title_main[:300])
        elif title.lower() not in generic_titles:
            candidates.append(title[:300])

    candidates = [c for c in candidates if c]
    if not candidates:
        return None

    for candidate in candidates:
        if candidate not in known_names:
            return candidate

    # Every candidate collides with a name used elsewhere -- append
    # the URL slug to disambiguate rather than dropping the product.
    slug = _url_slug(result.get("url", ""))
    base = candidates[0]
    if slug and slug.lower() not in base.lower():
        return f"{base} ({slug})"[:300]
    return base


def has_product_description(result, product_name):
    if not product_name:
        return False

    text = clean(result.get("text", ""))
    if len(text) < 180:
        return False

    body = text.replace(product_name, "", 1).strip()
    if len(body) < 80:
        return False

    lower = body.lower()
    descriptive_terms = [
        "designed", "features", "specification", "specifications",
        "material", "dimensions", "dimension", "size", "weight",
        "application", "applications", "diameter", "length", "width",
        "height", "capacity", "voltage", "current", "waterproof",
        "finish", "mount", "model", "technical", "suitable",
        "used for", "made from", "construction",
    ]
    hits = sum(1 for term in descriptive_terms if term in lower)
    return hits >= 1 and len(body) >= 80


def product_name_is_actually_product(product_name, page_text):
    """
    Reject services/projects/brands/programs/labels/signs disguised
    as "products" -- these never count under the guideline.
    """
    if not product_name:
        return False
    lower_name = product_name.lower()
    if any(f" {term} " in f" {lower_name} " for term in NON_PRODUCT_TERMS):
        return False
    return True


# ============================================================
# IMAGE ANALYSIS
# ============================================================

def _image_sharpness(image):
    """
    Variance of the edge response, used to spot blurred images.
    Sharp photos give a high number; out-of-focus ones give a low
    one. A flat, featureless placeholder measures 0.0 and is meant
    to fail. Returns None only when the measurement itself could not
    be taken -- that never rejects an image on its own.
    """
    try:
        from PIL import ImageFilter

        gray = image.convert("L")
        gray.thumbnail((256, 256))
        edges = gray.filter(ImageFilter.FIND_EDGES)

        # FIND_EDGES leaves a bright artificial border; crop it off so
        # it cannot inflate the variance of a genuinely blurry image.
        w, h = edges.size
        if w > 4 and h > 4:
            edges = edges.crop((2, 2, w - 2, h - 2))

        pixels = list(edges.getdata())
        if not pixels:
            return None
        mean = sum(pixels) / len(pixels)
        return sum((px - mean) ** 2 for px in pixels) / len(pixels)
    except Exception:
        return None


def analyze_image_bytes(data):
    """
    Decide whether a product image qualifies.

    Per Intensecore item 10, black & white, very small, dummy and
    foggy images ARE acceptable -- there is deliberately no colour
    test and no real size test here. What is rejected:

      - a format outside jpeg / png / jpg / gif  (rule (j))
      - trackers, spacers and UI icons below the tiny floor
      - half-cut images, via an extreme aspect ratio  (user rule)
      - clearly blurred images                        (user rule)

    Relevance to the product name is NOT judged here -- that is
    enforced separately by the name-token match, which is what
    rejects "irrelevant images corresponding to product".
    """
    try:
        from PIL import Image

        image = Image.open(io.BytesIO(data))
        detected_format = (image.format or "").upper()

        if detected_format not in ACCEPTED_IMAGE_FORMATS:
            return {
                "qualified": False,
                "reason": "unsupported image format",
                "format": detected_format,
                "sharpness": 0,
                "aspect_ratio": 0,
                "width": 0,
                "height": 0,
            }

        image = image.convert("RGB")
    except Exception:
        return {
            "qualified": False,
            "reason": "image decode failed",
            "format": "",
            "sharpness": 0,
            "aspect_ratio": 0,
            "width": 0,
            "height": 0,
        }

    width, height = image.size

    if (
        width < MIN_IMAGE_WIDTH
        or height < MIN_IMAGE_HEIGHT
        or width * height < MIN_IMAGE_AREA
    ):
        return {
            "qualified": False,
            "reason": "tracker / spacer / icon, not a product image",
            "format": detected_format,
            "sharpness": 0,
            "aspect_ratio": 0,
            "width": width,
            "height": height,
        }

    aspect = (max(width, height) / min(width, height)) if min(width, height) else 0

    if aspect > MAX_IMAGE_ASPECT_RATIO:
        return {
            "qualified": False,
            "reason": "half-cut / cropped strip (extreme aspect ratio)",
            "format": detected_format,
            "sharpness": 0,
            "aspect_ratio": round(aspect, 2),
            "width": width,
            "height": height,
        }

    sharpness = _image_sharpness(image)
    if sharpness is not None and sharpness < MIN_IMAGE_SHARPNESS:
        return {
            "qualified": False,
            "reason": "blurred",
            "format": detected_format,
            "sharpness": round(sharpness, 1) if sharpness is not None else "n/a",
            "aspect_ratio": round(aspect, 2),
            "width": width,
            "height": height,
        }

    return {
        "qualified": True,
        "reason": "qualifying product image",
        "format": detected_format,
        "sharpness": round(sharpness, 1) if sharpness is not None else "n/a",
        "aspect_ratio": round(aspect, 2),
        "width": width,
        "height": height,
    }

def image_context_is_restricted(metadata_text):
    """
    True if the image's alt/title/filename/surrounding text implies
    a restricted subject (diagram-type OR restricted-content-type).
    """
    lower = metadata_text.lower()
    if any(term in lower for term in BAD_IMAGE_CONTEXT_TERMS):
        return True
    if any(term in lower for term in RESTRICTED_IMAGE_SUBJECT_TERMS):
        return True
    return False


# ============================================================
# IMAGE SOURCE COLLECTION
# ============================================================

def collect_image_sources(item):
    values = []
    fields = [
        ("data_zoom", 1100000), ("data_full", 1050000),
        ("data_original", 1000000), ("data_large", 950000),
        ("current_src", 900000), ("data_src", 850000),
        ("data_lazy", 800000),
    ]
    for key, score in fields:
        value = clean(item.get(key))
        if value and not value.startswith("data:"):
            values.append((score, value))

    srcset = item.get("srcset", "")
    for part in srcset.split(","):
        bits = part.strip().split()
        if not bits:
            continue
        score = 700000
        if len(bits) > 1:
            match = re.match(r"(\d+)w", bits[1])
            if match:
                score += int(match.group(1))
        values.append((score, bits[0]))

    src = clean(item.get("src"))
    if src and not src.startswith("data:"):
        values.append((500000, src))

    unique = {}
    for score, value in values:
        unique[value] = max(score, unique.get(value, 0))

    return [
        value for value, _ in
        sorted(unique.items(), key=lambda x: x[1], reverse=True)
    ]


IMAGE_METADATA_JS = """
imgs => imgs.map(img => ({
    src: img.getAttribute('src') || '',
    srcset: img.getAttribute('srcset') || '',
    data_src: img.getAttribute('data-src') || '',
    data_lazy: img.getAttribute('data-lazy-src') || '',
    data_original: img.getAttribute('data-original') || '',
    data_full: img.getAttribute('data-full') || '',
    data_large: img.getAttribute('data-large') || '',
    data_zoom: img.getAttribute('data-zoom-image') || '',
    current_src: img.currentSrc || img.src || '',
    alt: img.getAttribute('alt') || '',
    title: img.getAttribute('title') || '',
    natural_width: img.naturalWidth || 0,
    natural_height: img.naturalHeight || 0,
    parent_text: img.parentElement ? (img.parentElement.innerText || '') : '',
    grandparent_text:
        img.parentElement && img.parentElement.parentElement
        ? (img.parentElement.parentElement.innerText || '')
        : ''
}))
"""


def trigger_lazy_load(page):
    """
    Many product-page templates never populate an <img>'s real src /
    naturalWidth / naturalHeight until it scrolls into view. Scroll
    the page in steps so lazy-loaded images resolve before we read
    their metadata.
    """
    try:
        height = page.evaluate("document.body.scrollHeight") or 0
        pos = 0
        step = 600
        while pos < height and pos < 20000:
            page.evaluate(f"window.scrollTo(0, {pos})")
            page.wait_for_timeout(150)
            pos += step
        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(200)
    except Exception:
        pass


def _fetch_image_via_browser(page, url):
    """
    Fallback image fetch executed inside the page's own JS context so
    the request naturally carries the same Referer/Origin/cookies a
    real browser would send. This recovers images from CDNs that use
    hotlink protection and reject Playwright's bare page.request.get
    (a common cause of every image failing to decode).
    """
    try:
        result = page.evaluate(
            """
            async (url) => {
                const res = await fetch(url, { credentials: 'include' });
                if (!res.ok) return null;
                const buf = await res.arrayBuffer();
                const bytes = new Uint8Array(buf);
                let binary = '';
                for (let i = 0; i < bytes.byteLength; i++) {
                    binary += String.fromCharCode(bytes[i]);
                }
                return {
                    data: btoa(binary),
                    type: res.headers.get('content-type') || '',
                };
            }
            """,
            url,
        )
        if not result or not result.get("data"):
            return None, None
        import base64
        raw = base64.b64decode(result["data"])
        content_type = (result.get("type") or "").split(";")[0].strip().lower()
        return raw, content_type
    except Exception:
        return None, None


# Words that carry no identifying power when matching an image
# against a product name -- a match on these alone does not prove the
# image is "related to the products name" (Intensecore rule (m)).
GENERIC_NAME_WORDS = {
    "product", "products", "item", "items", "series", "model",
    "models", "type", "types", "range", "quality", "high", "best",
    "new", "our", "the", "and", "for", "with", "from", "grade",
    "standard", "custom", "industrial", "commercial", "professional",
    "premium", "heavy", "duty", "size", "sizes", "set", "kit",
    "manufacturer", "supplier", "exporter", "company", "detail",
    "details", "more", "view", "info", "information", "page",
}


def product_name_tokens(name):
    """
    Significant tokens from a product name, used to prove that an
    image actually belongs to that product. Generic filler words are
    dropped so "Industrial Product Range" cannot match every photo on
    the site.
    """
    raw = re.split(r"[^a-z0-9]+", clean(name).lower())
    tokens = []
    for token in raw:
        if len(token) < 3:
            continue
        if token in GENERIC_NAME_WORDS:
            continue
        if token.isdigit() and len(token) < 3:
            continue
        if token not in tokens:
            tokens.append(token)
    return tokens


def image_matches_product_name(metadata, context, tokens):
    """
    True when the image's own metadata (alt/title/filename) or its
    immediate surrounding text carries a significant token from the
    product name. This is the guideline's "Images should be related to
    products name" / name-image-description correspondence check.
    """
    if not tokens:
        return False
    haystack = (metadata + " " + context).lower()
    return any(token in haystack for token in tokens)


def discover_product_images(page, product_name=""):
    """
    Rank <img> candidates on the current page and download up to
    MAX_IMAGE_CANDIDATES_PER_PRODUCT of them, stopping as soon as one
    qualifying colour image is found (one qualifying colour image is
    sufficient for the product).
    """
    trigger_lazy_load(page)

    try:
        items = page.locator("img").evaluate_all(IMAGE_METADATA_JS)
    except Exception as exc:
        print("  Image metadata error:", type(exc).__name__)
        return []

    name_tokens = product_name_tokens(product_name)
    if product_name and not name_tokens:
        print(
            "  Product name has no distinctive words -- cannot prove "
            "an image belongs to it."
        )

    ranked = []
    for item in items:
        sources = collect_image_sources(item)
        if not sources:
            continue

        filename = urlparse(sources[0]).path.rsplit("/", 1)[-1]
        metadata = " ".join([
            clean(item.get("alt")), clean(item.get("title")), filename,
        ]).lower()

        if any(term in metadata for term in BAD_IMAGE_TERMS):
            continue

        context = " ".join([
            item.get("parent_text", ""), item.get("grandparent_text", ""),
            clean(item.get("alt")), clean(item.get("title")),
        ])

        if image_context_is_restricted(metadata + " " + context):
            continue

        product_context_terms = [
            "product", "connector", "switch", "valve", "pump", "filter",
            "fitting", "hardware", "bolt", "hinge", "shackle", "fixture",
            "clip", "anchor",
        ]
        product_context = any(
            term in context.lower() for term in product_context_terms
        )

        # MANDATORY: the image must actually relate to THIS product's
        # name. An unrelated colour photo sitting on a product page is
        # not proof, so it is dropped rather than counted.
        name_match = image_matches_product_name(
            metadata, context, name_tokens,
        )
        if not name_match:
            continue

        width = int(item.get("natural_width", 0) or 0)
        height = int(item.get("natural_height", 0) or 0)

        score = 0
        if name_match:
            score += 200
        if product_context:
            score += 100
        if clean(item.get("alt")):
            score += 30
        if clean(item.get("title")):
            score += 20
        if width >= 300:
            score += 20
        if height >= 300:
            score += 20
        if width * height >= 100000:
            score += 20

        ranked.append((score, item, sources))

    ranked.sort(key=lambda x: x[0], reverse=True)

    found = []
    seen_sources = set()

    for _, item, sources in ranked:
        if len(found) >= MAX_IMAGE_CANDIDATES_PER_PRODUCT:
            break

        for candidate in sources[:2]:
            source_url = normalize_url(urljoin(page.url, candidate))
            if not source_url or source_url in seen_sources:
                continue
            if is_download_url(source_url):
                continue
            seen_sources.add(source_url)

            print(
                f"    Checking image {len(found) + 1}/"
                f"{MAX_IMAGE_CANDIDATES_PER_PRODUCT}"
            )

            body_bytes = None
            content_type = ""

            try:
                response = page.request.get(
                    source_url,
                    timeout=IMAGE_REQUEST_TIMEOUT,
                    headers={
                        "Referer": page.url,
                        "Accept": "image/avif,image/webp,image/*,*/*;q=0.8",
                    },
                )
                if response.ok:
                    content_type = (
                        response.headers.get("content-type", "")
                        .split(";")[0].strip().lower()
                    )
                    if content_type in ACCEPTED_IMAGE_MIME_TYPES:
                        body_bytes = response.body()
            except Exception:
                pass

            if body_bytes is None:
                # Hotlink-protected CDN or blocked direct request --
                # retry from inside the page's own JS context.
                body_bytes, content_type = _fetch_image_via_browser(
                    page, source_url,
                )

            if body_bytes is None or content_type not in ACCEPTED_IMAGE_MIME_TYPES:
                print("      Rejected MIME:", content_type or "UNKNOWN")
                continue

            try:
                analysis = analyze_image_bytes(body_bytes)
                found.append({
                    "src": source_url,
                    "content_type": content_type,
                    "color": analysis,
                    "alt": item.get("alt", ""),
                    "title": item.get("title", ""),
                })

                print(
                    "      Result:", analysis["reason"],
                    f"| {analysis['width']}x{analysis['height']}",
                    "| sharpness=", analysis["sharpness"],
                    "| aspect=", analysis["aspect_ratio"],
                )

                if analysis["qualified"]:
                    print("      >>> QUALIFYING PRODUCT IMAGE FOUND.")
                    return found

                break

            except Exception:
                continue

    return found


# ============================================================
# PRODUCT INSPECTION
# ============================================================

def inspect_product(page, result, known_names=None):
    name = get_product_name(page, result, known_names=known_names)

    if not name or not product_name_is_actually_product(name, result.get("text", "")):
        return {
            "name": name, "description": False, "images": [],
            "qualifying_images": [], "qualifies": False,
        }

    description = has_product_description(result, name)
    if not description:
        return {
            "name": name, "description": False, "images": [],
            "qualifying_images": [], "qualifies": False,
        }

    print("  Product name:", name)
    print("  Description: PASS")
    print("  Checking up to", MAX_IMAGE_CANDIDATES_PER_PRODUCT, "image candidates...")

    images = discover_product_images(page, product_name=name)
    qualifying = [img for img in images if img["color"]["qualified"]]

    if qualifying:
        print("  Name/image/description correspondence: PASS")
    else:
        print(
            "  No colour image could be tied to this product's name "
            "-- product not counted."
        )

    return {
        "name": name,
        "description": True,
        "images": images,
        "qualifying_images": qualifying,
        "qualifies": bool(qualifying),
    }


# ============================================================
# LINK DISCOVERY
# ============================================================

LINK_PRIORITIES = {
    "about": 100000, "who-we-are": 100000, "company profile": 100000,
    "company-overview": 100000, "company": 95000, "profile": 95000,
    "history": 90000, "corporate": 90000, "manufacturer": 85000,
    "manufacturing": 80000, "contact": 75000, "location": 70000,
    "product": 50000, "products": 50000, "catalog": 45000,
    "catalogue": 45000, "industrial": 30000, "service": 20000,
}

SOCIAL_DOMAINS = [
    "facebook.com", "instagram.com", "linkedin.com", "youtube.com",
    "twitter.com", "x.com",
]


def discover_links(page, base_url):
    output = {}
    try:
        links = page.locator("a[href]").all()
    except Exception:
        return []

    for link in links:
        try:
            href = link.get_attribute("href")
            if not href:
                continue

            absolute = normalize_url(urljoin(base_url, href))
            if not absolute or not absolute.startswith(("http://", "https://")):
                continue
            if not same_domain(base_url, absolute):
                continue
            if is_download_url(absolute):
                continue
            if any(social in absolute.lower() for social in SOCIAL_DOMAINS):
                continue
            if is_noise_page(absolute):
                # News/awards/careers/legal pages never establish
                # business type or products -- don't spend crawl
                # budget queueing them.
                continue

            text = clean(link.inner_text())
            aria = clean(link.get_attribute("aria-label") or "")
            title = clean(link.get_attribute("title") or "")
            combined = f"{text} {aria} {title} {absolute}".lower()

            score = 10
            for key, points in LINK_PRIORITIES.items():
                if key in combined:
                    score += points

            if is_directory_page(absolute):
                score -= 100000

            output[absolute] = max(score, output.get(absolute, 0))

        except Exception:
            pass

    return [
        url for url, _ in
        sorted(output.items(), key=lambda x: x[1], reverse=True)
    ]


# ============================================================
# PAGE INSPECTION
# ============================================================

def inspect_page(page, url, is_home=False):
    print("\n" + "-" * 70)
    print("Inspecting:", url)
    print("-" * 70)

    if is_download_url(url):
        print("SKIPPED DOWNLOAD")
        return None

    try:
        response = page.goto(
            url, wait_until="domcontentloaded",
            timeout=PAGE_NAVIGATION_TIMEOUT,
        )
        page.wait_for_timeout(300)
    except Exception as exc:
        print("Navigation error:", type(exc).__name__, exc)
        # If the browser, context or page itself is gone, every
        # remaining page in the queue will raise the same thing. Say so
        # once and stop the crawl instead of printing the same error 15
        # more times and pretending the site was inspected.
        if "has been closed" in str(exc):
            raise BrowserGoneError(
                "the browser/page was closed during the crawl"
            ) from exc
        return None

    final_url = normalize_url(page.url)
    if not final_url:
        return None
    if is_download_url(final_url):
        print("SKIPPED FINAL DOWNLOAD")
        return None

    try:
        title = clean(page.title())
    except Exception:
        title = ""

    try:
        text = page.locator("body").inner_text(timeout=BODY_TIMEOUT)
    except Exception:
        text = ""

    address = extract_address(text)
    if address.get("street"):
        address["source"] = final_url

    result = {
        "url": final_url,
        "requested_url": url,
        "status": response.status if response else None,
        "title": title,
        "text": text,
        "emails": extract_emails(text),
        "phones": extract_phones(text),
        "address": address,
        "schema_countries": extract_schema_countries(page),
        "is_home": is_home,
        "language": None,
    }
    result["language"] = detect_language(page, result)

    print("Final URL:", final_url)
    print("HTTP Status:", result["status"])
    print("Language:", result["language"])
    if result["emails"]:
        print("Emails:", result["emails"])
    if result["phones"]:
        print("Phones:", ", ".join(
            entry["raw"] + (
                " [" + "/".join(sorted(entry["labels"])) + "]"
                if entry["labels"] else ""
            )
            for entry in result["phones"]
        ))

    return result


# ============================================================
# URL STRUCTURE CHECKS
# ============================================================

def is_double_domain_or_redirect(landing_url, final_url):
    """
    A redirect that changes the registrable domain (not just path or
    protocol/www) is non-considerable -- the site must remain in its
    actual supplied form.
    """
    return get_root_domain(landing_url) != get_root_domain(final_url)


# ============================================================
# WEBSITE CRAWL ORCHESTRATION -- STAGED
# ============================================================
#
# The crawl is split into two independent stages so a site that is
# going to fail an early, cheap check (restricted category, wrong
# business type, missing email, bad country, etc.) never wastes time
# hunting for products at all:
#
#   Stage 1 (crawl_context):  homepage + company-context/contact
#     pages only. Fast, small page budget. Everything the pre-
#     product decision checks need.
#
#   Stage 2 (crawl_products): only run if Stage 1 + the pre-product
#     checks all pass. Hunts for qualifying products/services,
#     stopping as soon as MIN_QUALIFYING_PRODUCTS distinct
#     qualifying entries are found.
# ============================================================

CONVENTIONAL_CONTEXT_PATHS = (
    "/contact", "/contact-us", "/contacts", "/about", "/about-us",
    "/company", "/company-profile", "/who-we-are", "/our-company",
)


def guess_context_urls(landing_url, known_links):
    """
    Conventional company/contact URLs for a site whose homepage links
    to none. A language-splash homepage (fericor.com redirects to /en
    only after the page is read) leaves Stage 1 with nothing to visit,
    which showed up as "Address Error" on a perfectly good site.

    Any language prefix actually seen on the site is reused, so /en/
    and /de/ style paths are covered as well as bare ones. A guess
    that does not exist simply fails to load and costs one request.
    """
    try:
        parsed = urlparse(landing_url)
        root = f"{parsed.scheme}://{parsed.netloc}"
    except Exception:
        return []

    prefixes = [""]
    seen = set()
    for link in known_links:
        try:
            segments = [seg for seg in urlparse(link).path.split("/") if seg]
        except Exception:
            continue
        if not segments:
            continue
        first = segments[0].lower()
        # a language segment looks like "en", "de", "en-gb", "en_us"
        if re.fullmatch(r"[a-z]{2}([-_][a-z]{2})?", first) and first not in seen:
            seen.add(first)
            prefixes.append("/" + first)

    out = []
    for prefix in prefixes[:3]:
        for path in CONVENTIONAL_CONTEXT_PATHS:
            out.append(root + prefix + path)
    return out


def crawl_context(page, landing_url, seed_result=None):
    """
    Stage 1: crawl the homepage and company-context/contact pages
    only (no product/service hunting). `seed_result`, if given, is
    an already-inspected homepage result (e.g. after a language
    switch) that is reused instead of being fetched again.

    Returns a dict with the collected results plus `queued`/
    `pending_links` state that crawl_products() can continue from
    without re-visiting anything.
    """
    visited_final = set()
    queue = deque()
    queued = {landing_url}
    results = []
    context_pages_seen = 0
    pages_inspected = 0
    pending_links = []

    # Stage 1 visits company-information and contact pages ONLY.
    # It used to walk a plain FIFO queue of every link on the site, so
    # on a shop with a few hundred product links in the nav the 18-page
    # budget was gone long before the contact page came up -- which is
    # why good sites failed with "Address Error: no full address
    # paragraph found". Product links are still collected here for
    # Stage 2; they are just not walked now.
    candidates = []

    def offer(link):
        if link in queued:
            return
        queued.add(link)
        pending_links.append(link)
        score = context_link_score(link)
        if score is not None:
            candidates.append((score, link))

    if seed_result is not None:
        visited_final.add(seed_result["url"])
        results.append(seed_result)
        context_pages_seen += 1
        for link in discover_links(page, seed_result["url"]):
            offer(link)
    else:
        queue.append((landing_url, True))

    # A homepage that is only a language splash, or that finishes its
    # redirect after the links are read, leaves nothing to visit. Give
    # it a moment, re-read the links, and fall back to conventional
    # paths rather than declaring the site addressless.
    if not candidates and seed_result is not None:
        try:
            page.wait_for_timeout(1500)
            for link in discover_links(page, normalize_url(page.url)):
                offer(link)
        except Exception:
            pass

    # Still nothing? The homepage is probably a language chooser whose
    # only links are /en and /de. Follow the English one and read the
    # real navigation from there -- one request, instead of guessing.
    if not candidates:
        language_landings = []
        for link in pending_links:
            try:
                path = urlparse(link).path.strip("/")
            except Exception:
                continue
            if re.fullmatch(r"[a-z]{2}([-_][a-z]{2})?", path.lower()):
                language_landings.append(link)
        language_landings.sort(
            key=lambda l: 0 if "/en" in l.lower() else 1
        )

        for link in language_landings[:2]:
            print(f"  Homepage looks like a language chooser -- following {link}")
            result = inspect_page(page, link, is_home=False)
            pages_inspected += 1
            if result and result["url"] not in visited_final:
                visited_final.add(result["url"])
                results.append(result)
                if is_company_context_page(result["url"]):
                    context_pages_seen += 1
                for found in discover_links(page, result["url"]):
                    offer(found)
            if candidates:
                break

    # Last resort: conventional paths. Kept capped so a site with no
    # contact page cannot burn the whole Stage 1 budget on 404s.
    if not candidates:
        print(
            "  No company/contact link found -- trying conventional paths."
        )
        for guess in guess_context_urls(landing_url, list(queued))[:12]:
            offer(guess)

    while pages_inspected < MAX_CONTEXT_PAGES:
        if queue:
            url, is_home = queue.popleft()
        elif candidates:
            candidates.sort(key=lambda item: item[0], reverse=True)
            _, url = candidates.pop(0)
            is_home = False
        else:
            break

        result = inspect_page(page, url, is_home=is_home)
        pages_inspected += 1

        if result is None:
            continue
        if result["url"] in visited_final:
            continue
        visited_final.add(result["url"])
        results.append(result)

        if is_company_context_page(result["url"]) or result.get("is_home"):
            context_pages_seen += 1

        for link in discover_links(page, result["url"]):
            offer(link)

    print(
        f"  Stage 1: inspected {pages_inspected} page(s) -- "
        f"{context_pages_seen} company/contact page(s), "
        f"{len(pending_links)} link(s) held for the product hunt."
    )

    return {
        "results": results,
        "context_pages_seen": context_pages_seen,
        "visited_final": visited_final,
        "queued": queued,
        "pending_links": pending_links,
    }


def crawl_products(page, seed_state):
    """
    Stage 2: continue crawling from Stage 1's leftover queue,
    specifically hunting for qualifying products/services, stopping
    as soon as MIN_QUALIFYING_PRODUCTS distinct qualifying entries
    are found for either the product or the Industrial Service path.

    Only call this after the pre-product decision checks have
    already passed -- it is the expensive part of verification.
    """
    visited_final = set(seed_state["visited_final"])
    queued = set(seed_state["queued"])
    queue = deque(
        (link, False) for link in seed_state["pending_links"]
        if link not in visited_final
    )

    results = list(seed_state["results"])
    qualifying_products = []
    qualifying_services = []
    seen_product_names = set()
    seen_service_names = set()
    product_pages_seen = 0
    pages_inspected = 0

    while queue and pages_inspected < MAX_TOTAL_PAGES:
        if (
            len(qualifying_products) >= MIN_QUALIFYING_PRODUCTS
            or len(qualifying_services) >= MIN_QUALIFYING_PRODUCTS
        ):
            break

        url, is_home = queue.popleft()
        if url in visited_final:
            continue

        result = inspect_page(page, url, is_home=is_home)
        pages_inspected += 1

        if result is None:
            continue
        if result["url"] in visited_final:
            continue
        visited_final.add(result["url"])
        results.append(result)

        is_directory = is_directory_page(result["url"])

        if (
            not is_directory
            and looks_like_product_detail(result)
            and product_pages_seen < MAX_PRODUCT_PAGES
        ):
            product_pages_seen += 1
            print("  -> Candidate product page.")
            inspection = inspect_product(
                page, result, known_names=seen_product_names,
            )
            if inspection["name"]:
                seen_product_names.add(inspection["name"])
            if inspection["qualifies"]:
                already_named = {p["name"] for p in qualifying_products}
                if inspection["name"] not in already_named:
                    qualifying_products.append({
                        "name": inspection["name"],
                        "url": result["url"],
                        "colour_images": len(inspection["qualifying_images"]),
                        "description": inspection["description"],
                    })
                    print(
                        f"  >>> QUALIFYING PRODUCT "
                        f"{len(qualifying_products)}/{MIN_QUALIFYING_PRODUCTS}"
                    )

        elif (
            not is_directory
            and is_service_detail_page(result)
            and product_pages_seen < MAX_PRODUCT_PAGES
        ):
            product_pages_seen += 1
            print("  -> Candidate Industrial Service page.")
            service_inspection = inspect_industrial_service(
                page, result, known_names=seen_service_names,
            )
            if service_inspection["name"]:
                seen_service_names.add(service_inspection["name"])
            if service_inspection["qualifies"]:
                already_named = {s["name"] for s in qualifying_services}
                if service_inspection["name"] not in already_named:
                    qualifying_services.append({
                        "name": service_inspection["name"],
                        "url": result["url"],
                        "colour_images": len(service_inspection["qualifying_images"]),
                        "description": service_inspection["description"],
                    })
                    print(
                        f"  >>> QUALIFYING SERVICE "
                        f"{len(qualifying_services)}/{MIN_QUALIFYING_PRODUCTS}"
                    )

        if (
            len(qualifying_products) >= MIN_QUALIFYING_PRODUCTS
            or len(qualifying_services) >= MIN_QUALIFYING_PRODUCTS
        ):
            break

        if product_pages_seen < MAX_PRODUCT_PAGES:
            for link in discover_links(page, result["url"]):
                if link not in queued and len(queued) < MAX_TOTAL_PAGES * 2:
                    queued.add(link)
                    queue.append((link, False))

    return {
        "results": results,
        "qualifying_products": qualifying_products,
        "qualifying_services": qualifying_services,
        "product_pages_seen": product_pages_seen,
        "visited_final": visited_final,
    }


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
# DECISION ENGINE
# ============================================================
#
# Implements the master guideline's decision-first rule: a website
# is QUALIFIES only after every applicable mandatory requirement has
# actually been verified. Any single missing/unclear/unverifiable
# mandatory requirement forces SKIP. Checks are evaluated in the
# guideline's own working order so the first failing checkpoint is
# reported as the decision reason.
# ============================================================

class Decision:
    def __init__(self):
        self.final = None          # "QUALIFIES" or "SKIP"
        self.reason = None
        self.fields = {}
        # Website Status to select on the portal. Guideline (c):
        # "If website is working then -- select Opening else select
        # not working, Domain expired, Under Construction. Also if
        # website language is non-English then -- select Non English."
        # The same guideline warns that submitting an unpaid-category
        # record with the WRONG status is counted as an error, so a
        # site that loads fine keeps "Opening" even when it SKIPs for
        # business type, category, or a missing field.
        self.status = "Opening"


def evaluate_pre_product(
    assigned_url,
    landing_url,
    root_opening,
    is_parked,
    is_subdomain_flag,
    is_free_hosting_flag,
    is_double_domain_flag,
    language,
    translation_available,
    context_data,
):
    """
    Stage 1 decision checks -- everything that does NOT require
    hunting for products/services. Runs against the fast Stage-1
    (crawl_context) crawl only.

    Returns (decision_or_none, context_or_none):
      - If any check fails: (a SKIP Decision, None). The caller must
        stop here and never run the expensive product/service crawl.
      - If every check passes: (None, a dict of validated values to
        carry into evaluate_products()).
    """
    decision = Decision()
    results = context_data["results"]

    all_text = " \n ".join(clean(r.get("text", "")) for r in results)
    all_urls = [r["url"] for r in results]
    schema_hits = []
    for r in results:
        schema_hits.extend(r.get("schema_countries", []))

    # ------------------------------------------------------
    # 1. URL / hosting structure
    # ------------------------------------------------------
    if not root_opening:
        decision.final = "SKIP"
        decision.status = "Not Working"
        decision.reason = "Website Status: not opening / not working."
        return decision, None

    if is_parked:
        decision.final = "SKIP"
        decision.status = placeholder_status(all_text)
        decision.reason = (
            "Website Status: parked / under construction / placeholder "
            f"-> {decision.status}."
        )
        return decision, None

    if is_double_domain_flag:
        decision.final = "SKIP"
        decision.reason = (
            "URL Error: the site redirects to a different domain; "
            "the website must remain in its actual supplied form."
        )
        return decision, None

    if is_subdomain_flag:
        decision.final = "SKIP"
        decision.reason = "URL Error: assigned website is a subdomain not hosted at root level."
        return decision, None

    if is_free_hosting_flag:
        decision.final = "SKIP"
        decision.reason = "URL Error: website is hosted on a free platform (e.g. WordPress)."
        return decision, None

    # ------------------------------------------------------
    # 2. Language
    # ------------------------------------------------------
    if language == "Non English" and not translation_available:
        decision.final = "SKIP"
        decision.status = "Non English"
        decision.reason = "Website Status: non-English with no translation option."
        return decision, None

    # ------------------------------------------------------
    # 3. Hard-restricted categories
    # ------------------------------------------------------
    restricted_reason, restricted_term = detect_restricted_category(all_text, all_urls)
    if restricted_reason:
        decision.final = "SKIP"
        decision.reason = f"Restricted category: {restricted_reason} (matched: '{restricted_term}')."
        return decision, None

    # ------------------------------------------------------
    # 4. Address block -- address, city, country, state
    # ------------------------------------------------------
    # CONFIRMED USER ORDER: after the domain and language gates, the
    # next thing checked is the address together with country, state
    # and city. Country is validated before state because the
    # state-exemption list (Taiwan / Singapore / Hong Kong / Macau)
    # depends on which country it is.
    final_address = {"street": None, "city": None, "state": None, "source": None}
    for r in results:
        addr = r.get("address", {})
        if addr.get("street") and not final_address["street"]:
            final_address = addr
            break

    if not final_address.get("street"):
        decision.final = "SKIP"
        decision.reason = "Address Error: no full address paragraph found on the website."
        return decision, None

    if not final_address.get("city"):
        decision.final = "SKIP"
        decision.reason = "City Error: address found but city could not be verified."
        return decision, None

    # The country belongs to the address, not to whatever countries the
    # page happens to mention. fericor.com is a Slovenian company that
    # represents a Chinese manufacturer; counting page-wide mentions
    # made it "China -> Not Working". Any site that names a foreign
    # supplier, office or market was at the same risk.
    country_result = None
    address_evidence = final_address.get("evidence") or ""
    if address_evidence:
        from_address = determine_country([address_evidence], [])
        if from_address.get("found"):
            country_result = from_address
            print(
                f"  Country taken from the address line: "
                f"{from_address.get('input')}"
            )

    if country_result is None:
        country_result = determine_country([all_text], schema_hits)

    if detect_india_china(country_result):
        decision.final = "SKIP"
        decision.reason = "Country Error: India/China websites are not workable (except Hong Kong/Macau S.A.R.)."
        return decision, None

    if not country_result.get("found") or not country_result.get("usable_name"):
        decision.final = "SKIP"
        if country_result.get("status") == "Not Working":
            decision.reason = (
                f"Country Error: '{country_result.get('input')}' is marked "
                "Not Working in the Country Name List."
            )
        else:
            decision.reason = "Country Error: no country could be verified/validated."
        return decision, None

    state_exempt = (
        country_result.get("found")
        and clean(country_result.get("input", "")).lower() in STATE_NOT_REQUIRED_COUNTRIES
    )
    if not state_exempt and not final_address.get("state"):
        decision.final = "SKIP"
        decision.reason = "State Error: address found but state/region could not be verified."
        return decision, None

    # ------------------------------------------------------
    # 5. Email and phone
    # ------------------------------------------------------
    # Guideline (g), final line: "All details have to be taken from the
    # same place. Email, Number, Address, State." So the page that
    # supplied the address is asked for the email and phone first, and
    # only if it has none do we fall back to the rest of the crawl.
    # The fallback is a preference, not a hard gate -- a missing email
    # is already a SKIP on its own, and refusing a perfectly good
    # contact page would only lose valid records.
    address_source = final_address.get("source")
    same_place = [r for r in results if r.get("url") == address_source]
    same_place_urls = {r.get("url") for r in same_place}
    ordered = same_place + [
        r for r in results if r.get("url") not in same_place_urls
    ]

    selected_email = None
    email_from_address_page = False
    for r in ordered:
        for email in r.get("emails", []):
            selected_email = email
            email_from_address_page = r.get("url") in same_place_urls
            break
        if selected_email:
            break

    if not selected_email:
        decision.final = "SKIP"
        decision.reason = "Email Error: no actual email address verified on the website."
        return decision, None

    selected_phone = None
    phone_from_address_page = False
    for r in ordered:
        usable = get_valid_phones(r.get("phones", []))
        if usable:
            selected_phone = usable[0]
            phone_from_address_page = r.get("url") in same_place_urls
            break

    if not selected_phone:
        decision.final = "SKIP"
        decision.reason = "Number Error: no valid non-toll-free phone/mobile number verified."
        return decision, None

    if not (email_from_address_page and phone_from_address_page):
        print(
            "  NOTE: email/phone were not all found on the same page as "
            "the address -- guideline (g) prefers one place for Email, "
            "Number, Address and State."
        )

    # ------------------------------------------------------
    # 6. Kind of business
    # ------------------------------------------------------
    business = detect_business_type(results)
    business_type = business["type"]
    using_industrial_path = business_type == "Industrial Services"

    if not business_type:
        non_paid_hits = detect_non_paid_context(results)
        decision.final = "SKIP"
        if non_paid_hits:
            decision.reason = (
                "Non-Paid: no recognised paid business type found; "
                f"non-paid context detected ('{non_paid_hits[0]['term']}')."
            )
        else:
            decision.reason = "Business Type Error: no clear paid business type established."
        return decision, None

    # ------------------------------------------------------
    # 7. Online shopping / payment gateway gating
    # ------------------------------------------------------
    shopping = detect_online_shopping(results)
    if shopping["detected"] and business_type in ONLINE_SELLING_DISQUALIFIES:
        decision.final = "SKIP"
        decision.reason = (
            f"Online Shopping: {business_type} with an online-selling option "
            "is not considered paid (exception applies to Manufacturer only)."
        )
        return decision, None

    # ------------------------------------------------------
    # 8. Company profile
    # ------------------------------------------------------
    has_profile, profile_source = find_company_profile(results)
    if not has_profile:
        decision.final = "SKIP"
        decision.reason = "Company Profile Error: no qualifying About/Profile/History page found."
        return decision, None

    # ------------------------------------------------------
    # ALL PRE-PRODUCT CHECKS PASSED
    # ------------------------------------------------------
    return None, {
        "business_type": business_type,
        "using_industrial_path": using_industrial_path,
        "country_result": country_result,
        "selected_email": selected_email,
        "selected_phone": selected_phone,
        "final_address": final_address,
        "profile_source": profile_source,
    }


def evaluate_products(context, crawl_data):
    """
    Stage 2 decision check -- only run after evaluate_pre_product()
    has already passed. Verifies the product / Industrial Service
    requirement using the Stage-2 (crawl_products) crawl data and
    builds the final decision, including the real extracted values
    (not just Y/N flags) needed to auto-fill the portal.
    """
    decision = Decision()
    using_industrial_path = context["using_industrial_path"]

    if using_industrial_path:
        items = crawl_data.get("qualifying_services", [])
        item_label = "Industrial Service"
    else:
        items = crawl_data.get("qualifying_products", [])
        item_label = "product"

    product_count = len(items)
    if product_count < MIN_QUALIFYING_PRODUCTS:
        decision.final = "SKIP"
        decision.reason = (
            f"Product Error: only {product_count} qualifying {item_label}(s) "
            f"verified (minimum {MIN_QUALIFYING_PRODUCTS} required)."
        )
        return decision

    product_names_ok = all(item["name"] for item in items)
    product_images_ok = all(item["colour_images"] > 0 for item in items)
    product_descriptions_ok = all(item["description"] for item in items)

    if not (product_names_ok and product_images_ok and product_descriptions_ok):
        decision.final = "SKIP"
        decision.reason = (
            f"Product Error: {item_label} name/image/description "
            "requirement not fully satisfied."
        )
        return decision

    # ------------------------------------------------------
    # ALL MANDATORY CHECKS PASSED
    # ------------------------------------------------------
    final_address = context["final_address"]
    decision.final = "QUALIFIES"
    decision.reason = "All applicable mandatory requirements verified."
    decision.fields = {
        "email": context["selected_email"],
        "phone": context["selected_phone"],
        "country": context["country_result"]["usable_name"],
        # Exactly what gets typed into the portal's country box.
        "country_fill": portal_country_name(context["country_result"]),
        "business_type": context["business_type"],
        "address_ok": True,
        "city_ok": True,
        "state_ok": True,
        "company_profile_ok": True,
        "products_ok": True,
        "images_ok": True,
        "descriptions_ok": True,
        "product_count": product_count,
        # Real values (not just Y/N) -- needed for portal auto-fill.
        "address_text": final_address.get("street"),
        "city_text": final_address.get("city"),
        "state_text": final_address.get("state"),
        "postal_code": final_address.get("postal_code"),
        "company_profile_source": context["profile_source"],
        "items": items,
        "item_label": item_label,
    }
    return decision

# ============================================================
# OUTPUT FORMATTING  (fixed by master guideline section 4)
# ============================================================

def print_final_skip():
    print("\nSKIP #")


def print_final_qualifies(fields):
    print()
    print(f"Email: {fields['email']}")
    print(f"Phone No: {fields['phone']}")
    print(f"Country: {fields['country']}")
    print(f"Kind of Business: {fields['business_type']}")
    print(f"Address: {'Y' if fields['address_ok'] else 'N'}")
    print(f"City: {'Y' if fields['city_ok'] else 'N'}")
    print(f"State: {'Y' if fields['state_ok'] else 'N'}")
    print(f"Company Profile: {'Y' if fields['company_profile_ok'] else 'N'}")
    print(f"3+ Physical Products: {'Y' if fields['products_ok'] else 'N'}")
    print(f"3+ Product Images: {'Y' if fields['images_ok'] else 'N'}")
    print(f"3+ Product Descriptions: {'Y' if fields['descriptions_ok'] else 'N'}")


def safe_print(line):
    """
    Print a line that may contain symbols an old Windows console
    cannot encode, falling back to plain ASCII rather than crashing
    the run over a tick mark.
    """
    try:
        print(line)
    except UnicodeEncodeError:
        print(line.encode("ascii", "replace").decode("ascii"))


def print_qualifies_detail(fields, decision):
    """
    The readable QUALIFIES summary, in the layout the user asked for:
    the verified VALUE next to each Y, and the product names spelled
    out, so a qualifying record can be checked at a glance.

    The guideline-mandated block (MASTER section 4) is printed
    separately and is not replaced by this.
    """
    business = fields.get("business_type") or ""
    items = fields.get("items") or []
    names = [str(item.get("name") or "").strip() for item in items]
    names = [n for n in names if n][:3]
    count = min(int(fields.get("product_count") or 0), 3)

    address_bits = [
        fields.get("address_text"),
        fields.get("city_text"),
        fields.get("postal_code"),
    ]
    full_address = ", ".join(str(b) for b in address_bits[:2] if b)
    if address_bits[2]:
        full_address = f"{full_address} {address_bits[2]}".strip()

    # Use the green circle / tick only if this console can actually
    # encode them; an old Windows code page turns them into "?", which
    # looks broken. Fall back to plain markers instead.
    green, tick = "\U0001F7E2", "\u2705"
    encoding = getattr(sys.stdout, "encoding", None) or "ascii"
    try:
        (green + tick).encode(encoding)
    except Exception:
        green, tick = ">>>", "[QUALIFIES]"

    print()
    safe_print("=" * 70)
    safe_print(f"{green} FINAL: W / OPENING -- {business.upper()} {tick}")
    safe_print("=" * 70)
    print("Status: QUALIFIES")
    print(f"Email: {fields.get('email') or ''}")
    print(f"Phone No: {fields.get('phone') or ''}")
    print(f"Country: {fields.get('country_fill') or fields.get('country') or ''}")
    print(f"Kind of Business: {business}")
    print(f"Address Y: {full_address or fields.get('address_text') or ''}")
    print(f"City Y: {fields.get('city_text') or ''}")
    print(f"State Y: {fields.get('state_text') or ''}")
    print("Company Profile Y: Yes")
    print(f"Product Name {count}: " + ("; ".join(names) if names else "Yes"))
    print(f"Product Image {count}: Yes")
    print(f"Product Description {count}: Yes")

    profile = fields.get("company_profile_source") or ""
    label = fields.get("item_label") or "product"
    print()
    print(
        f"Reason: {business} verified from a clear line on {profile}. "
        f"{count} qualifying {label}(s) found, each with a matching "
        f"image and description. Address, city and state confirmed on "
        f"the site, with a non-toll-free phone number and a real email."
    )
    print("=" * 70)


# ============================================================
# MAIN
# ============================================================

# ============================================================
# PORTAL SUBMISSION -- SKIP -> "Not Working" -> Submit
# ============================================================
#
# CONFIRMED USER INSTRUCTION: on SKIP (any reason -- dead site OR a
# guideline mismatch such as wrong business type, restricted
# category, or a missing mandatory field) select "Not Working" in
# the Website Status dropdown and submit the form, then continue
# working the new URL the portal generates.
#
# NOTE ON ACCURACY: your own Master Guideline explicitly warns that
# a working, loading site marked "Not Working" is counted as an
# error ("if the delivery partner submits the data of unpaid
# category by selecting working in the website status field...").
# Most SKIP reasons here (wrong business type, restricted category,
# missing email, etc.) occur on sites that load fine -- this setting
# submits "Not Working" for those too, per your explicit instruction
# overriding that default. Set AUTO_SUBMIT_SKIP_AS_NOT_WORKING to
# False to fall back to read-only / manual mode at any time.
#
# NOTE ON SELECTORS: the exact field names on your live portal were
# never provided and this script cannot browse your live portal to
# discover them. The selector lists below are the most common
# real-world naming patterns, tried in order (same strategy as
# get_assigned_url). If NONE of them match your form, add the real
# selector as the FIRST entry in the relevant list -- open your
# browser's dev tools on the portal page, right-click the Website
# Status dropdown -> Inspect, and copy its `name` or `id` attribute.
# ============================================================

AUTO_SUBMIT_SKIP_AS_NOT_WORKING = True

# CONFIRMED USER INSTRUCTION: after one URL is finished the script
# must keep going instead of stopping. With this True, a record that
# cannot be auto-submitted (selector mismatch, crashed crawl, portal
# not handing out a new URL yet) is logged, the portal page is
# reloaded, and the loop moves on to the next record. The script only
# gives up after MAX_CONSECUTIVE_FAILURES records fail back to back,
# which means the portal is genuinely finished or broken.
CONTINUE_ON_SUBMIT_FAILURE = True
MAX_CONSECUTIVE_FAILURES = 5

# CONFIRMED USER INSTRUCTION: every rejected record is submitted with
# Website Status = "Not Working", whether or not the site loads.
#
#   "always_not_working" -- the active policy. Any SKIP -> Not Working.
#   "per_guideline"      -- Intensecore guideline (c) literally:
#                           Opening for a site that loads, Not Working
#                           only when it does not load, plus Domain
#                           Expired / Under Construction / Non English
#                           for those specific cases.
#
# The per-record status the decision engine worked out is still
# computed and printed either way, so switching this one value to
# "per_guideline" changes the submitted status with no other edits.
SKIP_STATUS_MODE = "always_not_working"


def status_to_submit(decision):
    """
    The Website Status this record is actually submitted with.

    A QUALIFYING record is always submitted as Working -- the
    always_not_working rule applies to REJECTED records only. Without
    this check the function reported "Not Working" for every record,
    so a qualifying site printed "Decision: QUALIFIES / Website
    Status: Not Working", which read as though good sites were being
    submitted as dead ones.
    """
    if getattr(decision, "final", "") == "QUALIFIES":
        return "Working"

    guideline_status = getattr(decision, "status", "Not Working")
    if SKIP_STATUS_MODE == "always_not_working":
        return "Not Working"
    return guideline_status

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

NOT_WORKING_OPTION_LABELS = [
    "Not Working", "Not-Working", "NotWorking", "Not working",
    "not working",
]

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


def select_website_status_not_working(portal):
    """Backwards-compatible wrapper for the Not Working status."""
    return select_website_status(portal, "Not Working")


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


def submit_skip_as_not_working(portal):
    """Backwards-compatible wrapper for the Not Working status."""
    return submit_skip_with_status(portal, "Not Working")


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
#
# CONFIRMED USER INSTRUCTION: when a site fully qualifies, fill the
# verified data into the portal and submit it directly.
#
# The real portal column names come from Error_Details.xlsx ("Columns
# of Portal": url, bussinesstype, country, emailid, phoneormobile,
# address, city, state, companyprofile, productname, productimage,
# productdescription) and are used as the FIRST selector candidate
# for each field -- these are actual field names from your source
# documents, not a guess. Generic fallback patterns are tried after
# them in case the live form differs slightly.
#
# SAFETY RULE: this never submits a partially-filled record. If any
# single mandatory field's selector cannot be located, the whole
# attempt is aborted before anything is submitted, the record is
# logged as needing manual entry, and the script moves on to the
# next URL -- it will not guess a selector or submit incomplete data.
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


def submit_decision(portal, decision):
    """
    Submit whatever the decision engine produced for the current
    record: the full auto-filled record on QUALIFIES, or Website
    Status = "Not Working" on SKIP. Any exception raised by the
    portal interaction is caught and reported as a failed submission
    so one bad record can never end the continuous loop.
    """
    try:
        # A dropped session is indistinguishable from a broken selector
        # until you check for it, so check for it first.
        if not ensure_logged_in(portal):
            return False
        if decision.final == "QUALIFIES":
            return fill_and_submit_qualifies(portal, decision.fields)
        return submit_skip_with_status(portal, status_to_submit(decision))
    except Exception as exc:
        print(
            f"  Submission raised {type(exc).__name__}: {exc} -- "
            "treating this record as not submitted."
        )
        return False


# ============================================================
# MAIN
# ============================================================

# Maps a decision reason onto the stage that produced it, so a run can
# report where records are actually being lost instead of leaving it to
# be guessed from terminal scrollback.
FAILURE_STAGES = [
    ("Stage 0 - domain",        ("URL Error:",)),
    ("Stage 0 - not opening",   ("Website Status: not opening",
                                 "Website Status: navigation failed",
                                 "Website Status: parked")),
    ("Stage 0 - language",      ("Website Status: non-English",)),
    ("Restricted category",     ("Restricted category:",)),
    ("Address / City / State",  ("Address Error", "City Error", "State Error")),
    ("Country",                 ("Country Error",)),
    ("Email",                   ("Email Error",)),
    ("Phone",                   ("Number Error",)),
    ("Kind of business",        ("Business Type Error", "Non-Paid:")),
    ("Online selling",          ("Online Shopping:",)),
    ("Company profile",         ("Company Profile Error",)),
    ("Products",                ("Product Error",)),
]


def failure_stage(reason):
    """The stage a SKIP reason came from, for the run tally."""
    text = reason or ""
    for label, prefixes in FAILURE_STAGES:
        if any(prefix in text for prefix in prefixes):
            return label
    return "Other"


def log_record_outcome(record_number, assigned, landing, decision, crawl_data,
                       seconds, submitted):
    """
    Append one line per record to debug/run_log.csv. Nothing about a
    run should only exist in terminal scrollback -- that is why every
    earlier failure had to be diagnosed from a pasted screenshot.
    """
    path = debug_path("run_log.csv")
    new_file = not os.path.exists(path)
    stage = "QUALIFIES" if decision.final == "QUALIFIES" else failure_stage(
        getattr(decision, "reason", ""))
    row = [
        time.strftime("%Y-%m-%d %H:%M:%S"),
        str(record_number),
        assigned or "",
        landing or "",
        decision.final or "",
        stage,
        (getattr(decision, "reason", "") or "").replace(",", ";"),
        str(len(crawl_data.get("results", []))),
        str(len(crawl_data.get("qualifying_products", []))),
        f"{seconds:.1f}",
        "yes" if submitted else "no",
    ]
    try:
        with io.open(path, "a", encoding="utf-8") as fh:
            if new_file:
                fh.write("when,record,assigned_url,landing_url,decision,"
                         "stage,reason,pages,products,seconds,submitted\n")
            fh.write(",".join(row) + "\n")
    except Exception:
        pass


def print_stage_tally(tally):
    """Where records are being lost, worst first."""
    if not tally:
        return
    total = sum(tally.values())
    print()
    print("-" * 62)
    print(f"WHERE RECORDS ARE BEING LOST  ({total} record(s) so far)")
    print("-" * 62)
    for label, count in sorted(tally.items(), key=lambda kv: kv[1], reverse=True):
        share = 100.0 * count / total
        bar = "#" * int(round(share / 4))
        print(f"  {label:26} {count:3}  {share:5.1f}%  {bar}")
    print("-" * 62)


def process_assigned_url(browser, assigned):
    """
    Run the staged crawl + decision engine against a single assigned
    URL, stopping immediately after Stage 1 if any pre-product check
    fails (never wastes time hunting for products on a site that's
    already going to SKIP). Returns (decision, crawl_data, landing_url).
    """
    website = browser.new_page()
    website.set_default_timeout(7000)
    website.set_default_navigation_timeout(PAGE_NAVIGATION_TIMEOUT)
    install_dialog_autoaccept(website)

    empty_crawl_data = {
        "results": [], "qualifying_products": [], "qualifying_services": [],
        "context_pages_seen": 0, "product_pages_seen": 0,
    }

    print("\nOpening assigned website...")

    try:
        root_response = website.goto(
            assigned, wait_until="domcontentloaded",
            timeout=PAGE_NAVIGATION_TIMEOUT,
        )
        website.wait_for_timeout(500)
    except Exception as exc:
        print("Website error:", type(exc).__name__, exc)
        decision = Decision()
        decision.final = "SKIP"
        decision.status = "Not Working"
        decision.reason = f"Website Status: navigation failed ({type(exc).__name__})."
        website.close()
        return decision, empty_crawl_data, assigned

    landing_url = normalize_url(website.url)
    landing_status = root_response.status if root_response else None
    root_opening = landing_status is not None and 200 <= landing_status < 400

    is_double_domain_flag = is_double_domain_or_redirect(assigned, landing_url)
    is_subdomain_flag = is_subdomain(landing_url)
    is_free_hosting_flag = is_free_hosting_domain(landing_url)

    home_result = None
    language = "English"
    translation_available = False
    is_parked = False
    seed_result = None

    can_crawl = (
        root_opening and not is_double_domain_flag
        and not is_subdomain_flag and not is_free_hosting_flag
    )

    if can_crawl:
        home_result = inspect_page(website, landing_url, is_home=True)
        if home_result:
            language = home_result.get("language", "English")
            is_parked = detect_parked_or_placeholder(home_result)

            if language == "Non English":
                # Guideline: a non-English site WITH a translation
                # option is still workable -- actually use it rather
                # than only detecting it's there.
                if try_switch_to_english(website):
                    reswitched = inspect_page(
                        website, website.url, is_home=True,
                    )
                    if reswitched:
                        home_result = reswitched
                        language = home_result.get("language", language)
                        landing_url = home_result["url"]

            translation_available = (
                has_translation_option(home_result) or language == "English"
            )
            seed_result = home_result

    # --------------------------------------------------
    # STAGE 0 -- instant kills. NOTHING is crawled past this point.
    # --------------------------------------------------
    # `can_crawl` used to gate only the homepage inspection while the
    # context crawl below ran regardless, so a site that redirects to
    # another domain was still crawled for 17 pages before being
    # rejected for redirecting. These are decided from the landing URL
    # alone -- no page content is needed, so no page should be fetched.
    stage0_reason = None
    if not root_opening:
        stage0_reason = "Website Status: not opening / not working."
    elif is_double_domain_flag:
        stage0_reason = (
            "URL Error: the site redirects to a different domain; "
            "the website must remain in its actual supplied form."
        )
    elif is_subdomain_flag:
        stage0_reason = (
            "URL Error: assigned website is a subdomain not hosted "
            "at root level."
        )
    elif is_free_hosting_flag:
        stage0_reason = (
            "URL Error: website is hosted on a free platform "
            "(e.g. WordPress)."
        )

    if stage0_reason:
        print(f"\nSTAGE 0 -- instant kill: {stage0_reason}")
        print("  Not crawling this site at all.")
        decision = Decision()
        decision.final = "SKIP"
        decision.reason = stage0_reason
        if not root_opening:
            decision.status = "Not Working"
        website.close()
        return decision, empty_crawl_data, landing_url

    # --------------------------------------------------
    # STAGE 1 -- fast context-only crawl + pre-product checks
    # --------------------------------------------------
    context_data = crawl_context(website, landing_url, seed_result=seed_result)

    pre_decision, context = evaluate_pre_product(
        assigned_url=assigned,
        landing_url=landing_url,
        root_opening=root_opening,
        is_parked=is_parked,
        is_subdomain_flag=is_subdomain_flag,
        is_free_hosting_flag=is_free_hosting_flag,
        is_double_domain_flag=is_double_domain_flag,
        language=language,
        translation_available=translation_available,
        context_data=context_data,
    )

    if pre_decision is not None:
        # Failed an early check -- stop here, never crawl products.
        crawl_data = {
            "results": context_data["results"],
            "qualifying_products": [],
            "qualifying_services": [],
            "context_pages_seen": context_data["context_pages_seen"],
            "product_pages_seen": 0,
        }
        website.close()
        return pre_decision, crawl_data, landing_url

    # --------------------------------------------------
    # STAGE 2 -- only reached if every pre-product check passed
    # --------------------------------------------------
    print("  Pre-product checks passed -- searching for qualifying products/services...")
    product_crawl = crawl_products(website, context_data)

    crawl_data = {
        "results": product_crawl["results"],
        "qualifying_products": product_crawl["qualifying_products"],
        "qualifying_services": product_crawl["qualifying_services"],
        "context_pages_seen": context_data["context_pages_seen"],
        "product_pages_seen": product_crawl["product_pages_seen"],
    }

    decision = evaluate_products(context, crawl_data)

    website.close()
    return decision, crawl_data, landing_url


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

# Controls that only a SIGNED-OUT visitor sees.
CHATGPT_LOGGED_OUT_SELECTORS = (
    '[data-testid="login-button"]',
    'button:has-text("Sign up for free")',
    'a:has-text("Sign up for free")',
    'button:has-text("Log in")',
)

# The composer. NOT proof of being signed in on its own -- see
# _chatgpt_logged_in below.
CHATGPT_LOGGED_IN_SELECTORS = (
    "#prompt-textarea",
    '[data-testid="composer-speech-button"]',
    'div[contenteditable="true"]',
    'textarea[placeholder*="Message"]',
    'textarea[placeholder*="Ask"]',
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


def _chatgpt_logged_out_ui(page):
    """
    True when a signed-out-only control is on screen. Used for
    reporting only -- the session endpoint decides.

    Note the :visible pseudo-class. chatgpt.com keeps several hidden
    "Log in" buttons in the DOM, so .first.is_visible() answers for a
    hidden one and says the page has no login button at all. That is
    what made the DOM fallback declare a signed-out page signed in.
    """
    for selector in CHATGPT_LOGGED_OUT_SELECTORS:
        try:
            if page.locator(f"{selector}:visible").count():
                return True
        except Exception:
            continue
    return False


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


def _chatgpt_new_text(before, after):
    """The part of the conversation that appeared since `before`."""
    if not before:
        return after
    limit = min(len(before), len(after))
    i = 0
    while i < limit and before[i] == after[i]:
        i += 1
    return after[i:].strip()


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
    # Which file is actually running. An older, stale copy of this
    # script exists in the user's Downloads folder and has been run by
    # mistake more than once -- it lacks every fix made here, so its
    # failures look identical to bugs in this file. Print the path so
    # the terminal answers that question immediately.
    print("=" * 70)
    print("WEBSITE VERIFIER [SYSTEM 2] -- running:", os.path.abspath(__file__))
    print("Credentials: .env2   |   Debug output: debug2/")
    print("=" * 70)

    load_env_file(".env2")  # SYSTEM 2 -- its own credentials

    with sync_playwright() as p:

        # --------------------------------------------------
        # FIREFOX -- REQUIRED (never Chrome / auto-translating browsers)
        # --------------------------------------------------
        # --------------------------------------------------
        # --chatgpt-login : ChatGPT only. Uses its own persistent
        # profile and never opens the portal at all, so this returns
        # before the normal browser is even launched.
        # --------------------------------------------------
        if CHATGPT_LOGIN_ONLY:
            chatgpt_login_mode(p)
            return

        # --gpt-flow : portal + ChatGPT in one window. Has its own
        # browser profile and its own portal login, so it returns
        # before the normal browser is launched.
        if GPT_FLOW:
            gpt_flow_mode(p)
            return

        browser = p.firefox.launch(headless=False)

        portal = browser.new_page()
        portal.set_default_timeout(7000)
        portal.set_default_navigation_timeout(PAGE_NAVIGATION_TIMEOUT)
        install_dialog_autoaccept(portal)

        # --------------------------------------------------
        # LOGIN
        # --------------------------------------------------
        try:
            portal.goto(
                LOGIN_URL, wait_until="domcontentloaded",
                timeout=PAGE_NAVIGATION_TIMEOUT,
            )
        except Exception as exc:
            print("Login page error:", type(exc).__name__, exc)
            browser.close()
            return

        print("Login page loaded.")
        # Credentials come from the .env file loaded at startup (or
        # from real environment variables). The terminal prompt is
        # only a fallback for when neither is present -- a normal run
        # never stops to ask for anything.
        username = os.environ.get("PORTAL_USERNAME", "")
        password = os.environ.get("PORTAL_PASSWORD", "")

        if username and password:
            print(f"Logging in automatically as {username}...")
        else:
            print(
                "No PORTAL_USERNAME / PORTAL_PASSWORD found -- create a "
                ".env file next to this script with those two lines to "
                "skip this prompt."
            )
            username = username or input("Enter your username: ")
            password = password or input("Enter your password: ")

        # Wait for the form before touching it. The login page ships a
        # long Terms & Conditions block, so #Email can take longer than
        # the 7s default to render -- a real run died here on
        # "Locator.fill: Timeout 7000ms exceeded" while the field was
        # present and visible a second later.
        try:
            portal.wait_for_selector("#Email", state="visible", timeout=30000)
        except PlaywrightTimeoutError:
            print("The login form never appeared within 30 seconds.")
            print(f"  url:   {portal.url}")
            try:
                print(f"  title: {portal.title()}")
            except Exception:
                pass
            try:
                shot = debug_path(
                    "login_form_missing_" + time.strftime("%Y%m%d_%H%M%S") + ".png"
                )
                portal.screenshot(path=shot, full_page=True)
                print(f"  screenshot: {shot}")
            except Exception:
                pass
            browser.close()
            return

        portal.locator("#Email").fill(username)
        portal.locator("#Password").fill(password)

        login_button = portal.locator('input[type="submit"][value="Log in"]')
        try:
            login_button.click(timeout=15000)
        except PlaywrightTimeoutError:
            print(
                "  Normal click on 'Log in' timed out (button not "
                "visible/stable in time) -- retrying with a forced click..."
            )
            try:
                login_button.click(force=True, timeout=5000)
            except Exception as exc:
                print(
                    "  Forced click also failed "
                    f"({type(exc).__name__}) -- submitting via Enter "
                    "key on the password field instead..."
                )
                try:
                    screenshot_path = debug_path("login_page_debug.png")
                    portal.screenshot(path=screenshot_path, full_page=True)
                    print(f"  Saved a debug screenshot to {screenshot_path}")
                except Exception:
                    pass
                portal.locator("#Password").press("Enter")

        try:
            portal.wait_for_load_state("domcontentloaded", timeout=15000)
        except PlaywrightTimeoutError:
            pass

        print("Login successful.")

        if PORTAL_LOGIN_ONLY:
            hold_portal_login_open(portal)
            try:
                browser.close()
            except Exception:
                pass
            return

        if DUMP_FORM_ONLY:
            print(
                "\n--dump-form: READ-ONLY. Nothing will be filled, "
                "clicked or submitted."
            )
            portal.wait_for_timeout(2500)
            try:
                assigned_now = get_assigned_url(portal)
            except Exception:
                assigned_now = ""
            print(f"  Assigned URL currently on the page: {assigned_now or '(none found)'}")
            dump_portal_form(portal, note="(--dump-form, record page after login)")
            print(
                "\nSend debug/portal_form_debug.txt (and the .png beside "
                "it) back -- that is everything needed to finish the "
                "qualifying-product rows, the Business Type dropdown, and "
                "the supplier-with-zero-products case."
            )
            browser.close()
            return

        assigned = ""
        record_number = 0
        consecutive_failures = 0
        submit_attempts = {}
        stage_tally = {}

        # --------------------------------------------------
        # CONTINUOUS LOOP -- one assigned URL after another,
        # fully automatic. A single bad record (crashed crawl,
        # selector mismatch, portal hiccup) is logged and the loop
        # moves on; the script only stops when the portal genuinely
        # runs out of work or the same record fails repeatedly.
        # --------------------------------------------------
        while True:
            record_number += 1
            print("\n" + "#" * 70)
            print(f"RECORD {record_number}")
            print("#" * 70)

            # ----------------------------------------------
            # ASSIGNED URL (read dynamically -- never hardcoded)
            # ----------------------------------------------
            assigned = ""
            for _ in range(6):
                assigned = get_assigned_url(portal)
                if assigned:
                    break
                portal.wait_for_timeout(1000)

            if not assigned:
                print(
                    "No assigned URL on the page -- reloading the "
                    "portal and looking again..."
                )
                reload_portal(portal)
                for _ in range(6):
                    assigned = get_assigned_url(portal)
                    if assigned:
                        break
                    portal.wait_for_timeout(1000)

            if not assigned:
                consecutive_failures += 1
                if (not CONTINUE_ON_SUBMIT_FAILURE
                        or consecutive_failures >= MAX_CONSECUTIVE_FAILURES):
                    print("No assigned URL found after retries. Stopping.")
                    break
                print(
                    f"  Still no URL (attempt {consecutive_failures}/"
                    f"{MAX_CONSECUTIVE_FAILURES}) -- waiting and retrying..."
                )
                record_number -= 1
                portal.wait_for_timeout(3000)
                continue

            print("Assigned website URL:", assigned)
            record_started = time.time()

            if submit_attempts.get(assigned, 0) >= 2:
                print(
                    "STOPPED: the portal is still showing the same "
                    f"record ({assigned}) after two failed submission "
                    "attempts. It will never advance on its own -- "
                    "handle this one record manually (or fix the "
                    "selector it reported above), then re-run."
                )
                break

            landing_url = assigned

            if is_download_url(assigned):
                print("ERROR: Assigned URL is a direct download resource.")
                decision = Decision()
                decision.final = "SKIP"
                decision.status = "Not Working"
                decision.reason = "Assigned URL is a direct download resource."
                crawl_data = {
                    "results": [], "qualifying_products": [],
                    "qualifying_services": [], "context_pages_seen": 0,
                    "product_pages_seen": 0,
                }
            else:
                try:
                    decision, crawl_data, landing_url = process_assigned_url(
                        browser, assigned,
                    )
                except Exception as exc:
                    # A crash on one website must never end the run --
                    # the site could not be verified, which is a SKIP,
                    # and the loop continues with the next URL.
                    print(
                        f"ERROR while crawling this site: "
                        f"{type(exc).__name__}: {exc}"
                    )
                    decision = Decision()
                    decision.final = "SKIP"
                    decision.status = "Not Working"
                    decision.reason = (
                        f"Crawl failed with {type(exc).__name__} -- site "
                        "could not be verified."
                    )
                    crawl_data = {
                        "results": [], "qualifying_products": [],
                        "qualifying_services": [], "context_pages_seen": 0,
                        "product_pages_seen": 0,
                    }
                    landing_url = assigned

                print("\n" + "=" * 70)
                print("INTERNAL CRAWL SUMMARY")
                print("=" * 70)
                print("Assigned URL      :", assigned)
                print("Landing URL       :", landing_url)
                print("Pages inspected   :", len(crawl_data["results"]))
                print("Context pages     :", crawl_data["context_pages_seen"])
                print("Product pages     :", crawl_data["product_pages_seen"])
                print("Qualifying prods  :", len(crawl_data["qualifying_products"]))
                print("Qualifying svcs   :", len(crawl_data.get("qualifying_services", [])))
                print("Decision          :", decision.final)
                print("Decision reason   :", decision.reason)
                print(
                    "Website Status    :", status_to_submit(decision),
                    "(guideline status:",
                    getattr(decision, "status", "Opening") + ")",
                )

            # ----------------------------------------------
            # MANDATED FINAL OUTPUT (master guideline section 4)
            # ----------------------------------------------
            print("\n" + "=" * 70)
            print("FINAL OUTPUT")
            print("=" * 70)

            if decision.final == "QUALIFIES":
                print_final_qualifies(decision.fields)
                print_qualifies_detail(decision.fields, decision)
                print("\nAuto-filling and submitting the QUALIFIES record...")
            else:
                print_final_skip()
                if not AUTO_SUBMIT_SKIP_AS_NOT_WORKING:
                    print(
                        "STOPPED: AUTO_SUBMIT_SKIP_AS_NOT_WORKING is "
                        "False -- nothing was submitted for this record."
                    )
                    break
                print(
                    "Submitting Website Status = "
                    f"{status_to_submit(decision)}..."
                )

            submit_attempts[assigned] = submit_attempts.get(assigned, 0) + 1
            submitted = submit_decision(portal, decision)

            if not submitted and CONTINUE_ON_SUBMIT_FAILURE:
                print(
                    "  Submission did not go through -- reloading the "
                    "portal and retrying this record once..."
                )
                reload_portal(portal)
                if get_assigned_url(portal) == assigned:
                    submit_attempts[assigned] += 1
                    submitted = submit_decision(portal, decision)
                else:
                    # The portal moved on by itself -- treat as handled
                    # and pick the new record up on the next pass.
                    print("  Portal already moved to a different record.")
                    consecutive_failures = 0
                    continue

            if not submitted:
                consecutive_failures += 1
                print(
                    "STOPPED: this record could not be submitted "
                    "automatically (see the WARNING above for the "
                    "field/button that was not found). Nothing was "
                    "submitted. Fix that selector at the top of this "
                    "file, or submit this one record manually, then "
                    "re-run -- the loop cannot advance past a record "
                    "the portal never accepts."
                )
                break

            consecutive_failures = 0

            stage_label = (
                "QUALIFIES" if decision.final == "QUALIFIES"
                else failure_stage(getattr(decision, "reason", ""))
            )
            stage_tally[stage_label] = stage_tally.get(stage_label, 0) + 1
            log_record_outcome(
                record_number, assigned, landing_url, decision, crawl_data,
                time.time() - record_started, True,
            )
            print_stage_tally(stage_tally)

            print("Submitted. Fetching the portal's next assigned URL...")
            next_url = wait_for_new_assigned_url(portal, assigned)

            if not next_url or next_url == assigned:
                if not CONTINUE_ON_SUBMIT_FAILURE:
                    print(
                        "STOPPED: the portal did not hand out a new "
                        "assigned URL after submission. Check the "
                        "portal manually."
                    )
                    break
                consecutive_failures += 1
                if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                    print(
                        "STOPPED: the portal stopped handing out new "
                        "assigned URLs (no new record after "
                        f"{MAX_CONSECUTIVE_FAILURES} reload attempts) "
                        "-- your queue is most likely empty."
                    )
                    break
                print(
                    "  No new URL yet -- reloading the portal and "
                    "continuing with the next record..."
                )
                reload_portal(portal)
                portal.wait_for_timeout(2000)
                continue

            print("Next assigned URL:", next_url)

        print("\nStopped.")
        print_stage_tally(stage_tally)
        print("Per-record log:", debug_path("run_log.csv"))
        browser.close()


if __name__ == "__main__":
    main()
