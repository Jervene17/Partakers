import os
import re
import json
import random
import calendar
import html
import datetime as dt
from collections import defaultdict
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)
# --- Auto-retry on Google Sheets rate limits (429) ---
# Patches gspread's transport layer once, at import time, so every read/write
# anywhere in this file gets automatic exponential-backoff retry — no need to
# wrap each of the many .get_all_records()/.append_rows()/etc. call sites
# individually. Only retries 429 (quota exceeded); any other error still
# raises immediately.
import time
import logging
from gspread.http_client import HTTPClient as _GspreadHTTPClient
from gspread.exceptions import APIError as _GspreadAPIError

_original_gspread_request = _GspreadHTTPClient.request


def _gspread_request_with_retry(self, *args, **kwargs):
    max_retries = 6  # waits ~2+4+8+16+30s: long enough to outlast the per-minute quota window
    for attempt in range(max_retries):
        try:
            return _original_gspread_request(self, *args, **kwargs)
        except _GspreadAPIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            if status == 429 and attempt < max_retries - 1:
                wait = min(2 ** (attempt + 1), 30) + random.random()
                logging.warning(f"Sheets API rate limited (429) — retrying in {wait:.1f}s (attempt {attempt + 1})")
                time.sleep(wait)
                continue
            raise


_GspreadHTTPClient.request = _gspread_request_with_retry


# --- Read cache: keep the bot under Google's Sheets read quota (~60/min) ---
# In gspread every ss.worksheet(), ss.worksheets(), get_all_records(),
# get_all_values(), col_values() and row_values() call is a separate API read,
# and a single button tap in this bot can trigger dozens of them (setup_sheet
# alone used to make ~16). This layer:
#   * reuses worksheet()/worksheets() metadata for META_CACHE_TTL seconds
#   * reuses read results for READ_CACHE_TTL seconds
#   * drops a sheet's cached reads the moment the bot writes to it, so the bot
#     always sees its own changes immediately
# Edits typed directly into the Sheet show up within READ_CACHE_TTL seconds
# (clear_read_cache() is also called before /sync_dashboard).
import copy
from gspread.spreadsheet import Spreadsheet as _GspreadSpreadsheet
from gspread.worksheet import Worksheet as _GspreadWorksheet

READ_CACHE_TTL = 10   # seconds
META_CACHE_TTL = 120  # seconds

_read_cache = {}  # ((spreadsheet_id, sheet_id), method, args) -> (timestamp, value)
_meta_cache = {}  # (spreadsheet_id, kind, arg) -> (timestamp, value)


def clear_read_cache():
    """Forget every cached read (forces fresh reads on the next access)."""
    _read_cache.clear()
    _meta_cache.clear()


def _invalidate_worksheet(ws):
    sheet_key = (ws.spreadsheet_id, ws.id)
    for k in [k for k in _read_cache if k[0] == sheet_key]:
        _read_cache.pop(k, None)


def _invalidate_meta(spreadsheet):
    for k in [k for k in _meta_cache if k[0] == spreadsheet.id]:
        _meta_cache.pop(k, None)


def _make_cached_reader(cls, name):
    original = getattr(cls, name)

    def wrapper(self, *args, **kwargs):
        key = ((self.spreadsheet_id, self.id), name, repr(args), repr(sorted(kwargs.items())))
        hit = _read_cache.get(key)
        if hit and time.time() - hit[0] < READ_CACHE_TTL:
            return copy.deepcopy(hit[1])
        value = original(self, *args, **kwargs)
        _read_cache[key] = (time.time(), copy.deepcopy(value))
        return value

    wrapper.__name__ = name
    return wrapper


def _make_invalidating_writer(cls, name):
    original = getattr(cls, name)

    def wrapper(self, *args, **kwargs):
        try:
            return original(self, *args, **kwargs)
        finally:
            _invalidate_worksheet(self)

    wrapper.__name__ = name
    return wrapper


_CACHED_READERS = ("get_all_records", "get_all_values", "col_values", "row_values")
_INVALIDATING_WRITERS = (
    "update", "update_cell", "update_cells", "update_acell", "batch_update",
    "append_row", "append_rows", "insert_row", "insert_rows", "delete_rows",
    "clear", "batch_clear", "resize", "add_rows", "add_cols",
)


def _install_read_cache(worksheet_cls, spreadsheet_cls):
    for name in _CACHED_READERS:
        setattr(worksheet_cls, name, _make_cached_reader(worksheet_cls, name))
    for name in _INVALIDATING_WRITERS:
        if hasattr(worksheet_cls, name):
            setattr(worksheet_cls, name, _make_invalidating_writer(worksheet_cls, name))

    orig_worksheet = spreadsheet_cls.worksheet
    orig_worksheets = spreadsheet_cls.worksheets
    orig_add_worksheet = spreadsheet_cls.add_worksheet
    orig_del_worksheet = spreadsheet_cls.del_worksheet

    def worksheet(self, title):
        key = (self.id, "ws", title)
        hit = _meta_cache.get(key)
        if hit and time.time() - hit[0] < META_CACHE_TTL:
            return hit[1]
        ws = orig_worksheet(self, title)  # raises WorksheetNotFound if missing (not cached)
        _meta_cache[key] = (time.time(), ws)
        return ws

    def worksheets(self, *args, **kwargs):
        key = (self.id, "all", repr(args) + repr(sorted(kwargs.items())))
        hit = _meta_cache.get(key)
        if hit and time.time() - hit[0] < META_CACHE_TTL:
            return list(hit[1])
        result = orig_worksheets(self, *args, **kwargs)
        _meta_cache[key] = (time.time(), list(result))
        return result

    def add_worksheet(self, *args, **kwargs):
        try:
            return orig_add_worksheet(self, *args, **kwargs)
        finally:
            _invalidate_meta(self)

    def del_worksheet(self, *args, **kwargs):
        try:
            return orig_del_worksheet(self, *args, **kwargs)
        finally:
            _invalidate_meta(self)

    spreadsheet_cls.worksheet = worksheet
    spreadsheet_cls.worksheets = worksheets
    spreadsheet_cls.add_worksheet = add_worksheet
    spreadsheet_cls.del_worksheet = del_worksheet


_install_read_cache(_GspreadWorksheet, _GspreadSpreadsheet)



# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

BOT_TOKEN = os.environ["SERVICE_PARTAKER_BOT_TOKEN"]
SPREADSHEET_ID = os.environ["SERVICE_PARTAKER_SHEET_ID"]
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]  # path or inline JSON, match your other bots' setup
BOT_USERNAME = os.environ["SERVICE_PARTAKER_BOT_USERNAME"]  # no @, e.g. "MyChurchSchedulerBot" — used to build deep links
# Telegram shows bot usernames as "@MyBot" inside the app, so it's an easy typo
# to paste that into the Railway variable with the @ still attached. A deep
# link with a literal @ in it (t.me/@MyBot?start=...) is invalid and Telegram
# falls back to a generic "open Telegram" page instead of the bot — so strip
# any @ (and stray whitespace) here, defensively, regardless of how it's set.
BOT_USERNAME = BOT_USERNAME.strip().lstrip("@")
# Link opened by the "📖 Guide to the bot" menu button. Override with a GUIDE_URL
# environment variable (e.g. if you host the guide elsewhere); set it to an empty
# value to hide the button.
GUIDE_URL = os.environ.get("GUIDE_URL", "https://jervene17.github.io/Partakers/").strip()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
CHURCH_TZ = ZoneInfo("Asia/Manila")  # adjust if the church is elsewhere


def today_local():
    """Today's date in the church's timezone (the server itself is usually UTC,
    which is a day behind Manila for ~8 hours every day)."""
    return dt.datetime.now(CHURCH_TZ).date()


# Selectable in place of a person for any role, on any service EXCEPT Sun
# Stop Sundays, when that service is a live broadcast relayed from another
# church rather than run locally. Never part of an eligible list used by
# random/equal-share generation — only offered as an extra manual option.
LIVE_BROADCAST = "Live broadcast"


def broadcast_button(service):
    """[[button]] to append to a manual picker's keyboard, or [] for Sun Stop Sundays."""
    if service == "SunStopSundays":
        return []
    return [[InlineKeyboardButton(f"📡 {LIVE_BROADCAST}", callback_data=LIVE_BROADCAST)]]

# ---------------------------------------------------------------------------
# Roster + role definitions
# NOTE: This is Phase 1 data (Sunday + Wednesday only). Predawn, Sun Stop
# Sundays, and Filipino Translation are included as constants for later
# phases but are not used by the generator yet.
# ---------------------------------------------------------------------------

# NOTE: "Preacher" is intentionally NOT in these dicts. Preacher is filled
# manually first (see set_preacher_* conversation below) and the generator
# treats the pre-filled preacher for each date as a hard exclusion from
# every other role that date, rather than randomly assigning it.
SUNDAY_ROLES = {
    "Praise Leader": ["M Azzel", "M Sarah", "M Jervene", "D Jabs", "M Rose"],
    "Presider": ["M Jhay", "M Cor", "M Jervene", "M Azzel", "M Ju Nara", "M Rose", "M Sarah"],
    "Representative Prayer": ["Dcn Ian", "M Cor", "M Jervene", "M Sarah", "M Rose", "M Jhay"],
}

WEDNESDAY_ROLES = {
    # Praise Leader temporarily removed — not enough people available yet;
    # will be re-added later.
    "Presider": ["Divine", "E Issa", "D Rue", "M Cor"],
    "Representative Prayer": ["M Cor", "D Rue", "Divine"],
}

# Eligible preachers per service, for the manual set_preacher flow.
# Default is P Auda for Sunday/Wednesday only — Predawn has NO default,
# it must be picked manually every time (still required before generating).
PREACHER_ELIGIBLE = {
    "Sunday": ["P Auda", "M Azzel", "M Sarah", "PP Bambi", "M Cor"],
    "Wednesday": ["P Auda", "M Azzel", "M Sarah", "PP Bambi", "M Cor"],
    "Predawn": ["P Auda", "M Azzel", "M Sarah", "PP Bambi", "M Cor"],
}
DEFAULT_PREACHER_BY_SERVICE = {
    "Sunday": "P Auda",
    "Wednesday": "P Auda",
    # Predawn intentionally has no entry — no default preacher
}

# Reserved for later phases (not used yet):
PREDAWN_ROLES = {
    # Praise Leader temporarily removed — not enough people available yet;
    # will be re-added later. Predawn currently has no randomized roles at
    # all: Preacher (manual, /set_preacher) and Tech (manual, /log_tech)
    # cover everything until Praise Leader comes back.
}
SUN_STOP_DEPTS = [
    "Male Career Dept", "Female Career Dept",
    "Male Family Dept", "Female Family Dept",
    "Male Campus Dept", "Female Campus Dept",
]
FILIPINO_TRANSLATION_ROLES = {
    "Initial Proofreading": ["Rue", "Grace", "Divine", "Pres Joan"],
    "2nd PR": ["M Jervene", "Rue", "Grace", "Divine", "Pres Joan"],
    "Filipino Preacher": ["M Jervene", "M Sarah", "Pres Joan", "E Issa", "Riza"],
}

ROLE_SETS = {
    "Sunday": SUNDAY_ROLES,
    "Wednesday": WEDNESDAY_ROLES,
    "Predawn": PREDAWN_ROLES,          # Preacher excluded — manually set first, same as Sunday/Wednesday
    "FilipinoTranslation": FILIPINO_TRANSLATION_ROLES,  # Sunday-only, gated on Sunday's Preacher being set
}

SUN_STOP_ROLES = {
    "Praise Leader": SUN_STOP_DEPTS,
    "Opening Prayer": SUN_STOP_DEPTS,
    "Welcome Remarks": SUN_STOP_DEPTS,
    "Testimony": SUN_STOP_DEPTS,
}

# Tech has its own schedule, entered manually — never randomized/equal-share.
# {service: {role: eligible_list}}
TECH_ROLES_BY_SERVICE = {
    "Sunday": {"Onsite Tech": ["D Mel", "M Azzel", "Andrea"]},
    "Wednesday": {"Onsite Tech": ["D Mel", "M Azzel"]},
    "Predawn": {
        "Onsite Tech": ["D Mel", "D Rue", "M Azzel", "Dcn Ian", "Andrea"],
        "Online Tech": ["Daryl", "D Mel", "Shaja"],
    },
    "SunStopSundays": {"Onsite Tech": ["D Mel", "M Azzel"]},
}
TECH_CADENCE = {"Sunday": "month", "Wednesday": "month", "Predawn": "week", "SunStopSundays": "month"}

SERVICE_WEEKDAY = {
    "Sunday": 6,     # Monday=0 ... Sunday=6
    "Wednesday": 2,
    "SunStopSundays": 6,
    "FilipinoTranslation": 6,  # tied to Sunday's dates
}
PREDAWN_WEEKDAYS = [0, 1, 2, 3, 4, 5]  # Monday-Saturday

SHEET_TABS = [
    "Roster",
    "Sunday",
    "Wednesday",
    "Predawn",
    "SunStopSundays",
    "FilipinoTranslation",
    "AdjustmentLog",
    "GroupChats",
    "ServiceConfig",
    "PreferenceRounds",
    "PredawnPattern",
    "Unavailability",
]

TAB_HEADERS = {
    "Roster": ["Name", "AssignmentCount"],
    "Sunday": ["Date", "Role", "Partaker", "Status"],
    "Wednesday": ["Date", "Role", "Partaker", "Status"],
    "Predawn": ["Date", "Role", "Partaker", "Status"],
    "SunStopSundays": ["Date", "Role", "Partaker", "Status"],
    "FilipinoTranslation": ["Date", "Role", "Partaker", "Status"],
    "AdjustmentLog": ["Timestamp", "Service", "Date", "Role", "OldPartaker", "NewPartaker", "Type", "Reason"],
    "GroupChats": ["Purpose", "ChatID"],
    # One row per role. Cadence: "month" or "week". Mode: "random" (equal-share
    # generator) or "manual" (logged only, like Preacher/Tech). Eligible is a
    # comma-separated list of names/departments.
    "ServiceConfig": ["ServiceType", "Cadence", "Role", "Eligible", "Mode"],
    # One row per preference-collection round. Status: "open"/"closed".
    # Deadline is an ISO datetime string with timezone offset.
    "PreferenceRounds": ["RoundID", "Service", "Year", "Month", "Deadline", "Status", "GroupChatID"],
    # Recurring weekly Predawn Preacher pattern: which Preacher covers which
    # weekday (0=Monday..5=Saturday), reused every month until changed.
    "PredawnPattern": ["Weekday", "Preacher"],
    # Dates a partaker said they can NOT do, collected in preference rounds.
    # One row per (service, date, partaker). Generation avoids these people on
    # those dates, and /cancel_role only offers replacements who aren't listed.
    "Unavailability": ["Service", "Date", "Partaker"],
}

DASHBOARD_MONTHS_AHEAD = 3  # how many upcoming months each dashboard tab shows


# ---------------------------------------------------------------------------
# Google Sheets helpers
# ---------------------------------------------------------------------------

_client = None


def get_client():
    """GOOGLE_CREDS_JSON can be either the full service-account JSON pasted
    directly into the env var (Railway-friendly — no file to manage) or a
    path to a JSON key file (handy for local testing). The authorized client
    is created once and reused (its token refreshes itself)."""
    global _client
    if _client is not None:
        return _client
    raw = GOOGLE_CREDS_JSON.strip()
    if raw.startswith("{"):
        creds = Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(raw, scopes=SCOPES)
    _client = gspread.authorize(creds)
    return _client


_setup_ss = None
_setup_checked_at = 0.0
SETUP_RECHECK_SECONDS = 600  # full tab/header/roster check at most every 10 minutes


def setup_sheet():
    """Idempotent: creates any missing tabs + headers, and seeds the Roster
    tab with every unique name across Sunday + Wednesday roles (count=0) if
    they aren't already present. Safe to re-run.

    This is called at the start of almost every command/button, and the full
    check costs ~16 Sheets reads, so it only really runs once per
    SETUP_RECHECK_SECONDS; in between the already-opened spreadsheet is returned."""
    global _setup_ss, _setup_checked_at
    if _setup_ss is not None and time.time() - _setup_checked_at < SETUP_RECHECK_SECONDS:
        return _setup_ss
    gc = get_client()
    ss = gc.open_by_key(SPREADSHEET_ID)
    existing = {ws.title for ws in ss.worksheets()}

    for tab in SHEET_TABS:
        if tab not in existing:
            ws = ss.add_worksheet(title=tab, rows=1000, cols=max(8, len(TAB_HEADERS[tab])))
            ws.append_row(TAB_HEADERS[tab])
        else:
            ws = ss.worksheet(tab)
            if ws.row_values(1) != TAB_HEADERS[tab]:
                ws.update(range_name="A1", values=[TAB_HEADERS[tab]])

    # seed roster (individuals across all services, plus Sun Stop Sundays departments,
    # so equal-share counts stay consistent whichever service/tab you generate from)
    roster_ws = ss.worksheet("Roster")
    existing_names = set(roster_ws.col_values(1)[1:])  # skip header
    all_names = set()
    for roles in (SUNDAY_ROLES, WEDNESDAY_ROLES, PREDAWN_ROLES, SUN_STOP_ROLES, FILIPINO_TRANSLATION_ROLES):
        for members in roles.values():
            all_names.update(members)
    for names in PREACHER_ELIGIBLE.values():
        all_names.update(names)
    for tech_roles in TECH_ROLES_BY_SERVICE.values():
        for members in tech_roles.values():
            all_names.update(members)
    new_names = sorted(all_names - existing_names)
    if new_names:
        roster_ws.append_rows([[name, 0] for name in new_names])

    # seed ServiceConfig with the built-in services' role setup, so custom
    # services added later via /add_service live in the same place and
    # /add_role can extend either built-in or custom services consistently.
    # Only seeds services not already present — never overwrites existing rows.
    config_ws = ss.worksheet("ServiceConfig")
    existing_services = {row["ServiceType"] for row in config_ws.get_all_records()}
    builtin_rows = []
    for service, roles in (("Sunday", SUNDAY_ROLES), ("Wednesday", WEDNESDAY_ROLES), ("Predawn", PREDAWN_ROLES)):
        if service in existing_services:
            continue
        cadence = "week" if service == "Predawn" else "month"
        for role, eligible in roles.items():
            builtin_rows.append([service, cadence, role, ", ".join(eligible), "random"])
        preacher_eligible = PREACHER_ELIGIBLE.get(service)
        if preacher_eligible:
            builtin_rows.append([service, cadence, "Preacher", ", ".join(preacher_eligible), "manual"])
        for role, eligible in TECH_ROLES_BY_SERVICE.get(service, {}).items():
            builtin_rows.append([service, cadence, role, ", ".join(eligible), "manual"])
    if "SunStopSundays" not in existing_services:
        for role, eligible in SUN_STOP_ROLES.items():
            builtin_rows.append(["SunStopSundays", "month", role, ", ".join(eligible), "random"])
        for role, eligible in TECH_ROLES_BY_SERVICE.get("SunStopSundays", {}).items():
            builtin_rows.append(["SunStopSundays", "month", role, ", ".join(eligible), "manual"])
    if "FilipinoTranslation" not in existing_services:
        for role, eligible in FILIPINO_TRANSLATION_ROLES.items():
            builtin_rows.append(["FilipinoTranslation", "month", role, ", ".join(eligible), "random"])
    if builtin_rows:
        config_ws.append_rows(builtin_rows)

    _setup_ss = ss
    _setup_checked_at = time.time()
    return ss


def load_service_configs(ss):
    """Reads ServiceConfig into {service_type: {"cadence": str,
    "roles": {role: {"eligible": [names], "mode": "random"|"manual"}}}}"""
    config_ws = ss.worksheet("ServiceConfig")
    configs = {}
    for row in config_ws.get_all_records():
        service = row["ServiceType"]
        entry = configs.setdefault(service, {"cadence": row["Cadence"], "roles": {}})
        eligible = [name.strip() for name in row["Eligible"].split(",") if name.strip()]
        entry["roles"][row["Role"]] = {"eligible": eligible, "mode": row["Mode"]}
    return configs


def add_service_config(ss, service_type, cadence, roles):
    """roles: list of (role_name, eligible_list, mode). Appends rows to
    ServiceConfig and creates the service's own Date/Role/Partaker/Status tab
    if it doesn't exist yet. Also tops up the Roster tab with any new names."""
    config_ws = ss.worksheet("ServiceConfig")
    config_ws.append_rows([
        [service_type, cadence, role, ", ".join(eligible), mode]
        for role, eligible, mode in roles
    ])

    existing_tabs = {ws.title for ws in ss.worksheets()}
    if service_type not in existing_tabs:
        ws = ss.add_worksheet(title=service_type, rows=1000, cols=4)
        ws.append_row(TAB_HEADERS["Sunday"])  # standard Date/Role/Partaker/Status shape

    roster_ws = ss.worksheet("Roster")
    existing_names = set(roster_ws.col_values(1)[1:])
    new_names = set()
    for _, eligible, _ in roles:
        new_names.update(eligible)
    new_names -= existing_names
    if new_names:
        roster_ws.append_rows([[name, 0] for name in sorted(new_names)])


def add_role_to_service(ss, service_type, role, eligible, mode):
    add_service_config(ss, service_type, load_service_configs(ss)[service_type]["cadence"], [(role, eligible, mode)])


# --- #6 Roster management: edit a role's Eligible list in ServiceConfig ---

def update_role_eligible(ss, service_type, role, new_eligible):
    """Rewrites the Eligible cell for one ServiceConfig row. Returns True if
    a matching row was found and updated."""
    config_ws = ss.worksheet("ServiceConfig")
    for i, r in enumerate(config_ws.get_all_records()):
        if r["ServiceType"] == service_type and r["Role"] == role:
            config_ws.update_cell(i + 2, 4, ", ".join(new_eligible))  # col 4 = Eligible
            return True
    return False


def add_member_to_roster(ss, name):
    """Registers a brand-new name in the Roster tab (count 0) if not already present."""
    roster_ws = ss.worksheet("Roster")
    existing = set(roster_ws.col_values(1)[1:])
    if name not in existing:
        roster_ws.append_rows([[name, 0]])
        return True
    return False


def add_member_to_role(ss, service_type, role, name):
    eligible = get_role_pool(ss, service_type, role)
    if name not in eligible:
        update_role_eligible(ss, service_type, role, eligible + [name])
    add_member_to_roster(ss, name)  # in case they weren't already registered


def remove_member_from_role(ss, service_type, role, name):
    eligible = get_role_pool(ss, service_type, role)
    if name in eligible:
        update_role_eligible(ss, service_type, role, [n for n in eligible if n != name])
        return True
    return False


def remove_member_from_all_roles(ss, name):
    """Strips `name` out of every role's Eligible list across every service.
    Returns a list of "Service - Role" strings they were removed from."""
    config_ws = ss.worksheet("ServiceConfig")
    removed_from = []
    for i, r in enumerate(config_ws.get_all_records()):
        eligible = [n.strip() for n in r["Eligible"].split(",") if n.strip()]
        if name in eligible:
            config_ws.update_cell(i + 2, 4, ", ".join(n for n in eligible if n != name))
            removed_from.append(f"{r['ServiceType']} - {r['Role']}")
    return removed_from


def load_assignment_counts(ss):
    """Reads current AssignmentCount per person from the Roster tab."""
    roster_ws = ss.worksheet("Roster")
    rows = roster_ws.get_all_records()  # list of {"Name":..., "AssignmentCount":...}
    counts = defaultdict(int)
    for row in rows:
        counts[row["Name"]] = int(row.get("AssignmentCount", 0) or 0)
    return counts


def save_assignment_counts(ss, counts):
    """Writes updated AssignmentCount values back to the Roster tab in a
    single batched update (instead of one API call per person), adding any
    names not already present."""
    roster_ws = ss.worksheet("Roster")
    names = [row["Name"] for row in roster_ws.get_all_records()]
    if names:
        roster_ws.update(
            range_name=f"B2:B{len(names) + 1}",
            values=[[counts.get(n, 0)] for n in names],
        )
    known = set(names)
    new_rows = [[n, c] for n, c in counts.items() if n not in known]
    if new_rows:
        roster_ws.append_rows(new_rows)


def append_schedule_rows(ss, service_type, rows, skip_preacher_rows=False, skip_keys=None):
    """rows: list of (date_str, role, partaker). Status defaults to 'scheduled'.
    skip_preacher_rows=True (Sunday/Wednesday) skips 'Preacher' rows, since
    those are written separately by set_preacher_* and must not be duplicated.
    skip_keys: an optional set/dict of (date_str, role) to also skip — used
    for slots claimed in a preference round (already_filled), which are
    included in `rows` for display purposes but already exist in the sheet."""
    ws = ss.worksheet(service_type)
    skip_keys = skip_keys or {}
    to_write = [
        [date_str, role, partaker, "scheduled"]
        for date_str, role, partaker in rows
        if not (skip_preacher_rows and role == "Preacher") and (date_str, role) not in skip_keys
    ]
    if to_write:
        ws.append_rows(to_write)


def get_role_assignments(ss, service_type, role, dates):
    """Reads existing rows for a given role in a given service/tab and
    returns {date_str: partaker} only for dates that already have one
    filled in. Generic — used for Preacher and for manual Tech logging."""
    ws = ss.worksheet(service_type)
    records = ws.get_all_records()  # [{"Date":..., "Role":..., "Partaker":..., "Status":...}, ...]
    date_strs = {d.isoformat() for d in dates}
    result = {}
    for row in records:
        if row.get("Role") == role and row.get("Date") in date_strs:
            result[row["Date"]] = row["Partaker"]
    return result


def write_role_assignments(ss, service_type, role, assignments):
    """assignments: list of (date_str, partaker). Writes rows for that role
    directly — used for Preacher and for manual Tech logging."""
    ws = ss.worksheet(service_type)
    ws.append_rows([[date_str, role, partaker, "scheduled"] for date_str, partaker in assignments])


def get_preacher_assignments(ss, service_type, dates):
    return get_role_assignments(ss, service_type, "Preacher", dates)


def missing_preacher_dates(service_type, dates, preacher_assignments):
    return [d for d in dates if d.isoformat() not in preacher_assignments]


def write_preacher_assignments(ss, service_type, assignments):
    write_role_assignments(ss, service_type, "Preacher", assignments)


# ---------------------------------------------------------------------------
# Core generator
# ---------------------------------------------------------------------------

def dates_in_month(year, month, weekday):
    """All dates in a given month that fall on the given weekday (0=Mon..6=Sun)."""
    _, last_day = calendar.monthrange(year, month)
    out = []
    for day in range(1, last_day + 1):
        d = dt.date(year, month, day)
        if d.weekday() == weekday:
            out.append(d)
    return out


def upcoming_mondays(n=6):
    """List of the next n Mondays (including today if today is Monday),
    used as week-picker options for Predawn."""
    today = today_local()
    days_ahead = (0 - today.weekday()) % 7
    first_monday = today + dt.timedelta(days=days_ahead)
    return [first_monday + dt.timedelta(weeks=i) for i in range(n)]


def dates_in_week(monday, weekdays):
    """Given a Monday date and a list of weekday numbers (0=Mon..6=Sun),
    returns the matching dates within that Mon-Sun week."""
    return [monday + dt.timedelta(days=w) for w in weekdays]


def week_dates(service, monday):
    """Dates covered by a weekly-cadence service for the week starting `monday`.
    Predawn only runs Monday-Saturday; other weekly services use all 7 days."""
    if service == "Predawn":
        return dates_in_week(monday, PREDAWN_WEEKDAYS)
    return [monday + dt.timedelta(days=i) for i in range(7)]


def pick_partaker(eligible, counts, taken_today, hard_exclude=None, soft_exclude=None, never=None):
    """Pick from eligible people, weighted toward whoever has the fewest
    total assignments so far (equal share, combined across all roles).

    - taken_today / hard_exclude: never picked for this role, full stop
      (taken_today = already has another role this date, e.g. the Preacher;
      hard_exclude = role-specific hard rule, e.g. Filipino Preacher -> Presider)
    - soft_exclude: avoided when a valid alternative exists, but allowed if
      it's the only option left (e.g. Filipino Preacher -> other roles)
    - never: people who must NEVER be picked for this role, even as a last
      resort (unlike hard_exclude, no fallback ever brings them back)
    Ties broken randomly. Returns None if the role has nobody eligible at all.
    """
    never = set(never or ())
    eligible = [p for p in eligible if p not in never]
    if not eligible:
        return None

    hard_exclude = hard_exclude or set()
    soft_exclude = soft_exclude or set()

    candidates = [p for p in eligible if p not in taken_today and p not in hard_exclude]
    if not candidates:
        # nobody eligible is free even ignoring same-day double-booking;
        # relax taken_today but keep hard_exclude, rather than leave the role empty
        candidates = [p for p in eligible if p not in hard_exclude]
    if not candidates:
        # hard_exclude ate the whole list (shouldn't happen with these
        # rosters, but don't crash) — fall back to full eligible list
        candidates = list(eligible)

    preferred = [p for p in candidates if p not in soft_exclude]
    pool = preferred if preferred else candidates  # soft exclusion only "if inevitable"

    min_count = min(counts[p] for p in pool)
    least_assigned = [p for p in pool if counts[p] == min_count]
    return random.choice(least_assigned)


def generate_schedule(service_type, year_month_list, counts, preacher_assignments,
                       filipino_preacher_assignments=None, roles=None, already_filled=None,
                       unavailable=None):
    """
    service_type: "Sunday" or "Wednesday"
    year_month_list: list of (year, month) tuples to generate for, in order.
                      Passing multiple months in one call keeps the running
                      counts balanced across all of them; generating one
                      month at a time (across separate calls, reusing the
                      same `counts` loaded from the sheet) achieves the same
                      "consider previous assignments" behavior requested.
    counts: dict[name] -> int, mutated in place and also returned.
    preacher_assignments: {date_str: partaker} — MUST already cover every
                      service date being generated (checked by the caller
                      via missing_preacher_dates before calling this).
                      The preacher is hard-excluded from every other role
                      that date, and is included in the returned rows for
                      display, but is NOT re-written to the sheet (it's
                      already there) — see append_schedule_rows.
    filipino_preacher_assignments: {date_str: partaker}, Sunday only. That
                      person is hard-excluded from Presider and soft-excluded
                      (avoid unless inevitable) from every other role.
    roles: {role: eligible_list}; defaults to the hardcoded ROLE_SETS but
                      callers should pass get_random_roles_for_service(ss,
                      service_type) so roster-management edits (#6) apply.
    already_filled: {(date_str, role): partaker} for slots already claimed
                      before generation — e.g. via a preference round. These
                      are skipped (not re-picked), included in the returned
                      rows for display but NOT re-written (already in the
                      sheet), their assignee hard-excluded from every other
                      role that date, and their equal-share count was
                      already applied when the slot was claimed — generation
                      must NOT increment counts for them again here.

    unavailable: {date_str: set(names)} — people who marked that date as one
                      they can't do. Hard-excluded from every role that date
                      (only ignored if that would leave a role with nobody).

    Returns: (list of (date_str, role, partaker), updated counts dict)
    """
    unavailable = unavailable or {}
    roles = roles if roles is not None else ROLE_SETS[service_type]
    weekday = SERVICE_WEEKDAY[service_type]
    filipino_preacher_assignments = filipino_preacher_assignments or {}
    already_filled = already_filled or {}
    schedule_rows = []

    for year, month in year_month_list:
        for d in dates_in_month(year, month, weekday):
            date_str = d.isoformat()
            preacher = preacher_assignments[date_str]  # precondition: must exist
            schedule_rows.append((date_str, "Preacher", preacher))

            taken_today = {preacher}  # hard: preacher gets no other role
            # anyone who already claimed ANY role this date (via preference
            # round) is also excluded from every other role that date
            for (fd, _frole), fperson in already_filled.items():
                if fd == date_str:
                    taken_today.add(fperson)
            fil_preacher = filipino_preacher_assignments.get(date_str)

            for role, eligible in roles.items():
                if (date_str, role) in already_filled:
                    schedule_rows.append((date_str, role, already_filled[(date_str, role)]))
                    continue
                hard_exclude = ({fil_preacher} if (role == "Presider" and fil_preacher) else set()) \
                    | set(unavailable.get(date_str, ()))
                soft_exclude = {fil_preacher} if (fil_preacher and role != "Presider") else set()
                person = pick_partaker(eligible, counts, taken_today, hard_exclude, soft_exclude)
                if person is None:
                    continue  # nobody eligible for this role — leave it unassigned
                taken_today.add(person)
                counts[person] += 1
                schedule_rows.append((date_str, role, person))

    return schedule_rows, counts


def generate_predawn_schedule(dates, counts, preacher_assignments, roles=None, already_filled=None):
    """Like generate_schedule, but for Predawn's weekly (Mon-Sat) cycle:
    Preacher is manually pre-filled (checked by the caller via
    missing_preacher_dates) and hard-excluded from every other role that
    date. No Filipino Preacher logic applies (Sunday only).
    `roles` defaults to the hardcoded PREDAWN_ROLES but callers should pass
    get_random_roles_for_service(ss, "Predawn") so roster-management edits
    to eligibility (#6) take effect.
    `already_filled`: see generate_schedule — same skip/exclude/no-recount
    behavior, for slots claimed in a preference round."""
    roles = roles if roles is not None else PREDAWN_ROLES
    already_filled = already_filled or {}
    rows = []
    for d in dates:
        date_str = d.isoformat()
        preacher = preacher_assignments[date_str]  # precondition: must exist
        rows.append((date_str, "Preacher", preacher))
        taken_today = {preacher}
        for (fd, _frole), fperson in already_filled.items():
            if fd == date_str:
                taken_today.add(fperson)
        for role, eligible in roles.items():
            if (date_str, role) in already_filled:
                rows.append((date_str, role, already_filled[(date_str, role)]))
                continue
            person = pick_partaker(eligible, counts, taken_today)
            if person is None:
                continue
            taken_today.add(person)
            counts[person] += 1
            rows.append((date_str, role, person))
    return rows, counts


def generate_simple_schedule(dates, roles, counts, already_filled=None, unavailable=None, never_by_role=None,
                             soft_avoid_from_role=None):
    """Generic equal-share generator for services with no manual-preacher
    precondition and no cross-role exclusions beyond "not more than 1 role
    per date" — used for Predawn's non-Preacher roles, Sun Stop Sundays
    (roles are department pools, plus an individual Onsite Tech pool), and
    any custom service.
    `already_filled`: see generate_schedule — same skip/exclude/no-recount
    behavior, for slots claimed in a preference round. `unavailable` is
    {date_str: set(names)} — hard-excluded that date, like generate_schedule.
    `never_by_role`: {role: {date_str: set(names)}} — people who must never get
    that role on that date (e.g. the English Preacher can't be Filipino Preacher).
    `soft_avoid_from_role`: {target_role: source_role} — once someone is given
    source_role on any date in this run, avoid also giving them target_role on
    a LATER date, unless nobody else eligible is available (e.g. someone who is
    the Filipino Preacher one Sunday is avoided, not blocked, for Initial
    Proofreading another Sunday, so they mainly just preach).
    Returns (rows, counts)."""
    already_filled = already_filled or {}
    unavailable = unavailable or {}
    never_by_role = never_by_role or {}
    soft_avoid_from_role = soft_avoid_from_role or {}
    role_history = defaultdict(set)  # role -> everyone already given that role earlier in this run
    rows = []
    for d in dates:
        date_str = d.isoformat()
        taken_today = set()
        for (fd, _frole), fperson in already_filled.items():
            if fd == date_str:
                taken_today.add(fperson)
        for role, eligible in roles.items():
            if (date_str, role) in already_filled:
                person = already_filled[(date_str, role)]
                rows.append((date_str, role, person))
                role_history[role].add(person)
                continue
            source_role = soft_avoid_from_role.get(role)
            soft_exclude = set(role_history[source_role]) if source_role else None
            person = pick_partaker(eligible, counts, taken_today, hard_exclude=set(unavailable.get(date_str, ())),
                                   soft_exclude=soft_exclude, never=never_by_role.get(role, {}).get(date_str))
            if person is None:
                continue  # nobody eligible for this role — leave it unassigned
            taken_today.add(person)
            counts[person] += 1
            role_history[role].add(person)
            rows.append((date_str, role, person))
    return rows, counts


async def warn_if_unavailable_scheduled(message, rows, unavailable):
    """Tells the admin when someone ended up on a date they marked as
    unavailable (only happens if nobody else eligible was free, or the slot
    was already filled before generation)."""
    clashes = [
        (d, role, person) for d, role, person in rows
        if role != "Preacher" and person in unavailable.get(d, ())
    ]
    if not clashes:
        return
    lines = "\n".join(
        f"- {dt.date.fromisoformat(d).strftime('%b %d')}: {role} — {person}" for d, role, person in clashes
    )
    await message.reply_text(
        "⚠️ Heads up — these people are scheduled on a date they marked as unavailable "
        "(no other eligible partaker was free, or the slot was already filled):\n" + lines
    )


def format_schedule_summary(service_type, schedule_rows):
    """Groups generated rows by date for a readable Telegram message."""
    by_date = defaultdict(list)
    for date_str, role, person in schedule_rows:
        by_date[date_str].append((role, person))

    lines = [f"*{service_type} Service Schedule*\n"]
    for date_str in sorted(by_date):
        d = dt.date.fromisoformat(date_str)
        lines.append(f"*{d.strftime('%B %d, %Y')}*")
        for role, person in by_date[date_str]:
            lines.append(f"{role} - {person}")
        lines.append("")
    return "\n".join(lines).strip()


# ---------------------------------------------------------------------------
# Telegram bot — "Generate Schedule" flow (Sunday/Wednesday only, Phase 1)
# ---------------------------------------------------------------------------

SELECT_SERVICE, SELECT_MONTH, CONFIRM = range(3)
PREACHER_SELECT_SERVICE, PREACHER_SELECT_MONTH, PREACHER_PICK = range(3, 6)
PREACHER_CONFIRM_CONFLICT = 140

MONTH_LOOKAHEAD = 6  # how many upcoming months to offer as buttons


def month_keyboard(prefix=""):
    today = today_local()
    options = []
    y, m = today.year, today.month
    for _ in range(MONTH_LOOKAHEAD):
        label = dt.date(y, m, 1).strftime("%B %Y")
        options.append([InlineKeyboardButton(label, callback_data=f"{prefix}{y}-{m}")])
        m += 1
        if m > 12:
            m = 1
            y += 1
    return InlineKeyboardMarkup(options)


# --- /set_preacher: manual preacher scheduling (must run before /generate) ---

async def set_preacher_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("Sunday", callback_data="Sunday")],
        [InlineKeyboardButton("Wednesday", callback_data="Wednesday")],
    ]
    await update.message.reply_text(
        "Set Preacher schedule — which service? (Predawn now uses /set_predawn_pattern instead)",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return PREACHER_SELECT_SERVICE


async def preacher_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service_type = query.data
    context.user_data["service_type"] = service_type
    await query.edit_message_text(
        f"{service_type} — pick a month:", reply_markup=month_keyboard()
    )
    return PREACHER_SELECT_MONTH


async def preacher_select_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year, month = map(int, query.data.split("-"))
    service_type = context.user_data["service_type"]
    weekday = SERVICE_WEEKDAY[service_type]
    dates = dates_in_month(year, month, weekday)

    ss = setup_sheet()
    existing = get_preacher_assignments(ss, service_type, dates)
    pending = [d for d in dates if d.isoformat() not in existing]

    context.user_data["preacher_ss"] = ss
    context.user_data["preacher_service_type"] = service_type
    context.user_data["preacher_pending_dates"] = pending
    context.user_data["preacher_new_assignments"] = []

    if not pending:
        await query.edit_message_text(
            f"All {service_type} dates in {dt.date(year, month, 1).strftime('%B %Y')} "
            f"already have a Preacher assigned. Nothing to do."
        )
        return ConversationHandler.END

    return await preacher_ask_next_date(query, context)


async def preacher_ask_next_date(query, context: ContextTypes.DEFAULT_TYPE, note=""):
    pending = context.user_data["preacher_pending_dates"]
    if not pending:
        ss = context.user_data["preacher_ss"]
        service_type = context.user_data["preacher_service_type"]
        write_preacher_assignments(ss, service_type, context.user_data["preacher_new_assignments"])
        lines = "\n".join(
            f"{d} - {p}" for d, p in context.user_data["preacher_new_assignments"]
        )
        await query.edit_message_text(f"Preacher schedule saved:\n{lines}")

        chat_id = get_group_chat_id(ss, "Service Partakers")
        if chat_id:
            await context.bot.send_message(
                chat_id=chat_id,
                text=f"✅ {service_type} Preacher schedule has been set.",
            )
        return ConversationHandler.END

    d = pending[0]
    service_type = context.user_data["preacher_service_type"]
    ss = context.user_data["preacher_ss"]
    eligible = get_role_pool(ss, service_type, "Preacher")
    default = DEFAULT_PREACHER_BY_SERVICE.get(service_type)
    buttons = []
    for name in eligible:
        label = f"{name} (default)" if name == default else name
        buttons.append([InlineKeyboardButton(label, callback_data=name)])
    buttons.extend(broadcast_button(service_type))
    await query.edit_message_text(
        f"{note}Preacher for {d.strftime('%B %d, %Y')}?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return PREACHER_PICK


async def preacher_pick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    name = query.data
    d = context.user_data["preacher_pending_dates"][0]
    ss = context.user_data["preacher_ss"]
    service_type = context.user_data["preacher_service_type"]
    date_str = d.isoformat()

    # Sunday's English Preacher can never also be that day's Filipino Preacher.
    if service_type == "Sunday" and forbidden_conflict(ss, "Sunday", "Preacher", date_str, name):
        return await preacher_ask_next_date(
            query, context,
            note=f"⛔ {name} is already the Filipino Preacher on {date_str}. The Preacher and the Filipino "
                 f"Preacher must be different people, so please pick someone else.\n\n",
        )

    conflicts = find_conflicts(ss, service_type, date_str, name)
    if conflicts:
        context.user_data["preacher_pending_name"] = name
        buttons = [
            [InlineKeyboardButton("Yes, assign as Preacher too", callback_data="yes")],
            [InlineKeyboardButton("No, pick someone else", callback_data="no")],
        ]
        await query.edit_message_text(
            conflict_warning_text(name, date_str, conflicts, "Preacher"), reply_markup=InlineKeyboardMarkup(buttons)
        )
        return PREACHER_CONFIRM_CONFLICT

    context.user_data["preacher_pending_dates"].pop(0)
    context.user_data["preacher_new_assignments"].append((date_str, name))
    return await preacher_ask_next_date(query, context)


async def preacher_confirm_conflict(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "yes":
        d = context.user_data["preacher_pending_dates"].pop(0)
        name = context.user_data.pop("preacher_pending_name")
        context.user_data["preacher_new_assignments"].append((d.isoformat(), name))
        return await preacher_ask_next_date(query, context)
    # "no" — re-show the same date's picker so they can choose someone else
    context.user_data.pop("preacher_pending_name", None)
    return await preacher_ask_next_date(query, context)


GENERATE_BUILTIN_SERVICES = ("Sunday", "Wednesday", "Predawn", "SunStopSundays", "FilipinoTranslation")
PREDAWN_PATTERN_BUTTON = "__predawn_pattern__"


async def generate_schedule_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("Sunday", callback_data="Sunday")],
        [InlineKeyboardButton("Wednesday", callback_data="Wednesday")],
        [InlineKeyboardButton("Predawn", callback_data="Predawn")],
        [InlineKeyboardButton("Sun Stop Sundays", callback_data="SunStopSundays")],
        [InlineKeyboardButton("Filipino Translation", callback_data="FilipinoTranslation")],
    ]
    # Custom services added with /add_service that have randomized roles.
    # If the sheet can't be read, the built-in buttons above still work.
    try:
        for service, cfg in load_service_configs(setup_sheet()).items():
            if service in GENERATE_BUILTIN_SERVICES:
                continue
            if any(r["mode"] == "random" for r in cfg["roles"].values()):
                keyboard.append([InlineKeyboardButton(service, callback_data=service)])
    except Exception:
        pass
    keyboard.append([InlineKeyboardButton("Set Predawn pattern", callback_data=PREDAWN_PATTERN_BUTTON)])
    await update.message.reply_text(
        "Which service would you like to generate a schedule for?",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return SELECT_SERVICE


async def select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "Predawn":
        ss = setup_sheet()
        context.user_data["pdp_ss"] = ss
        if not get_predawn_pattern(ss):
            await query.edit_message_text("No Predawn weekly pattern set yet — run /set_predawn_pattern first.")
            return ConversationHandler.END
        return await predawn_ask_month(update, context)
    if query.data == "SunStopSundays":
        await query.edit_message_text("Sun Stop Sundays — pick a month:", reply_markup=month_keyboard())
        return SELECT_SUNSTOP_MONTH
    if query.data == PREDAWN_PATTERN_BUTTON:
        context.user_data["pdp_ss"] = setup_sheet()
        context.user_data["pdp_weekday_idx"] = 0
        return await predawn_pattern_ask_day(update, context)
    if query.data not in ("Sunday", "Wednesday"):
        # Filipino Translation or a custom service: same flow as /generate_service
        ss = setup_sheet()
        context.user_data["gen_svc_ss"] = ss
        context.user_data["gen_svc_configs"] = load_service_configs(ss)
        return await generate_service_select(update, context)
    context.user_data["service_type"] = query.data
    await query.edit_message_text(
        f"{query.data} service selected. Pick a month to generate:",
        reply_markup=month_keyboard(),
    )
    return SELECT_MONTH


async def select_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year, month = map(int, query.data.split("-"))
    service_type = context.user_data["service_type"]
    weekday = SERVICE_WEEKDAY[service_type]
    dates = dates_in_month(year, month, weekday)

    ss = setup_sheet()

    # Precondition: every date in the month needs a Preacher already set.
    preacher_assignments = get_preacher_assignments(ss, service_type, dates)
    missing = missing_preacher_dates(service_type, dates, preacher_assignments)
    if missing:
        missing_list = ", ".join(d.strftime("%b %d") for d in missing)
        await query.edit_message_text(
            f"Can't generate yet — {service_type} has no Preacher scheduled for: "
            f"{missing_list}.\nRun /set_preacher first, then try /generate again."
        )
        return ConversationHandler.END

    # Precondition: the partaker preference round for this month must be
    # closed (deadline passed, or admin closed it early via /close_preferences).
    month_label = dt.date(year, month, 1).strftime("%B %Y")
    pref_round = find_round_for(ss, service_type, year, month)
    if not pref_round or pref_round.get("Status") != "closed":
        await query.edit_message_text(
            f"Can't generate yet — open (and then close) a preference round for {service_type} "
            f"{month_label} first: /open_preferences, then /close_preferences when ready."
        )
        return ConversationHandler.END

    await query.edit_message_text("Generating schedule...")

    # Filipino Preacher exclusions only apply to Sunday: whoever is the
    # Filipino Preacher that date can't be Presider (hard) and is avoided
    # for other roles when possible (soft). Read from the FilipinoTranslation tab.
    filipino_preacher_assignments = (
        get_role_assignments(ss, "FilipinoTranslation", "Filipino Preacher", dates)
        if service_type == "Sunday" else {}
    )

    # Slots already claimed via a preference round (or any manual write)
    # get skipped rather than re-picked. "Preacher" is excluded here since
    # it's already handled by preacher_assignments above.
    already_filled = get_already_filled(ss.worksheet(service_type), dates, exclude_roles=("Preacher",))

    # Dates partakers marked as unavailable in the preference round
    unavailable = get_unavailability(ss, service_type, dates)

    counts = load_assignment_counts(ss)
    schedule_rows, counts = generate_schedule(
        service_type, [(year, month)], counts, preacher_assignments, filipino_preacher_assignments,
        roles=get_random_roles_for_service(ss, service_type), already_filled=already_filled,
        unavailable=unavailable,
    )

    append_schedule_rows(ss, service_type, schedule_rows, skip_preacher_rows=True, skip_keys=already_filled)
    save_assignment_counts(ss, counts)

    summary = format_schedule_summary(service_type, schedule_rows)
    await query.message.reply_text(summary, parse_mode="Markdown")
    await warn_if_unavailable_scheduled(query.message, schedule_rows, unavailable)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Cancelled.")
    return ConversationHandler.END


WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

PREDAWN_PATTERN_DAY, PREDAWN_GEN_MONTH, PREDAWN_ADJUST = range(160, 163)


def get_predawn_pattern(ss):
    """{weekday_index (0=Mon..5=Sat): preacher_name}, from the recurring
    weekly pattern set via /set_predawn_pattern."""
    ws = ss.worksheet("PredawnPattern")
    return {int(r["Weekday"]): r["Preacher"] for r in ws.get_all_records()}


def set_predawn_pattern_day(ss, weekday_idx, preacher):
    ws = ss.worksheet("PredawnPattern")
    for i, r in enumerate(ws.get_all_records()):
        if int(r["Weekday"]) == weekday_idx:
            ws.update_cell(i + 2, 2, preacher)
            return
    ws.append_rows([[weekday_idx, preacher]])


def expand_predawn_pattern_to_month(ss, year, month):
    """(assignments, missing_weekday_names) — assignments is {date_str:
    preacher} for every Predawn date that month per the weekly pattern;
    missing_weekday_names lists any weekday with no pattern entry yet."""
    pattern = get_predawn_pattern(ss)
    dates = dates_matching_weekdays_in_month(year, month, PREDAWN_WEEKDAYS)
    assignments, missing = {}, set()
    for d in dates:
        wd = d.weekday()
        if wd in pattern:
            assignments[d.isoformat()] = pattern[wd]
        else:
            missing.add(WEEKDAY_NAMES[wd])
    return assignments, sorted(missing)


# --- /set_predawn_pattern: define the recurring weekly Preacher pattern ---
# (occasional — only needed when the arrangement actually changes), then
# hands off to the same month-generation tail as /generate_predawn.

async def set_predawn_pattern_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["pdp_ss"] = ss
    context.user_data["pdp_weekday_idx"] = 0
    return await predawn_pattern_ask_day(update, context)


async def predawn_pattern_ask_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = context.user_data["pdp_ss"]
    idx = context.user_data["pdp_weekday_idx"]
    send = update.callback_query.edit_message_text if update.callback_query else update.message.reply_text

    if idx >= len(WEEKDAY_NAMES):
        pattern = get_predawn_pattern(ss)
        lines = "\n".join(f"{WEEKDAY_NAMES[i]}: {pattern.get(i, 'TBA')}" for i in range(len(WEEKDAY_NAMES)))
        await send(f"Predawn weekly pattern set:\n{lines}\n\nNow let's generate a month from it.")
        # Post the month prompt as a NEW message so the pattern summary above
        # isn't immediately overwritten by an edit of the same message.
        target = update.callback_query.message if update.callback_query else update.message
        await target.reply_text("Generate Predawn for which month?", reply_markup=month_keyboard())
        return PREDAWN_GEN_MONTH

    eligible = get_role_pool(ss, "Predawn", "Preacher")
    current = get_predawn_pattern(ss).get(idx)
    buttons = [
        [InlineKeyboardButton(f"{name} (current)" if name == current else name, callback_data=name)]
        for name in eligible
    ]
    text = f"Preacher for {WEEKDAY_NAMES[idx]}s?" + (f" (currently {current})" if current else "")
    await send(text, reply_markup=InlineKeyboardMarkup(buttons))
    return PREDAWN_PATTERN_DAY


async def predawn_pattern_pick_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = context.user_data["pdp_ss"]
    idx = context.user_data["pdp_weekday_idx"]
    set_predawn_pattern_day(ss, idx, query.data)
    context.user_data["pdp_weekday_idx"] = idx + 1
    return await predawn_pattern_ask_day(update, context)


async def predawn_ask_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    send = update.callback_query.edit_message_text if update.callback_query else update.message.reply_text
    await send("Generate Predawn for which month?", reply_markup=month_keyboard())
    return PREDAWN_GEN_MONTH


# --- /generate_predawn: routine monthly use — expands the EXISTING pattern
# (no re-entry needed unless it changed) into that month's dates. ---

async def generate_predawn_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["pdp_ss"] = ss
    if not get_predawn_pattern(ss):
        await update.message.reply_text("No Predawn weekly pattern set yet — run /set_predawn_pattern first.")
        return ConversationHandler.END
    return await predawn_ask_month(update, context)


async def predawn_generate_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year, month = map(int, query.data.split("-"))
    ss = context.user_data["pdp_ss"]
    context.user_data["pdp_year"], context.user_data["pdp_month"] = year, month
    month_label = dt.date(year, month, 1).strftime("%B %Y")

    assignments, missing = expand_predawn_pattern_to_month(ss, year, month)
    if missing:
        await query.edit_message_text(
            f"No Preacher set for: {', '.join(missing)}. Run /set_predawn_pattern to fill those in first."
        )
        return ConversationHandler.END

    # Only write rows for dates that don't already have a Preacher, so a
    # re-run (or a prior /substitute) is never overwritten.
    ws = ss.worksheet("Predawn")
    existing_dates = {d.isoformat() for d in dates_matching_weekdays_in_month(year, month, PREDAWN_WEEKDAYS)}
    already = get_already_filled(ws, [dt.date.fromisoformat(d) for d in existing_dates])
    already_preacher_dates = {d for (d, role) in already if role == "Preacher"}
    new_rows = [(d, "Preacher", p) for d, p in assignments.items() if d not in already_preacher_dates]
    if new_rows:
        append_schedule_rows(ss, "Predawn", new_rows)

    buttons = [
        [InlineKeyboardButton("Yes, adjustments needed", callback_data="yes")],
        [InlineKeyboardButton("No, looks good", callback_data="no")],
    ]
    await query.edit_message_text(
        f"Predawn Preacher schedule set for {month_label} "
        f"({len(new_rows)} new, {len(already_preacher_dates)} already set).\n\n"
        f"Any dates need adjusting (swap/substitution) before sending the summary?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return PREDAWN_ADJUST


async def predawn_adjust_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = context.user_data["pdp_ss"]
    year, month = context.user_data["pdp_year"], context.user_data["pdp_month"]
    month_label = dt.date(year, month, 1).strftime("%B %Y")

    if query.data == "yes":
        await query.edit_message_text(
            "Use /swap or /substitute for any specific dates (already-set dates are never "
            "touched by re-running /generate_predawn). Sending the current summary to the group now."
        )
    else:
        await query.edit_message_text("Sending the summary to the group.")

    ws = ss.worksheet("Predawn")
    rows = rows_in_month(ws.get_all_records(), year, month)
    summary = format_schedule_summary(f"Predawn — {month_label}", records_to_rows(rows))
    chat_id = get_group_chat_id(ss, "Service Partakers")
    if chat_id:
        await context.bot.send_message(chat_id=chat_id, text=summary, parse_mode="Markdown")
    else:
        await query.message.reply_text(summary, parse_mode="Markdown")
    return ConversationHandler.END


# --- /generate_sunstop: monthly generator, every Sunday, department-based roles ---

SELECT_SUNSTOP_MONTH = 7


async def generate_sunstop_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Sun Stop Sundays — pick a month:", reply_markup=month_keyboard()
    )
    return SELECT_SUNSTOP_MONTH


async def select_sunstop_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year, month = map(int, query.data.split("-"))
    dates = dates_in_month(year, month, SERVICE_WEEKDAY["SunStopSundays"])

    await query.edit_message_text("Generating Sun Stop Sundays schedule...")

    ss = setup_sheet()
    already_filled = get_already_filled(ss.worksheet("SunStopSundays"), dates)
    counts = load_assignment_counts(ss)
    schedule_rows, counts = generate_simple_schedule(
        dates, get_random_roles_for_service(ss, "SunStopSundays"), counts, already_filled=already_filled
    )

    append_schedule_rows(ss, "SunStopSundays", schedule_rows, skip_keys=already_filled)
    save_assignment_counts(ss, counts)

    summary = format_schedule_summary("Sun Stop Sundays", schedule_rows)
    await query.message.reply_text(summary, parse_mode="Markdown")
    return ConversationHandler.END


# --- /log_tech: manual entry for Tech roles (never randomized) ---

TECH_SELECT_SERVICE, TECH_SELECT_ROLE, TECH_SELECT_PERIOD, TECH_PICK, TECH_CONFIRM_CONFLICT = range(8, 13)


async def log_tech_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in TECH_ROLES_BY_SERVICE]
    await update.message.reply_text(
        "Log Tech schedule — which service?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return TECH_SELECT_SERVICE


async def tech_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["tech_service"] = service
    roles = TECH_ROLES_BY_SERVICE[service]
    if len(roles) == 1:
        context.user_data["tech_role"] = next(iter(roles))
        return await tech_ask_period(query, context)
    buttons = [[InlineKeyboardButton(r, callback_data=r)] for r in roles]
    await query.edit_message_text(f"{service} — which Tech role?", reply_markup=InlineKeyboardMarkup(buttons))
    return TECH_SELECT_ROLE


async def tech_select_role(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["tech_role"] = query.data
    return await tech_ask_period(query, context)


async def tech_ask_period(query, context: ContextTypes.DEFAULT_TYPE):
    service = context.user_data["tech_service"]
    if TECH_CADENCE[service] == "week":
        mondays = upcoming_mondays()
        context.user_data["tech_mondays"] = mondays
        buttons = []
        for i, monday in enumerate(mondays):
            saturday = monday + dt.timedelta(days=5)
            buttons.append([InlineKeyboardButton(
                f"{monday.strftime('%b %d')} - {saturday.strftime('%b %d')}", callback_data=str(i)
            )])
        await query.edit_message_text("Pick a week:", reply_markup=InlineKeyboardMarkup(buttons))
    else:
        await query.edit_message_text("Pick a month:", reply_markup=month_keyboard())
    return TECH_SELECT_PERIOD


async def tech_select_period(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = context.user_data["tech_service"]
    role = context.user_data["tech_role"]
    if TECH_CADENCE[service] == "week":
        monday = context.user_data["tech_mondays"][int(query.data)]
        dates = dates_in_week(monday, PREDAWN_WEEKDAYS)
    else:
        year, month = map(int, query.data.split("-"))
        dates = dates_in_month(year, month, SERVICE_WEEKDAY[service])

    ss = setup_sheet()
    existing = get_role_assignments(ss, service, role, dates)
    pending = [d for d in dates if d.isoformat() not in existing]
    context.user_data["tech_ss"] = ss
    context.user_data["tech_pending"] = pending
    context.user_data["tech_new"] = []

    if not pending:
        await query.edit_message_text(f"{role} is already fully logged for that period.")
        return ConversationHandler.END
    return await tech_ask_next_date(query, context)


async def tech_ask_next_date(query, context: ContextTypes.DEFAULT_TYPE):
    pending = context.user_data["tech_pending"]
    if not pending:
        ss = context.user_data["tech_ss"]
        service = context.user_data["tech_service"]
        role = context.user_data["tech_role"]
        write_role_assignments(ss, service, role, context.user_data["tech_new"])
        lines = "\n".join(f"{d} - {p}" for d, p in context.user_data["tech_new"])
        await query.edit_message_text(f"{role} schedule saved:\n{lines}")
        return ConversationHandler.END

    d = pending[0]
    service = context.user_data["tech_service"]
    role = context.user_data["tech_role"]
    ss = context.user_data["tech_ss"]
    eligible = get_role_pool(ss, service, role)
    buttons = [[InlineKeyboardButton(name, callback_data=name)] for name in eligible]
    # No Live Broadcast option here — Tech stays blank until someone actually
    # submits their own schedule, never auto-filled.
    await query.edit_message_text(
        f"{role} for {d.strftime('%B %d, %Y')}?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return TECH_PICK


async def tech_pick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    name = query.data
    d = context.user_data["tech_pending"][0]
    ss = context.user_data["tech_ss"]
    service = context.user_data["tech_service"]
    role = context.user_data["tech_role"]
    date_str = d.isoformat()

    conflicts = find_conflicts(ss, service, date_str, name)
    if conflicts:
        context.user_data["tech_pending_name"] = name
        buttons = [
            [InlineKeyboardButton("Yes, assign anyway", callback_data="yes")],
            [InlineKeyboardButton("No, pick someone else", callback_data="no")],
        ]
        await query.edit_message_text(
            conflict_warning_text(name, date_str, conflicts, role), reply_markup=InlineKeyboardMarkup(buttons)
        )
        return TECH_CONFIRM_CONFLICT

    context.user_data["tech_pending"].pop(0)
    context.user_data["tech_new"].append((date_str, name))
    return await tech_ask_next_date(query, context)


async def tech_confirm_conflict(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "yes":
        d = context.user_data["tech_pending"].pop(0)
        name = context.user_data.pop("tech_pending_name")
        context.user_data["tech_new"].append((d.isoformat(), name))
        return await tech_ask_next_date(query, context)
    context.user_data.pop("tech_pending_name", None)
    return await tech_ask_next_date(query, context)


# --- /add_service: create a brand-new service type with custom roles ---
# --- /add_role: add one more role to an existing service (built-in or custom) ---
#
# Both write to the ServiceConfig tab. Once a service is in ServiceConfig,
# /generate_service (random-mode roles, equal-share) and /log_role
# (manual-mode roles, logged only — same as Tech/Preacher) can schedule it
# without any further code changes.

(ADD_SVC_NAME, ADD_SVC_CADENCE, ADD_SVC_ROLE_NAME, ADD_SVC_ELIGIBLE,
 ADD_SVC_MODE, ADD_SVC_MORE) = range(12, 18)

# New multi-phase /add_service states: define all roles first (name + mode),
# then collect eligible partakers per role, then offer to generate/upload.
(ADD_SVC_ROLE_NAME_ONLY, ADD_SVC_ROLE_MODE_ONLY, ADD_SVC_MORE_ROLES,
 ADD_SVC_ROLE_ELIGIBLE, ADD_SVC_READY, ADD_SVC_GEN_OR_UPLOAD, ADD_SVC_PERIOD) = range(95, 102)


async def add_service_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "What's the name of the new service? (e.g. 'Youth Fellowship')"
    )
    return ADD_SVC_NAME


async def add_service_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    context.user_data["new_service_name"] = name
    context.user_data["new_service_roles"] = []
    buttons = [
        [InlineKeyboardButton("Monthly", callback_data="month")],
        [InlineKeyboardButton("Weekly", callback_data="week")],
    ]
    await update.message.reply_text(
        f"How often does {name} happen?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return ADD_SVC_CADENCE


async def add_service_cadence(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["new_service_cadence"] = query.data
    context.user_data["new_service_roles_pending"] = []  # phase 1: (role_name, mode) only
    await query.edit_message_text("Role name? (e.g. 'Usher')")
    return ADD_SVC_ROLE_NAME_ONLY


# --- Phase 1: define every role's name + mode first, no partakers yet ---

async def add_service_role_name_only(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["current_role_name"] = update.message.text.strip()
    buttons = [
        [InlineKeyboardButton("Randomize (equal-share)", callback_data="random")],
        [InlineKeyboardButton("Manual entry only", callback_data="manual")],
    ]
    await update.message.reply_text(
        "Should this role be randomly/equal-share generated, or logged manually "
        "(like Preacher/Tech)?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return ADD_SVC_ROLE_MODE_ONLY


async def add_service_role_mode_only(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    role_name = context.user_data.pop("current_role_name")
    mode = query.data
    context.user_data["new_service_roles_pending"].append((role_name, mode))

    buttons = [
        [InlineKeyboardButton("Add another role", callback_data="more")],
        [InlineKeyboardButton("Done adding roles", callback_data="done")],
    ]
    await query.edit_message_text(
        f"Added role '{role_name}' ({mode}). Add another role for "
        f"{context.user_data['new_service_name']}?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return ADD_SVC_MORE_ROLES


async def add_service_more_roles(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "more":
        await query.edit_message_text("Role name? (e.g. 'Usher')")
        return ADD_SVC_ROLE_NAME_ONLY

    # done defining roles — move to phase 2: collect eligible partakers per role
    context.user_data["new_service_roles_final"] = []
    context.user_data["role_idx"] = 0
    return await add_service_ask_eligible(update, context)


# --- Phase 2: now go role by role and collect who's eligible for each ---

async def add_service_ask_eligible(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Shared by the 'Done adding roles' button and by the text-message path
    after each role's eligible list comes in — asks for the next role's
    partakers, or wraps up (saves config, asks to generate) once every role
    defined in phase 1 has its eligible list."""
    pending = context.user_data["new_service_roles_pending"]
    idx = context.user_data["role_idx"]
    send = update.callback_query.edit_message_text if update.callback_query else update.message.reply_text

    if idx >= len(pending):
        service = context.user_data["new_service_name"]
        cadence = context.user_data["new_service_cadence"]
        roles_final = context.user_data["new_service_roles_final"]
        ss = setup_sheet()
        context.user_data["add_svc_ss"] = ss
        add_service_config(ss, service, cadence, roles_final)

        role_summary = "\n".join(f"- {r} ({m}): {', '.join(e)}" for r, e, m in roles_final)
        buttons = [
            [InlineKeyboardButton("Yes", callback_data="yes")],
            [InlineKeyboardButton("Not yet", callback_data="no")],
        ]
        await send(
            f"'{service}' added ({cadence}).\n{role_summary}\n\nReady to generate a schedule now?",
            reply_markup=InlineKeyboardMarkup(buttons),
        )
        return ADD_SVC_READY

    role_name, _mode = pending[idx]
    await send(f"Who's eligible for '{role_name}'? Send names separated by commas.")
    return ADD_SVC_ROLE_ELIGIBLE


async def add_service_role_eligible(update: Update, context: ContextTypes.DEFAULT_TYPE):
    eligible = [n.strip() for n in update.message.text.split(",") if n.strip()]
    pending = context.user_data["new_service_roles_pending"]
    idx = context.user_data["role_idx"]
    role_name, mode = pending[idx]
    context.user_data["new_service_roles_final"].append((role_name, eligible, mode))
    context.user_data["role_idx"] = idx + 1
    return await add_service_ask_eligible(update, context)


# --- Once roles + partakers exist: ready to generate now, or upload instead? ---

async def add_service_ready_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = context.user_data["new_service_name"]
    if query.data == "no":
        await query.edit_message_text(
            f"'{service}' is all set up. Use /generate_service, /log_role, or /template whenever you're ready."
        )
        return ConversationHandler.END

    buttons = [
        [InlineKeyboardButton("Generate now (random pick)", callback_data="generate")],
        [InlineKeyboardButton("Upload an existing schedule", callback_data="upload")],
    ]
    await query.edit_message_text("How would you like to fill it in?", reply_markup=InlineKeyboardMarkup(buttons))
    return ADD_SVC_GEN_OR_UPLOAD


async def add_service_gen_or_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["add_svc_action"] = query.data
    ss = context.user_data["add_svc_ss"]
    service = context.user_data["new_service_name"]
    cadence = load_service_configs(ss)[service]["cadence"]

    if cadence == "week":
        mondays = upcoming_mondays()
        context.user_data["add_svc_mondays"] = mondays
        buttons = []
        for i, monday in enumerate(mondays):
            saturday = monday + dt.timedelta(days=6)
            buttons.append([InlineKeyboardButton(
                f"{monday.strftime('%b %d')} - {saturday.strftime('%b %d')}", callback_data=str(i)
            )])
        await query.edit_message_text("Which week?", reply_markup=InlineKeyboardMarkup(buttons))
    else:
        await query.edit_message_text("Which month?", reply_markup=month_keyboard())
    return ADD_SVC_PERIOD


async def add_service_period(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = context.user_data["add_svc_ss"]
    service = context.user_data["new_service_name"]
    configs = load_service_configs(ss)
    cadence = configs[service]["cadence"]

    if cadence == "week":
        monday = context.user_data["add_svc_mondays"][int(query.data)]
        dates = week_dates(service, monday)
    else:
        year, month = map(int, query.data.split("-"))
        # custom monthly services default to Sunday (see get_service_month_dates)
        dates = get_service_month_dates(ss, service, year, month)

    action = context.user_data["add_svc_action"]
    if action == "generate":
        random_roles = {r: c["eligible"] for r, c in configs[service]["roles"].items() if c["mode"] == "random"}
        if not random_roles:
            await query.edit_message_text(f"{service} has no randomized roles to generate — try /log_role instead.")
            return ConversationHandler.END
        await query.edit_message_text("Generating schedule...")
        already_filled = get_already_filled(ss.worksheet(service), dates)
        counts = load_assignment_counts(ss)
        rows, counts = generate_simple_schedule(dates, random_roles, counts, already_filled=already_filled)
        append_schedule_rows(ss, service, rows, skip_keys=already_filled)
        save_assignment_counts(ss, counts)
        summary = format_schedule_summary(service, rows)
        manual_roles = [r for r, c in configs[service]["roles"].items() if c["mode"] == "manual"]
        note = (f"\n\n(Manual-entry roles for this service — log via /log_role: {', '.join(manual_roles)})"
                if manual_roles else "")
        await query.message.reply_text(summary + note, parse_mode="Markdown")
    else:  # upload
        rows = build_template_rows(ss, service, dates)
        path = f"/tmp/{service}_template.csv"
        write_template_csv(path, rows)
        await query.edit_message_text("Here's your template — fill in the Partaker column and send it back to me.")
        with open(path, "rb") as f:
            await context.bot.send_document(chat_id=query.message.chat_id, document=f, filename=f"{service}_template.csv")
    return ConversationHandler.END


async def add_role_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    configs = load_service_configs(ss)
    if not configs:
        await update.message.reply_text("No services set up yet — use /add_service first.")
        return ConversationHandler.END
    context.user_data["add_role_ss"] = ss
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in configs]
    await update.message.reply_text(
        "Add a role to which service?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return ADD_ROLE_SELECT_SERVICE


ADD_ROLE_SELECT_SERVICE = 18


async def add_role_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["add_role_service"] = query.data
    await query.edit_message_text("Role name? (e.g. 'Usher')")
    return ADD_SVC_ROLE_NAME


async def add_service_role_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Single-role variant used only by /add_role (adding one role to an
    already-existing service) — /add_service uses the multi-role phase-1/
    phase-2 flow above instead."""
    context.user_data["current_role_name"] = update.message.text.strip()
    await update.message.reply_text(
        "Who's eligible for this role? Send names separated by commas."
    )
    return ADD_SVC_ELIGIBLE


async def add_service_eligible(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Pairs with add_service_role_name above, for /add_role only."""
    context.user_data["current_role_eligible"] = [n.strip() for n in update.message.text.split(",") if n.strip()]
    buttons = [
        [InlineKeyboardButton("Randomize (equal-share)", callback_data="random")],
        [InlineKeyboardButton("Manual entry only", callback_data="manual")],
    ]
    await update.message.reply_text(
        "Should this role be randomly/equal-share generated, or logged manually "
        "(like Preacher/Tech)?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return ADD_SVC_MODE


async def add_role_finish(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Terminal step for /add_role, reusing add_service_role_name/eligible/mode
    for the input steps, then writing just the one new role."""
    query = update.callback_query
    await query.answer()
    role_name = context.user_data.pop("current_role_name")
    eligible = context.user_data.pop("current_role_eligible")
    mode = query.data
    service = context.user_data["add_role_service"]

    ss = context.user_data.get("add_role_ss") or setup_sheet()
    add_role_to_service(ss, service, role_name, eligible, mode)
    await query.edit_message_text(f"Added '{role_name}' ({mode}) to {service}.")
    return ConversationHandler.END


# --- /generate_service: generic equal-share generator for any ServiceConfig service ---

GEN_SVC_SELECT_SERVICE, GEN_SVC_SELECT_PERIOD = range(20, 22)


async def generate_service_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    configs = load_service_configs(ss)
    # only offer services that have at least one random-mode role
    options = [s for s, cfg in configs.items() if any(r["mode"] == "random" for r in cfg["roles"].values())]
    if not options:
        await update.message.reply_text("No services with randomized roles found. Use /add_service first.")
        return ConversationHandler.END
    context.user_data["gen_svc_ss"] = ss
    context.user_data["gen_svc_configs"] = configs
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in options]
    await update.message.reply_text("Generate schedule for which service?", reply_markup=InlineKeyboardMarkup(buttons))
    return GEN_SVC_SELECT_SERVICE


async def generate_service_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["gen_svc_service"] = service
    cadence = context.user_data["gen_svc_configs"][service]["cadence"]
    if cadence == "week":
        mondays = upcoming_mondays()
        context.user_data["gen_svc_mondays"] = mondays
        buttons = []
        for i, monday in enumerate(mondays):
            saturday = monday + dt.timedelta(days=6)
            buttons.append([InlineKeyboardButton(
                f"{monday.strftime('%b %d')} - {saturday.strftime('%b %d')}", callback_data=str(i)
            )])
        await query.edit_message_text("Pick a week:", reply_markup=InlineKeyboardMarkup(buttons))
    else:
        await query.edit_message_text("Pick a month:", reply_markup=month_keyboard())
    return GEN_SVC_SELECT_PERIOD


async def generate_service_period(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = context.user_data["gen_svc_service"]
    cfg = context.user_data["gen_svc_configs"][service]
    ss = context.user_data["gen_svc_ss"]

    if cfg["cadence"] == "week":
        monday = context.user_data["gen_svc_mondays"][int(query.data)]
        dates = week_dates(service, monday)
    else:
        year, month = map(int, query.data.split("-"))
        # built-in services use their real weekday (e.g. FilipinoTranslation ->
        # Sunday); custom monthly services default to Sunday.
        dates = get_service_month_dates(ss, service, year, month)

    random_roles = {r: cfg["roles"][r]["eligible"] for r in cfg["roles"] if cfg["roles"][r]["mode"] == "random"}
    manual_roles = [r for r in cfg["roles"] if cfg["roles"][r]["mode"] == "manual"]

    # Manual-mode roles are informational only here (like Tech) — not a
    # generation blocker unless you want them to be; say so if that's wrong.
    await query.edit_message_text("Generating schedule...")
    already_filled = get_already_filled(ss.worksheet(service), dates)
    counts = load_assignment_counts(ss)
    unavailable = get_unavailability(ss, service, dates)

    # Filipino Translation: whoever is that Sunday's (English) Preacher can never
    # also be the Filipino Preacher, who translates the sermon online.
    never_by_role, soft_avoid_from_role, warnings = {}, {}, []
    if service == "FilipinoTranslation":
        sunday_preachers = get_role_assignments(ss, "Sunday", "Preacher", dates)
        never_by_role = {"Filipino Preacher": {
            d: {p} for d, p in sunday_preachers.items() if p and p != LIVE_BROADCAST
        }}
        unset = [d for d in dates if d.isoformat() not in sunday_preachers]
        if unset:
            warnings.append(
                "ℹ️ Sunday's Preacher isn't set yet for " + ", ".join(d.strftime("%b %d") for d in unset)
                + ", so I couldn't keep the Filipino Preacher different from the Preacher on those dates. "
                "Set the preachers first (/set_preacher)."
            )
        # Someone who is the Filipino Preacher on one Sunday is avoided (not
        # blocked) for Initial Proofreading on another — they should mainly
        # just preach; being 2nd PR sometimes is fine, so that pairing is left
        # to fairness alone.
        soft_avoid_from_role = {"Initial Proofreading": "Filipino Preacher"}

    schedule_rows, counts = generate_simple_schedule(
        dates, random_roles, counts, already_filled=already_filled, unavailable=unavailable,
        never_by_role=never_by_role, soft_avoid_from_role=soft_avoid_from_role,
    )
    append_schedule_rows(ss, service, schedule_rows, skip_keys=already_filled)
    save_assignment_counts(ss, counts)

    assigned = {(d, r) for d, r, _p in schedule_rows}
    missing = [f"{r} on {d.strftime('%b %d')}" for d in dates for r in random_roles if (d.isoformat(), r) not in assigned]
    if missing:
        warnings.append("⚠️ Nobody could be assigned: " + ", ".join(missing) + ".")

    summary = format_schedule_summary(service, schedule_rows)
    note = f"\n\n(Manual-entry roles for this service — log via /log_role: {', '.join(manual_roles)})" if manual_roles else ""
    await query.message.reply_text(summary + note, parse_mode="Markdown")
    for w in warnings:
        await query.message.reply_text(w)
    await warn_if_unavailable_scheduled(query.message, schedule_rows, unavailable)
    return ConversationHandler.END


# --- /log_role: generic manual entry for any manual-mode role in ServiceConfig ---

LOG_ROLE_SELECT_SERVICE, LOG_ROLE_SELECT_ROLE, LOG_ROLE_SELECT_PERIOD, LOG_ROLE_PICK, LOG_ROLE_CONFIRM_CONFLICT = range(22, 27)


async def log_role_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    configs = load_service_configs(ss)
    options = [s for s, cfg in configs.items() if any(r["mode"] == "manual" for r in cfg["roles"].values())]
    if not options:
        await update.message.reply_text("No manual-entry roles found.")
        return ConversationHandler.END
    context.user_data["log_role_ss"] = ss
    context.user_data["log_role_configs"] = configs
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in options]
    await update.message.reply_text("Log a role for which service?", reply_markup=InlineKeyboardMarkup(buttons))
    return LOG_ROLE_SELECT_SERVICE


async def log_role_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["log_role_service"] = service
    cfg = context.user_data["log_role_configs"][service]
    manual_roles = [r for r in cfg["roles"] if cfg["roles"][r]["mode"] == "manual"]
    if len(manual_roles) == 1:
        context.user_data["log_role_role"] = manual_roles[0]
        return await log_role_ask_period(query, context)
    buttons = [[InlineKeyboardButton(r, callback_data=r)] for r in manual_roles]
    await query.edit_message_text(f"{service} — which role?", reply_markup=InlineKeyboardMarkup(buttons))
    return LOG_ROLE_SELECT_ROLE


async def log_role_select_role(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["log_role_role"] = query.data
    return await log_role_ask_period(query, context)


async def log_role_ask_period(query, context: ContextTypes.DEFAULT_TYPE):
    service = context.user_data["log_role_service"]
    cadence = context.user_data["log_role_configs"][service]["cadence"]
    if cadence == "week":
        mondays = upcoming_mondays()
        context.user_data["log_role_mondays"] = mondays
        buttons = []
        for i, monday in enumerate(mondays):
            saturday = monday + dt.timedelta(days=6)
            buttons.append([InlineKeyboardButton(
                f"{monday.strftime('%b %d')} - {saturday.strftime('%b %d')}", callback_data=str(i)
            )])
        await query.edit_message_text("Pick a week:", reply_markup=InlineKeyboardMarkup(buttons))
    else:
        await query.edit_message_text("Pick a month:", reply_markup=month_keyboard())
    return LOG_ROLE_SELECT_PERIOD


async def log_role_select_period(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = context.user_data["log_role_service"]
    role = context.user_data["log_role_role"]
    cadence = context.user_data["log_role_configs"][service]["cadence"]
    ss = context.user_data["log_role_ss"]

    if cadence == "week":
        monday = context.user_data["log_role_mondays"][int(query.data)]
        dates = week_dates(service, monday)
    else:
        year, month = map(int, query.data.split("-"))
        # real weekday for built-ins (Wednesday -> Wednesdays), Sunday for custom
        dates = get_service_month_dates(ss, service, year, month)

    existing = get_role_assignments(ss, service, role, dates)
    pending = [d for d in dates if d.isoformat() not in existing]
    context.user_data["log_role_pending"] = pending
    context.user_data["log_role_new"] = []
    context.user_data["log_role_eligible"] = context.user_data["log_role_configs"][service]["roles"][role]["eligible"]

    if not pending:
        await query.edit_message_text(f"{role} is already fully logged for that period.")
        return ConversationHandler.END
    return await log_role_ask_next_date(query, context)


async def log_role_ask_next_date(query, context: ContextTypes.DEFAULT_TYPE):
    pending = context.user_data["log_role_pending"]
    if not pending:
        ss = context.user_data["log_role_ss"]
        service = context.user_data["log_role_service"]
        role = context.user_data["log_role_role"]
        write_role_assignments(ss, service, role, context.user_data["log_role_new"])
        lines = "\n".join(f"{d} - {p}" for d, p in context.user_data["log_role_new"])
        await query.edit_message_text(f"{role} schedule saved:\n{lines}")
        return ConversationHandler.END

    d = pending[0]
    eligible = context.user_data["log_role_eligible"]
    buttons = [[InlineKeyboardButton(name, callback_data=name)] for name in eligible]
    role = context.user_data["log_role_role"]
    buttons.extend(broadcast_button(context.user_data["log_role_service"]))
    await query.edit_message_text(
        f"{role} for {d.strftime('%B %d, %Y')}?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return LOG_ROLE_PICK


async def log_role_pick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    name = query.data
    d = context.user_data["log_role_pending"][0]
    ss = context.user_data["log_role_ss"]
    service = context.user_data["log_role_service"]
    role = context.user_data["log_role_role"]
    date_str = d.isoformat()

    conflicts = find_conflicts(ss, service, date_str, name)
    if conflicts:
        context.user_data["log_role_pending_name"] = name
        buttons = [
            [InlineKeyboardButton("Yes, assign anyway", callback_data="yes")],
            [InlineKeyboardButton("No, pick someone else", callback_data="no")],
        ]
        await query.edit_message_text(
            conflict_warning_text(name, date_str, conflicts, role), reply_markup=InlineKeyboardMarkup(buttons)
        )
        return LOG_ROLE_CONFIRM_CONFLICT

    context.user_data["log_role_pending"].pop(0)
    context.user_data["log_role_new"].append((date_str, name))
    return await log_role_ask_next_date(query, context)


async def log_role_confirm_conflict(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "yes":
        d = context.user_data["log_role_pending"].pop(0)
        name = context.user_data.pop("log_role_pending_name")
        context.user_data["log_role_new"].append((d.isoformat(), name))
        return await log_role_ask_next_date(query, context)
    context.user_data.pop("log_role_pending_name", None)
    return await log_role_ask_next_date(query, context)


# ---------------------------------------------------------------------------
# #2 Pull schedule per service, #3 Pull schedule per individual/department
# ---------------------------------------------------------------------------

BUILTIN_SCHEDULE_TABS = ["Sunday", "Wednesday", "Predawn", "SunStopSundays", "FilipinoTranslation"]


def get_all_schedule_tabs(ss):
    """Built-in tabs plus any custom services added via /add_service,
    de-duplicated and order-preserved."""
    custom = list(load_service_configs(ss).keys())
    return list(dict.fromkeys(BUILTIN_SCHEDULE_TABS + custom))


def get_roster_names(ss):
    roster_ws = ss.worksheet("Roster")
    return sorted(set(roster_ws.col_values(1)[1:]))


def rows_in_month(records, year, month):
    prefix = f"{year:04d}-{month:02d}"
    return [r for r in records if r.get("Date", "").startswith(prefix)]


def rows_in_dates(records, dates):
    date_strs = {d.isoformat() for d in dates}
    return [r for r in records if r.get("Date") in date_strs]


def nearest_date_rows(records):
    """All rows for the earliest date that is today or later. Empty list
    if nothing upcoming is scheduled."""
    today = today_local().isoformat()
    upcoming = sorted({r["Date"] for r in records if r.get("Date", "") >= today})
    if not upcoming:
        return []
    return [r for r in records if r["Date"] == upcoming[0]]


def records_to_rows(records):
    return [(r["Date"], r["Role"], r["Partaker"]) for r in records]


def person_schedule_by_service(ss, service_type, name):
    ws = ss.worksheet(service_type)
    records = ws.get_all_records()
    rows = [r for r in records if r.get("Partaker") == name]
    rows.sort(key=lambda r: r.get("Date", ""))
    return rows


def person_schedule_by_month(ss, all_tabs, year, month, name):
    """Returns list of (date_str, service, role) across every schedule tab."""
    prefix = f"{year:04d}-{month:02d}"
    results = []
    for service in all_tabs:
        ws = ss.worksheet(service)
        for r in ws.get_all_records():
            if r.get("Partaker") == name and r.get("Date", "").startswith(prefix):
                results.append((r["Date"], service, r["Role"]))
    results.sort()
    return results


def format_person_month_summary(name, results):
    if not results:
        return f"No assignments found for {name} that month."
    by_date = defaultdict(list)
    for date_str, service, role in results:
        by_date[date_str].append((service, role))
    lines = [f"*Schedule for {name}*\n"]
    for date_str in sorted(by_date):
        d = dt.date.fromisoformat(date_str)
        lines.append(f"*{d.strftime('%B %d, %Y')}*")
        for service, role in by_date[date_str]:
            lines.append(f"{service} - {role}")
        lines.append("")
    return "\n".join(lines).strip()


# --- /pull_schedule: per-service lookup (Nearest / This week / This month) ---

PULL_SELECT_SERVICE, PULL_SELECT_PERIOD, PULL_PICK_MONTH = range(30, 33)


async def pull_schedule_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["pull_ss"] = ss
    tabs = get_all_schedule_tabs(ss)
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
    await update.message.reply_text("Pull schedule for which service?", reply_markup=InlineKeyboardMarkup(buttons))
    return PULL_SELECT_SERVICE


async def pull_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["pull_service"] = service

    buttons = [[InlineKeyboardButton(f"Nearest {service}", callback_data="nearest")]]
    if service == "Predawn":
        buttons.append([InlineKeyboardButton("This week", callback_data="week")])
    buttons.append([InlineKeyboardButton("This month", callback_data="month")])
    buttons.append([InlineKeyboardButton("📅 Pick a month", callback_data="pick_month")])

    await query.edit_message_text(f"{service} — which period?", reply_markup=InlineKeyboardMarkup(buttons))
    return PULL_SELECT_PERIOD


async def pull_select_period(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "pick_month":
        service = context.user_data["pull_service"]
        await query.edit_message_text(f"{service} — which month?", reply_markup=month_keyboard())
        return PULL_PICK_MONTH

    service = context.user_data["pull_service"]
    ss = context.user_data["pull_ss"]
    ws = ss.worksheet(service)
    records = ws.get_all_records()

    if query.data == "nearest":
        rows = nearest_date_rows(records)
    elif query.data == "week":
        today = today_local()
        monday = today - dt.timedelta(days=today.weekday())
        dates = dates_in_week(monday, PREDAWN_WEEKDAYS)
        rows = rows_in_dates(records, dates)
    else:  # month
        today = today_local()
        rows = rows_in_month(records, today.year, today.month)

    if not rows:
        await query.edit_message_text(f"No {service} schedule found for that period.")
        return ConversationHandler.END

    summary = format_schedule_summary(service, records_to_rows(rows))
    await query.edit_message_text(summary, parse_mode="Markdown")
    return ConversationHandler.END


async def pull_pick_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year, month = map(int, query.data.split("-"))
    service = context.user_data["pull_service"]
    ss = context.user_data["pull_ss"]
    records = ss.worksheet(service).get_all_records()
    rows = rows_in_month(records, year, month)

    month_label = dt.date(year, month, 1).strftime("%B %Y")
    if not rows:
        await query.edit_message_text(f"No {service} schedule found for {month_label}.")
        return ConversationHandler.END

    summary = format_schedule_summary(f"{service} — {month_label}", records_to_rows(rows))
    await query.edit_message_text(summary, parse_mode="Markdown")
    return ConversationHandler.END


# --- /pull_person: per-individual or per-department lookup ---

PULL_PERSON_NAME, PULL_PERSON_MODE, PULL_PERSON_SERVICE, PULL_PERSON_MONTH = range(32, 36)


async def pull_person_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["pull_person_ss"] = ss
    names = get_roster_names(ss)
    buttons = [[InlineKeyboardButton(n, callback_data=n)] for n in names]
    await update.message.reply_text(
        "Pull schedule for which person/department?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return PULL_PERSON_NAME


async def pull_person_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["pull_person_name"] = query.data
    buttons = [
        [InlineKeyboardButton("Schedule Based on Service", callback_data="by_service")],
        [InlineKeyboardButton("Schedule for the Month", callback_data="by_month")],
    ]
    await query.edit_message_text(f"{query.data} — look up how?", reply_markup=InlineKeyboardMarkup(buttons))
    return PULL_PERSON_MODE


async def pull_person_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = context.user_data["pull_person_ss"]
    if query.data == "by_service":
        tabs = get_all_schedule_tabs(ss)
        buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
        await query.edit_message_text("Which service?", reply_markup=InlineKeyboardMarkup(buttons))
        return PULL_PERSON_SERVICE
    else:
        await query.edit_message_text("Which month?", reply_markup=month_keyboard())
        return PULL_PERSON_MONTH


async def pull_person_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = context.user_data["pull_person_ss"]
    name = context.user_data["pull_person_name"]
    service = query.data
    rows = person_schedule_by_service(ss, service, name)
    if not rows:
        await query.edit_message_text(f"No {service} assignments found for {name}.")
        return ConversationHandler.END
    summary = format_schedule_summary(f"{service} — {name}", records_to_rows(rows))
    await query.edit_message_text(summary, parse_mode="Markdown")
    return ConversationHandler.END


async def pull_person_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = context.user_data["pull_person_ss"]
    name = context.user_data["pull_person_name"]
    year, month = map(int, query.data.split("-"))
    all_tabs = get_all_schedule_tabs(ss)
    results = person_schedule_by_month(ss, all_tabs, year, month, name)
    summary = format_person_month_summary(name, results)
    await query.edit_message_text(summary, parse_mode="Markdown")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# #4 Day-before reminders to group chats
# ---------------------------------------------------------------------------

SERVICE_DISPLAY_NAME = {
    "Sunday": "Sunday Service",
    "Wednesday": "Wednesday Service",
    "Predawn": "Predawn Service",
    "SunStopSundays": "Sun Stop Sundays",
}

PREP_SCHEDULE_SUNDAY = (
    "8:30AM - Cleaning Time (All Members)\n"
    "9AM - Praise Practice \n"
    "9:15AM - Choir Practice\n"
    "9:20AM - Presider, Rep Prayer and Preacher Tech Run\n"
    "9:45AM - Prayer before the service"
)

# {pl} / {preacher} are filled in from that date's actual assignments.
PREP_SCHEDULE_WEDNESDAY_TEMPLATE = (
    "Cleaning Time - 6:30PM\n"
    "Tech Preparation - 7PM\n"
    "Prayer altogether - 7:20PM\n\n"
    "*before 7PM Rep Prayer should send the General Outline to {pl}\n"
    "Presider's notes to {preacher}"
)

REMINDER_FOOTER = (
    "would like to remind everyone for us to keep the schedule of service preparation "
    "including sending scripts, rep prayer, praise practice, Presider's part, Preacher, "
    "tech dry run and cleaning church ~ 😋🙂😄😘😘😘"
)

# Display label -> sheet Role name, in the order they should print.
# Sunday/Wednesday follow the exact abbreviations from the sample summary.
REMINDER_ROLE_ORDER = {
    "Sunday": [("Preacher", "Preacher"), ("Presider", "Presider"),
               ("Representative Prayer", "Representative Prayer"),
               ("Tech", "Onsite Tech"), ("PL", "Praise Leader")],
    "Wednesday": [("Preacher", "Preacher"), ("Presider", "Presider"),
                  ("Representative Prayer", "Representative Prayer"),
                  ("Tech", "Onsite Tech"), ("PL", "Praise Leader")],
}


def get_rows_for_date(ss, service_type, date):
    ws = ss.worksheet(service_type)
    date_str = date.isoformat()
    return [r for r in ws.get_all_records() if r.get("Date") == date_str]


def get_group_chat_id(ss, purpose):
    ws = ss.worksheet("GroupChats")
    for r in ws.get_all_records():
        if r.get("Purpose") == purpose:
            return r.get("ChatID") or None
    return None


def set_group_chat_id(ss, purpose, chat_id):
    ws = ss.worksheet("GroupChats")
    records = ws.get_all_records()
    for i, r in enumerate(records):
        if r.get("Purpose") == purpose:
            ws.update_cell(i + 2, 2, chat_id)  # +2: header row + 1-indexing
            return
    ws.append_rows([[purpose, chat_id]])


def format_daily_reminder(service_type, date, rows):
    """rows: list of {"Role":..., "Partaker":...} dicts for that single date.
    Sunday/Wednesday get the full prep-schedule block from the sample summary
    (with {pl}/{preacher} filled in for Wednesday's notes); Predawn and Sun
    Stop Sundays just list whatever roles are assigned, per your call to skip
    a prep block for those two."""
    lookup = {r["Role"]: r["Partaker"] for r in rows}
    header = f"{SERVICE_DISPLAY_NAME[service_type]}\n{date.strftime('%B %d, %Y')}\n\n"

    if service_type in REMINDER_ROLE_ORDER:
        lines = [f"{label} - {lookup.get(role, 'TBA')}" for label, role in REMINDER_ROLE_ORDER[service_type]]
        body = "\n".join(lines)
        if service_type == "Sunday":
            prep = PREP_SCHEDULE_SUNDAY
        else:
            prep = PREP_SCHEDULE_WEDNESDAY_TEMPLATE.format(
                pl=lookup.get("Praise Leader", "TBA"), preacher=lookup.get("Preacher", "TBA")
            )
        return f"{header}{body}\n\n{prep}\n\n{REMINDER_FOOTER}"

    if not lookup:
        body = "No roles assigned yet — please check the schedule."
    else:
        body = "\n".join(f"{role} - {partaker}" for role, partaker in lookup.items())
    return f"{header}{body}\n\nReminder: please prepare for tomorrow's service."


# NOTE: reminders are sent as plain text (no parse_mode). The Wednesday prep
# block contains a literal "*", which made Telegram's Markdown parser reject
# the whole message.

async def send_sunday_service_reminder(context: ContextTypes.DEFAULT_TYPE):
    """Runs daily at 8:30PM but only actually sends on Wednesday — reminds
    about the upcoming Sunday service (4 days out)."""
    now = dt.datetime.now(CHURCH_TZ)
    if now.weekday() != 2:  # Wednesday
        return
    ss = setup_sheet()
    chat_id = get_group_chat_id(ss, "Service Partakers")
    if not chat_id:
        return
    target = now.date() + dt.timedelta(days=4)
    rows = get_rows_for_date(ss, "Sunday", target)
    text = format_daily_reminder("Sunday", target, rows)
    await context.bot.send_message(chat_id=chat_id, text=text)


async def send_wednesday_service_reminder(context: ContextTypes.DEFAULT_TYPE):
    """Runs daily at 7PM but only actually sends on Sunday — reminds about
    the Wednesday service that follows (3 days out)."""
    now = dt.datetime.now(CHURCH_TZ)
    if now.weekday() != 6:  # Sunday
        return
    ss = setup_sheet()
    chat_id = get_group_chat_id(ss, "Service Partakers")
    if not chat_id:
        return
    target = now.date() + dt.timedelta(days=3)
    rows = get_rows_for_date(ss, "Wednesday", target)
    text = format_daily_reminder("Wednesday", target, rows)
    await context.bot.send_message(chat_id=chat_id, text=text)


async def send_friday_reminder(context: ContextTypes.DEFAULT_TYPE):
    """Runs daily at 8PM but only actually sends on Friday — reminds about
    the upcoming Sunday's Filipino Translation roles (to the Filipino
    Translators group) and Sun Stop Sundays roles (to Service Partakers),
    both 2 days out. Predawn has no reminder at all."""
    now = dt.datetime.now(CHURCH_TZ)
    if now.weekday() != 4:  # Friday
        return
    ss = setup_sheet()
    target = now.date() + dt.timedelta(days=2)

    fil_chat_id = get_group_chat_id(ss, "Filipino Translators")
    if fil_chat_id:
        rows = get_rows_for_date(ss, "FilipinoTranslation", target)
        body = ("\n".join(f"{r['Role']} - {r['Partaker']}" for r in rows)
                if rows else "No Filipino Translation roles assigned yet for this Sunday.")
        text = f"Filipino Translation Team\n{target.strftime('%B %d, %Y')}\n\n{body}"
        await context.bot.send_message(chat_id=fil_chat_id, text=text)

    partaker_chat_id = get_group_chat_id(ss, "Service Partakers")
    if partaker_chat_id:
        rows = get_rows_for_date(ss, "SunStopSundays", target)
        text = format_daily_reminder("SunStopSundays", target, rows)
        await context.bot.send_message(chat_id=partaker_chat_id, text=text)


# --- /set_group_chat: run inside a Telegram group to register it as the
# Service Partakers or Filipino Translators reminder destination ---

SET_GROUP_PURPOSE = 40


async def set_group_chat_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    context.user_data["group_chat_id"] = chat_id
    ss = setup_sheet()
    buttons = [
        [InlineKeyboardButton("Service Partakers (everyone)", callback_data="Service Partakers")],
        [InlineKeyboardButton("Filipino Translators", callback_data="Filipino Translators")],
    ] + [[InlineKeyboardButton(role, callback_data=role)] for role in role_group_purposes(ss)]
    await update.message.reply_text(
        "Register this group for which reminders? Pick 'Service Partakers' for a group covering "
        "everyone, or pick a role (e.g. Presider) if this group is only for people who do that role — "
        "it will then also get preference-round announcements for any service with that role.",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return SET_GROUP_PURPOSE


async def set_group_purpose(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    purpose = query.data
    chat_id = context.user_data["group_chat_id"]
    ss = setup_sheet()
    set_group_chat_id(ss, purpose, chat_id)
    if purpose in ("Service Partakers", "Filipino Translators"):
        await query.edit_message_text(f"This group is now registered for '{purpose}' reminders.")
    else:
        await query.edit_message_text(
            f"This group is now registered for the '{purpose}' role. It will get preference-round "
            f"announcements and close notices for any service that has a '{purpose}' role, alongside "
            f"the Service Partakers group."
        )
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# #5 Adjustments: swaps, substitutions, special requests
# ---------------------------------------------------------------------------

# --- Multiple group chats per service (e.g. a dedicated Presiders group,
# alongside the general Service Partakers group) ---
#
# A group is registered under a "Purpose": "Service Partakers" (everyone),
# "Filipino Translators" (that service's own group), or a role name (e.g.
# "Presider") for a group dedicated to people who do that one role. A
# preference round for a service notifies every group whose purpose matches:
# Service Partakers, plus any role of that service that has its own group.
# Filipino Translation keeps its single dedicated group instead (unchanged).

def role_group_purposes(ss):
    """Every role name across every service that actually collects preferences
    (built-in and custom, excluding Sun Stop Sundays — no round is ever opened
    for it — and Filipino Translation, which uses its own single group) —
    these are the extra "Purpose" options /set_group_chat can offer, one per
    role, e.g. "Presider", "Representative Prayer"."""
    services = [s for s in get_all_schedule_tabs(ss) if s not in ("SunStopSundays", "FilipinoTranslation")]
    return sorted({role for service in services for role in get_random_roles_for_service(ss, service)})


def collect_round_group_chat_ids(ss, service):
    """Every distinct, registered chat id that should hear about a preference
    round (or its close) for `service`, in a stable order: the general
    'Service Partakers' group, plus any group registered for one of that
    service's randomly-picked roles (e.g. a dedicated Presiders group).
    Filipino Translation uses only its own 'Filipino Translators' group."""
    if service == "FilipinoTranslation":
        purposes = ["Filipino Translators"]
    else:
        purposes = ["Service Partakers"] + list(get_random_roles_for_service(ss, service).keys())
    seen, ids = set(), []
    for purpose in purposes:
        chat_id = get_group_chat_id(ss, purpose)
        if chat_id and chat_id not in seen:
            seen.add(chat_id)
            ids.append(chat_id)
    return ids


def split_chat_ids(raw):
    """Parses a PreferenceRounds.GroupChatID cell back into a list of chat ids.
    Works whether it holds one id (older rounds) or several, comma-separated."""
    return [c.strip() for c in str(raw or "").split(",") if c.strip()]


def get_role_pool(ss, service_type, role):
    """Eligible list for any role, built-in or custom, random or manual.
    ServiceConfig (the sheet) is checked first since roster management
    (#6) edits eligibility there — it's the live source of truth once
    setup_sheet() has seeded it. The hardcoded dicts are only a fallback
    for the (normally unreachable) case a service isn't in the sheet yet."""
    configs = load_service_configs(ss)
    if service_type in configs and role in configs[service_type]["roles"]:
        return configs[service_type]["roles"][role]["eligible"]
    if service_type in ROLE_SETS and role in ROLE_SETS[service_type]:
        return ROLE_SETS[service_type][role]
    if service_type == "SunStopSundays" and role in SUN_STOP_ROLES:
        return SUN_STOP_ROLES[role]
    if role == "Preacher" and service_type in PREACHER_ELIGIBLE:
        return PREACHER_ELIGIBLE[service_type]
    if service_type in TECH_ROLES_BY_SERVICE and role in TECH_ROLES_BY_SERVICE[service_type]:
        return TECH_ROLES_BY_SERVICE[service_type][role]
    return []


def get_random_roles_for_service(ss, service_type):
    """{role: eligible_list} for only the random/equal-share roles of a
    service — built-in or custom. Config-first, same reasoning as
    get_role_pool above."""
    configs = load_service_configs(ss)
    if service_type in configs:
        return {r: c["eligible"] for r, c in configs[service_type]["roles"].items() if c["mode"] == "random"}
    if service_type in ROLE_SETS:
        return dict(ROLE_SETS[service_type])
    if service_type == "SunStopSundays":
        return dict(SUN_STOP_ROLES)
    return {}


def get_distinct_roles(ws):
    return sorted({r["Role"] for r in ws.get_all_records()})


def get_role_dates(ws, role):
    return sorted({r["Date"] for r in ws.get_all_records() if r["Role"] == role})


def get_already_filled(ws, dates, exclude_roles=()):
    """{(date_str, role): partaker} for every row already present on these
    dates — used so generation skips slots claimed in a preference round
    (or any other manual write) instead of overwriting them. exclude_roles
    lets callers keep a role (e.g. "Preacher") out of this dict when it's
    already handled by its own dedicated precondition/exclusion logic."""
    date_strs = {d.isoformat() for d in dates}
    return {
        (r["Date"], r["Role"]): r["Partaker"]
        for r in ws.get_all_records()
        if r.get("Date") in date_strs and r.get("Role") not in exclude_roles
    }


def get_person_roles_on_date(ss, service, date_str, person, ignore_role=None):
    """Every role `person` already holds on `date_str` within this service's
    own tab (Tech and regular partaker roles live in the same tab, so this
    single check covers Tech-vs-partaker crossover as well as any general
    double-role conflict)."""
    ws = ss.worksheet(service)
    return [
        r["Role"] for r in ws.get_all_records()
        if r.get("Date") == date_str and r.get("Partaker") == person and r.get("Role") != ignore_role
    ]


def get_cross_service_conflicts(ss, service, date_str, person):
    """The one cross-tab relationship in this system: Filipino Translation
    roles and Sunday roles both apply to the same date. Returns roles held
    in the OTHER tab."""
    other = {"Sunday": "FilipinoTranslation", "FilipinoTranslation": "Sunday"}.get(service)
    if not other:
        return []
    ws = ss.worksheet(other)
    return [r["Role"] for r in ws.get_all_records() if r.get("Date") == date_str and r.get("Partaker") == person]


# The English Preacher delivers the sermon; the Filipino Preacher translates it
# online at the same time. They must be two different people on the same Sunday.
# Unlike other double roles (which only trigger a warning), this pairing is
# never allowed.
FORBIDDEN_SAME_DAY = {
    ("Sunday", "Preacher"): ("FilipinoTranslation", "Filipino Preacher"),
    ("FilipinoTranslation", "Filipino Preacher"): ("Sunday", "Preacher"),
}


def forbidden_conflict(ss, service, role, date_str, person):
    """If `person` already holds the role that must never be combined with
    `role` on `date_str` (it lives in the other service's tab), returns that
    role's name, otherwise None."""
    pair = FORBIDDEN_SAME_DAY.get((service, role))
    if not pair or not person or person == LIVE_BROADCAST:
        return None
    other_service, other_role = pair
    for r in ss.worksheet(other_service).get_all_records():
        if r.get("Date") == date_str and r.get("Role") == other_role and r.get("Partaker") == person:
            return other_role
    return None


def preacher_pair_clashes(ss, dates):
    """[(date, person)] where the same person is both Sunday's Preacher and the
    Filipino Preacher on that date."""
    date_set = {d if isinstance(d, str) else d.isoformat() for d in dates}
    preachers = {r["Date"]: r["Partaker"] for r in ss.worksheet("Sunday").get_all_records()
                 if r.get("Role") == "Preacher" and r.get("Date") in date_set}
    out = []
    for r in ss.worksheet("FilipinoTranslation").get_all_records():
        person = r.get("Partaker")
        if (r.get("Role") == "Filipino Preacher" and r.get("Date") in date_set and person
                and person != LIVE_BROADCAST and preachers.get(r["Date"]) == person):
            out.append((r["Date"], person))
    return out


def find_conflicts(ss, service, date_str, person, ignore_role=None):
    """All roles `person` already holds on `date_str` that would conflict
    with giving them one more — same-tab (covers Tech-vs-partaker and any
    double-role case) plus the Sunday<->FilipinoTranslation cross-tab case.
    LIVE_BROADCAST is exempt — it's expected on every role the same date."""
    if person == LIVE_BROADCAST:
        return []
    return get_person_roles_on_date(ss, service, date_str, person, ignore_role) + \
        get_cross_service_conflicts(ss, service, date_str, person)


def conflict_warning_text(name, date_str, conflicts, new_role):
    return (
        f"⚠️ {name} already has {', '.join(conflicts)} on {date_str}. "
        f"Assign {new_role} too?"
    )


def get_role_row(ws, role, date_str):
    """Returns (sheet_row_number, current_partaker) or (None, None)."""
    for i, r in enumerate(ws.get_all_records()):
        if r["Role"] == role and r["Date"] == date_str:
            return i + 2, r["Partaker"]
    return None, None


def find_date_taken(ws, date_str, exclude_row=None):
    """Everyone already assigned some role on this date, for exclusion when
    picking a substitute — excludes the row being replaced itself."""
    taken = set()
    for i, r in enumerate(ws.get_all_records()):
        if r.get("Date") == date_str and (i + 2) != exclude_row:
            taken.add(r.get("Partaker"))
    return taken


def find_person_month_entries(ws, year, month, person, roles):
    """(row_number, record) pairs for a person's assignments in the given
    roles during that month, sorted earliest-first."""
    prefix = f"{year:04d}-{month:02d}"
    entries = []
    for i, r in enumerate(ws.get_all_records()):
        if r.get("Date", "").startswith(prefix) and r.get("Partaker") == person and r.get("Role") in roles:
            entries.append((i + 2, r))
    entries.sort(key=lambda e: e[1]["Date"])
    return entries


def log_adjustment(ss, service, date_str, role, old_partaker, new_partaker, adj_type, reason=""):
    ws = ss.worksheet("AdjustmentLog")
    ws.append_rows([[
        dt.datetime.now(CHURCH_TZ).isoformat(), service, date_str, role,
        old_partaker, new_partaker, adj_type, reason,
    ]])


async def announce_update(context, ss, text):
    """Posts the updated summary to the Service Partakers group if one is
    registered; the admin always gets it in their own reply regardless."""
    chat_id = get_group_chat_id(ss, "Service Partakers")
    if chat_id:
        await context.bot.send_message(chat_id=chat_id, text=text, parse_mode="Markdown")


# --- Double-role safety: used by /swap (and the Dashboard / CSV paths) ---
#
# The generator never gives the Preacher another role on the same date, but
# once a schedule exists, changing a Preacher (or anyone) by swap can put a
# person on two roles the same day. plan_swap() spots that BEFORE the swap and
# proposes moving the person's other (randomly-picked) role to a free partaker.

def apply_reassignment(ss, service, role, date_str, new_partaker, adj_type="swap_move", reason=""):
    """Changes who holds `role` on `date_str`: updates the sheet, logs it, and
    keeps the equal-share counts right (randomly-picked roles only). Returns
    the previous holder, or None if that slot doesn't exist."""
    ws = ss.worksheet(service)
    row_num, old_partaker = get_role_row(ws, role, date_str)
    if row_num is None:
        return None
    ws.update_cell(row_num, 3, new_partaker)
    log_adjustment(ss, service, date_str, role, old_partaker, new_partaker, adj_type, reason)
    if role in get_random_roles_for_service(ss, service) and new_partaker != LIVE_BROADCAST:
        counts = load_assignment_counts(ss)
        if old_partaker and old_partaker != LIVE_BROADCAST:
            counts[old_partaker] = max(0, counts[old_partaker] - 1)
        counts[new_partaker] += 1
        save_assignment_counts(ss, counts)
    return old_partaker


def plan_swap(ss, service, role, date_a, date_b, partaker_a, partaker_b):
    """Works out what a swap of `role` between date_a and date_b would cause.
    Returns {"role", "conflicts", "moves", "unmovable"}:
      conflicts - [{"person", "date", "roles"}]: someone who would end up with
                  another role on the same date
      moves     - [{"date", "role", "old", "new"}]: proposed fix, giving that
                  other role to a free, eligible person (fewest roles first)
      unmovable - conflicts that can't be fixed automatically (Preacher/Tech/
                  hand-picked roles, roles in the other service's tab, or
                  nobody free)"""
    plan = {"role": role, "conflicts": [], "moves": [], "unmovable": [], "blocked": []}
    if not partaker_a or not partaker_b or partaker_a == partaker_b:
        return plan

    random_roles = get_random_roles_for_service(ss, service)
    counts = load_assignment_counts(ss)
    chosen_on = defaultdict(set)  # date -> people already picked as a fix in this plan

    # the person arriving on each date, and the date they arrive on
    for person, target in ((partaker_b, date_a), (partaker_a, date_b)):
        if person == LIVE_BROADCAST:
            continue
        clash = forbidden_conflict(ss, service, role, target, person)
        if clash:
            plan["blocked"].append({"person": person, "date": target, "other": clash})
            continue
        same_tab = get_person_roles_on_date(ss, service, target, person, ignore_role=role)
        cross_tab = get_cross_service_conflicts(ss, service, target, person)
        if not same_tab and not cross_tab:
            continue
        plan["conflicts"].append({"person": person, "date": target, "roles": same_tab + cross_tab})

        for other_role in same_tab:
            if other_role not in random_roles:
                plan["unmovable"].append({"person": person, "date": target, "role": other_role,
                                          "reason": "it is not an automatically picked role"})
                continue
            candidates = [c for c in available_replacements(ss, service, target, other_role, person)
                          if c not in chosen_on[target]]
            if not candidates:
                plan["unmovable"].append({"person": person, "date": target, "role": other_role,
                                          "reason": "nobody else eligible is free that day"})
                continue
            fewest = min(counts[c] for c in candidates)
            new_person = random.choice([c for c in candidates if counts[c] == fewest])
            chosen_on[target].add(new_person)
            plan["moves"].append({"date": target, "role": other_role, "old": person, "new": new_person})
        for other_role in cross_tab:
            plan["unmovable"].append({"person": person, "date": target, "role": other_role,
                                      "reason": "it belongs to the other service's schedule"})
    return plan


def describe_swap_plan(plan):
    if plan["blocked"]:
        return "\n".join(
            f"⛔ {b['person']} is the {b['other']} on {b['date']}. The Preacher and the Filipino Preacher "
            f"must be different people, so this swap isn't allowed."
            for b in plan["blocked"]
        )
    lines = ["⚠️ Double role warning:"]
    for c in plan["conflicts"]:
        lines.append(f"- {c['person']} would take {plan['role']} on {c['date']} but already has "
                     f"{', '.join(c['roles'])} that day.")
    if plan["moves"] and not plan["unmovable"]:
        lines += ["", "Suggested fix (Swap + move their other role):"]
        for m in plan["moves"]:
            lines.append(f"- {m['date']} {m['role']}: {m['old']} -> {m['new']}")
    for u in plan["unmovable"]:
        lines.append(f"- Can't move {u['person']}'s {u['role']} on {u['date']}: {u['reason']}.")
    return "\n".join(lines)


def find_double_bookings(ss, service, dates):
    """[(date, person, [roles])] for anyone holding 2+ roles on the same date
    in this service. 'Live broadcast' rows are ignored."""
    date_set = set(dates)
    held = defaultdict(list)
    for r in ss.worksheet(service).get_all_records():
        person = r.get("Partaker")
        if r.get("Date") in date_set and person and person != LIVE_BROADCAST:
            held[(r["Date"], person)].append(r["Role"])
    return [(d, p, roles) for (d, p), roles in sorted(held.items()) if len(roles) > 1]


def double_booking_warnings(ss, service, changed):
    """changed: [(date_str, role, partaker)] just written by a sync/upload.
    Returns readable lines for people those edits left with two roles that day
    (older, untouched double roles are not repeated)."""
    if not changed:
        return []
    touched = {(d, p) for d, _role, p in changed}
    lines = [f"{d}: {p} has {' + '.join(roles)}"
             for d, p, roles in find_double_bookings(ss, service, {d for d, _r, _p in changed})
             if (d, p) in touched]
    if service in ("Sunday", "FilipinoTranslation"):
        lines += [f"{d}: {p} is both the Preacher and the Filipino Preacher (not allowed)"
                  for d, p in preacher_pair_clashes(ss, {d for d, _r, _p in changed}) if (d, p) in touched]
    return lines


# --- /swap: two dates trade partakers for the same role ---

SWAP_SERVICE, SWAP_ROLE, SWAP_DATE_A, SWAP_DATE_B, SWAP_CONFIRM = range(50, 55)


async def swap_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["swap_ss"] = ss
    tabs = get_all_schedule_tabs(ss)
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
    await update.message.reply_text("Swap dates for which service?", reply_markup=InlineKeyboardMarkup(buttons))
    return SWAP_SERVICE


async def swap_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["swap_service"] = service
    ws = context.user_data["swap_ss"].worksheet(service)
    roles = get_distinct_roles(ws)
    buttons = [[InlineKeyboardButton(r, callback_data=r)] for r in roles]
    await query.edit_message_text(f"{service} — which role?", reply_markup=InlineKeyboardMarkup(buttons))
    return SWAP_ROLE


async def swap_select_role(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    role = query.data
    context.user_data["swap_role"] = role
    ws = context.user_data["swap_ss"].worksheet(context.user_data["swap_service"])
    dates = get_role_dates(ws, role)
    if len(dates) < 2:
        await query.edit_message_text(f"Need at least 2 scheduled dates for {role} to swap.")
        return ConversationHandler.END
    context.user_data["swap_dates"] = dates
    buttons = [[InlineKeyboardButton(d, callback_data=d)] for d in dates]
    await query.edit_message_text("First date?", reply_markup=InlineKeyboardMarkup(buttons))
    return SWAP_DATE_A


async def swap_select_date_a(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["swap_date_a"] = query.data
    remaining = [d for d in context.user_data["swap_dates"] if d != query.data]
    buttons = [[InlineKeyboardButton(d, callback_data=d)] for d in remaining]
    await query.edit_message_text("Swap with which date?", reply_markup=InlineKeyboardMarkup(buttons))
    return SWAP_DATE_B


async def swap_select_date_b(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    date_b = query.data
    context.user_data["swap_date_b"] = date_b
    ss = context.user_data["swap_ss"]
    service = context.user_data["swap_service"]
    role = context.user_data["swap_role"]
    date_a = context.user_data["swap_date_a"]
    ws = ss.worksheet(service)
    _, partaker_a = get_role_row(ws, role, date_a)
    _, partaker_b = get_role_row(ws, role, date_b)
    plan = plan_swap(ss, service, role, date_a, date_b, partaker_a, partaker_b)

    text = (
        f"Swap {role}:\n{date_a}: {partaker_a}\n{date_b}: {partaker_b}\n\n"
        f"After swap: {date_a} -> {partaker_b}, {date_b} -> {partaker_a}"
    )
    if plan["blocked"]:
        text += "\n\n" + describe_swap_plan(plan)
        buttons = [[InlineKeyboardButton("Cancel", callback_data="no")]]
    elif not plan["conflicts"]:
        buttons = [
            [InlineKeyboardButton("Confirm swap", callback_data="yes")],
            [InlineKeyboardButton("Cancel", callback_data="no")],
        ]
    else:
        text += "\n\n" + describe_swap_plan(plan)
        buttons = []
        if plan["moves"] and not plan["unmovable"]:
            buttons.append([InlineKeyboardButton("Swap + move their other role", callback_data="move")])
        buttons.append([InlineKeyboardButton("Swap anyway (keep both roles)", callback_data="keep")])
        buttons.append([InlineKeyboardButton("Cancel", callback_data="no")])

    await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(buttons))
    return SWAP_CONFIRM


async def swap_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    choice = query.data  # "yes" (no conflicts), "keep", "move", or anything else = cancel
    if choice not in ("yes", "keep", "move"):
        await query.edit_message_text("Swap cancelled.")
        return ConversationHandler.END

    ss = context.user_data["swap_ss"]
    service = context.user_data["swap_service"]
    role = context.user_data["swap_role"]
    date_a = context.user_data["swap_date_a"]
    date_b = context.user_data["swap_date_b"]
    ws = ss.worksheet(service)

    row_a, partaker_a = get_role_row(ws, role, date_a)
    row_b, partaker_b = get_role_row(ws, role, date_b)

    # re-check right now: the sheet may have changed since the screen was shown
    plan = plan_swap(ss, service, role, date_a, date_b, partaker_a, partaker_b)
    if plan["blocked"]:
        await query.edit_message_text(describe_swap_plan(plan) + "\n\nNothing was changed.")
        return ConversationHandler.END
    moves = []
    if choice == "move":
        if plan["unmovable"]:
            await query.edit_message_text(
                "Things changed while you were deciding, so their other role can no longer be moved "
                "automatically. Nothing was changed — please run /swap again."
            )
            return ConversationHandler.END
        moves = plan["moves"]

    ws.update_cell(row_a, 3, partaker_b)
    ws.update_cell(row_b, 3, partaker_a)
    log_adjustment(ss, service, date_a, role, partaker_a, partaker_b, "swap")
    log_adjustment(ss, service, date_b, role, partaker_b, partaker_a, "swap")

    text = (
        f"Swap confirmed for {service} {role}:\n"
        f"{date_a}: {partaker_a} -> {partaker_b}\n"
        f"{date_b}: {partaker_b} -> {partaker_a}"
    )
    for m in moves:
        apply_reassignment(ss, service, m["role"], m["date"], m["new"], adj_type="swap_move",
                           reason=f"{m['old']} took {role} that day")
    if moves:
        text += "\n\nOther role moved because of the swap:\n" + "\n".join(
            f"{m['date']} {m['role']}: {m['old']} -> {m['new']}" for m in moves
        )
    elif plan["conflicts"]:
        text += "\n\nNote: " + "; ".join(
            f"{c['person']} has {', '.join(c['roles'])} and {role} on {c['date']}" for c in plan["conflicts"]
        )
    await query.edit_message_text(text)
    await announce_update(context, ss, text)
    return ConversationHandler.END


# --- /substitute: replace one partaker with another for one date+role ---

SUB_SERVICE, SUB_ROLE, SUB_DATE, SUB_NEW, SUB_CONFIRM_CONFLICT = range(55, 60)


async def substitute_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["sub_ss"] = ss
    tabs = get_all_schedule_tabs(ss)
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
    await update.message.reply_text("Substitute for which service?", reply_markup=InlineKeyboardMarkup(buttons))
    return SUB_SERVICE


async def substitute_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["sub_service"] = service
    ws = context.user_data["sub_ss"].worksheet(service)
    roles = get_distinct_roles(ws)
    buttons = [[InlineKeyboardButton(r, callback_data=r)] for r in roles]
    await query.edit_message_text(f"{service} — which role?", reply_markup=InlineKeyboardMarkup(buttons))
    return SUB_ROLE


async def substitute_select_role(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    role = query.data
    context.user_data["sub_role"] = role
    ws = context.user_data["sub_ss"].worksheet(context.user_data["sub_service"])
    dates = get_role_dates(ws, role)
    if not dates:
        await query.edit_message_text(f"No scheduled dates found for {role}.")
        return ConversationHandler.END
    buttons = [[InlineKeyboardButton(d, callback_data=d)] for d in dates]
    await query.edit_message_text("Which date?", reply_markup=InlineKeyboardMarkup(buttons))
    return SUB_DATE


async def substitute_select_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    date_str = query.data
    context.user_data["sub_date"] = date_str
    ss = context.user_data["sub_ss"]
    service = context.user_data["sub_service"]
    role = context.user_data["sub_role"]
    ws = ss.worksheet(service)
    _, current = get_role_row(ws, role, date_str)
    candidates = available_replacements(ss, service, date_str, role, current)
    buttons = [[InlineKeyboardButton(name, callback_data=name)] for name in candidates]
    if current != LIVE_BROADCAST and service != "SunStopSundays":
        buttons.extend(broadcast_button(service))
    if not buttons:
        await query.edit_message_text(
            f"No eligible, available replacement is listed for {role} on {date_str} "
            "based on preferences and existing assignments."
        )
        return ConversationHandler.END
    await query.edit_message_text(
        f"{role} on {date_str} is currently {current}. Choose an available replacement:",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return SUB_NEW


async def do_substitute(query, context, ss, service, role, date_str, new_partaker, adj_type="substitution"):
    ws = ss.worksheet(service)
    row_num, old_partaker = get_role_row(ws, role, date_str)
    ws.update_cell(row_num, 3, new_partaker)
    log_adjustment(ss, service, date_str, role, old_partaker, new_partaker, adj_type)

    random_roles = get_random_roles_for_service(ss, service)
    if role in random_roles and new_partaker != LIVE_BROADCAST:
        counts = load_assignment_counts(ss)
        if old_partaker != LIVE_BROADCAST:
            counts[old_partaker] = max(0, counts[old_partaker] - 1)
        counts[new_partaker] += 1
        save_assignment_counts(ss, counts)

    if adj_type == "cancellation":
        text = f"{service} {role} on {date_str}: {old_partaker} cancelled — {new_partaker} will cover."
    else:
        text = f"{service} {role} on {date_str}: {old_partaker} -> {new_partaker}"
    await query.edit_message_text(text)
    await announce_update(context, ss, text)


async def substitute_pick_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    new_partaker = query.data
    ss = context.user_data["sub_ss"]
    service = context.user_data["sub_service"]
    role = context.user_data["sub_role"]
    date_str = context.user_data["sub_date"]

    conflicts = find_conflicts(ss, service, date_str, new_partaker, ignore_role=role)
    if conflicts:
        context.user_data["sub_new_partaker"] = new_partaker
        buttons = [
            [InlineKeyboardButton("Yes, proceed", callback_data="yes")],
            [InlineKeyboardButton("No, cancel", callback_data="no")],
        ]
        await query.edit_message_text(
            conflict_warning_text(new_partaker, date_str, conflicts, role), reply_markup=InlineKeyboardMarkup(buttons)
        )
        return SUB_CONFIRM_CONFLICT

    await do_substitute(query, context, ss, service, role, date_str, new_partaker)
    return ConversationHandler.END


async def substitute_confirm_conflict(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data != "yes":
        await query.edit_message_text("Substitution cancelled.")
        return ConversationHandler.END

    ss = context.user_data["sub_ss"]
    service = context.user_data["sub_service"]
    role = context.user_data["sub_role"]
    date_str = context.user_data["sub_date"]
    new_partaker = context.user_data.pop("sub_new_partaker")
    await do_substitute(query, context, ss, service, role, date_str, new_partaker)
    return ConversationHandler.END


# --- /cancel_role: a partaker drops a role on one date; the bot lists who can cover ---
# Replacement candidates are people eligible for that role who did NOT mark the
# date as unavailable and aren't already scheduled that day. Nothing changes
# until the partaker taps a name and confirms — they're expected to ask that
# person first.

CANCEL_SERVICE, CANCEL_NAME, CANCEL_PICK, CANCEL_REPLACEMENT, CANCEL_CONFIRM = range(180, 185)


def upcoming_assignments(ws, name=None):
    """[(date_str, role, partaker)] from today on, earliest first. Optionally
    only for one person. Live Broadcast rows are skipped."""
    today = today_local().isoformat()
    out = []
    for r in ws.get_all_records():
        date_str, person = str(r.get("Date", "")), r.get("Partaker")
        if len(date_str) == 10 and date_str >= today and person and person != LIVE_BROADCAST:
            if name is None or person == name:
                out.append((date_str, r["Role"], person))
    return sorted(out)


def available_replacements(ss, service, date_str, role, current):
    """People who could take `role` on `date_str`: in the role's pool, not the
    person cancelling, not marked unavailable, and not already scheduled
    that day (this service, plus the Sunday <-> Filipino Translation link)."""
    pool = get_role_pool(ss, service, role)
    unavailable = get_unavailability(ss, service, [date_str]).get(date_str, set())
    busy = {
        r["Partaker"] for r in ss.worksheet(service).get_all_records()
        if r.get("Date") == date_str and r.get("Role") != role
    }
    other = {"Sunday": "FilipinoTranslation", "FilipinoTranslation": "Sunday"}.get(service)
    if other:
        busy |= {r["Partaker"] for r in ss.worksheet(other).get_all_records() if r.get("Date") == date_str}
    return [
        p for p in pool
        if p not in (current, LIVE_BROADCAST) and p not in unavailable and p not in busy
    ]


async def cancel_role_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["cr_ss"] = ss
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in get_all_schedule_tabs(ss)]
    await update.message.reply_text("Cancel a role — which service?", reply_markup=InlineKeyboardMarkup(buttons))
    return CANCEL_SERVICE


async def cancel_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = context.user_data["cr_ss"]
    service = query.data
    context.user_data["cr_service"] = service
    names = sorted({p for _, _, p in upcoming_assignments(ss.worksheet(service))})
    if not names:
        await query.edit_message_text(f"No upcoming {service} assignments found.")
        return ConversationHandler.END
    buttons = [[InlineKeyboardButton(n, callback_data=n)] for n in names]
    await query.edit_message_text(f"{service} — which name is yours?", reply_markup=InlineKeyboardMarkup(buttons))
    return CANCEL_NAME


async def cancel_select_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss, service = context.user_data["cr_ss"], context.user_data["cr_service"]
    name = query.data
    context.user_data["cr_name"] = name
    entries = upcoming_assignments(ss.worksheet(service), name)
    if not entries:
        await query.edit_message_text(f"{name} has no upcoming {service} roles.")
        return ConversationHandler.END
    context.user_data["cr_entries"] = entries
    buttons = [
        [InlineKeyboardButton(f"{dt.date.fromisoformat(d).strftime('%a %b %d')} — {role}", callback_data=str(i))]
        for i, (d, role, _) in enumerate(entries)
    ]
    await query.edit_message_text(
        f"{name}, which role do you want to cancel?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return CANCEL_PICK


async def cancel_pick_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss, service = context.user_data["cr_ss"], context.user_data["cr_service"]
    name = context.user_data["cr_name"]
    date_str, role, _ = context.user_data["cr_entries"][int(query.data)]
    context.user_data["cr_date"], context.user_data["cr_role"] = date_str, role

    candidates = available_replacements(ss, service, date_str, role, name)
    pretty = dt.date.fromisoformat(date_str).strftime("%A, %B %d")
    if not candidates:
        await query.edit_message_text(
            f"Nobody else is available for {role} on {pretty} based on the submitted preferences "
            f"and who's already scheduled that day.\nPlease talk to the admin — they can pick anyone with /substitute."
        )
        return ConversationHandler.END

    buttons = [[InlineKeyboardButton(p, callback_data=p)] for p in candidates]
    buttons.append([InlineKeyboardButton("Never mind, keep my role", callback_data="nevermind")])
    await query.edit_message_text(
        f"{name} is cancelling {role} on {pretty}.\n\n"
        f"Available partakers (based on preferences):\n"
        f"Please ask one of them first, then tap their name to confirm.",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return CANCEL_REPLACEMENT


async def cancel_pick_replacement(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "nevermind":
        await query.edit_message_text("OK — nothing changed.")
        return ConversationHandler.END
    context.user_data["cr_new"] = query.data
    name, role = context.user_data["cr_name"], context.user_data["cr_role"]
    date_str = context.user_data["cr_date"]
    buttons = [
        [InlineKeyboardButton("Confirm", callback_data="yes")],
        [InlineKeyboardButton("Cancel", callback_data="no")],
    ]
    await query.edit_message_text(
        f"Confirm: {query.data} takes over {role} on {date_str} from {name}?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return CANCEL_CONFIRM


async def cancel_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data != "yes":
        await query.edit_message_text("OK — nothing changed.")
        return ConversationHandler.END
    ss, service = context.user_data["cr_ss"], context.user_data["cr_service"]
    await do_substitute(
        query, context, ss, service, context.user_data["cr_role"], context.user_data["cr_date"],
        context.user_data["cr_new"], adj_type="cancellation",
    )
    return ConversationHandler.END


# --- /special_request: cap one person's assignments for a month, rebalance the rest ---

SPECIAL_SERVICE, SPECIAL_PERSON, SPECIAL_MONTH, SPECIAL_MAX = range(60, 64)


async def special_request_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["special_ss"] = ss
    # only services with random-mode roles make sense here — a cap has
    # nothing to rebalance against for manual-only services
    tabs = [s for s in get_all_schedule_tabs(ss) if get_random_roles_for_service(ss, s)]
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
    await update.message.reply_text(
        "Special request (limit assignments for a month) — which service?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return SPECIAL_SERVICE


async def special_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["special_service"] = service
    ss = context.user_data["special_ss"]
    names = get_roster_names(ss)
    buttons = [[InlineKeyboardButton(n, callback_data=n)] for n in names]
    await query.edit_message_text("Which person?", reply_markup=InlineKeyboardMarkup(buttons))
    return SPECIAL_PERSON


async def special_select_person(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["special_person"] = query.data
    await query.edit_message_text("Which month?", reply_markup=month_keyboard())
    return SPECIAL_MONTH


async def special_select_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year, month = map(int, query.data.split("-"))
    context.user_data["special_year"] = year
    context.user_data["special_month"] = month
    buttons = [[InlineKeyboardButton(str(n), callback_data=str(n))] for n in (1, 2, 3)]
    await query.edit_message_text(
        "Max total roles for this person this month?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return SPECIAL_MAX


async def special_apply(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    max_assignments = int(query.data)
    ss = context.user_data["special_ss"]
    service = context.user_data["special_service"]
    person = context.user_data["special_person"]
    year = context.user_data["special_year"]
    month = context.user_data["special_month"]

    await query.edit_message_text("Applying...")

    random_roles = get_random_roles_for_service(ss, service)
    ws = ss.worksheet(service)
    entries = find_person_month_entries(ws, year, month, person, random_roles)

    if len(entries) <= max_assignments:
        await query.message.reply_text(
            f"{person} already has {len(entries)} role(s) that month — no change needed."
        )
        return ConversationHandler.END

    to_remove = entries[max_assignments:]  # keep the earliest max_assignments, free up the rest
    counts = load_assignment_counts(ss)
    changes = []

    for row_num, rec in to_remove:
        date_str, role = rec["Date"], rec["Role"]
        eligible = random_roles[role]
        taken_today = find_date_taken(ws, date_str, exclude_row=row_num)
        candidates = [p for p in eligible if p != person and p not in taken_today]
        if not candidates:
            candidates = [p for p in eligible if p != person] or [person]
        min_count = min(counts[p] for p in candidates)
        new_person = random.choice([p for p in candidates if counts[p] == min_count])

        ws.update_cell(row_num, 3, new_person)
        counts[person] = max(0, counts[person] - 1)
        counts[new_person] += 1
        log_adjustment(ss, service, date_str, role, person, new_person, "special_request",
                        reason=f"capped at {max_assignments}/month")
        changes.append((date_str, role, person, new_person))

    save_assignment_counts(ss, counts)

    change_lines = "\n".join(f"{d} {r}: {old} -> {new}" for d, r, old, new in changes)
    month_label = dt.date(year, month, 1).strftime("%B %Y")
    header = f"{person}'s {service} schedule for {month_label} capped at {max_assignments} role(s):\n\n{change_lines}"
    await query.message.reply_text(header)

    updated_rows = rows_in_month(ws.get_all_records(), year, month)
    summary = format_schedule_summary(f"{service} ({month_label}, updated)", records_to_rows(updated_rows))
    await announce_update(context, ss, summary)
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# #6 Roster management
# ---------------------------------------------------------------------------

(ROSTER_MENU, ROSTER_ADD_NAME, ROSTER_SERVICE, ROSTER_ROLE,
 ROSTER_ADDROLE_NAME, ROSTER_REMOVEROLE_NAME, ROSTER_REMOVEALL_NAME) = range(70, 77)


async def roster_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    buttons = [
        [InlineKeyboardButton("Add new member", callback_data="add_member")],
        [InlineKeyboardButton("Add member to a role", callback_data="add_role")],
        [InlineKeyboardButton("Remove member from a role", callback_data="remove_role")],
        [InlineKeyboardButton("Remove member from all roles", callback_data="remove_all")],
    ]
    await update.message.reply_text(
        "Roster management — what would you like to do?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return ROSTER_MENU


async def roster_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = setup_sheet()
    context.user_data["roster_ss"] = ss
    action = query.data
    context.user_data["roster_action"] = action

    if action == "add_member":
        await query.edit_message_text("What's the new member's name?")
        return ROSTER_ADD_NAME

    if action == "remove_all":
        names = get_roster_names(ss)
        buttons = [[InlineKeyboardButton(n, callback_data=n)] for n in names]
        await query.edit_message_text(
            "Remove which member from ALL roles?", reply_markup=InlineKeyboardMarkup(buttons)
        )
        return ROSTER_REMOVEALL_NAME

    # add_role / remove_role both start with picking a service
    tabs = get_all_schedule_tabs(ss)
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
    label = "Add to which service?" if action == "add_role" else "Remove from which service?"
    await query.edit_message_text(label, reply_markup=InlineKeyboardMarkup(buttons))
    return ROSTER_SERVICE


async def roster_add_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    ss = context.user_data["roster_ss"]
    if add_member_to_roster(ss, name):
        await update.message.reply_text(
            f"Added {name} to the roster. Use 'Add member to a role' to assign them a role."
        )
    else:
        await update.message.reply_text(f"{name} is already on the roster.")
    return ConversationHandler.END


async def roster_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["roster_service"] = service
    ss = context.user_data["roster_ss"]
    ws = ss.worksheet(service)
    roles = get_distinct_roles(ws)
    if not roles:
        roles = list(load_service_configs(ss).get(service, {}).get("roles", {}).keys())
    buttons = [[InlineKeyboardButton(r, callback_data=r)] for r in roles]
    await query.edit_message_text(f"{service} — which role?", reply_markup=InlineKeyboardMarkup(buttons))
    return ROSTER_ROLE


async def roster_select_role(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    role = query.data
    context.user_data["roster_role"] = role
    ss = context.user_data["roster_ss"]
    service = context.user_data["roster_service"]

    if context.user_data["roster_action"] == "add_role":
        current = set(get_role_pool(ss, service, role))
        names = [n for n in get_roster_names(ss) if n not in current]
        if not names:
            await query.edit_message_text("Everyone on the roster is already eligible for this role.")
            return ConversationHandler.END
        buttons = [[InlineKeyboardButton(n, callback_data=n)] for n in names]
        await query.edit_message_text("Add whom?", reply_markup=InlineKeyboardMarkup(buttons))
        return ROSTER_ADDROLE_NAME
    else:
        eligible = get_role_pool(ss, service, role)
        if not eligible:
            await query.edit_message_text("Nobody is currently eligible for this role.")
            return ConversationHandler.END
        buttons = [[InlineKeyboardButton(n, callback_data=n)] for n in eligible]
        await query.edit_message_text("Remove whom?", reply_markup=InlineKeyboardMarkup(buttons))
        return ROSTER_REMOVEROLE_NAME


async def roster_addrole_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    name = query.data
    ss = context.user_data["roster_ss"]
    service = context.user_data["roster_service"]
    role = context.user_data["roster_role"]
    add_member_to_role(ss, service, role, name)
    await query.edit_message_text(f"Added {name} to {service} - {role}.")
    return ConversationHandler.END


async def roster_removerole_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    name = query.data
    ss = context.user_data["roster_ss"]
    service = context.user_data["roster_service"]
    role = context.user_data["roster_role"]
    remove_member_from_role(ss, service, role, name)
    await query.edit_message_text(f"Removed {name} from {service} - {role}.")
    return ConversationHandler.END


async def roster_removeall_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    name = query.data
    ss = context.user_data["roster_ss"]
    removed_from = remove_member_from_all_roles(ss, name)
    if removed_from:
        await query.edit_message_text(f"Removed {name} from:\n" + "\n".join(removed_from))
    else:
        await query.edit_message_text(f"{name} wasn't assigned to any roles.")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# #7 Bulk upload: downloadable CSV template, filled offline, uploaded back
# ---------------------------------------------------------------------------

TEMPLATE_SELECT_SERVICE, TEMPLATE_SELECT_PERIOD = range(80, 82)


def build_template_rows(ss, service_type, dates):
    """One row per (date, role) for every role configured on this service —
    pre-filled with whatever's already assigned, blank otherwise, so the
    person only needs to type into the empty cells (e.g. Tech's own
    irregular schedule) rather than re-enter everything."""
    ws = ss.worksheet(service_type)
    existing = {(r["Date"], r["Role"]): r["Partaker"] for r in ws.get_all_records()}
    configs = load_service_configs(ss)
    roles = list(configs.get(service_type, {}).get("roles", {}).keys()) or get_distinct_roles(ws)
    rows = []
    for d in dates:
        for role in roles:
            rows.append([service_type, d.isoformat(), role, existing.get((d.isoformat(), role), "")])
    return rows


def write_template_csv(path, rows):
    import csv
    with open(path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Service", "Date", "Role", "Partaker"])
        writer.writerows(rows)


async def template_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["template_ss"] = ss
    tabs = get_all_schedule_tabs(ss)
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
    await update.message.reply_text(
        "Download a fill-in template for which service?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return TEMPLATE_SELECT_SERVICE


async def template_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    service = query.data
    context.user_data["template_service"] = service
    ss = context.user_data["template_ss"]
    configs = load_service_configs(ss)
    cadence = configs.get(service, {}).get("cadence", "month")

    if cadence == "week":
        mondays = upcoming_mondays()
        context.user_data["template_mondays"] = mondays
        buttons = []
        for i, monday in enumerate(mondays):
            saturday = monday + dt.timedelta(days=6)
            buttons.append([InlineKeyboardButton(
                f"{monday.strftime('%b %d')} - {saturday.strftime('%b %d')}", callback_data=str(i)
            )])
        await query.edit_message_text("Which week?", reply_markup=InlineKeyboardMarkup(buttons))
    else:
        await query.edit_message_text("Which month?", reply_markup=month_keyboard())
    return TEMPLATE_SELECT_PERIOD


async def template_select_period(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss = context.user_data["template_ss"]
    service = context.user_data["template_service"]
    configs = load_service_configs(ss)
    cadence = configs.get(service, {}).get("cadence", "month")

    if cadence == "week":
        monday = context.user_data["template_mondays"][int(query.data)]
        dates = week_dates(service, monday)
    else:
        year, month = map(int, query.data.split("-"))
        dates = get_service_month_dates(ss, service, year, month)

    rows = build_template_rows(ss, service, dates)
    path = f"/tmp/{service}_template.csv"
    write_template_csv(path, rows)

    await query.edit_message_text("Here's your template — fill in the Partaker column and send it back to me.")
    with open(path, "rb") as f:
        await context.bot.send_document(chat_id=query.message.chat_id, document=f, filename=f"{service}_template.csv")
    return ConversationHandler.END


async def handle_schedule_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Global (non-conversation) handler: any .csv document sent to the bot
    is treated as a filled-in template from /template and upserted directly
    — no admin needs to remember which command started it, since the file
    carries its own Service/Date/Role columns."""
    doc = update.message.document
    if not doc.file_name or not doc.file_name.lower().endswith(".csv"):
        return

    import csv
    tg_file = await doc.get_file()
    path = f"/tmp/upload_{doc.file_unique_id}.csv"
    await tg_file.download_to_drive(path)

    clear_read_cache()
    ss = setup_sheet()
    by_service = defaultdict(list)
    # utf-8-sig strips the BOM that Excel adds, which would otherwise turn the
    # first header into "\ufeffService" and break the lookup below.
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            partaker = (row.get("Partaker") or "").strip()
            if not partaker:
                continue
            by_service[(row.get("Service") or "").strip()].append(
                ((row.get("Date") or "").strip(), (row.get("Role") or "").strip(), partaker)
            )

    if not by_service:
        await update.message.reply_text("No filled-in rows found in that file — nothing to update.")
        return

    known_tabs = {ws.title for ws in ss.worksheets()}
    total = 0
    skipped = []
    warnings = []
    for service, entries in by_service.items():
        if service not in known_tabs:
            skipped.append(service or "(blank)")
            continue
        ws = ss.worksheet(service)
        records = ws.get_all_records()
        row_index = {(r["Date"], r["Role"]): (i + 2, r["Partaker"]) for i, r in enumerate(records)}
        new_rows, changed = [], []
        for date_str, role, partaker in entries:
            key = (date_str, role)
            if key in row_index:
                row_num, current = row_index[key]
                if current != partaker:
                    ws.update_cell(row_num, 3, partaker)
                    changed.append((date_str, role, partaker))
            else:
                new_rows.append([date_str, role, partaker, "scheduled"])
                changed.append((date_str, role, partaker))
            total += 1
        if new_rows:
            ws.append_rows(new_rows)
        warnings.extend(f"{service} {w}" for w in double_booking_warnings(ss, service, changed))

    msg = f"Bulk upload processed: {total} entr{'y' if total == 1 else 'ies'} updated."
    if skipped:
        msg += f"\nSkipped unknown service(s): {', '.join(skipped)}"
    if warnings:
        msg += ("\n\n⚠️ This upload left someone with two roles on the same day:\n"
                + "\n".join(f"- {w}" for w in warnings)
                + "\nUse /swap or /substitute to fix it.")
    await update.message.reply_text(msg)


# ---------------------------------------------------------------------------
# Dashboard: one tab per service, laid out as monthly tables (Date rows x
# Role columns), admin-editable directly in Sheets.
# ---------------------------------------------------------------------------
#
# Two commands, two directions:
#   /refresh_dashboard  — rebuilds every "{service} Dashboard" tab from the
#                          service's own data tab (overwrites the dashboard;
#                          sync first if there are unsynced edits)
#   /sync_dashboard     — reads whatever admins typed into each dashboard's
#                          grid and upserts it back into the service's real
#                          data tab (blank cells are left alone, never erase)
#
# For real-time syncing instead of a manual command, see the companion
# Apps Script (dashboard_sync.gs) — an onEdit trigger pasted into the
# spreadsheet's Script Editor that calls sync automatically on every edit.

def dashboard_tab_name(service):
    return f"{service} Dashboard"


def dates_matching_weekdays_in_month(year, month, weekdays):
    _, last_day = calendar.monthrange(year, month)
    return [dt.date(year, month, d) for d in range(1, last_day + 1)
            if dt.date(year, month, d).weekday() in weekdays]


def get_service_month_dates(ss, service, year, month, existing_dates=()):
    """Which dates in this month belong to this service. Known cadences
    (Sunday/Wednesday/SunStopSundays/Predawn) use their real weekday
    pattern. A custom service with month cadence defaults to Sunday (same
    assumption used elsewhere for custom services). A custom service with
    week cadence has no fixed day-of-week on file, so falls back to
    whatever dates already have data that month."""
    if service == "Predawn":
        return dates_matching_weekdays_in_month(year, month, PREDAWN_WEEKDAYS)
    if service in SERVICE_WEEKDAY:
        return dates_matching_weekdays_in_month(year, month, [SERVICE_WEEKDAY[service]])
    configs = load_service_configs(ss)
    cadence = configs.get(service, {}).get("cadence", "month")
    if cadence == "month":
        return dates_matching_weekdays_in_month(year, month, [6])
    return sorted(existing_dates)  # week-cadence custom service: data-driven only


def month_range(start_year, start_month, count):
    out = []
    y, m = start_year, start_month
    for _ in range(count):
        out.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


def build_service_dashboard_grid(ss, service, months_ahead=DASHBOARD_MONTHS_AHEAD):
    """Returns (rows, bold_row_indices, roles) — rows is the full 2D grid to
    write (month-label rows, column-header rows, data rows, blank
    separators between months), bold_row_indices are the 0-based rows that
    should be bolded (month labels + column headers)."""
    ws = ss.worksheet(service)
    records = ws.get_all_records()
    existing = {(r["Date"], r["Role"]): r["Partaker"] for r in records}

    configs = load_service_configs(ss)
    roles = list(configs.get(service, {}).get("roles", {}).keys()) or get_distinct_roles(ws)

    today = today_local()
    months = month_range(today.year, today.month, months_ahead)
    # also include any month that already has data further out, so
    # already-generated future schedules never get hidden by the window
    all_data_months = {(dt.date.fromisoformat(d).year, dt.date.fromisoformat(d).month)
                        for (d, _r) in existing if d}
    for ym in sorted(all_data_months):
        if ym not in months:
            months.append(ym)
    months.sort()

    rows = []
    bold_rows = []
    for year, month in months:
        prefix = f"{year:04d}-{month:02d}"
        existing_dates_this_month = {dt.date.fromisoformat(d) for (d, _r) in existing if d.startswith(prefix)}
        dates = get_service_month_dates(ss, service, year, month, existing_dates_this_month)

        bold_rows.append(len(rows))
        rows.append([dt.date(year, month, 1).strftime("%B %Y")] + [""] * len(roles))
        bold_rows.append(len(rows))
        rows.append(["Date"] + roles)
        for d in dates:
            row = [d.isoformat()] + [existing.get((d.isoformat(), role), "") for role in roles]
            rows.append(row)
        rows.append([""] * (len(roles) + 1))  # blank separator between months

    return rows, bold_rows, roles


def refresh_service_dashboard(ss, service, months_ahead=DASHBOARD_MONTHS_AHEAD):
    rows, bold_rows, roles = build_service_dashboard_grid(ss, service, months_ahead)
    tab = dashboard_tab_name(service)
    existing_tabs = {ws.title for ws in ss.worksheets()}
    if tab not in existing_tabs:
        dash_ws = ss.add_worksheet(title=tab, rows=max(200, len(rows) + 10), cols=max(5, len(roles) + 1))
    else:
        dash_ws = ss.worksheet(tab)
        dash_ws.clear()

    if rows:
        dash_ws.append_rows(rows)

    try:
        dash_ws.freeze(rows=0)  # per-month headers, not one fixed top row
        for r in bold_rows:
            dash_ws.format(f"A{r + 1}:{chr(ord('A') + len(roles))}{r + 1}", {"textFormat": {"bold": True}})
    except Exception:
        pass  # cosmetic only — don't fail the refresh over formatting

    return len(rows)


def refresh_all_dashboards(ss, months_ahead=DASHBOARD_MONTHS_AHEAD):
    """Returns {service: row_count} for every schedule tab that actually exists."""
    counts = {}
    known_tabs = {ws.title for ws in ss.worksheets()}
    for service in get_all_schedule_tabs(ss):
        if service not in known_tabs:
            continue
        counts[service] = refresh_service_dashboard(ss, service, months_ahead)
    return counts


def is_month_header_label(text):
    try:
        dt.datetime.strptime(text.strip(), "%B %Y")
        return True
    except ValueError:
        return False


def parse_dashboard_grid(ws):
    """Reads a '{service} Dashboard' tab's current grid (month header rows,
    column header rows, data rows, blank separators) and returns
    [(date_str, role, partaker), ...] for every non-blank Partaker cell —
    parsed fresh from the sheet's own layout, so it works even after a bot
    restart between /refresh_dashboard and /sync_dashboard."""
    values = ws.get_all_values()
    entries = []
    roles = None
    i = 0
    while i < len(values):
        row = values[i]
        if not any(cell.strip() for cell in row):
            roles = None
            i += 1
            continue
        first_cell = row[0].strip()
        if is_month_header_label(first_cell):
            i += 1
            if i >= len(values):
                break
            roles = [h.strip() for h in values[i][1:] if h.strip()]
            i += 1
            continue
        if roles is not None and first_cell != "Date":
            try:
                dt.datetime.strptime(first_cell, "%Y-%m-%d")
            except ValueError:
                i += 1
                continue
            for col_idx, role in enumerate(roles, start=1):
                partaker = row[col_idx].strip() if col_idx < len(row) else ""
                if partaker:
                    entries.append((first_cell, role, partaker))
        i += 1
    return entries


def sync_service_dashboard(ss, service, warnings=None):
    """Saves what admins typed into a service's Dashboard tab. Only cells that
    actually differ are written. If `warnings` (a list) is given, lines about
    anyone these edits left with two roles on the same day are added to it."""
    tab = dashboard_tab_name(service)
    existing_tabs = {ws.title for ws in ss.worksheets()}
    if tab not in existing_tabs:
        return 0
    dash_ws = ss.worksheet(tab)
    entries = parse_dashboard_grid(dash_ws)
    if not entries:
        return 0

    ws = ss.worksheet(service)
    records = ws.get_all_records()
    row_index = {(r["Date"], r["Role"]): (i + 2, r["Partaker"]) for i, r in enumerate(records)}
    new_rows, changed = [], []
    total = 0
    for date_str, role, partaker in entries:
        key = (date_str, role)
        if key in row_index:
            row_num, current = row_index[key]
            if current != partaker:
                ws.update_cell(row_num, 3, partaker)
                changed.append((date_str, role, partaker))
        else:
            new_rows.append([date_str, role, partaker, "scheduled"])
            changed.append((date_str, role, partaker))
        total += 1
    if new_rows:
        ws.append_rows(new_rows)
    if warnings is not None:
        warnings.extend(f"{service} {w}" for w in double_booking_warnings(ss, service, changed))
    return total


def sync_all_dashboards(ss, warnings=None):
    """Returns {service: row_count} for every dashboard tab that exists."""
    counts = {}
    for service in get_all_schedule_tabs(ss):
        n = sync_service_dashboard(ss, service, warnings)
        if n:
            counts[service] = n
    return counts


async def refresh_dashboard_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    counts = refresh_all_dashboards(ss)
    lines = "\n".join(f"- {s}: {n} row(s)" for s, n in counts.items())
    await update.message.reply_text(
        f"Dashboards refreshed (next {DASHBOARD_MONTHS_AHEAD} months):\n{lines}\n\n"
        f"Edit any Partaker cell directly, then run /sync_dashboard when you're done."
    )


async def sync_dashboard_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_read_cache()  # admins just typed into the sheet by hand — never sync from a cached copy
    ss = setup_sheet()
    warnings = []
    counts = sync_all_dashboards(ss, warnings)
    if not counts:
        await update.message.reply_text("No dashboard edits found to sync.")
        return
    lines = "\n".join(f"- {s}: {n} entr{'y' if n == 1 else 'ies'}" for s, n in counts.items())
    msg = f"Synced back to service tabs:\n{lines}"
    if warnings:
        msg += ("\n\n⚠️ These edits left someone with two roles on the same day:\n"
                + "\n".join(f"- {w}" for w in warnings)
                + "\nUse /swap or /substitute to fix it.")
    await update.message.reply_text(msg)


# ---------------------------------------------------------------------------
# Yearly renewal: archive the outgoing year, empty the data tabs, and
# refresh dashboards so they roll into the new year's months.
# ---------------------------------------------------------------------------

def archive_year(ss, year):
    """Moves every row dated `year` out of each service's data tab into a
    single Archive_{year} tab (prefixed with which service it came from),
    then rewrites the original tab with only the rows that DON'T belong to
    that year (normally none, if renewal happens once a year as intended).
    Returns {service: rows_archived}."""
    archive_tab = f"Archive_{year}"
    existing_tabs = {ws.title for ws in ss.worksheets()}
    if archive_tab in existing_tabs:
        archive_ws = ss.worksheet(archive_tab)
    else:
        archive_ws = ss.add_worksheet(title=archive_tab, rows=2000, cols=5)
        archive_ws.append_row(["Service", "Date", "Role", "Partaker", "Status"])

    prefix = f"{year:04d}"
    counts = {}
    archived_rows = []
    known_tabs = {ws.title for ws in ss.worksheets()}
    for service in get_all_schedule_tabs(ss):
        if service not in known_tabs:
            continue  # e.g. a custom service in ServiceConfig whose tab hasn't been created yet
        ws = ss.worksheet(service)
        records = ws.get_all_records()
        moved = [r for r in records if r.get("Date", "").startswith(prefix)]
        if not moved:
            continue
        kept = [r for r in records if not r.get("Date", "").startswith(prefix)]
        archived_rows.extend([[service, r["Date"], r["Role"], r["Partaker"], r.get("Status", "scheduled")]
                               for r in moved])
        counts[service] = len(moved)

        ws.clear()
        ws.append_row(TAB_HEADERS.get(service, ["Date", "Role", "Partaker", "Status"]))
        if kept:
            ws.append_rows([[r["Date"], r["Role"], r["Partaker"], r.get("Status", "scheduled")] for r in kept])

    if archived_rows:
        archive_ws.append_rows(archived_rows)

    return counts


def reset_roster_counts(ss):
    """Zeroes every AssignmentCount in the Roster tab — used on yearly
    renewal so the new year's equal-share generation starts fresh rather
    than inheriting the outgoing year's balance."""
    roster_ws = ss.worksheet("Roster")
    n = len(roster_ws.get_all_records())
    if n:
        roster_ws.update(range_name=f"B2:B{n + 1}", values=[[0]] * n)
    return n


RENEW_CONFIRM = 110


async def renew_year_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    year = today_local().year
    buttons = [
        [InlineKeyboardButton(f"Yes, archive {year}", callback_data=str(year))],
        [InlineKeyboardButton("Cancel", callback_data="cancel")],
    ]
    await update.message.reply_text(
        f"This will move every {year} row out of each service tab into 'Archive_{year}', "
        f"empty the live tabs, reset everyone's equal-share count to 0, and refresh dashboards "
        f"for the months ahead. Proceed?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return RENEW_CONFIRM


async def renew_year_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "cancel":
        await query.edit_message_text("Cancelled — nothing archived.")
        return ConversationHandler.END

    year = int(query.data)
    await query.edit_message_text(f"Archiving {year}...")
    ss = setup_sheet()
    archived_counts = archive_year(ss, year)
    reset_count = reset_roster_counts(ss)
    dashboard_counts = refresh_all_dashboards(ss)

    archive_lines = "\n".join(f"- {s}: {n} row(s)" for s, n in archived_counts.items()) or "(nothing to archive)"
    await query.message.reply_text(
        f"Archived to 'Archive_{year}':\n{archive_lines}\n\n"
        f"Reset equal-share counts for {reset_count} roster member(s).\n"
        f"Dashboards refreshed for the months ahead — ready for {year + 1}."
    )
    return ConversationHandler.END




# ---------------------------------------------------------------------------
# Preference collection: admin opens a round for a service+month with a
# deadline; partakers claim (date, role) slots first-come-first-served via
# a private-chat flow reached through a deep link posted in the group.
# ---------------------------------------------------------------------------
#
# Telegram bots cannot message a user who hasn't opened a private chat with
# them first, so the group announcement uses a deep link
# (t.me/<bot>?start=pref_<RoundID>) rather than trying to DM directly. That
# link opens a private chat and routes straight into that round's flow.
#
# Reservations are written straight into the service's own data tab (same
# shape as everything else), so pull_schedule/dashboards/etc. show them
# immediately. The generator (see get_already_filled + the `already_filled`
# parameter on generate_schedule/generate_predawn_schedule/
# generate_simple_schedule above) skips any slot already filled this way.

DEADLINE_PRESETS = [("In 3 days", 3), ("In 1 week", 7), ("In 2 weeks", 14)]


def preacher_dates_for_gating(service, year, month):
    """Dates that need a Preacher set before this service+month can move
    forward, for services that have a Preacher role at all."""
    if service == "Predawn":
        return dates_matching_weekdays_in_month(year, month, PREDAWN_WEEKDAYS)
    if service in ("Sunday", "Wednesday"):
        return dates_matching_weekdays_in_month(year, month, [SERVICE_WEEKDAY[service]])
    return []  # FilipinoTranslation, SunStopSundays, custom services: no Preacher role


def is_preacher_fully_set(ss, service, year, month):
    dates = preacher_dates_for_gating(service, year, month)
    if not dates:
        return True
    preacher_assignments = get_preacher_assignments(ss, service, dates)
    return not missing_preacher_dates(service, dates, preacher_assignments)


def find_round_for(ss, service, year, month):
    """Most recently created round for this service+month, or None."""
    candidates = [
        r for r in ss.worksheet("PreferenceRounds").get_all_records()
        if r.get("Service") == service and str(r.get("Year")) == str(year) and str(r.get("Month")) == str(month)
    ]
    return candidates[-1] if candidates else None


def create_preference_round(ss, service, year, month, deadline_dt, chat_id):
    # Telegram deep-link payloads only allow letters, digits, "_" and "-", so
    # strip everything else from the service name (custom services may have
    # spaces or punctuation).
    safe_service = re.sub(r"[^A-Za-z0-9]", "", service) or "svc"
    round_id = f"{safe_service}_{year}{month:02d}_{dt.datetime.now(CHURCH_TZ).strftime('%H%M%S')}"
    ws = ss.worksheet("PreferenceRounds")
    ws.append_rows([[round_id, service, year, month, deadline_dt.isoformat(), "open", chat_id]])
    return round_id


def get_round(ss, round_id):
    ws = ss.worksheet("PreferenceRounds")
    for r in ws.get_all_records():
        if r.get("RoundID") == round_id:
            return r
    return None


def close_round(ss, round_id):
    ws = ss.worksheet("PreferenceRounds")
    for i, r in enumerate(ws.get_all_records()):
        if r.get("RoundID") == round_id:
            ws.update_cell(i + 2, 6, "closed")  # col 6 = Status
            return True
    return False


def is_round_open(round_):
    if not round_ or round_.get("Status") != "open":
        return False
    try:
        deadline = dt.datetime.fromisoformat(str(round_["Deadline"]))
    except ValueError:
        return False
    return dt.datetime.now(CHURCH_TZ) < deadline


def get_round_dates(ss, round_):
    service, year, month = round_["Service"], int(round_["Year"]), int(round_["Month"])
    ws = ss.worksheet(service)
    prefix = f"{year:04d}-{month:02d}"
    existing_dates_this_month = {
        dt.date.fromisoformat(r["Date"]) for r in ws.get_all_records() if r.get("Date", "").startswith(prefix)
    }
    return get_service_month_dates(ss, service, year, month, existing_dates_this_month)


def get_round_roles(ss, round_):
    configs = load_service_configs(ss)
    return list(configs.get(round_["Service"], {}).get("roles", {}).keys())


def get_open_roles_for_date(ss, round_, date_str, name):
    """Roles still unclaimed for this date that `name` is eligible for.
    Double-role conflicts are no longer hidden — the caller shows a
    confirmation instead of silently blocking (see find_conflicts)."""
    ws = ss.worksheet(round_["Service"])
    records = ws.get_all_records()
    taken_roles = {r["Role"] for r in records if r["Date"] == date_str}
    return [
        role for role in get_round_roles(ss, round_)
        if role not in taken_roles and name in get_role_pool(ss, round_["Service"], role)
    ]


def reserve_slot(ss, round_, date_str, role, name):
    """Re-checks the slot is still open right before writing, to shrink the
    race window. Returns True if reserved, False if someone beat them to it."""
    ws = ss.worksheet(round_["Service"])
    if any(r["Date"] == date_str and r["Role"] == role for r in ws.get_all_records()):
        return False
    ws.append_rows([[date_str, role, name, "scheduled"]])
    counts = load_assignment_counts(ss)
    counts[name] += 1
    save_assignment_counts(ss, counts)
    add_member_to_roster(ss, name)  # harmless no-op if already listed
    return True


async def do_close_round(context, round_id):
    ss = setup_sheet()
    round_ = get_round(ss, round_id)
    if not round_ or round_.get("Status") != "open":
        return  # already closed — avoid double notification
    close_round(ss, round_id)
    service = round_["Service"]
    year, month = int(round_["Year"]), int(round_["Month"])
    month_label = dt.date(year, month, 1).strftime("%B %Y")

    if service == "FilipinoTranslation":
        # Send the translation schedule-so-far to the Filipino Translators
        # group, then tell the Service Partakers group Sunday preferences
        # can now be opened.
        fil_chat_id = round_.get("GroupChatID") or get_group_chat_id(ss, "Filipino Translators")
        if fil_chat_id:
            ws = ss.worksheet("FilipinoTranslation")
            rows = rows_in_month(ws.get_all_records(), year, month)
            if rows:  # nothing to show until the month has actually been generated
                summary = format_schedule_summary(f"Filipino Translation — {month_label}", records_to_rows(rows))
                await context.bot.send_message(chat_id=fil_chat_id, text=summary, parse_mode="Markdown")

        partaker_chat_id = get_group_chat_id(ss, "Service Partakers")
        if partaker_chat_id:
            text = (
                f"✅ Filipino Translation preferences for *{month_label}* are closed.\n\n"
                f"Admin can now run /open_preferences for the Sunday service partakers."
            )
            await context.bot.send_message(chat_id=partaker_chat_id, text=text, parse_mode="Markdown")
        return

    chat_ids = split_chat_ids(round_.get("GroupChatID")) or collect_round_group_chat_ids(ss, service)
    if chat_ids:
        text = (
            f"⏰ Preference collection for {service} — {month_label} has closed.\n\n"
            f"The schedule can now be generated — nobody will be assigned on a date they marked as "
            f"not available.\n\n"
            f"For changes after this point, use /cancel_role, /swap, /substitute or /special_request "
            f"(please ask someone first, and submit changes in advance)."
        )
        for chat_id in chat_ids:
            await context.bot.send_message(chat_id=chat_id, text=text)


async def close_preference_round_job(context: ContextTypes.DEFAULT_TYPE):
    await do_close_round(context, context.job.data["round_id"])


async def startup_recover_rounds(context: ContextTypes.DEFAULT_TYPE):
    """Runs once shortly after the bot starts. job_queue.run_once jobs live
    in memory and don't survive a restart, so on startup this re-reads
    PreferenceRounds (the source of truth) and either closes any round
    whose deadline already passed while the bot was down, or re-schedules
    the ones still pending."""
    ss = setup_sheet()
    now = dt.datetime.now(CHURCH_TZ)
    for r in ss.worksheet("PreferenceRounds").get_all_records():
        if r.get("Status") != "open":
            continue
        round_id = r["RoundID"]
        try:
            deadline = dt.datetime.fromisoformat(str(r["Deadline"]))
        except ValueError:
            continue
        if deadline <= now:
            await do_close_round(context, round_id)
        else:
            context.job_queue.run_once(
                close_preference_round_job, when=deadline, data={"round_id": round_id}, name=f"close_{round_id}"
            )


# --- /open_preferences: admin opens a round ---

OPEN_PREF_SERVICE, OPEN_PREF_MONTH, OPEN_PREF_DEADLINE, OPEN_PREF_CUSTOM_TEXT = range(115, 119)


async def open_preferences_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["pref_ss"] = ss
    # Every service can collect unavailable dates except Sun Stop Sundays,
    # which is generated directly and adjusted afterwards.
    tabs = [s for s in get_all_schedule_tabs(ss) if s != "SunStopSundays"]
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
    await update.message.reply_text(
        "Open a preference round for which service?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return OPEN_PREF_SERVICE


async def open_pref_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["pref_service"] = query.data
    await query.edit_message_text(f"{query.data} — which month?", reply_markup=month_keyboard())
    return OPEN_PREF_MONTH


async def open_pref_select_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year, month = map(int, query.data.split("-"))
    service = context.user_data["pref_service"]
    ss = context.user_data["pref_ss"]
    month_label = dt.date(year, month, 1).strftime("%B %Y")

    # Gate 1: Preacher must already be set for this service+month (Sunday's
    # Preacher gates FilipinoTranslation too, since translation serves Sunday).
    # Predawn skips this gate: its Preacher comes from the weekly pattern and
    # it has no randomized roles that depend on preferences.
    preacher_service = "Sunday" if service == "FilipinoTranslation" else service
    if service != "Predawn" and not is_preacher_fully_set(ss, preacher_service, year, month):
        await query.edit_message_text(
            f"Can't open preferences yet — {preacher_service} has no Preacher scheduled for "
            f"{month_label}. Run /set_preacher first."
        )
        return ConversationHandler.END

    # Gate 2: Sunday partaker preferences require Filipino Translation's
    # round for the same month to already be closed.
    if service == "Sunday":
        fil_round = find_round_for(ss, "FilipinoTranslation", year, month)
        if not fil_round or fil_round.get("Status") != "closed":
            await query.edit_message_text(
                f"Can't open Sunday partaker preferences yet — open and close a Filipino Translation "
                f"preference round for {month_label} first (/open_preferences)."
            )
            return ConversationHandler.END

    context.user_data["pref_year"] = year
    context.user_data["pref_month"] = month
    buttons = [[InlineKeyboardButton(label, callback_data=str(days))] for label, days in DEADLINE_PRESETS]
    buttons.append([InlineKeyboardButton("Custom date/time", callback_data="custom")])
    await query.edit_message_text("When should the preference window close?", reply_markup=InlineKeyboardMarkup(buttons))
    return OPEN_PREF_DEADLINE


async def open_pref_finish(update: Update, context: ContextTypes.DEFAULT_TYPE, deadline, is_callback):
    ss = context.user_data["pref_ss"]
    service = context.user_data["pref_service"]
    year, month = context.user_data["pref_year"], context.user_data["pref_month"]
    reply = update.callback_query.edit_message_text if is_callback else update.message.reply_text

    # Filipino Translation is announced only to its own Filipino Translators
    # group, matching how its finished schedule and Friday reminder already
    # work. Every other service notifies the Service Partakers group PLUS any
    # group registered for one of that service's roles (e.g. a Presiders group).
    chat_ids = collect_round_group_chat_ids(ss, service)
    if not chat_ids:
        label = "Filipino Translators" if service == "FilipinoTranslation" else \
            "Service Partakers (or a role group, e.g. Presider)"
        await reply(f"No {label} group is registered yet — run /set_group_chat in that group first.")
        return ConversationHandler.END

    round_id = create_preference_round(ss, service, year, month, deadline, ",".join(str(c) for c in chat_ids))
    month_label = dt.date(year, month, 1).strftime("%B %Y")
    deep_link = f"https://t.me/{BOT_USERNAME}?start=pref_{round_id}"
    # HTML (not Markdown) so the underscores in /cancel_role etc. are safe
    announcement = (
        f"📋 <b>Preference collection is open for {html.escape(service)} — {month_label}!</b>\n\n"
        f"{html.escape(PARTAKER_PREFERENCE_NOTE)}\n{html.escape(PARTAKER_COMMANDS_HINT)}\n\n"
        f"Deadline: <b>{deadline.strftime('%B %d, %Y %I:%M %p')}</b>. If you don't respond by then, "
        f"you're treated as available on every date.\n\n"
        f'<a href="{html.escape(deep_link, quote=True)}">Set my availability</a>'
    )
    for chat_id in chat_ids:
        await context.bot.send_message(
            chat_id=chat_id, text=announcement, parse_mode="HTML", disable_web_page_preview=True
        )
    context.job_queue.run_once(
        close_preference_round_job, when=deadline, data={"round_id": round_id}, name=f"close_{round_id}"
    )

    await reply(f"Preference round opened for {service} — {month_label}, closing "
                f"{deadline.strftime('%b %d, %Y %I:%M %p')}.")
    return ConversationHandler.END


async def open_pref_select_deadline(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "custom":
        await query.edit_message_text("Send the deadline as YYYY-MM-DD HH:MM (24h, church local time).")
        return OPEN_PREF_CUSTOM_TEXT
    days = int(query.data)
    deadline = (dt.datetime.now(CHURCH_TZ) + dt.timedelta(days=days)).replace(hour=20, minute=0, second=0, microsecond=0)
    return await open_pref_finish(update, context, deadline, is_callback=True)


async def open_pref_custom_deadline(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        naive = dt.datetime.strptime(text, "%Y-%m-%d %H:%M")
    except ValueError:
        await update.message.reply_text("Couldn't parse that — send it as YYYY-MM-DD HH:MM, e.g. 2026-10-15 20:00.")
        return OPEN_PREF_CUSTOM_TEXT
    deadline = naive.replace(tzinfo=CHURCH_TZ)
    return await open_pref_finish(update, context, deadline, is_callback=False)


# --- /close_preferences: admin ends a round early ("everyone has responded") ---

CLOSE_PREF_SELECT = 119


async def close_preferences_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    open_rounds = [r for r in ss.worksheet("PreferenceRounds").get_all_records() if r.get("Status") == "open"]
    if not open_rounds:
        await update.message.reply_text("No open preference rounds right now.")
        return ConversationHandler.END
    context.user_data["close_pref_ss"] = ss
    buttons = []
    for r in open_rounds:
        month_label = dt.date(int(r["Year"]), int(r["Month"]), 1).strftime("%b %Y")
        buttons.append([InlineKeyboardButton(f"{r['Service']} — {month_label}", callback_data=r["RoundID"])])
    await update.message.reply_text("Close which round early?", reply_markup=InlineKeyboardMarkup(buttons))
    return CLOSE_PREF_SELECT


async def close_preferences_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("Closing...")
    await do_close_round(context, query.data)
    await query.message.reply_text("Round closed.")
    return ConversationHandler.END


# --- /start with a deep-link payload: partaker marks the dates they can NOT do ---
# Preferences are "unavailable dates", not slot claims: everyone starts out
# available, taps the dates they can't make, and generation (plus
# /cancel_role's replacement list) then avoids them on those dates.

PREF_SELECT_NAME, PREF_SELECT_DATE, PREF_SELECT_ROLE, PREF_CONFIRM_CONFLICT = range(120, 124)
# (PREF_SELECT_ROLE / PREF_CONFIRM_CONFLICT are no longer used; the numbers
# stay reserved so nothing else gets renumbered.)

PARTAKER_PREFERENCE_NOTE = (
    "Please select the dates you are NOT available to be a partaker. "
    "No need for preachers to set their preferences, as they are automatically waived for other partaker roles. "
    "For schedule swaps and substitutions, you can submit them later via 'swap', 'substitutions' "
    "and 'special request'. Please make sure to ask someone ahead of time for swaps and substitutions. "
    "If you cannot find a substitution, inform the Service Director right away. "
    "Please submit any changes in advance."
)
PARTAKER_COMMANDS_HINT = "(Commands: /cancel_role, /swap, /substitute, /special_request)"


def get_unavailability(ss, service, dates):
    """{date_str: set(names)} of partakers who marked themselves unavailable
    for any of these dates (date objects or ISO strings) in this service."""
    date_strs = {d if isinstance(d, str) else d.isoformat() for d in dates}
    out = defaultdict(set)
    for r in ss.worksheet("Unavailability").get_all_records():
        if r.get("Service") == service and r.get("Date") in date_strs:
            out[r["Date"]].add(r["Partaker"])
    return out


def toggle_unavailable(ss, service, date_str, name):
    """Flips one (service, date, name) row. Returns True if the person is now
    marked unavailable, False if the mark was just removed."""
    ws = ss.worksheet("Unavailability")
    for i, row in enumerate(ws.get_all_values()[1:], start=2):
        if row[:3] == [service, date_str, name]:
            ws.delete_rows(i)
            return False
    ws.append_rows([[service, date_str, name]])
    return True


def get_round_partaker_names(ss, round_):
    """Names worth showing in a round: anyone eligible for at least one
    non-Preacher role of that service (Preachers are waived). Falls back to
    the whole roster if the service has no eligible names on file."""
    service = round_["Service"]
    names = set()
    for role in get_round_roles(ss, round_):
        if role != "Preacher":
            names.update(get_role_pool(ss, service, role))
    names.discard(LIVE_BROADCAST)
    return sorted(names) or get_roster_names(ss)


async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args or not args[0].startswith("pref_"):
        await update.message.reply_text(
            "Hi! I'm the service scheduling bot. Choose an option below. "
            "If someone shared a preference link with you, tap that link to set your preferences.",
            reply_markup=InlineKeyboardMarkup(MAIN_MENU_ROWS),
        )
        return ConversationHandler.END

    round_id = args[0][len("pref_"):]
    ss = setup_sheet()
    round_ = get_round(ss, round_id)
    if not is_round_open(round_):
        await update.message.reply_text("Sorry, this preference round has closed or doesn't exist anymore.")
        return ConversationHandler.END

    context.user_data["pref_round"] = round_
    context.user_data["pref_ss"] = ss
    names = get_round_partaker_names(ss, round_)
    buttons = [[InlineKeyboardButton(n, callback_data=n)] for n in names]
    month_label = dt.date(int(round_["Year"]), int(round_["Month"]), 1).strftime("%B %Y")
    await update.message.reply_text(
        f"Setting preferences for {round_['Service']} — {month_label}.\n\n"
        f"{PARTAKER_PREFERENCE_NOTE}\n{PARTAKER_COMMANDS_HINT}\n\n"
        f"Which name is yours?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return PREF_SELECT_NAME


def pref_round_still_open(context):
    ss, round_ = context.user_data["pref_ss"], context.user_data["pref_round"]
    return is_round_open(get_round(ss, round_["RoundID"]))


async def pref_select_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ss, round_ = context.user_data["pref_ss"], context.user_data["pref_round"]
    name = query.data
    service = round_["Service"]

    if not pref_round_still_open(context):
        await query.edit_message_text("This preference round has closed — thanks for checking!")
        return ConversationHandler.END

    dates = get_round_dates(ss, round_)
    context.user_data["pref_name"] = name
    context.user_data["pref_dates"] = dates
    context.user_data["pref_marked"] = {
        ds for ds, names in get_unavailability(ss, service, dates).items() if name in names
    }
    context.user_data["pref_preacher_dates"] = {
        r["Date"] for r in ss.worksheet(service).get_all_records()
        if r.get("Role") == "Preacher" and r.get("Partaker") == name
    }
    return await pref_show_dates(update, context)


async def pref_show_dates(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = context.user_data["pref_name"]
    dates = context.user_data["pref_dates"]
    marked = context.user_data["pref_marked"]
    preacher_dates = context.user_data["pref_preacher_dates"]

    buttons, row = [], []
    for d in dates:
        date_str = d.isoformat()
        label = d.strftime("%a %b %d")
        if date_str in preacher_dates:
            label = f"🎤 {label}"
        elif date_str in marked:
            label = f"✖ {label}"
        row.append(InlineKeyboardButton(label, callback_data=date_str))
        if len(row) == 3:
            buttons.append(row)
            row = []
    if row:
        buttons.append(row)
    # Skip the bulk button once every non-preacher date is already marked — nothing left to add.
    can_mark_more = any(d.isoformat() not in preacher_dates and d.isoformat() not in marked for d in dates)
    if can_mark_more:
        buttons.append([InlineKeyboardButton("✖ I'm not available this month", callback_data="all_unavailable")])
    buttons.append([InlineKeyboardButton("I'm done", callback_data="done")])

    text = (
        f"Hi {name}! Tap each date you are NOT available (✖ = not available; tap again to undo). "
        f"Dates you leave alone count as available. Tap 'I'm done' when finished."
    )
    if can_mark_more:
        text += " Not available at all this month? Tap 'I'm not available this month' instead of every date."
    if preacher_dates:
        text += "\n🎤 = you're the Preacher that day, so you're already waived from other roles."
    await update.callback_query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(buttons))
    return PREF_SELECT_DATE


async def pref_select_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    ss, round_ = context.user_data["pref_ss"], context.user_data["pref_round"]
    name = context.user_data["pref_name"]

    if query.data == "done":
        await query.answer()
        marked = sorted(context.user_data["pref_marked"])
        listing = ", ".join(dt.date.fromisoformat(d).strftime("%b %d") for d in marked) if marked else "none"
        await query.edit_message_text(
            f"Thanks {name} — your unavailable dates are saved: {listing}.\n"
            f"You can reopen this link anytime before the deadline to change them."
        )
        return ConversationHandler.END

    if query.data == "all_unavailable":
        if not pref_round_still_open(context):
            await query.answer()
            await query.edit_message_text("This preference round just closed — sorry!")
            return ConversationHandler.END
        await query.answer("Marking every date as not available...")
        service = round_["Service"]
        marked = context.user_data["pref_marked"]
        preacher_dates = context.user_data["pref_preacher_dates"]
        for d in context.user_data["pref_dates"]:
            date_str = d.isoformat()
            if date_str in preacher_dates or date_str in marked:
                continue
            toggle_unavailable(ss, service, date_str, name)
            marked.add(date_str)
        return await pref_show_dates(update, context)

    date_str = query.data
    if date_str in context.user_data["pref_preacher_dates"]:
        await query.answer(
            "You're the Preacher that day — you're automatically waived from other roles.", show_alert=True
        )
        return PREF_SELECT_DATE

    if not pref_round_still_open(context):
        await query.answer()
        await query.edit_message_text("This preference round just closed — sorry!")
        return ConversationHandler.END

    now_unavailable = toggle_unavailable(ss, round_["Service"], date_str, name)
    marked = context.user_data["pref_marked"]
    if now_unavailable:
        marked.add(date_str)
    else:
        marked.discard(date_str)
    await query.answer("Marked as not available" if now_unavailable else "Marked as available again")
    return await pref_show_dates(update, context)


# ---------------------------------------------------------------------------
# /mark_broadcast: set every role for one service+date to Live Broadcast in
# one shot, for the days the service is relayed from another church rather
# than run locally. Not available for Sun Stop Sundays.
# ---------------------------------------------------------------------------

MARK_BC_SERVICE, MARK_BC_MONTH, MARK_BC_DATE, MARK_BC_CONFIRM = range(170, 174)


async def mark_broadcast_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ss = setup_sheet()
    context.user_data["bc_ss"] = ss
    tabs = [s for s in get_all_schedule_tabs(ss) if s != "SunStopSundays"]
    buttons = [[InlineKeyboardButton(s, callback_data=s)] for s in tabs]
    await update.message.reply_text(
        "Mark a date as Live Broadcast for which service?", reply_markup=InlineKeyboardMarkup(buttons)
    )
    return MARK_BC_SERVICE


async def mark_broadcast_select_service(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["bc_service"] = query.data
    await query.edit_message_text("Which month?", reply_markup=month_keyboard())
    return MARK_BC_MONTH


async def mark_broadcast_select_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year, month = map(int, query.data.split("-"))
    ss = context.user_data["bc_ss"]
    service = context.user_data["bc_service"]
    configs = load_service_configs(ss)
    cadence = configs.get(service, {}).get("cadence", "month")

    if service == "Predawn":
        dates = dates_matching_weekdays_in_month(year, month, PREDAWN_WEEKDAYS)
    elif cadence == "week":
        # unknown weekly pattern for a generic custom service — fall back to every day
        dates = dates_matching_weekdays_in_month(year, month, list(range(7)))
    else:
        dates = get_service_month_dates(ss, service, year, month)

    if not dates:
        await query.edit_message_text("No dates found for that service/month.")
        return ConversationHandler.END

    context.user_data["bc_dates"] = dates
    buttons = [[InlineKeyboardButton(d.strftime("%b %d, %Y"), callback_data=d.isoformat())] for d in dates]
    await query.edit_message_text("Which date?", reply_markup=InlineKeyboardMarkup(buttons))
    return MARK_BC_DATE


async def mark_broadcast_select_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    date_str = query.data
    context.user_data["bc_date"] = date_str
    ss = context.user_data["bc_ss"]
    service = context.user_data["bc_service"]

    ws = ss.worksheet(service)
    existing = [r for r in ws.get_all_records() if r.get("Date") == date_str]
    existing_summary = ", ".join(f"{r['Role']}: {r['Partaker']}" for r in existing) if existing else "(nothing set yet)"

    buttons = [
        [InlineKeyboardButton("Yes, mark as Live Broadcast", callback_data="yes")],
        [InlineKeyboardButton("Cancel", callback_data="no")],
    ]
    await query.edit_message_text(
        f"{service} on {date_str} currently: {existing_summary}\n\n"
        f"Set every non-Tech role for this date to Live Broadcast? This overwrites anything already "
        f"assigned (Tech is left alone — it stays blank unless already logged via /log_tech).",
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return MARK_BC_CONFIRM


async def mark_broadcast_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data != "yes":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END

    ss = context.user_data["bc_ss"]
    service = context.user_data["bc_service"]
    date_str = context.user_data["bc_date"]
    configs = load_service_configs(ss)
    roles = list(configs.get(service, {}).get("roles", {}).keys())
    random_roles = get_random_roles_for_service(ss, service)
    tech_roles = set(TECH_ROLES_BY_SERVICE.get(service, {}).keys())

    ws = ss.worksheet(service)
    records = ws.get_all_records()
    row_index = {(r["Date"], r["Role"]): i + 2 for i, r in enumerate(records)}
    counts = load_assignment_counts(ss)
    new_rows = []
    marked = 0

    for role in roles:
        if role in tech_roles:
            continue  # Tech is never auto-set — stays blank until logged via /log_tech, and is never overwritten if already logged
        key = (date_str, role)
        old_partaker = records[row_index[key] - 2]["Partaker"] if key in row_index else None
        if key in row_index:
            ws.update_cell(row_index[key], 3, LIVE_BROADCAST)
        else:
            new_rows.append([date_str, role, LIVE_BROADCAST, "scheduled"])
        log_adjustment(ss, service, date_str, role, old_partaker, LIVE_BROADCAST, "live_broadcast")
        if role in random_roles and old_partaker and old_partaker != LIVE_BROADCAST:
            counts[old_partaker] = max(0, counts[old_partaker] - 1)
        marked += 1

    if new_rows:
        ws.append_rows(new_rows)
    save_assignment_counts(ss, counts)

    text = (f"📡 {service} on {date_str} marked as Live Broadcast for {marked} role(s) "
            f"(Tech left as-is — blank until logged via /log_tech, unchanged if already set).")
    await query.edit_message_text(text)
    await announce_update(context, ss, text)
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# /menu — one inline keyboard for every admin command
# ---------------------------------------------------------------------------
# Each button re-uses the existing command's own flow. A command's start
# function normally replies through update.message, which doesn't exist on a
# button tap, so _ButtonAsCommand presents the tapped menu message as
# update.message — to the start function it looks like the command was typed.

class _ButtonAsCommand:
    def __init__(self, update):
        self._update = update
        self.message = update.callback_query.message
        self.callback_query = None

    def __getattr__(self, name):
        return getattr(self._update, name)


def from_menu(start_fn):
    async def _entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        try:
            await query.edit_message_reply_markup(reply_markup=None)  # retire the tapped menu
        except Exception:
            pass
        return await start_fn(_ButtonAsCommand(update), context)
    return _entry


def menu_entry(key, start_fn):
    """Extra ConversationHandler entry point: the /menu button for `key`."""
    return CallbackQueryHandler(from_menu(start_fn), pattern=f"^menu:{key}$")


def _menu_btn(label, key):
    return InlineKeyboardButton(label, callback_data=f"menu:{key}")


MENU_TITLE = "Scheduling menu — what would you like to do?"

MAIN_MENU_ROWS = [
    [_menu_btn("📅 Generate schedule", "generate")],
    [_menu_btn("🎤 Set preacher", "set_preacher"), _menu_btn("🌄 Set Predawn pattern", "set_predawn_pattern")],
    [_menu_btn("🎛 Log tech", "log_tech"), _menu_btn("📝 Log a role", "log_role")],
    [_menu_btn("🔄 Adjustments ›", "adjust")],
    [_menu_btn("🔎 Pull schedule", "pull_schedule"), _menu_btn("👤 Pull person", "pull_person")],
    [_menu_btn("➕ Add service", "add_service"), _menu_btn("➕ Add role", "add_role")],
    [_menu_btn("👥 Roster", "roster"), _menu_btn("🗳 Preferences ›", "prefs")],
    [_menu_btn("📤 Bulk upload (template)", "template")],
    [_menu_btn("🗓 Yearly renewal", "renew_year")],
    [_menu_btn("⚙️ More ›", "more")],
]

# A URL button opens the link directly (no bot round-trip), so it works from the
# /menu message and from the /start welcome message alike.
if GUIDE_URL:
    MAIN_MENU_ROWS.append([InlineKeyboardButton("📖 Guide to the bot", url=GUIDE_URL)])

SUBMENUS = {
    "adjust": ("Adjustments — what would you like to do?", [
        [_menu_btn("Cancel a role", "cancel_role")],
        [_menu_btn("Swap dates", "swap")],
        [_menu_btn("Substitute a partaker", "substitute")],
        [_menu_btn("Special request", "special_request")],
        [_menu_btn("Mark live broadcast", "mark_broadcast")],
    ]),
    "prefs": ("Preferences — what would you like to do?", [
        [_menu_btn("Open a preference round", "open_preferences")],
        [_menu_btn("Close a round early", "close_preferences")],
    ]),
    "more": ("More — what would you like to do?", [
        [_menu_btn("Refresh dashboards", "refresh_dashboard")],
        [_menu_btn("Sync dashboards", "sync_dashboard")],
        [_menu_btn("Register this group for reminders", "set_group_chat")],
    ]),
}


def _end_all_conversations(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Drop whatever command the user was in the middle of, so a menu tap
    always starts fresh instead of being read as an answer to the old step."""
    for handlers in context.application.handlers.values():
        for handler in handlers:
            if isinstance(handler, ConversationHandler):
                try:
                    handler._update_state(ConversationHandler.END, handler._get_key(update))
                except Exception:
                    pass  # unfamiliar library internals: skip, the menu still works


async def menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    _end_all_conversations(update, context)
    await update.message.reply_text(MENU_TITLE, reply_markup=InlineKeyboardMarkup(MAIN_MENU_ROWS))


async def menu_nav(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    key = query.data.split(":", 1)[1]
    if key == "main":
        await query.edit_message_text(MENU_TITLE, reply_markup=InlineKeyboardMarkup(MAIN_MENU_ROWS))
        return
    title, rows = SUBMENUS[key]
    await query.edit_message_text(
        title, reply_markup=InlineKeyboardMarkup(rows + [[_menu_btn("‹ Back", "main")]])
    )


# --- Global error handler: reply to the user instead of failing silently ---
# python-telegram-bot logs "No error handlers are registered" and swallows
# any unhandled exception otherwise — the person who triggered it never
# hears back at all. Registered at the end of build_app().
async def error_handler(update, context: ContextTypes.DEFAULT_TYPE):
    err = context.error
    logging.getLogger(__name__).error("Unhandled exception", exc_info=err)
    if not isinstance(update, Update):
        return

    status = getattr(getattr(err, "response", None), "status_code", None)
    if isinstance(err, _GspreadAPIError) and status == 429:
        text = ("⚠️ Google Sheets is rate-limiting me right now (too many requests). "
                "Please wait about a minute and try again.")
    else:
        text = (f"⚠️ Something went wrong ({type(err).__name__}). "
                f"Please try again in a moment — if it keeps happening, tell the admin.")

    # If this came from a button tap, replace the stuck message (e.g. "Generating
    # schedule...") so it never just hangs; otherwise send a fresh reply.
    try:
        if update.callback_query:
            await update.callback_query.edit_message_text(text)
        elif update.effective_message:
            await update.effective_message.reply_text(text)
    except Exception:
        try:
            if update.effective_message:
                await update.effective_message.reply_text(text)
        except Exception:
            pass


def build_app():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    # /menu and its navigation come first so they always win over an open command
    app.add_handler(CommandHandler("menu", menu_command))
    app.add_handler(CallbackQueryHandler(menu_nav, pattern=r"^menu:(main|adjust|prefs|more)$"))
    app.add_handler(CallbackQueryHandler(from_menu(refresh_dashboard_command), pattern=r"^menu:refresh_dashboard$"))
    app.add_handler(CallbackQueryHandler(from_menu(sync_dashboard_command), pattern=r"^menu:sync_dashboard$"))

    generate_conv = ConversationHandler(
        entry_points=[CommandHandler("generate", generate_schedule_start), menu_entry("generate", generate_schedule_start)],
        states={
            SELECT_SERVICE: [CallbackQueryHandler(select_service)],
            SELECT_MONTH: [CallbackQueryHandler(select_month)],
            PREDAWN_GEN_MONTH: [CallbackQueryHandler(predawn_generate_month)],
            PREDAWN_ADJUST: [CallbackQueryHandler(predawn_adjust_response)],
            SELECT_SUNSTOP_MONTH: [CallbackQueryHandler(select_sunstop_month)],
            PREDAWN_PATTERN_DAY: [CallbackQueryHandler(predawn_pattern_pick_day)],
            GEN_SVC_SELECT_PERIOD: [CallbackQueryHandler(generate_service_period)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    preacher_conv = ConversationHandler(
        entry_points=[CommandHandler("set_preacher", set_preacher_start), menu_entry("set_preacher", set_preacher_start)],
        states={
            PREACHER_SELECT_SERVICE: [CallbackQueryHandler(preacher_select_service)],
            PREACHER_SELECT_MONTH: [CallbackQueryHandler(preacher_select_month)],
            PREACHER_PICK: [CallbackQueryHandler(preacher_pick)],
            PREACHER_CONFIRM_CONFLICT: [CallbackQueryHandler(preacher_confirm_conflict)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    predawn_pattern_conv = ConversationHandler(
        entry_points=[CommandHandler("set_predawn_pattern", set_predawn_pattern_start), menu_entry("set_predawn_pattern", set_predawn_pattern_start)],
        states={
            PREDAWN_PATTERN_DAY: [CallbackQueryHandler(predawn_pattern_pick_day)],
            PREDAWN_GEN_MONTH: [CallbackQueryHandler(predawn_generate_month)],
            PREDAWN_ADJUST: [CallbackQueryHandler(predawn_adjust_response)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    predawn_conv = ConversationHandler(
        entry_points=[CommandHandler("generate_predawn", generate_predawn_start)],
        states={
            PREDAWN_GEN_MONTH: [CallbackQueryHandler(predawn_generate_month)],
            PREDAWN_ADJUST: [CallbackQueryHandler(predawn_adjust_response)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(predawn_pattern_conv)
    sunstop_conv = ConversationHandler(
        entry_points=[CommandHandler("generate_sunstop", generate_sunstop_start)],
        states={SELECT_SUNSTOP_MONTH: [CallbackQueryHandler(select_sunstop_month)]},
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    tech_conv = ConversationHandler(
        entry_points=[CommandHandler("log_tech", log_tech_start), menu_entry("log_tech", log_tech_start)],
        states={
            TECH_SELECT_SERVICE: [CallbackQueryHandler(tech_select_service)],
            TECH_SELECT_ROLE: [CallbackQueryHandler(tech_select_role)],
            TECH_SELECT_PERIOD: [CallbackQueryHandler(tech_select_period)],
            TECH_PICK: [CallbackQueryHandler(tech_pick)],
            TECH_CONFIRM_CONFLICT: [CallbackQueryHandler(tech_confirm_conflict)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    add_service_conv = ConversationHandler(
        entry_points=[CommandHandler("add_service", add_service_start), menu_entry("add_service", add_service_start)],
        states={
            ADD_SVC_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_service_name)],
            ADD_SVC_CADENCE: [CallbackQueryHandler(add_service_cadence)],
            ADD_SVC_ROLE_NAME_ONLY: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_service_role_name_only)],
            ADD_SVC_ROLE_MODE_ONLY: [CallbackQueryHandler(add_service_role_mode_only)],
            ADD_SVC_MORE_ROLES: [CallbackQueryHandler(add_service_more_roles)],
            ADD_SVC_ROLE_ELIGIBLE: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_service_role_eligible)],
            ADD_SVC_READY: [CallbackQueryHandler(add_service_ready_response)],
            ADD_SVC_GEN_OR_UPLOAD: [CallbackQueryHandler(add_service_gen_or_upload)],
            ADD_SVC_PERIOD: [CallbackQueryHandler(add_service_period)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    add_role_conv = ConversationHandler(
        entry_points=[CommandHandler("add_role", add_role_start), menu_entry("add_role", add_role_start)],
        states={
            ADD_ROLE_SELECT_SERVICE: [CallbackQueryHandler(add_role_select_service)],
            ADD_SVC_ROLE_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_service_role_name)],
            ADD_SVC_ELIGIBLE: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_service_eligible)],
            ADD_SVC_MODE: [CallbackQueryHandler(add_role_finish)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    generate_service_conv = ConversationHandler(
        entry_points=[CommandHandler("generate_service", generate_service_start)],
        states={
            GEN_SVC_SELECT_SERVICE: [CallbackQueryHandler(generate_service_select)],
            GEN_SVC_SELECT_PERIOD: [CallbackQueryHandler(generate_service_period)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    log_role_conv = ConversationHandler(
        entry_points=[CommandHandler("log_role", log_role_start), menu_entry("log_role", log_role_start)],
        states={
            LOG_ROLE_SELECT_SERVICE: [CallbackQueryHandler(log_role_select_service)],
            LOG_ROLE_SELECT_ROLE: [CallbackQueryHandler(log_role_select_role)],
            LOG_ROLE_SELECT_PERIOD: [CallbackQueryHandler(log_role_select_period)],
            LOG_ROLE_PICK: [CallbackQueryHandler(log_role_pick)],
            LOG_ROLE_CONFIRM_CONFLICT: [CallbackQueryHandler(log_role_confirm_conflict)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    pull_schedule_conv = ConversationHandler(
        entry_points=[CommandHandler("pull_schedule", pull_schedule_start), menu_entry("pull_schedule", pull_schedule_start)],
        states={
            PULL_SELECT_SERVICE: [CallbackQueryHandler(pull_select_service)],
            PULL_SELECT_PERIOD: [CallbackQueryHandler(pull_select_period)],
            PULL_PICK_MONTH: [CallbackQueryHandler(pull_pick_month)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    pull_person_conv = ConversationHandler(
        entry_points=[CommandHandler("pull_person", pull_person_start), menu_entry("pull_person", pull_person_start)],
        states={
            PULL_PERSON_NAME: [CallbackQueryHandler(pull_person_name)],
            PULL_PERSON_MODE: [CallbackQueryHandler(pull_person_mode)],
            PULL_PERSON_SERVICE: [CallbackQueryHandler(pull_person_service)],
            PULL_PERSON_MONTH: [CallbackQueryHandler(pull_person_month)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    set_group_chat_conv = ConversationHandler(
        entry_points=[CommandHandler("set_group_chat", set_group_chat_start), menu_entry("set_group_chat", set_group_chat_start)],
        states={SET_GROUP_PURPOSE: [CallbackQueryHandler(set_group_purpose)]},
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    swap_conv = ConversationHandler(
        entry_points=[CommandHandler("swap", swap_start), menu_entry("swap", swap_start)],
        states={
            SWAP_SERVICE: [CallbackQueryHandler(swap_select_service)],
            SWAP_ROLE: [CallbackQueryHandler(swap_select_role)],
            SWAP_DATE_A: [CallbackQueryHandler(swap_select_date_a)],
            SWAP_DATE_B: [CallbackQueryHandler(swap_select_date_b)],
            SWAP_CONFIRM: [CallbackQueryHandler(swap_confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    substitute_conv = ConversationHandler(
        entry_points=[CommandHandler("substitute", substitute_start), menu_entry("substitute", substitute_start)],
        states={
            SUB_SERVICE: [CallbackQueryHandler(substitute_select_service)],
            SUB_ROLE: [CallbackQueryHandler(substitute_select_role)],
            SUB_DATE: [CallbackQueryHandler(substitute_select_date)],
            SUB_NEW: [CallbackQueryHandler(substitute_pick_new)],
            SUB_CONFIRM_CONFLICT: [CallbackQueryHandler(substitute_confirm_conflict)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    special_request_conv = ConversationHandler(
        entry_points=[CommandHandler("special_request", special_request_start), menu_entry("special_request", special_request_start)],
        states={
            SPECIAL_SERVICE: [CallbackQueryHandler(special_select_service)],
            SPECIAL_PERSON: [CallbackQueryHandler(special_select_person)],
            SPECIAL_MONTH: [CallbackQueryHandler(special_select_month)],
            SPECIAL_MAX: [CallbackQueryHandler(special_apply)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    roster_conv = ConversationHandler(
        entry_points=[CommandHandler("roster", roster_start), menu_entry("roster", roster_start)],
        states={
            ROSTER_MENU: [CallbackQueryHandler(roster_menu)],
            ROSTER_ADD_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, roster_add_name)],
            ROSTER_SERVICE: [CallbackQueryHandler(roster_select_service)],
            ROSTER_ROLE: [CallbackQueryHandler(roster_select_role)],
            ROSTER_ADDROLE_NAME: [CallbackQueryHandler(roster_addrole_name)],
            ROSTER_REMOVEROLE_NAME: [CallbackQueryHandler(roster_removerole_name)],
            ROSTER_REMOVEALL_NAME: [CallbackQueryHandler(roster_removeall_name)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    template_conv = ConversationHandler(
        entry_points=[CommandHandler("template", template_start), menu_entry("template", template_start)],
        states={
            TEMPLATE_SELECT_SERVICE: [CallbackQueryHandler(template_select_service)],
            TEMPLATE_SELECT_PERIOD: [CallbackQueryHandler(template_select_period)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(generate_conv)
    app.add_handler(preacher_conv)
    app.add_handler(predawn_conv)
    app.add_handler(sunstop_conv)
    app.add_handler(tech_conv)
    app.add_handler(add_service_conv)
    app.add_handler(add_role_conv)
    app.add_handler(generate_service_conv)
    app.add_handler(log_role_conv)
    app.add_handler(pull_schedule_conv)
    app.add_handler(pull_person_conv)
    app.add_handler(set_group_chat_conv)
    app.add_handler(swap_conv)
    app.add_handler(substitute_conv)
    app.add_handler(special_request_conv)
    app.add_handler(roster_conv)
    app.add_handler(template_conv)
    app.add_handler(MessageHandler(filters.Document.ALL, handle_schedule_upload))
    app.add_handler(CommandHandler("refresh_dashboard", refresh_dashboard_command))
    app.add_handler(CommandHandler("sync_dashboard", sync_dashboard_command))
    renew_year_conv = ConversationHandler(
        entry_points=[CommandHandler("renew_year", renew_year_start), menu_entry("renew_year", renew_year_start)],
        states={RENEW_CONFIRM: [CallbackQueryHandler(renew_year_confirm)]},
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(renew_year_conv)

    # #4: nightly reminder jobs (Asia/Manila). run_daily fires once per day;
    # each callback itself decides whether today/tomorrow actually needs a
    # reminder, so a single daily job safely covers every service.
    app.job_queue.run_daily(send_sunday_service_reminder, time=dt.time(hour=20, minute=30, tzinfo=CHURCH_TZ))
    app.job_queue.run_daily(send_wednesday_service_reminder, time=dt.time(hour=19, minute=0, tzinfo=CHURCH_TZ))
    app.job_queue.run_daily(send_friday_reminder, time=dt.time(hour=20, minute=0, tzinfo=CHURCH_TZ))

    open_preferences_conv = ConversationHandler(
        entry_points=[CommandHandler("open_preferences", open_preferences_start), menu_entry("open_preferences", open_preferences_start)],
        states={
            OPEN_PREF_SERVICE: [CallbackQueryHandler(open_pref_select_service)],
            OPEN_PREF_MONTH: [CallbackQueryHandler(open_pref_select_month)],
            OPEN_PREF_DEADLINE: [CallbackQueryHandler(open_pref_select_deadline)],
            OPEN_PREF_CUSTOM_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, open_pref_custom_deadline)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    claim_preferences_conv = ConversationHandler(
        entry_points=[CommandHandler("start", start_command)],
        states={
            PREF_SELECT_NAME: [CallbackQueryHandler(pref_select_name)],
            PREF_SELECT_DATE: [CallbackQueryHandler(pref_select_date)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(open_preferences_conv)
    app.add_handler(claim_preferences_conv)
    close_preferences_conv = ConversationHandler(
        entry_points=[CommandHandler("close_preferences", close_preferences_start), menu_entry("close_preferences", close_preferences_start)],
        states={CLOSE_PREF_SELECT: [CallbackQueryHandler(close_preferences_select)]},
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(close_preferences_conv)
    app.job_queue.run_once(startup_recover_rounds, when=5)

    mark_broadcast_conv = ConversationHandler(
        entry_points=[CommandHandler("mark_broadcast", mark_broadcast_start), menu_entry("mark_broadcast", mark_broadcast_start)],
        states={
            MARK_BC_SERVICE: [CallbackQueryHandler(mark_broadcast_select_service)],
            MARK_BC_MONTH: [CallbackQueryHandler(mark_broadcast_select_month)],
            MARK_BC_DATE: [CallbackQueryHandler(mark_broadcast_select_date)],
            MARK_BC_CONFIRM: [CallbackQueryHandler(mark_broadcast_confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(mark_broadcast_conv)

    cancel_role_conv = ConversationHandler(
        entry_points=[CommandHandler("cancel_role", cancel_role_start), menu_entry("cancel_role", cancel_role_start)],
        states={
            CANCEL_SERVICE: [CallbackQueryHandler(cancel_select_service)],
            CANCEL_NAME: [CallbackQueryHandler(cancel_select_name)],
            CANCEL_PICK: [CallbackQueryHandler(cancel_pick_entry)],
            CANCEL_REPLACEMENT: [CallbackQueryHandler(cancel_pick_replacement)],
            CANCEL_CONFIRM: [CallbackQueryHandler(cancel_confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(cancel_role_conv)

    app.add_error_handler(error_handler)
    return app


if __name__ == "__main__":
    app = build_app()
    app.run_polling()