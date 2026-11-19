#!/usr/bin/env python
"""upkeep.py (upkeep-beta) - keeps the apps on a Windows machine current.

Most people update an app when they need something from it; the apps nobody
needed to touch stay unpatched. This tool lists what winget can update, keeps
those apps current, and names the apps it cannot reach.

Subcommands:
  setup [--every .. --at ..]   create the home folder and, when asked, the schedule
  status                       summary of the last run
  apps [--match] [--json]      every installed app and what upkeep does with it
  select [...]                 choose the apps (default: all catalogue apps)
  schedule [...]               choose when it runs (default: no task)
  mode [report|auto]           report only (default) or install by itself
  daily [--dry-run]            the scheduled run: list or install, check, record, toast
  apply [--only id,id]         install the listed updates now
  check [--json]               the item check only; installs and records nothing
  claude [--dry-run]           Claude Code CLI and plugins only
  pins [--apply]               held package ids that lack a winget pin
  doctor                       is everything this tool needs in place
  diag                         a redacted block to send as feedback

Every install is judged by effect: installed versions are read before and
after, and an upgrade counts only when the version changed. One installer runs
at a time. Nothing here elevates. In report mode nothing is installed.

Every external call goes through an injected `runner` (subprocess.run by
default), so tests run with fakes. Standard library only. Windows only.
"""
from __future__ import annotations

import argparse
import base64
import csv
import fnmatch
import hashlib
import io
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from types import SimpleNamespace


__version__ = "0.1.1"
APP = "upkeep-beta"
MIN_PYTHON = (3, 9)

if sys.version_info < MIN_PYTHON:   # the message a too-old Python gets, instead of a traceback
    sys.stderr.write(
        "upkeep-beta needs Python %d.%d or newer; this is Python %s.\n"
        "Update Python, then run the command again: https://www.python.org/downloads/windows/\n"
        % (MIN_PYTHON[0], MIN_PYTHON[1], sys.version.split()[0]))
    sys.exit(3)


def home_dir() -> Path:
    """%LOCALAPPDATA%\\upkeep-beta, or UPKEEP_BETA_HOME when set."""
    env = os.environ.get("UPKEEP_BETA_HOME")
    if env:
        return Path(env)
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / APP


HOME = home_dir()
POLICY = HOME / "config" / "policy.json"
ITEMS = HOME / "config" / "items.json"
KNOWN = HOME / "config" / "known_apps.json"
STATE = HOME / "state"
SHIPPED = Path(__file__).resolve().parent.parent / "config"   # the defaults shipped beside scripts/
TASK_NAME = "upkeep-beta"
SKILL = "/upkeep-beta:upkeep"


def _resolve_winget() -> str:
    found = shutil.which("winget.exe")
    if found:
        return found
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return str(Path(base) / "Microsoft" / "WindowsApps" / "winget.exe")


def _resolve_powershell() -> str:
    root = os.environ.get("SystemRoot") or r"C:\Windows"
    return str(Path(root) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe")


WINGET = _resolve_winget()
POWERSHELL = _resolve_powershell()
PY_MANAGER = "py"   # the Python install manager; its `list --online` names the newest patch release


TOAST_TAG = "upkeep-beta"


TOAST_GROUP = "upkeep-beta"
STALE_HOURS = {"daily": 36, "weekdays": 84, "weekly": 180}   # a missed run is red after this long
DAYS = ("mon", "tue", "wed", "thu", "fri", "sat", "sun")
FAIL_LIMIT = 2              # consecutive failed runs before a package waits for the user


WINGET_UPGRADE_TIMEOUT = 900   # default per package; policy `upgrade_timeout_s` overrides
WINGET_TIME_BUDGET = 7200   # no new upgrade attempt starts this many seconds after the first one
DEFER_LIMIT = 3             # consecutive deferrals before a package waits for the user
CLAUDE_UPDATE_TIMEOUT = 900
CLAUDE_PLUGIN_TIMEOUT = 300
CLAUDE_PLUGINS_FILE = "%USERPROFILE%/.claude/plugins/installed_plugins.json"
ADMIN_NOTE = "must run in an administrator PowerShell"
RANK = {"red": 3, "yellow": 2, "unknown": 1, "green": 0}

# Windows PowerShell 5.1 native WinRT toast.
_TOAST_PS = r"""
$ErrorActionPreference = 'Stop'
try {
  [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
  [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
  $app = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
  $n = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app)
  if ([string]$n.Setting -ne 'Enabled') { [Console]::Out.Write('NOTIFICATIONS ' + $n.Setting); exit 3 }
  $t = [Security.SecurityElement]::Escape($env:UPKEEP_TOAST_TITLE)
  $b = [Security.SecurityElement]::Escape($env:UPKEEP_TOAST_BODY)
  if ($env:UPKEEP_TOAST_STICKY -eq '1') {
    $xml = "<toast scenario='reminder'><visual><binding template='ToastGeneric'><text>$t</text><text>$b</text></binding></visual><actions><action content='OK' arguments='dismiss' activationType='system'/></actions></toast>"
  } else {
    $xml = "<toast><visual><binding template='ToastGeneric'><text>$t</text><text>$b</text></binding></visual></toast>"
  }
  $d = New-Object Windows.Data.Xml.Dom.XmlDocument
  $d.LoadXml($xml)
  $toast = [Windows.UI.Notifications.ToastNotification]::new($d)
  if ($env:UPKEEP_TOAST_TAG) { $toast.Tag = $env:UPKEEP_TOAST_TAG; $toast.Group = $env:UPKEEP_TOAST_GROUP }
  $n.Show($toast)
  [Console]::Out.Write('SHOWN')
  exit 0
} catch { [Console]::Out.Write('ERROR ' + $_.Exception.Message); exit 2 }
"""


# ---------------------------------------------------------------- helpers -----

def now_local() -> datetime:
    return datetime.now()


def xp(s) -> str:
    """Expand {python} and %ENV% in a configured string."""
    return os.path.expandvars(str(s).replace("{python}", sys.executable))


def xpath(s) -> Path:
    return Path(xp(s))


def parse_date(s):
    """'YYYY-MM-DD' (or an ISO datetime) -> date, else None."""
    if not s:
        return None
    m = re.match(r"\s*(\d{4})-(\d{2})-(\d{2})", str(s))
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def parse_datetime(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(str(s).replace("Z", "+00:00")).replace(tzinfo=None)
    except ValueError:
        d = parse_date(s)
        return datetime(d.year, d.month, d.day) if d else None


def read_json(path, default=None):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def write_json(path: Path, obj) -> None:
    """Atomic rewrite (tmp + replace)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
        f.write("\n")
    os.replace(tmp, path)


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    os.replace(tmp, path)


def tail_lines(text: str, n: int = 5) -> list:
    lines = [ln.rstrip() for ln in (text or "").replace("\r", "\n").split("\n")]
    return [ln for ln in lines if ln.strip()][-n:]


def run_cmd(runner, cmd, timeout=120, env_extra=None, cwd=None):
    """Run through the injected runner; never raises. A launch failure comes
    back as a negative returncode with the reason in stderr."""
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    if env_extra:
        env.update(env_extra)
    kw = {}
    if os.name == "nt" and sys.executable.lower().endswith("pythonw.exe"):
        kw["creationflags"] = 0x08000000   # CREATE_NO_WINDOW: the scheduled run opens no console
    try:
        r = runner([str(c) for c in cmd], capture_output=True, text=True,
                   encoding="utf-8", errors="replace", timeout=timeout,
                   env=env, cwd=cwd, **kw)
    except subprocess.TimeoutExpired:
        return SimpleNamespace(returncode=-998, stdout="", stderr="timeout after %s s" % timeout)
    except FileNotFoundError as e:
        return SimpleNamespace(returncode=-997, stdout="", stderr="not found: %s" % e)
    except Exception as e:  # noqa: BLE001 - report, never crash the pass
        return SimpleNamespace(returncode=-999, stdout="", stderr="launch failed: %s" % e)
    return SimpleNamespace(returncode=getattr(r, "returncode", 0),
                           stdout=getattr(r, "stdout", "") or "",
                           stderr=getattr(r, "stderr", "") or "")


def match_any(pkg_id: str, patterns) -> bool:
    return any(fnmatch.fnmatchcase(pkg_id, p) for p in (patterns or []))


def lookup(mapping: dict, pkg_id: str):
    """Exact key first, then fnmatch patterns (e.g. 'JetBrains.*')."""
    if not mapping:
        return None
    if pkg_id in mapping:
        return mapping[pkg_id]
    for k, v in mapping.items():
        if any(ch in k for ch in "*?[") and fnmatch.fnmatchcase(pkg_id, k):
            return v
    return None


def load_policy(path=None) -> dict:
    p = read_json(path or POLICY)
    if not isinstance(p, dict):
        return None
    for k, v in (("held", []), ("held_override", []), ("admin", []), ("quiet", []),
                 ("check_after", {}), ("in_use", {})):
        p.setdefault(k, v)
    if p.get("mode") not in ("report", "auto"):
        p["mode"] = "report"
    sel = p.get("selection") if isinstance(p.get("selection"), dict) else {}
    p["selection"] = {"mode": "named" if sel.get("mode") == "named" else "all",
                      "excluded": list(sel.get("excluded") or []),
                      "only": list(sel.get("only") or [])}
    if not isinstance(p.get("schedule"), dict):
        p["schedule"] = None
    p.setdefault("upgrade_timeout_s", WINGET_UPGRADE_TIMEOUT)
    p.setdefault("fail_limit", FAIL_LIMIT)
    return p


def load_known() -> dict:
    d = read_json(KNOWN, {})
    return d if isinstance(d, dict) else {}


def is_selected(pkg_id: str, policy: dict) -> bool:
    """Default: every catalogue app is covered. `excluded` narrows it; mode
    `named` covers only the ids (or patterns) in `only`."""
    sel = policy.get("selection") or {}
    if sel.get("mode") == "named":
        return match_any(pkg_id, sel.get("only"))
    return not match_any(pkg_id, sel.get("excluded"))


def load_items(path=None) -> list:
    d = read_json(path or ITEMS, {})
    items = d.get("items") if isinstance(d, dict) else d
    return items if isinstance(items, list) else []


def item_by_id(items, item_id):
    for it in items:
        if it.get("id") == item_id:
            return it
    return None


# ------------------------------------------------------- winget text parsing ----

def parse_winget_tables(text: str) -> list:
    """Parse every winget text table in `text`.

    A table is a header line followed by a line of dashes. Columns are cut by
    the character positions of the header words, so the layout works in
    Portuguese ('Nome ID Versão Disponível Origem') and English alike. Returns
    a list of tables: {"header": [words], "rows": [dict word.lower() -> str]}.
    A table body ends at a blank line or at winget's count footer
    ('N atualizações disponíveis.'), which is never a row."""
    lines = []
    for raw in (text or "").replace("\r\n", "\n").split("\n"):
        lines.append(raw.split("\r")[-1].rstrip())  # drop spinner overwrites
    tables = []
    i = 0
    while i < len(lines):
        if re.fullmatch(r"-{5,}", lines[i].strip() or "x") and i > 0 and lines[i - 1].strip():
            header = lines[i - 1]
            words = [(m.start(), m.group(0)) for m in re.finditer(r"\S+", header)]
            starts = [w[0] for w in words]
            rows = []
            j = i + 1
            while j < len(lines) and lines[j].strip() and not _FOOTER_RE.match(lines[j]):
                ln = lines[j]
                cells = {}
                for k, (st, word) in enumerate(words):
                    en = starts[k + 1] if k + 1 < len(starts) else len(ln)
                    cells[word.lower()] = ln[st:en].strip()
                rows.append(cells)
                j += 1
            tables.append({"header": [w[1] for w in words], "rows": rows})
            i = j
        else:
            i += 1
    return tables


_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.+\-]*$")
_FOOTER_RE = re.compile(r"^\s*\d+\s+\S.*(dispon[ií]v|available|upgrades?)", re.I)
NO_UPGRADES_PHRASES = (
    "nenhum pacote instalado", "nenhuma atualiza", "não há atualiza", "nao ha atualiza",
    "no installed package found", "no available upgrade", "no newer package",
)
NO_PINS_PHRASES = ("não há", "nao ha", "no pins", "nenhum", "no pin")


def _wide(table) -> bool:
    return len(table["header"]) >= 3


def _cols(table, kind="upgrade") -> dict:
    """Column keys by position, so the header language does not matter: name, id,
    version; then `available` (4th column of an upgrade table, or of a list table
    that has five columns) and `source` (last column). A header word 'id' wins."""
    keys = [h.lower() for h in table["header"]]
    get = lambda i: keys[i] if len(keys) > i else ""  # noqa: E731
    out = {"name": get(0), "id": "id" if "id" in keys else get(1), "version": get(2)}
    if kind == "upgrade":
        out["available"] = get(3)
        out["source"] = get(4)
    else:
        out["available"] = get(3) if len(keys) >= 5 else ""
        out["source"] = keys[-1] if len(keys) >= 4 else ""
    return out


def _table_ids(table, installed=None, idkey="id"):
    """-> (rows with a usable id, count of rows whose id was cut off).
    An id is usable when it has an internal dot and does not end with one
    (Publisher.Package), or when it is a key of the installed-version map."""
    out, bad = [], 0
    for row in table["rows"]:
        pid = row.get(idkey, "")
        if _ID_RE.match(pid) and (("." in pid[1:] and not pid.endswith(".")) or pid in (installed or {})):
            out.append(row)
        elif pid and "…" in pid:
            bad += 1
    return out, bad


def _first_table(text):
    tables = [t for t in parse_winget_tables(text) if _wide(t)]
    return next((t for t in tables if "id" in [h.lower() for h in t["header"]]), tables[0] if tables else None)


def parse_upgrade_list(text: str, installed=None, rc=None):
    """-> (list of {id,name,version,available}, warnings, ok).
    ok False means the text could not be understood (a RED result). No table
    together with exit code 0 is an empty list in any language; without the
    exit code the Portuguese and English phrases decide.
    `installed` ({id: version}) lets an id without a dot through when it is installed."""
    first = _first_table(text)
    if first is None:
        low = (text or "").lower()
        if rc == 0 or any(p in low for p in NO_UPGRADES_PHRASES):
            return [], [], True
        return [], [], False
    c = _cols(first, "upgrade")
    rows, bad = _table_ids(first, installed, c["id"])
    warnings = []
    if bad:
        warnings.append("%d upgrade row(s) with a truncated id were skipped" % bad)
    pkgs = [{"id": r[c["id"]], "name": r.get(c["name"], ""), "version": r.get(c["version"], ""),
             "available": r.get(c["available"], "") if c["available"] else ""} for r in rows]
    return pkgs, warnings, True


def parse_upgrade_notes(text: str, installed=None) -> list:
    """What winget lists but this pass leaves alone: packages in a second table
    (they need explicit targeting) and packages whose version winget cannot read."""
    notes = []
    tables = [t for t in parse_winget_tables(text) if _wide(t)]
    extra = sum(len(_table_ids(t, installed, _cols(t)["id"])[0]) for t in tables[1:])
    if extra:
        notes.append("%d package(s) that winget lists as needing explicit targeting were left alone" % extra)
    for ln in (text or "").splitlines():
        if "--include-unknown" in ln:
            m = re.match(r"\s*(\d+)", ln)
            notes.append("%s package(s) whose version winget cannot read were not listed "
                         "(winget upgrade --include-unknown shows them)" % (m.group(1) if m else "some"))
            break
    return notes


def parse_pin_list(text: str, installed=None):
    """-> (set of pinned ids, ok)."""
    first = _first_table(text)
    if first is None:
        low = (text or "").lower()
        return set(), any(p in low for p in NO_PINS_PHRASES)
    idkey = _cols(first)["id"]
    rows, _ = _table_ids(first, installed, idkey)
    return {r[idkey] for r in rows}, True


def parse_export(text: str):
    """winget export JSON -> {PackageIdentifier: Version}, or None when unreadable."""
    try:
        d = json.loads(text)
    except (ValueError, TypeError):
        return None
    out = {}
    for src in d.get("Sources", []) or []:
        for p in src.get("Packages", []) or []:
            pid = p.get("PackageIdentifier")
            if pid:
                out[pid] = p.get("Version", "")
    return out


def read_installed(runner):
    """winget export -> ({id: version} or None, error text)."""
    tmpdir = tempfile.mkdtemp(prefix="upkeep_")
    out = Path(tmpdir) / "export.json"
    try:
        r = run_cmd(runner, [WINGET, "export", "-o", out, "--include-versions",
                             "--accept-source-agreements", "--disable-interactivity"],
                    timeout=300)
        if not out.exists():
            return None, "winget export wrote no file (rc=%s): %s" % (
                r.returncode, " | ".join(tail_lines(r.stdout + "\n" + r.stderr, 2)))
        try:
            text = out.read_text(encoding="utf-8-sig")
        except OSError as e:
            return None, "winget export unreadable: %s" % e
        versions = parse_export(text)
        if versions is None:
            return None, "winget export JSON could not be parsed"
        return versions, ""
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def read_running(runner):
    """tasklist /fo csv /nh -> set of lowercase image names, or None."""
    r = run_cmd(runner, ["tasklist", "/fo", "csv", "/nh"], timeout=60)
    if r.returncode != 0 or not r.stdout.strip():
        return None
    names = set()
    for row in csv.reader(io.StringIO(r.stdout)):
        if row:
            names.add(row[0].strip().lower())
    return names


def decide(pkg_id: str, policy: dict, running, pinned=(), busy_folder=()):
    """Per-package decision: held | needs_admin | deferred | upgrade (with its check_after).
    A package listed under `admin` is never attempted: it needs an administrator shell.
    `busy_folder` names running executables found inside the app's install folder."""
    held = match_any(pkg_id, policy.get("held")) and pkg_id not in (policy.get("held_override") or [])
    if held or pkg_id in pinned:
        return {"action": "held", "reason": "held by policy" if held else "pinned in winget"}
    if match_any(pkg_id, policy.get("admin")):
        return {"action": "needs_admin", "reason": "listed as admin in policy"}
    procs = lookup(policy.get("in_use"), pkg_id) or []
    if procs:
        if running is None:
            return {"action": "deferred", "reason": "deferred: process list unreadable",
                    "processes": list(procs)}
        busy = [p for p in procs if p.lower() in running]
        if busy:
            return {"action": "deferred", "reason": "deferred: in use", "processes": busy}
    if busy_folder:
        return {"action": "deferred", "reason": "deferred: in use", "processes": list(busy_folder)}
    return {"action": "upgrade", "check": lookup(policy.get("check_after"), pkg_id)}


# ------------------------------------------------- registry and process paths ----

_UNINSTALL_KEYS = (
    ("HKLM", r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "machine"),
    ("HKLM", r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "machine"),
    ("HKCU", r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "user"),
)


def read_uninstall() -> dict:
    """Registry uninstall entries, read only: display name ->
    {location, date, publisher, version, scope}. Empty where winreg is missing."""
    try:
        import winreg
    except ImportError:
        return {}
    hives = {"HKLM": winreg.HKEY_LOCAL_MACHINE, "HKCU": winreg.HKEY_CURRENT_USER}
    out = {}
    for hive, sub, scope in _UNINSTALL_KEYS:
        try:
            key = winreg.OpenKey(hives[hive], sub)
        except OSError:
            continue
        i = 0
        while True:
            try:
                name = winreg.EnumKey(key, i)
            except OSError:
                break
            i += 1
            try:
                s = winreg.OpenKey(key, name)
                dn = str(winreg.QueryValueEx(s, "DisplayName")[0])
            except OSError:
                continue
            vals = {}
            for field, value in (("location", "InstallLocation"), ("date", "InstallDate"),
                                 ("publisher", "Publisher"), ("version", "DisplayVersion")):
                try:
                    vals[field] = str(winreg.QueryValueEx(s, value)[0])
                except OSError:
                    vals[field] = ""
            vals["scope"] = scope
            out[dn] = vals
    return out


_PROC_PS = "Get-CimInstance Win32_Process | ForEach-Object { $_.ExecutablePath } | Where-Object { $_ }"


def read_process_paths(runner):
    """Executable paths of the running processes, lowercase; None when unreadable."""
    r = run_cmd(runner, [POWERSHELL, "-NoProfile", "-NonInteractive", "-Command", _PROC_PS], timeout=90)
    if r.returncode != 0:
        return None
    return {ln.strip().lower() for ln in r.stdout.splitlines() if ln.strip()}


def _shared_roots() -> set:
    """Folders many apps share: an install folder equal to one of these says nothing."""
    out = set()
    for var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", "ProgramData", "LOCALAPPDATA",
                "APPDATA", "USERPROFILE", "SystemRoot", "CommonProgramFiles", "CommonProgramFiles(x86)"):
        v = os.environ.get(var)
        if v:
            out.add(os.path.normpath(v).lower())
    la = os.environ.get("LOCALAPPDATA")
    if la:
        out.add(os.path.normpath(os.path.join(la, "Programs")).lower())
        out.add(os.path.normpath(os.path.join(la, "Microsoft", "WindowsApps")).lower())
    return out


def folder_busy(name: str, uninstall: dict, proc_paths) -> list:
    """Names of running executables inside the app's registry install folder.
    Empty when the folder is unknown, shared, or a drive root: this check is partial."""
    loc = ((uninstall or {}).get(name) or {}).get("location", "").strip().strip('"')
    if not loc or not proc_paths:
        return []
    root = os.path.normpath(loc).lower().rstrip("\\")
    if root.count("\\") < 2 or root in _shared_roots():
        return []
    root += "\\"
    return sorted({p.rsplit("\\", 1)[-1] for p in proc_paths if p.startswith(root)})


def load_learned() -> dict:
    """state/learned.json: {"fails": {package id: consecutive failed runs}}."""
    d = read_json(STATE / "learned.json", {})
    d = d if isinstance(d, dict) else {}
    if not isinstance(d.get("fails"), dict):
        d["fails"] = {}
    return d


def classify_outcome(rc, before, after) -> str:
    """Judge by effect: upgraded / not_applied / failed."""
    if after is not None and after != before:
        return "upgraded"
    return "not_applied" if rc == 0 else "failed"


def upgrade_cmd(pkg_id: str) -> list:
    return [WINGET, "upgrade", "--id", pkg_id, "--exact", "--silent",
            "--accept-package-agreements", "--accept-source-agreements",
            "--disable-interactivity"]


def held_command(pkg_id: str) -> str:
    return "winget upgrade --id %s --exact" % pkg_id


def _check_version_cmd(runner, cmd, name):
    r = run_cmd(runner, cmd, timeout=60)
    lines = tail_lines(r.stdout + "\n" + r.stderr, 1)
    first = (r.stdout.strip().splitlines() or [""])[0].strip() or (lines[0] if lines else "")
    if r.returncode == 0 and first:
        return {"name": name, "status": "green", "detail": first}
    return {"name": name, "status": "red",
            "detail": "%s failed (rc=%s): %s" % (" ".join(map(str, cmd)), r.returncode, first)}


def check_node_version(runner, items):
    return _check_version_cmd(runner, ["node", "--version"], "node_version")


def check_git_version(runner, items):
    return _check_version_cmd(runner, ["git", "--version"], "git_version")


CHECKS = {
    "node_version": check_node_version,
    "git_version": check_git_version,
}


def _prev_streaks(prev_winget) -> dict:
    """id -> consecutive deferrals recorded by the previous run (a waiting_close
    entry keeps counting). An entry without a streak counts as one deferral."""
    out = {}
    if isinstance(prev_winget, dict):
        for key in ("deferred", "waiting_close"):
            for x in prev_winget.get(key) or []:
                if isinstance(x, dict) and x.get("id"):
                    try:
                        s = int(x.get("streak", 1))
                    except (TypeError, ValueError):
                        s = 1
                    out[x["id"]] = max(out.get(x["id"], 0), s)
    return out


def _defer(r: dict, base: dict, reason: str, procs, streaks: dict) -> None:
    """Record a deferral. From DEFER_LIMIT consecutive days on, the package
    waits for the user (waiting_close) instead of being deferred again."""
    procs = list(procs or [])
    streak = streaks.get(base["id"], 0) + 1
    rec = dict(base, reason=reason, processes=procs, streak=streak)
    if streak >= DEFER_LIMIT:
        rec["command"] = "%s%s" % ("close %s, then run: " % ", ".join(procs) if procs else "run: ",
                                   held_command(base["id"]))
        r["waiting_close"].append(rec)
    else:
        r["deferred"].append(rec)


def _is_admin_failure(res) -> bool:
    """winget's own evidence that the installer needed elevation: exit 1603, or
    its installer-failure code 2316632112 together with the text 1603."""
    if res.returncode == 1603:
        return True
    text = "%s\n%s" % (res.stdout, res.stderr)
    return "2316632112" in ("%s %s" % (res.returncode, text)) and "1603" in text


def _reconcile(prev_winget, before, policy) -> list:
    """Packages the previous run recorded as failed or unconfirmed whose installed
    version moved since: the installer finished after the run."""
    done, seen_ids = [], set()
    if not isinstance(prev_winget, dict) or before is None:
        return done
    for key in ("failed", "unconfirmed"):
        for x in prev_winget.get(key) or []:
            pid, frm = x.get("id"), x.get("from")
            if not pid or pid in seen_ids or pid not in before or frm is None:
                continue
            if before[pid] != frm:
                seen_ids.add(pid)
                done.append({"id": pid, "name": x.get("name", ""), "from": frm, "to": before[pid]})
    return done


def winget_pass(runner=subprocess.run, policy=None, dry_run=False, items=None,
                prev_winget=None, clock=time.monotonic, only=None, learned=None,
                uninstall=None) -> dict:
    """The whole winget pass. Never raises; returns a result dict with
    status green|yellow|red and the per-package outcomes.

    `prev_winget` is the previous run's winget dict (reconciles late completions
    and counts deferral streaks; None keeps a pass isolated). `clock` measures
    the overall time budget. `only` restricts the pass to the named ids and
    lets them past the selection and the failure limit. `learned` carries the
    consecutive-failure counts; the updated counts come back in r["learned"].
    `uninstall` is the registry map for the folder check (read when None)."""
    r = {"status": "green", "dry_run": bool(dry_run), "errors": [], "warnings": [], "notes": [],
         "available": [], "upgraded": [], "failed": [], "not_applied": [],
         "deferred": [], "held": [], "needs_admin": [], "unconfirmed": [],
         "waiting_close": [], "waiting_fail": [], "not_selected": [],
         "completed_late": [], "would_upgrade": [], "checks": [], "raw_head": []}
    policy = policy if policy is not None else load_policy()
    if policy is None:
        r["errors"].append("policy.json is missing or unreadable: %s" % POLICY)
        r["status"] = "red"
        return r
    only = set(only or [])
    fails = {k: int(v) for k, v in ((learned or {}).get("fails") or {}).items()
             if isinstance(v, int) and not isinstance(v, bool)}
    limit = int(policy.get("fail_limit") or FAIL_LIMIT)
    timeout = int(policy.get("upgrade_timeout_s") or WINGET_UPGRADE_TIMEOUT)

    before, err = read_installed(runner)
    if before is None:
        r["errors"].append(err)

    r["completed_late"] = _reconcile(prev_winget, before, policy)
    for x in r["completed_late"]:
        fails.pop(x["id"], None)
    streaks = _prev_streaks(prev_winget)
    if dry_run:   # a pass that installs nothing carries a streak forward, it does not lengthen it
        streaks = {k: v - 1 for k, v in streaks.items()}

    ul = run_cmd(runner, [WINGET, "upgrade", "--include-pinned", "--accept-source-agreements",
                          "--disable-interactivity"], timeout=300)
    pkgs, warns, ok = parse_upgrade_list(ul.stdout, before, ul.returncode)
    r["warnings"].extend(warns)
    if not ok:
        r["raw_head"] = (ul.stdout + "\n" + ul.stderr).strip().splitlines()[:8]
        r["errors"].append("winget upgrade list could not be parsed (rc=%s)" % ul.returncode)
        pkgs = []
    else:
        r["notes"].extend(parse_upgrade_notes(ul.stdout, before))
    r["available"] = pkgs

    pl = run_cmd(runner, [WINGET, "pin", "list", "--accept-source-agreements",
                          "--disable-interactivity"], timeout=120)
    pinned, pins_ok = parse_pin_list(pl.stdout, before)
    if not pins_ok:
        r["warnings"].append("winget pin list could not be parsed")

    if before is not None and pins_ok:
        missing = missing_pins(policy, before, pinned)
        if missing:   # a note: held packages are listed anyway (--include-pinned), so this alone is not yellow
            r["notes"].append("held but not pinned in winget: %s (upkeep pins --apply)"
                              % ", ".join(missing))

    if only:
        scope = [p for p in pkgs if p["id"] in only]
        absent = sorted(only - {p["id"] for p in pkgs})
        if absent:
            r["notes"].append("named but not in winget's upgrade list: %s" % ", ".join(absent))
    else:
        scope = []
        for p in pkgs:
            if is_selected(p["id"], policy):
                scope.append(p)
            else:
                r["not_selected"].append({"id": p["id"], "name": p["name"], "from": p["version"],
                                          "available": p["available"],
                                          "command": "upkeep select --add %s" % p["id"]})

    needs_tasklist = any(lookup(policy["in_use"], p["id"]) for p in scope)
    running = read_running(runner) if needs_tasklist else set()
    proc_paths = None
    if scope:
        uninstall = uninstall if uninstall is not None else read_uninstall()
        if any((uninstall.get(p["name"]) or {}).get("location") for p in scope):
            proc_paths = read_process_paths(runner)
            if proc_paths is None:
                r["notes"].append("running-process paths unreadable: the install-folder check was skipped")

    todo = []
    for p in scope:
        d = decide(p["id"], policy, running, pinned, folder_busy(p["name"], uninstall, proc_paths))
        base = {"id": p["id"], "name": p["name"], "from": p["version"], "available": p["available"]}
        if d["action"] == "held":
            r["held"].append(dict(base, reason=d["reason"], command=held_command(p["id"])))
        elif d["action"] == "needs_admin":
            r["needs_admin"].append(dict(base, reason=d["reason"], command=held_command(p["id"]), note=ADMIN_NOTE))
        elif d["action"] == "deferred":
            _defer(r, base, d["reason"], d["processes"], streaks)
        elif fails.get(p["id"], 0) >= limit and p["id"] not in only:
            r["waiting_fail"].append(dict(base, fails=fails[p["id"]], command=held_command(p["id"]),
                                          reason="failed %d runs in a row" % fails[p["id"]]))
        else:
            todo.append((p, d))
            r["would_upgrade"].append(dict(base, check=d["check"]))

    r["learned"] = {"fails": fails}
    if dry_run:
        _finish_status(r)
        return r
    if r["errors"] or before is None:
        if before is None and not r["errors"]:
            r["errors"].append("installed versions unreadable")
        _finish_status(r)
        return r

    attempts, started_at = [], None
    for p, d in todo:
        t = clock()
        if started_at is None:
            started_at = t
        elif t - started_at >= WINGET_TIME_BUDGET:
            base = {"id": p["id"], "name": p["name"], "from": p["version"], "available": p["available"]}
            _defer(r, base, "deferred: time budget", [], streaks)
            continue
        res = run_cmd(runner, upgrade_cmd(p["id"]), timeout=timeout)   # one installer at a time
        attempts.append((p, d, res))

    after = None
    if attempts:
        after, err = read_installed(runner)
        if after is None:
            r["errors"].append("could not verify upgrades: " + err)

    needed_checks = []
    for x in r["completed_late"]:   # an installer that finished after the last run still gets its check
        chk = lookup(policy.get("check_after"), x["id"])
        if chk and chk not in needed_checks:
            needed_checks.append(chk)
    for p, d, res in attempts:
        v_before = before.get(p["id"], p["version"])
        v_after = after.get(p["id"]) if after is not None else None
        if after is None:
            continue
        outcome = classify_outcome(res.returncode, v_before, v_after)
        rec = {"id": p["id"], "name": p["name"], "from": v_before, "to": v_after,
               "rc": res.returncode, "tail": tail_lines(res.stdout + "\n" + res.stderr, 5)}
        if outcome == "upgraded":
            fails.pop(p["id"], None)
            r["upgraded"].append(rec)
            if d["check"] and d["check"] not in needed_checks:
                needed_checks.append(d["check"])
            continue
        if outcome != "not_applied":
            fails[p["id"]] = fails.get(p["id"], 0) + 1
            rec["fails"] = fails[p["id"]]
        if res.returncode == -998:   # timed out, version unchanged: unknown until a later pass reads it
            r["unconfirmed"].append(rec)
        elif outcome == "failed" and _is_admin_failure(res):
            r["needs_admin"].append(dict(rec, available=p["available"], reason="installer failed (1603)",
                                         command=held_command(p["id"]), note=ADMIN_NOTE))
        elif outcome == "not_applied":
            r["not_applied"].append(rec)
        else:
            r["failed"].append(rec)

    for name in needed_checks:
        fn = CHECKS.get(name)
        if fn is None:
            r["checks"].append({"name": name, "status": "red", "detail": "unknown check name"})
            continue
        try:
            r["checks"].append(fn(runner, items if items is not None else load_items()))
        except Exception as e:  # noqa: BLE001
            r["checks"].append({"name": name, "status": "red", "detail": "check crashed: %s" % e})

    _finish_status(r)
    return r


def _finish_status(r: dict) -> None:
    if r["errors"] or r["failed"] or any(c["status"] == "red" for c in r["checks"]):
        r["status"] = "red"
    elif (r["held"] or r["deferred"] or r["not_applied"] or r["warnings"] or r["needs_admin"]
          or r["unconfirmed"] or r["waiting_close"] or r["waiting_fail"]
          or (r["dry_run"] and r["would_upgrade"])
          or any(c["status"] == "unknown" for c in r["checks"])):
        # notes, completed_late and not_selected are information: they never colour the pass
        r["status"] = "yellow"
    else:
        r["status"] = "green"


def missing_pins(policy: dict, installed: dict, pinned) -> list:
    """Held package ids (exact ids, plus wildcard patterns expanded against the
    installed set) that have no winget pin."""
    ids = set()
    for pat in policy.get("held", []):
        if any(ch in pat for ch in "*?["):
            ids.update(i for i in installed if fnmatch.fnmatchcase(i, pat))
        elif pat in installed:
            ids.add(pat)
    return sorted(i for i in ids if i not in set(pinned))


def winget_summary(r: dict) -> str:
    bits = []
    if r["errors"]:
        bits.append("errors: " + "; ".join(r["errors"]))
    if r["failed"]:
        bits.append("failed: " + ", ".join(x["id"] for x in r["failed"]))
    if r["upgraded"]:
        bits.append("installed: " + ", ".join("%s %s->%s" % (x["id"], x["from"], x["to"]) for x in r["upgraded"]))
    if r["not_applied"]:
        bits.append("not applied: " + ", ".join(x["id"] for x in r["not_applied"]))
    if r.get("unconfirmed"):
        bits.append("unconfirmed (timed out): " + ", ".join(x["id"] for x in r["unconfirmed"]))
    if r.get("completed_late"):
        bits.append("completed after the run: " + ", ".join(
            "%s %s->%s" % (x["id"], x["from"], x["to"]) for x in r["completed_late"]))
    if r.get("would_upgrade") and r["dry_run"]:
        bits.append("ready to install: " + ", ".join(x["id"] for x in r["would_upgrade"]))
    if r["deferred"]:
        bits.append("deferred: " + ", ".join(x["id"] for x in r["deferred"]))
    if r["held"]:
        bits.append("waiting for you: " + ", ".join(x["id"] for x in r["held"]))
    if r.get("needs_admin"):
        bits.append("waiting for you (admin): " + ", ".join(x["id"] for x in r["needs_admin"]))
    if r.get("waiting_close"):
        bits.append("waiting for you (close the program first): " + ", ".join(x["id"] for x in r["waiting_close"]))
    if r.get("waiting_fail"):
        bits.append("waiting for you (failed twice or more): " + ", ".join(x["id"] for x in r["waiting_fail"]))
    if r.get("not_selected"):
        bits.append("not selected: " + ", ".join(x["id"] for x in r["not_selected"]))
    for w in r["warnings"]:
        bits.append("warning: " + w)
    for n in r.get("notes", []):
        bits.append("note: " + n)
    return "; ".join(bits) or "nothing to do"


# --------------------------------------------------------------- claude pass ----

def _claude_run(runner, args, timeout):
    """`claude <args>`, falling back to the native install's own path."""
    r = run_cmd(runner, ["claude"] + list(args), timeout=timeout)
    if r.returncode == -997:
        fb = Path(os.environ.get("USERPROFILE") or Path.home()) / ".local" / "bin" / "claude.exe"
        r = run_cmd(runner, [fb] + list(args), timeout=timeout)
    return r


def claude_version(runner):
    r = _claude_run(runner, ["--version"], 60)
    m = re.search(r"(\d+\.\d+\.\d+)", r.stdout + "\n" + r.stderr)
    return m.group(1) if m else None


def plugin_versions(path=None) -> dict:
    """name@marketplace -> 'version/commit' from installed_plugins.json."""
    data = read_json(xpath(path or CLAUDE_PLUGINS_FILE), {})
    plugins = data.get("plugins", {}) if isinstance(data, dict) else {}
    out = {}
    for name, entries in (plugins.items() if isinstance(plugins, dict) else []):
        if isinstance(entries, list) and entries and isinstance(entries[0], dict):
            e = entries[0]
            out[name] = "%s/%s" % (e.get("version"), (e.get("gitCommitSha") or "")[:12])
    return out


def _outcome(rc, before, after) -> str:
    if before is not None and after is not None and after != before:
        return "upgraded"
    return "current" if rc == 0 else "failed"


def claude_pass(runner=subprocess.run, dry_run=False, now=None, plugins_file=None) -> dict:
    """claude update, then marketplaces, then each installed plugin. Never raises.
    A failed step is yellow (it waits for the user with the command); nothing is red."""
    now = now or now_local()
    out = {"schema": "upkeep-claude/1", "at": now.isoformat(timespec="seconds"),
           "dry_run": bool(dry_run), "status": "green", "cli": {}, "marketplaces": {}, "plugins": []}
    before = claude_version(runner)
    if dry_run:
        out["cli"] = {"before": before, "after": before, "outcome": "skipped (dry run)"}
    else:
        r = _claude_run(runner, ["update"], CLAUDE_UPDATE_TIMEOUT)
        after = claude_version(runner)
        out["cli"] = {"before": before, "after": after, "rc": r.returncode,
                      "outcome": _outcome(r.returncode, before, after),
                      "tail": tail_lines(r.stdout + "\n" + r.stderr, 4), "command": "claude update"}
    pv = plugin_versions(plugins_file)
    if dry_run:
        out["marketplaces"] = {"outcome": "skipped (dry run)"}
        out["plugins"] = [{"name": n, "before": v, "after": v, "outcome": "skipped (dry run)"}
                          for n, v in sorted(pv.items())]
    else:
        r = _claude_run(runner, ["plugin", "marketplace", "update"], CLAUDE_PLUGIN_TIMEOUT)
        out["marketplaces"] = {"rc": r.returncode, "outcome": "ok" if r.returncode == 0 else "failed",
                               "tail": tail_lines(r.stdout + "\n" + r.stderr, 3),
                               "command": "claude plugin marketplace update"}
        for name in sorted(pv):
            r = _claude_run(runner, ["plugin", "update", name], CLAUDE_PLUGIN_TIMEOUT)
            after = plugin_versions(plugins_file).get(name)
            out["plugins"].append({"name": name, "before": pv[name], "after": after, "rc": r.returncode,
                                   "outcome": _outcome(r.returncode, pv[name], after),
                                   "tail": tail_lines(r.stdout + "\n" + r.stderr, 3),
                                   "command": "claude plugin update %s" % name})
    failed = (out["cli"].get("outcome") == "failed" or out["marketplaces"].get("outcome") == "failed"
              or any(p["outcome"] == "failed" for p in out["plugins"]))
    out["status"] = "yellow" if failed else "green"
    return out


def claude_summary(c: dict) -> list:
    lines = []
    cli = c.get("cli") or {}
    o = cli.get("outcome")
    if o == "upgraded":
        lines.append("Claude Code updated: %s -> %s" % (cli.get("before"), cli.get("after")))
    elif o == "failed":
        lines.append("FAILED: claude update rc=%s; last output: %s" % (cli.get("rc"), " | ".join(cli.get("tail") or [])))
    else:
        lines.append("Claude Code %s: %s" % (cli.get("after") or cli.get("before") or "?", o))
    m = c.get("marketplaces") or {}
    if m.get("outcome") == "failed":
        lines.append("FAILED: plugin marketplace update rc=%s; last output: %s" % (m.get("rc"), " | ".join(m.get("tail") or [])))
    for p in c.get("plugins") or []:
        if p["outcome"] == "upgraded":
            lines.append("Plugin updated: %s %s -> %s" % (p["name"], p["before"], p["after"]))
        elif p["outcome"] == "failed":
            lines.append("FAILED: plugin %s rc=%s; last output: %s" % (p["name"], p.get("rc"), " | ".join(p.get("tail") or [])))
    n_cur = sum(1 for p in c.get("plugins") or [] if p["outcome"] == "current")
    if n_cur:
        lines.append("%d plugin(s) already current" % n_cur)
    n_skip = sum(1 for p in c.get("plugins") or [] if p["outcome"] == "skipped (dry run)")
    if n_skip:
        lines.append("%d plugin(s) installed, not updated (dry run)" % n_skip)
    return lines


# --------------------------------------------------------------- item checks ----

def make_ctx(runner=subprocess.run, seen=None, now=None, winget=None, prev_last=None, claude=None) -> dict:
    return {"runner": runner, "seen": seen or {}, "now": now or now_local(),
            "winget": winget, "prev_last": prev_last, "claude": claude}


def _recent_claude_pass(ctx, max_days):
    """The Claude pass of this run, or of a recent run, when it really ran."""
    c = ctx.get("claude")
    if not isinstance(c, dict) or c.get("dry_run"):
        return None
    at = parse_datetime(c.get("at"))
    if at is None or (ctx["now"] - at).total_seconds() > max(1, max_days) * 86400:
        return None
    return c


def res(item, status, value=None, detail="", action="", **extra) -> dict:
    out = {"id": item.get("id"), "label": item.get("label") or item.get("id"),
           "status": status, "value": value, "detail": detail,
           "action": action or item.get("action", "")}
    out.update(extra)
    return out


def _age_status(item, age_days, what):
    yd, rd = item.get("yellow_days"), item.get("red_days")
    if rd is not None and age_days > rd:
        return "red", "%s is %d days old (red after %s)" % (what, age_days, rd)
    if yd is not None and age_days > yd:
        return "yellow", "%s is %d days old (yellow after %s)" % (what, age_days, yd)
    return "green", "%s is %d days old" % (what, age_days)


def check_command_version(item, ctx):
    cmd = [xp(c) for c in item["cmd"]]
    r = run_cmd(ctx["runner"], cmd, timeout=60)
    if r.returncode in (-997, -998, -999):
        return res(item, "unknown", detail="cannot run %s: %s" % (cmd[0], r.stderr.strip()))
    text = r.stdout + "\n" + r.stderr
    m = re.search(item.get("regex", r"(\S+)"), text, re.M)
    if not m:
        return res(item, "unknown", detail="no version found in output (rc=%s): %s" % (
            r.returncode, " | ".join(tail_lines(text, 2))))
    value = m.group(1) if m.groups() else m.group(0)
    return res(item, "green", value=value, detail="installed %s" % value)


def check_file_age(item, ctx):
    now = ctx["now"]
    paths = []
    for pat in item.get("paths") or [item["path"]]:
        p = xpath(pat)
        if any(ch in p.name for ch in "*?["):
            paths.extend(p.parent.glob(p.name) if p.parent.exists() else [])
        elif p.exists():
            paths.append(p)
    if not paths:
        return res(item, "unknown", detail="file not found: %s" % ", ".join(xp(x) for x in (item.get("paths") or [item["path"]])))
    stamp, src = None, ""
    key = item.get("json_key")
    if key:
        data = read_json(paths[0])
        if isinstance(data, dict) and key in data:
            d = parse_date(data[key])
            stamp = datetime(d.year, d.month, d.day) if d else None
            src = "%s %s" % (paths[0].name, key)
    key2 = item.get("json_newest_key")
    if key2 and stamp is None:
        data = read_json(paths[0])
        found = []

        def walk(o):
            if isinstance(o, dict):
                for k, v in o.items():
                    if k == key2:
                        dt = parse_datetime(v)
                        if dt:
                            found.append(dt)
                    else:
                        walk(v)
            elif isinstance(o, list):
                for v in o:
                    walk(v)
        walk(data)
        if found:
            stamp, src = max(found), "%s newest %s" % (paths[0].name, key2)
    if stamp is None:
        m = max(p.stat().st_mtime for p in paths)
        stamp = datetime.fromtimestamp(m)
        src = "mtime of %s" % (max(paths, key=lambda p: p.stat().st_mtime).name)
    age = max(0, int((now - stamp).total_seconds() // 86400))
    status, detail = _age_status(item, age, src)
    return res(item, status, value=stamp.date().isoformat(), detail=detail)


def check_text_date(item, ctx):
    p = xpath(item["path"])
    try:
        text = p.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return res(item, "unknown", detail="file not found: %s" % p)
    m = re.search(item["regex"], text)
    d = parse_date(m.group(1)) if m else None
    if d is None:
        return res(item, "unknown", detail="date pattern not found in %s" % p.name)
    age = (ctx["now"].date() - d).days
    status, detail = _age_status(item, age, "%s date %s" % (p.name, d.isoformat()))
    return res(item, status, value=d.isoformat(), detail=detail)


def check_text_date_due(item, ctx):
    p = xpath(item["path"])
    try:
        text = p.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return res(item, "unknown", detail="file not found: %s" % p)
    m = re.search(item["regex"], text)
    d = parse_date(m.group(1)) if m else None
    if d is None:
        return res(item, "unknown", detail="due-date pattern not found in %s" % p.name)
    today = ctx["now"].date()
    if d < today:
        return res(item, "red", value=d.isoformat(),
                   detail="review was due %s (%d days ago)" % (d.isoformat(), (today - d).days))
    if d <= today + timedelta(days=7):
        return res(item, "yellow", value=d.isoformat(),
                   detail="review due %s (in %d days)" % (d.isoformat(), (d - today).days))
    return res(item, "green", value=d.isoformat(), detail="review due %s" % d.isoformat())


def check_json_flag(item, ctx):
    files = item.get("paths") or [item["path"]]
    bad, missing = [], []
    for f in files:
        p = xpath(f)
        data = read_json(p)
        if not isinstance(data, dict):
            missing.append(p.name)
        elif data.get(item["key"]) != item["expect"]:
            bad.append("%s: %s=%r" % (p.name, item["key"], data.get(item["key"])))
    if bad:
        return res(item, "red", detail="; ".join(bad))
    if missing:
        return res(item, "unknown", detail="unreadable: %s" % ", ".join(missing))
    return res(item, "green", value=str(item["expect"]), detail="%s == %s in %d file(s)" % (item["key"], item["expect"], len(files)))


def check_dir_latest(item, ctx):
    d = xpath(item["path"])
    if not d.is_dir():
        return res(item, "unknown", detail="folder not found: %s" % d)
    names = [p.name for p in d.iterdir() if p.is_dir() and p.name.startswith(item.get("prefix", ""))]
    if not names:
        return res(item, "unknown", detail="no folder starting with %r in %s" % (item.get("prefix", ""), d))

    def vkey(n):
        m = re.search(r"(\d+(?:\.\d+)+)", n[len(item.get("prefix", "")):])
        return tuple(int(x) for x in m.group(1).split(".")) if m else (-1,)
    best = max(names, key=vkey)
    m = re.search(r"(\d+(?:\.\d+)+)", best[len(item.get("prefix", "")):])
    value = m.group(1) if m else best
    return res(item, "green", value=value, detail="newest folder %s" % best)


def check_dir_list(item, ctx):
    d = xpath(item["path"])
    if not d.is_dir():
        return res(item, "unknown", detail="folder not found: %s" % d)
    names = sorted(p.name for p in d.iterdir() if p.is_dir() and p.name.startswith(item.get("prefix", "")))
    if not names:
        return res(item, "unknown", detail="no folders in %s" % d)
    return res(item, "green", value=", ".join(names), detail="%d folders" % len(names))


def check_claude_code(item, ctx):
    r = run_cmd(ctx["runner"], ["claude", "--version"], timeout=60)
    if r.returncode == -997:
        fb = Path(os.environ.get("USERPROFILE") or Path.home()) / ".local" / "bin" / "claude.exe"
        r = run_cmd(ctx["runner"], [fb, "--version"], timeout=60)
    m = re.search(r"(\d+\.\d+\.\d+)", r.stdout + "\n" + r.stderr)
    if not m:
        return res(item, "unknown", detail="claude --version gave no version (rc=%s): %s" % (
            r.returncode, r.stderr.strip()[:120]))
    value = m.group(1)
    c = _recent_claude_pass(ctx, item.get("stale_days", 14))
    if c:
        cli = c.get("cli") or {}
        o = cli.get("outcome")
        if o == "failed":
            return res(item, "yellow", value=value, waiting=True,
                       detail="claude update failed (rc=%s, %s): %s" % (cli.get("rc"), c.get("at"), " | ".join(cli.get("tail") or [])),
                       action="claude update")
        if o == "upgraded":
            return res(item, "green", value=value,
                       detail="updated by upkeep: %s -> %s (%s)" % (cli.get("before"), cli.get("after"), c.get("at")))
        if o == "current":
            return res(item, "green", value=value,
                       detail="Claude Code %s; claude update found nothing newer (%s)" % (value, c.get("at")))
    prev = ctx["seen"].get(item["id"])
    now = ctx["now"]
    if prev and prev.get("value") == value:
        since = parse_datetime(prev.get("since")) or now
        days = int((now - since).total_seconds() // 86400)
        if days > item.get("stale_days", 14):
            return res(item, "yellow", value=value, waiting=True,
                       detail="Claude Code %s unchanged for %d days (no network lookup is made)" % (value, days),
                       action=item.get("action", "claude update"))
        return res(item, "green", value=value, detail="Claude Code %s (%d days on this version)" % (value, days))
    if prev:
        return res(item, "green", value=value, detail="changed: %s -> %s" % (prev.get("value"), value))
    return res(item, "green", value=value, detail="Claude Code %s (first record)" % value)


def check_claude_plugins(item, ctx):
    """The Claude pass decides when it ran recently; the file age is the backstop."""
    c = _recent_claude_pass(ctx, item.get("yellow_days", 30))
    if not c:
        return check_file_age(item, ctx)
    m = c.get("marketplaces") or {}
    plugins = c.get("plugins") or []
    failed = [p for p in plugins if p["outcome"] == "failed"]
    if m.get("outcome") == "failed" or failed:
        parts = []
        if m.get("outcome") == "failed":
            parts.append("marketplace update rc=%s: %s" % (m.get("rc"), " | ".join(m.get("tail") or [])))
        parts += ["%s rc=%s: %s" % (p["name"], p.get("rc"), " | ".join(p.get("tail") or [])) for p in failed]
        cmds = ([m.get("command")] if m.get("outcome") == "failed" else []) + [p["command"] for p in failed]
        return res(item, "yellow", waiting=True, detail="plugin update failed (%s): %s" % (c.get("at"), "; ".join(parts)),
                   action="; ".join(x for x in cmds if x))
    up = [p for p in plugins if p["outcome"] == "upgraded"]
    if up:
        return res(item, "green", value=str(len(plugins)),
                   detail="updated by upkeep (%s): %s; restart Claude Code sessions to load them" % (
                       c.get("at"), ", ".join("%s %s -> %s" % (p["name"], p["before"], p["after"]) for p in up)))
    return res(item, "green", value=str(len(plugins)),
               detail="%d plugin(s) current after the marketplace refresh (%s)" % (len(plugins), c.get("at")))


def _vkey(s):
    m = re.search(r"(\d+(?:\.\d+)+)", s or "")
    return tuple(int(x) for x in m.group(1).split(".")) if m else None


def check_claude_extension_lag(item, ctx):
    """Watch only: VS Code updates the extension itself. Yellow when the newest
    installed extension has stayed behind the CLI for more than `lag_days`."""
    d = xpath(item.get("path", "%USERPROFILE%/.vscode/extensions"))
    prefix = item.get("prefix", "anthropic.claude-code-")
    if not d.is_dir():
        return res(item, "unknown", detail="folder not found: %s" % d)
    exts = [p.name[len(prefix):] for p in d.iterdir() if p.is_dir() and p.name.startswith(prefix)]
    exts = [e for e in exts if _vkey(e)]
    if not exts:
        return res(item, "unknown", detail="no %s* folder in %s" % (prefix, d))
    ext = ".".join(str(x) for x in _vkey(max(exts, key=_vkey)))
    c = _recent_claude_pass(ctx, 2)
    cli = ((c or {}).get("cli") or {}).get("after") or claude_version(ctx["runner"])
    if not cli:
        return res(item, "unknown", detail="extension %s; CLI version not readable" % ext)
    if _vkey(ext) >= _vkey(cli):
        return res(item, "green", value="in step %s" % ext, detail="extension %s, CLI %s" % (ext, cli))
    value = "%s<%s" % (ext, cli)
    prev = ctx["seen"].get(item["id"])
    since = parse_datetime(prev.get("since")) if prev and prev.get("value") == value else None
    days = int((ctx["now"] - since).total_seconds() // 86400) if since else 0
    lag = item.get("lag_days", 2)
    if days > lag:
        return res(item, "yellow", value=value, waiting=True,
                   detail="VS Code Claude extension %s behind the CLI %s for %d days (yellow after %d); VS Code's own extension update has not caught up" % (
                       ext, cli, days, lag))
    return res(item, "green", value=value,
               detail="extension %s behind the CLI %s for %d day(s) (yellow after %d)" % (ext, cli, days, lag))


def _norm_col(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def parse_schtasks(text: str):
    """schtasks /query /fo CSV /v -> list of {name, last_result(int|None), last_run,
    state, task_state}. Header names are matched accent-insensitively and tolerate
    a mangled accent, in Portuguese or English. Repeated header rows are skipped.
    `state` is the Status column, `task_state` the scheduled-task-state column
    ('Estado de tarefa agendada'); either may say disabled."""
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return []
    header = rows[0]
    idx = {"name": None, "result": None, "run": None, "state": None, "task_state": None}
    for i, h in enumerate(header):
        n = _norm_col(h)
        if idx["name"] is None and n in ("nomedatarefa", "taskname"):
            idx["name"] = i
        elif idx["result"] is None and ("resultado" in n or "lastresult" in n):
            idx["result"] = i
        elif idx["run"] is None and ("timaexecu" in n or "lastrun" in n):
            idx["run"] = i
        elif idx["state"] is None and n == "status":
            idx["state"] = i
        elif idx["task_state"] is None and n in ("estadodetarefaagendada", "scheduledtaskstate"):
            idx["task_state"] = i
    if idx["name"] is None or idx["result"] is None:
        return None
    out = []
    for r in rows[1:]:
        if len(r) <= max(i for i in idx.values() if i is not None) or r == header or r[idx["name"]] == header[idx["name"]]:
            continue
        try:
            lr = int(r[idx["result"]].strip())
        except ValueError:
            lr = None
        out.append({"name": r[idx["name"]].lstrip("\\"), "last_result": lr,
                    "last_run": r[idx["run"]] if idx["run"] is not None else "",
                    "state": r[idx["state"]].strip() if idx["state"] is not None else "",
                    "task_state": r[idx["task_state"]].strip() if idx["task_state"] is not None else ""})
    return out


_DISABLED = ("desabilitado", "desativado", "disabled")


def _task_disabled(t) -> bool:
    return any(_norm_col(t.get(k, "")) in _DISABLED for k in ("state", "task_state"))


def _task_rank(t, item):
    lr = t["last_result"]
    if lr == 0:
        return 0
    return 1 if lr in set(item.get("soft_results", [267009, 267011])) else 2


def check_scheduled_tasks(item, ctx):
    r = run_cmd(ctx["runner"], ["schtasks", "/query", "/fo", "CSV", "/v"], timeout=120)
    tasks = parse_schtasks(r.stdout) if r.stdout.strip() else None
    if tasks is None:
        return res(item, "unknown", detail="schtasks output not understood (rc=%s)" % r.returncode)
    keys = [k.lower() for k in item.get("name_contains", [TASK_NAME])]
    worst = {}  # one row per trigger: keep one entry per task, the worst result
    for t in tasks:
        if _task_disabled(t) or not any(k in t["name"].lower() for k in keys):
            continue   # a disabled row is dropped before the per-task aggregation
        old = worst.get(t["name"])
        if old is None or _task_rank(t, item) > _task_rank(old, item):
            worst[t["name"]] = t
    watched = list(worst.values())
    soft = set(item.get("soft_results", [267009, 267011]))
    configured = item.get("task_results") or {}
    red, yellow, soft_yellow, accepted = [], [], [], 0
    alert_rows = []
    for t in watched:
        lr = t["last_result"]
        mapped = (configured.get(t["name"]) or {}).get(str(lr)) if lr is not None else None
        if isinstance(mapped, list) and len(mapped) in (2, 3) and mapped[0] in ("red", "yellow", "green"):
            why = mapped[1]
            if mapped[0] == "red":
                red.append((t, why))
            elif mapped[0] == "yellow":
                yellow.append((t, why))
            else:
                accepted += 1
        elif lr is None or (lr != 0 and lr not in soft):
            red.append((t, ""))
        elif lr in soft:
            soft_yellow.append(t)
    fmt = lambda t, why="": "%s=%s%s" % (t["name"], t["last_result"], " (%s)" % why if why else "")  # noqa: E731
    extra = {"alerts": alert_rows} if alert_rows else {}
    if red:
        return res(item, "red", value="%d red of %d" % (len(red), len(watched)),
                   detail="last result not 0: " + ", ".join(fmt(t, why) for t, why in red[:8]), **extra)
    if yellow or soft_yellow:
        parts = []
        if soft_yellow:
            parts.append("running or never run: " + ", ".join(fmt(t) for t in soft_yellow[:8]))
        parts += ["%s: %s" % (fmt(t), why) for t, why in yellow[:8]]
        return res(item, "yellow", value="%d yellow of %d" % (len(yellow) + len(soft_yellow), len(watched)),
                   detail="; ".join(parts), **extra)
    return res(item, "green", value="%d tasks" % len(watched),
               detail="%d watched tasks, all last results 0%s" % (len(watched), " or accepted" if accepted else ""),
               **extra)


def check_winget_pass(item, ctx):
    w = ctx.get("winget")
    if w is None:
        prev = (ctx.get("prev_last") or {}).get("winget_item")
        if prev:
            return res(item, prev.get("status", "unknown"), value=prev.get("value"),
                       detail="from the last daily run: " + (prev.get("detail") or ""))
        return res(item, "unknown", detail="the winget pass has not run yet")
    return res(item, w["status"], value=w["status"], detail=winget_summary(w),
               action="upkeep pins --apply" if any("not pinned" in x for x in w.get("notes", [])) else "")


def apply_change(item, result, prev):
    """Values seen on the previous run: a change turns a green item yellow. The
    check's own finding stays in the detail, with the change appended, and its
    own action stays when it set one."""
    tracked = item.get("track", item.get("type") == "command_version")
    if not tracked or not prev or result["status"] != "green":
        return result
    if prev.get("value") != result["value"]:
        change = "changed: %s -> %s" % (prev.get("value"), result["value"])
        result["status"] = "yellow"
        result["detail"] = "%s (%s)" % (result["detail"], change) if result.get("detail") else change
        result["action"] = result.get("action") or item.get("on_change") or "review the change"
    return result


def run_items(items, ctx) -> list:
    out = []
    for item in items:
        fn = CHECK_TYPES.get(item.get("type"))
        if fn is None:
            out.append(res(item, "unknown", detail="unknown check type %r" % item.get("type")))
            continue
        try:
            result = fn(item, ctx)
        except Exception as e:  # noqa: BLE001 - one broken check must not hide the rest
            result = res(item, "unknown", detail="check crashed: %s: %s" % (type(e).__name__, e))
        out.append(apply_change(item, result, ctx["seen"].get(item.get("id"))))
    return out


# -------------------------------------------------------------- state files ----

def last_path() -> Path:
    return STATE / "last.json"


def previous_toast_item(prev_last):
    t = (prev_last or {}).get("toast") or {}
    if t.get("attempted") and not t.get("ok"):
        return res({"id": "previous_toast", "label": "Previous toast"}, "yellow",
                   detail="the previous run's toast failed: %s" % t.get("detail", ""),
                   action="check Windows notification settings (Focus Assist / Do Not Disturb)")
    return None


def update_seen(seen: dict, results, now) -> dict:
    new = {}
    stamp = now.isoformat(timespec="seconds")
    for r in results:
        if r.get("value") is None or r["id"] in ("previous_toast", "upkeep_stale"):
            if r["id"] in seen:
                new[r["id"]] = seen[r["id"]]
            continue
        old = seen.get(r["id"])
        new[r["id"]] = {"value": r["value"],
                        "since": old["since"] if old and old.get("value") == r["value"] and old.get("since") else stamp}
    return new


def counts(results) -> dict:
    c = {"red": 0, "yellow": 0, "green": 0, "unknown": 0}
    for r in results:
        c[r["status"]] += 1
    return c


# ------------------------------------------------------------ report + toast ----

def render_report(record: dict) -> str:
    c = record["counts"]
    lines = ["# Upkeep report %s" % record["date"], "",
             "Run finished %s. %d red, %d yellow, %d unknown, %d green." % (
                 record["finished"], c["red"], c["yellow"], c["unknown"], c["green"]), ""]
    if record.get("run"):
        lines += ["Mode: %s. This run (%s) %s." % (
            record.get("mode"), record["run"],
            "listed the updates and installed nothing" if record.get("dry_run") else "installed updates"), ""]
    for status, title in (("red", "Red"), ("yellow", "Yellow"), ("unknown", "Unknown"), ("green", "Green")):
        group = [r for r in record["items"] if r["status"] == status]
        if not group:
            continue
        lines.append("## %s (%d)" % (title, len(group)))
        lines.append("")
        for r in group:
            line = "- **%s** (`%s`): %s" % (r["label"], r["id"], r["detail"] or r.get("value") or "")
            if r.get("action") and status != "green":
                line += " | Action: %s" % r["action"]
            lines.append(line)
        lines.append("")
    w = record.get("winget")
    if w:
        lines += ["## Winget pass", "", "Status: %s%s" % (w["status"], " (dry run)" if w.get("dry_run") else ""), ""]
        for x in w["upgraded"]:
            lines.append("- Installed: %s %s -> %s" % (x["id"], x["from"], x["to"]))
        for x in w["failed"]:
            lines.append("- FAILED: %s rc=%s; last output: %s" % (x["id"], x["rc"], " | ".join(x["tail"])))
        for x in w["not_applied"]:
            lines.append("- Not applied (exit 0, version unchanged): %s" % x["id"])
        for x in w.get("unconfirmed", []):
            lines.append("- Unconfirmed (timed out; checked again on the next run): %s %s rc=%s; last output: %s" % (
                x["id"], x["from"], x["rc"], " | ".join(x.get("tail", []))))
        for x in w.get("completed_late", []):
            lines.append("- Completed after the run: %s %s -> %s" % (x["id"], x["from"], x["to"]))
        for x in w["deferred"]:
            lines.append("- Deferred (%s): %s %s -> %s (run %s)" % (
                x["reason"], x["id"], x["from"], x["available"], x.get("streak", 1)))
        for x in w["held"]:
            lines.append("- Waiting for you: %s %s -> %s. Run: %s" % (x["id"], x["from"], x["available"], x["command"]))
        for x in w.get("needs_admin", []):
            lines.append("- Waiting for you (admin, run in an administrator PowerShell): %s %s -> %s. Run: %s" % (
                x["id"], x["from"], x.get("available", ""), x["command"]))
        for x in w.get("waiting_close", []):
            lines.append("- Waiting for you (close the program first): %s %s -> %s. %s" % (
                x["id"], x["from"], x["available"], x["command"]))
        if w.get("dry_run"):
            for x in w.get("would_upgrade", []):
                lines.append("- Ready to install: %s %s -> %s" % (x["id"], x["from"], x["available"]))
        for x in w.get("waiting_fail", []):
            lines.append("- Waiting for you (%s; not retried by itself): %s %s -> %s. Run: %s" % (
                x["reason"], x["id"], x["from"], x["available"], x["command"]))
        for x in w.get("not_selected", []):
            lines.append("- Not selected: %s %s -> %s. To cover it: %s" % (
                x["id"], x["from"], x["available"], x["command"]))
        for x in w["checks"]:
            lines.append("- Post-upgrade check %s: %s (%s)" % (x["name"], x["status"], x["detail"]))
        for e in w["errors"]:
            lines.append("- ERROR: %s" % e)
        for e in w["warnings"]:
            lines.append("- Warning: %s" % e)
        for e in w.get("notes", []):
            lines.append("- Note: %s" % e)
        if w.get("raw_head"):
            lines += ["", "First output lines that could not be parsed:", "", "```"] + w["raw_head"] + ["```"]
        lines.append("")
    cp = record.get("claude")
    if cp:
        lines += ["## Claude pass", "", "Status: %s%s" % (cp.get("status"), " (dry run)" if cp.get("dry_run") else ""), ""]
        lines += ["- %s" % s for s in claude_summary(cp)]
        lines.append("")
    t = record.get("toast")
    if t:
        lines.append("Toast: %s" % (("shown" if t.get("ok") else "FAILED: %s" % t.get("detail")) if t.get("attempted") else "not sent"))
    return "\n".join(lines).rstrip() + "\n"


def _xml_safe(s):
    return "".join(ch for ch in str(s) if ch in "\t\n\r" or ord(ch) >= 0x20)


def send_toast(title, body, sticky, runner=subprocess.run, tag=TOAST_TAG):
    """Native toast; text travels in environment variables. -> (ok, detail)."""
    env = {"UPKEEP_TOAST_TITLE": _xml_safe(title)[:120],
           "UPKEEP_TOAST_BODY": _xml_safe(body)[:600],
           "UPKEEP_TOAST_TAG": tag,
           "UPKEEP_TOAST_GROUP": TOAST_GROUP,
           "UPKEEP_TOAST_STICKY": "1" if sticky else "0"}
    enc = base64.b64encode(_TOAST_PS.encode("utf-16-le")).decode("ascii")
    r = run_cmd(runner, [POWERSHELL, "-NoProfile", "-NonInteractive", "-EncodedCommand", enc],
                timeout=30, env_extra=env)
    out = (r.stdout or "").strip()
    if r.returncode == 0 and out.endswith("SHOWN"):
        return True, "shown"
    return False, ("rc=%s %s %s" % (r.returncode, out, (r.stderr or "").strip()))[:300]


# ------------------------------------------------------------------ commands ----

def load_seen() -> dict:
    d = read_json(STATE / "seen.json", {})
    return d if isinstance(d, dict) else {}


def sort_results(results):
    return sorted(results, key=lambda r: -RANK[r["status"]])


def new_run_file(stamp: str) -> Path:
    runs = STATE / "runs"
    runs.mkdir(parents=True, exist_ok=True)
    for n in range(1, 100):
        p = runs / ("%s%s.json" % (stamp, "" if n == 1 else "_%d" % n))
        if not p.exists():
            return p
    raise RuntimeError("no free run file name for %s" % stamp)


def cmd_claude(args, runner=subprocess.run) -> int:
    c = claude_pass(runner, dry_run=args.dry_run)
    print("claude pass: %s%s" % (c["status"], " (dry run)" if args.dry_run else ""))
    for line in claude_summary(c):
        print("  " + line)
    return 0


def cmd_pins(args, runner=subprocess.run) -> int:
    policy = load_policy()
    if policy is None:
        print("policy.json missing or unreadable: %s" % POLICY, file=sys.stderr)
        return 1
    installed, err = read_installed(runner)
    if installed is None:
        print("cannot read installed packages: %s" % err, file=sys.stderr)
        return 1
    pl = run_cmd(runner, [WINGET, "pin", "list", "--accept-source-agreements", "--disable-interactivity"], timeout=120)
    pinned, ok = parse_pin_list(pl.stdout)
    if not ok:
        print("cannot parse winget pin list", file=sys.stderr)
        return 1
    missing = missing_pins(policy, installed, pinned)
    if not missing:
        print("every installed held package is pinned")
        return 0
    rc = 0
    for pid in missing:
        cmd = [WINGET, "pin", "add", "--id", pid, "--exact", "--accept-source-agreements"]
        if args.apply:
            r = run_cmd(runner, cmd, timeout=120)
            print("%s: pin add rc=%s" % (pid, r.returncode))
            rc = rc or (1 if r.returncode != 0 else 0)
        else:
            print("would run: winget pin add --id %s --exact --accept-source-agreements" % pid)
    if not args.apply:
        print("(dry run; add --apply to pin)")
    return rc


# ------------------------------------------------------- report-only checks ----

def _python_arch_tag() -> str:
    if platform.machine().lower() in ("arm64", "aarch64"):
        return "arm64"
    return "64" if sys.maxsize > 2 ** 32 else "32"


def _release_tuple(s):
    """'3.14.10' -> (3, 14, 10); None for a pre-release or anything else."""
    m = re.fullmatch(r"(\d+)\.(\d+)\.(\d+)", str(s or "").strip())
    return tuple(int(x) for x in m.groups()) if m else None


def newest_python_patch(runner, major: int, minor: int, arch: str):
    """Newest release of major.minor for this architecture, from the Python install
    manager (`py list --online`). -> ((x, y, z), None) or (None, reason)."""
    r = run_cmd(runner, [PY_MANAGER, "list", "--online", "%d.%d" % (major, minor), "--format=json"],
                timeout=120)
    try:
        data = json.loads(r.stdout)
    except ValueError:
        data = None
    if not isinstance(data, dict) or not isinstance(data.get("versions"), list):
        return None, "the Python install manager (py) was not found or gave no online list"
    tag = "%d.%d-%s" % (major, minor, arch)
    best = None
    for e in data["versions"]:
        if not isinstance(e, dict) or e.get("tag") != tag:
            continue
        t = _release_tuple(e.get("sort-version"))
        if t and t[:2] == (major, minor) and (best is None or t > best):
            best = t
    if best is None:
        return None, "the online list has no %s release" % tag
    return best, None


def check_python_version(item, ctx):
    """Yellow when this Python is below the floor, or when a newer patch release
    of the same minor version is out. Read only: never installs."""
    v = sys.version_info
    value = "%d.%d.%d" % (v[0], v[1], v[2])
    floor = tuple(int(x) for x in str(item.get("yellow_below", "3.10")).split(".")[:2])
    action = item.get("action") or ("update Python: py install --update (Python install manager), "
                                     "or https://www.python.org/downloads/windows/")
    if (v[0], v[1]) < floor:
        return res(item, "yellow", value=value, waiting=True,
                   detail="Python %s is older than %d.%d" % (value, floor[0], floor[1]), action=action)
    if not item.get("check_online", True):
        return res(item, "green", value=value, detail="Python %s" % value)
    newest, why = newest_python_patch(ctx["runner"], v[0], v[1], _python_arch_tag())
    if newest is None:
        return res(item, "green", value=value,
                   detail="Python %s (newer patch release not checked: %s)" % (value, why))
    if newest > (v[0], v[1], v[2]):
        return res(item, "yellow", value=value, waiting=True,
                   detail="Python %s; %s is out" % (value, ".".join(map(str, newest))), action=action)
    return res(item, "green", value=value, detail="Python %s, the newest %d.%d release" % (value, v[0], v[1]))


def check_pip_outdated(item, ctx):
    """Count only. Python packages are never updated here: that breaks working environments."""
    r = run_cmd(ctx["runner"], [sys.executable, "-m", "pip", "list", "--outdated", "--format=json",
                                "--disable-pip-version-check"], timeout=item.get("timeout_s", 240))
    try:
        data = json.loads(r.stdout)
    except ValueError:
        data = None
    if not isinstance(data, list):
        return res(item, "unknown", detail="pip could not list outdated packages (rc=%s): %s" % (
            r.returncode, " | ".join(tail_lines(r.stderr, 1))))
    n = len(data)
    return res(item, "green", value=str(n),
               detail="%d outdated Python package(s); reported only, never updated here "
                      "(one by hand: python -m pip install --upgrade <name>)" % n)


_REBOOT_KEYS = (
    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
    r"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
)


def reg_key_exists(sub: str):
    """True / False for a key under HKEY_LOCAL_MACHINE; None when it cannot be told."""
    try:
        import winreg
    except ImportError:
        return None
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub))
        return True
    except FileNotFoundError:
        return False
    except OSError:
        return None


def check_pending_reboot(item, ctx):
    states = [reg_key_exists(k) for k in _REBOOT_KEYS]
    if any(s is True for s in states):
        return res(item, "yellow", value="pending", waiting=True,
                   detail="Windows is waiting for a restart to finish installing updates",
                   action=item.get("action") or "restart the computer")
    if all(s is None for s in states):
        return res(item, "unknown", detail="the restart flags could not be read")
    return res(item, "green", value="none", detail="no restart is pending")


disk_usage = shutil.disk_usage


def check_disk_free(item, ctx):
    drive = xp(item.get("drive", "%SystemDrive%")).rstrip("\\/") + "\\"
    try:
        u = disk_usage(drive)
    except OSError as e:
        return res(item, "unknown", detail="cannot read free space on %s: %s" % (drive, e))
    pct = 100.0 * u.free / u.total if u.total else 0.0
    detail = "%.0f GB free of %.0f GB on %s (%.0f%%)" % (u.free / 2 ** 30, u.total / 2 ** 30, drive, pct)
    yl, rl = item.get("yellow_below_pct", 10), item.get("red_below_pct", 5)
    status = "red" if pct < rl else ("yellow" if pct < yl else "green")
    return res(item, status, value="%.0f%%" % pct, detail=detail)


_WU_PS = ("$s=(New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher(); "
          "$r=$s.Search('IsInstalled=0 and IsHidden=0'); [Console]::Out.Write('COUNT ' + $r.Updates.Count)")


def check_windows_update(item, ctx):
    """Count of updates Windows has not installed yet. Read only."""
    r = run_cmd(ctx["runner"], [POWERSHELL, "-NoProfile", "-NonInteractive", "-Command", _WU_PS],
                timeout=item.get("timeout_s", 180))
    m = re.search(r"COUNT (\d+)", r.stdout)
    if not m:
        return res(item, "unknown", detail="Windows Update could not be queried (rc=%s): %s" % (
            r.returncode, " | ".join(tail_lines(r.stderr, 1))))
    n = int(m.group(1))
    if n:
        return res(item, "yellow", value=str(n), waiting=True,
                   detail="%d Windows update(s) not installed yet" % n,
                   action=item.get("action") or "Settings > Windows Update > install")
    return res(item, "green", value="0", detail="Windows Update has nothing waiting")


# ------------------------------------------------------------ the inventory ----

WEB_HOSTS = ("chrome", "edge", "msedge", "brave-browser", "brave", "chromium", "vivaldi", "opera")


def read_winget_list(runner):
    """winget list -> rows {name, id, version, source}; None when unreadable."""
    r = run_cmd(runner, [WINGET, "list", "--accept-source-agreements", "--disable-interactivity"], timeout=300)
    t = _first_table(r.stdout)
    if t is None:
        return None
    c = _cols(t, "list")
    return [{"name": row.get(c["name"], ""), "id": row.get(c["id"], ""), "version": row.get(c["version"], ""),
             "source": row.get(c["source"], "") if c["source"] else ""} for row in t["rows"]]


def unmatched_kind(name: str, info: dict, known: dict) -> tuple:
    """-> (kind, who updates it) for an app winget cannot match:
    web app | driver | manual | updates itself | no known updater."""
    low = name.lower()
    plow = ((info or {}).get("publisher") or "").strip().lower()
    if "\\" in plow and plow.rsplit("\\", 1)[-1] in WEB_HOSTS:
        return "web app", "nothing to update: a page installed through the browser"
    if any(w.lower() in low for w in known.get("driver_words") or []):
        return "driver", "Windows Update or the maker's own tool"
    for key, kind in (("manual", "manual"), ("self_updating", "updates itself")):
        for e in known.get(key) or []:
            n, p = (e.get("name") or "").lower(), (e.get("publisher") or "").lower()
            if (n and fnmatch.fnmatchcase(low, n)) or (p and plow and fnmatch.fnmatchcase(plow, p)):
                return kind, e.get("by", "")
    return "no known updater", "nobody that this tool knows of: check the maker's site"


def app_class(pid: str, name: str, policy: dict, fails: dict, uninstall: dict) -> str:
    if not is_selected(pid, policy):
        return "not selected"
    d = decide(pid, policy, set(), ())
    if d["action"] == "held":
        return "held: reported, waits for you"
    if d["action"] == "needs_admin":
        return "needs an administrator shell: waits for you"
    if fails.get(pid, 0) >= int(policy.get("fail_limit") or FAIL_LIMIT):
        return "failed %d runs in a row: waits for you" % fails[pid]
    if lookup(policy.get("in_use"), pid) or (uninstall.get(name) or {}).get("location"):
        return "updated; put off while it is running"
    return "updated"


def classify_installed(rows, policy, learned, uninstall, known) -> dict:
    """Split winget's installed list: catalogue apps (what upkeep updates), other
    sources, Windows components (a count) and apps winget cannot match."""
    fails = (learned or {}).get("fails") or {}
    cat, store, unmatched, components, seen_names = {}, [], [], 0, set()
    for row in rows:
        pid, src = row["id"], row["source"].strip().lower()
        up = pid.upper()
        if up.startswith("MSIX\\"):
            components += 1
        elif src == "winget" and not up.startswith("ARP\\"):
            if pid not in cat:
                cat[pid] = dict(row, **{"class": app_class(pid, row["name"], policy, fails, uninstall)})
        elif src and not up.startswith("ARP\\"):
            store.append(row)
        elif row["name"] not in seen_names:
            seen_names.add(row["name"])
            info = uninstall.get(row["name"]) or {}
            kind, by = unmatched_kind(row["name"], info, known)
            unmatched.append(dict(row, kind=kind, by=by, date=info.get("date", "")))
    return {"catalogue": sorted(cat.values(), key=lambda x: x["id"].lower()), "store": store,
            "unmatched": sorted(unmatched, key=lambda x: (x["kind"], x["name"].lower())),
            "components": components}


def _install_date(s):
    s = str(s or "").strip()
    if len(s) == 8 and s.isdigit():
        try:
            return datetime(int(s[:4]), int(s[4:6]), int(s[6:]))
        except ValueError:
            return None
    return None


def check_unmatched_stale(item, ctx):
    """Apps winget cannot match that show no version change for `stale_days`.
    Drivers and web apps are left out. The clock of an app with no known
    updater starts at its registry install date when this tool first sees it;
    an app that updates itself is counted from that first sight, because its
    install date often stays put when it updates."""
    rows = ctx.get("installed_rows")
    if rows is None:
        rows = read_winget_list(ctx["runner"])
    if rows is None:
        return res(item, "unknown", detail="winget list could not be read")
    policy = ctx.get("policy") or load_policy() or {}
    uninstall = ctx.get("uninstall")
    if uninstall is None:
        uninstall = read_uninstall()
    old = ctx.get("apps_seen")
    if old is None:
        old = read_json(STATE / "apps_seen.json", {})
    old = old if isinstance(old, dict) else {}
    now, limit = ctx["now"], item.get("stale_days", 180)
    quiet = {q.lower() for q in policy.get("quiet") or []}
    stamp = now.isoformat(timespec="seconds")
    track, stale = {}, []
    for a in classify_installed(rows, policy, {}, uninstall, load_known())["unmatched"]:
        if a["kind"] in ("driver", "web app"):
            continue
        prev = old.get(a["name"])
        if isinstance(prev, dict) and prev.get("version") == a["version"] and prev.get("since"):
            since = prev["since"]
        elif prev is None and a["kind"] in ("no known updater", "manual") and _install_date(a["date"]):
            since = _install_date(a["date"]).isoformat(timespec="seconds")
        else:
            since = stamp
        track[a["name"]] = {"version": a["version"], "since": since}
        days = (now - (parse_datetime(since) or now)).days
        if days > limit and a["name"].lower() not in quiet:
            stale.append((days, a["name"], a["kind"]))
    stale.sort(reverse=True)
    if stale:
        shown = ", ".join("%s (%d days, %s)" % (n, d, k) for d, n, k in stale[:6])
        more = " and %d more" % (len(stale) - 6) if len(stale) > 6 else ""
        return res(item, "yellow", value=str(len(stale)), apps_seen=track,
                   detail="%d app(s) outside winget with no version change for over %d days: %s%s" % (
                       len(stale), limit, shown, more),
                   action=item.get("action") or "check each maker for an update; silence one with: upkeep select --quiet \"<name>\"")
    return res(item, "green", value="0", apps_seen=track,
               detail="%d app(s) outside winget watched; none unchanged for over %d days" % (len(track), limit))


CHECK_TYPES = {
    "command_version": check_command_version,
    "file_age": check_file_age,
    "text_date": check_text_date,
    "text_date_due": check_text_date_due,
    "json_flag": check_json_flag,
    "dir_latest": check_dir_latest,
    "dir_list": check_dir_list,
    "claude_code": check_claude_code,
    "claude_plugins": check_claude_plugins,
    "claude_extension_lag": check_claude_extension_lag,
    "scheduled_tasks": check_scheduled_tasks,
    "winget_pass": check_winget_pass,
    "python_version": check_python_version,
    "pip_outdated": check_pip_outdated,
    "pending_reboot": check_pending_reboot,
    "disk_free": check_disk_free,
    "windows_update": check_windows_update,
    "unmatched_stale": check_unmatched_stale,
}


# ------------------------------------------------------------- the schedule ----

def stale_hours(policy):
    """Hours after which a missed scheduled run is an alarm; None without a schedule."""
    return STALE_HOURS.get(((policy or {}).get("schedule") or {}).get("every"))


def stale_run_alarm(now=None, policy=None):
    """Red item when the last run is older than the schedule allows; None when
    fresh, and None when the tester chose no schedule."""
    now = now or now_local()
    hours_allowed = stale_hours(policy)
    if hours_allowed is None:
        return None
    last = read_json(last_path(), None)
    item = {"id": "upkeep_stale", "label": "Scheduled upkeep ran",
            "action": "check the '%s' scheduled task (upkeep schedule), or run: upkeep daily" % TASK_NAME}
    if not isinstance(last, dict):
        return res(item, "yellow", detail="no upkeep run is recorded yet")
    fin = parse_datetime(last.get("finished"))
    if fin is None:
        return None
    hours = (now - fin).total_seconds() / 3600
    if hours > hours_allowed:
        return res(item, "red", value=fin.isoformat(timespec="seconds"),
                   detail="the scheduled upkeep did not run: last run %s (%.0f hours ago)" % (
                       fin.isoformat(timespec="seconds"), hours))
    return None


def valid_schedule(every, day=None, at=None):
    """-> (schedule dict, "") or (None, error text)."""
    every = (every or "").lower()
    if every == "manual":
        return {"every": "manual"}, ""
    if every not in ("daily", "weekdays", "weekly"):
        return None, "periodicity must be daily, weekdays, weekly or manual"
    m = re.fullmatch(r"\s*(\d{1,2}):(\d{2})\s*", at or "")
    if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
        return None, "time of day must be HH:MM (24 hours), for example 08:00"
    s = {"every": every, "at": "%02d:%s" % (int(m.group(1)), m.group(2))}
    if every == "weekly":
        d = (day or "").lower()[:3]
        if d not in DAYS:
            return None, "weekly needs a day: mon, tue, wed, thu, fri, sat or sun"
        s["day"] = d
    return s, ""


def task_python() -> str:
    exe = Path(sys.executable)
    w = exe.with_name("pythonw.exe")
    return str(w if w.exists() else exe)


def stable_script() -> Path:
    return HOME / "bin" / "upkeep.py"


def schtasks_create_cmd(schedule: dict, python_exe=None, script=None) -> list:
    """The schtasks line that registers the task for the current user (no elevation)."""
    run = '"%s" "%s" daily' % (python_exe or task_python(), script or stable_script())
    cmd = ["schtasks", "/create", "/f", "/tn", TASK_NAME, "/tr", run, "/st", schedule["at"]]
    if schedule["every"] == "daily":
        cmd += ["/sc", "daily"]
    elif schedule["every"] == "weekdays":
        cmd += ["/sc", "weekly", "/d", "MON,TUE,WED,THU,FRI"]
    else:
        cmd += ["/sc", "weekly", "/d", schedule["day"].upper()]
    return cmd


def describe_schedule(s) -> str:
    if not s:
        return "not chosen yet (no task)"
    if s.get("every") == "manual":
        return "manual (no task)"
    if s.get("every") == "weekly":
        return "weekly on %s at %s" % (s.get("day"), s.get("at"))
    return "%s at %s" % (s.get("every"), s.get("at"))


def task_registered(runner) -> bool:
    return run_cmd(runner, ["schtasks", "/query", "/tn", TASK_NAME], timeout=60).returncode == 0


# ------------------------------------------------------------- home folder ----

def ensure_home() -> list:
    """Create the home folders and copy the shipped defaults that are missing.
    Never overwrites a config file. Refreshes the stable copy of the script when
    one exists and this run comes from a different file."""
    made = []
    for d in (HOME / "config", STATE, HOME / "bin"):
        d.mkdir(parents=True, exist_ok=True)
    for name in ("policy.json", "items.json", "known_apps.json"):
        dst, src = HOME / "config" / name, SHIPPED / name
        if not dst.exists() and src.exists():
            shutil.copyfile(src, dst)
            made.append(name)
    if stable_script().exists():
        install_stable_copy()
    return made


def install_stable_copy() -> str:
    """Copy this script to <home>/bin/upkeep.py: the path the scheduled task runs,
    which stays the same when the plugin folder changes with a new version."""
    src, dst = Path(__file__).resolve(), stable_script()
    try:
        if dst.exists() and (src == dst.resolve() or src.read_bytes() == dst.read_bytes()):
            return "stable copy is current: %s" % dst
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, dst)
        return "stable copy written: %s" % dst
    except OSError as e:
        return "stable copy NOT written (%s): %s" % (e, dst)


def edit_policy(fn) -> dict:
    """Read the tester's policy.json, apply fn to it, write it back."""
    raw = read_json(POLICY)
    if not isinstance(raw, dict):
        raise SystemExit("policy.json is missing or unreadable: %s (run: upkeep setup)" % POLICY)
    fn(raw)
    write_json(POLICY, raw)
    return raw


def redact(text: str) -> str:
    """Replace the profile path, the user name and the machine name by their variable names."""
    for var in ("USERPROFILE", "COMPUTERNAME", "USERNAME"):
        v = os.environ.get(var)
        if v and len(v) >= 3:
            for form in {v, v.replace("\\", "/")}:
                text = re.sub(re.escape(form), lambda m, t="%" + var + "%": t, text, flags=re.I)
    return text


# ------------------------------------------------------------ report + toast ----

def build_toast(record: dict, details_path) -> tuple:
    """-> (title, body, sticky). Body is at most 600 characters; the Details
    line is always kept whole."""
    c = record["counts"]
    w = record.get("winget") or {}
    ready = w.get("would_upgrade", []) if w.get("dry_run") else []
    if c["red"]:
        head = "%d red" % c["red"] + (", %d to review" % c["yellow"] if c["yellow"] else "")
    elif ready:
        head = "%d update(s) ready" % len(ready)
    elif c["yellow"]:
        head = "%d to review" % c["yellow"]
    else:
        head = "all current"
    title = "Upkeep %s: %s" % (record["date"], head)
    parts = ["RED %s: %s" % (r["label"], r["detail"]) for r in record["items"] if r["status"] == "red"]
    if ready:
        parts.append("Ready to install (in Claude Code: %s apply): %s" % (SKILL, ", ".join(x["id"] for x in ready)))
    if w.get("upgraded"):
        parts.append("Installed: " + ", ".join("%s %s" % (x["id"], x["to"]) for x in w["upgraded"]))
    waiting = ["%s (%s)" % (r["label"], r.get("action") or "see report")
               for r in record["items"] if r.get("waiting")]
    waiting += ["%s %s (winget upgrade --id %s)" % (x["id"], x["available"], x["id"]) for x in w.get("held", [])]
    waiting += ["%s %s (admin PowerShell: winget upgrade --id %s)" % (x["id"], x.get("available", ""), x["id"])
                for x in w.get("needs_admin", [])]
    waiting += ["%s %s (close %s first)" % (x["id"], x["available"], ", ".join(x.get("processes") or ["the program"]))
                for x in w.get("waiting_close", [])]
    waiting += ["%s %s (%s)" % (x["id"], x["available"], x["reason"]) for x in w.get("waiting_fail", [])]
    if waiting:
        parts.append("Waiting for you: " + "; ".join(waiting))
    if w.get("unconfirmed"):
        parts.append("Unconfirmed (timed out): " + ", ".join(x["id"] for x in w["unconfirmed"]))
    if w.get("completed_late"):
        parts.append("Completed after the run: " + ", ".join(x["id"] for x in w["completed_late"]))
    if w.get("deferred"):
        parts.append("Deferred: " + ", ".join(
            "%s (%s)" % (x["id"], x.get("reason", "deferred: in use").replace("deferred: ", "")) for x in w["deferred"]))
    details = "Details: %s" % details_path
    budget = 600 - len(details) - 1
    body = "\n".join(parts)
    if len(body) > budget:
        body = body[:max(0, budget - 1)].rstrip() + "…"
    body = (body + "\n" + details) if body else details
    return title, body, c["red"] > 0


# ------------------------------------------------------------------ commands ----

def collect(runner, now, install=False, include_winget=True, only=None, policy=None):
    """Winget pass + Claude pass + item check.
    -> (winget_result|None, results, items, claude_result|None, apps_seen|None).
    `install` False reads and classifies only. Without the passes (check mode)
    the items read the previous run's Claude pass."""
    items = load_items()
    prev_last = read_json(last_path(), None)
    prev = prev_last if isinstance(prev_last, dict) else {}
    policy = policy if policy is not None else load_policy()
    w = c = None
    c_ctx = prev.get("claude")
    if include_winget:
        w = winget_pass(runner, policy=policy, dry_run=not install, items=items,
                        prev_winget=prev.get("winget"), only=only, learned=load_learned())
        c = claude_pass(runner, dry_run=(not install) or bool(only), now=now)
        c_ctx = c
    ctx = make_ctx(runner, load_seen(), now, w, prev_last, claude=c_ctx)
    ctx["policy"] = policy
    results = run_items(items, ctx)
    apps_seen = None
    for r in results:
        if "apps_seen" in r:
            apps_seen = r.pop("apps_seen")
    pt = previous_toast_item(prev_last)
    if pt:
        results.append(pt)
    return w, results, items, c, apps_seen


def run_pass(runner, now, clock, install, only=None, label="daily", toast=True, print_only=False) -> int:
    """One run: the passes, the checks, the record, the report and the toast.
    `install` False lists and classifies; `print_only` writes nothing."""
    policy = load_policy()
    started = now.isoformat(timespec="seconds")
    w, results, _, c, apps_seen = collect(runner, now, install=install, only=only, policy=policy)
    results = sort_results(results)
    record = {"schema": "upkeep-run/1", "version": __version__, "date": now.date().isoformat(),
              "started": started, "finished": clock().isoformat(timespec="seconds"),
              "mode": (policy or {}).get("mode", "report"), "run": label,
              "dry_run": not install, "winget": w, "claude": c,
              "winget_item": next((r for r in results if r["id"] == "winget_pass"), None),
              "items": results, "counts": counts(results),
              "toast": {"attempted": False, "ok": None, "detail": ""}}
    record["exit"] = 1 if record["counts"]["red"] else 0
    if print_only:
        print(render_report(record))
        return record["exit"]

    title, body, sticky = build_toast(record, STATE / "last_report.md")
    # Everything the toast points to is on disk before the toast is sent.
    write_text(STATE / "last_report.md", render_report(record))
    run_path = new_run_file(now.strftime("%Y-%m-%d_%H%M%S"))
    with open(run_path, "x", encoding="utf-8") as f:     # created once, never overwritten by another run
        json.dump(record, f, ensure_ascii=False, indent=2)
        f.write("\n")
    write_json(last_path(), record)
    write_json(STATE / "seen.json", update_seen(load_seen(), results, now))
    if w and isinstance(w.get("learned"), dict):
        write_json(STATE / "learned.json", w["learned"])
    if apps_seen is not None:
        write_json(STATE / "apps_seen.json", apps_seen)
    if toast:
        ok, detail = send_toast(title, body, sticky, runner)
        record["toast"] = {"attempted": True, "ok": ok, "detail": detail,
                           "title": title, "sticky": sticky}
        write_json(run_path, record)
        write_json(last_path(), record)
        write_text(STATE / "last_report.md", render_report(record))
    print(title)
    print(body)
    return record["exit"]


def cmd_daily(args, runner=subprocess.run, now=None, clock=now_local) -> int:
    """The scheduled run. Report mode lists; auto mode installs."""
    policy = load_policy() or {}
    install = policy.get("mode") == "auto" and not args.dry_run
    return run_pass(runner, now or now_local(), clock, install, label="daily",
                    toast=not args.no_toast, print_only=args.dry_run)


def cmd_apply(args, runner=subprocess.run, now=None, clock=now_local) -> int:
    """Install now, with the tester present, in either mode."""
    only = [x.strip() for x in (args.only or "").split(",") if x.strip()]
    return run_pass(runner, now or now_local(), clock, True, only=only or None, label="apply", toast=False)


def cmd_check(args, runner=subprocess.run, now=None) -> int:
    now = now or now_local()
    policy = load_policy()
    _, results, _, _, _ = collect(runner, now, include_winget=False, policy=policy)
    alarm = stale_run_alarm(now, policy)
    if alarm:
        results.append(alarm)
    results = sort_results(results)
    record = {"date": now.date().isoformat(), "finished": now.isoformat(timespec="seconds"),
              "items": results, "counts": counts(results), "toast": None}
    if args.json:
        print(json.dumps(record, ensure_ascii=False, indent=2))
    else:
        print(render_report(record))
    if args.toast:
        title, body, sticky = build_toast(record, STATE / "last_report.md")
        ok, detail = send_toast(title, body, sticky, runner)
        if not ok:
            print("toast failed: %s" % detail, file=sys.stderr)
    return 1 if record["counts"]["red"] else 0


def cmd_status(args) -> int:
    policy = load_policy() or {}
    print("mode: %s; schedule: %s; apps: %s" % (
        policy.get("mode", "report"), describe_schedule(policy.get("schedule")), describe_selection(policy)))
    last = read_json(last_path(), None)
    if not isinstance(last, dict):
        print("no upkeep run recorded yet (try: upkeep daily --dry-run)")
        return 0
    c = last["counts"]
    print("last run %s (%s): %d red, %d yellow, %d unknown, %d green" % (
        last.get("finished"), "listed only" if last.get("dry_run") else "installed",
        c["red"], c["yellow"], c["unknown"], c["green"]))
    for r in last["items"]:
        if r["status"] in ("red", "yellow"):
            print("  %s %s: %s" % (r["status"].upper(), r["label"], r["detail"]))
    t = last.get("toast") or {}
    if t.get("attempted"):
        print("toast: %s" % ("shown" if t.get("ok") else "FAILED " + str(t.get("detail"))))
    print("report: %s" % (STATE / "last_report.md"))
    return 0


def describe_selection(policy) -> str:
    sel = (policy or {}).get("selection") or {}
    if sel.get("mode") == "named":
        return "only %s" % (", ".join(sel.get("only") or []) or "(none named)")
    ex = sel.get("excluded") or []
    return "all catalogue apps" + (" except %s" % ", ".join(ex) if ex else "")


def cmd_mode(args) -> int:
    if args.value:
        edit_policy(lambda p: p.__setitem__("mode", args.value))
    policy = load_policy() or {}
    print("mode: %s" % policy.get("mode"))
    print("report = the scheduled run lists updates and installs nothing; install with: upkeep apply"
          if policy.get("mode") == "report" else
          "auto = the scheduled run installs the updates of the selected apps by itself")
    return 0


def _ids(s) -> list:
    return [x.strip() for x in (s or "").split(",") if x.strip()]


def cmd_select(args) -> int:
    """Choose the apps. Default: all catalogue apps."""
    def change(p):
        sel = p.get("selection") if isinstance(p.get("selection"), dict) else {}
        sel.setdefault("mode", "all")
        sel.setdefault("excluded", [])
        sel.setdefault("only", [])
        override = p.setdefault("held_override", [])
        if args.all:
            sel.update(mode="all", excluded=[], only=[])
        if args.only:
            sel.update(mode="named", only=_ids(args.only))
        for pid in _ids(args.remove):
            if sel["mode"] == "named":
                sel["only"] = [x for x in sel["only"] if x != pid]
            elif pid not in sel["excluded"]:
                sel["excluded"].append(pid)
            if pid in override:
                override.remove(pid)
        for pid in _ids(args.add):
            sel["excluded"] = [x for x in sel["excluded"] if x != pid]
            if sel["mode"] == "named" and pid not in sel["only"]:
                sel["only"].append(pid)
            if match_any(pid, p.get("held")) and pid not in override:
                override.append(pid)
                print("%s is held by default (an upgrade can close a program you are working in, or "
                      "change the Python your tools run on); it is now updated like the others." % pid)
        p["selection"] = sel
        quiet = p.setdefault("quiet", [])
        if args.quiet and args.quiet not in quiet:
            quiet.append(args.quiet)
        if args.unquiet and args.unquiet in quiet:
            quiet.remove(args.unquiet)
    if any((args.all, args.only, args.remove, args.add, args.quiet, args.unquiet)):
        edit_policy(change)
    policy = load_policy() or {}
    print("apps: %s" % describe_selection(policy))
    held = [h for h in policy.get("held", []) if h not in policy.get("held_override", [])]
    print("held (reported, never installed): %s" % (", ".join(held) or "none"))
    if policy.get("held_override"):
        print("held by default but added by you: %s" % ", ".join(policy["held_override"]))
    if policy.get("quiet"):
        print("silenced in the stale check: %s" % ", ".join(policy["quiet"]))
    return 0


def set_schedule(schedule: dict, runner, dry_run=False) -> int:
    """Store the tester's schedule and register (or remove) the task."""
    if schedule["every"] == "manual":
        if dry_run:
            print("would remove the task '%s' if it exists" % TASK_NAME)
            return 0
        edit_policy(lambda p: p.__setitem__("schedule", schedule))
        if task_registered(runner):
            r = run_cmd(runner, ["schtasks", "/delete", "/f", "/tn", TASK_NAME], timeout=60)
            print("task removed" if r.returncode == 0 else "task NOT removed (rc=%s): %s" % (
                r.returncode, " | ".join(tail_lines(r.stdout + "\n" + r.stderr, 2))))
            return 0 if r.returncode == 0 else 1
        print("schedule: manual (no task)")
        return 0
    cmd = schtasks_create_cmd(schedule)
    if dry_run:
        print("would run: " + subprocess.list2cmdline(cmd))
        return 0
    edit_policy(lambda p: p.__setitem__("schedule", schedule))
    print(install_stable_copy())
    r = run_cmd(runner, cmd, timeout=60)
    if r.returncode != 0:
        print("task NOT registered (rc=%s): %s" % (r.returncode, " | ".join(tail_lines(r.stdout + "\n" + r.stderr, 2))))
        return 1
    print("task '%s' registered: %s" % (TASK_NAME, describe_schedule(schedule)))
    return 0


def cmd_schedule(args, runner=subprocess.run) -> int:
    every = "manual" if args.off else args.every
    if every:
        schedule, err = valid_schedule(every, args.day, args.at)
        if schedule is None:
            print(err, file=sys.stderr)
            return 2
        return set_schedule(schedule, runner, dry_run=args.dry_run)
    policy = load_policy() or {}
    print("schedule: %s" % describe_schedule(policy.get("schedule")))
    print("task '%s': %s" % (TASK_NAME, "registered" if task_registered(runner) else "not registered"))
    return 0


def cmd_setup(args, runner=subprocess.run) -> int:
    print("home: %s" % HOME)
    print(install_stable_copy())
    rc = 0
    if args.every:
        schedule, err = valid_schedule(args.every, args.day, args.at)
        if schedule is None:
            print(err, file=sys.stderr)
            return 2
        rc = set_schedule(schedule, runner)
    policy = load_policy() or {}
    print("mode: %s; schedule: %s; apps: %s" % (
        policy.get("mode"), describe_schedule(policy.get("schedule")), describe_selection(policy)))
    return rc


def cmd_apps(args, runner=subprocess.run) -> int:
    policy = load_policy()
    rows = read_winget_list(runner)
    if policy is None or rows is None:
        print("cannot read %s" % ("the policy" if policy is None else "winget list"), file=sys.stderr)
        return 1
    inv = classify_installed(rows, policy, load_learned(), read_uninstall(), load_known())
    if args.match:
        return _apps_match(inv, runner)
    if args.json:
        print(json.dumps(inv, ensure_ascii=False, indent=2))
        return 0
    cat, un = inv["catalogue"], inv["unmatched"]
    print("Covered by upkeep: %d app(s) in the winget catalogue" % len(cat))
    for a in cat:
        print("  %-42s %-18s %s" % (a["id"][:42], a["version"][:18], a["class"]))
    if inv["store"]:
        print("\nOther winget sources (Microsoft Store): %d, updated when winget offers an upgrade" % len(inv["store"]))
        for a in inv["store"]:
            print("  %-42s %s" % (a["name"][:42], a["version"]))
    print("\nWindows and Store components: %d (Windows and the Microsoft Store update these)" % inv["components"])
    print("\nNot covered: %d app(s) winget cannot match. Upkeep watches them and installs nothing." % len(un))
    for a in un:
        print("  %-42s %-18s %-17s %s" % (a["name"][:42], a["version"][:18], a["kind"], a["by"]))
    print("\nSome of these may exist in the winget catalogue under the same name: upkeep apps --match")
    return 0


def _apps_match(inv, runner) -> int:
    """Search the winget catalogue by exact name for the apps winget did not match."""
    found = 0
    todo = [a for a in inv["unmatched"] if a["kind"] not in ("driver", "web app")]
    for a in todo:
        r = run_cmd(runner, [WINGET, "search", "--name", a["name"], "--source", "winget", "--exact",
                             "--accept-source-agreements", "--disable-interactivity"], timeout=60)
        t = _first_table(r.stdout)
        if t is None:
            continue
        c = _cols(t, "list")
        rows, _ = _table_ids(t, None, c["id"])
        for row in rows[:3]:
            found += 1
            print("%s (installed %s): catalogue id %s, version %s" % (
                a["name"], a["version"], row[c["id"]], row.get(c["version"], "")))
            print("    to bring it under winget, if you choose to: winget install --id %s --exact" % row[c["id"]])
    print("%d of %d unmatched app(s) have an exact-name entry in the winget catalogue." % (found, len(todo)))
    print("Upkeep never installs over an unmatched app: the line above is yours to run or not.")
    return 0


def _doctor_rows(runner) -> list:
    """-> [(name, ok True|False|None, detail)]. None is information, not a problem."""
    rows = [("python", True, "%s (needs %d.%d or newer)" % (sys.version.split()[0], MIN_PYTHON[0], MIN_PYTHON[1]))]
    r = run_cmd(runner, [WINGET, "--version"], timeout=60)
    m = re.search(r"v?(\d+\.\d+(?:\.\d+)*)", r.stdout)
    rows.append(("winget", bool(m and r.returncode == 0),
                 m.group(1) if m else "not found: install 'App Installer' from the Microsoft Store"))
    cv = claude_version(runner)
    rows.append(("claude CLI", True if cv else None, cv or "not found on PATH (the Claude pass will report it)"))
    rows.append(("powershell", Path(POWERSHELL).exists(), POWERSHELL))
    try:
        probe = HOME / ".write_probe"
        probe.write_text("x", encoding="utf-8")
        probe.unlink()
        rows.append(("home folder", True, str(HOME)))
    except OSError as e:
        rows.append(("home folder", False, "%s is not writable: %s" % (HOME, e)))
    policy = load_policy()
    rows.append(("config", policy is not None and bool(load_items()),
                 "%d item(s); mode %s" % (len(load_items()), (policy or {}).get("mode"))
                 if policy is not None else "policy.json missing: run upkeep setup"))
    sched = (policy or {}).get("schedule")
    reg = task_registered(runner)
    wants = bool(sched) and sched.get("every") != "manual"
    rows.append(("schedule", None if not wants else reg,
                 "%s; task %s" % (describe_schedule(sched), "registered" if reg else "not registered")))
    rows.append(("stable copy", None, str(stable_script()) if stable_script().exists()
                 else "not needed yet: choosing a schedule writes it (only the scheduled task runs it)"))
    return rows


def cmd_doctor(args, runner=subprocess.run) -> int:
    rows = _doctor_rows(runner)
    for name, ok, detail in rows:
        print("%-8s %-12s %s" % ("ok" if ok else ("PROBLEM" if ok is False else "info"), name, detail))
    return 1 if any(ok is False for _, ok, _ in rows) else 0


def cmd_diag(args, runner=subprocess.run) -> int:
    """A block to paste into a feedback issue. It names the installed package
    ids that had updates; read it before sending."""
    policy = load_policy() or {}
    out = ["upkeep-beta %s diagnostic" % __version__, "windows: %s" % platform.platform()]
    out += ["%s: %s (%s)" % (n, "ok" if ok else ("PROBLEM" if ok is False else "info"), d)
            for n, ok, d in _doctor_rows(runner)]
    out.append("apps: %s" % describe_selection(policy))
    last = read_json(last_path(), None)
    if isinstance(last, dict):
        c = last.get("counts", {})
        out.append("last run %s (%s, %s): %s red, %s yellow, %s unknown, %s green" % (
            last.get("finished"), last.get("run"), "listed only" if last.get("dry_run") else "installed",
            c.get("red"), c.get("yellow"), c.get("unknown"), c.get("green")))
        for r in last.get("items", []):
            if r.get("status") != "green":
                out.append("  %s %s: %s" % (r.get("status"), r.get("id"), r.get("detail")))
        w = last.get("winget") or {}
        for key in ("upgraded", "failed", "not_applied", "unconfirmed", "deferred", "held", "needs_admin",
                    "waiting_close", "waiting_fail", "not_selected", "would_upgrade"):
            if w.get(key):
                out.append("  winget %s: %s" % (key, ", ".join(
                    "%s%s" % (x.get("id"), " rc=%s" % x["rc"] if "rc" in x else "") for x in w[key])))
        for key in ("errors", "warnings", "notes", "raw_head"):
            for x in w.get(key) or []:
                out.append("  winget %s: %s" % (key, x))
    else:
        out.append("no run recorded yet")
    out.append("learned failures: %s" % json.dumps(load_learned().get("fails"), ensure_ascii=False))
    print(redact("\n".join(out)))
    return 0


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(prog="upkeep", description=__doc__.split("\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)

    def sched_args(p):
        p.add_argument("--every", choices=["daily", "weekdays", "weekly", "manual"], help="how often it runs")
        p.add_argument("--day", help="for weekly: mon, tue, wed, thu, fri, sat or sun")
        p.add_argument("--at", help="time of day, HH:MM")
    st = sub.add_parser("setup", help="create the home folder and the stable copy; with --every, the schedule")
    sched_args(st)
    sc = sub.add_parser("schedule", help="show or choose when upkeep runs")
    sched_args(sc)
    sc.add_argument("--off", action="store_true", help="no scheduled task (manual)")
    sc.add_argument("--dry-run", action="store_true", help="print the task line; register nothing")
    se = sub.add_parser("select", help="show or choose the apps (default: all catalogue apps)")
    se.add_argument("--add", help="ids to cover again, comma separated")
    se.add_argument("--remove", help="ids to leave out, comma separated")
    se.add_argument("--only", help="cover only these ids (and no app installed later)")
    se.add_argument("--all", action="store_true", help="back to the default: all catalogue apps")
    se.add_argument("--quiet", help="silence one app name in the stale check")
    se.add_argument("--unquiet", help="undo --quiet for one app name")
    mo = sub.add_parser("mode", help="show or set the mode")
    mo.add_argument("value", nargs="?", choices=["report", "auto"])
    a = sub.add_parser("apps", help="every installed app and what upkeep does with it")
    a.add_argument("--match", action="store_true", help="search the winget catalogue for the unmatched apps")
    a.add_argument("--json", action="store_true")
    d = sub.add_parser("daily", help="the scheduled run: list (report mode) or install (auto mode), check, record, toast")
    d.add_argument("--dry-run", action="store_true", help="read and classify only; install, record and toast nothing")
    d.add_argument("--no-toast", action="store_true")
    ap_ = sub.add_parser("apply", help="install the listed updates now")
    ap_.add_argument("--only", help="install only these ids, comma separated")
    c = sub.add_parser("check", help="item check only")
    c.add_argument("--json", action="store_true")
    c.add_argument("--toast", action="store_true")
    cl = sub.add_parser("claude", help="Claude pass only: claude update, marketplaces, plugins")
    cl.add_argument("--dry-run", action="store_true", help="read versions only; update nothing")
    p = sub.add_parser("pins", help="held package ids that lack a winget pin")
    p.add_argument("--apply", action="store_true")
    sub.add_parser("doctor", help="is everything this tool needs in place")
    sub.add_parser("diag", help="a redacted block to send as feedback")
    sub.add_parser("status", help="summary of the last run")
    sub.add_parser("version", help="print the version")
    return ap


def main(argv=None, runner=subprocess.run) -> int:
    args = build_parser().parse_args(argv)
    if args.cmd == "version":
        print(__version__)
        return 0
    try:
        ensure_home()
    except OSError as e:
        print("cannot prepare the home folder %s: %s" % (HOME, e), file=sys.stderr)
        return 2
    with_runner = {"daily": cmd_daily, "apply": cmd_apply, "check": cmd_check, "claude": cmd_claude,
                   "pins": cmd_pins, "apps": cmd_apps, "schedule": cmd_schedule, "setup": cmd_setup,
                   "doctor": cmd_doctor, "diag": cmd_diag}
    if args.cmd in with_runner:
        return with_runner[args.cmd](args, runner)
    return {"select": cmd_select, "mode": cmd_mode, "status": cmd_status}[args.cmd](args)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass
    sys.exit(main())
