"""Tests for upkeep-beta. Every external call goes through a fake runner and every
path lives under pytest's tmp_path: nothing here reads or changes the machine.
All names, versions and tables are invented."""
import json
import subprocess
import sys
from collections import namedtuple
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import upkeep  # noqa: E402

NOW = datetime(2026, 1, 15, 8, 0, 0)
EN = ["Name", "Id", "Version", "Available", "Source"]
NONE_EN = "No installed package found matching input criteria.\n"


def table(header, rows):
    cols = list(zip(*([header] + [list(r) for r in rows])))
    widths = [max(len(str(c)) for c in col) + 2 for col in cols]
    line = lambda r: "".join(str(c).ljust(w) for c, w in zip(r, widths)).rstrip()  # noqa: E731
    return "\n".join([line(header), "-" * sum(widths)] + [line(r) for r in rows]) + "\n"


def out(stdout="", rc=0, stderr=""):
    return SimpleNamespace(returncode=rc, stdout=stdout, stderr=stderr)


class World:
    """A fake machine: what winget, tasklist, PowerShell, claude, schtasks and pip answer."""

    def __init__(self):
        self.installed = {"Acme.Editor": "1.0", "Acme.Player": "2.0", "Globex.Tool": "3.0"}
        self.upgrades = [("Acme Editor", "Acme.Editor", "1.0", "1.1"),
                         ("Acme Player", "Acme.Player", "2.0", "2.5"),
                         ("Globex Tool", "Globex.Tool", "3.0", "3.1")]
        self.outcomes = {}            # id -> ok | fail | timeout | noop | admin
        self.attempted = []
        self.running = set()          # process image names
        self.proc_paths = []
        self.list_text = None
        self.list_rc = 0
        self.installed_rows = []      # winget list rows: (name, id, version, source)
        self.search = {}              # app name -> [(name, id, version)]
        self.claude = "2.1.0"
        self.claude_after = None
        self.wu_count = 0
        self.pip = "[]"
        self.task = None
        self.schtasks = []
        self.toasts = []
        self.calls = []

    def __call__(self, cmd, **kw):
        self.calls.append(cmd)
        a = cmd[1:]
        if cmd[0] == "WINGET.EXE":
            if a[0] == "export":
                pkgs = [{"PackageIdentifier": k, "Version": v} for k, v in self.installed.items()]
                Path(a[a.index("-o") + 1]).write_text(json.dumps({"Sources": [{"Packages": pkgs}]}), encoding="utf-8")
                return out()
            if a[0] == "upgrade" and "--id" in a:
                pid = a[a.index("--id") + 1]
                self.attempted.append(pid)
                how = self.outcomes.get(pid, "ok")
                if how == "timeout":
                    raise subprocess.TimeoutExpired(cmd, 1)
                if how == "ok":
                    self.installed[pid] = next(u[3] for u in self.upgrades if u[1] == pid)
                return out("installer said things", {"ok": 0, "noop": 0, "fail": 1, "admin": 1603}[how])
            if a[0] == "upgrade":
                if self.list_text is not None:
                    return out(self.list_text, self.list_rc)
                rows = [u for u in self.upgrades if self.installed.get(u[1]) != u[3]]
                return out(table(EN, [u + ("winget",) for u in rows]) if rows else NONE_EN)
            if a[0] == "pin":
                return out("No pins configured.\n")
            if a[0] == "list":
                return out(table(["Name", "Id", "Version", "Source"], self.installed_rows))
            if a[0] == "search":
                hits = self.search.get(a[a.index("--name") + 1])
                return out(table(["Name", "Id", "Version"], hits) if hits else "No package found.\n", 0 if hits else 1)
            if a[0] == "--version":
                return out("v1.9.0\n")
        if cmd[0] == "tasklist":
            return out("".join('"%s","1","Console","1","10 K"\n' % n for n in sorted(self.running)))
        if cmd[0] == "POWERSHELL.EXE":
            text = " ".join(a)
            if "Win32_Process" in text:
                return out("\n".join(self.proc_paths))
            if "Microsoft.Update.Session" in text:
                return out("COUNT %s" % self.wu_count) if self.wu_count is not None else out("", 1, "access denied")
            self.toasts.append(kw["env"]["UPKEEP_TOAST_TITLE"])
            return out("SHOWN")
        if cmd[0] == "claude":
            if a == ["--version"]:
                return out("%s (Claude Code)\n" % self.claude)
            if a == ["update"] and self.claude_after:
                self.claude = self.claude_after
            return out()
        if cmd[0] == "schtasks":
            self.schtasks.append(cmd)
            if a[0] == "/query":
                return out(rc=0 if self.task else 1)
            self.task = cmd if a[0] == "/create" else None
            return out()
        if a[:2] == ["-m", "pip"]:
            return out(self.pip)
        return out(rc=-1, stderr="unexpected command %r" % (cmd,))


@pytest.fixture
def world(tmp_path, monkeypatch):
    home = tmp_path / "home"
    for name, val in (("HOME", home), ("POLICY", home / "config" / "policy.json"),
                      ("ITEMS", home / "config" / "items.json"), ("KNOWN", home / "config" / "known_apps.json"),
                      ("STATE", home / "state"), ("SHIPPED", ROOT / "config"),
                      ("WINGET", "WINGET.EXE"), ("POWERSHELL", "POWERSHELL.EXE"),
                      ("CLAUDE_PLUGINS_FILE", str(tmp_path / "no_plugins.json"))):
        monkeypatch.setattr(upkeep, name, val)
    monkeypatch.setattr(upkeep, "read_uninstall", lambda: w.uninstall)
    upkeep.ensure_home()
    upkeep.write_json(upkeep.ITEMS, {"items": [{"id": "winget_pass", "label": "App updates", "type": "winget_pass"}]})
    w = World()
    w.uninstall = {}
    w.home = home
    return w


def policy(**kw):
    p = upkeep.load_policy()
    p.update(kw)
    return p


def set_policy(**kw):
    upkeep.edit_policy(lambda p: p.update(kw))


def args(**kw):
    base = dict(dry_run=False, no_toast=False, only=None, json=False, toast=False, match=False,
                every=None, day=None, at=None, off=False, add=None, remove=None, all=False,
                quiet=None, unquiet=None, value=None)
    base.update(kw)
    return SimpleNamespace(**base)


# ------------------------------------------------------------------ listing ----

def test_upgrade_list_english_and_portuguese_headers():
    pkgs, warns, ok = upkeep.parse_upgrade_list(table(EN, [("Acme Editor", "Acme.Editor", "1.0", "1.1", "winget")]))
    assert ok and not warns and pkgs == [{"id": "Acme.Editor", "name": "Acme Editor", "version": "1.0", "available": "1.1"}]
    pt = table(["Nome", "ID", "Versão", "Disponível", "Origem"], [("Acme Editor", "Acme.Editor", "1.0", "1.1", "winget")])
    assert upkeep.parse_upgrade_list(pt + "1 atualizações disponíveis.\n")[0][0]["available"] == "1.1"


def test_upgrade_list_reads_columns_by_position_in_an_unknown_language():
    # A synthetic header: no word of it is known, not even 'id'.
    text = table(["Nimi", "Tunnus", "Versio", "Saatavilla", "Lahde"],
                 [("Acme Editor", "Acme.Editor", "1.0", "1.1", "winget")])
    pkgs, _, ok = upkeep.parse_upgrade_list(text)
    assert ok and pkgs[0] == {"id": "Acme.Editor", "name": "Acme Editor", "version": "1.0", "available": "1.1"}


def test_empty_list_is_ok_by_exit_code_in_any_language_and_not_ok_otherwise():
    assert upkeep.parse_upgrade_list("Ei paivityksia.\n", rc=0) == ([], [], True)
    assert upkeep.parse_upgrade_list("Ei paivityksia.\n", rc=1) == ([], [], False)
    assert upkeep.parse_upgrade_list(NONE_EN) == ([], [], True)        # the English phrase, no exit code
    assert upkeep.parse_upgrade_list("garbage") == ([], [], False)


def test_footer_is_not_a_row_and_a_truncated_id_is_skipped_with_a_warning():
    text = table(EN, [("Acme Editor", "Acme.Editor", "1.0", "1.1", "winget"),
                      ("Long Name", "Globex.VeryLo…", "1.0", "2.0", "winget")]) + "2 upgrades available.\n"
    pkgs, warns, ok = upkeep.parse_upgrade_list(text)
    assert ok and [p["id"] for p in pkgs] == ["Acme.Editor"] and "truncated" in warns[0]


def test_notes_count_the_explicit_targeting_table_and_the_unknown_versions():
    text = (table(EN, [("Acme Editor", "Acme.Editor", "1.0", "1.1", "winget")]) + "1 upgrades available.\n\n"
            + "The following packages require explicit targeting:\n"
            + table(EN, [("Globex Tool", "Globex.Tool", "3.0", "3.1", "winget")])
            + "\n2 package(s) have version numbers that cannot be determined. Use --include-unknown to see all results.\n")
    notes = upkeep.parse_upgrade_notes(text)
    assert notes[0].startswith("1 package(s)") and notes[1].startswith("2 package(s)")
    assert [p["id"] for p in upkeep.parse_upgrade_list(text)[0]] == ["Acme.Editor"]


# ----------------------------------------------------------------- decisions ----

POL = {"held": ["Held.App", "Py.*"], "held_override": [], "admin": ["Adm.App"],
       "in_use": {"Acme.Editor": ["editor.exe"]}, "check_after": {"Acme.Editor": "node_version"}}


def test_decide_held_admin_pinned_and_the_override():
    assert upkeep.decide("Held.App", POL, set())["action"] == "held"
    assert upkeep.decide("Py.Thon.3", POL, set())["action"] == "held"
    assert upkeep.decide("Other.App", POL, set(), pinned={"Other.App"})["reason"] == "pinned in winget"
    assert upkeep.decide("Adm.App", POL, set())["action"] == "needs_admin"
    assert upkeep.decide("Held.App", dict(POL, held_override=["Held.App"]), set())["action"] == "upgrade"


def test_decide_in_use_table_folder_check_and_idle():
    assert upkeep.decide("Acme.Editor", POL, {"editor.exe"})["action"] == "deferred"
    assert upkeep.decide("Acme.Editor", POL, None)["reason"] == "deferred: process list unreadable"
    assert upkeep.decide("Acme.Editor", POL, set()) == {"action": "upgrade", "check": "node_version"}
    d = upkeep.decide("Other.App", POL, set(), busy_folder=["other.exe"])
    assert d["action"] == "deferred" and d["processes"] == ["other.exe"]


def test_folder_busy_matches_only_inside_a_specific_install_folder(monkeypatch):
    monkeypatch.setenv("ProgramFiles", r"C:\Program Files")
    un = {"Acme Editor": {"location": r"C:\Program Files\Acme\Editor"},
          "Shared": {"location": r"C:\Program Files"}, "Root": {"location": "C:\\"}, "None": {"location": ""}}
    paths = {r"c:\program files\acme\editor\bin\editor.exe", r"c:\program files\other\x.exe"}
    assert upkeep.folder_busy("Acme Editor", un, paths) == ["editor.exe"]
    for name in ("Shared", "Root", "None", "Missing"):
        assert upkeep.folder_busy(name, un, paths) == []
    assert upkeep.folder_busy("Acme Editor", un, None) == []


def test_selection_default_all_excluded_and_named():
    assert upkeep.is_selected("Any.App", {"selection": {"mode": "all", "excluded": []}})
    assert not upkeep.is_selected("Acme.Player", {"selection": {"mode": "all", "excluded": ["Acme.*"]}})
    named = {"selection": {"mode": "named", "only": ["Acme.Editor"]}}
    assert upkeep.is_selected("Acme.Editor", named) and not upkeep.is_selected("New.App", named)


# --------------------------------------------------------------- winget pass ----

def test_dry_run_lists_and_installs_nothing(world):
    r = upkeep.winget_pass(world, policy=policy(), dry_run=True)
    assert [x["id"] for x in r["would_upgrade"]] == ["Acme.Editor", "Acme.Player", "Globex.Tool"]
    assert world.attempted == [] and r["status"] == "yellow"


def test_install_is_judged_by_effect_one_at_a_time(world):
    world.outcomes = {"Acme.Player": "noop", "Globex.Tool": "fail"}
    r = upkeep.winget_pass(world, policy=policy())
    assert world.attempted == ["Acme.Editor", "Acme.Player", "Globex.Tool"]
    assert [x["id"] for x in r["upgraded"]] == ["Acme.Editor"] and r["upgraded"][0]["to"] == "1.1"
    assert [x["id"] for x in r["not_applied"]] == ["Acme.Player"]
    assert [x["id"] for x in r["failed"]] == ["Globex.Tool"] and r["status"] == "red"
    assert r["learned"]["fails"] == {"Globex.Tool": 1}


def test_two_failed_runs_then_the_package_waits_and_only_names_it_again(world):
    world.outcomes = {"Globex.Tool": "timeout"}
    r1 = upkeep.winget_pass(world, policy=policy())
    assert [x["id"] for x in r1["unconfirmed"]] == ["Globex.Tool"]
    r2 = upkeep.winget_pass(world, policy=policy(), learned=r1["learned"])
    assert r2["learned"]["fails"] == {"Globex.Tool": 2}
    world.attempted.clear()
    r3 = upkeep.winget_pass(world, policy=policy(), learned=r2["learned"])
    assert world.attempted == [] and [x["id"] for x in r3["waiting_fail"]] == ["Globex.Tool"]
    assert r3["status"] == "yellow" and "winget upgrade --id Globex.Tool" in r3["waiting_fail"][0]["command"]
    world.outcomes = {}
    r4 = upkeep.winget_pass(world, policy=policy(), learned=r3["learned"], only=["Globex.Tool"])
    assert world.attempted == ["Globex.Tool"] and r4["learned"]["fails"] == {}


def test_excluded_app_is_listed_and_never_installed(world):
    p = policy(selection={"mode": "all", "excluded": ["Acme.Player"], "only": []})
    r = upkeep.winget_pass(world, policy=p)
    assert "Acme.Player" not in world.attempted
    assert r["not_selected"][0]["id"] == "Acme.Player" and "select --add Acme.Player" in r["not_selected"][0]["command"]
    assert r["status"] == "green"


def test_named_selection_covers_the_named_few_and_no_other(world):
    r = upkeep.winget_pass(world, policy=policy(selection={"mode": "named", "excluded": [], "only": ["Globex.Tool"]}))
    assert world.attempted == ["Globex.Tool"] and len(r["not_selected"]) == 2


def test_held_app_stays_held_even_when_named_in_only(world):
    r = upkeep.winget_pass(world, policy=policy(held=["Acme.Editor"]), only=["Acme.Editor"])
    assert world.attempted == [] and r["held"][0]["id"] == "Acme.Editor"


def test_app_running_from_its_install_folder_is_put_off(world):
    world.uninstall = {"Acme Player": {"location": r"D:\Apps\Acme\Player"}}
    world.proc_paths = [r"D:\Apps\Acme\Player\player.exe"]
    r = upkeep.winget_pass(world, policy=policy())
    assert "Acme.Player" not in world.attempted
    assert r["deferred"][0]["id"] == "Acme.Player" and r["deferred"][0]["processes"] == ["player.exe"]


def test_a_dry_run_does_not_lengthen_a_deferral_streak(world):
    p = policy(in_use={"Acme.Editor": ["editor.exe"]})
    world.running = {"editor.exe"}
    first = upkeep.winget_pass(world, policy=p)
    assert first["deferred"][0]["streak"] == 1
    listed = upkeep.winget_pass(world, policy=p, dry_run=True, prev_winget=first)
    assert listed["deferred"][0]["streak"] == 1
    second = upkeep.winget_pass(world, policy=p, prev_winget=listed)
    assert second["deferred"][0]["streak"] == 2


def test_unparseable_list_is_red_and_installs_nothing(world):
    world.list_text, world.list_rc = "something unexpected", 5
    r = upkeep.winget_pass(world, policy=policy())
    assert r["status"] == "red" and world.attempted == [] and r["raw_head"]


# -------------------------------------------------------------------- modes ----

def test_report_mode_lists_records_and_toasts_without_installing(world):
    rc = upkeep.cmd_daily(args(), world, now=NOW, clock=lambda: NOW)
    last = upkeep.read_json(upkeep.last_path())
    assert rc == 0 and world.attempted == [] and last["dry_run"] is True and last["mode"] == "report"
    assert world.toasts == ["Upkeep 2026-01-15: 3 update(s) ready"]
    report = (upkeep.STATE / "last_report.md").read_text(encoding="utf-8")
    assert "Ready to install: Acme.Editor 1.0 -> 1.1" in report and "installed nothing" in report
    assert not any(c[:2] == ["claude", "update"] for c in world.calls)


def test_auto_mode_installs_on_the_scheduled_run(world):
    set_policy(mode="auto")
    world.claude_after = "2.2.0"
    upkeep.cmd_daily(args(), world, now=NOW, clock=lambda: NOW)
    last = upkeep.read_json(upkeep.last_path())
    assert world.attempted == ["Acme.Editor", "Acme.Player", "Globex.Tool"] and last["dry_run"] is False
    assert last["claude"]["cli"]["outcome"] == "upgraded"
    assert upkeep.read_json(upkeep.STATE / "learned.json") == {"fails": {}}


def test_apply_installs_in_report_mode_and_only_restricts_it(world):
    upkeep.cmd_apply(args(only="Acme.Player"), world, now=NOW, clock=lambda: NOW)
    assert world.attempted == ["Acme.Player"] and world.toasts == []
    assert not any(c[:2] == ["claude", "update"] for c in world.calls)     # a named apply leaves Claude alone
    upkeep.cmd_apply(args(), world, now=NOW + timedelta(minutes=1), clock=lambda: NOW)
    assert world.attempted == ["Acme.Player", "Acme.Editor", "Globex.Tool"]


def test_daily_dry_run_writes_nothing(world, capsys):
    upkeep.cmd_daily(args(dry_run=True), world, now=NOW, clock=lambda: NOW)
    assert not upkeep.last_path().exists() and world.toasts == [] and "Upkeep report" in capsys.readouterr().out


# ------------------------------------------------------------------ schedule ----

def test_valid_schedule_accepts_the_four_periodicities_and_rejects_the_rest():
    assert upkeep.valid_schedule("daily", None, "8:05")[0] == {"every": "daily", "at": "08:05"}
    assert upkeep.valid_schedule("weekly", "Friday", "19:30")[0] == {"every": "weekly", "at": "19:30", "day": "fri"}
    assert upkeep.valid_schedule("manual")[0] == {"every": "manual"}
    for bad in (("hourly", None, "08:00"), ("daily", None, "25:00"), ("daily", None, None), ("weekly", None, "08:00")):
        assert upkeep.valid_schedule(*bad)[0] is None


def test_task_line_for_each_periodicity():
    mk = lambda s: upkeep.schtasks_create_cmd(s, "C:\\py\\pythonw.exe", "C:\\home\\bin\\upkeep.py")  # noqa: E731
    daily = mk({"every": "daily", "at": "08:00"})
    assert daily[:4] == ["schtasks", "/create", "/f", "/tn"] and daily[4] == "upkeep-beta"
    assert daily[daily.index("/tr") + 1] == '"C:\\py\\pythonw.exe" "C:\\home\\bin\\upkeep.py" daily'
    assert daily[-2:] == ["/sc", "daily"] and daily[daily.index("/st") + 1] == "08:00"
    assert mk({"every": "weekdays", "at": "08:00"})[-4:] == ["/sc", "weekly", "/d", "MON,TUE,WED,THU,FRI"]
    assert mk({"every": "weekly", "at": "08:00", "day": "sat"})[-4:] == ["/sc", "weekly", "/d", "SAT"]


def test_stale_alarm_follows_the_chosen_period(world):
    upkeep.write_json(upkeep.last_path(), {"finished": (NOW - timedelta(hours=50)).isoformat()})
    assert upkeep.stale_run_alarm(NOW, {"schedule": None}) is None
    assert upkeep.stale_run_alarm(NOW, {"schedule": {"every": "manual"}}) is None
    assert upkeep.stale_run_alarm(NOW, {"schedule": {"every": "daily", "at": "08:00"}})["status"] == "red"
    assert upkeep.stale_run_alarm(NOW, {"schedule": {"every": "weekdays", "at": "08:00"}}) is None
    assert upkeep.stale_run_alarm(NOW, {"schedule": {"every": "weekly", "at": "08:00", "day": "mon"}}) is None


def test_schedule_dry_run_registers_nothing_and_the_real_one_stores_and_registers(world, capsys):
    assert upkeep.cmd_schedule(args(every="daily", at="07:30", dry_run=True), world) == 0
    assert world.schtasks == [] and upkeep.load_policy()["schedule"] is None
    assert "would run: schtasks /create" in capsys.readouterr().out
    assert upkeep.cmd_schedule(args(every="weekly", day="mon", at="07:30"), world) == 0
    assert upkeep.load_policy()["schedule"] == {"every": "weekly", "at": "07:30", "day": "mon"}
    assert world.task[-2:] == ["/d", "MON"] and upkeep.stable_script().exists()
    assert upkeep.cmd_schedule(args(off=True), world) == 0
    assert world.task is None and upkeep.load_policy()["schedule"] == {"every": "manual"}
    assert upkeep.cmd_schedule(args(every="weekly", at="07:30"), world) == 2


def test_nothing_is_scheduled_by_default(world):
    assert upkeep.load_policy()["schedule"] is None and upkeep.stale_hours(upkeep.load_policy()) is None
    upkeep.cmd_setup(args(), world)
    assert world.schtasks == [] and upkeep.stable_script().exists()


# -------------------------------------------------------------------- select ----

def test_select_remove_add_only_all_and_quiet(world, capsys):
    upkeep.cmd_select(args(remove="Acme.Player,Globex.Tool"))
    assert upkeep.load_policy()["selection"] == {"mode": "all", "excluded": ["Acme.Player", "Globex.Tool"], "only": []}
    upkeep.cmd_select(args(add="Acme.Player"))
    assert upkeep.load_policy()["selection"]["excluded"] == ["Globex.Tool"]
    upkeep.cmd_select(args(only="Acme.Editor"))
    assert upkeep.load_policy()["selection"]["mode"] == "named"
    upkeep.cmd_select(args(add="Acme.Player"))
    assert upkeep.load_policy()["selection"]["only"] == ["Acme.Editor", "Acme.Player"]
    upkeep.cmd_select(args(all=True))
    assert upkeep.load_policy()["selection"] == {"mode": "all", "excluded": [], "only": []}
    upkeep.cmd_select(args(quiet="Old App"))
    assert upkeep.load_policy()["quiet"] == ["Old App"]
    assert "all catalogue apps" in capsys.readouterr().out


def test_adding_a_default_held_app_by_name_lifts_the_hold_and_says_why(world, capsys):
    assert upkeep.decide("Microsoft.VisualStudioCode", upkeep.load_policy(), set())["action"] == "held"
    upkeep.cmd_select(args(add="Microsoft.VisualStudioCode"))
    assert "held by default" in capsys.readouterr().out
    assert upkeep.decide("Microsoft.VisualStudioCode", upkeep.load_policy(), set())["action"] == "upgrade"
    upkeep.cmd_select(args(remove="Microsoft.VisualStudioCode"))
    assert upkeep.load_policy()["held_override"] == []


def test_the_three_default_held_apps_and_report_mode_ship_as_defaults(world):
    p = upkeep.load_policy()
    assert p["mode"] == "report" and p["selection"]["mode"] == "all" and p["admin"] == []
    for pid in ("Microsoft.VisualStudioCode", "Anthropic.Claude", "Python.Python.3.13"):
        assert upkeep.decide(pid, p, set())["action"] == "held"


# ---------------------------------------------------------------------- apps ----

KNOWN = {"driver_words": ["driver"], "manual": [{"name": "python 3*", "by": "you"}],
         "self_updating": [{"name": "zeta browser*", "by": "itself"}, {"publisher": "maker corp*", "by": "Maker Update"}]}


def test_unmatched_kinds():
    k = lambda n, pub="": upkeep.unmatched_kind(n, {"publisher": pub}, KNOWN)[0]  # noqa: E731
    assert k("Photo Page", "Vendor\\Chrome") == "web app"
    assert k("Audio Driver") == "driver"
    assert k("Python 3.12.1") == "manual"
    assert k("Zeta Browser") == "updates itself" and k("Helper Service", "Maker Corp., Ltd.") == "updates itself"
    assert k("Signing Tool", "Some Vendor") == "no known updater"


def test_classify_installed_splits_the_winget_list(world):
    rows = [{"name": "Acme Editor", "id": "Acme.Editor", "version": "1.0", "source": "winget"},
            {"name": "Acme Editor", "id": "Acme.Editor", "version": "1.0", "source": "winget"},
            {"name": "VS Code", "id": "Microsoft.VisualStudioCode", "version": "1.0", "source": "winget"},
            {"name": "Reader", "id": "XP123ABC", "version": "9", "source": "msstore"},
            {"name": "Part", "id": "MSIX\\Vendor.Part_1.0_x64__abc", "version": "1.0", "source": ""},
            {"name": "Signing Tool", "id": "ARP\\Machine\\X64\\{GUID}", "version": "2.0", "source": ""}]
    inv = upkeep.classify_installed(rows, policy(), {"fails": {"Acme.Editor": 2}}, {}, KNOWN)
    assert [a["id"] for a in inv["catalogue"]] == ["Acme.Editor", "Microsoft.VisualStudioCode"]
    assert inv["catalogue"][0]["class"].startswith("failed 2 runs") and inv["catalogue"][1]["class"].startswith("held")
    assert len(inv["store"]) == 1 and inv["components"] == 1
    assert inv["unmatched"][0]["name"] == "Signing Tool" and inv["unmatched"][0]["kind"] == "no known updater"


def test_apps_match_prints_an_exact_name_hit_and_installs_nothing(world, capsys):
    world.installed_rows = [("Signing Tool", "ARP\\Machine\\X64\\{A}", "2.0", ""),
                            ("Other Tool", "ARP\\Machine\\X64\\{B}", "1.0", "")]
    world.search = {"Signing Tool": [("Signing Tool", "Vendor.SigningTool", "2.1")]}
    assert upkeep.cmd_apps(args(match=True), world) == 0
    text = capsys.readouterr().out
    assert "catalogue id Vendor.SigningTool, version 2.1" in text and "1 of 2 unmatched" in text
    assert not any("install" in c for c in world.calls)


# -------------------------------------------------------------------- checks ----

def ctx(world, **kw):
    c = upkeep.make_ctx(world, {}, NOW)
    c.update(kw)
    return c


def test_pending_reboot(monkeypatch, world):
    item = {"id": "pending_reboot"}
    monkeypatch.setattr(upkeep, "reg_key_exists", lambda k: False)
    assert upkeep.check_pending_reboot(item, ctx(world))["status"] == "green"
    monkeypatch.setattr(upkeep, "reg_key_exists", lambda k: "RebootRequired" in k)
    assert upkeep.check_pending_reboot(item, ctx(world))["status"] == "yellow"
    monkeypatch.setattr(upkeep, "reg_key_exists", lambda k: None)
    assert upkeep.check_pending_reboot(item, ctx(world))["status"] == "unknown"


@pytest.mark.parametrize("free,status", [(50, "green"), (8, "yellow"), (3, "red")])
def test_disk_free_thresholds(monkeypatch, world, free, status):
    usage = namedtuple("usage", "total used free")
    monkeypatch.setattr(upkeep, "disk_usage", lambda d: usage(100 * 2 ** 30, 0, free * 2 ** 30))
    assert upkeep.check_disk_free({"id": "disk_free", "drive": "C:"}, ctx(world))["status"] == status


def test_windows_update_count_and_unknown(world):
    item = {"id": "windows_update"}
    assert upkeep.check_windows_update(item, ctx(world))["status"] == "green"
    world.wu_count = 3
    r = upkeep.check_windows_update(item, ctx(world))
    assert r["status"] == "yellow" and r["value"] == "3"
    world.wu_count = None
    assert upkeep.check_windows_update(item, ctx(world))["status"] == "unknown"


def test_pip_outdated_reports_a_count_and_never_updates(world):
    world.pip = json.dumps([{"name": "a"}, {"name": "b"}])
    r = upkeep.check_pip_outdated({"id": "pip_outdated"}, ctx(world))
    assert r["status"] == "green" and r["value"] == "2"
    assert not any("install" in c for c in world.calls)
    world.pip = "not json"
    assert upkeep.check_pip_outdated({"id": "pip_outdated"}, ctx(world))["status"] == "unknown"


def test_python_version_is_yellow_below_the_floor(world):
    assert upkeep.check_python_version({"id": "python", "yellow_below": "3.0"}, ctx(world))["status"] == "green"
    r = upkeep.check_python_version({"id": "python", "yellow_below": "99.0"}, ctx(world))
    assert r["status"] == "yellow" and "update Python" in r["action"]


def stale(world, rows, uninstall, seen=None, quiet=()):
    upkeep.write_json(upkeep.KNOWN, KNOWN)
    c = ctx(world, installed_rows=[{"name": n, "id": "ARP\\Machine\\X64\\" + n, "version": v, "source": ""}
                                   for n, v in rows],
            uninstall=uninstall, apps_seen=seen or {}, policy={"quiet": list(quiet)})
    return upkeep.check_unmatched_stale({"id": "unmatched_stale", "stale_days": 180}, c)


def test_unmatched_stale_seeds_the_clock_from_the_install_date_only_without_an_updater(world):
    old = (NOW - timedelta(days=400)).strftime("%Y%m%d")
    un = {"Signing Tool": {"date": old, "publisher": "V"}, "Zeta Browser": {"date": old, "publisher": "Z"},
          "Audio Driver": {"date": old, "publisher": "R"}, "Photo Page": {"date": old, "publisher": "V\\Chrome"}}
    r = stale(world, [("Signing Tool", "2.0"), ("Zeta Browser", "5"), ("Audio Driver", "1"), ("Photo Page", "1")], un)
    assert r["status"] == "yellow" and "Signing Tool (400 days" in r["detail"] and "Zeta" not in r["detail"]
    assert set(r["apps_seen"]) == {"Signing Tool", "Zeta Browser"}       # drivers and web apps are left out
    assert stale(world, [("Signing Tool", "2.0")], un, quiet=["signing tool"])["status"] == "green"


def test_unmatched_stale_counts_from_first_sight_and_a_version_change_resets(world):
    seen = {"Zeta Browser": {"version": "5", "since": (NOW - timedelta(days=200)).isoformat()}}
    assert stale(world, [("Zeta Browser", "5")], {}, seen)["status"] == "yellow"
    r = stale(world, [("Zeta Browser", "6")], {}, seen)
    assert r["status"] == "green" and r["apps_seen"]["Zeta Browser"]["since"] == NOW.isoformat(timespec="seconds")


# ----------------------------------------------------- home, doctor, feedback ----

def test_ensure_home_copies_defaults_once_and_never_overwrites(world):
    assert upkeep.load_policy()["mode"] == "report" and upkeep.KNOWN.exists()
    set_policy(mode="auto")
    upkeep.ensure_home()
    assert upkeep.load_policy()["mode"] == "auto"


def test_stable_copy_is_refreshed_when_this_script_differs(world):
    upkeep.install_stable_copy()
    upkeep.stable_script().write_text("old", encoding="utf-8")
    upkeep.ensure_home()
    assert upkeep.stable_script().read_bytes() == Path(upkeep.__file__).read_bytes()


def test_redact_and_diag_hide_the_user_and_machine(world, monkeypatch, capsys):
    monkeypatch.setenv("USERPROFILE", r"C:\Users\testperson")
    monkeypatch.setenv("USERNAME", "testperson")
    monkeypatch.setenv("COMPUTERNAME", "TESTBOX")
    assert upkeep.redact(r"c:/users/TESTPERSON/x on testbox") == "%USERPROFILE%/x on %COMPUTERNAME%"
    monkeypatch.setattr(upkeep, "HOME", Path(r"C:\Users\testperson\AppData\Local\upkeep-beta"))
    monkeypatch.setattr(upkeep, "_doctor_rows", lambda r: [("home folder", True, str(upkeep.HOME))])
    upkeep.cmd_diag(args(), world)
    text = capsys.readouterr().out
    assert "testperson" not in text.lower() and "%USERPROFILE%" in text


def test_doctor_reports_winget_and_the_missing_task(world, monkeypatch, capsys):
    monkeypatch.setattr(upkeep, "POWERSHELL", sys.executable)     # doctor only checks that the file exists
    assert upkeep.cmd_doctor(args(), world) == 0
    text = capsys.readouterr().out
    assert "winget" in text and "1.9.0" in text and "not chosen yet" in text


def test_toast_puts_ready_updates_first_and_stays_within_600(world):
    w = upkeep.winget_pass(world, policy=policy(), dry_run=True)
    rec = {"date": "2026-01-15", "counts": {"red": 0, "yellow": 1, "green": 0, "unknown": 0}, "items": [], "winget": w}
    title, body, sticky = upkeep.build_toast(rec, "D:\\x\\last_report.md")
    assert title.endswith("3 update(s) ready") and "apply" in body and len(body) <= 600 and not sticky

