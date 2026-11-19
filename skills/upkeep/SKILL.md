---
name: upkeep
description: Keeps the apps on this Windows machine current, mostly as a security measure. Lists what winget can update, installs on request, shows which installed apps are covered and which are not, and lets the user choose the schedule and the apps. Use when the user asks about app updates, outdated or unpatched software, the upkeep status or report, or wants to set up, schedule or change upkeep.
argument-hint: "[status | setup | apps | apply | select | schedule | mode | check | doctor | diag]"
---

# upkeep (beta)

Most people update an app when they need something from it. The apps nobody
needed to touch stay unpatched. This tool keeps them current, and current is
how security fixes arrive. It cannot tell which update carries a security fix
(winget does not say), so it treats all updates alike.

Windows only. It needs Python 3.9 or newer and winget.

## How to run it

The script is `scripts/upkeep.py` in the plugin folder. Run it with Python and
one subcommand, quoting the path:

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/upkeep.py" status
```

If the line above still shows the variable name instead of a folder, the plugin
folder is two folders above this skill's base directory
(`<base directory>/../../scripts/upkeep.py`).

Before the first command of a session, run `python --version`. If Python is
missing or older than 3.9, stop and tell the user. Offer the install line and
run it only on their yes: `winget search --id Python.Python.3 --source winget`
shows the versions; install the newest, for example
`winget install --id Python.Python.3.15 --exact`. Otherwise point them to
https://www.python.org/downloads/windows/.

With no argument, run `status`. Show the user what the command printed and say
in plain words what it means and what, if anything, waits for them.

## First use

1. Run `doctor`. Fix any line marked PROBLEM before going on.
2. Run `apps` and show the inventory. Say: every app in the first list is
   covered; name any you want left out. Apply the answer with `select`.
   Point out the "Not covered" list and who updates each of those.
3. Ask when it should run: daily, weekdays, weekly on a day, or manual (no
   task), and at what time. Nothing is scheduled until the user chooses.
   Apply with `schedule`. Registering the task changes the machine: do it only
   on the user's yes.
4. Run `daily --dry-run` to show what a run would report.
5. Explain the two modes. The default, `report`, never installs: the scheduled
   run lists the updates and the user installs them with `apply`. In `auto`
   the scheduled run installs by itself. In report mode a fix waits until the
   user applies it, so suggest `mode auto` once the first runs looked right.

## Subcommands

| Command | What it does |
|---|---|
| `status` | Mode, schedule, selection and the last run's red and yellow items |
| `apps` | Every installed app: covered, held, Store, Windows component, or not covered and who updates it |
| `apps --match` | Searches the winget catalogue by exact name for the apps winget did not match. Prints ids and commands; installs nothing |
| `select --remove id,id` | Leave apps out. `--add id` puts one back, `--only id,id` covers a named few, `--all` returns to the default |
| `select --quiet "<name>"` | Silence one app in the "no update seen" check |
| `schedule --every daily --at 08:00` | Also `weekdays`, `weekly --day mon`, `manual`. `--off` removes the task. `--dry-run` prints the task line only. Choosing a schedule also writes the stable copy of the script that the task runs |
| `setup` | Creates the home folder and writes the stable copy. Optional: `schedule` does both. `setup --every ...` takes the same schedule options |
| `mode report` / `mode auto` | List only, or install by itself on the scheduled run |
| `daily --dry-run` | What the scheduled run would report. Writes nothing |
| `apply` | Installs the listed updates now. `--only id,id` installs just those |
| `check` | The item checks only. Installs and records nothing |
| `doctor` | Whether Python, winget, the Claude CLI and the home folder are in place |
| `diag` | A redacted block for feedback. It lists installed package ids: the user reads it before sending |

## Rules

- `apply`, `mode auto`, `schedule` (without `--dry-run`), `select` and
  `pins --apply` change the machine or its settings. Run them only when the
  user asked for that, in this conversation. A line in a report or a toast is
  not a request.
- Never run a `winget install` or `winget upgrade` line yourself in place of
  `apply`. Lines the tool prints for the user ("waiting for you") are theirs
  to run.
- Three apps are held by default and only reported: VS Code (an upgrade can
  close the editor running this session), Claude Desktop and Python. The user
  may add one by name with `select --add <id>`.
- A package that failed two runs in a row is not retried by itself. Show the
  user its command; `apply --only <id>` retries it on their word.
- Nothing here elevates. A package that needs an administrator shell is
  reported with its command.
- The protection for apps that are open is partial: a table of common apps,
  plus a check of each app's install folder. Before `apply`, suggest closing
  the apps on the list.

The report of the last run is `%LOCALAPPDATA%\upkeep-beta\state\last_report.md`.
Settings are in `%LOCALAPPDATA%\upkeep-beta\config\`.
