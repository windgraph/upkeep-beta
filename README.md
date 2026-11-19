# upkeep-beta

A Claude Code plugin that keeps the apps on a Windows machine current.

**This is a beta.** It has run daily on one machine. You are among the first to
run it on another. Read "What is not tested yet" before you switch it to
automatic.

## Why

Developers usually update an app when they need something from it. The apps
nobody needed to touch stay at the version they were installed with, unpatched.
That is mostly a security problem, not a features problem.

upkeep-beta keeps those apps current, because current is how security fixes
arrive. Three limits, stated up front:

- It cannot tell which update carries a security fix. winget does not say, so
  all updates are treated alike.
- It only updates apps that winget can match to its catalogue. The rest are
  named, with who updates each one, and watched for going stale.
- In the default mode it installs nothing by itself. A fix waits until you
  apply it.

## What it does

- **Lists** the updates winget has for your apps, one line each.
- **Installs** them one at a time, when you say so (`apply`) or on a schedule
  you choose (`auto` mode). Each install is judged by effect: the version is
  read before and after, and exit codes are only evidence.
- **Shows the inventory** (`apps`): which installed apps it covers, which it
  holds, and which it cannot reach.
- **Checks** a few things that never install anything: Windows Update count,
  restart pending, free disk space, the Claude Code CLI and its plugins, the
  VS Code Claude extension, Python, outdated Python packages (a count), Node,
  Git, the NVIDIA driver.
- **Sends one toast** per scheduled run and writes a report.

## What it never does

- It never elevates. A package that needs an administrator shell is reported
  with the command for you to run.
- It never updates VS Code, Claude Desktop or Python unless you add them by
  name. An upgrade of VS Code can close the editor your Claude Code session
  runs in.
- It never updates Python packages. It reports how many are outdated.
- It never installs over an app winget did not match.
- It sends nothing anywhere. The only network traffic is what `winget`,
  `claude` and `pip` do themselves.

## Requirements

- Windows 10 or 11 with winget ("App Installer" from the Microsoft Store)
- Python 3.9 or newer on PATH
- Claude Code

## Install

In Claude Code:

```
/plugin marketplace add windgraph/upkeep-beta
/plugin install upkeep-beta@upkeep-beta
```

The repository is private: your GitHub account must have been invited, and git
on your machine must be signed in to it.

## First run

Type:

```
/upkeep-beta:upkeep setup
```

Claude walks you through it:

1. `doctor` checks that Python, winget and the Claude CLI are in place.
2. `apps` shows the inventory. By default every catalogue app is covered,
   including ones you install later. Name any you want left out.
3. You choose when it runs: daily, weekdays, weekly on a day, or manual, and
   at what time. Nothing is scheduled until you choose.
4. A dry run shows what a report looks like.

## Modes

| Mode | The scheduled run | You |
|---|---|---|
| `report` (default) | Lists updates, runs the checks, sends the toast. Installs nothing | Install with `/upkeep-beta:upkeep apply` when you are at the desk |
| `auto` | Installs the updates of the covered apps by itself | Read the toast |

Start in `report`. Once the first runs look right, `auto` is what keeps fixes
from waiting on you: `/upkeep-beta:upkeep mode auto`.

## Commands

All through the skill, for example `/upkeep-beta:upkeep apps`.

| Command | What it does |
|---|---|
| `status` | Mode, schedule, selection and the last run's red and yellow items |
| `apps` | The inventory |
| `apps --match` | Searches the winget catalogue by exact name for apps winget did not match |
| `select --remove id,id` | Leave apps out. Also `--add`, `--only`, `--all`, `--quiet "<name>"` |
| `schedule --every weekly --day mon --at 08:00` | Also `daily`, `weekdays`, `manual`, `--off` |
| `mode report` or `mode auto` | Switch the mode |
| `apply` | Install the listed updates now. `--only id,id` for some |
| `check` | The checks only |
| `doctor` | Is everything in place |
| `diag` | A redacted block for feedback |

## When an install does not go through

- **The app is open.** The update is put off to the next run. After three runs
  in a row it waits for you, with the command.
- **It failed or timed out.** After two runs in a row it is not retried by
  itself. The report shows the command; `apply --only <id>` retries it.
- **It needs an administrator.** It is reported with the command to run in an
  administrator PowerShell.

## Where things are

| What | Where |
|---|---|
| Settings | `%LOCALAPPDATA%\upkeep-beta\config\` |
| Last report | `%LOCALAPPDATA%\upkeep-beta\state\last_report.md` |
| Run records | `%LOCALAPPDATA%\upkeep-beta\state\runs\` |
| The script the task runs | `%LOCALAPPDATA%\upkeep-beta\bin\upkeep.py` |
| Scheduled task | `upkeep-beta`, in Task Scheduler, under your user |

Your settings are never overwritten by a plugin update.

## What is not tested yet

- **Other Windows languages.** winget's tables are read by column position, so
  the language should not matter, but only Portuguese and English output have
  been seen for real. If the report says the upgrade list could not be parsed,
  send the `diag` block.
- **Apps that are open.** The protection is partial. It is a table of about 35
  common apps written from general knowledge, plus a check of each app's
  install folder. On the one machine measured, the registry gave that folder
  for 16 of 45 apps. An app outside both can be upgraded while open. In `report` mode you are at the
  desk when you apply, so close what matters first.
- **The scheduled task on a second machine.** Registering it is covered by
  tests with a fake `schtasks`, not yet by a second real machine.
- **The "who updates it" table** for apps outside winget is a starting list,
  not verified app by app. Correct `known_apps.json` freely.
- **A machine without winget.** `doctor` says so; nothing else works.

## Feedback

Open an issue on this repository. Paste the output of
`/upkeep-beta:upkeep diag`. It replaces your user name, profile path and
machine name, but it does list app names and package ids from the last run:
read it before you send it.

## Uninstall

```
/upkeep-beta:upkeep schedule --off
/plugin uninstall upkeep-beta@upkeep-beta
```

Then delete the folder `%LOCALAPPDATA%\upkeep-beta`.

## Development

```
python -m pytest tests -q
```

The tests use a fake runner and temporary folders. They read and change
nothing on the machine. Standard library only; pytest for the tests.

## Licence

MIT. See `LICENSE`.
