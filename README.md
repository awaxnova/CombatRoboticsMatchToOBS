# CombatMatchManager

Portable Windows PowerShell 5.1 WinForms app for managing combat robotics matches and writing OBS text source files.

## Requirements
- Windows 10/11
- Windows PowerShell 5.1
- No install needed

## Launch
1. Place files in a folder (or USB drive).
2. Double-click `RunApp.bat`.

The app auto-creates these folders if missing:
- `divisions/`
- `data/`
- `output/`
- `logs/`

## CSV Input
Put one CSV per division in `divisions/`.
- Division name = filename without `.csv`
- Required column: `BotName`
- Optional columns: `Team`, `Seed`, `Driver`, `Notes`, `Arena`, `Division`
- Blank `BotName` rows are ignored

Example:
```csv
BotName,Team,Seed
Rust Bucket,Team Alpha,1
Gear Grinder,Team Beta,2
Tiny Menace,Team Gamma,3
```

## Match Management
- Select a division
- Multi-select bots and create a match (2+ bots required)
- Reorder, edit, delete, mark done/queued
- Set one match live globally

Keyboard shortcuts:
- `Enter` set selected match live (when match list is focused)
- `Ctrl+Up` move match up
- `Ctrl+Down` move match down
- `Delete` delete selected match
- `Ctrl+L` clear live
- `Ctrl+R` reload divisions
- `Ctrl+N` create match from selected bots

## OBS Output Files
When a live match is set, files in `output/` are overwritten:
- `current_division.txt`
- `current_match.txt`
- `current_bot_1.txt` ... `current_bot_6.txt`
- `current_status.txt`
- `current_vs.txt`

Rules:
- Unused bot files are blank
- `current_status.txt` = `LIVE` when a match is live, otherwise blank
- `current_vs.txt` = `vs` for 2 bots, `Battle Royale` for 3+ bots, blank when no live match

## Persistence
- Match plans: `data/matchPlans.json`
- UI/live state: `data/appState.json`

State is saved immediately after create/edit/delete/reorder/live/status changes.

## Logs
- Log file: `logs/app.log`
- Includes startup, CSV load issues, persistence issues, match actions, and OBS output write errors.
