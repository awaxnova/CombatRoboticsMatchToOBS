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
- Preferred headered format includes `BotName`
- Headerless rows are also supported (first column = bot name, second = team)
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
- Manage the bot list directly (add, rename, remove)
- Reorder, edit, delete (single or multi-select), mark done/queued
- Set one match live globally

Keyboard shortcuts:
- `Enter` set selected match live (when match list is focused)
- `Ctrl+Up` move match up
- `Ctrl+Down` move match down
- `Delete` delete selected match(es)
- `Ctrl+L` clear live
- `Ctrl+R` reload divisions
- `Ctrl+N` create match from selected bots
- `Ctrl+B` add bot
- `Ctrl+E` edit selected bot
- `Ctrl+I` import matches from Challonge SVG
- `Delete` removes selected bots when the bot list is focused

## Challonge SVG Import
- Use `Import Challonge SVG` in the match actions (or `Ctrl+I`)
- Paste a printer-friendly Challonge SVG URL (for example `https://challonge.com/part_nyk_beetles.svg`) or a local `.svg` file path
- The app parses bracket matches and appends found 2+ bot matches into the selected division as `Queued`
- Imported bot names are auto-added to that division's bot list
- Duplicate bot names are deduped automatically (case/spacing-insensitive)
- Placeholders (`TBD`, `BYE`, `Winner of ...`, `Loser of ...`) are skipped

## OBS Output Files
When a live match is set, files in `output/` are overwritten:
- `current_division.txt`
- `current_match.txt`
- `current_bot_1.txt` ... `current_bot_N.txt` (one per bot; minimum 6 files maintained)
- `current_status.txt`
- `current_vs.txt`

Rules:
- Unused bot files are blank
- `current_status.txt` = `LIVE` when a match is live, otherwise blank
- `current_vs.txt` = `vs` for 2 bots, `Battle Royale` for 3+ bots, blank when no live match
- Marking a live match `Done` clears live OBS output files

## Persistence
- Match plans: `data/matchPlans.json`
- Bot roster overrides: `data/botOverrides.json`
- UI/live state: `data/appState.json`

State is saved immediately after bot or match create/edit/delete/reorder/live/status changes.

## Logs
- Log file: `logs/app.log`
- Includes startup, CSV load issues, persistence issues, match actions, and OBS output write errors.
