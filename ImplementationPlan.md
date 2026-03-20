Implementable Spec: Portable Windows Match Manager for OBS Text Files

Build a portable Windows desktop application for running combat robotics matches at an event. The app must be designed to run on a Windows PC where no new software can be installed. For that reason, implement it as a single PowerShell 5.1 WinForms application with supporting files only. Do not depend on Python, Node, .NET SDK installation, external modules, databases, or web services.

Primary purpose

The app ingests multiple CSV roster files, where each CSV represents a division, typically associated with a specific arena or site. The user can create and arrange matches within each division, then mark one match as live. When a match is marked live, the app must overwrite a set of separate output text files that OBS can read as text sources.

Platform constraints

Target:

Windows 10/11

Windows PowerShell 5.1

WinForms GUI

No installers

No external packages

App should run from a folder or USB drive

Deliverables should be:

CombatMatchManager.ps1

RunApp.bat

supporting folders created automatically if missing

High-level behavior

The app must:

Load one or more CSV roster files from a divisions folder

Treat each CSV file as one division

Display divisions in the GUI

Show bots for the selected division

Allow the user to create matches from selected bots

Support matches with 2 or more bots

Allow matches to be reordered

Allow the user to set one match as live

When a match is live, overwrite separate OBS text output files

Persist match plans and state locally so the event can be resumed after restart

Folder structure

The app should expect or create this structure relative to the script location:

CombatMatchManager/
  CombatMatchManager.ps1
  RunApp.bat
  divisions/
  data/
  output/
  logs/
Purpose of folders

divisions/ contains roster CSV files, one per division

data/ contains saved match plans and app state

output/ contains OBS-readable text files

logs/ contains plain text logs for troubleshooting

Launch behavior

RunApp.bat should launch the script with execution-policy bypass for that run only.

Example behavior:

determine script directory

launch powershell.exe

run CombatMatchManager.ps1

The app must not require command-line arguments.

CSV input specification

Each CSV file in divisions/ is one division.

Division naming

Use the CSV filename without extension as the default division name.

Examples:

Beetles Arena 1.csv → division name Beetles Arena 1

Ants South.csv → division name Ants South

Required columns

At minimum, support:

BotName

Optional columns

If present, preserve and display:

Team

Seed

Driver

Notes

Arena

Division

Unknown additional columns should not break the app. Preserve them in memory if feasible, but only BotName is required for function.

CSV assumptions

First row is header

UTF-8 or ANSI text should be handled as reasonably as PowerShell allows

Blank bot rows should be ignored

Leading/trailing whitespace in names should be trimmed

Internal data model

Use simple PowerShell custom objects or hashtables.

Division object

Each division should contain:

Id stable internal identifier

Name

SourceCsvPath

Bots collection

Matches collection

Bot object

Each bot should contain:

Id stable internal identifier

BotName

Team optional

Seed optional

Driver optional

Notes optional

RawData optional hashtable for extra columns

Match object

Each match should contain:

Id stable internal identifier

DivisionId

MatchNumber display/order number

Bots array of bot references or names

Status one of Queued, Live, Done

ArenaLabel optional, default from division name

Notes optional

CreatedAt

UpdatedAt

App state object

Persist:

divisions loaded

matches per division

current selected division

current live match id

optionally current selected match in UI

Persistence

Use JSON files under data/.

Required files

data\matchPlans.json

data\appState.json

Persistence rules

On startup, load divisions from CSVs

Then load saved match plans if present

Match plans are separate from source CSVs

Do not modify the original roster CSVs

Save after every meaningful change:

create match

edit match

reorder match

delete match

set live

mark done

If JSON is missing or invalid, app should continue with empty match plans and show a warning.

OBS output files

When a match is set live, overwrite separate text files in the output/ folder.

Required output files

current_division.txt

current_match.txt

current_bot_1.txt

current_bot_2.txt

Additional output files

Support up to at least 6 bots per match by also writing:

current_bot_3.txt

current_bot_4.txt

current_bot_5.txt

current_bot_6.txt

If fewer bots exist in the live match, overwrite unused bot files with an empty string.

Strongly recommended extra files

Also write:

current_status.txt

current_vs.txt

Where:

current_status.txt contains LIVE

current_vs.txt contains vs for 2-bot matches, otherwise empty or Battle Royale

Match text rules

current_match.txt should contain a user-friendly label such as:

Match 1

Match 12

Division text rules

current_division.txt should contain the selected division or arena label.

Bot file contents

Each bot file should contain only the bot display name unless a formatting option is added later.

Examples:

current_bot_1.txt → Rust Bucket

current_bot_2.txt → Gear Grinder

On startup

Do not automatically set a live match unless one was persisted as live in saved state and still exists.

On clearing live match

If user clears the live match, overwrite all current output files with empty strings except optionally current_status.txt, which may also be blank.

User interface specification

Use a single main WinForms window.

Main layout

Use a simple three-column layout:

Left panel: divisions

Displays all loaded divisions in a list.
Actions:

select division

reload divisions button

Middle panel: bots in division

Displays bots in selected division.
Supports:

multi-select

create match from selected bots

optional search/filter text box

Right panel: matches in division

Displays matches for selected division in order.
Each match should show:

match number

bot names

status

Actions:

set live

mark done

mark queued

move up

move down

edit match

delete match

duplicate match optional

Bottom panel: live preview

Show exactly what is currently being written to OBS files:

division

match

bot 1

bot 2

bot 3 etc. as needed

status

UX behavior details
Division selection

When user selects a division:

load its bots into the bots list

load its matches into the matches list

preserve unsaved changes by saving immediately on prior actions

Creating a match

User can select 2 or more bots from the division bot list and click Create Match.

Behavior:

create one queued match

bots appear in selected order if determinable, otherwise list order

append to end of match list

auto-assign match number based on display order

save immediately

Match numbering

Match numbers are display-order based, not fixed immutable ids.
After reorder, renumber visible matches sequentially:

Match 1

Match 2

Match 3

Reordering

User can reorder matches with:

Move Up

Move Down

After reorder:

renumber match labels

save immediately

Set live

When user clicks Set Live on a match:

all other matches in all divisions should lose Live status

selected match status becomes Live

write OBS output files

save immediately

refresh live preview

Mark done

When user marks match done:

set status to Done

if it was live, do not necessarily clear live automatically unless explicitly chosen; preferred behavior is:

mark done

keep it live until another match is set live or user clears live

save immediately

Clear live

Provide a Clear Live button.
Behavior:

no match remains live

overwrite current output files with blanks

save immediately

Edit match

Allow editing the bots assigned to an existing match.
Simple version:

open dialog with list of bots from that division

allow reselecting participants

require at least 2 bots

save immediately after confirmation

Delete match

Delete selected match after confirmation prompt.
After delete:

renumber

save immediately

Keyboard controls

Implement these keyboard shortcuts at minimum:

Up/Down arrows: navigate lists

Enter: set selected match live when match list has focus

Ctrl+Up: move selected match up

Ctrl+Down: move selected match down

Delete: delete selected match

Ctrl+L: clear live

Ctrl+R: reload divisions

Ctrl+N: create match from selected bots

Keyboard shortcuts should work without interfering with text entry when a search box has focus.

Validation and warnings
Required validation

cannot create a match with fewer than 2 bots

cannot set live if no match is selected

invalid or empty CSV should not crash the app

division with no bots should still appear but with warning or empty state

Duplicate bot usage

Allow a bot to appear in multiple queued matches, but show a non-blocking warning in UI if the same bot already appears in another match in that division.

Missing BotName

If a row lacks a usable BotName, skip it.

Logging

Write simple logs to logs\app.log.

Log events like:

startup

division load

CSV parse errors

match create/edit/delete

reorder

set live

clear live

persistence load/save failures

OBS output write failures

Use plain text with timestamps.

Error handling

The app must fail gracefully.

If a CSV cannot be read

log the error

show a warning

continue loading other divisions

If output files cannot be written

log the error

show warning in status bar or message box

app remains usable

If JSON save fails

log the error

notify user

Display formatting
Bot display name

Use:

BotName only for v1

Optional future enhancement:

BotName (Team)

Match display in list

For a 2-bot match:

Match 4 | Rust Bucket vs Gear Grinder | Queued

For 3+ bots:

Match 7 | Rust Bucket vs Gear Grinder vs Tiny Menace | Queued

Or use a cleaner separator if it fits WinForms better.

Startup sequence

On startup:

determine app root folder

ensure required folders exist

initialize log

load CSV divisions

load saved match plans

reconcile saved divisions with current CSV-derived divisions

load saved app state

populate GUI

if saved live match still exists, write output files to reflect it

otherwise leave output files unchanged or optionally clear them; preferred v1 behavior: clear them on startup unless saved live match is valid

Reconciliation rules

Since CSVs may change between runs:

division identity should primarily follow filename

if a division CSV disappears, keep its saved data only if you explicitly support orphaned divisions; for v1, simplest is:

only load divisions currently present in divisions/

ignore stale saved divisions not backed by a current CSV

if bot roster changes, existing saved matches may reference bot names no longer present

keep those matches but show them as-is

optionally flag missing bots in UI

Non-goals for v1

Do not implement:

bracket generation

automatic tournament seeding logic

drag-and-drop reorder

network sync

web UI

database

OBS websocket integration

editing source CSV files

multiple simultaneous live matches

installer or EXE packaging

Suggested implementation structure

Within the PowerShell script, separate logic into functions:

Core areas

path initialization

logging

CSV loading

JSON persistence

OBS output writing

division/match state helpers

GUI construction

event handlers

UI refresh functions

Suggested function names

Use names like:

Initialize-AppPaths

Write-Log

Load-DivisionCsvFiles

Load-MatchPlans

Save-MatchPlans

Load-AppState

Save-AppState

Write-ObsOutputFiles

Clear-ObsOutputFiles

Renumber-Matches

Set-LiveMatch

Refresh-DivisionList

Refresh-BotList

Refresh-MatchList

Refresh-LivePreview

Show-EditMatchDialog

Acceptance criteria

The build is acceptable when all of the following work:

The app launches from RunApp.bat on a normal Windows machine without installing anything.

CSV files placed in divisions/ appear as divisions in the app.

Selecting a division shows its bots.

User can select 2 or more bots and create a match.

Matches can be reordered up and down.

Match list renumbers properly after reorder or delete.

User can set a match live.

Setting a match live overwrites:

current_division.txt

current_match.txt

current_bot_1.txt

current_bot_2.txt

and clears unused bot files

Live preview reflects actual written content.

Match plans survive restart using JSON persistence.

App handles bad CSVs without crashing.

App logs important operations to logs\app.log.

Example input CSV
BotName,Team,Seed
Rust Bucket,Team Alpha,1
Gear Grinder,Team Beta,2
Tiny Menace,Team Gamma,3
Servo Smasher,Team Delta,4
Example output files for a 2-bot live match

output\current_division.txt

Beetles Arena 1

output\current_match.txt

Match 3

output\current_bot_1.txt

Rust Bucket

output\current_bot_2.txt

Gear Grinder

output\current_bot_3.txt

output\current_status.txt

LIVE

output\current_vs.txt

vs
Example output files for a 4-bot live match

output\current_match.txt

Match 8

output\current_bot_1.txt

Rust Bucket

output\current_bot_2.txt

Gear Grinder

output\current_bot_3.txt

Tiny Menace

output\current_bot_4.txt

Servo Smasher

output\current_vs.txt

Battle Royale
Implementation style requirements

Prefer clarity and maintainability over cleverness

Keep all code in one main script for v1 unless a second helper ps1 is truly useful

Comment major sections

Avoid advanced PowerShell patterns that reduce compatibility

Use WinForms controls that are standard and dependable

Keep the UI functional even on modest-resolution displays

Codex build prompt

Use this as the direct prompt to Codex:

Build a portable Windows PowerShell 5.1 WinForms desktop app called CombatMatchManager.

Requirements:
- No external dependencies
- No installers
- Must run from RunApp.bat
- Main script is CombatMatchManager.ps1
- App root contains folders: divisions, data, output, logs
- Load all CSV files from divisions folder
- Each CSV is one division
- Use filename without extension as division name
- Required CSV column: BotName
- Optional columns: Team, Seed, Driver, Notes, Arena, Division
- Ignore blank bot rows
- GUI must have:
  - left panel divisions list
  - middle panel division bot roster with multi-select
  - right panel match list
  - bottom live preview
- Buttons/actions:
  - reload divisions
  - create match from selected bots (2 or more bots required)
  - set live
  - clear live
  - move up
  - move down
  - edit match
  - delete match
  - mark done
  - mark queued
- Persist matches to data\matchPlans.json
- Persist UI/live state to data\appState.json
- Do not modify source CSVs
- Match numbering is based on current order and should renumber after reorder/delete
- Only one live match total across the whole app
- When setting a match live, overwrite OBS text output files in output folder:
  - current_division.txt
  - current_match.txt
  - current_bot_1.txt through current_bot_6.txt
  - current_status.txt
  - current_vs.txt
- For unused bot slots, write blank text
- current_status.txt should contain LIVE when a match is live
- current_vs.txt should contain:
  - vs for 2-bot matches
  - Battle Royale for 3+ bot matches
  - blank if no live match
- Provide live preview in GUI showing actual current output values
- Implement keyboard shortcuts:
  - Enter sets selected match live
  - Ctrl+Up moves match up
  - Ctrl+Down moves match down
  - Delete deletes selected match
  - Ctrl+L clears live
  - Ctrl+R reloads divisions
  - Ctrl+N creates match from selected bots
- Add plain text logging to logs\app.log
- Handle bad CSV or JSON gracefully without crashing
- Save immediately after create/edit/delete/reorder/live-state changes
- On startup:
  - create missing folders
  - load divisions
  - load saved matches/state
  - restore valid live match if present
  - otherwise clear output files
- Keep implementation readable and well-commented
- Use standard WinForms controls only
- Produce:
  1. full CombatMatchManager.ps1
  2. RunApp.bat
  3. short README.md with usage instructions