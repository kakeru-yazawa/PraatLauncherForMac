# Praat Launcher for Mac

A compact macOS tool that launches Praat from an Excel spreadsheet. This is a reimplementation of [Praat Launcher](https://www2.ninjal.ac.jp/past-events/2009_2021/event/specialists/project-meeting/files/JCLWorkshop_no2_papers/JCLWorkshop2012_2_30.pdf) by Ken'ya Nishikawa, which has been discontinued and whose source no longer seems to be available.

## What it does

When annotating a large speech corpus, you typically work from a spreadsheet of items — a word or segment you want to inspect, or where you stopped working. Praat Launcher turns each of those records into a shortcut. Click a cell in the row you want and trigger the launcher, and Praat opens the matching sound and/or TextGrid file, zoomed directly to that point.

## Requirements

| Item | Details |
| --- | --- |
| Microsoft Excel | Required |
| Praat | Must be located at `/Applications/Praat.app` |

Both Microsoft Excel and Praat also need **Full Disk Access** and **Accessibility** permissions (**System Settings** → **Privacy & Security**), since the launcher controls both apps. If a prompt appears the first time you run the launcher, click **OK** to allow access.

## Setup

The script (`PraatLauncher.applescript`) works on its own — open it with Script Editor and hit Run. It's more convenient, though, to wrap it in an Automator Quick Action with a keyboard shortcut, so you can trigger it without leaving Excel. The steps below cover that route.

### 1. Create the Automator Quick Action

1. Open **Automator** → **New Document → Quick Action**.
2. Set **Workflow receives current** to **no input**.
3. Add the **Run AppleScript** action.
4. Paste in the entire contents of `PraatLauncher.applescript`.
5. Save with a name of your choice (e.g. "PraatLauncher").

The workflow is stored in `~/Library/Services/`.

### 2. Assign a keyboard shortcut

In **System Settings → Keyboard → Keyboard Shortcuts → Services**, find your workflow under **General**, click **Add Shortcut**, and press your desired key combination.

### 3. Set up the PraatSettings sheet

On the first run with a workbook open, the launcher offers to create the `PraatSettings` sheet for you. Click **Create**, then fill it in:

| Row | Item | Example |
| --- | --- | --- |
| 1 | Sound file path | `/Users/xxx/<basename>.wav` |
| 2 | TextGrid file path | `/Users/xxx/<basename>.TextGrid` |
| 3 | Sound object type | `Sound` or `LongSound` |
| 4 | Column for basename | `A` |
| 5 | Column for start time | `B` |
| 6 | Column for end time | `C` |
| 7 | Color of launched row | `Blue` / `Green` / `Gray` / `Red` / `Unchanged` |

Use `<basename>` as a placeholder in the file paths; it is replaced with the value in the basename column of the selected row. Paths can be absolute, or relative to the folder the workbook is saved in.

For the sound object type, use `Sound` in most cases; switch to `LongSound` only for very long recordings that are impractical to load fully into memory.

## Sandbox

The `Sandbox` folder contains a ready-to-run example. Open the workbook `Test.xlsx`, click a cell in any row, and trigger the launcher to see it in action.

## Acknowledgements

Full credit for the idea and design of Praat Launcher goes to Nishikawa-san, who was my supervisor during my time at the Laboratory for Language Development, RIKEN Brain Science Institute.
