# MPR Combiner (Usage Tab)

Combine the `Usage` or `Usage (Part B)` sheets from many MPR Excel files into **one workbook** in the same folder you pick.

## Run the app (GUI)

**From the project folder in Terminal** (shortest):

```bash
./gui
```

**On a Mac**, you can double-click **`MPR Combiner.command`** in Finder (opens Terminal and starts the app).

The first launch runs Maven to build the JAR; later launches are quicker. Same as `./scripts/run.sh` if you prefer that path.

Advanced (after you have built once with `mvn package`):

```bash
java -jar target/mpr-combiner-1.0.0-all.jar
```

**Company logo (optional):** Put your mark at `src/main/resources/branding/logo.png` or `logo.jpg`. It appears in the top header next to the title when you run the app. A PNG with a transparent background and roughly square or horizontal proportions works well (the UI scales it to a fixed height). To share the artwork with someone helping in Cursor, add that file in the repo or attach the image in chat.

## What to do in the UI

1. Use **Open library** (or **Browse…**) to point at the folder that contains your MPR workbooks (for example your synced OneDrive or SharePoint library). Pick the **Combine from** folder in the tree or path field.
2. Turn **Search subfolders for workbooks** on or off depending on whether workbooks are nested.
3. Click **Combine**. In **Output**, optionally set **Output file** or **Output folder**, choose whether to open the workbook and copy the Copilot prompt, then click **Combine** again to start.

Files whose names start with `99_combined` and end with `.xlsx` are **not** read as inputs, so you can re-run a combine without pulling the last combined workbook into the next one.

## Next step — Excel Copilot

Excel does not offer a supported way for external apps to type into Copilot for you. The app can still help:

- In the **Output** dialog (after **Combine**), enable **Copy Copilot prompt to clipboard** (on by default) and optionally **Open workbook when finished**. After a successful run, paste into Copilot with **⌘V** (Mac) or **Ctrl+V** (Windows).

The exact text lives in `src/main/resources/copilot-excel-prompt.txt` (bundled in the JAR). Open or copy from that file, or rely on the app’s clipboard option after combine, then paste into Excel Copilot.

## Notes

- **Combine timeout**: a single combine run stops automatically after **15 minutes** (wall clock) so a stuck read (for example OneDrive/SharePoint files that are not fully synced yet) does not hang forever. Override with e.g. `java -Dmpr.combiner.timeoutMinutes=30 -jar …` if you routinely combine very large libraries.
- **SharePoint / OneDrive**: sync the library to this computer, then choose that folder in the app. There is no Microsoft sign-in; files are read from disk only.
- **Sheet selection**: uses `Usage` if present; otherwise `Usage (Part B)`.
- **Headers**: common typos and broken spacing in column names are normalized.
- **Worksheet**: combines the `Usage` sheet (or `Usage (Part B)` when `Usage` is missing) with header row 1.
