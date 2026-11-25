# Setup & Usage

## 1. Install Python (one-time)

Install Python 3.10+ from python.org, ticking "Add python.exe to PATH" during install.

## 2. Install dependencies (one-time)

Open Command Prompt in this repository folder and run:

    pip install -r src/python/requirements.txt

Selenium 4.15+ includes Selenium Manager, which downloads the matching ChromeDriver
automatically the first time you run the script - no manual ChromeDriver install, as
long as Google Chrome itself is installed.

## 3. Prepare the workbook (one-time)

- Copy `templates/GST_Returns_Automation_TEMPLATE.xlsx` somewhere you'll keep using it
  (don't edit the template copy in the repo directly - keep that one as a clean blank
  starting point for next time).
- Open it in Excel, `Alt+F11` to open the VBA editor, right-click
  **VBAProject (this file)** -> **Import File...** -> select
  `src/vba/GSTAutomation.bas`.
- **File -> Save As -> Excel Macro-Enabled Workbook (*.xlsm)**. Use this `.xlsm` copy
  from now on.

## 4. Add the Developer tab (if not already visible)

File -> Options -> Customize Ribbon -> tick "Developer".

## 5. Insert the four buttons

On the "Clients" sheet, near row 4: Developer tab -> Insert -> Button (Form Control) ->
draw it -> in "Assign Macro" pick the matching macro -> rename the button's caption:

| Button caption          | Macro to assign  |
|--------------------------|------------------|
| Download GSTR-1          | `DownloadGSTR1`  |
| Download GSTR-3B         | `DownloadGSTR3B` |
| Download GSTR-1 & 3B     | `DownloadBoth`   |
| Download Notices         | `DownloadNotices`|

## 6. Fill in the Settings sheet

Your Python path (`where python` in Command Prompt tells you this), the full path to
`gst_portal_bot.py` on your machine, a jobs/temp folder, and your default Downloads
folder.

## 7. Enable macros

File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> "Disable
all macros with notification" (then click "Enable Content" only on this workbook).

## 8. Test on your own GSTIN first

Fill a row with your own GSTIN and portal login, mark Consent Obtained = Y, pick a
Financial Year, select that row, and click a button. Watch the first run closely - it
proves out the login-field and download-button selectors, which are best-effort (see
the Known Limitations section of the main README).

## Using an editor to change the VBA code going forward (optional)

Instead of re-importing a `.bas` file by hand every time, `xlwings`
(`pip install xlwings`) has an `xlwings vba edit` command: it exports every VBA module
to `.bas`/`.cls` files next to the workbook and watches them - edit a file in VS Code,
save, and it's pushed live into the open Excel workbook's VBA project, no manual
import step. Requires "Trust access to the VBA project object model" ticked once
(File -> Options -> Trust Center -> Macro Settings) and the workbook open.

## If something breaks

- Check the **Run Log** sheet first (status + message for every run).
- Check the **black console window** - Python prints progress and errors there.
- All page-lookup logic is centralised at the top of `src/python/gst_portal_bot.py`
  under `CONFIG` - update the relevant XPath there if the portal's screen changed.
