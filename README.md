# GST Automation Package

A supervised (human-in-the-loop) tool that bulk-downloads filed GST returns (GSTR-1,
GSTR-3B) and Notices & Orders from the Indian GST portal, across any number of client
GSTINs, driven from a simple Excel control sheet.

**This cannot be fully unattended.** The GST portal requires solving a CAPTCHA on every
login. No automation here bypasses that - a person must type the CAPTCHA into the real
browser window on every run. This tool automates everything around that one step.

## How it works

1. You list clients in the "Clients" sheet of the Excel workbook: name, GSTIN, portal
   login, password, financial year, and a Consent Obtained (Y/N) flag.
2. You select a row (or several) and click a button: Download GSTR-1 / Download GSTR-3B /
   Download GSTR-1 & 3B / Download Notices.
3. Excel VBA (`src/vba/GSTAutomation.bas`) checks consent, then shells out to a Python/
   Selenium script (`src/python/gst_portal_bot.py`) for each selected row.
4. A visible Chrome window opens on the GST login page with the username/password already
   filled in. You type the CAPTCHA and click Login, then confirm in the console.
5. The script navigates the portal (by clicking, never by loading a URL directly - doing
   that logs you out of this portal) and saves whatever it finds as PDFs into your
   Downloads folder.
6. Results are written back into the Excel sheet and an append-only Run Log.

## Repository layout

```
src/vba/GSTAutomation.bas        VBA module - import into the workbook's VBA editor
src/python/gst_portal_bot.py     The Selenium automation engine
src/python/requirements.txt      Python dependencies
templates/                       The blank control workbook (placeholder data only)
docs/SETUP.md                    Full step-by-step setup and usage instructions
```

## Quick start

See [`docs/SETUP.md`](docs/SETUP.md) for the full walkthrough. In short:

1. `pip install -r src/python/requirements.txt`
2. Open `templates/GST_Returns_Automation_TEMPLATE.xlsx`, import
   `src/vba/GSTAutomation.bas` via the VBA editor (Alt+F11), Save As `.xlsm`.
3. Add the four Form Control buttons and fill in the Settings sheet with your paths.
4. Test on your own GSTIN first.

## Before using this on a client

This tool stores a client's real portal username and password and logs in on their
behalf. Get the client's explicit written authorisation, a signed engagement letter
covering it, and have a supervisor present for every run (required anyway, for the
CAPTCHA), before marking that client's row Consent Obtained = Y.

## Known limitations

- Login-page field selectors and the exact "download a filed return" button are
  best-effort, built from general knowledge of the portal's conventions rather than a
  confirmed working example (the GSTIN used during development had no returns ever
  filed under it). Test carefully on a GSTIN with real filed returns and fix selectors
  in `src/python/gst_portal_bot.py` (all centralised near the top of the file) as needed.
- The portal's UI can change without notice.

## License

MIT - see [`LICENSE`](LICENSE).
