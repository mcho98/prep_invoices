# prep_invoices

Turns the monthly timesheet workbook into an invoice-import CSV.

## Installing (no Python or other software needed)

Everything the app needs (Python and openpyxl) is bundled inside it.

- **Windows:** download `PrepInvoices-Setup.exe`, double-click it, and click Next / Install.
  If Windows says "Windows protected your PC", click **More info**, then **Run anyway**.
  Afterwards open **Prep Invoices** from the desktop icon or Start menu.
- **Mac:** download the `.dmg` that matches your Mac (**Mac-AppleSilicon** for M1 or newer,
  **Mac-Intel** for older ones; Apple menu > About This Mac shows which). Open it and drag
  **PrepInvoices** onto **Applications**. The first time, right-click the app and choose **Open**.
  If macOS still refuses, go to System Settings > Privacy & Security and click **Open Anyway**.

The app is not code-signed, so these one-time warnings are expected.

## Using the app

Choose the timesheet, adjust the dates if needed, and click **Create invoice CSV**. The CSV is
saved next to the timesheet; **Show file** opens its folder.

## Publishing the installers (for the person maintaining this)

The GitHub Actions workflow in `.github/workflows/build.yml` builds the Windows installer and
both Mac disk images. Push the project to GitHub, then either:

- Actions tab > **Build apps** > **Run workflow**, and download the files from the finished run, or
- tag a release (`git tag v1.0 && git push --tags`) to publish them on the repo's Releases page,
  which is the easiest link to give users.

To build locally instead (must be done on each OS; PyInstaller can't cross-compile):

    pip install -r requirements.txt pyinstaller
    python build_app.py

## Command line

    python prep_invoices.py mock_data.xlsx --start 20260810 --end 20260820

Run `python prep_invoices.py --help` for all options. `python app.py` opens the window from source.
