# prep_invoices

Turns the monthly timesheet workbook into an invoice-import CSV.

## Installing (no Python or other software needed)

Everything the app needs (Python and openpyxl) is bundled inside it.

- **Windows:** download `PrepInvoices.exe` and double-click it. There is nothing to install; keep
  the file wherever is convenient (e.g. the Desktop). If Windows says "Windows protected your
  PC", click **More info**, then **Run anyway**. The first start takes a few seconds.
- **Mac (M1 or newer only):** download `PrepInvoices-Mac.dmg`, open it, and drag
  **PrepInvoices** onto **Applications**. The first time, right-click the app and choose **Open**.
  If macOS still refuses, go to System Settings > Privacy & Security and click **Open Anyway**.
  Older Intel Macs are not supported.

The app is not code-signed, so these one-time warnings are expected.

## Using the app

Choose the timesheet, adjust the dates if needed, and click **Create invoice CSV**. The CSV is
saved next to the timesheet, together with an Excel summary (client name, total hours, total amount)
to check against after uploading. Both are named `date_time_timesheet name`, e.g.
`20261008_143005_mock_data.csv` and `.xlsx`. **Show file** opens their folder.

## Client chart (English names)

Customer names on the invoices come from the client chart (`Family# Client# Chart` sheet), matched on
each client's family and client numbers in the timesheet; the chart name is used as written, with leading and trailing spaces removed. The chart
path is saved in `config/prep_invoices.json` next to the app. If it is missing or the file has moved,
the app asks for it again. From the command line use `--chart FILE.xlsx` to set or change it. A client
not found in the chart keeps the timesheet name, with a warning.

## Publishing the installers (for the person maintaining this)

The GitHub Actions workflow in `.github/workflows/build.yml` builds the Windows exe and
the Mac disk image. Push the project to GitHub, then either:

- Actions tab > **Build apps** > **Run workflow**, and download the files from the finished run, or
- tag a release (`git tag v1.0 && git push --tags`) to publish them on the repo's Releases page,
  which is the easiest link to give users.

To build locally instead (must be done on each OS; PyInstaller can't cross-compile):

    pip install -r requirements.txt pyinstaller
    python build_app.py

## Command line

    python prep_invoices.py mock_data.xlsx --start 20260810 --end 20260820

Run `python prep_invoices.py --help` for all options. `python app.py` opens the window from source.
