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

Customer names on the invoices come from the client chart (the `Family# Client# Chart` sheet, or the
first sheet). Each client is matched in this order:

1. **AHS ID.** Read from the cell right of the "AHS#" label in the client's section (or, if that is
   empty, the cell next to the client number above "Client Name"), and looked up in the chart's `AHS ID#`
   column.
2. **Family + client number.** If the AHS ID is missing or not in the chart, the family and client
   numbers above "Client Name" are looked up in the chart's `Family #` and `Client #` columns. A warning
   is shown.
3. **Timesheet name, cleaned.** If neither matches, the timesheet name is used with all numbers and
   Chinese characters removed, and a warning is shown.

Chart names are used as written, with leading and trailing spaces removed.

The chart path is saved in `config/prep_invoices.json` next to the app. If it is missing or the file
has moved, the app asks for it again. From the command line use `--chart FILE.xlsx` to set or change it.

## Office

A section whose name is exactly "Office" (any capitals, spaces around it ignored) is not a client: it is
skipped, with a "Skipping Office" line in the log, whatever its hours, rate or numbers say. Names like
"Office 2" are treated as normal clients.

## Large files (QuickBooks 1000-row limit)

QuickBooks accepts at most 1000 rows per CSV upload and it is unclear whether the header counts, so each
CSV holds at most 999 invoice lines plus the header. If there are more, the output is split into
`..._part1.csv`, `..._part2.csv`, and so on. A client's invoice is never split between files. The Excel
summary is still a single file covering everything.

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
