# Rent IIF Converter

A small Windows GUI for converting an Avena rent transaction Excel report into a QuickBooks `.iif` import file.

## Requirements

- Windows
- PowerShell 7
- A rent transaction report named like `RentTransactionDetailReport 06122026.xlsx`

This tool reads `.xlsx` files directly and does not require Excel to be installed.

## Use

1. Double-click `Run-RentIifConverter.cmd`.
2. Choose the rent transaction report `.xlsx`.
3. Confirm the processing date.
4. Confirm the output folder.
5. Confirm the QuickBooks account names.
6. Click `Preview`.
7. Click `Create IIF`.

The generated file is tab-delimited text, not an Excel workbook renamed with `.iif`.

## Notes

- The output folder defaults to the same folder as the selected rent report.
- The app remembers the last folders and account names.
- `Open IIF` opens the generated file in Notepad.
- `Copy Path` copies the generated file path to the clipboard.
- The app validates that the selected workbook has the expected `Tenant`, `Datetime`, and `Payment` columns.
