# PC-DMIS Report Extractor

Open-source PC-DMIS report exporter maintained by Townsend Precision Labs Ltd.

> **Status:** working validation build. It is useful in its current state, but it is not a complete replacement for every PC-DMIS report path.

The extractor reads evaluated dimensions from the measurement routine currently open in PC-DMIS and creates:

- a formatted Excel workbook (`.xlsx`);
- a matching emergency/readable CSV (`.csv`);
- an export evidence log (`.export.log`).

It has been validated with PC-DMIS 2026.1 against a copied/offline validation routine containing legacy dimensions and standalone ISO `SIZE/` commands.

## Current coverage

- Evaluated legacy dimensions.
- Standalone ISO `SIZE/` commands.
- ASME `SIZE/` commands when PC-DMIS exposes the evaluated UAME and local-size results.
- PASS/FAIL, nominal, tolerances, measured value, deviation, out-of-tolerance, bonus, minimum and maximum values where PC-DMIS exposes them.
- Automatic report filenames for successive components in a PC-DMIS loop.

## Known limitation

New `GEOMETRIC_TOLERANCE/` commands (new GeoTol/new FCF) have not yet been validated because the current test routine does not contain a representative example. Legacy geometric dimensions are included. Do not treat this release as complete new-GeoTol support until position, profile and multi-segment examples have been tested.

## Requirements

- Windows workstation running PC-DMIS.
- Microsoft Excel desktop, used through Excel COM to create `.xlsx` files.
- Windows PowerShell.
- Permission to create `C:\PCDMIS_Report_Fallback` and write to the chosen report folder.

## Package files

- `Run-PCDMIS-Dimension-Extractor.bas` — automatic PC-DMIS BASIC script. This is the file normally called by the measurement routine.
- `Export-PCDMIS-Dimensions.bas` — reusable BASIC extraction library without the automatic `Main` entry point.
- `Export-PCDMIS-Dimensions.ps1` — turns the extracted data into CSV and Excel reports.
- `Install-PCDMIS-Dimension-Extractor.ps1` — installs the working files and preserves timestamped backups of previous versions.
- `Required-PCDMIS-Variables.txt` — the five required PC-DMIS variable names.
- `SHA256SUMS.txt` — integrity hashes for the packaged files.
- `LICENSE` — GNU General Public License version 3 or later.

## Installation

1. Download the repository as a ZIP or clone it.
2. Open PowerShell in the downloaded folder.
3. Run:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Install-PCDMIS-Dimension-Extractor.ps1"
   ```

4. If Windows refuses to create the target folder, run PowerShell as Administrator and repeat the command.
5. The installer copies and hash-verifies the working files in:

   ```text
   C:\PCDMIS_Report_Fallback
   ```

6. The BASIC script to call from PC-DMIS is:

   ```text
   C:\PCDMIS_Report_Fallback\Run-PCDMIS-Dimension-Extractor.bas
   ```

The fixed local working folder gives the BASIC and PowerShell parts a predictable handoff location. The final report location and filename are controlled by variables in the measurement routine.

## The five required PC-DMIS variables

Define and populate these variables before the extractor runs:

1. `EXPORT_PATH` — destination folder only, with no filename. Example: `W:\Inspection_Reports`.
2. `FILE_PREFIX` — routine, part or drawing prefix. Example: `PART123_REV_A`.
3. `WORKS_ORDER` — works order or P number. Example: `WO000123`.
4. `SERIAL` — the **starting serial number** for the loop. It must be a whole number.
5. `LOOPCOUNT` — the current one-based report counter. Set it to `1` before the loop and increment it manually at the end of each loop iteration.

The current serial number is calculated as:

```text
current serial = SERIAL + LOOPCOUNT - 1
```

Example with `SERIAL = 1001`:

- `LOOPCOUNT = 1` produces serial `1001`.
- `LOOPCOUNT = 2` produces serial `1002`.
- `LOOPCOUNT = 5` produces serial `1005`.

`LOOPCOUNT` must be a whole number of 1 or greater. The report script must run before `LOOPCOUNT` is incremented so the current component receives the correct serial number.

## Output filename

The workbook filename is constructed as:

```text
FILE_PREFIX_WORKS_ORDER_CURRENT_SERIAL.xlsx
```

With:

```text
EXPORT_PATH = W:\Inspection_Reports
FILE_PREFIX = PART123_REV_A
WORKS_ORDER = WO000123
SERIAL = 1001
LOOPCOUNT = 3
```

The current serial is `1001 + 3 - 1 = 1003`, so the output is:

```text
W:\Inspection_Reports\PART123_REV_A_WO000123_1003.xlsx
```

A CSV and export log with the same base filename are written beside it.

## Example: beginning of a looped PC-DMIS routine

Ask the operator how many parts to run, convert the answer to an integer, initialise the five extractor variables and then start the PC-DMIS loop. `PARTS_TO_RUN` and `LOOPINDEX` control the routine; they are not additional variables required by the extractor.

```text
C1=COMMENT/INPUT,NO,FULL SCREEN=NO,
How many parts should this routine run? Enter a whole number of 1 or greater.
ASSIGN/PARTS_TO_RUN=INT(C1.INPUT)
ASSIGN/EXPORT_PATH="W:\Inspection_Reports"
ASSIGN/FILE_PREFIX="PART123_REV_A"
ASSIGN/WORKS_ORDER="WO000123"
ASSIGN/SERIAL=1001
ASSIGN/LOOPCOUNT=1
LOOPINDEX=LOOP/START,ID=NO,NUMBER=PARTS_TO_RUN,START=1,SKIP=,
```

The exact spacing and optional fields displayed by PC-DMIS can vary. The important points are:

- the operator input is stored as `C1.INPUT` and converted to the numeric `PARTS_TO_RUN` value;
- `PARTS_TO_RUN` is used as the loop's `NUMBER` value;
- `LOOPCOUNT` is explicitly initialised to `1` before the loop;
- the PC-DMIS loop can use a separate variable such as `LOOPINDEX` for its own iteration value;
- `SERIAL` is the first component's numeric serial number.

## Example: end of each loop iteration

Place the BASIC script call **after all dimensions for the current component have executed** and immediately **before `LOOP/END`**.

A PC-DMIS script block will resemble:

```text
CS1=SCRIPT/FILENAME=C:\PCDMIS_Report_Fallback\Run-PCDMIS-Dimension-Extractor.bas
FUNCTION/Main,SHOW=NO,,
STARTSCRIPT/
ENDSCRIPT/
ASSIGN/LOOPCOUNT=LOOPCOUNT+1
LOOP/END
```

Insert the BASIC Script command using the PC-DMIS interface for your installed version, select the installed `.bas` file, choose the `Main` function and compile it. PC-DMIS may format the command block differently, but the script must execute before `LOOPCOUNT` is increased and before the loop advances to the next component.

## How to produce reports

1. Install the package.
2. Add an input comment asking how many parts to run and assign `PARTS_TO_RUN=INT(C1.INPUT)`.
3. Set the five extractor variables, including `LOOPCOUNT=1`, before the loop.
4. Use `PARTS_TO_RUN` as the PC-DMIS loop's `NUMBER` value.
5. Put the BASIC Script call after the current component's dimensions.
6. Immediately after the script call, add `ASSIGN/LOOPCOUNT=LOOPCOUNT+1`, followed by `LOOP/END`.
7. Compile the BASIC script when setting up the routine.
8. Run the routine. PC-DMIS asks once for the number of parts, then the exporter reads its five variables without additional prompts.
9. Confirm that the `.xlsx`, `.csv` and `.export.log` files were created in `EXPORT_PATH`.
10. Review the report before relying on it as controlled inspection evidence.

Reusing the same five values produces the same filename and replaces the previous report. Use a unique works-order/serial combination where inspection records must not be overwritten.

## Working and diagnostic files

The extractor temporarily uses these files in `C:\PCDMIS_Report_Fallback`:

- `last-basic-bridge.tsv`
- `last-command-mode.txt`
- `export-request.txt`
- `last-basic-error.log` if BASIC fails
- `last-export-error.log` if rendering fails

These are working or diagnostic files, not controlled inspection reports.

## Restart guidance

PC-DMIS does not normally need restarting between clean exports. Restart PC-DMIS and reopen the routine if:

- PC-DMIS crashes or becomes unresponsive;
- a native/object error is reported;
- an output or diagnostic file is missing or incomplete;
- results appear blank, zero or stale;
- you replaced the installed `.bas` file after a failed test or crash.

PC-DMIS can retain an older BASIC script in memory during an active routine or loop, so restart once after replacing the script before production validation.

## Record-control recommendation

For controlled work, retain each generated `.xlsx`, `.csv` and `.export.log` beside the relevant job records. Keep the PC-DMIS routine revision, drawing revision, works order and serial number traceable to the same inspection evidence.

## Contributing

Bug reports, test routines stripped of customer/confidential data, compatibility findings and pull requests are welcome. Do not upload customer drawings, measurement results, serial numbers, proprietary routines or other controlled information.

When changing extraction behaviour, state the PC-DMIS version and command type used for validation. A change should not claim support for a command family without a representative test.

## License

Copyright (C) 2026 Townsend Precision Labs Ltd.

This project is licensed under the **GNU General Public License v3.0 or later** (`GPL-3.0-or-later`). You may use, study, modify and redistribute it. If you distribute this program or a modified version, the GPL requires the corresponding source to remain available under the GPL's terms. See `LICENSE` for the complete terms.

PC-DMIS is a trademark of its respective owner. This independent project is not affiliated with or endorsed by Hexagon.
