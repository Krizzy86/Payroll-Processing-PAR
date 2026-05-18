# Payroll Processing - PAR

Phase 1 extracts payroll-ready employee data from the weekly ExactaBKS timecard export.

## What It Does

The script reads `Employee Timecard.xlsx` from a pay-period folder and creates:

- `timesheet_config.xlsx`
- `Phase 1 Payroll Extract.xlsx`

The output columns are:

1. `Name`
2. `Total Hours`
3. `Base Pay`
4. `Guaranteed Rate`

## Weekly Process

1. Put the new timecard export in the weekly pay-period folder.

   The folder should be named for the ending Sunday date in `yyyymmdd` format.

   Example:

   20260524

   The file inside that folder must be named exactly:

   Employee Timecard.xlsx

2. Open the project folder:

   D:\Software\Codex\Payroll-Processing-PAR

3. If needed, update the config file in the pay-period folder:

   timesheet_config.xlsx

   The config file has these sections:

   Salary Employees  
   Non Tip Pool Employees  
   Ignore These Employees

   Names should match the timecard format:

   Last, First

   Examples:

   Leon, Tina  
   Hernandez, Aydelin

   If `timesheet_config.xlsx` does not exist yet, the script will create it the first time you run the process.

4. Close these files before running the process:

   Employee Timecard.xlsx  
   timesheet_config.xlsx  
   Phase 1 Payroll Extract.xlsx

   Excel locks open files, so the script may fail if any of them are open.

5. Open PowerShell.

6. Go to the project folder:

   cd "D:\Software\Codex\Payroll-Processing-PAR"

7. Run the script, replacing the folder path with the current weekly pay-period folder:

   python extract_timecard_phase1.py "C:\Users\kirk\OneDrive - Rizzy Group\Shared Documents\General\40_Store Operations\TX004 RoundRock\Vendors\Payroll\ExactaBKS\2026\02_AprMayJun\20260524"

8. The script will create or update this file in the same pay-period folder:

   Phase 1 Payroll Extract.xlsx

## Sorting Rules

The output is sorted in this order:

1. Salary employees
2. Non-tip-pool employees
3. All remaining hourly employees

Each group is sorted alphabetically.

Employees listed under `Ignore These Employees` are excluded from the output.

## Guaranteed Rate

`Guaranteed Rate` is calculated as:

- Non-tip-pool employees: same as `Base Pay`
- Employees with `Base Pay > 0`: `Base Pay + 3`
- Employees with `Base Pay = 0`: `0`

## Example Run

For the May 17, 2026 pay period:

cd "D:\Software\Codex\Payroll-Processing-PAR"

python extract_timecard_phase1.py "C:\Users\kirk\OneDrive - Rizzy Group\Shared Documents\General\40_Store Operations\TX004 RoundRock\Vendors\Payroll\ExactaBKS\2026\02_AprMayJun\20260517"

For the next week, change only the final folder date:

python extract_timecard_phase1.py "C:\Users\kirk\OneDrive - Rizzy Group\Shared Documents\General\40_Store Operations\TX004 RoundRock\Vendors\Payroll\ExactaBKS\2026\02_AprMayJun\20260524"

## Common Problems

### File is locked

If the script fails with a permission error, close:

Employee Timecard.xlsx  
timesheet_config.xlsx  
Phase 1 Payroll Extract.xlsx

Then run the script again.

### Wrong folder date

The pay-period folder date should match the ending Sunday of the pay period.

Example:

20260524

### Wrong source file name

The input file must be named exactly:

Employee Timecard.xlsx