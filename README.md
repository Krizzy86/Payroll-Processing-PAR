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

## Pay Period Folder

The pay-period folder name should be the ending Sunday date in `yyyymmdd` format.

Example:

```text
20260517
```

The source file inside that folder should be named:

```text
Employee Timecard.xlsx
```

## Configuration

The script creates `timesheet_config.xlsx` if it does not already exist.

The config contains these editable sections:

- `Salary Employees`
- `Non Tip Pool Employees`
- `Ignore These Employees`

Employee names should match the timecard format:

```text
Last, First
```

## Sorting

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

## Run

```powershell
python extract_timecard_phase1.py "C:\path\to\pay-period-folder"
```

For the current development sample:

```powershell
python extract_timecard_phase1.py "C:\Users\kirk\OneDrive - Rizzy Group\Shared Documents\General\40_Store Operations\TX004 RoundRock\Vendors\Payroll\ExactaBKS\2026\02_AprMayJun\20260517"
```
