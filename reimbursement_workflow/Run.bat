@echo off
REM ============================================================
REM  Fleet Gas Reimbursement - run the monthly processor.
REM  1) Fill in the console workbook (Monthly Input / Employees /
REM     Locations) and pick the month on the Dashboard.
REM  2) SAVE and CLOSE Gas_Reimbursement_Console.xlsx.
REM  3) Double-click this file.
REM  4) Re-open the workbook to see the updated tabs.
REM ============================================================
cd /d "%~dp0scripts"
echo Running the gas-reimbursement processor...
echo.
python process_console.py
if errorlevel 1 (
  echo.
  echo *** Something went wrong. Make sure the workbook is CLOSED, then try again. ***
)
echo.
pause
