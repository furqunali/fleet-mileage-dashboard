@echo off
REM ============================================================
REM  REFRESH THE EMBEDDED MATRIX  —  Gas Reimbursement Calculator
REM
REM  Use this ONLY when a new site (or corrected distance) has been
REM  added to the master distance workbook and you want the
REM  calculator to use it.
REM
REM  It KEEPS all your data (employees + every month's trips) and
REM  makes a timestamped backup in the "backups" folder first.
REM
REM  STEPS:
REM   1) Make sure Fleet_Distance_Dashboard_v5.xlsx (the master,
REM      in the Houston_Distance_Dashboard folder) is up to date.
REM   2) CLOSE Gas_Reimbursement_Calculator.xlsx in Excel.
REM   3) Double-click this file.
REM   4) Re-open the calculator — the new site is now in the
REM      dropdowns and Matrix, and all tabs recalculate.
REM ============================================================
cd /d "%~dp0scripts"
echo Refreshing the embedded company Matrix...
echo (Your employees and monthly data are preserved. A backup is made first.)
echo.
python build_calculator.py --rebuild
if errorlevel 1 (
  echo.
  echo *** Something went wrong. Make sure the calculator is CLOSED in Excel, then run again. ***
)
echo.
pause
