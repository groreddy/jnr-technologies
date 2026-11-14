@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0update-prices.ps1"
if errorlevel 1 (
  echo.
  echo Price update failed. Read the error above, then check the Excel workbook.
) else (
  echo.
  echo Website prices are ready. Refresh the browser to see them.
)
pause
