@echo off
setlocal
set "REPO=%~dp0..\.."
cd /d "%REPO%"
set "PWQC_REPO_ROOT=%CD%"
set "PWQC_APPSETTINGS=%CD%\appsettings.json"
set "PWQC_SQL_SERVER=192.168.22.90"
set "PWQC_SQL_DATABASE=QC_Pipeline"
set "PWQC_SQL_TRUST_CERT=yes"
npx @modelcontextprotocol/inspector "%~dp0.venv\Scripts\python.exe" "%~dp0server.py"
