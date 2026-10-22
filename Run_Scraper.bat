@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Universal Lead Extractor
set "SCRAPER_EXIT_CODE=1"

rem Double-click the VBS launcher for no console at all. Redirect old shortcuts
rem to it and close this bootstrap console immediately.
if "%~1"=="" if exist "%~dp0Run_Scraper.vbs" (
    start "" "%SystemRoot%\System32\wscript.exe" //nologo "%~dp0Run_Scraper.vbs"
    exit /b 0
)

rem Always use the application folder, including when launched by a shortcut.
pushd "%~dp0"
if errorlevel 1 goto folder_error
if not exist "Scraper.py" goto files_error

rem Prefer an environment installed beside the application.
set "SCRAPER_PYTHON="
set "SCRAPER_PY_ARGS="
if exist ".venv\Scripts\python.exe" set "SCRAPER_PYTHON=%CD%\.venv\Scripts\python.exe"
if defined SCRAPER_PYTHON goto ready
if exist "env\Scripts\python.exe" set "SCRAPER_PYTHON=%CD%\env\Scripts\python.exe"
if defined SCRAPER_PYTHON goto ready
if exist "venv\Scripts\python.exe" set "SCRAPER_PYTHON=%CD%\venv\Scripts\python.exe"
if defined SCRAPER_PYTHON goto ready

where py.exe >nul 2>&1
if errorlevel 1 goto try_python
set "SCRAPER_PYTHON=py.exe"
set "SCRAPER_PY_ARGS=-3"
goto ready

:try_python
where python.exe >nul 2>&1
if errorlevel 1 goto python_error
set "SCRAPER_PYTHON=python.exe"

:ready
"%SCRAPER_PYTHON%" %SCRAPER_PY_ARGS% -c "import sys; sys.exit(0 if sys.version_info[0] == 3 else 1)"
if not "%ERRORLEVEL%"=="0" goto interpreter_error
echo Checking application requirements...
"%SCRAPER_PYTHON%" %SCRAPER_PY_ARGS% -c "import Scraper, business_email, company_public_data, company_search_dom, google_maps_dom, email_domain_check, maps_office_scope, benchmark, run_support, enrichment_session, linkedin_links, scraper_ui, link_import, worksheet_editor; print('Universal Lead Extractor v' + Scraper.APP_VERSION + ' is ready.')"
if not "%ERRORLEVEL%"=="0" goto requirements_error
if /i "%~1"=="--check" goto success

echo Opening the application.
"%SCRAPER_PYTHON%" %SCRAPER_PY_ARGS% -u "Scraper.py"
set "SCRAPER_EXIT_CODE=%ERRORLEVEL%"
if "%SCRAPER_EXIT_CODE%"=="0" goto success
echo.
echo The application stopped with an error. The details are shown above.
goto failure

:interpreter_error
echo.
echo The selected Python interpreter could not start as Python 3:
echo "%SCRAPER_PYTHON%" %SCRAPER_PY_ARGS%
echo Repair this Python installation or recreate its local environment.
goto failure

:requirements_error
echo.
echo The requirements check failed. Review the error above.
echo Keep all extracted release files in this folder.
echo If Python packages are missing, install them with:
echo "%SCRAPER_PYTHON%" %SCRAPER_PY_ARGS% -m pip install customtkinter requests openpyxl playwright
goto failure

:files_error
echo Scraper.py is missing. Keep this launcher beside Scraper.py and the release files.
goto failure

:python_error
echo Python 3 was not found. Install Python 3 with Tkinter and add it to PATH,
echo or place your Python environment in this folder as .venv, env, or venv.
goto failure

:folder_error
echo Could not open the application folder.
if /i "%~1"=="--check" exit /b 1
if /i "%~1"=="--background" exit /b 1
pause
exit /b 1

:failure
if /i "%~1"=="--check" goto check_failure
if /i "%~1"=="--background" goto check_failure
pause
:check_failure
popd
exit /b %SCRAPER_EXIT_CODE%

:success
popd
exit /b 0
