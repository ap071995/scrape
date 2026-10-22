# Scraper v42.8.1 — windowless launcher

The new Run_Scraper.vbs starts the existing launcher with its console hidden.
It preserves interpreter selection, requirement checks, the application working
directory and failure exit codes. Normal Run_Scraper.bat double-clicks delegate
to the new launcher; --check and --console remain available.

Startup output is redirected to a separate per-launch log in
%LOCALAPPDATA%\UniversalLeadExtractor\LauncherLogs. A startup/application failure
shows a message with its log path. UTF-8 logging supports non-ASCII output.
Background failure paths return instead of waiting at an invisible pause prompt.

Validation: **19 launcher tests passed in 9.300 seconds**, including five new
tests for successful windowless startup, silent preflight, diagnostic retention,
background failure without pause, and delegation from the old batch launcher.
Fixtures verify interpreter and working-directory selection in paths containing
spaces, exclamation marks, ampersands and parentheses. They do not open websites
or process user workbooks.

The actual application's windowless preflight also passed. Its version is
42.8.1. The previous v42.8.0 scraper and UI behavior is retained. The complete
extraction regression suite was not rerun for this launcher-only update.

Use Run_Scraper.vbs directly to avoid any console flash. Windows itself may
briefly show a console when starting a .bat file before delegation completes.
