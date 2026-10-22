# Start Lead Extractor

1. Extract the complete ZIP into a folder.
2. Double-click **Run_Scraper.vbs**.

Only the application window opens. Keep the VBS launcher, batch launcher and
Python files together. You can create a desktop shortcut to Run_Scraper.vbs.

Version 42.9.0 adds large-file import counts, a 15-second email lookup budget,
per-stage progress, and a worksheet editor for custom columns and manual rows.
See [the worksheet guide](README_WORKSHEET.md) for the new controls.

Chromium sandboxing remains enabled for the LinkedIn browser and both
enrichment sessions, including the Chromium fallback. Close the previous
application and the Chrome windows it opened before starting this release.
Already-running browser processes cannot pick up the new launch setting.

The batch file also redirects ordinary double-clicks to the windowless launcher,
but Windows may briefly show its bootstrap console. Use the VBS file directly
to avoid that flash.

Startup and console diagnostics are saved to:

`%LOCALAPPDATA%\UniversalLeadExtractor\LauncherLogs`

Each launch creates a separate log. If startup fails or the application exits
with an error, a message gives the log path. Normal application activity remains
visible inside the app and in its existing output activity log.

For troubleshooting from a terminal:

```text
Run_Scraper.bat --check
Run_Scraper.bat --console
```

The windowless launcher uses Windows Script Host. If your organization's policy
disables it, use the console launcher and ask your IT team about permitted
application launch options. No system policies are changed by this release.
