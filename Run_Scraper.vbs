' Open the desktop application without a visible CMD window.
' The batch launcher still selects the interpreter and checks requirements.
Option Explicit
Dim shell, files, folder, batch, logFolder, logPath, stream, mode, checking, command, result
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
folder = files.GetParentFolderName(WScript.ScriptFullName)
batch = files.BuildPath(folder, "Run_Scraper.bat")
checking = False
If WScript.Arguments.Count > 0 Then checking = (LCase(WScript.Arguments(0)) = "--check")

If Not files.FileExists(batch) Then
    If Not checking Then MsgBox "Run_Scraper.bat is missing. Keep all extracted application files together.", 16, "Lead Extractor"
    WScript.Quit 1
End If

On Error Resume Next
logFolder = files.BuildPath(shell.ExpandEnvironmentStrings("%LOCALAPPDATA%"), "UniversalLeadExtractor")
If Not files.FolderExists(logFolder) Then files.CreateFolder logFolder
logFolder = files.BuildPath(logFolder, "LauncherLogs")
If Not files.FolderExists(logFolder) Then files.CreateFolder logFolder
logPath = files.BuildPath(logFolder, Replace(files.GetTempName, ".tmp", ".log"))
Set stream = files.CreateTextFile(logPath, True)
stream.WriteLine "Lead Extractor startup: " & Now
stream.Close
If Err.Number <> 0 Then
    If Not checking Then MsgBox "Could not create the startup log: " & Err.Description, 16, "Lead Extractor"
    WScript.Quit 1
End If
On Error GoTo 0

mode = "--background"
If checking Then mode = "--check"
' Environment substitutions keep spaces, ampersands and percent signs in paths
' inside quotes without interpreting their contents as shell commands.
shell.Environment("PROCESS")("SCRAPER_LAUNCH_BATCH") = batch
shell.Environment("PROCESS")("SCRAPER_LAUNCH_LOG") = logPath
shell.Environment("PROCESS")("PYTHONIOENCODING") = "utf-8"
command = Quote(shell.ExpandEnvironmentStrings("%ComSpec%")) & " /d /v:off /s /c " & _
    Quote(Quote("%SCRAPER_LAUNCH_BATCH%") & " " & mode & " >> " & Quote("%SCRAPER_LAUNCH_LOG%") & " 2>&1")
On Error Resume Next
result = shell.Run(command, 0, True)
If Err.Number <> 0 Then
    If Not checking Then MsgBox "Could not start the application: " & Err.Description & vbCrLf & vbCrLf & _
        "Startup log: " & logPath, 16, "Lead Extractor"
    WScript.Quit 1
End If
On Error GoTo 0
If result <> 0 And Not checking Then
    MsgBox "The application could not start or stopped with an error." & vbCrLf & vbCrLf & _
        "Open this startup log for details:" & vbCrLf & logPath, 16, "Lead Extractor"
End If
WScript.Quit result

Function Quote(value)
    Quote = Chr(34) & value & Chr(34)
End Function
