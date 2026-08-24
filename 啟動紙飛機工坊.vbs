Option Explicit

Dim shell, fso, currentDir, pythonw, appFile, heartbeatUrl, started

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = fso.BuildPath(currentDir, "python_embed\pythonw.exe")
appFile = fso.BuildPath(currentDir, "app.py")
heartbeatUrl = "http://127.0.0.1:8080/api/heartbeat"

If Not fso.FileExists(pythonw) Then
    MsgBox "Portable Python was not found. Please run RUN.bat once to prepare the environment.", 16, "PaperSwitch"
    WScript.Quit 1
End If

If Not fso.FileExists(appFile) Then
    MsgBox "app.py was not found. Please extract the complete PaperSwitch folder before starting.", 16, "PaperSwitch"
    WScript.Quit 1
End If

shell.CurrentDirectory = currentDir
On Error Resume Next
shell.Run Quote(pythonw) & " " & Quote(appFile), 0, False
If Err.Number <> 0 Then
    MsgBox "PaperSwitch could not start. Please use RUN.bat to view the startup message.", 16, "PaperSwitch"
    WScript.Quit 1
End If
On Error GoTo 0

started = WaitForPaperSwitch(heartbeatUrl, 60, 200)
If Not started Then
    MsgBox "PaperSwitch did not respond within 12 seconds. Please use RUN.bat to view the startup message.", 16, "PaperSwitch"
    WScript.Quit 1
End If

On Error Resume Next
shell.Run "msedge.exe --app=http://127.0.0.1:8080 --window-size=1280,860", 1, False
If Err.Number <> 0 Then
    Err.Clear
    shell.Run "http://127.0.0.1:8080", 1, False
End If
On Error GoTo 0

Function WaitForPaperSwitch(url, attempts, delayMs)
    Dim request, i
    WaitForPaperSwitch = False
    For i = 1 To attempts
        On Error Resume Next
        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
        request.SetTimeouts 500, 500, 500, 500
        request.Open "GET", url, False
        request.Send
        If Err.Number = 0 And request.Status = 200 Then
            If InStr(request.ResponseText, "alive") > 0 Then
                WaitForPaperSwitch = True
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
        WScript.Sleep delayMs
    Next
End Function

Function Quote(value)
    Quote = Chr(34) & value & Chr(34)
End Function
