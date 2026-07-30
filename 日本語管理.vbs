Option Explicit
Dim fso, shell, rootDir, appDir, launcher, psExe, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
rootDir = fso.GetParentFolderName(WScript.ScriptFullName)
appDir = fso.BuildPath(rootDir, "app")
launcher = fso.BuildPath(appDir, "launch.ps1")
If Not fso.FileExists(launcher) Then
  MsgBox "ReportBinder launcher was not found:" & vbCrLf & launcher, vbCritical, "ReportBinder"
  WScript.Quit 1
End If
psExe = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
If Not fso.FileExists(psExe) Then psExe = "powershell.exe"
cmd = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & launcher & """ -Mode ja"
shell.Run cmd, 0, False
