Sub RunPowerShellScript()
  Dim objShell As Object
  Set objShell = CreateObject("WScript.Shell")
  objShell.Run "powershell.exe -File C:\path\to\script.ps1"
End Sub
