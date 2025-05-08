Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run WshShell.CurrentDirectory & "\Examples.cmd", 0 
Set WshShell = Nothing 
MsgBox "It's Change for Examples" 
