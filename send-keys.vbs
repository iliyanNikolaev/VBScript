Option Explicit
Dim x
Set x = CreateObject("wscript.shell")
x.run "cmd.exe"
WScript.Sleep 2000
x.SendKeys "cd C:\Users\ilich\Desktop\bgato018-tracking-system\server"
x.SendKeys "{ENTER}"
x.SendKeys "node index"
x.SendKeys "{ENTER}"
