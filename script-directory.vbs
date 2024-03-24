Option Explicit
Dim a, b, c

a = WScript.ScriptName
b = WScript.ScriptFullName
Set c = CreateObject("wscript.shell")
MsgBox a
MsgBox b
MsgBox c.CurrentDirectory
