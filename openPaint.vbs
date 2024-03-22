Option Explicit
Dim runner
Dim choice

Set runner=CreateObject("wscript.shell")
choice=MsgBox("Do you really want to open a Paint?", vbYesNo)

If choice=vbYes Then
    runner.run("mspaint.exe")
End If