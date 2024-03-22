Option Explicit

Dim q

q=MsgBox("Do you know any programming language?", vbYesNoCancel, "Choice")

If q=6 Then 
    MsgBox "OK Hacker!", vbInformation, "Anonymous"
ElseIf q=7 Or q=2 Then 
    MsgBox "Hahaha looser!", vbCritical, "L" 
End If