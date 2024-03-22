Option Explicit

Dim q

q=MsgBox("Do you know any programming language?", vbYesNoCancel, "Choice")

If q=6 Then 
    MsgBox "OK Hacker!", vbInformation, "Anonymous"
ElseIf q=7 Then 
    MsgBox "Hahaha looser!", vbCritical, "L" 
Else 
    MsgBox "Dont cheat next time!", vbExclamation, "Cheater"
End If