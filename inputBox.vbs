' InputBox "message", "title", "input field", xposition, yposition
Option Explicit
Dim name
Dim age

name = InputBox("What is your name?", "Name")

If name="" Then
    MsgBox "Name cannot be empty!", vbExclamation, "Error"
    WScript.Quit
End If

age = InputBox("What is your age?", "Age")

If Not IsNumeric(age) Then
    MsgBox "Invalid age!", vbExclamation, "Error"
    WScript.Quit
End If

MsgBox name + ", " + age