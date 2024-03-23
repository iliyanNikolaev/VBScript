' InputBox "message", "title", "input field", xposition, yposition
Option Explicit
Dim name
Dim age

Do
    name = InputBox("What is your name?", "Name")
If name="" Then
    MsgBox "Name cannot be empty!", vbExclamation, "Error"
Else
    Exit Do
End If
Loop

Do
    age = InputBox("What is your age?", "Age")
If Not IsNumeric(age) Then
    MsgBox "Invalid age!", vbExclamation, "Error"
ElseIf IsNumeric(age) Then
    Exit Do
End If
Loop 

MsgBox name + ", " + age