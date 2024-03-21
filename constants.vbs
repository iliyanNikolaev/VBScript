Dim x
x=MsgBox("Hello vbs",_
vbAbortRetryIgnore+vbExclamation+vbDefaultButton2+vbSystemModal,_
"Constants Example")

If x=3 Then MsgBox "your choice is Abort"
If x=4 Then MsgBox "your choice is Retry"
If x=5 Then MsgBox "your choice is Ignore", vbCritical