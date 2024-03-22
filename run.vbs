Option Explicit
Dim obj

Set obj=CreateObject("wscript.shell")
obj.run "mspaint.exe"
obj.run "C:\Users\ilich\Desktop\test.txt"
obj.run """C:\Users\ilich\Desktop\New folder""" 'when file name or folder name contains spaces