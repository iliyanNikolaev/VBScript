Dim person
Set person = CreateObject("Scripting.Dictionary")
person.Add "Name", "Ilich"
person.Add "Age", 26
person.Add "City", "Pleven&&Trastenik"
person("City") = "Sofia"
' MsgBox(person("City"))

If person.Exists("Name") Then
    MsgBox "is exist"
End If

For Each key in person.Keys
    MsgBox(person(key))
Next
