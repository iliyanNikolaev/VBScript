Dim numbers(4)
numbers(0) = 11
numbers(1) = 22
numbers(2) = 33
numbers(3) = 44
numbers(4) = 55
For i = LBound(numbers) To UBound(numbers) Step 2
    MsgBox(numbers(i))
Next