Sub SortArray()
Dim arr() As Variant
Dim i As Long, j As Long
Dim temp As Variant
arr = Array(5, 3, 8, 1, 2)
For i = LBound(arr) To UBound(arr) - 1
For j = i + 1 To UBound(arr)
If arr(i) > arr(j) Then
temp = arr(i)
arr(i) = arr(j)
arr(j) = temp
End If
Next j
Next i
For i = LBound(arr) To UBound(arr)
Debug.Print arr(i)
Next i
End Sub
Sub FilterArray()
Dim arr() As Variant
Dim result() As Variant
Dim i As Long, count As Long
arr = Array(5, 3, 8, 1, 2)
ReDim result(0)
count = 0
For i = LBound(arr) To UBound(arr)
If arr(i) > 2 Then
ReDim Preserve result(count)
result(count) = arr(i)
count = count + 1
End If
Next i
For i = LBound(result) To UBound(result)
Debug.Print result(i)
Next i
End Sub
Sub SumArray()
Dim arr() As Variant
Dim total As Double
Dim i As Long
arr = Array(5, 3, 8, 1, 2)
total = 0
For i = LBound(arr) To UBound(arr)
total = total + arr(i)
Next i
Debug.Print total
End Sub
Sub ReverseArray()
Dim arr() As Variant
Dim i As Long
arr = Array(5, 3, 8, 1, 2)
For i = LBound(arr) To (UBound(arr) - LBound(arr)) / 2
Dim temp As Variant
temp = arr(i)
arr(i) = arr(UBound(arr) - i)
arr(UBound(arr) - i) = temp
Next i
For i = LBound(arr) To UBound(arr)
Debug.Print arr(i)
Next i
End Sub
