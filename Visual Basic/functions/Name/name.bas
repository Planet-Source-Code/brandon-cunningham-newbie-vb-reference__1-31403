Attribute VB_Name = "Name"
Public Function namex(file As String) As String
filex = Split(file, "\")
x = UBound(filex)
filewext = filex(x)
namesplit = Split(filewext, ".")
firstx = 0
For i = LBound(namesplit) To UBound(namesplit) - 1
If firstx = 0 Then
namex = namesplit(i)
firstx = 1
Else
namex = namex & "." & namesplit(i)
End If
Next
End Function
