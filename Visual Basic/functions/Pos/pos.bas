Attribute VB_Name = "Module1"
Public Function Pos(string1 As String, string2 As String) As String
blah = string2
blen = Len(blah)
blahx = blah
i = 1
Do While Not i = blen
xleft = Left(blahx, i)
xright = Right(xleft, 1)
If xright = string1 Then
Pos = i
Exit Function
Else
i = i + 1
End If
Loop
End Function
