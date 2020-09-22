Attribute VB_Name = "Substr"
Public Function xsubstr(string1 As String, pos As Integer, countback As Integer) As String
leftstr = Left(string1, pos)
xsubstr = Right(leftstr, countback)
End Function
