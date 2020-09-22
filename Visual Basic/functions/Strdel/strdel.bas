Attribute VB_Name = "Module1"
Public Function strdel(string1 As String, pos1 As Integer, pos2 As Integer) As String
leftstr = Left(string1, pos1)
dox = Len(string1) - pos2
rightstr = Right(string1, dox)
strdel = leftstr & rightstr
End Function
