Attribute VB_Name = "Module1"
Public Function strins(string1 As String, pos1 As Integer, insert1 As String) As String
leftstr = Left(string1, pos1)
dox = Len(string1) - pos1
rightstr = Right(string1, dox)
strins = leftstr & insert1 & rightstr
End Function
