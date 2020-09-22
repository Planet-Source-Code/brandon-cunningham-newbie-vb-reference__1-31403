Attribute VB_Name = "Module1"
Public Function textoflist(Listx As Object) As String
Dim Count As Long
For Count = 0 To Listx.ListCount - 1
If Count = 0 Then
textoflist = Listx.List(Count)
Else
textoflist = textoflist & Chr(13) + Chr(10) & Listx.List(Count)
End If
Next Count
End Function
