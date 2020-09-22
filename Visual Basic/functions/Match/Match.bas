Attribute VB_Name = "Module1"
Public Function FirstMatch(listx As Object, stringx As String)
l = listx.ListIndex

Do While Not l = listx.ListCount

Item = listx.List(l)
x = Split(Item, stringx)

If UBound(x) > 0 Then
listx.ListIndex = l
Exit Function
Else
End If
l = l + 1
Loop
End Function

Public Function NextMatch(listx As Object, stringx As String)
l = listx.ListIndex + 1

Do While Not l = listx.ListCount

Item = listx.List(l)
x = Split(Item, stringx)

If UBound(x) > 0 Then
listx.ListIndex = l
Exit Function
Else
End If
l = l + 1
Loop
End Function

Public Function ExactMatch(List As Object, String1 As String)
l = 0
Do While Not l = List.ListCount
Item = List.List(l)
l = l + 1
If Item = String1 Then
List.ListIndex = l - 1
Exit Function
Else
End If
Loop
End Function
