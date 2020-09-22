Attribute VB_Name = "List"
Public Function ListAssign(List_1 As Object, list_2 As Object)
Dim x As Variant
For x = 0 To List_1.ListCount - 1
list_2.AddItem List_1.List(x)
Next x
End Function

Public Function ListCopy(List_1 As Object)
Dim Count As Long
Dim Copy As String
For Count = 0 To List_1.ListCount - 1
If Count = 0 Then
Copy = List_1.List(Count)
Else
Copy = Copy & Chr(13) + Chr(10) & List_1.List(Count)
End If
Next Count
Clipboard.Clear
Clipboard.SetText Copy
End Function
Public Function ListLoadfile(List_1 As Object, File_1 As String)
List_1.Clear
Dim lstInput As String
Open File_1 For Input As #1
While Not EOF(1)
Input #1, lstInput$
List_1.AddItem lstInput$
Wend
Close #1
End Function
Public Function ListText(List_1 As Object) As String
Dim Count As Long
For Count = 0 To List_1.ListCount - 1
If Count = 0 Then
ListText = List_1.List(Count)
Else
ListText = Copy & Chr(13) + Chr(10) & List_1.List(Count)
End If
Next Count
End Function

