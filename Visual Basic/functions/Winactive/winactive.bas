Attribute VB_Name = "MyMod"
Private Declare Function M_GetActiveWindow Lib "user32" _
Alias "GetActiveWindow" () As Long

Public Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" _
    (ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
    
Function GetActiveWindow() As Long
Dim xHwnd As Long
xHwnd = M_GetActiveWindow()
GetActiveWindow = xHwnd
End Function
