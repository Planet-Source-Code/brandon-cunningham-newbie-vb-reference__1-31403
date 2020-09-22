Attribute VB_Name = "Module1"
Public Declare Function ShellExecute& Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal _
lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long)

Public Const SW_NORMAL = 1
Public Sub Shell(ShellOp As String, ShellFile As String, Perc1 As String)
  If ShellExecute(&O0, ShellOp, ShellFile, Perc1, vbNullString, _
  SW_NORMAL) < 33 Then
  End If
End Sub

