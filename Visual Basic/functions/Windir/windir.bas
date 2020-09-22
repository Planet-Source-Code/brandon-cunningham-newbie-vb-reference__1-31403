Attribute VB_Name = "MyMod"
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Function WinDir(Optional ByVal AddSlash As Boolean = False) As String
    Dim t As String * 255
    Dim i As Long
    i = GetWindowsDirectory(t, Len(t))
    WinDir = Left(t, i)


    If (AddSlash = True) And (Right(WinDir, 1) <> "\") Then
        WinDir = WinDir & "\"
    ElseIf (AddSlash = False) And (Right(WinDir, 1) = "\") Then
        WinDir = Left(WinDir, Len(WinDir) - 1)
    End If
End Function


Public Function SysDir(Optional ByVal AddSlash As Boolean = False) As String
    Dim t As String * 255
    Dim i As Long
    i = GetSystemDirectory(t, Len(t))
    SysDir = Left(t, i)


    If (AddSlash = True) And (Right(SysDir, 1) <> "\") Then
        SysDir = SysDir & "\"
    ElseIf (AddSlash = False) And (Right(SysDir, 1) = "\") Then
        SysDir = Left(SysDir, Len(SysDir) - 1)
    End If
End Function
