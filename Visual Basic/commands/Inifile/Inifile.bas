Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function InifileWrite(IniFile As String, KeySection As String, KeyKey As String, KeyValue As String) As Boolean
    Dim lngResult As Long, strFileName As String, AppPath As String
    
    ' Get the full path to the ini-file.
    AppPath = App.Path
    If Right(AppPath, 1) = "\" Then
        strFileName = AppPath & InifileOpen
    Else
        strFileName = AppPath & "\" & InifileOpen
    End If

    ' Write to ini-file
    lngResult = WritePrivateProfileString(KeySection, KeyKey, KeyValue, strFileName)
    
    ' Check if the write was successful
    If lngResult = 0 Then
        ' If an error occured, return False.
        InifileWrite = False
        MsgBox "Internal Error"
    Else
        ' Return True (Successful write).
        InifileWrite = True
    End If
End Function

Public Function Iniread(KeySection As String, KeyKey As String) As String
Dim lngResult As Long, strFileName As String, AppPath As String
    Dim strResult As String * 50
    
    ' Get the full path to the ini-file.
    AppPath = App.Path
    If Right(AppPath, 1) = "\" Then
        strFileName = AppPath & InifileOpen
    Else
        strFileName = AppPath & "\" & InifileOpen
    End If
    
    ' Read the ini-file
    lngResult = GetPrivateProfileString(KeySection, KeyKey, strFileName, strResult, Len(strResult), strFileName)
    
    ' Check if the read was successful
    If lngResult = 0 Then
        ' If an error occured, return "error".
        Iniread = "error"
    Else
        ' Return the value.
        Iniread = Left(strResult, lngResult)
    End If
End Function
