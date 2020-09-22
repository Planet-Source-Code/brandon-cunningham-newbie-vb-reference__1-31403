Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Function ShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long
    Dim sShortPathName As String
    Dim iLen As Integer
    
    'Set up buffer area for API function cal
    '     l return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)
    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    ShortName = Left(sShortPathName, lRetVal)
End Function