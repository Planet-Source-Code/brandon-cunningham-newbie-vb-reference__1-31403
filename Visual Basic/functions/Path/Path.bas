Attribute VB_Name = "Module1"
Public Function Path(ByVal fname As String, Optional delim As String = "\", Optional keeplast As Boolean = True) As String
    Dim outstr As String
    Dim llen As Long
    llen = InStrRev(fname, delim)


    If (Not keeplast) Then
        llen = llen - 1
    End If


    If (llen > 0) Then
        BasePath = Mid(fname, 1, llen)
    Else
        BasePath = fname
    End If
End Function
