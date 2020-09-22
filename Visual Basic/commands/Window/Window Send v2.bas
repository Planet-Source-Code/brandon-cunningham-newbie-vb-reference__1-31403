Attribute VB_Name = "Module1"
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Const WM_CHAR = &H102

Public Function WindowSend(Window As String, Control As String, Send As String)
Window = FindWindow(Window, vbNullString)
Control = FindWindowEx(Window, 0&, Control, vbNullString)
Dim numLoop, strLeft As String
  Dim strRight As String
    For numLoop = 1 To Len(Send)
    strLeft = Left(Send, numLoop)
    strRight = Right(strLeft, 1)
       Pause 0.01
       Call SendMessageByNum(Control, WM_CHAR, Asc(strRight), 0)
  Next numLoop
End Function

Public Function WindowSendToTitle(Window As String, Control As String, Send As String)
Window = FindWindow(vbNullString, Window)
Control = FindWindowEx(Window, 0&, Control, vbNullString)
Dim numLoop, strLeft As String
  Dim strRight As String
    For numLoop = 1 To Len(Send)
    strLeft = Left(Send, numLoop)
    strRight = Right(strLeft, 1)
       Pause 0.01
       Call SendMessageByNum(Control, WM_CHAR, Asc(strRight), 0)
  Next numLoop
End Function

Sub Pause(Duration)
  Dim numTime
  numTime = Timer
  Do While Timer - numTime < Duration
    DoEvents
  Loop
End Sub
