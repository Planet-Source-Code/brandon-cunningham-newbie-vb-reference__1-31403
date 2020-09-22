Attribute VB_Name = "Module1"
Rem My Very First Module!! Yay!
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Const WM_PASTE = &H302

Public Const WM_ACTIVATE = &H6

Public Const WM_Close = &H10

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Function WindowSend(Window As String, Control As String, Send As String)
Dim x As String
x = Clipboard.GetText
Clipboard.SetText Send
Window = FindWindow(Window, vbNullString)
Control = FindWindowEx(Window, 0&, Control, vbNullString)
Call SendMessage(Control, WM_PASTE, 0&, "")
Clipboard.SetText x
End Function

Public Function WindowSendToTitle(Window As String, Control As String, Send As String)
Dim x As String
x = Clipboard.GetText
Clipboard.SetText Send
Window = FindWindow(vbNullString, Window)
Control = FindWindowEx(Window, 0&, Control, vbNullString)
Call SendMessage(Control, WM_PASTE, 0&, "")
Clipboard.SetText x
End Function
Public Function WindowClose(Window As String)
Window = FindWindow(Window, vbNullString)
x = SendMessage(Window, WM_Close, 0, 0)
End Function

Public Function WindowCloseByTitle(Window As String)
Window = FindWindow(vbNullString, Window)
x = SendMessage(Window, WM_Close, 0, 0)
End Function

Rem failed experiment
Public Function WindowSendType(Window As String, Control As String, Send As String)
Dim Win2 As Long
Dim CTRL2 As Long
Dim SendIt As Long
Win2& = FindWindow(Window, vbNullString)
CTRL2& = FindWindowEx(Win2&, 0&, Control, vbNullString)
Call SendMessageByString(CTRL2&, WM_SETTEXT, 0&, Send)
End Function

