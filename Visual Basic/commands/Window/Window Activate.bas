Attribute VB_Name = "Module2"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Public Type POINTAPI
  x       As Long
  y       As Long
End Type

Public Type WINDOWPLACEMENT
  Length            As Long
  flags             As Long
  showCmd           As Long
  ptMinPosition     As POINTAPI
  ptMaxPosition     As POINTAPI
  rcNormalPosition  As RECT
End Type

Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4

Public Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" _
    (ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
    
Public Declare Function BringWindowToTop Lib "user32" _
  (ByVal hWnd As Long) As Long

Public Declare Function FindWindow Lib "user32" _
   Alias "FindWindowA" _
  (ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long
   
Public Declare Function GetWindowPlacement Lib "user32" _
   (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function SetForegroundWindow Lib "user32" _
   (ByVal hWnd As Long) As Long

Public Declare Function SetWindowPlacement Lib "user32" _
   (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long


Public Sub WindowActivateByName(sWindowTitle As String)

   Dim hWndCtlApp As Long
   Dim currWinP As WINDOWPLACEMENT
    
  'obtain the handle to the control app
   hWndCtlApp = FindWindow(sWindowTitle, vbNullString)
   
   If hWndCtlApp Then
   
     'prepare the WINDOWPLACEMENT type
      currWinP.Length = Len(currWinP)
   
      If GetWindowPlacement(hWndCtlApp, currWinP) > 0 Then
      
        'determine the window state
         If currWinP.showCmd = SW_SHOWMINIMIZED Then
      
           'minimized, so restore
            currWinP.Length = Len(currWinP)
            currWinP.flags = 0&
            currWinP.showCmd = SW_SHOWNORMAL
            Call SetWindowPlacement(hWndCtlApp, currWinP)
         
         Else
           
           'on screen, so assure visible
            Call SetForegroundWindow(hWndCtlApp)
            Call BringWindowToTop(hWndCtlApp)
         
         End If
      
      End If
   
   End If
   
End Sub


Public Sub WindowActivateByTitle(sWindowTitle As String)

   Dim hWndCtlApp As Long
   Dim currWinP As WINDOWPLACEMENT
    
  'obtain the handle to the control app
   hWndCtlApp = FindWindow(vbNullString, sWindowTitle)
   
   If hWndCtlApp Then
   
     'prepare the WINDOWPLACEMENT type
      currWinP.Length = Len(currWinP)
   
      If GetWindowPlacement(hWndCtlApp, currWinP) > 0 Then
      
        'determine the window state
         If currWinP.showCmd = SW_SHOWMINIMIZED Then
      
           'minimized, so restore
            currWinP.Length = Len(currWinP)
            currWinP.flags = 0&
            currWinP.showCmd = SW_SHOWNORMAL
            Call SetWindowPlacement(hWndCtlApp, currWinP)
         
         Else
           
           'on screen, so assure visible
            Call SetForegroundWindow(hWndCtlApp)
            Call BringWindowToTop(hWndCtlApp)
         
         End If
      
      End If
   
   End If
   
End Sub



