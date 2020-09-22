Attribute VB_Name = "modBrowseForFolder"
Option Explicit

'API Declares
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

'API Constants
Private Const MAX_PATH = 260
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_STATUSTEXT = 4

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
'Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)
'Private Const BFFM_SETSELECTIONW = (WM_USER + 103)
'Private Const BFFM_SETSTATUSTEXTW = (WM_USER + 104)

'BrowseInfo Type
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'Private Variables
Private m_sDefaultFolder As String

'Displays the Windows 95 BrowseForFolder dialog box
Public Function Dirdlg(DefaultFolder As String, Optional Parent As Long = 0, Optional Caption As String = "") As String
    Dim bi As BrowseInfo
    Dim sResult As String, nResult As Long

    bi.hwndOwner = Parent
    bi.pIDLRoot = 0
    bi.pszDisplayName = String$(MAX_PATH, Chr$(0))
    If Len(Caption) > 0 Then
        bi.lpszTitle = Caption
    End If
    bi.ulFlags = BIF_RETURNONLYFSDIRS   'Or BIF_STATUSTEXT
    bi.lpfn = GetAddress(AddressOf BrowseCallbackProc)
    bi.lParam = 0
    bi.iImage = 0
    'Set local default folder string
    '(will be set in callback after dialog initializes)
    m_sDefaultFolder = DefaultFolder
    'Call API
    nResult = SHBrowseForFolder(bi)
    'Get result if successful
    If nResult <> 0 Then
        sResult = String(MAX_PATH, 0)
        If SHGetPathFromIDList(nResult, sResult) Then
            Dirdlg = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
        End If
        'Free memory allocated by SHBrowseForFolder
        CoTaskMemFree nResult
    End If
End Function

'CAUTION: This function is called by the system to intialize the SHBrowseForFolder
'dialog box. Attempting to set breakpoints or adding other debugging code to this
'routine may cause unexpected problems.
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Select Case uMsg
        Case BFFM_INITIALIZED
            'Note: This code was crashing VB when the default folder was empty
            If Len(m_sDefaultFolder) > 0 Then
                'Set default folder when dialog has initialized
                SendMessage hwnd, BFFM_SETSELECTIONA, True, ByVal m_sDefaultFolder
            End If
    End Select
End Function

'Return the argument to workaround limitations of AddressOf operator
Private Function GetAddress(nAddress As Long) As Long
    GetAddress = nAddress
End Function
