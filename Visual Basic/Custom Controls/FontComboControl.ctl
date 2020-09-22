VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FontComboControl 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   ScaleHeight     =   315
   ScaleWidth      =   3345
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageCombo FontCombo 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FontComboControl.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FontComboControl.ctx":0338
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FontComboControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event FontClick()

Public Property Get curfont() As String
curfont = FontCombo.SelectedItem
End Property

Public Property Let TheFont(font As String)
FontCombo.Text = font
End Property
Private Sub FontCombo_Click()
RaiseEvent FontClick
End Sub

Private Sub UserControl_Initialize()
FontCombo.ImageList = ImageList1
  Dim i As Long
  For i = 0 To Screen.FontCount - 1
  List1.AddItem CStr(Screen.Fonts(i))
  Next i

 Dim savei As Long
   For i = 0 To List1.ListCount - 1
   If List1.List(i) = "Times New Roman" Then
  'MsgBox List1.List(i)
  savei = i
  Else
  End If
      FontCombo.ComboItems.Add , , List1.List(i), 1
        Next 'i
   FontCombo.SelectedItem = FontCombo.ComboItems(savei + 1)
End Sub

Private Sub UserControl_Resize()
FontCombo.Top = 0
FontCombo.Left = 0
FontCombo.Width = UserControl.Width
FontCombo.Height = UserControl.Height
End Sub


