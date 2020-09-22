VERSION 5.00
Begin VB.UserControl LabeledText 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   675
   ScaleWidth      =   4800
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   250
      Width           =   4515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Default"
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4785
   End
End
Attribute VB_Name = "LabeledText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Get Caption() As String
Caption = Frame1.Caption
End Property

Public Property Let Caption(Caption As String)
Frame1.Caption = Caption
End Property

Public Property Get Text() As String
Text = Text1.Text
End Property

Public Property Let Text(Text As String)
Text1.Text = Text
End Property


Private Sub UserControl_Resize()
On Error GoTo er:
UserControl.Height = 675
Frame1.Top = 0
Frame1.Left = 0
Frame1.Width = UserControl.Width
Frame1.Height = UserControl.Height
Text1.Top = 250
Text1.Left = 135
Text1.Height = UserControl.Height - 400
Text1.Width = UserControl.Width - 300
er:
Exit Sub
End Sub
