VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ColorComboControl 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   ScaleHeight     =   330
   ScaleWidth      =   2430
   Begin MSComctlLib.ImageList ColorList 
      Left            =   2565
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":0278
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":04F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":0768
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":09E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":0ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":13C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":1638
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":18B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":1B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":1DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":2018
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":2290
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCombo.ctx":2508
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo ColorCombo 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "ColorList"
   End
End
Attribute VB_Name = "ColorComboControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Click()

Public Property Get Color() As String
Color = ColorCombo.SelectedItem
End Property

Private Sub ColorCombo_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
 ColorCombo.ComboItems.Add , , "Black", 1
        ColorCombo.ComboItems.Add , , "Maroon", 2
        ColorCombo.ComboItems.Add , , "Green", 3
        ColorCombo.ComboItems.Add , , "Olive", 4
        ColorCombo.ComboItems.Add , , "Navy", 5
        ColorCombo.ComboItems.Add , , "Purple", 6
        ColorCombo.ComboItems.Add , , "Teal", 7
        ColorCombo.ComboItems.Add , , "Gray", 8
        ColorCombo.ComboItems.Add , , "Silver", 9
        ColorCombo.ComboItems.Add , , "Red", 10
        ColorCombo.ComboItems.Add , , "Lime", 11
        ColorCombo.ComboItems.Add , , "Yellow", 12
        ColorCombo.ComboItems.Add , , "Blue", 13
        ColorCombo.ComboItems.Add , , "Fuchsia", 14
       ColorCombo.ComboItems.Add , , "Aqua", 15
       ColorCombo.ComboItems.Add , , "White", 16
       ColorCombo.SelectedItem = ColorCombo.ComboItems(1)
End Sub

Private Sub UserControl_Resize()
ColorCombo.Top = 0
ColorCombo.Left = 0
ColorCombo.Width = UserControl.Width
ColorCombo.Height = UserControl.Height
End Sub
