VERSION 5.00
Begin VB.UserControl UserControl3 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image Image6 
      Height          =   1005
      Left            =   0
      Picture         =   "UserControl3.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "UserControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = Image6.Width
UserControl.Height = Parent.Height
UserControl.Height = Parent.Height
Image6.Height = UserControl.Height
End Sub
