VERSION 5.00
Begin VB.UserControl UserControl2 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image Image3 
      Height          =   2250
      Left            =   0
      Picture         =   "UserControl2.ctx":0000
      Stretch         =   -1  'True
      Top             =   350
      Width           =   60
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   0
      Picture         =   "UserControl2.ctx":0096
      Top             =   0
      Width           =   60
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
UserControl.Width = Image3.Width
UserControl.Height = UserControl.Height
Image3.Height = UserControl.Height + 9999
End Sub
