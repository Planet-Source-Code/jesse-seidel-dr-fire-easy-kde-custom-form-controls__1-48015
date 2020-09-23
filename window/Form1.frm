VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl4 UserControl41 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   3
      Top             =   2265
      Width           =   7185
      _extentx        =   12674
      _extenty        =   212
      font            =   "Form1.frx":0000
   End
   Begin Project1.UserControl3 UserControl31 
      Align           =   4  'Align Right
      Height          =   2385
      Left            =   7140
      TabIndex        =   2
      Top             =   300
      Width           =   45
      _extentx        =   79
      _extenty        =   4207
   End
   Begin Project1.UserControl2 UserControl21 
      Align           =   3  'Align Left
      Height          =   1965
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   60
      _extentx        =   106
      _extenty        =   3466
   End
   Begin Project1.UserControl1 UserControl11 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      _extentx        =   12674
      _extenty        =   529
      font            =   "Form1.frx":002C
   End
   Begin Project1.kCombo kCombo1 
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   6735
      _extentx        =   11880
      _extenty        =   529
      font            =   "Form1.frx":0058
      list0           =   "KList - By Jesse Seidel"
      listindex       =   -1
      text            =   "I would like to"
   End
   Begin Project1.KdeButton KdeButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      font            =   "Form1.frx":0080
      caption         =   "&Done"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Note: Double click the item you wish to have selected"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
kCombo1.AddItem "Enter the program"
kCombo1.AddItem "Exit"
kCombo1.AddItem "Read the terms and conditions before viewing the program"
kCombo1.RemoveItem (0)
End Sub

Private Sub KdeButton1_Click()
If kCombo1.Text = "Exit" Then End
If kCombo1.Text = "Enter the program" Then Form2.Visible = True
If kCombo1.Text = "Read the terms and conditions before viewing the program" Then Form3.Visible = True
End Sub
