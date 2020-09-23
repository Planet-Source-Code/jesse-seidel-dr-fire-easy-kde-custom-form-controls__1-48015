VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KdeButton By Jesse Seidel"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin KButton.KdeButton KdeButton3 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KButton.KdeButton KdeButton2 
      Height          =   1695
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "KdeButton"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enable"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin KButton.KdeButton KdeButton1 
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
      _ExtentX        =   18865
      _ExtentY        =   14208
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Note: The gradient angle changes by size"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Long Example:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Big Sample:"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Sample Button:"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2400
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
KdeButton1.Enabled = True
End Sub

Private Sub Command2_Click()
KdeButton1.Enabled = False
End Sub

Private Sub Text1_Change()
KdeButton1.Caption = Text1.Text
End Sub
