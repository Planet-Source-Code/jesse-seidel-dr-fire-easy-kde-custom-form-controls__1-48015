VERSION 5.00
Begin VB.UserControl KdeButton 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   ScaleHeight     =   1890
   ScaleWidth      =   3600
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      Picture         =   "UserControl1.ctx":0312
      ScaleHeight     =   1815
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "KdeButton"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "KdeButton"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Image7 
         Height          =   45
         Left            =   2160
         Picture         =   "UserControl1.ctx":0378
         Stretch         =   -1  'True
         Top             =   120
         Width           =   30
      End
      Begin VB.Image Image8 
         Height          =   60
         Left            =   960
         Picture         =   "UserControl1.ctx":03D2
         Top             =   0
         Width           =   60
      End
      Begin VB.Image Image6 
         Height          =   30
         Left            =   120
         Picture         =   "UserControl1.ctx":0444
         Stretch         =   -1  'True
         Top             =   0
         Width           =   30
      End
      Begin VB.Image Image5 
         Height          =   60
         Left            =   120
         Picture         =   "UserControl1.ctx":0496
         Top             =   240
         Width           =   60
      End
      Begin VB.Image Image4 
         Height          =   30
         Left            =   50
         Picture         =   "UserControl1.ctx":0508
         Stretch         =   -1  'True
         Top             =   250
         Width           =   60
      End
      Begin VB.Image Image3 
         Height          =   45
         Left            =   0
         Picture         =   "UserControl1.ctx":0562
         Top             =   0
         Width           =   60
      End
      Begin VB.Image Image1 
         Height          =   45
         Left            =   0
         Picture         =   "UserControl1.ctx":05C8
         Top             =   240
         Width           =   60
      End
      Begin VB.Image Image2 
         Height          =   195
         Left            =   0
         Picture         =   "UserControl1.ctx":0609
         Stretch         =   -1  'True
         Top             =   45
         Width           =   60
      End
   End
End
Attribute VB_Name = "KdeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ClsGradient As New CGradient
'Event Declarations:
Event Click() 'MappingInfo=Label2,Label2,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Label2,Label2,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label2,Label2,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label2,Label2,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label2,Label2,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."


Private Sub Label2_Click()
    RaiseEvent Click

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    With ClsGradient
        .Angle = -9999
        .Color1 = RGB(255, 250, 255)
        .Color2 = RGB(200, 200, 200)
        .Draw Picture1
    End With
    Picture1.Refresh
Label1(0).ForeColor = vbBlack
Text1.Text = Label1(0).Caption
Label1(0).Caption = " " & Text1.Text
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
Image7.Top = Image8.Height
    With ClsGradient
        .Angle = -9999
        .Color2 = RGB(255, 250, 255)
        .Color1 = RGB(200, 200, 200)
        .Draw Picture1
    End With
    Picture1.Refresh
Label1(0).ForeColor = vbBlack
Label1(0).Caption = Text1.Text
End Sub

Private Sub Timer1_Timer()
Label1(1).Caption = Label1(0).Caption
End Sub

Private Sub UserControl_Resize()
Label2.Height = UserControl.Height
Label2.Width = UserControl.Width
Picture1.Height = UserControl.Height
Picture1.Width = UserControl.Width
Label1(0).Top = UserControl.Height / 2 - 100
Label1(0).Width = UserControl.Width
Label1(1).Top = Label1(0).Top
Label1(1).Width = Label1(0).Width
Label1(1).Left = Label1(0).Left + 10
Label1(1).Caption = Label1(0).Caption
Image1.Top = UserControl.Height - Image1.Height
Image2.Height = UserControl.Height - Image1.Height - Image1.Height
Image4.Width = UserControl.Width - Image1.Width - Image1.Width + 10
Image4.Top = UserControl.Height - Image4.Height
Image5.Top = UserControl.Height - Image5.Height
Image5.Left = UserControl.Width - Image5.Width
Image6.Left = 0 + Image3.Width
Image6.Width = UserControl.Width - Image5.Width
Image8.Left = UserControl.Width - Image8.Width
Image7.Left = UserControl.Width - Image7.Width
Image7.Height = UserControl.Height - Image8.Height - Image5.Height
Image7.Top = Image8.Height
    With ClsGradient
        .Angle = -9999
        .Color2 = RGB(255, 250, 255)
        .Color1 = RGB(200, 200, 200)
        .Draw Picture1
    End With
    Picture1.Refresh

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1(0),Label1(0),-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1(0).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1(0).ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
If New_Enabled = False Then
Label1(1).Visible = True
Else
Label1(1).Visible = False
End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1(0),Label1(0),-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1(0).Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub Label2_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1(0),Label1(0),-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1(0).Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1(0).Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1(0).ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label1(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1(0).Caption = PropBag.ReadProperty("Caption", "KButton")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1(0).ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", Label1(0).Caption, "KButton")
End Sub

