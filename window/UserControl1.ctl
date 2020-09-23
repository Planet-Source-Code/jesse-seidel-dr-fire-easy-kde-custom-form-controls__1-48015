VERSION 5.00
Begin VB.UserControl UserControl1 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ScaleHeight     =   3600
   ScaleWidth      =   6015
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Text            =   "down"
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   4200
      Picture         =   "UserControl1.ctx":0000
      Top             =   0
      Width           =   540
   End
   Begin VB.Image Image2 
      Height          =   245
      Left            =   70
      Picture         =   "UserControl1.ctx":08B2
      Stretch         =   -1  'True
      Top             =   50
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6015
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   0
      Picture         =   "UserControl1.ctx":111F
      Top             =   0
      Width           =   60
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Minimize"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KDE Window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   380
      TabIndex        =   0
      Top             =   45
      Width           =   1020
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   360
      Picture         =   "UserControl1.ctx":1251
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1140
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   -240
      Picture         =   "UserControl1.ctx":1383
      Top             =   480
      Width           =   14160
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Picture         =   "UserControl1.ctx":F105
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4740
   End
   Begin VB.Menu d 
      Caption         =   "d"
      Begin VB.Menu min 
         Caption         =   "&Minimize"
         Shortcut        =   ^M
      End
      Begin VB.Menu sml 
         Caption         =   "Make &Small / &Restore"
         Shortcut        =   ^R
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu cls 
         Caption         =   "&Close"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Default Property Values:
Dim InitHeight As Integer
'Const m_def_Icon = 0
'Const m_def_BackColor = 0
'Const m_def_ForeColor = 0
'Const m_def_Enabled = 0
'Const m_def_BackStyle = 0
'Const m_def_BorderStyle = 0
'Property Variables:
'Dim m_ForeColor As OLE_COLOR
'Dim m_Icon As Variant
'Dim m_BackColor As Long
'Dim m_ForeColor As Long
'Dim m_Enabled As Boolean
'Dim m_Font As Font
'Dim m_BackStyle As Integer
'Dim m_BorderStyle As Integer
'Event Declarations:
Event About() 'MappingInfo=Command1,Command1,-1,Click
Attribute About.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
'Event Click()
'Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1
Dim OnTop As Integer





Private Sub cls_Click()
Unload Parent
End Sub

Private Sub d2_Click()
Parent.WindowState = 2
End Sub

Private Sub Image2_Click()
PopupMenu d, , 60, 300
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Parent.hwnd, &H112, &HF012&, 0
End Sub

Private Sub Label3_Click()
Parent.WindowState = 1
End Sub

Private Sub Label4_Click()
Unload Parent
End Sub


Private Sub Label5_DblClick()

If Text1.Text = "down" Then
InitHeight = Parent.Height
Parent.Height = 465
Text1.Text = "up"
Else
Parent.Height = InitHeight
Text1.Text = "down"
End If
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Parent.hwnd, &H112, &HF012&, 0
End Sub

Private Sub min_Click()
Parent.WindowState = 1
End Sub

Private Sub sml_Click()
If Text1.Text = "down" Then
InitHeight = Parent.Height
Parent.Height = 465
Text1.Text = "up"
Else
Parent.Height = InitHeight
Text1.Text = "down"
End If
End Sub

Private Sub UserControl_Resize()
Label5.Height = Image4.Height
Label5.Width = UserControl.Width
UserControl.Height = 300
Image9.Left = UserControl.Width - Image9.Width
Image1.Width = UserControl.Width
Label1.Width = UserControl.Width
Label1.Top = 50
Line1.X1 = Image9.Left + Image9.Width
Line1.X2 = Image9.Left + Image9.Width
Label4.Left = UserControl.Width - Label4.Width - 20
Label3.Left = Label4.Left - Label3.Width
Image5.Width = Label1.Width
Image4.Visible = True
OnTop = False
End Sub
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,0
'Public Property Get BackColor() As Long
'    BackColor = m_BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As Long)
'    m_BackColor = New_BackColor
'    PropertyChanged "BackColor"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=8,0,0,0
''Public Property Get ForeColor() As Long
''    ForeColor = m_ForeColor
''End Property
''
''Public Property Let ForeColor(ByVal New_ForeColor As Long)
''    m_ForeColor = New_ForeColor
''    PropertyChanged "ForeColor"
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=0,0,0,0
''Public Property Get Enabled() As Boolean
''    Enabled = m_Enabled
''End Property
''
''Public Property Let Enabled(ByVal New_Enabled As Boolean)
''    m_Enabled = New_Enabled
''    PropertyChanged "Enabled"
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=6,0,0,0
''Public Property Get Font() As Font
''    Set Font = m_Font
''End Property
''
''Public Property Set Font(ByVal New_Font As Font)
''    Set m_Font = New_Font
''    PropertyChanged "Font"
''End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get BackStyle() As Integer
'    BackStyle = m_BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    m_BackStyle = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=7,0,0,0
''Public Property Get BorderStyle() As Integer
''    BorderStyle = m_BorderStyle
''End Property
''
''Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
''    m_BorderStyle = New_BorderStyle
''    PropertyChanged "BorderStyle"
''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_BackColor = m_def_BackColor
'    m_ForeColor = m_def_ForeColor
'    m_Enabled = m_def_Enabled
'    Set m_Font = Ambient.Font
'    m_BackStyle = m_def_BackStyle
'    m_BorderStyle = m_def_BorderStyle
'    m_ForeColor = m_def_ForeColor
'    m_Icon = m_def_Icon
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
'    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
'    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
'    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
'    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    Label1.Caption = PropBag.ReadProperty("Caption", "KDE Window")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set Picture = PropBag.ReadProperty("BarPicture", Nothing)
'    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
'    m_Icon = PropBag.ReadProperty("Icon", m_def_Icon)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set Picture = PropBag.ReadProperty("Icon", Nothing)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
'    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
'    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
'    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
'    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
'    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "KDE Window")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BarPicture", Picture, Nothing)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Image2,Image2,-1,Picture
'Public Property Get Picture() As Picture
'    Set Picture = Image2.Picture
'End Property
'
'Public Property Set Picture(ByVal New_Picture As Picture)
'    Set Image2.Picture = New_Picture
'    PropertyChanged "Picture"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Image9,Image9,-1,ForeColor
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = Image9.ForeColor
'End Property
'
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)


'    Image9.ForeColor() = New_ForeColor
'    PropertyChanged "ForeColor"
'    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
'    Call PropBag.WriteProperty("Icon", m_Icon, m_def_Icon)

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
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get BarPicture() As Picture
Attribute BarPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set BarPicture = Image1.Picture
End Property

Public Property Set BarPicture(ByVal New_BarPicture As Picture)
    Set Image1.Picture = New_BarPicture
    PropertyChanged "BarPicture"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=10,0,0,
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = m_ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    m_ForeColor = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get Icon() As Variant
'    Icon = m_Icon
'End Property
'
'Public Property Let Icon(ByVal New_Icon As Variant)
'    m_Icon = New_Icon
'    PropertyChanged "Icon"
'End Property
'
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
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image2,Image2,-1,Picture
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Icon = Image2.Picture
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set Image2.Picture = New_Icon
    PropertyChanged "Icon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property


Private Sub Command1_Click()
    RaiseEvent About
    MsgBox "Made by Jesse Seidel", vbOKOnly, "About..."
End Sub

