VERSION 5.00
Begin VB.UserControl FrameHide 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   3045
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   170
      Left            =   0
      MouseIcon       =   "FrameHide.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   170
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   75
         X2              =   75
         Y1              =   40
         Y2              =   120
      End
      Begin VB.Line Line1 
         X1              =   50
         X2              =   130
         Y1              =   80
         Y2              =   80
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   165
         Left            =   0
         Top             =   0
         Width           =   170
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1935
      Left            =   60
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape shape3 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   2175
      Left            =   50
      Top             =   100
      Width           =   2415
   End
End
Attribute VB_Name = "FrameHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ClickBoxRoll()
Event Show()
Event Resize()
'Default Property Values:
Const m_def_isMinimize = False
Const m_def_MaximizeSize = 200
'Const m_def_Minimize = False
'Property Variables:
Private m_isMinimize As Boolean
Private m_MaximizeSize As Long
Private m_BackColor As OLE_COLOR
Private MisMinimize As Boolean

Private Sub Frame2_Click()
If MisMinimize Then
    Maximize
Else
    Minimize
End If
RaiseEvent ClickBoxRoll
End Sub

Public Sub Minimize()
isMinimize = True
End Sub

Private Sub mMinimize()
If MisMinimize Then Exit Sub
MisMinimize = True
'm_MaximizeSize = UserControl.Height
Line2.Visible = True
UserControl.Height = 200
Shape2.Height = UserControl.Height - 190
shape3.Height = UserControl.Height - 180
'UserControl_Resize
End Sub

Public Sub Maximize()
isMinimize = False
End Sub

Private Sub mMaximize()
If MisMinimize = False Then Exit Sub
MisMinimize = False
UserControl.Height = m_MaximizeSize
Line2.Visible = False
'UserControl_Resize
End Sub

Private Sub Label1_Click()
Frame2_Click
End Sub


Private Sub UserControl_InitProperties()

Label1.MouseIcon = Frame2.MouseIcon

Caption = UserControl.Extender.Name

UserControl.BackColor() = UserControl.Ambient.BackColor
'Shape4.BackColor = UserControl.BackColor()
Frame3.BackColor() = UserControl.BackColor()
Shape1.BackColor() = UserControl.BackColor()

m_MaximizeSize = m_def_MaximizeSize
m_isMinimize = m_def_isMinimize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
If isMinimize = False Then
m_MaximizeSize = UserControl.Height 'm_def_MaximizeSize
shape3.Height = UserControl.Height - 150
Shape2.Height = UserControl.Height - 145
Else
UserControl.Height = 200
Shape2.Height = UserControl.Height - 190
shape3.Height = UserControl.Height - 180
Line2.Visible = True
MisMinimize = True
End If
shape3.Width = UserControl.Width - 93
Shape2.Width = UserControl.Width - 90
RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption = New_Caption
    If Label1.Caption <> "" Then
    Frame3.Width = Len(Label1.Caption) * 120
    Label1.Width = Frame3.Width '- 160
    End If
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    
    Shape1.BorderColor = Label1.ForeColor
    Line1.BorderColor = Label1.ForeColor
    Line2.BorderColor = Label1.ForeColor
    
    PropertyChanged "ForeColor"
End Property


Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.Caption = PropBag.ReadProperty("Caption", "")
    
    If Label1.Caption <> "" Then
    Frame3.Width = Len(Label1.Caption) * 120
    Label1.Width = Frame3.Width '- 160
    End If
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    
    Shape1.BorderColor = Label1.ForeColor
    Line1.BorderColor = Label1.ForeColor
    Line2.BorderColor = Label1.ForeColor
    
    Label1.FontName = PropBag.ReadProperty("FontName", "arial")
    Label1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Label1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Label1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'    Shape4.BackColor() = UserControl.BackColor
    Frame3.BackColor() = UserControl.BackColor
    Shape1.BackColor() = UserControl.BackColor
    
    m_MaximizeSize = PropBag.ReadProperty("MaximizeSize", m_def_MaximizeSize)
    m_isMinimize = PropBag.ReadProperty("isMinimize", m_def_isMinimize)
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", Label1.Caption, "")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("FontName", Label1.FontName, "")
    Call PropBag.WriteProperty("FontItalic", Label1.FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", Label1.FontBold, 0)
    Call PropBag.WriteProperty("FontUnderline", Label1.FontUnderline, 0)
'    Call PropBag.WriteProperty("BackColor", Shape4.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("MaximizeSize", m_MaximizeSize, m_def_MaximizeSize)
    Call PropBag.WriteProperty("isMinimize", m_isMinimize, m_def_isMinimize)
    
End Sub

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Label1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Label1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Label1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Label1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Label1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Label1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = shape3.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
'    Shape4.BackColor() = New_BackColor
    Frame3.BackColor() = New_BackColor
    Shape1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get MaximizeSize() As Long
    MaximizeSize = m_MaximizeSize
End Property

Public Property Let MaximizeSize(ByVal New_MaximizeSize As Long)
    m_MaximizeSize = New_MaximizeSize
    PropertyChanged "MaximizeSize"
End Property

Public Property Get isMinimize() As Boolean
    isMinimize = m_isMinimize
End Property

Public Property Let isMinimize(ByVal New_isMinimize As Boolean)
'    If Ambient.UserMode = False Then Err.Raise 387
    m_isMinimize = New_isMinimize
    
    If m_isMinimize Then
    mMinimize
    Else
    mMaximize
    End If
    
    PropertyChanged "isMinimize"
End Property

