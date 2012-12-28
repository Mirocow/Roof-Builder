VERSION 5.00
Begin VB.UserControl SplitHV 
   Alignable       =   -1  'True
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   MousePointer    =   99  'Custom
   PaletteMode     =   4  'None
   ScaleHeight     =   2445
   ScaleWidth      =   285
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   135
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   2280
   End
End
Attribute VB_Name = "SplitHV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************
'© MDinc; 2004
'dmms@narod.ru
'***********************************

Option Explicit

Dim bMoving As Boolean

'Dim R1 As RECT, R2 As RECT ' Split control
'Dim P1 As RECT ' Parent

'Default Property Values:
Const m_def_PositionSave = True
Const m_def_SplitWidth = 75
'Const m_def_ResizeLimit = 1000
'Const m_def_BackColor = vbButtonFace

Const m_def_SplitLimit = 500
Const m_def_LimitBorder = 50
Const m_def_isRollUp = False

'Const m_def_BackColorDown = vbButtonShadow

'Property Variables:
Dim m_PositionSave As Boolean
Private m_SplitWidth As Single
'Private m_ResizeLimit As Single
'Private m_BackColor As OLE_COLOR
'Dim m_MouseOver As OLE_COLOR

Private m_SplitLimit As Single
Private m_LimitBorder As Single
Private m_isRollUp As Boolean

Private m_SplitLimitWidth As Single
Private m_obj1() As Object
Private m_obj2() As Object
'Private m_BackColorDown As OLE_COLOR

Const m_def_Style = 1
Public Enum Style
    Vertical = 1
    Horizontal = 0
End Enum
Private m_Style As Style

Const m_def_Binding = 0
Public Enum Binding
    LeftTop = 1
    RightBottom = 0
End Enum
Private m_Binding As Binding

'Public Enum SizeArea
'    [As Control] = 1
'    [As Parent] = 0
'End Enum
'Private m_flagSizeArea As SizeArea

Private m_barclick As Boolean

Private ControlAddress As String
Dim SplitPositionSave As Single
Dim FlagDontSavePosition As Boolean

'Private Bar_ As Object
'Private BarHot_ As Object
'Private BarPres_ As Object

Private ControlS() As Integer

'Event Declarations:
Event MoveEnd() 'MappingInfo=UserControl,UserControl,-1,Resize
Event DblClick()

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_MemberFlags = "40"
   Load aboutfrm
   aboutfrm.Show (vbModal)
End Sub


'***********************************
'Свойства
'***********************************
Public Property Get SplitLimit() As Single
Attribute SplitLimit.VB_ProcData.VB_Invoke_Property = ";SplitV"
    SplitLimit = m_SplitLimit
End Property

Public Property Let SplitLimit(ByVal New_SplitLimit As Single)
    m_SplitLimit = New_SplitLimit
    PropertyChanged "SplitLimit"
End Property

'Public Property Get LimitBorder() As Single
'    LimitBorder = m_LimitBorder
'End Property
'
'Public Property Let LimitBorder(ByVal New_LimitBorder As Single)
'    m_LimitBorder = New_LimitBorder
'    PropertyChanged "LimitBorder"
'End Property

Public Property Get isRollUp() As Boolean
    isRollUp = m_isRollUp
End Property

Public Property Let isRollUp(ByVal is_RollUp As Boolean)
    m_isRollUp = is_RollUp
    PropertyChanged "isRollUp"
    bMoving = True
'    m_barclick = False
    UserControl_MouseMove 0, 0, 0, 0
    bMoving = False
    Resize
End Property

Public Property Get Style() As Style
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As Style)
    m_Style = New_Style
    PropertyChanged "Style"
    If m_Style = Horizontal Then
        Dim w As Long
        w = UserControl.Width
        UserControl.Width = UserControl.Height
        UserControl.Height = w
    Else
        Dim h As Long
        h = UserControl.Height
        UserControl.Height = UserControl.Width
        UserControl.Width = h
    End If
    UserControl_Resize
End Property

Public Property Get Binding() As Binding
    Binding = m_Binding
End Property

Public Property Let Binding(ByVal New_Binding As Binding)
    m_Binding = New_Binding
    PropertyChanged "Binding"
    If IsDebugMode = False Then ResizeControl
End Property

Function IsDebugMode() As Boolean
Debug.Assert zSetTrue(IsDebugMode)
End Function

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

Public Property Set obj1(ByVal New_obj1 As Object)
Attribute obj1.VB_ProcData.VB_Invoke_PropertyPutRef = ";SplitV"
On Error GoTo ERR
    Dim n As Integer
    n = Amount(m_obj1, Push)
    Set m_obj1(n) = New_obj1
Exit Property

ERR:
MsgBox "Not set " & New_obj1.Name & ", " & ERR.Description
On Error Resume Next
End Property

Public Property Set obj2(ByVal New_obj2 As Object)
Attribute obj2.VB_ProcData.VB_Invoke_PropertyPutRef = ";SplitV"
On Error GoTo ERR
    Dim n As Integer
    n = Amount(m_obj2, Push)
    Set m_obj2(n) = New_obj2
Exit Property

ERR:
MsgBox "Not set " & New_obj2.Name & ", " & ERR.Description
On Error Resume Next
End Property

'Public Property Get BackColorDown() As OLE_COLOR
'    BackColorDown = m_BackColorDown
'End Property
'
'Public Property Let BackColorDown(ByVal New_BackColorDown As OLE_COLOR)
'    m_BackColorDown = New_BackColorDown
'    PropertyChanged "BackColorDown"
'End Property

Public Sub ResizeControl()
bMoving = True
'If isRollUp = False Then
m_barclick = True

'If size > 0 Then
'SplitPositionSave = 0
'Else
If PositionSave = True Then
SplitPositionSave = GetSetting(App.ProductName, "Position", ControlAddress, SplitLimit)
End If
'End If

Dim move As Single

    If Style = Vertical Then
    
'        If SplitPositionSave > SplitLimit Then SplitPositionSave = SplitLimit
        
        If Binding = RightBottom And SplitPositionSave > 0 Then
            ' положение справа
            move = Extender.Parent.Width - SplitPositionSave '- UserControl.Extender.Width
        ElseIf Binding = LeftTop And SplitPositionSave > 0 Then
            ' положение слева
            move = Extender.Parent.ScaleLeft + SplitPositionSave '+ UserControl.Extender.Width
        End If
        
        UserControl_MouseMove 0, 0, move, 0
        UserControl_MouseUp 0, 0, move, 0

    Else
        
'        If SplitPositionSave > SplitLimit Then SplitPositionSave = SplitLimit
        
        If Binding = RightBottom Then
            move = Extender.Parent.ScaleHeight - SplitPositionSave '- LimitBorder
        ElseIf Binding = LeftTop Then
            move = Extender.Parent.ScaleHeight - SplitPositionSave '+ UserControl.Extender.Height '+ LimitBorder
        End If
        
        UserControl_MouseMove 0, 0, 0, move
        UserControl_MouseUp 0, 0, 0, move

    End If
    
'If PositionSave = True Then SaveSetting App.ProductName, "Position", ControlAddress, SplitPositionSave

bMoving = False
m_barclick = False
End Sub


Private Sub UserControl_DblClick()
RaiseEvent DblClick
If isRollUp Then
ResizeControl
isRollUp = False
Else
FlagDontSavePosition = False
isRollUp = True
FlagDontSavePosition = True
End If
End Sub

'***********************************
'UserControl
'***********************************
Private Sub UserControl_InitProperties()

    ReDim m_obj1(0): ReDim m_obj2(0)
'    m_BackColor = m_def_BackColor
    m_SplitLimit = m_def_SplitLimit
'    m_BackColorDown = m_def_BackColorDown
    m_Style = m_def_Style
    m_SplitWidth = m_def_SplitWidth
    m_PositionSave = m_def_PositionSave
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.Extender.ZOrder 0
    m_barclick = False
    bMoving = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim sglPos As Single
  
  Dim ParentLeft, ParentWidth, ParentTop, ParentHeight As Single
  
  If Button > 0 Then FlagDontSavePosition = False: m_isRollUp = False
  
  On Error Resume Next
  
  ParentLeft = Extender.Parent.ScaleLeft
  ParentWidth = Extender.Parent.ScaleWidth
  ParentTop = Extender.Parent.ScaleTop
  ParentHeight = Extender.Parent.ScaleHeight
  
  If ParentLeft Is Empty Then ParentLeft = 0
  If ParentTop Is Empty Then ParentTop = 0
  
'  m_isRollUp = False
  
  If bMoving Then

'    SetRect R1, Extender.Parent.ScaleLeft, Extender.Parent.ScaleTop, Extender.Parent.ScaleWidth, Extender.Parent.ScaleHeight
'    SetRect R2, Extender.Parent.ScaleLeft, Extender.Parent.ScaleTop, Extender.Parent.ScaleWidth, Extender.Parent.ScaleHeight
  
  If Style = Vertical Then
  
        sglPos = X + UserControl.Extender.Left - CInt(UserControl.Extender.Width / 2)
        
        If m_barclick = False Then
            
            If SplitLimit > 0 Then
                
                If Binding = RightBottom Then
                    
                    If sglPos < (ParentWidth - SplitLimit) And isRollUp = False Then
'                        ' справа
                        UserControl.Extender.Left = ParentWidth - SplitLimit '- LimitBorder
                    Else
                        GoTo BorderRight
                    End If
                    
                ElseIf Binding = LeftTop Then
                
                    If sglPos > (ParentLeft + SplitLimit) And isRollUp = False Then
'                        ' слева
                        UserControl.Extender.Left = ParentLeft + SplitLimit '+ LimitBorder
                    Else
                        GoTo BorderLeft
                    End If
                    
                End If

            Else
            
            If Binding = RightBottom Then
                GoTo BorderRight
            ElseIf Binding = LeftTop Then
                GoTo BorderLeft
            End If
            
BorderRight:
                If sglPos > (ParentWidth - UserControl.Extender.Width) Or isRollUp Then
                    ' ограничение справа от границы элемента
                    UserControl.Extender.Left = ParentWidth - UserControl.Extender.Width '- LimitBorder
                    Exit Sub
                ElseIf SplitLimit = 0 And sglPos < (ParentLeft - UserControl.Extender.Width) Then
                    UserControl.Extender.Left = 0
                    Exit Sub
                Else
                    UserControl.Extender.Left = sglPos '- LimitBorder
                    Exit Sub
                End If
                
BorderLeft:

                If sglPos < (ParentLeft - UserControl.Extender.Width) Or isRollUp Then
                    ' ограничение слева от границы элемента
                    UserControl.Extender.Left = 0 '+ LimitBorder
                    Exit Sub
                ElseIf SplitLimit = 0 And sglPos > (ParentWidth - UserControl.Extender.Width) Then
                    UserControl.Extender.Left = ParentWidth - UserControl.Extender.Width '- LimitBorder
                    Exit Sub
                Else
                    UserControl.Extender.Left = sglPos '- LimitBorder
                    Exit Sub
                End If

            End If
            
        Else
        
           If Binding = RightBottom Then
           UserControl.Extender.Left = ParentWidth - SplitPositionSave
           Else
           UserControl.Extender.Left = ParentLeft + SplitPositionSave ' UserControl.Extender.Width
           End If
           
        End If
        
  Else
  
        sglPos = Y + UserControl.Extender.Top - CInt(UserControl.Extender.Height / 2)
        
        If m_barclick = False Then

            If SplitLimit > 0 Then
                
                If Binding = LeftTop Then
                    If sglPos > (ParentTop + SplitLimit) And isRollUp = False Then
                        ' сверху
                        UserControl.Extender.Top = ParentTop + SplitLimit '+ LimitBorder '- UserControl.Extender.Height
                    Else
                        GoTo BorderTop
                    End If
                ElseIf Binding = RightBottom Then
                    If sglPos < (ParentHeight - SplitLimit) And isRollUp = False Then
                        ' снизу
                        UserControl.Extender.Top = ParentHeight - SplitLimit '- LimitBorder '- UserControl.Extender.Width
                    Else
                        GoTo BorderBottom
                    End If
                End If
            
            Else
            
            If Binding = LeftTop Then
                GoTo BorderTop
            ElseIf Binding = RightBottom Then
                GoTo BorderBottom
            End If
            
BorderTop:

                If sglPos < UserControl.Extender.Height Or isRollUp Then
                    ' ограничение сверху
                    UserControl.Extender.Top = 0 '+ LimitBorder
                    Exit Sub
                ElseIf SplitLimit = 0 And sglPos > (ParentHeight - UserControl.Extender.Height) Or isRollUp Then
                    UserControl.Extender.Top = ParentHeight - UserControl.Extender.Height '- LimitBorder
                    Exit Sub
                Else
                    UserControl.Extender.Top = sglPos '+ LimitBorder
                    Exit Sub
                End If

BorderBottom:

                If sglPos > (ParentHeight - UserControl.Extender.Height) Or isRollUp Then
                    ' ограничение снизу
                    UserControl.Extender.Top = ParentHeight - UserControl.Extender.Height '- LimitBorder
                    Exit Sub
                ElseIf SplitLimit = 0 And sglPos < UserControl.Extender.Height Or isRollUp Then
                    UserControl.Extender.Top = 0 '+ LimitBorder
                    Exit Sub
                Else
                    UserControl.Extender.Top = sglPos '- LimitBorder
                    Exit Sub
                End If

            End If
            
        Else
        
           If Binding = RightBottom Then
           UserControl.Extender.Top = ParentHeight - SplitPositionSave
           Else
           UserControl.Extender.Top = ParentTop + SplitPositionSave
           End If
           
        End If
  End If
    
  End If
  
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bMoving = False
'm_isRollUp = False
Resize
'UserControl.Extender.ZOrder 1
If PositionSave = True And FlagDontSavePosition = False Then SaveSetting App.ProductName, "Position", ControlAddress, SplitPositionSave
End Sub

Private Sub Resize()

'    bMoving = False
    Dim i As Integer
    Dim sizeobj1 As Integer
    Dim sizeobj2 As Integer
    
   On Error GoTo DONTRUN
    sizeobj1 = UBound(m_obj1): sizeobj2 = UBound(m_obj2)

On Error GoTo ERR

    Dim SplitControlsSize As Integer
    If sizeobj1 > sizeobj2 Then
        SplitControlsSize = sizeobj1
    ElseIf sizeobj2 > sizeobj1 Then
        SplitControlsSize = sizeobj2
    Else
        SplitControlsSize = sizeobj1
    End If
    
    For i = 0 To SplitControlsSize '- 1
    
    If Style = Vertical Then
        If i <= UBound(m_obj2) And UBound(m_obj1) < i Then
            SizeControls UserControl.Extender.Left, m_obj1(0), m_obj2(i)
        ElseIf i <= UBound(m_obj1) And UBound(m_obj2) < i Then
            SizeControls UserControl.Extender.Left, m_obj1(i), m_obj2(0)
        ElseIf i <= UBound(m_obj1) And i <= UBound(m_obj2) Then
            SizeControls UserControl.Extender.Left, m_obj1(i), m_obj2(i)
        End If
        If m_barclick = False Then
            If Binding = LeftTop Then
                SplitPositionSave = Extender.Parent.ScaleLeft + UserControl.Extender.Left
            ElseIf Binding = RightBottom Then
                SplitPositionSave = Extender.Parent.ScaleWidth - UserControl.Extender.Left
            End If
        End If
    Else
        If i <= UBound(m_obj2) And UBound(m_obj1) < i Then
            SizeControls UserControl.Extender.Top, m_obj1(0), m_obj2(i)
        ElseIf i <= UBound(m_obj1) And UBound(m_obj2) < i Then
            SizeControls UserControl.Extender.Top, m_obj1(i), m_obj2(0)
        ElseIf i <= UBound(m_obj1) And i <= UBound(m_obj2) Then
            SizeControls UserControl.Extender.Top, m_obj1(i), m_obj2(i)
        End If
        If m_barclick = False Then
            If Binding = LeftTop Then
                SplitPositionSave = Extender.Parent.ScaleTop + UserControl.Extender.Top
            ElseIf Binding = RightBottom Then
                SplitPositionSave = Extender.Parent.ScaleHeight - UserControl.Extender.Top
            End If
        End If
        
    End If
    Next
    
'    If PositionSave = True Then SaveSetting App.ProductName, "Position", ControlAddress, SplitPositionSave
    
'    UserControl.BackColor() = BackColor
    
    RaiseEvent MoveEnd
Exit Sub
ERR:
'
MsgBox ERR.Description & "Data: " & UBound(m_obj1) & ", " & UBound(m_obj2)
On Error Resume Next
DONTRUN:
End Sub

Private Sub UserControl_Resize()
'On Error Resume Next
On Error GoTo ERR
If Style = Vertical Then
UserControl.Width = SplitWidth

Line2.Y1 = 0: Line2.Y2 = UserControl.Height
Line1.Y1 = 0: Line1.Y2 = UserControl.Height

Line2.X1 = UserControl.Width / 2: Line2.X2 = UserControl.Width / 2
Line1.X1 = UserControl.Width / 2 - 10: Line1.X2 = UserControl.Width / 2 - 10

Else

UserControl.Height = SplitWidth

Line2.X1 = 0: Line2.X2 = UserControl.Width
Line1.X1 = 0: Line1.X2 = UserControl.Width

Line1.Y1 = UserControl.Height / 2: Line1.Y2 = UserControl.Height / 2
Line2.Y1 = UserControl.Height / 2 - 10: Line2.Y2 = UserControl.Height / 2 - 10

End If
Exit Sub
ERR:
MsgBox "Err resize " & ERR.Description, vbCritical
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'On Error Resume Next
'    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_SplitLimit = PropBag.ReadProperty("SplitLimit", m_def_SplitLimit)
    m_LimitBorder = PropBag.ReadProperty("LimitBorder", m_def_LimitBorder)
'    m_flagSizeArea = PropBag.ReadProperty("SizeArea", 0)
'    m_MouseOver = PropBag.ReadProperty("m_MouseOver", m_def_MouseOver)
'    Set m_obj1 = PropBag.ReadProperty("obj1", Nothing)
'    Set m_obj2 = PropBag.ReadProperty("obj2", Nothing)
'    m_BackColorDown = PropBag.ReadProperty("BackColorDown", m_def_BackColorDown)
'    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
'    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Binding = PropBag.ReadProperty("Binding", m_def_Binding)
'    m_ResizeLimit = PropBag.ReadProperty("ResizeLimit", m_def_ResizeLimit)
'    If m_ResizeLimit = 0 Then m_ResizeLimit = m_def_ResizeLimit
    m_SplitWidth = PropBag.ReadProperty("SplitWidth", m_def_SplitWidth)
    m_PositionSave = PropBag.ReadProperty("PositionSave", m_def_PositionSave)
    
    ControlAddress = Extender.Parent.Name & UserControl.Extender.Name
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SplitLimit", m_SplitLimit, m_def_SplitLimit)
    Call PropBag.WriteProperty("LimitBorder", m_LimitBorder, m_def_LimitBorder)
'    Call PropBag.WriteProperty("SizeArea", m_flagSizeArea, 0)
'    Call PropBag.WriteProperty("m_MouseOver", m_MouseOver, m_def_MouseOver)
'    Call PropBag.WriteProperty("obj1", m_obj1, Nothing)
'    Call PropBag.WriteProperty("obj2", m_obj2, Nothing)
'    Call PropBag.WriteProperty("BackColorDown", m_BackColorDown, m_def_BackColorDown)
'    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
'    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Binding", m_Binding, m_def_Binding)
'    Call PropBag.WriteProperty("ResizeLimit", m_ResizeLimit, m_def_ResizeLimit)
    Call PropBag.WriteProperty("SplitWidth", m_SplitWidth, m_def_SplitWidth)
    Call PropBag.WriteProperty("PositionSave", m_PositionSave, m_def_PositionSave)
    
End Sub

Private Sub UserControl_Show()
ini
'ControlAddress = Extender.Parent.Name & UserControl.Extender.Name
'SaveSetting App.ProductName, "Main", "version", App.Major & "," & App.Minor & App.Revision
'UserControl.Extender.ZOrder 0
'Subclass Picture1, 0, style
'ResizeControl
End Sub

'***********************************
'Внутренние процедуры
'***********************************
Private Sub SizeControls(ch As Single, Optional ByRef obj1 As Object, Optional ByRef obj2 As Object)
On Error Resume Next
If Style = Vertical Then

    obj1.Width = UserControl.Extender.Left - obj1.Left ') + UserControl.Extender.Height 'ch - R1.Top
    obj2.Left = UserControl.Extender.Left + UserControl.Extender.Width
    obj2.Width = Extender.Parent.ScaleWidth - obj2.Left
    
Else
    
    obj1.Height = UserControl.Extender.Top - obj1.Top ') + UserControl.Extender.Height 'ch - R1.Top
    obj2.Top = UserControl.Extender.Top + UserControl.Extender.Height
    obj2.Height = Extender.Parent.ScaleHeight - obj2.Top

End If

End Sub

'Public Property Get BorderStyle() As Integer
'    BorderStyle = UserControl.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    UserControl.BorderStyle() = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'Public Property Get BackColor() As OLE_COLOR
'    BackColor = m_BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    If bMoving = False Then m_BackColor = New_BackColor
'    UserControl.BackColor() = m_BackColor
'    PropertyChanged "BackColor"
'End Property


Private Sub ini()
Dim Pic As Object
    If Style = Vertical Then
'        Set Pic = LoadResPicture(101, 0)
'        Picture1.Left = 0
'        Set BarHot_ = LoadResPicture(102, 0)
        Set UserControl.MouseIcon = LoadResPicture(103, 2)
    Else
'        Set Pic = LoadResPicture(103, 0)
'        Picture1.Top = 0
'        Set BarHot_ = LoadResPicture(104, 0)
        Set UserControl.MouseIcon = LoadResPicture(102, 2)
    End If
    UserControl.MousePointer = 99
End Sub

'Public Property Get ResizeLimit() As Single
'    ResizeLimit = m_ResizeLimit
'End Property
'
'Public Property Let ResizeLimit(ByVal New_ResizeLimit As Single)
'    m_ResizeLimit = New_ResizeLimit
'    PropertyChanged "ResizeLimit"
'End Property

Public Property Get SplitWidth() As Single
    SplitWidth = m_SplitWidth
End Property

Public Property Let SplitWidth(ByVal New_SplitWidth As Single)
    m_SplitWidth = New_SplitWidth
    PropertyChanged "SplitWidth"
End Property

Public Property Get PositionSave() As Boolean
    PositionSave = m_PositionSave
End Property

Public Property Let PositionSave(ByVal New_PositionSave As Boolean)
    m_PositionSave = New_PositionSave
    PropertyChanged "PositionSave"
End Property

