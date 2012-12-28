VERSION 5.00
Begin VB.UserControl SuperRuler 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "SuperRuler.ctx":0000
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderStyle     =   3  'Dot
      X1              =   1680
      X2              =   1680
      Y1              =   960
      Y2              =   2040
   End
   Begin VB.Menu mnuScaleModeMenu 
      Caption         =   "ScaleModeMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuScaleMode 
         Caption         =   "Twips"
         Index           =   0
      End
      Begin VB.Menu mnuScaleMode 
         Caption         =   "Pixels"
         Index           =   1
      End
      Begin VB.Menu mnuScaleMode 
         Caption         =   "Milimeters"
         Index           =   2
      End
      Begin VB.Menu mnuScaleMode 
         Caption         =   "Inches"
         Index           =   3
      End
   End
End
Attribute VB_Name = "SuperRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***[Enumerations]***************************************************************************************************
Public Enum enuOrientation
    orHorizontal = 0
    orVertical = 1
End Enum

Public Enum enuDirection
    Forward_Direction = 0
    Reverse_Direction = 1
End Enum

Public Enum enuScaleMode
    smTwips = 0
    smPixels = 1
    smMilimeters = 2
    smInches = 3
    smUser = 4
End Enum

Public Enum enuBorderStyle
    bsNoBorder = 0
    bsSingle = 1
End Enum

Public Enum enuTypeLen
    sm = 0
    mm = 1
    m = 2
End Enum

'***[Default Constants]******************************************************************************************************
Private Const mvar_def_Orientation As Long = orHorizontal
Private Const mvar_def_BorderStyle As Long = bsNoBorder
Private Const mvar_def_ScaleMode As Long = smTwips
Private Const mvar_def_MouseTrackingOn As Boolean = False


'***[Shared Variables]******************************************************************************************************
Private mvarOrientation As Long
Private mvarBorderStyle As Long
Private mvarMouseTrackingOn As Boolean
Private mvarMeasure As Integer
Private mvarScale As Long

Private mvarScaleTop As Single
Private mvarScaleWidth As Single
Private mvarScaleLeft As Single
Private mvarScaleHeight As Single

Private is_on_control As Boolean
Private LineColor As Long
Private MidleLineColor As Long

'***[Events]*********************************************************************************************************
Public Event ScaleModeChanged(Mode As enuScaleMode)
Public Event HooverValue(Value As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, Value As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, Value As Single)
Public Event Resize()
Public Event Click()
Public Event Move(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Default Property Values:
Const m_def_MaxH = 6400
Const m_def_MaxV = 6400
'Property Variables:
Dim m_MaxH As Single
Dim m_MaxV As Single


'***[Properties]*****************************************************************************************************
Public Property Get Orientation() As enuOrientation
    Orientation = mvarOrientation
End Property

Public Property Let Orientation(ByVal Value As enuOrientation)
    mvarOrientation = Value
    RenderControl
    PropertyChanged "Orientation"
End Property

Public Property Get BorderStyle() As enuBorderStyle
    BorderStyle = mvarBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As enuBorderStyle)
    mvarBorderStyle = Value
    UserControl.BorderStyle = mvarBorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    RenderControl
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    RenderControl
    PropertyChanged "BackColor"
End Property

Public Property Get MouseTrackingOn() As Boolean
    MouseTrackingOn = mvarMouseTrackingOn
End Property

Public Property Let MouseTrackingOn(ByVal Value As Boolean)
    mvarMouseTrackingOn = Value
    If Value Then
    Line1.Visible = True
    Else
    Line1.Visible = False
    End If
    PropertyChanged "MouseTrackingOn"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mvarOrientation = PropBag.ReadProperty("Orientation", mvar_def_Orientation)
    BorderStyle = PropBag.ReadProperty("BorderStyle", mvar_def_BorderStyle)
    mvarMouseTrackingOn = PropBag.ReadProperty("MouseTrackingOn", mvar_def_MouseTrackingOn)
    ScaleMode = PropBag.ReadProperty("ScaleMode", mvar_def_ScaleMode)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    RenderControl
    m_MaxH = PropBag.ReadProperty("MaxH", m_def_MaxH)
    m_MaxV = PropBag.ReadProperty("MaxV", m_def_MaxV)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    mvarMeasure = PropBag.ReadProperty("Measure", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Orientation", mvarOrientation, mvar_def_Orientation)
    Call PropBag.WriteProperty("BorderStyle", mvarBorderStyle, mvar_def_BorderStyle)
    Call PropBag.WriteProperty("MouseTrackingOn", mvarMouseTrackingOn, mvar_def_MouseTrackingOn)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("MaxH", m_MaxH, m_def_MaxH)
    Call PropBag.WriteProperty("MaxV", m_MaxV, m_def_MaxV)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Measure", mvarMeasure, 0)
End Sub


Private Sub UserControl_Initialize()
mvarScale = 100
ScaleMode = smUser
LineColor = RGB(136, 136, 136)
MidleLineColor = RGB(172, 172, 172)
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
'        PopupMenu mnuScaleModeMenu
    Else
        RaiseEvent MouseDown(Button, Shift, CalculateValue(X, Y))
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RenderTrackLine X, Y, Button, Shift
    RaiseEvent HooverValue(CalculateValue(X, Y))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, CalculateValue(X, Y))
End Sub

Public Function CalculateValue(X As Single, Y As Single) As Single
    Dim myValue As Single

    Select Case mvarOrientation
    Case orHorizontal
        Dim X1 As Single
        Dim Ratio1 As Single
        Ratio1 = 10
        X1 = Int(X / Ratio1) * Ratio1
        If X > X1 - Ratio1 / 2 And X < X1 + Ratio1 / 2 Then
            myValue = X1
        Else
            myValue = X1 + Ratio1
        End If
    Case orVertical
        Dim Ratio2 As Single
        Dim Y1 As Single
        Ratio2 = 10
        Y1 = Int(Y / Ratio2) * Ratio2
        If Y > Y1 - Ratio2 / 2 And Y < Y1 + Ratio2 / 2 Then
            myValue = Y1
        Else
            myValue = Y1 + Ratio2
        End If
    End Select
    
    ' В зависимости от еденицы измерения преобразуем данные
    ' 1 - 1м (м)
    ' 100 - 1м (см)
    ' 1000 - 1м (мм)
    
'    Select Case mvarMeasure
'    Case 0
'        myValue = myValue * 1
'    Case 1
'        myValue = myValue * 10
'    Case 2
'        myValue = myValue * 0.01
'    End Select

CalculateValue = myValue
        
End Function

Public Sub RenderTrackLine(X As Single, Y As Single, Optional Button As Integer, Optional Shift As Integer)
    If mvarMouseTrackingOn = True Then
    
        'Optionaly render Mouse tracking line
        Select Case Orientation
        Case orHorizontal
            Line1.X1 = X
            Line1.Y1 = 0
            Line1.X2 = X
            Line1.Y2 = Height
        Case orVertical
            Line1.X1 = 0
            Line1.Y1 = Y
            Line1.X2 = Width
            Line1.Y2 = Y
        End Select
        
        RaiseEvent Move(Button, Shift, X, Y)
        
    End If
End Sub

Public Function GetCurrentPos() As Single
    Select Case Orientation
    Case orHorizontal
        GetCurrentPos = Line1.X1
    Case orVertical
        GetCurrentPos = Line1.Y1
    End Select
End Function

Private Sub mnuScaleMode_Click(Index As Integer)
    ScaleMode = Index
    RenderControl
End Sub

Public Sub Refresh()
    RenderControl
End Sub

Private Sub RenderControl()
    Dim mySmallScale As Long
    Dim myValue As Single
    
    On Error Resume Next
       
    If Ambient.UserMode = False Then Exit Sub

    Cls
    
    If m_MaxH = 0 And m_MaxV = 0 Then Exit Sub
    
    ' В зависимости от еденицы измерения преобразуем данные
    ' 1 - 1м (м)
    ' 100 - 1м (см)
    ' 1000 - 1м (мм)
    
    Dim PmvarScale As Single
    Select Case mvarMeasure
    Case 0
        PmvarScale = 1
    Case 1
        PmvarScale = 10
    Case 2
        PmvarScale = 0.01
    End Select
    
    Dim ScaleTop As Single
    Dim ScaleWidth As Single
    Dim ScaleLeft As Single
    Dim ScaleHeight As Single
    
    mySmallScale = mvarScale / 10
    Select Case mvarOrientation
    Case orHorizontal
        If mvarScaleWidth = 0 Then Exit Sub
        UserControl.ScaleWidth = mvarScaleWidth
        UserControl.ScaleLeft = mvarScaleLeft
    Case orVertical
        If mvarScaleHeight = 0 Then Exit Sub
        UserControl.ScaleTop = mvarScaleTop
        UserControl.ScaleHeight = mvarScaleHeight
    End Select
    
    Dim From As Long
    Dim Till As Long
    Dim i As Long
    Dim j As Long
    
    Select Case mvarOrientation
    Case orHorizontal

        For j = 0 To m_MaxH + mvarScale Step mvarScale
            'Draw big line
            Line (j, 0)-(j, UserControl.ScaleHeight), LineColor
            myValue = j / mvarScale
            UserControl.CurrentY = 0
            Print CLng(myValue * mvarScale * PmvarScale)
            'Draw small lines
            For i = j + mySmallScale To j + mvarScale - mySmallScale Step mySmallScale
                If i = j + mvarScale / 2 Then
                    Line (i, UserControl.ScaleHeight / 2)-(i, UserControl.ScaleHeight), MidleLineColor
                Else
                    Line (i, UserControl.ScaleHeight - UserControl.ScaleHeight / 3)-(i, UserControl.ScaleHeight), LineColor
                End If
            Next i
        Next j

        For j = 0 To -m_MaxH - mvarScale Step -mvarScale
            'Draw big line
            Line (j, 0)-(j, UserControl.ScaleHeight), LineColor
            myValue = j / mvarScale
            UserControl.CurrentY = 0
            Print CLng(myValue * mvarScale * PmvarScale)
            'Draw small lines
            For i = j + mySmallScale To j + mvarScale - mySmallScale Step mySmallScale
                If i = j + mvarScale / 2 Then
                    Line (i, UserControl.ScaleHeight / 2)-(i, UserControl.ScaleHeight), MidleLineColor
                Else
                    Line (i, UserControl.ScaleHeight - UserControl.ScaleHeight / 3)-(i, UserControl.ScaleHeight), LineColor
                End If
            Next i
        Next j
        
    Case orVertical
        
        For j = 0 To m_MaxV + mvarScale Step mvarScale
            'Draw big line
            Line (0, j)-(UserControl.ScaleWidth, j), LineColor
            myValue = j / mvarScale
            UserControl.CurrentX = 0
            Print CLng(myValue * mvarScale * PmvarScale)
            'Draw small lines
            For i = j + mySmallScale To j + mvarScale - mySmallScale Step mySmallScale
                If i = j + mvarScale / 2 Then
                    Line (UserControl.ScaleWidth / 2, i)-(UserControl.ScaleWidth, i), MidleLineColor
                Else
                    Line (UserControl.ScaleWidth - UserControl.ScaleWidth / 3, i)-(UserControl.ScaleWidth, i), LineColor
                End If
            Next i
        Next j
        
        For j = 0 To -m_MaxV - mvarScale Step -mvarScale
            'Draw big line
            Line (0, j)-(UserControl.ScaleWidth, j), LineColor
            myValue = j / mvarScale
            UserControl.CurrentX = 0
            Print CLng(myValue * mvarScale * PmvarScale)
            'Draw small lines
            For i = j + mySmallScale To j + mvarScale - mySmallScale Step mySmallScale
                If i = j + mvarScale / 2 Then
                    Line (UserControl.ScaleWidth / 2, i)-(UserControl.ScaleWidth, i), MidleLineColor
                Else
                    Line (UserControl.ScaleWidth - UserControl.ScaleWidth / 3, i)-(UserControl.ScaleWidth, i), LineColor
                End If
            Next i
        Next j
        
    End Select
    
    
End Sub

' ---------=--------------------=----------------------------=-----------------------
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = mvarScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    mvarScaleWidth = New_ScaleWidth
End Property

' ---------=--------------------=----------------------------=-----------------------
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = mvarScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    mvarScaleTop = New_ScaleTop
End Property

' ---------=--------------------=----------------------------=-----------------------
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = mvarScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    mvarScaleLeft = New_ScaleLeft
End Property

' ---------=--------------------=----------------------------=-----------------------
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = mvarScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    mvarScaleHeight = New_ScaleHeight
End Property

' ---------=--------------------=----------------------------=-----------------------
Public Property Get MaxH() As Single
    If m_MaxH = 0 Then Exit Property
    MaxH = Round_to_big(m_MaxH / mvarScale)
    MaxH = MaxH * mvarScale
End Property

Public Property Let MaxH(ByVal New_MaxH As Single)
    m_MaxH = New_MaxH
    PropertyChanged "MaxH"
End Property

' ---------=--------------------=----------------------------=-----------------------
Public Property Get MaxV() As Single
    If MaxV = 0 Then Exit Property
    MaxV = Round_to_big(m_MaxV / mvarScale)
    MaxV = MaxV * mvarScale
End Property

Public Property Let MaxV(ByVal New_MaxV As Single)
    m_MaxV = New_MaxV
    PropertyChanged "MaxV"
End Property

' ---------=--------------------=----------------------------=-----------------------
Private Function Round_to_big(Number)
  Round_to_big = Number
  If Number > Int(Number) Then Round_to_big = Abs(Int(Number)) + 1
End Function

' ---------=--------------------=----------------------------=-----------------------
Public Property Let UserScale(ByVal New_UserScale As Long)
    mvarScale = New_UserScale
End Property

Public Property Get UserScale() As Long
    UserScale = mvarScale
End Property

' ---------=--------------------=----------------------------=-----------------------
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

' ---------=--------------------=----------------------------=-----------------------
Public Property Get Measure() As enuTypeLen
Attribute Measure.VB_Description = "Sets a custom mouse icon."
    Measure = mvarMeasure
End Property

Public Property Let Measure(ByVal New_mvarMeasure As enuTypeLen)
    mvarMeasure = New_mvarMeasure
    PropertyChanged "Measure"
End Property

' ---------=--------------------=----------------------------=-----------------------
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property


