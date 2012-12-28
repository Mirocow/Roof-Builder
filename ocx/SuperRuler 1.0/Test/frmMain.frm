VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11805
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11805
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   9360
      Width           =   9615
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0C0C0&
            Height          =   495
            Left            =   720
            Picture         =   "frmMain.frx":0000
            TabIndex        =   17
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton Комманда1 
            DisabledPicture =   "frmMain.frx":0442
            DownPicture     =   "frmMain.frx":0884
            DragIcon        =   "frmMain.frx":0CC6
            Height          =   495
            Left            =   720
            Picture         =   "frmMain.frx":1108
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Move Up"
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton Комманда2 
            DisabledPicture =   "frmMain.frx":154A
            DownPicture     =   "frmMain.frx":198C
            DragIcon        =   "frmMain.frx":1DCE
            Height          =   495
            Left            =   1200
            Picture         =   "frmMain.frx":2210
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Move Right"
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton Комманда3 
            DisabledPicture =   "frmMain.frx":2652
            DownPicture     =   "frmMain.frx":2A94
            DragIcon        =   "frmMain.frx":2ED6
            Height          =   495
            Left            =   240
            Picture         =   "frmMain.frx":3318
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Move left"
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton Комманда4 
            DisabledPicture =   "frmMain.frx":375A
            DownPicture     =   "frmMain.frx":3B9C
            DragIcon        =   "frmMain.frx":3FDE
            Height          =   495
            Left            =   720
            Picture         =   "frmMain.frx":4420
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Move Down"
            Top             =   1320
            Width           =   495
         End
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   100
         Left            =   2040
         Max             =   6400
         Min             =   1
         SmallChange     =   100
         TabIndex        =   11
         Top             =   480
         Value           =   10
         Width           =   7455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   1335
         Left            =   5520
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "TypeLen"
         Height          =   1455
         Left            =   7080
         TabIndex        =   6
         Top             =   840
         Width           =   2295
         Begin VB.OptionButton TypeLen 
            Caption         =   "sm"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton TypeLen 
            Caption         =   "mm"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton TypeLen 
            Caption         =   "m"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   2040
         TabIndex        =   3
         Top             =   840
         Width           =   3375
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   255
            Left            =   1920
            TabIndex        =   21
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   255
            Left            =   1920
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   7680
         TabIndex        =   18
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      Height          =   8775
      Left            =   840
      ScaleHeight     =   8775
      ScaleWidth      =   8775
      TabIndex        =   0
      Top             =   600
      Width           =   8775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Measure As Integer

Const RatioH = 1 ' 1.51
Const RatioV = 1.2 ' 1.2


Private Sub Command1_Click()
SuperRuler1.Refresh
SuperRuler2.Refresh
End Sub

Private Sub Command2_Click()
Picture1.Cls

'Picture1.ForeColor = vbRed
Picture1.DrawWidth = 2

Picture1.Line (0, 0)-(0, 1200), vbRed '
Picture1.Line (0, 1200)-(1200, 1200), vbRed '
Picture1.Line (1200, 1200)-(1200, 0), vbRed '
Picture1.Line (1200, 0)-(0, 0), vbRed '

Picture1.Line (-100, -100)-(-100, 1500), vbGreen '
Picture1.Line (-100, 1500)-(1500, 1500), vbGreen '
Picture1.Line (1500, 1500)-(1500, -100), vbGreen '
Picture1.Line (1500, -100)-(-100, -100), vbGreen '

End Sub

Private Sub Command5_Click()
'#     Переинициализация окна
Call init
SetCenter Picture1, HScroll2, RatioH, RatioV
Command2_Click
RollesRefresh
End Sub

Public Sub SetCenter(Obj, HScroll, prcwidth As Single, prcheight As Single)
    Dim XMax As Single
    Dim XMin As Single
    Dim YMax As Single
    Dim YMin As Single

    Dim FigureLenght
    Dim FigureHeight

    Dim l00A6 As Single
    Dim l00AE As Single
 
    On Error Resume Next
 
    If Obj.ScaleWidth <= 100 Or Abs(Obj.ScaleHeight) <= 100 Then Exit Sub

    Dim P
  
      XMax = 1200
      XMin = 0
      YMax = 1200
      YMin = 0

      FigureLenght = XMax - XMin
      FigureHeight = YMax - YMin

      If FigureHeight <= 0 Then FigureHeight = 1
      If FigureLenght <= 0 Then FigureLenght = 1

      If FigureLenght / FigureHeight > prcwidth Then
  
          l00A6 = prcheight * FigureLenght
  
          If l00A6 > HScroll.Max Then
              HScroll.Value = HScroll.Max
              GoTo L96EE
          End If
  
          Obj.ScaleLeft = XMin - 0.1 * FigureLenght
          Obj.ScaleWidth = l00A6
          l00AE = l00A6 / prcwidth
          Obj.ScaleHeight = -l00AE
          Obj.ScaleTop = (YMax - (FigureHeight / 2)) + (l00AE / 2)
  
      Else
  
          l00AE = -prcheight * FigureHeight
          l00A6 = l00AE * -prcwidth
  
          If l00A6 > HScroll.Max Then
              HScroll.Value = HScroll.Max
              GoTo L96EE
          End If

          Obj.ScaleTop = YMax + 0.1 * FigureHeight
          Obj.ScaleHeight = l00AE
          Obj.ScaleWidth = l00A6
          Obj.ScaleLeft = (XMax - (FigureLenght / 2)) - (l00A6 / 2)
  
      End If

      If HScroll.Max < l00A6 Then HScroll.Max = 64

      If HScroll.Min > l00A6 Then
          HScroll.Value = HScroll.Min
      Else
  
          ' Не центрировать если уже все выровнено
          If HScroll.Value <> Round(l00A6) Then
              HScroll.Value = l00A6
          End If
  
      End If

L96EE:
End Sub

Private Sub Form_Load()
'#      Инициализация окна
Call init

'SuperRuler1.ScaleMode = smUser
SuperRuler1.MouseTrackingOn = True
'SuperRuler2.ScaleMode = smUser
SuperRuler2.MouseTrackingOn = True

Command5.Value = True
End Sub

Sub init()

Picture1.ScaleLeft = 0
Picture1.ScaleWidth = 6400
Picture1.ScaleTop = 600
Picture1.ScaleHeight = -600


SuperRuler1.MaxH = 6400
SuperRuler2.MaxV = 6400
End Sub

Private Sub HScroll1_Change()
    SuperRuler1.Refresh
End Sub

Private Sub HScroll1_Scroll()
    SuperRuler1.Refresh
End Sub

Private Sub HScroll2_Change()

Change_scrol Me.Picture1, Me.HScroll2

Label1 = HScroll2.Value

Command2_Click

Label4 = Round_to_big(HScroll2.Value / (Screen.TwipsPerPixelX * 100)) * 100

SuperRuler1.UserScale = Label4
SuperRuler2.UserScale = Label4

RollesRefresh
End Sub

Function Round_to_big(Number)
  Round_to_big = Number
  If Number > Int(Number) Then Round_to_big = Abs(Int(Number)) + 1
End Function

Sub RollesRefresh()
SuperRuler1.Width = Picture1.Width
SuperRuler1.ScaleLeft = Picture1.ScaleLeft
SuperRuler1.ScaleWidth = Picture1.ScaleWidth
SuperRuler2.Height = Picture1.Height
SuperRuler2.ScaleTop = Picture1.ScaleTop
SuperRuler2.ScaleHeight = Picture1.ScaleHeight

SuperRuler1.Refresh
SuperRuler2.Refresh
End Sub

Public Sub Change_scrol(Pic As Object, scroll As Object)
On Error Resume Next

    Dim l00C2 As Single
    Dim l00C4 As Single
    Dim l00C6 As Single
    Dim l00C8 As Single
    Dim l00CA As Single

    l00C2 = Pic.ScaleLeft + Pic.ScaleWidth
    l00C4 = (Pic.ScaleLeft + l00C2) / 2
    l00C6 = Pic.ScaleTop + (Pic.ScaleHeight / 2)
    l00C8 = scroll.Value
    l00CA = l00C8 / RatioH ' 1.51

    Pic.ScaleTop = l00C6 + (l00CA / 2) '+ 100
    Pic.ScaleHeight = -l00CA '+ 100
    Pic.ScaleWidth = l00C8 '+ 100
    Pic.ScaleLeft = l00C4 - (l00C8 / 2) '+ 100
End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub


Private Sub Option1_Click()

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SuperRuler1.RenderTrackLine X, 0
    SuperRuler2.RenderTrackLine 0, Y

    Command1.Caption = "X=" & Int(SuperRuler1.GetCurrentPos) & " / Y=" & Int(SuperRuler2.GetCurrentPos)

    Label2 = SuperRuler1.CalculateValue(X, Y)
    Label3 = SuperRuler2.CalculateValue(X, Y)
End Sub


Private Sub SuperRuler1_HooverValue(Value As Long)
Label5 = Value
End Sub

Private Sub SuperRuler2_HooverValue(Value As Long)
Label6 = Value
End Sub

'Private Sub SuperRuler1_MouseDown(Button As Integer, Shift As Integer, value As Single)
'MsgBox value
'End Sub
'
'Private Sub SuperRuler2_MouseDown(Button As Integer, Shift As Integer, value As Single)
'MsgBox value
'End Sub

Private Sub TypeLen_Click(Index As Integer)
Measure = Index
Select Case Index
Case 0
    SuperRuler1.Measure = 0
    SuperRuler2.Measure = 0
Case 1
    SuperRuler1.Measure = 1
    SuperRuler2.Measure = 1
Case 2
    SuperRuler1.Measure = 2
    SuperRuler2.Measure = 2
End Select
Command5_Click
End Sub

Private Sub Комманда1_Click()
Picture1.ScaleTop = Picture1.ScaleTop + Picture1.ScaleHeight * 0.1
RollesRefresh
Command2_Click
End Sub

Private Sub Комманда2_Click()
Picture1.ScaleLeft = Picture1.ScaleLeft - Picture1.ScaleWidth * 0.1
RollesRefresh
Command2_Click
End Sub

Private Sub Комманда3_Click()
Picture1.ScaleLeft = Picture1.ScaleLeft + Picture1.ScaleWidth * 0.1
RollesRefresh
Command2_Click
End Sub

Private Sub Комманда4_Click()
Picture1.ScaleTop = Picture1.ScaleTop - Picture1.ScaleHeight * 0.1
RollesRefresh
Command2_Click
End Sub

'
' SuperRuler1
'
Private Sub SuperRuler1_MouseDown(Button As Integer, Shift As Integer, Value As Single)
SuperRuler1_Move 1, 0, SuperRuler1.GetCurrentPos, 0
SuperRuler1.MousePointer = 9
End Sub

Private Sub SuperRuler1_MouseUp(Button As Integer, Shift As Integer, Value As Single)
SuperRuler1.MousePointer = 0
Command2_Click
End Sub

Private Sub SuperRuler1_Move(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And SuperRuler1.MousePointer = 9 Then
Picture1.ScaleLeft = Picture1.ScaleLeft - X
SuperRuler1.ScaleLeft = Picture1.ScaleLeft
SuperRuler1.Refresh
End If
End Sub

'
' SuperRuler2
'
Private Sub SuperRuler2_MouseDown(Button As Integer, Shift As Integer, Value As Single)
SuperRuler2_Move 1, 0, 0, SuperRuler2.GetCurrentPos
SuperRuler2.MousePointer = 7
End Sub

Private Sub SuperRuler2_MouseUp(Button As Integer, Shift As Integer, Value As Single)
SuperRuler2.MousePointer = 0
Command2_Click
End Sub

Private Sub SuperRuler2_Move(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And SuperRuler2.MousePointer = 7 Then
Picture1.ScaleTop = Picture1.ScaleTop - Y
SuperRuler2.ScaleTop = Picture1.ScaleTop
SuperRuler2.Refresh
End If
End Sub

Public Function ConvertData(ByVal Value, Optional divide As Boolean = True) As Single
If divide = True Then
    Select Case Measure 'Setup.Combo4.ListIndex
    Case 0 ' - см
        Value = Value / 100
    Case 1 ' - мм
        Value = Value / 1000
    Case 2 ' - метры
        Value = Value
    End Select
Else
    Select Case Measure 'Setup.Combo4.ListIndex
    Case 0 ' - см
        Value = Value * 100
    Case 1 ' - мм
        Value = Value * 1000
    Case 2 ' - метры
        Value = Value
    End Select
End If
ConvertData = Value
End Function
