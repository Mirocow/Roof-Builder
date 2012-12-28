VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form movepic 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move Picture"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1920
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      Picture         =   "movepic.frx":0000
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Комманда4 
      DisabledPicture =   "movepic.frx":0442
      DownPicture     =   "movepic.frx":0884
      DragIcon        =   "movepic.frx":0CC6
      Height          =   495
      Left            =   720
      Picture         =   "movepic.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move Down"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Комманда3 
      DisabledPicture =   "movepic.frx":154A
      DownPicture     =   "movepic.frx":198C
      DragIcon        =   "movepic.frx":1DCE
      Height          =   495
      Left            =   240
      Picture         =   "movepic.frx":2210
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Move left"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Комманда2 
      DisabledPicture =   "movepic.frx":2652
      DownPicture     =   "movepic.frx":2A94
      DragIcon        =   "movepic.frx":2ED6
      Height          =   495
      Left            =   1200
      Picture         =   "movepic.frx":3318
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Move Right"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Комманда1 
      DisabledPicture =   "movepic.frx":375A
      DownPicture     =   "movepic.frx":3B9C
      DragIcon        =   "movepic.frx":3FDE
      Height          =   495
      Left            =   720
      Picture         =   "movepic.frx":4420
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Move Up"
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "movepic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Lapepic.Command5 = True
End Sub


Private Sub Form_Load()
    Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
    Me.Top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.Top)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
    SaveSetting App.ProductName, "Position", Me.name & "top", Me.Top
    OfficeStart.menu_view_m_Click 9
End Sub


Private Sub Комманда1_Click()
    Lapepic.Picture1.ScaleTop = Lapepic.Picture1.ScaleTop - Lapepic.Picture1.ScaleHeight * 0.1
    'draw_rules
    Lapepic.Draw_Systems Lapepic.Picture1
End Sub


Private Sub Комманда2_Click()
    Lapepic.Picture1.ScaleLeft = Lapepic.Picture1.ScaleLeft + Lapepic.Picture1.ScaleWidth * 0.1
    'draw_rules
    Lapepic.Draw_Systems Lapepic.Picture1
End Sub


Private Sub Комманда3_Click()
    Lapepic.Picture1.ScaleLeft = Lapepic.Picture1.ScaleLeft - Lapepic.Picture1.ScaleWidth * 0.1
    'draw_rules
    Lapepic.Draw_Systems Lapepic.Picture1
End Sub


Private Sub Комманда4_Click()
    Lapepic.Picture1.ScaleTop = Lapepic.Picture1.ScaleTop + Lapepic.Picture1.ScaleHeight * 0.1
    'draw_rules
    Lapepic.Draw_Systems Lapepic.Picture1
End Sub

