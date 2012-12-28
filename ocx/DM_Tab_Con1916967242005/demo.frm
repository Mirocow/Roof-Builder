VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "DM TabControl"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Set disabled tab 1"
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   4140
      Width           =   1935
   End
   Begin VB.CheckBox chkhand 
      Caption         =   "Show Hotracking Hand Cursor"
      Height          =   255
      Left            =   210
      TabIndex        =   19
      Top             =   3120
      Width           =   3555
   End
   Begin VB.CheckBox chkrect 
      Caption         =   "Show focus rect"
      Height          =   255
      Left            =   210
      TabIndex        =   18
      Top             =   3720
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Tab 2"
      Height          =   375
      Left            =   195
      TabIndex        =   17
      Top             =   4140
      Width           =   1710
   End
   Begin VB.CheckBox chkunderline 
      Caption         =   "Show Hot Tracking with underline"
      Height          =   255
      Left            =   210
      TabIndex        =   16
      Top             =   2820
      Width           =   3555
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Show Selected tab captions in bold"
      Height          =   255
      Left            =   210
      TabIndex        =   15
      Top             =   3435
      Width           =   3555
   End
   Begin VB.CheckBox chkHot 
      Caption         =   "Hot Tracking"
      Height          =   255
      Left            =   210
      TabIndex        =   14
      Top             =   2535
      Width           =   3555
   End
   Begin Project1.sTabFx sTabFx1 
      Height          =   2130
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   4740
      _extentx        =   8361
      _extenty        =   3757
      trackingcolor   =   255
      boldselection   =   0   'False
      font            =   "demo.frx":0000
      forecolor       =   0
      showrect        =   0   'False
      mouseicon       =   "demo.frx":0028
      showtrackinghand=   0   'False
      Begin VB.PictureBox PicTab 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   3
         Left            =   4545
         ScaleHeight     =   1335
         ScaleWidth      =   2895
         TabIndex        =   10
         Top             =   1455
         Visible         =   0   'False
         Width           =   2895
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Version 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   510
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DM TabControl Replacement"
            Height          =   195
            Left            =   150
            TabIndex        =   12
            Top             =   225
            Width           =   2070
         End
      End
      Begin VB.PictureBox PicTab 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   2
         Left            =   4545
         ScaleHeight     =   1335
         ScaleWidth      =   2895
         TabIndex        =   6
         Top             =   1455
         Visible         =   0   'False
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Mouse"
            Height          =   240
            Index           =   8
            Left            =   270
            TabIndex        =   9
            Top             =   210
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Keyboard"
            Height          =   240
            Index           =   7
            Left            =   270
            TabIndex        =   8
            Top             =   510
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Monitor"
            Height          =   240
            Index           =   6
            Left            =   270
            TabIndex        =   7
            Top             =   795
            Width           =   1425
         End
      End
      Begin VB.PictureBox PicTab 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   1
         Left            =   4545
         ScaleHeight     =   1335
         ScaleWidth      =   2895
         TabIndex        =   5
         Top             =   1455
         Visible         =   0   'False
         Width           =   2895
         Begin VB.DirListBox Dir1 
            Height          =   765
            Left            =   105
            TabIndex        =   11
            Top             =   195
            Width           =   2445
         End
      End
      Begin VB.PictureBox PicTab 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   0
         Left            =   4545
         ScaleHeight     =   1335
         ScaleWidth      =   2895
         TabIndex        =   1
         Top             =   1455
         Visible         =   0   'False
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Quake IV"
            Height          =   240
            Index           =   2
            Left            =   270
            TabIndex        =   4
            Top             =   795
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "DOOM III"
            Height          =   240
            Index           =   1
            Left            =   270
            TabIndex        =   3
            Top             =   510
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Duke 3D"
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   2
            Top             =   210
            Width           =   1425
         End
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TabIdx As Integer
Dim b As Boolean

Sub ArrangeTabs(index As Integer)
Dim x As Integer

    For x = 0 To PicTab.Count - 1
        PicTab(x).Visible = False
    Next
    
    x = 0
    
    PicTab(index).Visible = True
    PicTab(index).Left = 135
    PicTab(index).Top = 450
    
End Sub

Private Sub chkhand_Click()
    sTabFx1.ShowTrackingHand = chkhand
End Sub

Private Sub chkrect_Click()
    sTabFx1.ShowRect = chkrect
End Sub

Private Sub cmdabout_Click()
    MsgBox "Tab Replacement control by DreamVb.", vbInformation
    
End Sub

Private Sub cmdadd_Click()
    sTabFx2.AddTab "Tab " & sTabFx2.TabCount
    
End Sub

Private Sub cmdCap_Click()
   sTabFx1.TabCaption(TabIdx) = txtCaption.Text
End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub Command1_Click()
    sTabFx1.SelectTab 2
    MsgBox sTabFx1.SelectTab
End Sub

Private Sub chkBold_Click()
    sTabFx1.BoldSelection = chkBold
End Sub

Private Sub chkHot_Click()
    sTabFx1.HotTracking = chkHot
End Sub

Private Sub chkunderline_Click()
    sTabFx1.TrackUnderLine = chkunderline
End Sub

Private Sub Command2_Click()
    If sTabFx1.TabDisabled(1) = True Then
        sTabFx1.TabDisabled(1) = False
    Else
        sTabFx1.TabDisabled(1) = True
    End If
End Sub

Private Sub Form_Load()
    DoEvents
    'sTabFx1.TabCaption(0) = "Games"
    sTabFx1.AddTab "Games ", "A"
    sTabFx1.AddTab "Test Tab ", "B"
    sTabFx1.AddTab "Hardware", "C"
    sTabFx1.AddTab "About", "D"
    sTabFx1.SelectTab 1 'Select tab 1
    
'    sTabFx3.AddTab "Test"
'    sTabFx3.AddTab "Test"
'    sTabFx3.AddTab "Test"
'    sTabFx3.AddTab "Test"
'    sTabFx3.AddTab "Test"
'    sTabFx3.AddTab "Test"
'    sTabFx3.AddTab "Test"
'    sTabFx3.AddTab "Test"
            
End Sub

Private Sub lblIndex_Click()

End Sub

Private Sub sTabFx1_Click(index As Integer, Key As String, Caption As String)
    ArrangeTabs index
    TabIdx = index
    
'    lblkey.Caption = "Tab Key: " & Key
'    lblcap.Caption = "Tab Caption: " & Caption
'    lblIndex.Caption = "Tab Index: " & index
    
    'If index = 2 Then sTabFx1.HightLight(index) = True
End Sub

Private Sub sTabFx1_TabMouseMove(index As Integer, Selected As Boolean, Key As String, Caption As String)
'    lblkey.Caption = "Tab Key: " & Key
'    lblcap.Caption = "Tab Caption: " & Caption
'    lblIndex.Caption = "Tab Index: " & index
End Sub
