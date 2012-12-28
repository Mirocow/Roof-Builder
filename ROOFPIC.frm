VERSION 5.00
Object = "{D4055E62-5507-43CA-B528-924FB94C4FF4}#1.0#0"; "SplitterModern.ocx"
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form ROOFPIC 
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10140
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Форма1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10140
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin SplitterHV.SplitHV SplitHV1 
      Height          =   75
      Left            =   0
      TabIndex        =   8
      Top             =   5160
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   132
      SplitLimit      =   2195
      Style           =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DCFBFC&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5145
      ScaleWidth      =   10125
      TabIndex        =   7
      Top             =   0
      Width           =   10155
      Begin VB.Line Line1 
         BorderStyle     =   2  'Dash
         Visible         =   0   'False
         X1              =   3240
         X2              =   6240
         Y1              =   3360
         Y2              =   2760
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   9615
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   2160
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   480
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin roof.sTabFx sTabFx1 
         Height          =   1215
         Left            =   4440
         TabIndex        =   9
         Top             =   60
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2143
         BoldSelection   =   0   'False
         Border3DStyle   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         ShowRect        =   0   'False
         ShowToolTip     =   0   'False
         ShowTrackingHand=   0   'False
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   4215
            Begin СhameleonButton.chameleonButton Command2 
               Height          =   375
               Left            =   0
               TabIndex        =   19
               Top             =   120
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               BTYPE           =   7
               TX              =   "chameleonButton1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "ROOFPIC.frx":0000
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
               ICONS           =   16
            End
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   2860
         TabIndex        =   12
         Top             =   0
         Width           =   1455
         Begin СhameleonButton.chameleonButton Command5 
            Height          =   405
            Left            =   645
            TabIndex        =   14
            Top             =   525
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "ROOFPIC.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Комманда4 
            Height          =   405
            Left            =   645
            TabIndex        =   15
            Top             =   915
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "ROOFPIC.frx":0038
            PICN            =   "ROOFPIC.frx":0054
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Комманда3 
            Height          =   405
            Left            =   240
            TabIndex        =   16
            Top             =   525
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "ROOFPIC.frx":04A6
            PICN            =   "ROOFPIC.frx":04C2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Комманда1 
            Height          =   405
            Left            =   645
            TabIndex        =   17
            Top             =   120
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "ROOFPIC.frx":0914
            PICN            =   "ROOFPIC.frx":0930
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Комманда2 
            Height          =   405
            Left            =   1050
            TabIndex        =   18
            Top             =   525
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "ROOFPIC.frx":0D82
            PICN            =   "ROOFPIC.frx":0D9E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   5000
         Left            =   120
         Max             =   10000
         Min             =   1
         SmallChange     =   500
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Value           =   1000
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00DCFBFC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         MouseIcon       =   "ROOFPIC.frx":11F0
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label7 
         Caption         =   "X ="
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Метка2 
         Caption         =   "Y ="
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "ROOFPIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Const colorstr = &HFF8080
Dim PSx
Dim PSy

'**************************
'*** УПРАВЛЯЮЩИЕ
'**************************
Private FindPoint As Integer ' Найденная точка
Private Current_P As Integer ' Выбранная точка
Private Current_L As Integer ' Выбранная линия

Private SinA As Single

Sub Command4_Click()
    OfficeStart.TabStrip1.Tabs(1).Selected = True
End Sub


Private Sub Command2_Click()

ScaleLeft_Main = 0
ScaleWidth_Main = 1200
ScaleTop_Main = 625
ScaleHeight_Main = -625

ROOFPIC.Picture1.ScaleLeft = ScaleLeft_Main
ROOFPIC.Picture1.ScaleTop = ScaleTop_Main
ROOFPIC.Picture1.ScaleWidth = ScaleWidth_Main
ROOFPIC.Picture1.ScaleHeight = ScaleHeight_Main

OfficeStart.Clear_project

SetChange True

Draw_Systems Me.Picture1
End Sub

Sub Command5_Click()
If Me.Visible = True Then
SetPicCentr Picture1
Draw_Systems Me.Picture1
End If
End Sub


Public Sub SetPicCentr(P As PictureBox)

    Dim l0032
    Dim l0034
    Dim l0036
    Dim l0038
    Dim l003A
    Dim l003C
    Dim l004A As Single
    Dim l004C As Single
    Dim p00ae As Single
    Dim p00b0 As Single
    Dim p00b2 As Single
    Dim p00b4 As Single
    Dim p00b6 As Single
    Dim p00b8 As Single
  
    On Error Resume Next
  
    If MainCountOfLines < 1 Then GoTo ERR
   
    Dim l00BA
    p00ae = 0: p00b0 = 99999: p00b2 = 0: p00b4 = 99999
    For l00BA = 1 To MainCountOfPoints Step 1
        If Main_Points_X(l00BA) < p00b0 Then p00b0 = Main_Points_X(l00BA)
        If Main_Points_X(l00BA) > p00ae Then p00ae = Main_Points_X(l00BA)
        If Main_Points_Y(l00BA) < p00b4 Then p00b4 = Main_Points_Y(l00BA)
        If Main_Points_Y(l00BA) > p00b2 Then p00b2 = Main_Points_Y(l00BA)
    Next l00BA

    p00b6 = p00ae - p00b0
    p00b8 = p00b2 - p00b4
    l0032 = p00ae
    l0034 = p00b0
    l0036 = p00b2
    l0038 = p00b4
    l003A = p00b6
    l003C = p00b8

    If l003C = 0 Then l003C = 1
'    If RatioW * l003A >= HScroll1.MAX Then GoTo ERR
    If l003A / l003C > RatioH Then
        P.ScaleLeft = l0034 - 0.1 * l003A
        P.ScaleWidth = 1.2 * l003A
        l004A = P.ScaleWidth
        l004C = P.ScaleWidth / RatioH
        P.ScaleHeight = -l004C
        P.ScaleTop = (l0036 - (l003C / 2)) + (l004C / 2)
    Else
        P.ScaleTop = l0036 + 0.1 * l003C
        P.ScaleHeight = -1.2 * l003C
        l004A = P.ScaleHeight * -RatioH
        P.ScaleWidth = l004A
        P.ScaleLeft = (l0032 - (l003A / 2)) - (l004A / 2)
    End If
    
    If l004A > HScroll1.MAX Then HScroll1.MAX = l004A 'GoTo L706C
    HScroll1.value = l004A

ERR:
End Sub


Private Sub Command6_Click()
    OfficeStart.TabStrip1.Tabs(3).Selected = True
End Sub


Sub Form_Load()
    On Error GoTo ERR

    SetFont Me

    ROOFPIC.Label2.Caption = lng.GetResIDstring(1015)
    
    sTabFx1.AddTab lng.GetResIDstring(1016)
    sTabFx1.AddTab lng.GetResIDstring(1017)
    sTabFx1.AddTab lng.GetResIDstring(1044)
    
'    ROOFPIC.Command6.Caption = lng.GetResIDstring(1045)
'    ROOFPIC.Command4.Caption = lng.GetResIDstring(3003)
    Text1.ToolTipText = lng.GetResIDstring(1123)
    Text5.ToolTipText = lng.GetResIDstring(1124)
    Text4.ToolTipText = Text5.ToolTipText
    
    Command2.Caption = lng.GetResIDstring(9177)
    
    Picture1.BackColor = Setup.Command9.BackColor

    If ScaleTop_Main <> 0 And ScaleWidth_Main <> 0 And ScaleHeight_Main <> 0 Then
    
        On Error GoTo SETSCALE
        ROOFPIC.Picture1.ScaleLeft = ScaleLeft_Main
        ROOFPIC.Picture1.ScaleTop = ScaleTop_Main
        ROOFPIC.Picture1.ScaleWidth = ScaleWidth_Main
        ROOFPIC.Picture1.ScaleHeight = ScaleHeight_Main
        ERR.Clear
        On Error GoTo ERR
    Else
    
SETSCALE:
        
        ScaleLeft_Main = 0
        ScaleWidth_Main = 1200
        ScaleTop_Main = 625
        ScaleHeight_Main = -625
        
        ROOFPIC.Picture1.ScaleLeft = ScaleLeft_Main
        ROOFPIC.Picture1.ScaleTop = ScaleTop_Main
        ROOFPIC.Picture1.ScaleWidth = ScaleWidth_Main
        ROOFPIC.Picture1.ScaleHeight = ScaleHeight_Main
        
       If Me.Visible Then Command5.value = True
        
    End If

    Set SplitHV1.obj1 = Picture1
    Set SplitHV1.obj2 = Frame1
    
    HistoryClear True
    
'    OfficeStart.HistoryWorking = False

    Exit Sub
ERR:
'    STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.22." & ERR.Source & "]", ERR.Number, ERR.Description
    Resume Next
End Sub





Private Sub Form_Resize()
    On Error Resume Next
    SplitHV1.Width = Me.Width
    Picture1.Width = Me.Width
    Frame1.Width = Me.ScaleWidth
    sTabFx1.Width = Frame1.Width - 4520
'    Shape2.Width = sTabFx1.Width - 250
    
'    Shape2.Width = frame1.Width - 200 '- 3200
    Label2.Width = Frame1.Width - 250
    
    Text1.Width = Frame1.Width - 200
    SplitHV1.ResizeControl
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HistoryClear False
End Sub

Sub HScroll1_Change()
    scroll Picture1
    Draw_Systems Picture1
End Sub


Function Find_Point(p0096 As Single, p0098 As Single) As Integer
    On Error Resume Next
    Dim l009A As Single
    Dim l009C
    Dim l00AC As Single
        Find_Point = 0
        l009A = 99999
        For l009C = 1 To MainCountOfPoints Step 1
            l00AC = Sqr((p0096 - Main_Points_X(l009C)) ^ 2 + (p0098 - Main_Points_Y(l009C)) ^ 2)
            If l00AC < l009A Then
                l009A = l00AC
                Find_Point = l009C
            End If

        Next l009C

End Function





Private Sub Functions_Options(X As Single, Y As Single)
    Dim l00C4 As Single
    Dim l00C6 As Single
    Dim l00C8 As Single
    On Error Resume Next
    
    Select Case OptionDMM
        Case "Mdraw"
            ROOFPIC.Text4 = Format$(X - Main_Points_X(Current_P), "#####")
            ROOFPIC.Text5 = Format$(Y - Main_Points_Y(Current_P), "#####")
            l00C4 = Main_Points_X(Current_P) - X
            l00C6 = Main_Points_Y(Current_P) - Y
            If l00C4 = 0 Then
                l00C8 = 0
            Else
                l00C8 = 57.295778 * Atn(l00C6 / l00C4)
            End If

            Label5.Caption = Format$(l00C8, "##0.0#")
        Case "Msel"
            ROOFPIC.Text4 = Format$(Main_Points_X(P_B) - Main_Points_X(P_A), "#####")
            ROOFPIC.Text5 = Format$(Main_Points_Y(P_B) - Main_Points_Y(P_A), "#####")
            l00C4 = Main_Points_X(P_B) - Main_Points_X(P_A)
            l00C6 = Main_Points_Y(P_B) - Main_Points_Y(P_A)
            If l00C4 = 0 Then
                l00C8 = 0
            Else
                l00C8 = 57.295778 * Atn(l00C6 / l00C4)
            End If

            Label5.Caption = Format$(l00C8, "##0.0#")
    End Select

End Sub


Function Find_lape_label(p00D6 As Single, p00D8 As Single) As Single
    Dim xCentr As Single
    Dim l00DC As Single
    Dim l00DE
    
    On Error Resume Next
    
        Find_lape_label = 0
        xCentr = ScaleWidth_Main / 49
        l00DC = 0.9 * xCentr
        For l00DE = 1 To KolvoScatov Step 1
            If p00D6 > Label_X(l00DE) - l00DC Then
                If p00D6 < Label_X(l00DE) + l00DC Then
                    If p00D8 < Label_Y(l00DE) + l00DC Then
                        If p00D8 > Label_Y(l00DE) - l00DC Then
                            Find_lape_label = l00DE
                            Picture1.MousePointer = 15
                            Exit For
                        End If

                    End If

                End If

            End If

        Next l00DE

End Function

Sub Option1_Click()
    On Error Resume Next
    OptionDMM = "Mdraw"
    isSave = True
    P_A = 0: P_B = 0
    'Option1.value = -1
    Picture1.MousePointer = 2
    Text1 = lng.GetResIDstring(1434)
    '  Text1 = lng.GetResIDstring(1445)
    Draw_Systems Me.Picture1
End Sub


Sub Option3_Click()
    On Error Resume Next
    OptionDMM = "Msel"
    P_A = 0: P_B = 0
    Picture1.MousePointer = 1
    'Option3.value = -1
    isSave = True
    Draw_Systems Me.Picture1
    '  Text1 = lng.GetResIDstring(1445)
    Text1 = lng.GetResIDstring(1435)
    '  Picture1.SetFocus
End Sub


Sub Option4_Click()
    On Error Resume Next
    OptionDMM = "Mlabel"
    'Option4.value = -1
    isSave = True
    Picture1.MousePointer = 1
    Draw_Systems Me.Picture1
    Text1 = lng.GetResIDstring(1485)
    '  Text1 = lng.GetResIDstring(1405)
    '  Picture1.SetFocus
End Sub


Private Sub HScroll2_Change()

End Sub

Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim l00F8 As Single
    Dim l00FA As Single
    Dim Slope As Integer
    
    On Error Resume Next
    
    Select Case KeyCode
    Case 112
        sTabFx1.SelectTab 0
        Exit Sub
    Case 113
        sTabFx1.SelectTab 1
        Exit Sub
    Case 114
        sTabFx1.SelectTab 3
        Exit Sub
    End Select
    
    isSave = True

    If OptionDMM = "Mdraw" And FlagDraw = 0 And KeyCode = 46 Then
        OptionDMM = "Mdel"
        Picture1.MousePointer = 10
    End If

    If OptionDMM = "Mdraw" And FlagDraw = 0 And KeyCode = vbKeyControl Then
        OptionDMM = "Mdel"
        Picture1.MousePointer = 10
    End If

    If OptionDMM = "Mlabel" And KeyCode = 46 And KolvoScatov > 0 Then
        
        If KolvoScatov > 26 Then
            Slope = KolvoScatov + 70
        Else
            Slope = KolvoScatov + 64
        End If

        l00F8 = MsgBox(lng.GetResIDstring(1438) & Chr$(Slope) & ") ?", 4)

        If l00F8 = 6 Then
            Call Clear_lape_label(KolvoScatov)
            KolvoScatov = KolvoScatov - 1
            SetChange True
        End If

    End If

    If OptionDMM = "Msel" Then
        
        If Shift = 1 Then
            l00FA = 10
        Else
            If Shift = 2 Then
                l00FA = 100
            Else
                l00FA = 1
            End If

        End If

        Select Case KeyCode
            Case 38
                Main_Points_Y(P_B) = Main_Points_Y(P_B) + l00FA
                SetChange True
            Case 40
                Main_Points_Y(P_B) = Main_Points_Y(P_B) - l00FA
                SetChange True
            Case 37
                Main_Points_X(P_B) = Main_Points_X(P_B) - l00FA
                SetChange True
            Case 39
                Main_Points_X(P_B) = Main_Points_X(P_B) + l00FA
                SetChange True
        End Select

        Draw_Systems Me.Picture1
        Call Functions_Options(0, 0)
        
    End If

    If OptionDMM = "Msel" And FlagDraw = 0 And KeyCode = 45 Then
        Picture1.MousePointer = 5
    End If

End Sub


Private Sub Picture1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Command4_Click
End Sub


Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If OptionDMM = "Mdel" And KeyCode = 46 Then
        OptionDMM = "Mdraw"
        Picture1.MousePointer = 2
        Option1_Click
    End If

    If OptionDMM = "Mdel" And KeyCode = vbKeyControl Then
        OptionDMM = "Mdraw"
        Picture1.MousePointer = 2
        Option1_Click
    End If

    If OptionDMM = "Msel" And KeyCode = 45 Then
        OptionDMM = "Msel"
        Picture1.MousePointer = 1
    End If

    If OptionDMM = "Mlabel" And KeyCode = 46 Then
        Picture1.MousePointer = 1
    End If

End Sub


Private Sub Picture1_LostFocus()
On Error Resume Next

    ScaleLeft_Main = ROOFPIC.Picture1.ScaleLeft
    ScaleTop_Main = ROOFPIC.Picture1.ScaleTop
    ScaleWidth_Main = ROOFPIC.Picture1.ScaleWidth
    ScaleHeight_Main = ROOFPIC.Picture1.ScaleHeight
    
End Sub

Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim l0114 As Single
    Dim l0116
    Dim l0118 As Double
    Dim l011A As Double
    
    On Error Resume Next
  
      If Button = 1 Then
      
          If OptionDMM = "Mdraw" Then
              isSave = True

              If FlagDraw = 0 Then
                
                MainCountOfPoints = MainCountOfPoints + 1
                Main_Points_X(MainCountOfPoints) = X
                Main_Points_Y(MainCountOfPoints) = Y
                Current_P = MainCountOfPoints
                Me.Drawing_cross Current_P, Me.Picture1
                MainCountOfLines = MainCountOfLines + 1
                Points_m_A(MainCountOfLines) = MainCountOfPoints
                Current_P = MainCountOfPoints
                FlagDraw = -1
                GoSub Draw
'                SetChange true
                
              Else
              
                ' Рисование прямых
                If Shift = 1 Then
                    ' первая строка 0 - 90 и 270 - 360
                    If (0.75 <= SinA And SinA <= 1) And (1 >= SinA And SinA >= 0.8) Or _
                    (-0.8 >= SinA And SinA >= -1) And (-0.8 >= SinA And SinA >= -1) Then
                        Y = Main_Points_Y(Current_P)
                    ' вторая строка 90 - 180 и 180 - 270
                    ElseIf (-0.93 <= SinA And SinA >= 0) And (0 <= SinA And SinA <= 0.8) Or _
                    (0 >= SinA And SinA >= -0.8) And (-0.8 <= SinA And SinA <= 0) Then
                        X = Main_Points_X(Current_P)
                    End If
                End If
              
                MainCountOfPoints = MainCountOfPoints + 1
                Main_Points_X(MainCountOfPoints) = X: Main_Points_Y(MainCountOfPoints) = Y
                Points_m_B(MainCountOfLines) = MainCountOfPoints
                Current_P = MainCountOfPoints
                FlagDraw = 0
                GoSub L88BC
                SetChange True
                  
              End If
    
          End If

      End If

      If Button = 2 Then
      
          If OptionDMM = "Mdraw" Then
              If FlagDraw = 0 Then
                  Current_P = Find_Point(X, Y)
                  MainCountOfLines = MainCountOfLines + 1
                  Points_m_A(MainCountOfLines) = Current_P
                  FlagDraw = -1
                  GoSub Draw
              Else
                  l0114 = Find_Point(X, Y)
                  Points_m_B(MainCountOfLines) = l0114
                  FlagDraw = 0
                  GoSub L88BC
                  SetChange True
              End If

          End If

      End If

      If OptionDMM = "Mdel" Then
          Call DEL_P(Find_Point(X, Y))
          OptionDMM = "Mdraw"
          Picture1.MousePointer = 2
          Picture1.SetFocus
      End If

      If OptionDMM = "Msel" Then
          If Picture1.MousePointer = 5 Then
              For l0116 = 1 To MainCountOfLines Step 1
                  If Points_m_A(l0116) = P_A And Points_m_B(l0116) = P_B Or Points_m_A(l0116) = P_B And Points_m_B(l0116) = P_A Then
                      MainCountOfPoints = MainCountOfPoints + 1
                      If Main_Points_X(P_B) = Main_Points_X(P_A) Then
                          Main_Points_X(MainCountOfPoints) = Main_Points_X(P_A)
                          Main_Points_Y(MainCountOfPoints) = Y
                      Else
                          If Main_Points_Y(P_B) = Main_Points_Y(P_A) Then
                              Main_Points_Y(MainCountOfPoints) = Main_Points_Y(P_A)
                              Main_Points_X(MainCountOfPoints) = X
                          Else
                              l0118# = (Main_Points_Y(P_B) - Main_Points_Y(P_A)) / (Main_Points_X(P_B) - Main_Points_X(P_A))
                              l011A# = (Main_Points_X(P_A) - Main_Points_X(P_B)) / (Main_Points_Y(P_B) - Main_Points_Y(P_A))
                              Main_Points_X(MainCountOfPoints) = (Y - Main_Points_Y(P_A) + l0118# * Main_Points_X(P_A) - l011A# * X) / (l0118# - l011A#)
                              Main_Points_Y(MainCountOfPoints) = l0118# * Main_Points_X(MainCountOfPoints) - l0118# * Main_Points_X(P_A) + Main_Points_Y(P_A)
                          End If

                      End If

                      Current_P = MainCountOfPoints
                      Me.Drawing_cross Current_P, Me.Picture1
                      MainCountOfLines = MainCountOfLines + 1
                      Points_m_A(MainCountOfLines) = MainCountOfPoints
                      Points_m_B(MainCountOfLines) = Points_m_B(l0116)
                      Points_m_B(l0116) = MainCountOfPoints
                      P_B = MainCountOfPoints
                      Exit For
                  End If

              Next l0116

          Else
              If Button = 1 Then
                  P_A = Find_Point(X, Y)
              Else
                  P_B = Find_Point(X, Y)
              End If

              Draw_Systems Me.Picture1
          End If

      End If

      If OptionDMM = "Mlabel" Then
          isSave = True
  
          If KolvoScatov > 0 Then LapeName = Find_lape_label(X, Y)
          
              If LapeName > 0 Then ' Передвижения имени ската
                  Clear_lape_label (LapeName) '
                  OptionDMM = "Mlabmov" '
              Else
                                
                  If KolvoScatov < MAXSLOPES Then
                      
                      KolvoScatov = KolvoScatov + 1
                      Label_X(KolvoScatov) = X
                      Label_Y(KolvoScatov) = Y
                      Call Module10.Draw_lape_label(Me.Picture1, KolvoScatov, X, Y)
                      SetChange True
    
                  Else
LIM:
                      MsgBox lng.GetResIDstring(1439) & MAXSLOPES
                      Exit Sub
                  End If
    
              End If

      End If

      Exit Sub

L88BC:
      Line1.Visible = False
      If Points_m_A(MainCountOfLines) = Points_m_B(MainCountOfLines) Then
          MainCountOfLines = MainCountOfLines - 1
      End If

      Draw_Systems Me.Picture1
      Return
Draw:
      Line1.Visible = True
      Call Draw_Line(Main_Points_X(Current_P), Main_Points_Y(Current_P), X, Y)
      Return
End Sub


Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If isActive.isFormFocus(Me.hwnd) Then
        isActive.SetFormFocus Picture1.hwnd
    End If
  
    If FlagDraw = -1 Then
    
        ' Рисование прямых
        If Shift = 1 Then

            SinA = (X - Main_Points_X(Current_P)) / Sqr((X - Main_Points_X(Current_P)) ^ 2 + _
            (Y - Main_Points_Y(Current_P)) ^ 2)

            ' первая строка 0 - 90 и 270 - 360
            If (0.75 <= SinA And SinA <= 1) And (1 >= SinA And SinA >= 0.8) Or _
                (-0.8 >= SinA And SinA >= -1) And (-0.8 >= SinA And SinA >= -1) Then
'                    Y = Lape_Points_Y(N_Slope, Current_P)
                    Y = Main_Points_Y(Current_P)
            ' вторая строка 90 - 180 и 180 - 270
            ElseIf (-0.93 <= SinA And SinA >= 0) And (0 <= SinA And SinA <= 0.8) Or _
                (0 >= SinA And SinA >= -0.8) And (-0.8 <= SinA And SinA <= 0) Then
'                    X = Lape_Points_X(N_Slope, Current_P)
                    X = Main_Points_X(Current_P)
            End If
        End If
        
        Call Draw_Line(Main_Points_X(Current_P), Main_Points_Y(Current_P), X, Y)
        
    End If
  
    Call Functions_Options(X, Y)

End Sub


Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If OptionDMM = "Mlabmov" Then
        Call Module10.Draw_lape_label(Me.Picture1, LapeName, X, Y)
        Label_X(LapeName) = X
        Label_Y(LapeName) = Y
        OptionDMM = "Mlabel"
        Picture1.MousePointer = 1
        LapeName = 0
    End If

End Sub


Sub DEL_P(p014A As Integer)
    Dim l014C
    Dim l014E As Single
    Dim l0150
    On Error Resume Next
    
        If p014A < 1 Then Exit Sub
        For l014C = p014A To MainCountOfPoints - 1 Step 1
            Main_Points_X(l014C) = Main_Points_X(l014C + 1)
            Main_Points_Y(l014C) = Main_Points_Y(l014C + 1)
        Next l014C

        Main_Points_X(MainCountOfPoints) = 0: Main_Points_Y(MainCountOfPoints) = 0
        MainCountOfPoints = MainCountOfPoints - 1
        l014E = -1
        Do While l014E = -1
            l014E = 0
            For l014C = 1 To MainCountOfLines Step 1
                If Points_m_A(l014C) = p014A Or Points_m_B(l014C) = p014A Then
                    GoSub L90F6
                    Exit For
                End If

            Next l014C

        Loop

        For l014C = 1 To MainCountOfLines Step 1
            If Points_m_A(l014C) > p014A Then Points_m_A(l014C) = Points_m_A(l014C) - 1
            If Points_m_B(l014C) > p014A Then Points_m_B(l014C) = Points_m_B(l014C) - 1
        Next l014C

        Call Draw_Systems(Me.Picture1)
        Exit Sub

L90F6:
        l014E = -1
        For l0150 = l014C To MainCountOfLines - 1 Step 1
            Points_m_A(l0150) = Points_m_A(l0150 + 1)
            Points_m_B(l0150) = Points_m_B(l0150 + 1)
        Next l0150

        MainCountOfLines = MainCountOfLines - 1
        Return
End Sub


Sub Draw_Line(ByVal px1 As Integer, ByVal py1 As Integer, ByVal px2 As Single, ByVal py2 As Single, Optional Clear As Boolean)
On Error Resume Next

    Line1.BorderColor = Setup.Command10.BackColor
    Line1.X1 = px1
    Line1.x2 = px2
    Line1.Y1 = py1
    Line1.y2 = py2
    PSx = px2
    PSy = py2
End Sub


Public Sub scroll(P As PictureBox)
    Dim l008A As Single
    Dim l008C As Single
    Dim l008E As Single
    Dim l0090 As Single
    Dim l0092 As Single
    
    On Error Resume Next
    
    l008A = P.ScaleLeft + P.ScaleWidth
    l008C = (P.ScaleLeft + l008A) / 2
    l008E = P.ScaleTop + (P.ScaleHeight / 2)
    l0090 = HScroll1.value
    l0092 = l0090 / RatioH
    P.ScaleTop = l008E + (l0092 / 2)
    P.ScaleHeight = -l0092
    P.ScaleWidth = l0090
    P.ScaleLeft = l008C - (l0090 / 2)
    ScaleLeft_Main = P.ScaleLeft
    ScaleWidth_Main = P.ScaleWidth
    ScaleTop_Main = P.ScaleTop
    ScaleHeight_Main = P.ScaleHeight
End Sub


Sub Drawing_cross(ByVal p0146 As Integer, p0136 As Object)
    Dim Len_AB_X As Single
    Static tcolor As Integer
    On Error GoTo ERR

        Len_AB_X = p0136.ScaleWidth / 350
        tcolor = p0136.FillColor
        p0136.FillColor = vbRed
        p0136.FillStyle = 0
        p0136.Circle (Main_Points_X(p0146), Main_Points_Y(p0146)), Len_AB_X, vbBlack
        
'    aspect = ScaleY(1, vbUser, vbPixels) / ScaleX(1, vbUser, vbPixels)
'    Circle (0, 0), 1, , , , asp
        
        p0136.FillColor = tcolor
        p0136.FillStyle = 1
ERR:
End Sub


Private Sub Picture1_Resize()
    Command5.value = True
End Sub


Private Sub sTabFx1_Click(Index As Integer, Key As String, Caption As String)

SetChange True

Select Case Index
Case 0
Option1_Click
Frame3.ZOrder 0
Frame3.Visible = True
Case 1
Frame3.Visible = False
Option3_Click
Case 2
Frame3.Visible = False
Option4_Click
End Select
End Sub


Private Sub Text1_Click()
On Error Resume Next
Load Teksti
Teksti.Text1 = Text1.Text
Teksti.Show vbModal, OfficeStart
Unload Teksti
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Text4 <> "" And KeyAscii = 13 Then
        Main_Points_X(P_B) = Main_Points_X(P_A) + Text4
        SetChange True
        Draw_Systems Me.Picture1
    End If
End Sub



Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Text5 <> "" And KeyAscii = 13 Then
        Main_Points_Y(P_B) = Main_Points_Y(P_A) + Text5
        SetChange True
        Draw_Systems Me.Picture1
    End If
End Sub


Sub LapeMenuNext()
    On Error GoTo ERR

    Static az As Integer

        If KolvoScatov = 0 Then
            MsgBox lng.GetResIDstring(1458), vbCritical
            OfficeStart.TabStrip1.Tabs(2).Selected = True
            sTabFx1.SelectTab 3
            Exit Sub
        End If

        FlagDraw = 0
'        Lapemenu.Command4.Enabled = 0
        Lapemenu.Show

        Unload Me
        Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.26." & ERR.Source & "]", ERR.Number, ERR.Description
        Resume Next
End Sub


Private Sub Комманда1_Click()
On Error Resume Next
    Picture1.ScaleTop = Picture1.ScaleTop + Picture1.ScaleHeight * 0.1
    Draw_Systems Picture1
End Sub

Private Sub Комманда2_Click()
On Error Resume Next
    Picture1.ScaleLeft = Picture1.ScaleLeft - Picture1.ScaleWidth * 0.1
    Draw_Systems Picture1
End Sub

Private Sub Комманда3_Click()
On Error Resume Next
    Picture1.ScaleLeft = Picture1.ScaleLeft + Picture1.ScaleWidth * 0.1
    Draw_Systems Picture1
End Sub

Private Sub Комманда4_Click()
On Error Resume Next
    Picture1.ScaleTop = Picture1.ScaleTop - Picture1.ScaleHeight * 0.1
    Draw_Systems Picture1
End Sub
