VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form Project 
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   8835
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Project.frx":0000
   LinkTopic       =   "‘ÓÏ‡1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7005
   ScaleWidth      =   8835
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   1020
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1095
         Left            =   0
         TabIndex        =   20
         Top             =   2160
         Width           =   8655
         Begin —hameleonButton.chameleonButton chameleonButton1 
            Height          =   735
            Left            =   7800
            TabIndex        =   21
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
            BTYPE           =   7
            TX              =   "+"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
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
            MICON           =   "Project.frx":030A
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   7575
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   7575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   290
            Width           =   7575
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1440
         Left            =   0
         TabIndex        =   24
         Top             =   3240
         Width           =   8655
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1245
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   2050
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1260
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   2050
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1260
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   2050
         End
         Begin VB.Label Label9 
            Caption         =   "“ÓÎ˘ËÌ‡"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "œÓÍ˚ÚËÂ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "÷‚ÂÚ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   1095
         End
      End
      Begin —hameleonButton.chameleonButton Command4 
         Height          =   615
         Left            =   6240
         TabIndex        =   17
         Top             =   6120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
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
         MICON           =   "Project.frx":0326
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
      Begin —hameleonButton.chameleonButton Check1 
         Height          =   300
         Left            =   6840
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
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
         MICON           =   "Project.frx":0342
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
         ICONS           =   16
      End
      Begin VB.CommandButton Command13 
         Caption         =   "ƒÓ·‡‚ËÚ¸/»ÁÏÂÌËÚ¸"
         Enabled         =   0   'False
         Height          =   300
         Left            =   6840
         TabIndex        =   1
         Top             =   9075
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1800
         MaxLength       =   249
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin —hameleonButton.chameleonButton Command5 
         Height          =   300
         Left            =   6840
         TabIndex        =   18
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
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
         MICON           =   "Project.frx":035E
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
      Begin —hameleonButton.chameleonButton Command3 
         Height          =   300
         Left            =   5160
         TabIndex        =   19
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   529
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
         MICON           =   "Project.frx":037A
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Label5"
         Height          =   735
         Left            =   120
         TabIndex        =   32
         Top             =   2400
         Width           =   8535
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "»Ïˇ ÍÎËÂÌÚ‡:"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   9075
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2760
         TabIndex        =   5
         Top             =   9075
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         TabIndex        =   4
         Top             =   9075
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "œÛÚ¸:"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label Label19 
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1215
         Left            =   0
         MouseIcon       =   "Project.frx":0396
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   4800
         Width           =   8655
      End
      Begin VB.Label Label22 
         Caption         =   "ŒÔËÒ‡ÌËÂ / ÍÓÏÏÂÌÚ‡ËË:"
         Height          =   615
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "œÓÂÍÚ ÔÓ‰„ÓÚÓ‚ËÎ:"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "»Ïˇ Ù‡ÈÎ‡:"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
Load ChangeProfil
'Label3.Caption = "*"
ChangeProfil.lstprof.ListIndex = 0
ChangeProfil.lstprof_Click
ChangeProfil.Show vbModal, OfficeStart

If ChangeProfil.ListView1.ListItems.Count = 0 Then
Unload ChangeProfil
Exit Sub
End If

' Factory
Label2.Caption = ChangeProfil.Label24.Caption
Label2.Tag = ChangeProfil.Label24.Tag

' Profil
Label3.Caption = ChangeProfil.ComboBox1.List(ChangeProfil.ComboBox1.ListIndex)
Label3.Tag = ChangeProfil.ListView1.ListItems(1).SubItems(1)

Unload ChangeProfil

SwitchProfil
End Sub

Private Sub Check1_Click()
    If Check1.value = True Then
        Text2.Text = Text6.Text
        Text2.Visible = True
        Text6.SetFocus
    Else
        Text6.Locked = True
        Text2.Visible = False
        Text6.Text = Text2.Text
        Text6.BackColor = "&H8000000F"
    End If

End Sub


Private Sub Combo1_Change(Index As Integer)
    Select Case Index
        Case 1
            width1 = Combo1(1).Text
        Case 2
            cover = Combo1(2).Text
        Case 3
            ColorRoof = Combo1(3).Text
    End Select
End Sub


Private Sub Command3_Click()
    SetProjectName.Show vbModal, OfficeStart
End Sub


Public Sub Combo1_Click(Index%)
    Select Case Index
        Case 1
            width1 = Project.Combo1(1).Text
      
        Case 2
            cover = Project.Combo1(2).Text
      
        Case 3
            ColorRoof = Project.Combo1(3).Text
      
    End Select

End Sub


Public Sub Command4_Click()
    OfficeStart.TabStrip1.Tabs(2).Selected = True
End Sub


Private Sub Command5_Click()
    Setup.Show
    Setup.TabStrip2.Tabs(1).Selected = True
End Sub


Sub Form_Load()
    Dim name As String
    'On Error Resume Next
    On Error GoTo ERR

    Label14 = Date

    SetFont Me

splash.SetProgress 8

    ' Á‡„ÛÁÍ‡ ÔÓËÁ‚Ó‰ËÚÂÎˇ
    Factory_Name = GetSetting(App.ProductName, "Main", "Factory_Name", "NOGROUP")
    Label2.Caption = Factory_Name
    
splash.SetProgress 9

    Project.Command4.Caption = lng.GetResIDstring(1033)
    Check1.Caption = lng.GetResIDstring(9378)
    Command5.Caption = lng.GetResIDstring(9641)
    Label12.Caption = lng.GetResIDstring(9642)
    Command3.Caption = lng.GetResIDstring(9379)
    Command13.Caption = lng.GetResIDstring(9403)
    Label9.Caption = lng.GetResIDstring(9390)
    Label16.Caption = lng.GetResIDstring(9391)
    Label17.Caption = lng.GetResIDstring(9392)
    Label27.Caption = lng.GetResIDstring(9404)
    Label22.Caption = lng.GetResIDstring(9405)
    Label26.Caption = lng.GetResIDstring(9382)
    Label1.Caption = lng.GetResIDstring(1027)
    Label19.Caption = lng.GetResIDstring(30002)
    Frame2.Caption = lng.GetResIDstring(9027)
    Label4.Caption = lng.GetResIDstring(1008)
    Label5.Caption = lng.GetResIDstring(1023)
    
    TimeStart = Now
'    OfficeStart.Timer1.Enabled = True

    Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ":load) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.33." & ERR.Source & "]", ERR.Number, ERR.Description
        Resume Next
End Sub



Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OfficeStart.OpenFilePreload Data.Files.Item(1)
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Frame1.Top = (Me.ScaleHeight / 2 - Frame1.Height / 2) + 400
    Frame1.Left = Me.ScaleWidth / 2 - Frame1.Width / 2
End Sub


Sub Form_Unload(Cancel As Integer)
Dim ans As Integer
    If FEXIT = False Then
        Cancel = -1
        If Me.Visible = True Then Me.Hide
    End If
End Sub

Private Sub Label19_Click()
    Navigate Me, "http://www.firma-ms.ru/forum/index.php?showforum=13"
End Sub


Sub SwitchProfil()
If Gl.FileNameExtension = ".rfd" Then

    Project.Frame2.Visible = True
            
    If Project.Label3.Caption = "" Then
        Project.Label4.Visible = True
    Else
        Project.Label4.Visible = False
        
        Project.Combo1(3).Clear
        Project.Combo1(1).Clear
        Project.Combo1(2).Clear
        
        LoadProfilAdditionalData "Color", Project.Combo1(3), Project.Label3.Tag, CurrentLocale
        LoadProfilAdditionalData "Thickness", Project.Combo1(1), Project.Label3.Tag, CurrentLocale
        LoadProfilAdditionalData "Coating", Project.Combo1(2), Project.Label3.Tag, CurrentLocale
    End If
    
Else
    Project.Frame2.Visible = False
End If
End Sub
