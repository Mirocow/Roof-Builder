VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form SetProjectName 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5925
   ClipControls    =   0   'False
   Icon            =   "SetProjectName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin —hameleonButton.chameleonButton Command2 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "&Cancel"
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
      MICON           =   "SetProjectName.frx":030A
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
   Begin —hameleonButton.chameleonButton Command1 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "&Ok"
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
      MICON           =   "SetProjectName.frx":0326
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
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "¬˚·Ó ÙÓÏ‡Ú‡ Ù‡ÈÎ‡ (“ËÔ ÔÓÂÍÚ‡)"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label1 
      Height          =   900
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "SetProjectName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Combo1 <> "" Then Gl.FileNameExtension = Combo1

    Dim F As String
        F = OfficeStart.CheckFileName(Text1)
        If F <> "" Then
            If Right(Text1, 4) = Gl.FileNameExtension Then
                CurrentFile = F
            Else
                CurrentFile = F & Gl.FileNameExtension
            End If

        End If
        
        Project.SwitchProfil
        
        Unload Me
End Sub


Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    On Error Resume Next

    SetFont Me

    Call VarPtr("VMProtect begin")
    If Gl.PV = "Prof  " Then
        Label2 = lng.GetResIDstring(9558)
        Combo1.Visible = True
        Label2.Visible = True
        Combo1.AddItem ".rfd"
        Combo1.AddItem ".rbp"
        If Gl.FileNameExtension = ".rfd" Then
            Combo1.ListIndex = 0
        Else
            Combo1.ListIndex = 1
        End If

    Else
        Combo1.Enabled = False
    End If
    Call VarPtr("VMProtect end")

    Label1 = lng.GetResIDstring(1110)
    Dim L As Integer
        L = InStr(Len(Project.Text1) - 4, Project.Text1, ".")
        If L > 0 Then
            Text1 = Left(Project.Text1, L - 1)
        Else
            Text1 = Project.Text1
        End If

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.value = True
    End If

End Sub

