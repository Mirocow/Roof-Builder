VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form FactoryNames 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Factory"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4740
   Icon            =   "ListCountry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ÑhameleonButton.chameleonButton Command3 
      Height          =   300
      Left            =   4080
      TabIndex        =   5
      Top             =   45
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      BTYPE           =   7
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      MICON           =   "ListCountry.frx":030A
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
   Begin ÑhameleonButton.chameleonButton Command1 
      Height          =   300
      Left            =   3120
      TabIndex        =   4
      Top             =   45
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      BTYPE           =   7
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      MICON           =   "ListCountry.frx":0326
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4510
      _ExtentX        =   7964
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Factory"
         Object.Width           =   9596
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   2780
      Width           =   4640
      Begin VB.Label Label2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
   End
   Begin ÑhameleonButton.chameleonButton Command2 
      Height          =   300
      Left            =   3600
      TabIndex        =   6
      Top             =   50
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      BTYPE           =   7
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      MICON           =   "ListCountry.frx":0342
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2355
      Left            =   45
      Top             =   405
      Width           =   4650
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   2355
      Left            =   60
      Top             =   420
      Width           =   4650
   End
End
Attribute VB_Name = "FactoryNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Factory_Name = InputBox(lng.GetResIDstring(1499), "")
    
    On Error GoTo ERR

    If Factory_Name <> "" Then
    ' ïðîèçâîäèòåëÿ íåò, íàäî äîáàâèòü
    Dim Count As Integer
    Dim PDataRS As Recordset
    
        Set PDataRS = RequestSQL("select ID from FirmFactory where Name=" & "'" & Trim(Factory_Name) & "'")
    
        If PDataRS Is Nothing Then
            Set PDataRS = RequestSQL("select max(id) from FirmFactory")
            Count = IIf(IsNull(PDataRS.Fields(0)), 1, PDataRS.Fields(0) + 1)
            PDataRS.Close
            ' Äîáàâèòü èìÿ ïðîôèëÿ
            Connect Gl.FileName, -Val(iData(0)) 'And (GetAttr(strFileName) And vbReadOnly)
            Execute "insert into `FirmFactory` (ID,NAME,URL) values " & _
                "('" & Count & "','" & Factory_Name & "','')"
            Connect Gl.FileName, True
        End If
    
        LoadFactories
    End If
    
ERR:
End Sub


Private Sub Command2_Click()
    On Error GoTo ERR
    If ListView1.ListItems.Count = 0 Then
        ChangeProfil.Label24 = "*"
        ChangeProfil.Command2.Enabled = False
        Exit Sub
    End If
    Connect Gl.FileName, -Val(iData(0))
    DelBaseData "select id from FirmFactory where id=" & ListView1.SelectedItem.Text
    Connect Gl.FileName, True
    LoadFactories
ERR:
End Sub


Private Sub Command3_Click()
    If ListView1.ListItems.Count = 0 Then
        ChangeProfil.Label24 = "*"
        ChangeProfil.Command2.Enabled = False
        Exit Sub
    End If
    Factory_Name = InputBox(lng.GetResIDstring(1499), , ListView1.SelectedItem.ListSubItems(1))
    If Factory_Name <> "" Then
        ListView1.SelectedItem.ListSubItems(1) = Factory_Name
    End If
End Sub


Private Sub Form_Load()
    SetFont Me
    Me.Caption = lng.GetResIDstring(9201)
    LoadFactories
    If ListView1.ListItems.Count > 0 Then
        ListView1.ListItems(1).Selected = True
        ListView1_Click
    End If
End Sub


Sub LoadFactories()

    Dim RS As Recordset

    Set RS = RequestSQL("select * from FirmFactory order by FirmFactory.id")
    Dim itmXlist As ListItem

    ListView1.ListItems.Clear
    If Not RS Is Nothing Then
        Do While Not RS.EOF
            If RS.Fields(1) <> "Null" Then
                Set itmXlist = ListView1.ListItems.Add(, , RS.Fields(0))
                itmXlist.SubItems(1) = RS.Fields(1)
            End If
            RS.MoveNext
        Loop
        RS.Close
        Set RS = Nothing
        Set itmXlist = Nothing
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    If ListView1.ListItems.Count = 0 Then
        ChangeProfil.Command2.Enabled = False
        ChangeProfil.Label24 = "*"
    End If
    ChangeProfil.lstprof_Click
'    If ChangeProfil.ComboBox1.Enabled = False Then ChangeProfil.Label24 = "*"
End Sub


Private Sub ListView1_Click()
    On Error Resume Next
    
    If ListView1.ListItems.Count = 0 Then
        ChangeProfil.Command2.Enabled = False
        ChangeProfil.Label24 = "*"
        Exit Sub
    End If
    Dim RS As Recordset
    Set RS = RequestSQL("select url from FirmFactory where name='" & ListView1.SelectedItem.SubItems(1) & "'")
    If Not RS Is Nothing Then If RS!url <> "Null" Then Label2 = RS!url Else Label2 = "..." Else Label2 = "..."
    Label1 = ListView1.SelectedItem.SubItems(1)
End Sub


Private Sub ListView1_DblClick()
    On Error Resume Next
    
    If ListView1.ListItems.Count = 0 Then
    ChangeProfil.Command2.Enabled = False
    Unload Me
    Exit Sub
    End If
    
    ' Î÷èùàåì âñå ïîëÿ
    ChangeProfil.ListView1.ListItems.Clear
'    ChangeProfil.lstprof.Clear
    ChangeProfil.ComboBox1.Clear
    
    ChangeProfil.Command2.Enabled = True
    ChangeProfil.Label24.Tag = ListView1.SelectedItem.Text
    ChangeProfil.Label24 = ListView1.SelectedItem.SubItems(1)
    Factory_Name = ChangeProfil.Label24
    SaveSetting App.ProductName, "Main", "Factory_Name", ListView1.SelectedItem.SubItems(1)
    
    Connect Gl.FileName, -Val(iData(0))
    SaveFactoryData ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.Text
    Connect Gl.FileName, True
    
    Unload Me
End Sub

