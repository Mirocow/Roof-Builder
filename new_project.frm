VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form newproject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New project"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6345
   Icon            =   "new_project.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6345
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   34
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "new_project.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "new_project.frx":04A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      CausesValidation=   0   'False
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      View            =   2
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      MouseIcon       =   "new_project.frx":0652
      OLEDragMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Recent"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "newproject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
If ListView1.ListItems.Count = 0 Then Exit Sub

Dim indexselected As String
indexselected = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
Select Case TabStrip1.SelectedItem.Index
    Case 1
    
        Gl.FileNameExtension = indexselected
        If CurrentFile = "" Then CurrentFile = Project.Label14 & Gl.FileNameExtension '"Project1" & Gl.FileNameExtension
        OfficeStart.StatusBar.Panels(2) = ""
        
        Select Case indexselected
        Case ".rfd"
'            Project.Label3.Enabled = True
'            Project.Label2.Enabled = True
'            Project.Label4.Enabled = True
'            Project.chameleonButton1.Enabled = True
            Project.Label3.Caption = ""
        Case ".rbp"
'            Project.Label3.Enabled = False
'            Project.Label2.Enabled = False
'            Project.Label4.Enabled = False
'            Project.chameleonButton1.Enabled = False
            Project.Label3.Caption = ""
        End Select
        
        Project.SwitchProfil
        
    Case 2
    
        Dim tempname As String
        tempname = Right(indexselected, Len(indexselected) - InStrRev(indexselected, "\", -1))
        CurrentFile = Left(tempname, Len(tempname) - 4)
        ProjectsDir = Left$(indexselected, InStrRev(indexselected, "\", -1))
        If OfficeStart.OpenFilePreload(indexselected) = "" Then Exit Sub
        
End Select

isSave = True
Unload Me
Project.Show
End Sub


Private Sub Command3_Click()
    Unload Me
End Sub




Private Sub Form_Load()
    'Label3 = Gl.Firm_name & " " & chr$(174)

    SetFont Me

    'Me.Left = GetSetting(App.ProductName, "Position", Me.Name & "left", Me.Left)
    'Me.Top = GetSetting(App.ProductName, "Position", Me.Name & "top", Me.Top)
    'Me.Width = GetSetting(App.ProductName, "Position", Me.name & "width", Me.Width)
    'Me.Height = GetSetting(App.ProductName, "Position", Me.name & "height", Me.Height)

    Me.Caption = lng.GetResIDstring(9277) '"New Workname"
    Label3 = Me.Caption
    Me.Command3.Caption = lng.GetResIDstring(3002)
    Me.Command2.Caption = lng.GetResIDstring(3017)
    TabStrip1.Tabs(1).Caption = lng.GetResIDstring(9278)
    TabStrip1.Tabs(2).Caption = lng.GetResIDstring(9279)

    TabStrip1.Tabs(1).Selected = True
    ListView1_ItemClick ListView1.ListItems(1)

    'If Dir(ProjectsDir, vbDirectory) = "" Then
    'MsgBox lng.GetResIDstring(1469), vbCritical
    'DeleteSetting App.ProductName, "Main", "url_work"
    'OfficeStart.initializa
    'Exit Sub
    'End If

    'dir1.Path = ProjectsDir
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'SaveSetting App.ProductName, "Position", Me.Name & "left", Me.Left
    'SaveSetting App.ProductName, "Position", Me.Name & "top", Me.Top
    'SaveSetting App.ProductName, "Position", Me.name & "width", Me.Width
    'SaveSetting App.ProductName, "Position", Me.name & "height", Me.Height
End Sub


'Private Sub Text2_Change()
''Text2 = LCase(Text2)
'If Text2 = "" Then Exit Sub
'If Right(Text2, 4) = Combo1 Then Text2 = Left(Text2, Len(Text2) - 4)
'If Len(ProjectsDir) > 3 And Right(ProjectsDir, 1) <> "\" Then ProjectsDir = ProjectsDir + "\"
'If Dir$(ProjectsDir & LCase(Text2) & Combo1) <> "" Then
'Label1 = lng.GetResIDstring(1436) '"Workname already exists!" &  vbNewLine  & "Will rewrite it?"
'Command2.Caption = lng.GetResIDstring(3016)
'Else
'If Len(ProjectsDir) > 3 Then
'Label1 = Dir1.Path & "\" & Text2 & Combo1
'Else
'Label1 = Dir1.Path & Text2 & Combo1
'End If
'Command2.Caption = lng.GetResIDstring(3010)
'End If
'End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2_Click
    End If

End Sub


Private Sub ListView1_DblClick()
    Command2_Click
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If TabStrip1.SelectedItem.Index = 1 Then

        Select Case Item.Index
            Case 1
                Label1 = lng.GetResIDstring(9009)
            Case 2
                Label1 = lng.GetResIDstring(9010)
        End Select

    Else
        If dir(Item.SubItems(1)) <> "" Then
            Label1 = lng.GetResIDstring(9276) & FileDateTime(Item.SubItems(1)) & vbNewLine & _
lng.GetResIDstring(9275) & Format(FileLen(Item.SubItems(1)) / 1024, "0.00") & " kb"
        Else

        End If

    End If

End Sub


Private Sub TabStrip1_Click()
    Dim itmX As ListItem
        On Error GoTo ERR
        ListView1.ListItems.Clear
        Label1 = ""
        Select Case TabStrip1.SelectedItem.Index
            Case 1

                ListView1.View = lvwIcon
                ListView1.Icons = ImageList1
                Set itmX = ListView1.ListItems.Add(, , "RFD", 2)
                itmX.SubItems(1) = ".rfd"
                
                Call VarPtr("VMProtect begin")
                If Gl.PV = "Prof  " Then
                    Set itmX = ListView1.ListItems.Add(, , "Roof Builder Project", 1)
                    itmX.SubItems(1) = ".rbp"
                End If
                Call VarPtr("VMProtect end")
    
            Case 2

                ListView1.View = lvwReport
    Dim i As Integer
        For i = 0 To UBound(RecentlyFiles)
            If RecentlyFiles(i) <> "" Then
                Set itmX = ListView1.ListItems.Add(, , Right(RecentlyFiles(i), Len(RecentlyFiles(i)) - InStrRev(RecentlyFiles(i), "\", -1)))
                itmX.SubItems(1) = RecentlyFiles(i) 'Left$(RecentlyFiles(i), InStrRev(RecentlyFiles(i), "\", -1))
            End If

        Next
    
End Select

ERR:
End Sub

