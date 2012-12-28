VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5DBA0776-6A79-11D9-B3E9-00301B254912}#1.0#0"; "FrameHide.ocx"
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form SO 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search & Open"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11340
   Icon            =   "search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11340
   StartUpPosition =   1  'CenterOwner
   Begin СhameleonButton.chameleonButton Command2 
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   6240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "..."
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
      MICON           =   "search.frx":0E42
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6810
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   661
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17640
            MinWidth        =   17640
            Text            =   "..."
            TextSave        =   "..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   1
            Text            =   "0"
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   5535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6015
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10610
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483646
      BackColor       =   14745339
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "f"
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "p"
         Text            =   "Project"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Create Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "a"
         Text            =   "Agent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "pr"
         Text            =   "Product"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "pa"
         Text            =   "Path"
         Object.Width           =   9596
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Ver"
         Object.Width           =   882
      EndProperty
   End
   Begin Frame.FrameHide FrameHide1 
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   344
      Caption         =   "Date list"
      ForeColor       =   0
      FontName        =   "MS Sans Serif"
      MaximizeSize    =   5775
      isMinimize      =   -1  'True
      Begin VB.CommandButton Command4 
         Height          =   255
         Left            =   480
         Picture         =   "search.frx":0E5E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   120
         Picture         =   "search.frx":10A8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5055
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   8916
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         PathSeparator   =   "."
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Frame.FrameHide FrameHide2 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10186
      Caption         =   "Main picture"
      ForeColor       =   0
      FontName        =   "MS Sans Serif"
      MaximizeSize    =   5775
      Begin VB.PictureBox ShowMP 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0FEFB&
         Height          =   3015
         Left            =   120
         ScaleHeight     =   2955
         ScaleWidth      =   3435
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0FEFB&
         Height          =   2370
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3240
         Width           =   3480
      End
   End
   Begin СhameleonButton.chameleonButton Check1 
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   6240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "Сканировать подпапки"
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
      MICON           =   "search.frx":12F2
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
   Begin СhameleonButton.chameleonButton Command6 
      Height          =   495
      Left            =   8760
      TabIndex        =   12
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "&Open"
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
      MICON           =   "search.frx":130E
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
   Begin СhameleonButton.chameleonButton Command1 
      Height          =   495
      Left            =   9990
      TabIndex        =   13
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "&Exit"
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
      MICON           =   "search.frx":132A
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
   Begin VB.Menu menuw 
      Caption         =   "View"
      Begin VB.Menu menshdatlist 
         Caption         =   "Show sort date list"
      End
      Begin VB.Menu menshprojinfo 
         Caption         =   "Show project info"
      End
   End
   Begin VB.Menu m 
      Caption         =   "Menu"
      Begin VB.Menu mscan 
         Caption         =   "Scan"
         Begin VB.Menu scanrbp 
            Caption         =   "*.rbp (Roof Builder Project)"
            Checked         =   -1  'True
         End
         Begin VB.Menu scanrfd 
            Caption         =   "*.rfd (RFD)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mopen 
         Caption         =   "Open"
      End
      Begin VB.Menu mdel 
         Caption         =   "Del"
      End
   End
   Begin VB.Menu mencurdir 
      Caption         =   "Current:"
   End
End
Attribute VB_Name = "SO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StopSearch As Boolean

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private strFolder As String

Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const MAX_PATH = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                                                              lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                                                            lpFindFileData As WIN32_FIND_DATA) As Long

Private columselect As Integer
Private itemselect As Object

Private MonthNames(12) As String

Private Sub Check1_Click()
If dir(Text2.Text, vbDirectory) <> "" Then FilesFil Text2.Text
End Sub


Private Sub Command1_Click()
    OfficeStart.MousePointer = 0
    OfficeStart.Enabled = True
    Me.Hide
End Sub


Private Sub Command2_Click()
On Error Resume Next
Dim sdir As String
sdir = Dialog.BrowseFolders(hwnd, "Select a Folder", BrowseForFolders, CSIDL_DESKTOP, Text2) '+весь компьютер
If sdir <> "" Then Text2 = sdir
If dir(Text2.Text, vbDirectory) <> "" Then FilesFil Text2.Text
End Sub

Private Sub Command3_Click()
    Me.MousePointer = 11
    DatSorting True
    Me.MousePointer = 1
End Sub


Private Sub Command4_Click()
    Me.MousePointer = 11
    DatSorting
    Me.MousePointer = 1
End Sub


Private Sub Command6_Click()
    ListView1_DblClick
End Sub

Private Sub Form_Load()
    On Error GoTo ERR

    SetFont Me
    
    Text2.Text = ProjectsDir

    Me.Caption = lng.GetResIDstring(9481)
    Check1.Caption = lng.GetResIDstring(9483)
    Command6.Caption = lng.GetResIDstring(9484)
'    Command6.Caption = lng.GetResIDstring(9486)
    Command1.Caption = lng.GetResIDstring(9487)
'    Label2.Caption = lng.GetResIDstring(9489)
    FrameHide1.Caption = lng.GetResIDstring(9490)
    FrameHide2.Caption = lng.GetResIDstring(9491)

    menuw.Caption = lng.GetResIDstring(9506)

    menshdatlist.Caption = lng.GetResIDstring(9507)
    menshprojinfo.Caption = lng.GetResIDstring(9509)

'    Dim hSysMenu As Long
'        hSysMenu = GetSystemMenu(Me.hwnd, 0)
'        RemoveMenu hSysMenu, 6, MF_BYPOSITION

        MonthNames(0) = lng.GetResIDstring(9511)
        MonthNames(1) = lng.GetResIDstring(9512)
        MonthNames(2) = lng.GetResIDstring(9513)
        MonthNames(3) = lng.GetResIDstring(9514)
        MonthNames(4) = lng.GetResIDstring(9515)
        MonthNames(5) = lng.GetResIDstring(9516)
        MonthNames(6) = lng.GetResIDstring(9517)
        MonthNames(7) = lng.GetResIDstring(9518)
        MonthNames(8) = lng.GetResIDstring(9519)
        MonthNames(9) = lng.GetResIDstring(9520)
        MonthNames(10) = lng.GetResIDstring(9521)
        MonthNames(11) = lng.GetResIDstring(9522)

        ProjectsDir = GetSetting(App.ProductName, "Main", "url_work", ProjectsDir)
        FilesFil ProjectsDir
        
        columselect = 0

        Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.38." & ERR.Source & "]", ERR.Number, ERR.Description
        Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If FEXIT = False Then Cancel = -1
    SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
    SaveSetting App.ProductName, "Position", Me.name & "top", Me.Top
'    SaveSetting App.ProductName, "Position", Me.name & "width", Me.Width
'    SaveSetting App.ProductName, "Position", Me.name & "height", Me.Height
End Sub


Private Sub FrameHide2_ClickBoxRoll()
    If FrameHide1.isMinimize Then
        FrameHide1.isMinimize = False
        FrameHide2.Top = FrameHide1.MaximizeSize + 200
    Else
        FrameHide1.isMinimize = True
        FrameHide2.Top = FrameHide1.Height + 200
    End If
End Sub


Private Sub FrameHide1_ClickBoxRoll()
    If FrameHide1.isMinimize Then
        FrameHide2.isMinimize = False
        FrameHide2.Top = FrameHide1.Height + 200 '+ 150
    Else
        FrameHide2.isMinimize = True
        FrameHide2.Top = FrameHide1.MaximizeSize + 200
    End If
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    If ListView1.SortOrder = 0 Then
        ListView1.SortOrder = 1
    Else
        ListView1.SortOrder = 0
    End If

    ListView1.Sorted = True
End Sub


Private Sub ListView1_DblClick()
If Not Command1.Enabled = False Then
    Me.Hide
    OfficeStart.OpenFilePreload itemselect.ListSubItems.Item(5), , True
End If
End Sub



Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim str As String
Dim dstr As Double
Dim leninfo As Integer
Dim i As Integer
Dim strn As Single

    On Error Resume Next

    If Not Command1.Enabled = False Then
    
        OfficeStart.Clear_project
        
        str = ""
        strn = 0
        i = 1
        Set itemselect = Item
        If Item <> "" Then Text3 = Item.ListSubItems.Item(1): If Text3 = "" Then Option2.Enabled = False Else Option2.Enabled = True
        
        Gl.FileNameExtension = GetFileExtension(Item.ListSubItems.Item(5))
        
        Dim cf As Object
        Set cf = Setup.ws_Getdata
        If cf.FOpen(Item.ListSubItems.Item(5), 0) Then
            
            If cf.FN = 0 Then GoTo ERR
            If cf.FLOF() = 0 Then GoTo ERR
            
            ' Начало загрузки данных проекта
            GetDataProject cf, False
            ' Начало загрузки переменных главного рисунка
            GetMainData cf
            cf.FRead KolvoScatov
            cf.FClose
        End If
        Set cf = Nothing
        
        Text3 = ""
        If KolvoScatov > 52 Then Text3 = Text3 & "FILE IS CORRUPTED" & vbCrLf
        Text3 = Text3 & "Project`s name: " & Item & vbCrLf
        Text3 = Text3 & "File size: " & Item.ListSubItems.Item(6) & vbCrLf
        Text3 = Text3 & "File version: " & FileVersion & vbCrLf
        Text3 = Text3 & "Describ: " & PrjDescrib & vbCrLf
        Text3 = Text3 & "Agent: " & UserCreatProject & vbCrLf
        Text3 = Text3 & "Factory: " & Factory_Name & vbCrLf
        Text3 = Text3 & "ProfilName`s name: " & Profil_Name & " " & width1 & " " & cover & " " & ColorRoof & " " & vbCrLf
        Text3 = Text3 & "Created: " & Item.ListSubItems.Item(2) & vbCrLf
        Text3 = Text3 & "Amount of slopes: " & KolvoScatov & vbCrLf
        
        Me.ShowMP.ScaleLeft = ScaleLeft_Main
        Me.ShowMP.ScaleWidth = ScaleWidth_Main
        Me.ShowMP.ScaleTop = ScaleTop_Main
        Me.ShowMP.ScaleHeight = ScaleHeight_Main
        
        Load ROOFPIC
        ROOFPIC.SetPicCentr ShowMP
        If MainCountOfLines > 0 Then
            Module10.Draw_Systems Me.ShowMP
        Else
            Module10.Draw_Systems Me.ShowMP, "PICTURE ERROR"
        End If
        Unload ROOFPIC
        
    End If

    Exit Sub

ERR:
    'Label4.Visible = True
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    'Label4 = Lng.GetResIDstring(1452, "%CATALOGUE%", "", "%FILE%", Item.ListSubItems.Item(6))
    'OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.39." & ERR.Source & "]", ERR.Number, ERR.Description
End Sub


Function open_file(FILE As String) As String
    Dim itmX As ListItem

    On Error GoTo ERR

    Set itmX = ListView1.ListItems.Add(, , Right(FILE, Len(FILE) - InStrRev(FILE, "\", -1)))
    
    itmX.SubItems(2) = FileDateTime(FILE)  ' Время создания файла

    Dim cf As Object
    Set cf = Setup.ws_Getdata
    If cf.FOpen(FILE, 0) Then
        
        If cf.FN = 0 Then GoTo ERR
        If cf.FLOF() = 0 Then GoTo ERR
        
        ' Начало загрузки данных проекта
        GetDataProject cf, False
    
        itmX.SubItems(1) = PrjDescrib
        itmX.SubItems(3) = UserCreatProject
        itmX.SubItems(4) = Profil_Name & " " & width1 & " " & cover & " " & ColorRoof & " "
        itmX.SubItems(5) = LCase(FILE)
        itmX.SubItems(6) = Format(cf.FLOF() / 1024, "0.00") & " kb"
        itmX.SubItems(7) = FileVersion
        cf.FClose
    End If
    Set cf = Nothing
    
    Label1 = "Found: " & Me.ListView1.ListItems.Count
    open_file = "[ok]"

    Exit Function
ERR:
'    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.40." & ERR.Source & "]", ERR.Number, ERR.Description
    open_file = "[err]"
    ListView1.ListItems.Remove itmX.Index
    Resume Next
End Function


Private Sub ListView1_KeyDown(KeyCode As Integer, _
                              Shift As Integer)
    If KeyCode = 13 Then
        ListView1_DblClick
    ElseIf KeyCode = 46 Then
'        Me.Command2.value = True
    End If

End Sub



Private Sub mdel_Click()
    If ListView1.SelectedItem.Index <> 0 Then
        Kill ListView1.SelectedItem.SubItems(6)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    End If
End Sub

Private Sub mencurdir_Click()
        Dim dir As String
        dir = Dialog.BrowseFolders(hwnd, "Select a Folder", BrowseForFolders, CSIDL_DESKTOP, ProjectsDir) '+весь компьютер
        If dir <> "" Then ProjectsDir = dir: FilesFil ProjectsDir
End Sub


Private Sub menshdatlist_Click()
    FrameHide1.isMinimize = False
End Sub


Private Sub menshprojinfo_Click()
    FrameHide1.isMinimize = True
End Sub


Private Sub mopen_Click()
ListView1_DblClick
End Sub

Private Sub Option1_Click()
    ShowMP.Visible = True
End Sub


Private Sub Option2_Click()
    ShowMP.Visible = False
End Sub


Private Sub Text1_Change()
    Dim itmX As ListItem
        'ListView1.FindItem(Text3, 1, 1, 0)
        'ListView1.FindItem(Text1, 1, , 1)
        'If Combo1.ListIndex = -1 Then
        'MsgBox "Please choose a method of search"
        'Combo1.SetFocus
        'Exit Sub
        'End If
        Set itmX = ListView1.FindItem(Text1.Text, 0, 1, 1)
        If itmX Is Nothing Then  ' If no match, inform user and exit.
            Exit Sub
        Else
            Set itemselect = itmX
            itmX.EnsureVisible ' Scroll ListView to show found ListItem.
            itmX.Selected = True   ' Select the ListItem.
            '       ListView1.SetFocus
        End If

End Sub


Private Sub SearchForFolders(fName As String, Path As String)
    'fName - указывает, какие подкаталоги будем искать("*" - все подкаталоги, как и в нашем случае); Path - указывает, в какой папке будем искать; File - указывает, какой файл будем искать.
    If StopSearch = True Then Exit Sub 'переменная StopSearch указывает, должен ли быть прерван поиск.
    Dim Atr As Integer
    Dim hFnd As Long
    Dim WFD As WIN32_FIND_DATA
    Dim i As Integer
    Dim flag As Boolean
        
    If Right(Path, 1) <> "\" Then Path = Path & "\"

    ' Каталог
    hFnd = FindFirstFile(Path & fName, WFD) 'ищем первый подкаталог.
    Me.StatusBar1.Panels(1).Text = Path
'        Me.StatusBar1.Refresh

    If hFnd = INVALID_HANDLE_VALUE Then Exit Sub 'если подкаталог не найден, то выходим из функции.

    ' Файлы
    If scanrfd.Checked Then SearchForFiles "*.rfd", Path
    If scanrbp.Checked Then SearchForFiles "*.rbp", Path
'    If scanrfd.Checked Then SearchForFiles "*.*", Path
    
    If Check1.value = False Then Exit Sub
    
    Do
        Atr = (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) 'узнаём, является ли найденный файл папкой
        If TrimNull(WFD.cFileName) <> "." And TrimNull(WFD.cFileName) <> ".." Then 'если папка не является корневой на диске, то функция FindNextFile возвращает эти два значения.
            If Atr > 0 Then
                'DoEvents
                SearchForFolders fName, Path & TrimNull(WFD.cFileName) & "\"  'Если файл является папкой, то снова вызываем функцию поиска подкаталогов в этой папке.
            End If

        End If

    Loop While FindNextFile(hFnd, WFD) 'производим поиск до конца.

    FindClose hFnd 'освобождаем память.
End Sub


Private Sub SearchForFiles(fName As String, Path As String) 'Path - указывает в какой папке будет производится поиск фалов, указанных в параметре fName.
    If StopSearch = True Then Exit Sub

    On Error Resume Next

    Dim Atr As Integer
    Dim hFnd As Long
    Dim WFD As WIN32_FIND_DATA
    If Not Right(Path, 1) = "\" Then Path = Path & "\"
    hFnd = FindFirstFile(Path & fName, WFD) 'ищем первый файл.
    If hFnd = INVALID_HANDLE_VALUE Then Exit Sub

    Do
    Atr = (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) 'узнаём, является ли найденный файл папкой
    If TrimNull(WFD.cFileName) <> "." And TrimNull(WFD.cFileName) <> ".." Then
        If Atr > 0 Then 'если это папка

        Else 'если это файл

        Dim strk As String
        Dim FILE As String
        Dim DateCreated As String
        FILE = LCase(Path & TrimNull(WFD.cFileName))
        
'        Me.StatusBar1.Panels(1).Text = FILE
'        Me.StatusBar1.Refresh
  
        ' разбивка по годам
        DateCreated = FileDateTime(FILE)
        strk = Year(DateCreated) & ":"
  
        If Not TreeView1.Nodes(strk).Key = strk Then
            TreeView1.Nodes.Add , tvwChild, strk, Year(DateCreated) ', 1, 1
        End If
  
        '  TreeView1.Nodes.Add strk, tvwChild, strk & "\" & MonthNames(Month(DateCreated) - 1), MonthNames(Month(DateCreated) - 1), 3, 3
        TreeView1.Nodes.Add strk, tvwChild, strk & "." & MonthNames(Month(DateCreated) - 1), MonthNames(Month(DateCreated) - 1) ', 2, 2
        Dim SomeNode As Node
        Dim FileName As String
        FileName = Right(FILE, Len(FILE) - InStrRev(FILE, "\", -1))
        Set SomeNode = TreeView1.Nodes.Add(strk & "." & MonthNames(Month(DateCreated) - 1), tvwChild, FILE, FileName) ', 3, 3)

        SomeNode.Text = FileName & ": " & open_file(FILE)
        
        Me.StatusBar1.Panels(2).Text = ListView1.ListItems.Count

        End If
    
    End If

Loop While FindNextFile(hFnd, WFD)

FindClose hFnd
End Sub


Private Function TrimNull(Start As String) As String
    Dim pos As Integer
        pos = InStr(Start, Chr$(0))
        If pos Then
            TrimNull = Left$(Start, pos - 1)
            Exit Function
        End If

        TrimNull = Start
End Function


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListView1_DblClick
    End If

End Sub


Sub DatSorting(Optional NodeVisible As Boolean)
    Dim anyNode As Node
        For Each anyNode In TreeView1.Nodes
            If anyNode.Children > 0 Then
                '    If anyNode.Parent Is Nothing Then
                '    anyNode.Text = anyNode.Text & ": " & anyNode.Children * anyNode.Child.Children
                '    Else
                anyNode.Text = anyNode.Key & ": " & anyNode.Children
    
                If NodeVisible Then
                    anyNode.Expanded = True
                Else
                    anyNode.Expanded = False
                End If
    
                '    End If
                '    Label8 = "Sorting: " & anyNode.Text
                anyNode.Sorted = True
            End If

        Next

End Sub


Sub FilesFil(d As String)
    ListView1.ListItems.Clear
    TreeView1.Nodes.Clear
    Command1.Enabled = False
    Me.MousePointer = 11
    If scanrbp.Checked Or scanrfd.Checked Then
        SearchForFolders "*", d
    End If
    Command1.Enabled = True
    DatSorting
    Me.MousePointer = 1
End Sub


'Sub fillfiles()
'ProjectsDir = strFolder
'Command1.Enabled = False
'Command2.Enabled = False
'ListView1.MousePointer = 11
'TreeView1.Nodes.Clear
'SearchForFolders "*", strFolder, "*" & ".rfd"
''SearchForFolders "*", strFolder, "*" & ".rbp"
'Command1.Enabled = True
'DatSorting
'Label8 = ""
'ListView1.MousePointer = 0
'End Sub

'Private Sub List(ByVal strPath As String)
'    Dim strName As String, strFile As String
'    strName = Dir(strPath, vbDirectory)
'    List1.Clear
'    If Len(strPath) < 3 Then
'        Dim lngIndex As Long, lngType As Long
'        For lngIndex = 65 To 90
'            lngType = GetDriveType(Chr(lngIndex) & ":\")
'            If lngType = 3 Then List1.AddItem chr$(lngIndex) & ":\"
'        Next
'    Else
'
'        List1.AddItem "[..]"
''        If Right(strPath, 1) <> "\" Then strPath = strPath + "\"
'        Do While strName <> vbNullString
'            If strName <> "." And strName <> ".." Then If (GetAttr(strPath & strName) And vbDirectory) = vbDirectory Then List1.AddItem LCase(strName)
'            strName = Dir
'        Loop
''        strName = Dir(strPath & "*.mp3")
''        While strName <> vbNullString
''            lstMain.Add strName, lngMp3
''            strName = Dir()
''        Wend
'    End If
'
''    If strPath <> "" Then List1.AddItem strPath
'    strFolder = strPath
''    lstMain.Selected = 0
''    lstMain.Update
'End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i As Integer, n As Integer
        On Error Resume Next

        'Node.Text = Node.Key & ": " & Node.Children
        ListView1.ListItems.Clear
        ListView1.MousePointer = 11

        If Node.Children = 0 Then open_file Node.Key: n = 1: GoTo FOUND

    Dim nextcild As Node
        Set nextcild = Node.Child
        n = 0
        For i = 0 To Node.Children

            If Not nextcild.Parent.Parent Is Nothing Then
                ' находится не на вершине узлов
                open_file nextcild.Key ' Left$(Node.Child.Key, InStrRev(Node.Child.Key, "\", -1))
                If Not nextcild Is Nothing Then n = n + 1
            Else
 
    Dim inCounter As Integer
        inCounter = nextcild.Child.FirstSibling.Index
        While inCounter <> nextcild.Child.LastSibling.Index
            open_file TreeView1.Nodes(inCounter).Key
            inCounter = TreeView1.Nodes(inCounter).Next.Index
            n = n + 1
        Wend
    
        If inCounter = nextcild.Child.LastSibling.Index Then
            open_file TreeView1.Nodes(inCounter).Key
            n = n + 1
        End If
    
    End If
 
    Set nextcild = nextcild.Next
Next

FOUND:
ListView1.MousePointer = 1
End Sub

