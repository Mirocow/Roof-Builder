VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form ChangeProfil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin СhameleonButton.chameleonButton chameleonButton1 
      Height          =   615
      Left            =   5760
      TabIndex        =   15
      Top             =   5760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      BTYPE           =   7
      TX              =   "&OK"
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
      MICON           =   "ChangeProfil.frx":0000
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
   Begin VB.Frame Frame3 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame4 
         Height          =   4935
         Left            =   120
         TabIndex        =   4
         Top             =   630
         Width           =   8415
         Begin VB.ComboBox lstprof 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   4935
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   350
            Left            =   120
            TabIndex        =   7
            Top             =   650
            Width           =   8175
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   75
               Width           =   7935
            End
         End
         Begin VB.ComboBox ComboBox1 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   8175
         End
         Begin СhameleonButton.chameleonButton Command2 
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   4320
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   873
            BTYPE           =   7
            TX              =   "chameleonButton2"
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
            MICON           =   "ChangeProfil.frx":001C
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
            Height          =   2535
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Properties name"
               Object.Width           =   5645
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   1411
            EndProperty
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Caption         =   "Выбор профиля"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   8175
         End
         Begin VB.Label Label23 
            Caption         =   "Сортировка по группам:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6375
         Begin VB.Label Label24 
            BackColor       =   &H00808080&
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   75
            Width           =   6255
         End
      End
      Begin СhameleonButton.chameleonButton Command8 
         Height          =   375
         Left            =   6600
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   7
         TX              =   "chameleonButton3"
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
         MICON           =   "ChangeProfil.frx":0038
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
      Begin VB.Label Label6 
         Caption         =   "Метка2"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Метка1"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   3240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "ChangeProfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public objComboBox As Collection


Public Sub AddItemComboBox(Data() As String)
Dim Item As ItemComboBox
Set Item = New ItemComboBox
Item.AddColumns Data
objComboBox.Add Item, CStr(objComboBox.Count)
ComboBox1.AddItem Data(0)
End Sub


Public Sub ClearComboBox()
ComboBox1.Clear
Set objComboBox = New Collection
End Sub


Private Sub chameleonButton1_Click()
Me.Hide
End Sub

Private Sub ComboBox1_Change()
If ComboBox1.Text = "" Then ListView1.ListItems.Clear
End Sub


Private Sub ProfilFillData()
        
    Frame3.Visible = True
    
    ListView1.ListItems.Clear
    
    Dim itmXlist As ListItem
    Dim PDataRS As Recordset
    Dim selectid As Integer
    Dim FactoryID As Integer
    
'    If Label24.Caption = "NOGROUP" Or Label24.Caption = "*" Then Factory_Name = ""

        If ComboBox1.ListIndex > -1 Then
            Profil_Name = ""
            Factory_Name = ""
            FactoryID = ItemComboBox(ComboBox1.ListIndex, 3)
            Label2.Caption = GetGroupName(ItemComboBox(ComboBox1.ListIndex, 2))
        Else
            If Factory_Name <> "" Then FactoryID = GetFactoryID(Factory_Name) Else FactoryID = 0
        End If
        
        Set PDataRS = GetProfilData(ComboBox1.Text, FactoryID)
    
        If Not PDataRS Is Nothing Then
        
            Do While Not PDataRS.EOF
            
                selectid = CheckNullNomber(PDataRS.Fields(1))
                
                Set itmXlist = ListView1.ListItems.Add(, , "ID")
                itmXlist.SubItems(1) = selectid
                
                Profil_Name = GetProfilName(selectid)
                
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1503))
                itmXlist.SubItems(1) = FactoryID 'GetFactoryName(GetFactoryIDFromProfilID(selectid))
                
                Label24.Caption = GetFactoryName(itmXlist.SubItems(1))   'itmXlist.SubItems(1)
                Label24.Tag = FactoryID
                
                Factory_Name = Label24.Caption
                
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1034))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(2)))
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1035))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(3)))
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1036))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(4)))
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1055))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(5)))
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1056))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(6)))
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1057))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(7)))
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1079))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(8)))
            
                ' Интерфейс невыполнимых длин
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1135))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(9)))
                Set itmXlist = ListView1.ListItems.Add(, , lng.GetResIDstring(1136))
                itmXlist.SubItems(1) = ConvertData(CheckNullNomber(PDataRS.Fields(10)))
            
                PDataRS.MoveNext
            Loop

            PDataRS.Close
            Set PDataRS = Nothing
        Else
            ListView1.ListItems.Clear
        End If

End Sub


Public Function ItemComboBox(lItem As Integer, Col As Integer)
If objComboBox.Count = 0 Then Exit Function
Dim Item As ItemComboBox
Set Item = New ItemComboBox
Set Item = objComboBox.Item(lItem + 1)
ItemComboBox = Item.Column(Col)
Set Item = Nothing
End Function


Public Sub ComboBox1_Click()
            ListView1.ListItems.Clear
    
            '
            ' Если расчет не производился
            '
            If ComboBox1.Text <> "" And ((Profil_Name = ComboBox1.Text) Or KolvoScatov = 0) Then
            
                If Profil_Name <> "" Or KolvoScatov = 0 Then

                    ' Заполнение харрактеристик профиля и вывод в listwiev
                    ProfilFillData
                    
                    If ListView1.ListItems.Count > 2 Then
                        
                        If KolvoScatov <> 0 Then
                            ' Заполнение массива имен профилей
                            loadProfil Profil_Name
                            ' Заполнение массива имен производителей
                            loadFactory Factory_Name
                        End If
                        
                    Else
                        GoTo ERR ' Если такого профиля нет в БД
                    End If

                Else
                    GoTo ERR
                End If

            Else
        
                If ComboBox1.Text <> "" Then
                
                    'ListIndex
                
                    ' Заполнение харрактеристик профиля и вывод в listwiev
                    ProfilFillData
                    
                Else
                    ' В файле проекта содержится название профиля
                    ' но данные по нему отсутствуют в бд
                    If Profil_Name <> "" Then GoTo ERR
                    ComboBox1.Text = ""
                    Profil_Name = ""
                End If
                Exit Sub

ERR:

                ComboBox1.Text = ""
                MsgBox lng.GetResIDstring(1437, "%PROFIL%", Profil_Name), vbCritical, App.ProductName
                Profil_Name = ""
                Exit Sub
                
            End If
End Sub


Private Sub Command2_Click()
Dim bRes As Boolean
Dim ismatedit As Boolean
On Error GoTo ERR

If IsLic = False Then GoTo ERR

If IsAdmin = False Then GoTo ERR

Dim pwd As String, check As String
Dim a As New clsRegistry
check = a.GetStringValue(HKEY_LOCAL_MACHINE, "Software\" & App.ProductName, "adm", 0)
    
If check <> "" Then
pwd = InputBox(lng.GetResIDstring(9707), "Please input password", "")
If RC4(StrConv("PASSWORD", vbFromUnicode), pwd) <> check Then
    Set a = Nothing
    If pwd <> "" Then MsgBox lng.GetResIDstring(9708), vbInformation
    Exit Sub
End If
End If

Set a = Nothing


Screen.MousePointer = 11
Command2.Enabled = False
Me.Enabled = False
OfficeStart.Enabled = False
        
    Load matedit
    matedit.Label13 = ComboBox1.Text

    If matedit.Label13 <> "" Then
    
        If ComboBox1.ListIndex > -1 Then
            matedit.Combo1.ListIndex = ItemComboBox(ComboBox1.ListIndex, 2)
        Else
            matedit.Combo1.ListIndex = 0
        End If
    
        matedit.Text2 = ListView1.ListItems(CInt(iData(3))).ListSubItems(1).Text
        matedit.Text3 = ListView1.ListItems(CInt(iData(4))).ListSubItems(1).Text
        matedit.Text4 = ListView1.ListItems(CInt(iData(5))).ListSubItems(1).Text
        matedit.Text5 = ListView1.ListItems(CInt(iData(6))).ListSubItems(1).Text
        matedit.Text6 = ListView1.ListItems(CInt(iData(7))).ListSubItems(1).Text
        matedit.Text7 = ListView1.ListItems(CInt(iData(8))).ListSubItems(1).Text
        matedit.text8 = ListView1.ListItems(CInt(iData(9))).ListSubItems(1).Text
        
'        If matedit.Text6 = 0 Then matedit.Text9.Enabled = False: matedit.Text10.Enabled = False Else _
'        matedit.Text9.Enabled = True
'        matedit.Text10.Enabled = True
        
        matedit.Text9 = ListView1.ListItems(CInt(iData(10))).ListSubItems(1).Text
        matedit.Text10 = ListView1.ListItems(CInt(iData(11))).ListSubItems(1).Text
        
        matedit.Check2.Enabled = True
        
    Else
        matedit.Check2.Enabled = False
        matedit.Label13 = "New"
    End If

    matedit.Label9 = Label24

    Connect Gl.FileName, -Val(iData(0)) ' Val(iData(0)) = FALSE
    Screen.MousePointer = 0
    matedit.Show vbModal, OfficeStart
    Connect Gl.FileName, True

    Screen.MousePointer = 0
    Command2.Enabled = True
    OfficeStart.Enabled = True
    Me.Enabled = True
    Exit Sub
    
ERR:
    Connect Gl.FileName, True
    Screen.MousePointer = 0
    Module10.withoutl
    Command2.Enabled = True
    Me.Enabled = True
    OfficeStart.Enabled = True
End Sub


Private Sub Command8_Click()
    Dim CP As POINTAPI
    ClientToScreen Command8.hwnd, CP
    FactoryNames.Left = (CP.X * Screen.TwipsPerPixelY) - FactoryNames.Width + Command8.Width
    FactoryNames.Top = CP.Y * Screen.TwipsPerPixelY + Command8.Height
    FactoryNames.Show vbModal, OfficeStart
    'lstprof_Click
End Sub


Private Sub Form_Load()

    Set objComboBox = New Collection

    Label5.Caption = lng.GetResIDstring(1030)
    Label6.Caption = lng.GetResIDstring(1031)
    Command2.Caption = lng.GetResIDstring(9389)
    Label21.Caption = lng.GetResIDstring(9395)
    Command8.Caption = lng.GetResIDstring(9397)
    Label23.Caption = lng.GetResIDstring(9396)
    
    Connect Gl.FileName, True

    ' По умолчанию
    Dim RS As Recordset
    Set RS = RequestSQL("select ID from FirmFactory where name = '" & Factory_Name & "'")
    If Not RS Is Nothing Then
        Label24.Caption = Factory_Name
        Label24.Tag = RS!id
        Command2.Enabled = True
    Else
        Command2.Enabled = False
        Label24.Tag = 0
        Label24.Caption = "*"
    End If
    Set RS = Nothing

    Set RS = RequestSQL("select GroupName.id, GroupName.Name from GroupName where LNG='" & lng.GetResIDstring(100) & "' order by GroupName.id")
    If Not RS Is Nothing Then
        Do While Not RS.EOF
            lstprof.AddItem RS.Fields(1)
            RS.MoveNext
        Loop
        RS.Close
        lstprof.ListIndex = 0
    End If
    Set RS = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objComboBox = Nothing
End Sub

Sub lstprof_Click()
    Dim PNamesRS As Recordset
    Dim sql As String
    Dim SQLF As String
    On Error Resume Next
    
    
    Project.Combo1(3).Clear
    Project.Combo1(1).Clear
    Project.Combo1(2).Clear
    
    ListView1.ListItems.Clear
    ClearComboBox
    
    If Factory_Name <> "" And Factory_Name <> "NOGROUP" Then
        Dim FactoryID As Integer
        FactoryID = GetFactoryID(Factory_Name)
        SQLF = "p.IDFACTORY=" & FactoryID
        Label24.Caption = Factory_Name
        Label24.Tag = FactoryID
    Else
        SQLF = "p.IDFACTORY=0"
        Label24.Caption = "*"
        Label24.Tag = 0
    End If
 
    If lstprof.ListIndex <> 0 Then
        If SQLF <> "" Then SQLF = "and " & SQLF
        ' Вывод в соответствии с категорией
        sql = "select p.id, p.name, p.idfactory, p.idgroup from ProfiName p where p.IDGROUP=" & _
        lstprof.ListIndex & " " & SQLF & " order by p.id"
    Else
        If SQLF <> "" Then SQLF = "where " & SQLF
        ' Вывод без категорий
        sql = "select p.id, p.name, p.idfactory, p.idgroup from ProfiName p " & SQLF & " order by p.id"
    End If

    ' Вывод имен профилей с сортировкой по id
    Set PNamesRS = RequestSQL(sql)

    If Not PNamesRS Is Nothing Then
        
        ' забивка наименований профилей
        Dim Recitem As ListItem
        Dim i As Integer
        Do While Not PNamesRS.EOF
            
            Dim Data(3) As String
            Data(0) = CheckNull(PNamesRS.Fields(1)) ' NAME
            Data(1) = CheckNull(PNamesRS.Fields(0)) ' ID
            Data(2) = CheckNull(PNamesRS.Fields(3)) ' IDGROUP
            Data(3) = CheckNull(PNamesRS.Fields(2)) ' IDFACTORY
            AddItemComboBox Data
            
            PNamesRS.MoveNext
        Loop

        lstprof.Enabled = True
        ChangeProfil.ComboBox1.Enabled = True
        PNamesRS.Close
        
    Else
        ChangeProfil.ComboBox1.Enabled = False
    End If
    
    Label2.Caption = lstprof.Text

End Sub


