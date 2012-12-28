VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   5400
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   9735
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Height          =   2370
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "manmat.frx":0000
         Left            =   120
         List            =   "manmat.frx":000D
         TabIndex        =   13
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add"
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Dell"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "Интерфейс установки/удаления редактирования данных."
         Height          =   1335
         Left            =   4800
         TabIndex        =   17
         Top             =   1320
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Описательные параметры материала неиспользуемые в дальнейшем расчете."
         Height          =   1095
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Choice of an edited position:"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   9735
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove"
         Height          =   255
         Left            =   8160
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   255
         Left            =   6480
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7011
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14482428
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Working width"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "General width"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Step"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Overlaping"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Min lenght"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Max lenght"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "General height"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OK"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8493
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Material edit"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Thickness - Coating - Color"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Choice of a material for editing (*.own - editor)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10215
   End
   Begin VB.Menu setfile 
      Caption         =   "Set file CFG"
   End
   Begin VB.Menu BU 
      Caption         =   "Base Utilites"
      Begin VB.Menu Import_cfg 
         Caption         =   "Clear mdb & Import cfg from *.own"
      End
      Begin VB.Menu Import_add 
         Caption         =   "Import from *.own & Add to mdb"
      End
      Begin VB.Menu ClearDB 
         Caption         =   "Clear mdb"
      End
   End
End
Attribute VB_Name = "mat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flagopt As Boolean

Private op As Integer

Private Sub ClearDB_Click()
BASE.Dell_All_Data_in_DB Gl.FileName, "Profils"
BASE.Dell_All_Data_in_DB Gl.FileName, "Thickness"
BASE.Dell_All_Data_in_DB Gl.FileName, "Coating"
BASE.Dell_All_Data_in_DB Gl.FileName, "Color"
mat.ListView1.ListItems.Clear
mat.List1.Clear
Project.Combo1(1).Clear
Project.Combo1(2).Clear
Project.Combo1(3).Clear
End Sub

Private Sub Combo1_Click()
Gl.FileName = Combo1.Text
Module10.get_data
End Sub

Private Sub Combo2_Click()
Dim i As Integer
List1.Clear
For i = 0 To Project.Combo1(Combo2.ListIndex + 1).ListCount - 1
List1.AddItem Project.Combo1(Combo2.ListIndex + 1).List(i)
Next
End Sub

Private Sub Command1_Click()

'' ЗАНЕСЕНИЕ ДАННЫХ В ФАЙЛ КОНФИГУРАЦИИ *.OWN
'If Right(Gl.FileName, 3) = "own" Then
'Dim i As Integer
'Dim strt As String
'Dim tmpstrt As String
'Dim ii As Integer
'Dim FileN As Integer
'
'If Me.ListView1.ListItems.count = 0 Then
'MsgBox ResolveResstring(1461), vbCritical, ResolveResstring(1413)
'Me.Hide
'Exit Sub
'End If
'
'strt = 0 & Len(Gl.Firm_name & Gl.Firm_r) & Gl.Firm_name & Gl.Firm_r
'
'For ii = 0 To 3
'If ii = 0 Then
'    strt = strt & Format(Me.ListView1.ListItems.count - 1, "000")
'    For i = 1 To Me.ListView1.ListItems.count
'    tmpstrt = Me.ListView1.ListItems(i).ListSubItems(1)
'    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(4))
'    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(2))
'    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(3))
'    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(5))
'    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(6))
'    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(7))
''    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(8))
''    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(9))
'
'    strt = strt & chr$(Len(tmpstrt)) & tmpstrt
'    Next i
'    strt = Trim$(strt)
'Else
'    If Project.Combo1(ii).ListCount <> 0 Then
'    strt = strt & Format(Project.Combo1(ii).ListCount - 1, "000")
'    For i = 0 To Project.Combo1(ii).ListCount - 1
'    tmpstrt = Trim$(Project.Combo1(ii).List(i))
'    strt = strt & chr$(Len(tmpstrt)) & tmpstrt
'    Next i
'    Else
'    strt = strt & " -1"
'    End If
'End If
'Next ii
'
'strt = DopOWN.Shifr(RTrim$(strt))
'FileN = FreeFile
''If Not Dir(Gl.FileName) = "" Then Kill Gl.FileName
'Open Gl.FileName For Binary As #FileN
'Put #FileN, , strt
'Close #FileN
'
'Call Module10.GetProfils_Properties(Gl.FileName) ' Получение служебной информации
'
'
'' ЗАНЕСЕНИЕ ДАННЫХ В БАЗУ ДАННЫХ *.MDB
'ElseIf Right(Gl.FileName, 3) = "mdb" Then

'BASE.Dell_All_Data_in_DB Gl.FileName, "Profils"
'BASE.Dell_All_Data_in_DB Gl.FileName, "Thickness"
'BASE.Dell_All_Data_in_DB Gl.FileName, "Coating"
'BASE.Dell_All_Data_in_DB Gl.FileName, "Color"

For i = 1 To ListView1.ListItems.count
    
    
    Dim n As Integer
    Dim datarow()
    ReDim datarow(ListView1.ListItems(i).ListSubItems.count)
    For n = 1 To ListView1.ListItems(i).ListSubItems.count
        datarow(n) = ListView1.ListItems(i).SubItems(n)
    Next n
    
BASE.Add_to_Data_DB Gl.FileName, "Profils", datarow, True

Next i

For ii = 1 To 3
    
    If Project.Combo1(ii).ListCount <> 0 Then
    
    ReDim datarow(Project.Combo1(ii).ListCount - 1)
    For i = 0 To Project.Combo1(ii).ListCount - 1
    datarow(i) = Trim$(Project.Combo1(ii).List(i))
    Next i
    
    If ii = 1 Then BASE.Add_to_Data_DB Gl.FileName, "Thickness", datarow, False
    If ii = 2 Then BASE.Add_to_Data_DB Gl.FileName, "Coating", datarow, False
    If ii = 3 Then BASE.Add_to_Data_DB Gl.FileName, "Color", datarow, False

    End If

Next ii

ReDim datarow(2)

datarow(1) = Gl.Firm_name
datarow(2) = Gl.Firm_r

BASE.Add_to_Data_DB Gl.FileName, "Main", datarow, True, True

'End If

MsgBox (ResolveResstring(1468))
Project.lstprof_Click
'For i = 0 To Project.Combo1(0).ListCount - 1 Step 1
'    If Profil = Project.Combo1(0).list(i) Then Project.Combo1(0).ListIndex = i: Exit For
'Next

Me.Hide
End Sub

Private Sub Command2_Click()
flagopt = True
matedit.Show vbModal, Me
End Sub

Private Sub Command3_Click()
Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
End Sub

Private Sub Command4_Click()
flagopt = False
matedit.Text1 = Me.ListView1.SelectedItem.ListSubItems(1)

matedit.Text2 = Me.ListView1.SelectedItem.ListSubItems(2)
matedit.Text3 = Me.ListView1.SelectedItem.ListSubItems(3)
matedit.Text4 = Me.ListView1.SelectedItem.ListSubItems(8)

matedit.Text5 = Me.ListView1.SelectedItem.ListSubItems(4)
matedit.Text6 = Me.ListView1.SelectedItem.ListSubItems(5)
matedit.Text7 = Me.ListView1.SelectedItem.ListSubItems(6)

matedit.Text8 = Me.ListView1.SelectedItem.ListSubItems(7)

'matedit.Text9 = Left(Me.ListView1.SelectedItem.ListSubItems(8), Len(Me.ListView1.SelectedItem.ListSubItems(8)) - 1)
'If Me.ListView1.SelectedItem.ListSubItems(8) <> "0" Then matedit.Combo1.ListIndex = Asc(Right(Me.ListView1.SelectedItem.ListSubItems(8), 1)) - 65

matedit.Show vbModal, Me
End Sub

Private Sub Command5_Click()
Setd.Show vbModal, matedit
End Sub

Private Sub Command6_Click()
op = 1
List1.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End Sub

Private Sub Command7_Click()
op = 2
List1.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End Sub

Private Sub Command8_Click()
op = 3
If List1.ListIndex <> -1 Then
Project.Combo1(Combo2.ListIndex + 1).RemoveItem List1.ListIndex
List1.RemoveItem List1.ListIndex

Command6.Enabled = True
Command7.Enabled = False
Command8.Enabled = False
End If
End Sub

Private Sub Form_Load()

'Me.Left = GetSetting(App.ProductName, "Position", Me.Name & "left", Me.Left)
'Me.top = GetSetting(App.ProductName, "Position", Me.Name & "top", Me.top)
'Me.Width = GetSetting(App.ProductName, "Position", Me.name & "width", Me.Width)
'Me.Height = GetSetting(App.ProductName, "Position", Me.name & "height", Me.Height)

Me.Caption = ResolveResstring(1085)

TabStrip1.Tabs(1).Caption = ResolveResstring(1086)
TabStrip1.Tabs(2).Caption = ResolveResstring(1087)
'TabStrip1.Tabs(3).Caption = ResolveResstring(1088)
'TabStrip1.Tabs(4).Caption = ResolveResstring(1122)

Label1.Caption = ResolveResstring(1085)

ListView1.ColumnHeaders(2).Text = ResolveResstring(1089)
ListView1.ColumnHeaders(3).Text = ResolveResstring(1090)
ListView1.ColumnHeaders(4).Text = ResolveResstring(1091)
ListView1.ColumnHeaders(5).Text = ResolveResstring(1092)
ListView1.ColumnHeaders(6).Text = ResolveResstring(1093)
ListView1.ColumnHeaders(7).Text = ResolveResstring(1094)
ListView1.ColumnHeaders(8).Text = ResolveResstring(1095)
ListView1.ColumnHeaders(9).Text = ResolveResstring(1096)

'ListView1.ColumnHeaders(9).Text = ResolveResstring(1088)

'Command5.Caption = ResolveResstring(1134)
'Me.cfg.Caption = ResolveResstring(1135)
Me.setfile.Caption = ResolveResstring(1137)
Me.BU.Caption = ResolveResstring(1136)
Me.Import_cfg.Caption = ResolveResstring(1138)
Me.Import_add.Caption = ResolveResstring(1139)
Me.ClearDB.Caption = ResolveResstring(1140)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If fexit = False Then Cancel = -1: Exit Sub
'SaveSetting App.ProductName, "Position", Me.Name & "left", Me.Left
'SaveSetting App.ProductName, "Position", Me.Name & "top", Me.top
'SaveSetting App.ProductName, "Position", Me.name & "width", Me.Width
'SaveSetting App.ProductName, "Position", Me.name & "height", Me.Height
End Sub

Private Sub Import_add_Click()
Import (False)
Module10.get_data
End Sub

Private Sub Import_cfg_Click()
Import (True)
Module10.get_data
End Sub

Private Sub List1_Click()
Text1 = List1.List(List1.ListIndex)
Command6.Enabled = False
Command7.Enabled = True
Command8.Enabled = True
End Sub

Private Sub List1_LostFocus()
'Command6.Enabled = True
'Command7.Enabled = False
'Command8.Enabled = False
End Sub

Private Sub ListView1_DblClick()
Command4_Click
End Sub

Private Sub set_Click()

End Sub

Private Sub setfile_Click()
Gl.FileName = ""
Module10.get_data
End Sub

Sub Import(flag As Boolean)
Dim FileNameCFG As String
Dim FileNameMDB As String

    FileNameCFG = Dialog.GetFileName("", "Roofcalc.own (Roofcalc.own)|Roofcalc.own|*.own (*.own)|*.own|", "", True)

    FileNameMDB = Gl.FileName

If flag = True Then

' clear
'ListView1.ListItems.Clear
'mat.List1.Clear
'Project.Combo1(1).Clear
'Project.Combo1(2).Clear
'Project.Combo1(3).Clear

BASE.Dell_All_Data_in_DB FileNameMDB, "Thickness"
BASE.Dell_All_Data_in_DB FileNameMDB, "Coating"
BASE.Dell_All_Data_in_DB FileNameMDB, "Color"
BASE.Dell_All_Data_in_DB FileNameMDB, "Profils"
End If

'Combo1.AddItem FileName
Module10.Set_DataProfils_to_DB FileNameCFG, FileNameMDB
'End If
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
Case 1
Me.Frame1.Visible = True
Me.Frame2.Visible = False
'Me.Cparts.Visible = False
'Me.Frame3.Visible = False
Case 2
Me.Frame1.Visible = False
Me.Frame2.Visible = True
Combo2.ListIndex = 0
'Me.Cparts.Visible = False
'Me.Frame3.Visible = False
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

If Text1 <> "" Then

Select Case op
Case 1
List1.AddItem Text1
Project.Combo1(Combo2.ListIndex + 1).AddItem Text1

Case 2
If List1.ListIndex <> -1 Then
Project.Combo1(Combo2.ListIndex + 1).List(List1.ListIndex) = Text1
List1.List(List1.ListIndex) = Text1
End If
End Select

List1.Enabled = True
Command6.Enabled = True
Command7.Enabled = False
Command8.Enabled = False
End If
End Sub
