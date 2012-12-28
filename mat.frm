VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material wizard"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9090
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Cparts 
      Caption         =   "Component parts"
      Height          =   3855
      Left            =   1800
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton Command12 
         Caption         =   "Edit"
         Height          =   495
         Left            =   3480
         TabIndex        =   21
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Remove"
         Height          =   495
         Left            =   6960
         TabIndex        =   20
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add"
         Height          =   495
         Left            =   5280
         TabIndex        =   19
         Top             =   3240
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2895
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14482428
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Lenght"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   5470
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   5655
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Установка данных в верхней части страницы"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   8895
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Dell"
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "mat.frx":0000
         Left            =   120
         List            =   "mat.frx":000D
         TabIndex        =   7
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Height          =   2370
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Choice of an edited position:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   8655
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   255
         Left            =   3480
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove"
         Height          =   255
         Left            =   6960
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   255
         Left            =   5280
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5106
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
         NumItems        =   10
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
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Cost"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3570
         Width           =   3255
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Material edit"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Thickness - Coating - Color"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Component parts"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OK"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "V"
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   5280
      Width           =   495
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9000
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9000
      Y1              =   5175
      Y2              =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "CFG:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "mat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public flagopt As Integer
'Private posm As ListItem

Private Sub Combo1_Click()
Dim i As Integer
List1.Clear
For i = 0 To Project.Combo1(Combo1.ListIndex + 1).ListCount - 1
List1.AddItem Project.Combo1(Combo1.ListIndex + 1).list(i)
Next
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim strt As String
Dim tmpstrt As String
Dim ii As Integer
Dim FileN As Integer

If Me.ListView1.ListItems.count = 0 Then
MsgBox ResolveResstring(1461), vbCritical, ResolveResstring(1413)
Me.Hide
Exit Sub
End If

strt = 0 & Len(Gl.Firm_name & Gl.Firm_r) & Gl.Firm_name & Gl.Firm_r

For ii = 0 To 3
If ii = 0 Then
    strt = strt & Format(Me.ListView1.ListItems.count - 1, "000")
    For i = 1 To Me.ListView1.ListItems.count
    tmpstrt = Me.ListView1.ListItems(i).ListSubItems(1)
    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(4))
    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(2))
    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(3))
    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(5))
    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(6))
    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(7))
    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(8))
    tmpstrt = tmpstrt & "," & Trim$(Me.ListView1.ListItems(i).ListSubItems(9))
    strt = strt & chr$(Len(tmpstrt)) & tmpstrt
    Next i
    strt = Trim$(strt)
'    Debug.Print strt
Else
    If Project.Combo1(ii).ListCount <> 0 Then
    strt = strt & Format(Project.Combo1(ii).ListCount - 1, "000")
    For i = 0 To Project.Combo1(ii).ListCount - 1
    tmpstrt = Trim$(Project.Combo1(ii).list(i))
    strt = strt & chr$(Len(tmpstrt)) & tmpstrt
    Next i
    Else
    strt = strt & " -1"
    End If
'    strt = Rtrim$(strt)
'    Debug.Print strt
End If
Next ii
'Debug.Print strt
strt = DopOWN.Shifr(RTrim$(strt))
FileN = FreeFile
If Not Dir(Text2.Text) = "" Then Kill Text2.Text
Open Text2.Text For Binary As #FileN
'  If LOF(#FileN) = 0 Then
'    Select Case language
'      Case "RU"
'        MsgBox ("Файл ROOFCALC.OWN неверен или испорчен.")
'      Case "En"
'        MsgBox ("File ROOFCALC.OWN is corrupted")
'      Case "Fi"
'        MsgBox ("#FileN ROOFCALC.OWN дr fдlaktig")
'    End Select
'   Exit Sub
'  End If
'Debug.Print strt
Put #FileN, , strt
Close #FileN

Call OfficeStart.GetProfils_Properties ' Получение служебной информации
MsgBox (ResolveResstring(1468))
Project.lstprof_Click
For i = 0 To Project.Combo1(0).ListCount - 1 Step 1
    If Profil = Project.Combo1(0).list(i) Then Project.Combo1(0).ListIndex = i: Exit For
Next
Me.Hide
End Sub

Private Sub Command2_Click()
flagopt = 1
matedit.Show vbModal, Me
End Sub

Private Sub Command3_Click()
Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
End Sub

Private Sub Command4_Click()
flagopt = 0
matedit.Text1 = Me.ListView1.SelectedItem.ListSubItems(1)

matedit.Text2 = Me.ListView1.SelectedItem.ListSubItems(2)
matedit.Text3 = Me.ListView1.SelectedItem.ListSubItems(3)
matedit.Text4 = Me.ListView1.SelectedItem.ListSubItems(8)

matedit.Text5 = Me.ListView1.SelectedItem.ListSubItems(4)
matedit.Text6 = Me.ListView1.SelectedItem.ListSubItems(5)
matedit.Text7 = Me.ListView1.SelectedItem.ListSubItems(6)
matedit.Text8 = Me.ListView1.SelectedItem.ListSubItems(7)

matedit.Text9 = Left(Me.ListView1.SelectedItem.ListSubItems(9), Len(Me.ListView1.SelectedItem.ListSubItems(9)) - 1)
If Me.ListView1.SelectedItem.ListSubItems(9) <> "0" Then matedit.Combo1.ListIndex = Asc(Right(Me.ListView1.SelectedItem.ListSubItems(9), 1)) - 65
matedit.Show vbModal, Me
End Sub

Private Sub Command5_Click()
Setd.Show vbModal, matedit
End Sub

Private Sub Command6_Click()
If Text1 <> "" Then
List1.AddItem Text1
Project.Combo1(Combo1.ListIndex + 1).AddItem Text1
End If
End Sub

Private Sub Command7_Click()
If List1.ListIndex <> -1 Then
Project.Combo1(Combo1.ListIndex + 1).list(List1.ListIndex) = Text1
List1.list(List1.ListIndex) = Text1
End If
End Sub

Private Sub Command8_Click()
If List1.ListIndex <> -1 Then
Project.Combo1(Combo1.ListIndex + 1).RemoveItem List1.ListIndex
List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub Command9_Click()
'Dim FileName As String
With OfficeStart.CommonDialog1
            .DialogTitle = "Open"
            .CancelError = False
            .Filter = "Roofcalc.own (Roofcalc.own)|Roofcalc.own|*.own (*.own)|*.own|"
            .ShowOpen
            If Len(.FileName) = 0 Then
                Unload OfficeStart
            End If
            Gl.FileName = .FileName
        End With
If Gl.FileName <> "" Then
Text2 = Gl.FileName
Call OfficeStart.GetProfils_Properties
End If
End Sub

Private Sub Form_Load()
Me.Caption = ResolveResstring(1085)

TabStrip1.Tabs(1).Caption = ResolveResstring(1086)
TabStrip1.Tabs(2).Caption = ResolveResstring(1087)
TabStrip1.Tabs(3).Caption = ResolveResstring(1088)

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

Label3.Caption = ResolveResstring(1098)
Label2.Caption = ResolveResstring(1097)

If Gl.FileName = "" Then Me.Show: Exit Sub
Text2 = Gl.FileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
If fexit = False Then Cancel = -1: Me.Hide
End Sub


Private Sub List1_Click()
Text1 = List1.list(List1.ListIndex)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1
'columselect = ColumnHeader.Index
If ListView1.SortOrder = 0 Then
ListView1.SortOrder = 1
Else
ListView1.SortOrder = 0
End If
ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
Command4_Click
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
Case 1
Me.Frame1.Visible = True
Me.Frame2.Visible = False
Me.Cparts.Visible = False
Case 2
Me.Frame1.Visible = False
Me.Frame2.Visible = True
Combo1.ListIndex = 0
Me.Cparts.Visible = False
Case 3
Me.Cparts.Visible = True
Me.Frame1.Visible = False
Me.Frame2.Visible = False
End Select
End Sub
