VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A490304E-1BC4-4F6C-97AC-D2383955DD55}#2.0#0"; "SplitterModern.ocx"
Begin VB.Form cust 
   Caption         =   "Customers` manager"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   10665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   5280
      TabIndex        =   5
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton Command8 
         Caption         =   "New Cusnomer"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Height          =   735
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Add or Edit Customer"
         Height          =   495
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Remove Customer"
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Tel / FAx:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Description:"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Customer`s code:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin SplitterHV.SplitHV SplitHV1 
      Height          =   7215
      Left            =   5160
      TabIndex        =   4
      Top             =   0
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   12726
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   10821
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14482428
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tel / Fax"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Files"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label16 
      Caption         =   "Search of the customer:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customers` manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "cust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const FILES = 100
Const CUST = 200
Dim Count_Cust As Integer
Private Type Customer
    Date_C As Date
    Name_C As String
    Tel As String
    Descrp As String
    Count_F As Integer
    FILES(FILES) As String
End Type
Dim Customers(CUST) As Customer

Private Sub Combo1_Change()
Combo1_Click
End Sub

Private Sub Combo1_Click()
Dim FILE As String
FILE = Combo1
If Combo1 <> "" Then
FILE = open_file(FILE)
If Frame1.Visible = False Then
Label15 = "Файл [" & Combo1 & "] пустой или содержит ошибку."
End If
If FILE <> "" Then
Frame1.Visible = True
Command7.Enabled = True
Else
Frame1.Visible = False
Command7.Enabled = False
End If
'If Combo1 <> "" And FILE <> "" Then
Command6.Enabled = True
'Else
'Command6.Enabled = False
'End If
Else
Command7.Enabled = False
Frame1.Visible = False
Label15 = "На данного клиента расчеты не проводились."
End If
End Sub

Private Sub Command1_Click()
'If Me.ListView1.ListItems.Count = 0 Then

'End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
'newproject.Show vbModal, Me
Me.Hide
OfficeStart.menu_uusi_Click
Me.Combo1.AddItem Gl.Katalog_files & FileName_data & Gl.FileName_extension
Combo1 = Combo1.List(Combo1.ListCount - 1)
Project.Text3 = Text2
Label14 = Combo1.ListCount
Command4_Click
Unload Me
'Combo1.Text = Combo1.list(Combo1.ListCount - 1)
End Sub

Function open_file(FILE As String)
Dim str As String
Dim strn As Single
Dim str1 As String
Dim leninfo As Integer
'Dim itmX As ListItem
Dim i As Integer
Dim FileN As Integer

On Error GoTo ERR
FileN = FreeFile
Open FILE For Binary As #FileN
  If LOF(FileN) = 0 Then
    Close #FileN
    Frame1.Visible = False
    Exit Function
  End If
  str = string$(8, " ")
  Get #FileN, , str
  If Right(str, 4) = chr$(2) & chr$(0) & "N" & chr$(9) Then
  Setup.Check2.value = 1
  Get #FileN, , strn
    Text4 = Input$(strn, FileN)
  Else
  Setup.Check2.value = 0
  Seek FileN, 1
  Get #FileN, , strn
    Text4 = Input$(strn, FileN)
  End If
  Get #FileN, , strn
    Text5 = Input$(strn, FileN)
  Get #FileN, , strn
  str = Trim(Input$(strn, 1))
  Get #FileN, , strn
  str = str & "- " & Input$(strn, FileN)
  Get #FileN, , strn
  str = str & "- " & Input$(strn, FileN)
  Get #FileN, , strn
  str = str & "- " & Input$(strn, FileN)
Close #FileN
Label1 = str
Me.Combo1.Text = FILE
Frame1.Visible = True
open_file = FILE
Exit Function
ERR:
Close #FileN
Label15 = "Файл [" & FILE & "] пустой или содержит ошибку."
'MsgBox "File is " & FILE & " corrupted."
open_file = ""
Exit Function
End Function

Private Sub Command4_Click()
Dim itmX As ListItem
Dim i As Integer
If Me.Text6 <= CUST Then
Customers(Text6).Count_F = Me.Combo1.ListCount - 1
If ListView1.ListItems.count < Me.Text6 Then
Set itmX = ListView1.ListItems.Add(, , Text6)
itmX.SubItems(1) = Text2
Customers(Text6).Name_C = Text2
itmX.SubItems(2) = Text3
Customers(Text6).Tel = Text3
itmX.SubItems(3) = Text1
Customers(Text6).Descrp = Text1
itmX.SubItems(4) = Label11
Customers(Text6).Date_C = Label11
'ReDim Customers(Text6).FILES(Me.Combo1.ListCount)
For i = 0 To Me.Combo1.ListCount - 1
Customers(Text6).FILES(i) = Me.Combo1.List(i)
OfficeStart.StatusBar.Panels(2) = "The file [" & Combo1.List(i) & "] of the project is added to the customer " & Text2
Next
Else
ListView1.SelectedItem.SubItems(1) = Text2
Customers(Text6).Name_C = Text2
ListView1.SelectedItem.SubItems(2) = Text3
Customers(Text6).Tel = Text3
ListView1.SelectedItem.SubItems(3) = Text1
Customers(Text6).Descrp = Text1
ListView1.SelectedItem.SubItems(4) = Label11
Customers(Text6).Date_C = Label11
'Customers(Text6)
'ReDim Customers(Text6).FILES(Me.Combo1.ListCount)
For i = 0 To Me.Combo1.ListCount - 1
Customers(Text6).FILES(i) = LCase(Me.Combo1.List(i))
OfficeStart.StatusBar.Panels(2) = "The file [" & Combo1.List(i) & "] of the project is added to the customer " & Text2
Next
End If
'itmX.Selected = True
'ListView1_ItemClick itmX
Else
Select Case Gl.language
Case "En"
 MsgBox "The limit [" & CUST & "] of amount of allowable recordings is exceeded.", vbCritical, "System error"
Case "RU"
 MsgBox "Превышен предел [" & CUST & "] количества допустимых записей.", vbCritical, "System error"
End Select
End If
'itmX = Text6
'Me.ListView1.ListItems.Item(itmX).Selected = True
End Sub

Private Sub Command5_Click()
Dim i As Integer
If ListView1.ListItems.count <> 0 Then
If Text6 = Me.ListView1.SelectedItem.Index Then
Text6 = Me.ListView1.SelectedItem.Index
DELL:
Customers(Text6).Count_F = 0
Customers(Text6).Date_C = Empty
Customers(Text6).Descrp = ""
Customers(Text6).Name_C = ""
Customers(Text6).Tel = ""
Me.ListView1.ListItems.Remove CInt(Text6)
For i = 1 To Me.Combo1.ListCount - 1
Customers(Me.ListView1.SelectedItem.Index).FILES(i) = ""
OfficeStart.StatusBar.Panels(2) = "The file [" & Combo1.List(i) & "] of the project is removed at the customer " & Text2
Next
Else
Text6 = Me.ListView1.ListItems.count
GoTo DELL
End If
Else
Combo1.Enabled = False
Command3.Enabled = False
End If
End Sub

Private Sub Command6_Click()
Dim i As Integer
For i = 0 To Combo1.ListCount - 1
If Combo1.List(i) = Combo1 Then Exit For
Next
If i <> Combo1.ListCount Then
Combo1.RemoveItem i
Label14 = Combo1.ListCount                    '
Combo1 = ""
End If
Command4_Click
exit_with_save
'Unload Me
End Sub

Private Sub Command7_Click()
Dim ans As Integer
If Gl.FileName_data = "" Then
OP:
ans = Module10.Close_Project
If ans = 2 Then Exit Sub
Gl.FileName_data = Combo1

Gl.FileName_data = Right(Combo1, Len(Combo1) - InStrRev(Combo1, "\", -1))
Gl.Katalog_files = Left(Combo1, InStrRev(Combo1, "\", -1))
OfficeStart.StatusBar.Panels(2) = Gl.Katalog_files & Gl.FileName_data

     Gl.FileName_data = Module10.Load_f

If Gl.FileName_data = "" Then Exit Sub
Project.Combo1(0) = Gl.profil
Project.Combo1(1) = Gl.Tolshina
Project.Combo1(2) = Gl.Pokritie
Project.Combo1(3) = Gl.ColorRoof
Project.Label11.Caption = Gl.Katalog_files
Project.Label14.Caption = Gl.FileName_extension
'Me.Hide

exit_with_save
Unload Me
Project.Show
Else
If Right(Combo1, Len(Gl.FileName_data)) = Gl.FileName_data Then
exit_with_save
Unload Me
Exit Sub
End If
Select Case language
Case "RU"
ans = MsgBox("Нельзя открыть проект пока открыт файл" & SC & FileName_data & Gl.FileName_extension & SC & "Вы желаете закрыть его", vbYesNo, "System question")
Case "En"
ans = MsgBox("The project doesn`t be open while the file is open" & SC & FileName_data & Gl.FileName_extension & SC & "Do you wish to close it?", vbYesNo, "System question")
End Select
If ans = 7 Then
Exit Sub
Else
GoTo OP
End If
End If
End Sub

Private Sub Command8_Click()
'If Text6 < Me.ListView1.ListItems.Count Then
Text6 = Me.ListView1.ListItems.count + 1
'Else
'Text6 = Text6 + 1
'End If
Command4_Click
End Sub

Private Sub exit_with_save()
Dim MyRecord As Customer
Dim FileN As Integer
Dim i As Integer
FileN = FreeFile
'Kill Gl.Katalog & "\data\customers.dat"
Open Gl.Katalog & "\data\customers.dat" For Random As #FileN Len = Len(MyRecord)
For i = 1 To ListView1.ListItems.count
Put #FileN, , Customers(i)
Next
Close FileN
If ListView1.ListItems.count = 0 Then Kill Gl.Katalog & "\data\customers.dat"
'Me.Hide
'Unload Me
End Sub

Private Sub Command9_Click()
exit_with_save
Unload Me
End Sub

Private Sub Form_Load()
'Me.ListView1.ListItems.Clear
'If Me.load <> 0 Then Exit Sub
'Label15=""
End Sub

Private Sub Form_Unload(Cancel As Integer)
exit_with_save
End Sub

Private Sub Label14_Change()
Combo1_Click
End Sub

'Private Sub Form_Activate()
'Dim itmX As ListItem
'itmX = Text6
'ListView1_ItemClick itmX
'Text2 = Project.Text3
'End Sub

Public Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Text2 = Item.ListSubItems(1)
Text3 = Item.ListSubItems(2)
Text1 = Item.ListSubItems(3)
Label11 = Item.SubItems(4)
Combo1.Enabled = True
Command3.Enabled = True
Combo1.clear
For i = 0 To Customers(Item).Count_F
If Not Customers(Item).FILES(i) = "" Then Combo1.AddItem LCase(Customers(Item).FILES(i))
Next
Label14 = Combo1.ListCount
Combo1.Text = Combo1.List(0)
Text6 = Item
'If Combo1 <> "" Then
'Frame1.Visible = True
'Else
'Frame1.Visible = False
'Label15 = "The customer [" & Text2 & "] has no projects."
'End If
End Sub

Function load()
Dim MyRecord As Customer
Dim FileN As Integer
Dim recn As Integer
Dim i As Integer
FileN = FreeFile
On Error GoTo ERR
'recn = 0
ProgressBar1.Max = CUST
Open Gl.Katalog & "\data\customers.dat" For Random As #FileN Len = Len(MyRecord)
Do While Not EOF(FileN)
recn = recn + 1
ProgressBar1.value = recn
Get #FileN, recn, Customers(recn)
Loop
Close FileN
'If Not Customers(recn).Count_F = 0 Then

For i = 1 To recn

If Customers(i).Name_C <> "" Then
Set itmX = ListView1.ListItems.Add(, , i)
itmX.SubItems(1) = Customers(i).Name_C
itmX.SubItems(2) = Customers(i).Tel
itmX.SubItems(3) = Customers(i).Descrp
itmX.SubItems(4) = Customers(i).Date_C
End If

Next
'End If
load = -1
ProgressBar1.value = 0
Exit Function
ERR:
Close FileN
load = 0
End Function

'Private Sub Text6_Change()
'If Text6 = "" Then Text6 = 0
'If Me.ListView1.ListItems.Count >= Text6 And Text6 <> 0 Then
''Me.ListView1.ListItems(Text6).Selected = True
'End If
'End Sub

Private Sub Text7_Change()
Dim itmX As ListItem
'ListView1.FindItem(Text3, 1, 1, 0)
'ListView1.FindItem(Text1, 1, , 1)
'If Combo1.ListIndex = -1 Then
'MsgBox "Please choose a method of search"
'Combo1.SetFocus
'Exit Sub
'End If
Set itmX = ListView1.FindItem(CStr(Text7), 1, , 1)
If itmX Is Nothing Then  ' If no match, inform user and exit.
      Exit Sub
   Else
       Set itemselect = itmX
       itmX.EnsureVisible ' Scroll ListView to show found ListItem.
       itmX.Selected = True   ' Select the ListItem.
'       ListView1.SetFocus
   End If
End Sub
