VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Select_detail_item 
   Caption         =   "Select line"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select line for calc"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Adjacent"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Lenght"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   6068
      EndProperty
   End
End
Attribute VB_Name = "Select_detail_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selectitem As Integer



Private Sub Command1_Click()
Dim calc As Single
Dim i As Integer
'If ListView1.SelectedItem Is Nothing Then Exit Sub
   For i = 1 To ListView1.ListItems.count
If DetailItems(Gl.N_Slope, i, 0) = 0 Then
      If ListView1.ListItems(i).Checked = True Then
        DetailItems(Gl.N_Slope, i, 0) = Lapepic.ListView1.SelectedItem
        calc = Lapepic.ListView1.SelectedItem.ListSubItems(2)
        SizeItem(Lapepic.ListView1.SelectedItem) = SizeItem(Lapepic.ListView1.SelectedItem) + ListView1.ListItems(i).ListSubItems(2)
        calc = calc + ListView1.ListItems(i).ListSubItems(2)
        Lapepic.ListView1.SelectedItem.ListSubItems(2) = calc
        DetailPoints(N_Slope) = DetailPoints(N_Slope) + 1
      End If
End If
   Next i
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
Me.top = OfficeStart.top + 1210
Me.Left = Lapepic.Width - 1760
Me.ListView1.ToolTipText = "Для выбора смежной линии кликните дважды или выберите желаемый скат в нижнем списке."
'Me.Caption = Me.Caption & " (" & Lapepic.ListView1.SelectedItem.ListSubItems(2) & ")"
'Lapepic.Filling
End Sub

Private Sub Form_Resize()
Command1.Width = Me.ScaleWidth
ListView1.Width = Me.ScaleWidth
If Me.ScaleHeight > (Command1.Height + 100) Then ListView1.Height = Me.ScaleHeight - Command1.Height
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Lapepic.setdetitems = Right(Item, Len(Item) - 1)
Lapepic.Draw_Point Lapepic
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
