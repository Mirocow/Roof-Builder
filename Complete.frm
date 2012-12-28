VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Complete 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CODE"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NAME"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "LENGTH"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DEVIDE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "OVERCLOAK"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "TYPECALC"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   " TYPE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "DESCRIPTION"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Description:"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   6975
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Ãðóïïû ýëèìåíòîâ:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   3855
      Left            =   60
      Top             =   60
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3855
      Left            =   45
      Top             =   45
      Width           =   7095
   End
End
Attribute VB_Name = "Complete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Dim RSC As Recordset
Dim itmX As ListItem

Set RSC = MainBaseFunction.RequestSQL("select c.ID, c.CODE, c.NAME, c.LENGTH, c.DEVIDE, c.OVERCLOAK, c.TYPECALC, c.TYPE, c.DESCRIPTION  from Completes c where c.TYPE=" & Combo1.ListIndex & " order by c.id")

ListView1.ListItems.Clear

If Not RSC Is Nothing Then
Do While Not RSC.EOF
    Set itmX = ListView1.ListItems.Add(, , CheckNull(RSC.Fields(0)))
    itmX.SubItems(1) = CheckNull(RSC.Fields(1))
    itmX.SubItems(2) = CheckNull(RSC.Fields(2))
    itmX.SubItems(3) = CheckNull(RSC.Fields(3))
    itmX.SubItems(4) = CheckNull(RSC.Fields(4))
    itmX.SubItems(5) = CheckNull(RSC.Fields(5))
    itmX.SubItems(6) = CheckNull(RSC.Fields(6))
    itmX.SubItems(7) = CheckNull(RSC.Fields(7))
    itmX.SubItems(8) = CheckNull(RSC.Fields(8))
'    itmX.SubItems(9) = CheckNull(RSC.Fields(9))
    RSC.MoveNext
Loop
RSC.Close
End If
End Sub

Private Sub Form_Load()
Dim RSC As Recordset
Set RSC = MainBaseFunction.RequestSQL("select c.ID, c.NAME from CompleteType c  order by c.ID")

If Not RSC Is Nothing Then
Do While Not RSC.EOF
    Combo1.AddItem RSC.Fields(1)
    RSC.MoveNext
Loop
RSC.Close
End If

Combo1.ListIndex = 1
End Sub

Private Sub ListView1_DblClick()
Project.MSFlexGrid1.Rows = Project.MSFlexGrid1.Rows + 1
Project.MSFlexGrid1.TextMatrix(Project.MSFlexGrid1.Rows - 1, 0) = ListView1.SelectedItem
Project.MSFlexGrid1.TextMatrix(Project.MSFlexGrid1.Rows - 1, 1) = ListView1.SelectedItem.SubItems(2)
Me.Hide
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Text1 = ""
Text1 = Text1 & Item.ListSubItems(1) & vbCrLf
Text1 = Text1 & Item.ListSubItems(3) & vbCrLf
Text1 = Text1 & Item.ListSubItems(4) & vbCrLf
Text1 = Text1 & Item.ListSubItems(5) & vbCrLf
'Text1 = Text1 & Item.ListSubItems(7) & vbCrLf
Text1 = Text1 & Item.ListSubItems(8) & vbCrLf
End Sub
