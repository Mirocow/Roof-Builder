VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Plugins 
   Caption         =   "Plugins"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14482428
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descriptions"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Plugins"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "Plugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
If fexit = False Then Cancel = -1: Me.Hide
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim strHcode1 As String
Dim strHcode2(16) As Byte
Dim ans As Long
Dim i, n As Integer
If Item.Checked = True Then Item.Checked = False: Exit Sub

If LoadDLLC.loaddlls(Item.Text, strHcode1) Then
Item.Checked = True
Else
strHcode1 = InputBox("Отсутствует лицензия. Отправьте приведенный ниже код по адресу roofbuilder@narod.ru. Полученный в ответ код введите в окне и нажмите Ok и перезапустите программу." & vbCrLf & strHcode1, "Введите лицензионный код.", "Код из письма")
If strHcode1 <> "Код из письма" Then
i = 0: n = 1
Do
    strHcode2(i) = "&H" & (Mid(strHcode1, n, 2)) ' Перевод HEX to DEC
    n = n + 2: i = i + 1
Loop Until i = Len(strHcode1) / 2
'strHcode2(i + 1) = Null
'ans = osSetCode(strHcode2(0), UBound(strHcode2))

ans = osSetCode(strHcode2(0), UBound(strHcode2))
End If
End If

End Sub
