VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{E9335107-9E4B-4A70-BDE4-6A18106C27BA}#1.0#0"; "SplitterModern.ocx"
Begin VB.Form ClientsList 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin SplitterModern.SplitHV SplitHV1 
      Height          =   6495
      Left            =   5400
      TabIndex        =   10
      Top             =   0
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   11456
      BackColor       =   8421376
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   0
      TabIndex        =   26
      Top             =   1560
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8493
      _Version        =   393216
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   5520
      TabIndex        =   11
      Top             =   1440
      Width           =   5775
      Begin VB.Frame Frame4 
         BackColor       =   &H00808000&
         ForeColor       =   &H00FFFFFF&
         Height          =   4935
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   5775
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   3000
            TabIndex        =   29
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1560
            TabIndex        =   27
            Top             =   1260
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   2280
            Width           =   5535
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1560
            TabIndex        =   17
            Top             =   1560
            Width           =   4095
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1560
            TabIndex        =   16
            Top             =   960
            Width           =   4095
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   4080
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
            Caption         =   "..."
            Height          =   285
            Left            =   5400
            TabIndex        =   13
            Top             =   240
            Width           =   255
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Left            =   120
            TabIndex        =   19
            Top             =   3240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   2778
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Path"
               Object.Width           =   6068
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer type"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Tel"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   30
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Org`s name"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1275
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Date of addition:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2880
            TabIndex        =   25
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Done projects"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   3000
            Width           =   5535
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact`s information:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Addres:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer`s Name:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer`s ID"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5390
      Begin VB.CommandButton Command3 
         Caption         =   "Seek"
         Height          =   495
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Text            =   "*"
         Top             =   1080
         Width           =   5175
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00808000&
         Caption         =   "Seek by date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "Seek by name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Seek by code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   7800
      TabIndex        =   5
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command6 
         Caption         =   "Update"
         Height          =   495
         Left            =   0
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "New"
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Remove"
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "ClientsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cinf As Recordset

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim requesr As String
If Option2 Then requesr = "c.cname"
If Option1 Then requesr = "c.cname"
If Option3 Then requesr = "c.date"

MSFlexGrid1.Clear

Set cinf = MainBaseFunction.RequestSQL("select * from CustomerInfo c where " & requesr & " like '" & Text1 & "' order by c.id")

If Not cinf Is Nothing Then

MSFlexGrid1.Cols = 3
MSFlexGrid1.Rows = cinf.RecordCount + 1


'Dim c As Integer
'For c = 0 To 2
MSFlexGrid1.TextMatrix(0, 0) = "ID"
MSFlexGrid1.TextMatrix(0, 1) = "Äàòà ñîçäàíèÿ"
MSFlexGrid1.TextMatrix(0, 2) = "Èìÿ êëèåíòà"
'Next

r = 1
'Dim r As Integer
Do While Not cinf.EOF
    For c = 0 To 2
    MSFlexGrid1.TextMatrix(r, c) = CheckNull(cinf.Fields(c))
    Next
    r = r + 1
    cinf.MoveNext
Loop
'cinf.Close
End If
'Set cinf = Nothing
End Sub

Private Sub Command4_Click()
Dim RS As Recordset
Dim LastID As Integer

Set RS = MainBaseFunction.RequestSQL("select max([ID]) as maxid from CustomerInfo")

If Not RS Is Nothing Then
LastID = CheckNull(RS!maxid) + 1
End If
RS.Close

Text6 = LastID

Set RS = MainBaseFunction.RequestSQL("select * from CustomerInfo where id=" & LastID)
If Not RS.EOF Then
    RS.Close
    MsgBox "Error: This ID already exists, input a different ID."
Exit Sub
End If

On Error GoTo ERR
dao.BeginTrans

RS.AddNew
RS!id = CheckNull(Text6, True)
RS!Date = CheckNull(Text2, True)
RS!cname = CheckNull(Text3, True)
RS!IDCTYPE = CheckNull(Combo1.ListIndex, True)
RS!oname = CheckNull(Text7, True)
RS!COORDINATE = CheckNull(Text4, True)
RS!Description = CheckNull(Text5, True)
RS!tel = CheckNull(Text8)

RS.Update
dao.CommitTrans
RS.Close
'cinf.Close
'
'cinf = MainBaseFunction.RequestSQL("select * from CustomerInfo c where " & requesr & " like '" & Text1 & "' order by c.id")

'cinf.AddNew
Exit Sub
ERR:
MsgBox ERR.Description
dao.Rollback
RS.Close
Exit Sub
End Sub

Private Sub Command5_Click()
'MonthView1.Visible = True
End Sub

Private Sub Form_Load()
Text2 = MonthView1

MSFlexGrid1.Cols = 3
MSFlexGrid1.ColWidth(0) = 500
MSFlexGrid1.ColWidth(2) = 1900
MSFlexGrid1.ColWidth(2) = 3000
MSFlexGrid1.ColAlignment(0) = 1

Set SplitHV1.obj1 = Frame1
Set SplitHV1.obj1 = MSFlexGrid1
'Set SplitHV1.obj2 = Frame2
Set SplitHV1.obj2 = Frame3

Dim RS As Recordset
Set RS = MainBaseFunction.RequestSQL("select * from CustomerType order by id")
If Not RS Is Nothing Then
Do While Not RS.EOF
    Combo1.AddItem RS.Fields(1), RS.Fields(0) - 1
RS.MoveNext
Loop
End If
Combo1.ListIndex = 0
Command3_Click
End Sub

Private Sub Form_Resize()
SplitHV1.Height = Me.Height
Frame3.Height = Me.Height
'MSFlexGrid1.Height = Me.Height - Frame1.Height
SplitHV1.ResizeControl
End Sub

Private Sub Form_Unload(Cancel As Integer)
cinf.Close
Set cinf = Nothing
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text2 = DateClicked
MonthView1.Visible = False
End Sub

Private Sub MSFlexGrid1_Click()
'cinf.MoveFirst
cinf.FindFirst "id = " & CLng(MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, 0))
Text6 = CheckNull(cinf!id)
Combo1.ListIndex = CheckNull(cinf!IDCTYPE)
Text2 = CheckNull(cinf!Date)
Text3 = CheckNull(cinf!cname)
Text7 = CheckNull(cinf!oname)
Text4 = CheckNull(cinf!COORDINATE)
Text5 = CheckNull(cinf!Description)
Text8 = cinf!tel

Dim RS As Recordset

Set RS = MainBaseFunction.RequestSQL("select * from Projectsfile c where c.IDCUSTOMER like '" & cinf!id & "' order by c.id")


If Not RS Is Nothing Then
Dim itmX As ListItem
ListView1.ListItems.Clear

If RS.RecordCount > 0 Then
Do While Not RS.EOF
    Set itmX = ListView1.ListItems.Add(, , RS!id)
    itmX.SubItems(1) = 0
    itmX.SubItems(2) = IIf(IsNull(RS!Path), "", RS!Path)
RS.MoveNext
Loop
End If

RS.Close
End If

Set RS = Nothing
End Sub

Private Sub SplitHV1_MoveEnd()
Frame2.Left = Me.Width - Frame2.Width
End Sub

