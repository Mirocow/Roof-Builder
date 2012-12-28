VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "from grid"
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   3120
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "From base"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   2760
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5760
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2160
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   5760
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Option1.Value = True Then
    Dim rs As Recordset
    If MainBaseFunction.Connect(App.Path & "\cfg\materials.mdb") Then
    Set rs = MainBaseFunction.RequestSQL("select * from profils p where p.id=(select ProfiName.ID from ProfiName where ProfiName.Name=" & """" & Combo1.Text & """)")
    If rs Is Nothing Then Exit Sub
    Dim nr As Integer, i As Integer
    List1.Clear
    Do While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            List1.AddItem rs.Fields(i)
        Next
        rs.MoveNext
    Loop
    End If
Else
'    MSFlexGrid1.f
'    rs.FindFirst
'    ws.
'    db.
'    MSFlexGrid1.
End If
End Sub

Private Sub Command1_Click()
On Error GoTo ERR

Dim rs As Recordset
If MainBaseFunction.Connect(App.Path & "\cfg\materials.mdb") Then
Set rs = MainBaseFunction.RequestSQL(IIf(Text1.Text <> "", Text1.Text, _
"select p.id, pn.name from profils p, ProfiName pn  where pn.id=p.id" _
))
rs.MoveFirst

'MSFlexGrid1.Clear
MSFlexGrid1.Cols = rs.Fields.Count
MSFlexGrid1.Rows = rs.RecordCount

' Enumerate the specified Recordset object.
Dim nr As Integer, nc As Integer
With rs
    For nc = 0 To .Fields.Count - 1
        MSFlexGrid1.TextMatrix(0, nc) = .Fields(nc).Name
    Next
    
'    nr = 1
    Do While Not .EOF
        nr = nr + 1
        MSFlexGrid1.Rows = nr + 1
        For nc = 0 To .Fields.Count - 1
            MSFlexGrid1.TextMatrix(nr, nc) = .Fields(nc)
            If nc = 1 Then Combo1.AddItem .Fields(1)
            MSFlexGrid1.ColWidth(0) = 0
        Next
        .MoveNext
    Loop
End With


End If

ERR:
If Not rs Is Nothing Then rs.Close
Set rs = Nothing
End Sub
