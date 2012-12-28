VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form integr 
   Caption         =   "Integration calc"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6150
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   4455
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   615
         Left            =   3120
         TabIndex        =   24
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Creat integration calc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   22
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3360
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   735
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "integr.frx":0000
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "integr.frx":004F
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Calc"
         Height          =   735
         Left            =   3120
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   8
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Set"
         Height          =   615
         Left            =   3120
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "3"
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
         TabIndex        =   25
         Top             =   3000
         Width           =   135
      End
      Begin VB.Line Line2 
         X1              =   4320
         X2              =   360
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Запись файла"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         TabIndex        =   23
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "шт."
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
         Left            =   2880
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "2"
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
         TabIndex        =   18
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "1"
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
         TabIndex        =   17
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Select integration "
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Len."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1575
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   3930
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   3930
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Line"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Point"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label Label6 
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6240
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "integr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# INTEGATION
Dim Lines_integration(MAXL, 1) As Single
Dim IntegrationS(40) As Single


Private Sub Combo1_Change()
Label9 = Combo1.ListIndex + 1
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Not Me.List1.ListIndex = -1 Then
Lines_integration(Me.List1.ListIndex + 1, 0) = Combo1.ListIndex + 1
Lines_integration(Me.List1.ListIndex + 1, 1) = Text1
ElseIf Not Me.List2.ListIndex = -1 Then

End If

'Command3_Click

End Sub

Private Sub Command3_Click()
Dim i As Integer
For i = 0 To 40
    IntegrationS(i) = 0
Next
For i = 0 To 40
    Select Case Lines_integration(i, 0)
    Case 1, 5, 4, 7
        IntegrationS(Lines_integration(i, 0)) = IntegrationS(Lines_integration(i, 0)) + (Lines_integration(i, 1))
'    Case 7
'
'    Case 4
    End Select
'Lines_integration(i, 0)
'Lines_integration(i, 1)
'If Lines_integration(i + 1, 0) = 0 Then Exit For
Next

For i = 0 To 40
    Select Case i
    Case 1, 5, 4
        IntegrationS(i) = IntegrationS(i) \ 200
    End Select
Next
End Sub

Private Sub Form_Load()
Dim dtemp As String
Dim i As Integer

On Error GoTo ERR
For i = 1 To P_Bs
Me.List1.AddItem i
Next
For i = 1 To N_Points_Main
Me.List2.AddItem i
Next
Open App.Path & "\data\integr.dat" For Input As #1
While Not EOF(1)
Line Input #1, dtemp
Me.Combo1.AddItem dtemp
Wend
Close #1

'Me.Combo1.Text

'Global Line_PX(127, 100) As Single
'Global Line_PY(127, 100) As Single

'Global Points_m_X(127) As Single
'Global Points_m_Y(127) As Single

'P_Bs
Exit Sub
ERR:
MsgBox "The file " & App.Path & "\data\integr.dat" & " is needed for this program."
End Sub

Private Sub List1_Click()
Nline_S = List1.Text
Npoint_S = 0
Me.Label1 = "Линия " & Nline_S & " выбрана для редактирования."
If Lines_integration(Nline_S, 0) <> 0 Then
  Me.Combo1.Text = Me.Combo1.list(Lines_integration(Nline_S, 0) - 1)
  Me.Text1 = Lines_integration(Nline_S, 1)
  Me.Label1 = Me.Label1 & "  [" & Lines_integration(Nline_S, 0) & "]"
  Me.Label6 = Me.Combo1.Text & " = " & IntegrationS(Lines_integration(Nline_S, 0))
  Text4 = IntegrationS(Lines_integration(Nline_S, 0))
Else
  Me.Label6 = "Нет расчетов. Нажмите Calc."
  Me.Combo1.Text = "Линия не обозначена"
End If
ROOFPIC.Draw_Point ROOFPIC
End Sub

Private Sub List2_Click()
Npoint_S = List2.Text
Nline_S = 0
Me.Label1 = "Вершина " & Nline_S & " выбрана для редактирования."
ROOFPIC.Draw_Point ROOFPIC
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command2_Click
End Sub
