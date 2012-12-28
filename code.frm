VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As String
i = 1
Text2 = ""
While i <= Len(Text1)
    Text2 = Text2 & "Chr(" & Asc(Mid(Text1, i, 1)) & ") & "
    i = i + 1
Wend
End Sub

