VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form Set_prop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Height          =   2760
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OK"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Set items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Set_prop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()
Unload Me
End Sub
