VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form Set_mat_prop 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   2370
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Max\Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Min\Length"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Set: "
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
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Set_mat_prop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim itmX As ListItem
'If Gl.set_prop = False Then
Set itmX = matedit.ListView2.ListItems.Add(, , Text1)
itmX.SubItems(1) = Text2
'Else
'Set itmX = matedit.ListView1.ListItems.Add(, , Text1)
'itmX.SubItems(1) = Text2
'End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Gl.set_prop = False Then
Label1 = "Length"
Label2 = "Quantity"
Else
Label2 = "Min"
Label1 = "Max"
End If
End Sub
