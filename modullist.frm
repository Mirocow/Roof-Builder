VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form modullist 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Description"
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
      TabIndex        =   4
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label modullist 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "modullist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
