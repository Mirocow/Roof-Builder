VERSION 5.00
Begin VB.Form aboutfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   $"about.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   720
      Y1              =   369
      Y2              =   369
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   720
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Splitter Modern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "aboutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label3 = App.Major & "." & App.Minor & "." & App.Revision
End Sub
