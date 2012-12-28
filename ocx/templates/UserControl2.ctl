VERSION 5.00
Begin VB.UserControl UserControl2 
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   ScaleHeight     =   4035
   ScaleWidth      =   4080
   Begin VB.Frame Frame13 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   41
         Left            =   2280
         TabIndex        =   5
         Text            =   "0"
         Top             =   285
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   42
         Left            =   120
         TabIndex        =   4
         Text            =   "0"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   43
         Left            =   1800
         TabIndex        =   3
         Text            =   "0"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   44
         Left            =   2760
         TabIndex        =   2
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   45
         Left            =   600
         TabIndex        =   1
         Text            =   "0"
         Top             =   2880
         Width           =   615
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   120
         X2              =   3600
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line43 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   3600
         X2              =   3600
         Y1              =   3360
         Y2              =   720
      End
      Begin VB.Line Line44 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   120
         X2              =   1440
         Y1              =   3360
         Y2              =   720
      End
      Begin VB.Line Line45 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   1440
         X2              =   3600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line46 
         X1              =   1440
         X2              =   1440
         Y1              =   720
         Y2              =   3360
      End
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

