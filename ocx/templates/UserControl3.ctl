VERSION 5.00
Begin VB.UserControl UserControl3 
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   ScaleHeight     =   4050
   ScaleWidth      =   4050
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   16
         Left            =   2520
         TabIndex        =   5
         Text            =   "0"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   4
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   19
         Left            =   1440
         TabIndex        =   3
         Text            =   "0"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   20
         Left            =   2880
         TabIndex        =   2
         Text            =   "0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   21
         Left            =   840
         TabIndex        =   1
         Text            =   "0"
         Top             =   285
         Width           =   615
      End
      Begin VB.Line Line18 
         X1              =   2280
         X2              =   2280
         Y1              =   720
         Y2              =   3360
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   120
         X2              =   2280
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line20 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   3600
         X2              =   2280
         Y1              =   3360
         Y2              =   720
      End
      Begin VB.Line Line21 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   120
         X2              =   120
         Y1              =   3360
         Y2              =   720
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   120
         X2              =   3600
         Y1              =   3360
         Y2              =   3360
      End
   End
End
Attribute VB_Name = "UserControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

