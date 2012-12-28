VERSION 5.00
Begin VB.UserControl UserControl10 
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   ScaleHeight     =   3900
   ScaleWidth      =   4245
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   30
         Left            =   240
         TabIndex        =   5
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   31
         Left            =   2760
         TabIndex        =   4
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   32
         Left            =   720
         TabIndex        =   3
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   33
         Left            =   1320
         TabIndex        =   2
         Text            =   "0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   34
         Left            =   1440
         TabIndex        =   1
         Text            =   "0"
         Top             =   0
         Width           =   615
      End
      Begin VB.Line Line32 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   1800
         X2              =   360
         Y1              =   3600
         Y2              =   480
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   360
         X2              =   3360
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line34 
         X1              =   1440
         X2              =   1800
         Y1              =   480
         Y2              =   3600
      End
      Begin VB.Line Line35 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   3360
         X2              =   1800
         Y1              =   480
         Y2              =   3600
      End
   End
End
Attribute VB_Name = "UserControl10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

