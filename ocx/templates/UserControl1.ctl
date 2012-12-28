VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   ScaleHeight     =   3600
   ScaleWidth      =   4095
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   5
         Text            =   "0"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   3
         Text            =   "0"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   2
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Text            =   "0"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   1320
         X2              =   3480
         Y1              =   240
         Y2              =   3360
      End
      Begin VB.Line Line4 
         X1              =   1320
         X2              =   1800
         Y1              =   240
         Y2              =   3360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   480
         X2              =   3480
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   480
         X2              =   1320
         Y1              =   3360
         Y2              =   240
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

