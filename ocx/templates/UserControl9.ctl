VERSION 5.00
Begin VB.UserControl UserControl9 
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ScaleHeight     =   4020
   ScaleWidth      =   4065
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   3
         Text            =   "0"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   2
         Text            =   "0"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   5
         Left            =   600
         TabIndex        =   1
         Text            =   "0"
         Top             =   2040
         Width           =   615
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   480
         X2              =   3480
         Y1              =   240
         Y2              =   3360
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   480
         X2              =   3480
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   480
         X2              =   480
         Y1              =   3360
         Y2              =   240
      End
   End
End
Attribute VB_Name = "UserControl9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

