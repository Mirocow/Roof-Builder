VERSION 5.00
Begin VB.UserControl UserControl5 
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ScaleHeight     =   4020
   ScaleWidth      =   4065
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   24
         Left            =   600
         TabIndex        =   3
         Text            =   "0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   25
         Left            =   2160
         TabIndex        =   2
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   26
         Left            =   1560
         TabIndex        =   1
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.Line Line26 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   480
         X2              =   480
         Y1              =   3360
         Y2              =   360
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   480
         X2              =   3480
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line28 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   3480
         X2              =   480
         Y1              =   360
         Y2              =   3360
      End
   End
End
Attribute VB_Name = "UserControl5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

