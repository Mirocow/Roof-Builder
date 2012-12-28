VERSION 5.00
Begin VB.UserControl UserControl8 
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   ScaleHeight     =   3960
   ScaleWidth      =   4080
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   18
         Left            =   2640
         TabIndex        =   3
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   22
         Left            =   1320
         TabIndex        =   2
         Text            =   "0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   23
         Left            =   1440
         TabIndex        =   1
         Text            =   "0"
         Top             =   3480
         Width           =   615
      End
      Begin VB.Line Line23 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   3480
         X2              =   3480
         Y1              =   3360
         Y2              =   240
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   480
         X2              =   3480
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line25 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   3480
         X2              =   480
         Y1              =   240
         Y2              =   3360
      End
   End
End
Attribute VB_Name = "UserControl8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

