VERSION 5.00
Begin VB.UserControl UserControl11 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   ScaleHeight     =   3600
   ScaleWidth      =   4050
   Begin VB.Frame Frame7 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   15
         Left            =   120
         TabIndex        =   2
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   14
         Left            =   1800
         TabIndex        =   1
         Text            =   "0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   840
         X2              =   3360
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   3360
         X2              =   3360
         Y1              =   3000
         Y2              =   1200
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   840
         X2              =   3360
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   840
         X2              =   840
         Y1              =   1200
         Y2              =   3000
      End
   End
End
Attribute VB_Name = "UserControl11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

