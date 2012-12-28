VERSION 5.00
Begin VB.UserControl UserControl7 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   ScaleHeight     =   3600
   ScaleWidth      =   4050
   Begin VB.Frame Frame10 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   27
         Left            =   1560
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   28
         Left            =   1320
         TabIndex        =   2
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   29
         Left            =   2760
         TabIndex        =   1
         Text            =   "0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.Line Line29 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   480
         X2              =   3480
         Y1              =   360
         Y2              =   3360
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   480
         X2              =   3480
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line31 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   3480
         X2              =   3480
         Y1              =   3360
         Y2              =   360
      End
   End
End
Attribute VB_Name = "UserControl7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

