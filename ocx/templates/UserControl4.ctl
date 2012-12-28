VERSION 5.00
Begin VB.UserControl UserControl4 
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   ScaleHeight     =   4035
   ScaleWidth      =   4080
   Begin VB.Frame Frame12 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox value 
         Height          =   285
         Index           =   35
         Left            =   1680
         TabIndex        =   6
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   36
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   37
         Left            =   480
         TabIndex        =   4
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   38
         Left            =   1680
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   39
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox value 
         Height          =   285
         Index           =   40
         Left            =   1680
         TabIndex        =   1
         Text            =   "0"
         Top             =   3480
         Width           =   615
      End
      Begin VB.Line Line36 
         X1              =   1320
         X2              =   1320
         Y1              =   720
         Y2              =   3360
      End
      Begin VB.Line Line37 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   1320
         X2              =   2640
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line38 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   2640
         X2              =   3600
         Y1              =   3360
         Y2              =   720
      End
      Begin VB.Line Line39 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   1320
         X2              =   120
         Y1              =   3360
         Y2              =   720
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         X1              =   120
         X2              =   3600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line41 
         X1              =   2640
         X2              =   2640
         Y1              =   720
         Y2              =   3360
      End
      Begin VB.Label Label4 
         Caption         =   "0"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "UserControl4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

