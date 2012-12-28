VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form calc 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Ok"
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command18 
         Caption         =   "C"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1920
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Text            =   "0"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Text            =   "0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   1320
         TabIndex        =   4
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Labelf 
         Caption         =   "g"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   32
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Labelf 
         Caption         =   "b"
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   31
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Labelf 
         Caption         =   "a"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   30
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Labelf 
         Caption         =   "F"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   29
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Labelf 
         Caption         =   "D"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   28
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Labelf 
         Caption         =   "B"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   27
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Labelf 
         Caption         =   "A"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   26
         Top             =   1680
         Width           =   255
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         X1              =   240
         X2              =   1320
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line9 
         X1              =   1320
         X2              =   240
         Y1              =   240
         Y2              =   1680
      End
      Begin VB.Line Line10 
         X1              =   1320
         X2              =   2520
         Y1              =   240
         Y2              =   1680
      End
      Begin VB.Line Line11 
         X1              =   1320
         X2              =   1320
         Y1              =   1680
         Y2              =   240
      End
      Begin VB.Label Labelf 
         Caption         =   "C"
         Height          =   255
         Index           =   2
         Left            =   1250
         TabIndex        =   25
         Top             =   1430
         Width           =   255
      End
      Begin VB.Line Line12 
         X1              =   2520
         X2              =   1320
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label24 
         Caption         =   "A"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label25 
         Caption         =   "B"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label26 
         Caption         =   "C"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "D"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "F"
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label29 
         Caption         =   "a"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label33 
         Caption         =   "b"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label34 
         Caption         =   "g"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label S 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label h 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "h"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   255
      End
   End
End
Attribute VB_Name = "calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
