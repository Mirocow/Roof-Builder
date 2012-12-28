VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   5655
      Begin VB.ListBox List1 
         BackColor       =   &H00DCFBFC&
         Height          =   2595
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "in milimeters"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "in moduls"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.ListBox List2 
         Height          =   2595
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Usable lengths"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "All modul lengths"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   0
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
