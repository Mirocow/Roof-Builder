VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "—hameleonButton.ocx"

Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CheckBox Check1 
         Caption         =   "—Í‡ÌËÓ‚‡Ú¸ ÔÓ‰Ô‡ÔÍË"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4560
         TabIndex        =   5
         Top             =   720
         Width           =   5175
         Begin VB.CheckBox Check2 
            BackColor       =   &H00808000&
            Caption         =   "*.rfd (RFD)"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Value           =   2  'Grayed
            Width           =   4095
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00808000&
            Caption         =   "*.rbp (Roof Builder Project)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Open"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Dell"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   9840
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search by name a file:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
