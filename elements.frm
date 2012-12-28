VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
