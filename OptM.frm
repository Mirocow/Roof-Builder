VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form OptM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Опции рассчета"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3600
   LinkTopic       =   "Форма1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Комманда1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtN 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtW 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox Pr 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Комбо1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Метка2 
      Caption         =   "Нахлест:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Метка1 
      Caption         =   "Ширина:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "OptM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Pr_Click()
txtW = len_L(Pr.ListIndex)
txtN = len_w_L(Pr.ListIndex)
End Sub

Private Sub Комманда1_Click()
Unload Me
End Sub
