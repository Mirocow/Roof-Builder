VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form DButil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Base Config Utilite"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Repair Data Base"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compact Data Base"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "DButil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MainBaseFunction.CloseDB
CompactDatabase Gl.FileName, Catalogue & "\Croorbuilder.mdb", True, False
MainBaseFunction.Connect Gl.FileName
End Sub

'Private Sub Command2_Click()
'RepairDatabase Gl.FileName
'End Sub
Private Sub Form_Load()

End Sub
