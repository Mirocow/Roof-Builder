VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form SetPlugins 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Description:"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   6975
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   6855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   3495
      Left            =   60
      Top             =   60
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3495
      Left            =   45
      Top             =   45
      Width           =   7095
   End
End
Attribute VB_Name = "SetPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
SetFont Me
Me.Caption = GetResIDstring(1037)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If fexit Then Exit Sub
Cancel = -1
'Me.Hide
End Sub

Private Sub List1_Click()
On Error Resume Next
LNÑ = List1.ListIndex
Text1 = arr(LNÑ).Dll.About
Text1 = Text1 & vbCrLf & "Library: " & arr(LNÑ).Pname & " Ver: " & arr(LNÑ).Dll.RBLibVer
Project.Text2 = List1.List(LNÑ)
Label2 = "Copyright: " & arr(LNÑ).Dll.Copyright
End Sub

Private Sub List1_DblClick()
Me.Hide
End Sub

