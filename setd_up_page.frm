VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form Setd 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings of  data`s page "
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Height          =   1335
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Height          =   375
      Left            =   120
      MaxLength       =   150
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Ok"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   7695
   End
End
Attribute VB_Name = "Setd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Gl.Firm_name = Text1
Gl.Firm_r = Text2
Unload Me
End Sub

Private Sub Form_Load()
Me.Text1 = Gl.Firm_name
Me.Text2 = Gl.Firm_r
Me.Caption = ResolveResstring(1134)
End Sub

