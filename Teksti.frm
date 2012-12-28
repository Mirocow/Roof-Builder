VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form Teksti 
   BackColor       =   &H00C8D0D4&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8895
   Icon            =   "Teksti.frx":0000
   LinkTopic       =   "‘ÓÏ‡1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8895
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin —hameleonButton.chameleonButton  ÓÏÏ‡Ì‰‡1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "&Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Teksti.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      ICONS           =   16
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "Teksti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetFont Me
End Sub

Private Sub  ÓÏÏ‡Ì‰‡1_Click()
    Me.Hide
End Sub

