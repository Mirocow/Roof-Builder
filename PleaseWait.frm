VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form PleaseWait 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Interupt"
      Height          =   375
      Left            =   3750
      TabIndex        =   1
      Top             =   650
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "www.roof-builder.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   50
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      Height          =   1140
      Left            =   0
      Top             =   0
      Width           =   5700
   End
End
Attribute VB_Name = "PleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Wait(Optional ByVal descr As String)
Me.Show vbModeless, OfficeStart
Label3 = descr
Label3.Refresh
Me.Refresh
OfficeStart.MousePointer = 99
End Sub


Public Sub SetDescription(descr As String)
Label3 = descr
Label3.Refresh
End Sub


Public Sub CloseForm()
Unload Me
End Sub


Private Sub Command1_Click()
Arr(LNC).Dll.InteruptCalc
OfficeStart.Enabled = True
Unload Me
End Sub


Private Sub Form_Load()
Command1.Caption = GetResIDstring(1493)
End Sub
