VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form splash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splash.frx":030A
   ScaleHeight     =   2010
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Max             =   20
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "[-]"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Screen.MousePointer = 0
    splash.ProgressBar1.MAX = 10
End Sub

Public Sub SetProgress(value As Single)
Label2.Caption = value
ProgressBar1.value = CInt(value)
End Sub
