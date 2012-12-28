VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   -15
   ClientWidth     =   8370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   10
   HasDC           =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Tausta.frx":0000
   ScaleHeight     =   5220
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8400
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   3495
      Left            =   8040
      MouseIcon       =   "Tausta.frx":775C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PATENT: 2009611392"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5880
      MouseIcon       =   "Tausta.frx":78AE
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "+7 (495) 514-63-53"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   195
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Call VarPtr("VMProtect begin")
Label5 = "v" & Gl.Ver '& " - " & IIf(Gl.PV = "Prof  ", "Registered", "Unregistered")
'If Not IsLic Then
''    L = "Roof Builder " & OfficeStart.Label1 & vbNewLine
''    L = L & "Name: " & Gl.UserName & vbNewLine
'Else
''    L = lng.GetResIDstring(9670, "%NAME%", Gl.UserName)
'End If
'L = L & vbNewLine & vbNewLine & "MAINMAXSLOPELINE = " & MAINMAXSLOPELINE & vbNewLine
'L = L & "MAXSLOPELINE = " & MAXSLOPELINE & vbNewLine
'
'L = L & "MAXSLOPES = " & MAXSLOPES & vbNewLine
'L = L & "MAXSLOPELISTS = " & MAXSLOPELISTS & vbNewLine

Label5 = Label5 & vbNewLine & Plgs(LNC).Pname & " Ver: " & Plgs(LNC).Dll.RBLibVer

' http://roof-builder.ru/rbcurrent.php?ver=3.0.168
'Navigate Me, "http://roof-builder.ru/rbcurrent.php?ver=" & Gl.Ver
'Label6.Visible = False
'Label4.Visible = False

Call VarPtr("VMProtect end")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub L_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Label2_Click()
    Navigate Me, "http://roof-builder.ru/pics/roofbuilder.jpg"
End Sub

Private Sub Label3_Click()
    Navigate Me, "http://roof-builder.ru"
End Sub

Private Sub Label5_Click()
Unload Me
End Sub
