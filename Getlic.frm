VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form Getlic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check licence"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5145
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Text            =   "Insert your registration key"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Check"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00DCFBFC&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "1124470744112"
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00DCFBFC&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Compu Name"
      Top             =   480
      Width           =   3015
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   -120
      X2              =   5160
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   -120
      X2              =   5160
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label5 
      Caption         =   "Registration key"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Check licence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "MDinc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Licence number"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Getlic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim test1 As String
Dim test2 As String

getlicn.licnumber = Text2
getlicn.regnumber = Text3
If getlicn.ChekLicence(App.ProductName) Then
MsgBox GetResIDstring(9151)
SaveSetting App.ProductName, "REG", "Regnumber", Text3
SaveSetting App.ProductName, "Main", "licence", Text2
'About.process = About.process & vbcrlf & "Спасибо за то, что лицензировали программу."
SaveSetting App.ProductName, "REG", "datereg", Date
Gl.LicenceN = Text3
About.Command1.Visible = False
'OfficeStart.Command1.Visible = False
'About.Label1 = "Registered by"
Unload Me
About.Hide
Exit Sub
Else
MsgBox GetResIDstring(9152): End
End If
End Sub

Private Sub Command2_Click()
Me.Hide
'End
End Sub

Private Sub Form_Load()
Label1 = GetResIDstring(9145)
Label2 = GetResIDstring(9149)
Label5 = GetResIDstring(9146)

Label4 = GetResIDstring(9143)
Command1.Caption = GetResIDstring(9144)
Text2 = Gl.ProductID
Me.Text1 = Gl.UserName
End Sub

