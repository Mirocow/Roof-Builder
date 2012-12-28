VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form GetActiveData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Activation Data"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get activation data"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "GetActiveData.frx":0000
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "GetActiveData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim mvarCOMPUTERNAME As String
        mvarCOMPUTERNAME = Gl.UserName
        Text1 = App.ProductName & " " & "V" & Ver & vbCrLf
        Text1 = Text1 & "Name: " & mvarCOMPUTERNAME & vbCrLf
        Text1 = Text1 & "Product ID: " & OfficeStart.PM.XorStr("" & Chr(1) & SystemData & Chr(1), mvarCOMPUTERNAME) & vbCrLf
        'Text1 = Text1 & Gl.licence & vbCrLf
        'Text1 = Text1 & Gl.Regnumber & vbCrLf
        'Text1 = Text1 & Gl.SystemData & vbCrLf
End Sub


Private Sub Command3_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Command1.Caption = Lng.GetResIDstring(9650)
    Me.Caption = Lng.GetResIDstring(9650)
    Command3.Caption = Lng.GetResIDstring(9120)
    Text1 = Lng.GetResIDstring(9651)
    'Me.Combo1.ListIndex = 0
End Sub

