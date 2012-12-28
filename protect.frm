VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function callbycode Lib "protect.dll" (code As Long) As Long
Private Declare Function START_CRYPT Lib "protect.dll" Alias "start_crypt" (code As Long) As Long
Private Declare Function END_CRYPT Lib "protect.dll" Alias "end_crypt" (code As Long) As Long

'static long start_crypt_long [] = { 0x12345600, 0x12345601, 0x12345602, 0x12345603, 0x12345604};
'static long end_crypt_long [] = { 0x12345610, 0x12345611, 0x12345612, 0x12345613, 0x12345614};

Const cSTART_CRYPT1 = &H12345600
Const cEND_CRYPT1 = &H12345610

Const cSTART_CRYPT2 = &H12345601
Const cEND_CRYPT2 = &H12345611

Const cSTART_CRYPT3 = &H12345602
Const cEND_CRYPT3 = &H12345612

Const cSTART_CRYPT4 = &H12345603
Const cEND_CRYPT4 = &H12345613

Const cSTART_CRYPT5 = &H12345604
Const cEND_CRYPT5 = &H12345614

Private Sub Command1_Click()

START_CRYPT cSTART_CRYPT1

callbycode 0
t

END_CRYPT cEND_CRYPT1

End Sub

Sub t()
MsgBox "tttttt"
End Sub

Private Sub Command3_Click()
START_CRYPT cSTART_CRYPT2

Dim d As Long
Dim f As Integer
If d = 4 Then
d = CLng(f * f)
End If

END_CRYPT cEND_CRYPT2
End Sub

Private Sub Command4_Click()
START_CRYPT cSTART_CRYPT4

Dim d As Long
d = 45
d = d * 55555
Dim b As Long
b = d ^ d

END_CRYPT cEND_CRYPT4
End Sub

Private Sub Form_Load()

START_CRYPT cSTART_CRYPT3

Dim s As String
s = test("ddddddddd", "eeeeeeeeee")
MsgBox s

END_CRYPT cEND_CRYPT3

End Sub
