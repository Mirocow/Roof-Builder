VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form ProgressBarf 
   BorderStyle     =   0  'None
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   2
   LinkTopic       =   "Form2"
   ScaleHeight     =   1005
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   10
      Top             =   10
      Width           =   4935
   End
End
Attribute VB_Name = "ProgressBarf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim descrip As String
Dim proges_color As String
Private bformS As Form
Private mouse As Integer

Sub ProgressBar1(bform As Form, ByVal perc As Integer, ByVal max As Integer)
Dim Num$

If perc = 0 Then Exit Sub
Screen.MousePointer = 11
If perc > max Then Unload Me: Exit Sub
Set bformS = bform
Me.Show vbModeless, bform
bform.Enabled = False
Me.SetFocus
Picture1.Cls
Picture1.ScaleWidth = 100
Picture1.DrawMode = 10
If descrip <> "" Then
Label1 = descrip
Else
Label1 = perc & " ... " & max
End If
'allprc = Picture1.ScaleWidth
'Picture1.ScaleWidth = max
perc = (perc / max) * 100
Num$ = Format$(perc, "###") + "%"
'Picture1.CurrentX = 50 - Picture1.TextWidth(Num$) / 2
'Picture1.CurrentY = (Picture1.ScaleHeight - Picture1.TextHeight(Num$)) / 2
If proges_color = "" Then proges_color = vbBlue
Picture1.Line (0, 0)-(perc, Picture1.ScaleHeight), proges_color, BF
Picture1.PSet (50 - Picture1.TextWidth(Num$) / 2, (Picture1.ScaleHeight - Picture1.TextHeight(Num$)) / 2), vbWhite
Picture1.Print Num$
Picture1.Refresh
DoEvents
End Sub

Sub ProgressBar_end()
Unload Me
End Sub

Sub set_descrip(descr As String)
descrip = descr
End Sub

Sub set_color(color As String, colorstr As String, colorback As String)
proges_color = color
Picture1.ForeColor = colorstr
Picture1.BackColor = colorback
End Sub

Private Sub Form_Load()
mouse = Screen.MousePointer
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not bformS Is Nothing Then
descrip = ""
bformS.Enabled = True
Set bformS = Nothing
End If
Screen.MousePointer = 0
End Sub

Private Sub Picture1_Click()
MsgBox "Engine: dmms@narod.ru"
'Unload Me
End Sub
