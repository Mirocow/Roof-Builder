VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form SetPoint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set point"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3675
   StartUpPosition =   1  'CenterOwner
   Begin ÑhameleonButton.chameleonButton chameleonButton2 
      Height          =   680
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1191
      BTYPE           =   2
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
      MICON           =   "Set_point.frx":0000
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
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1335
   End
   Begin ÑhameleonButton.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "&Cancel"
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
      MICON           =   "Set_point.frx":001C
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
   Begin VB.Label Label2 
      Caption         =   "Y ="
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "X ="
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   150
      Width           =   375
   End
End
Attribute VB_Name = "SetPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
If Lapepic.chameleonButton1.value = True Then
    Lapepic.Picture1_KeyDown 116, 0
Else
    Lapepic.Picture1_KeyDown 117, 0
End If
If FlagDraw = -1 Then
Lapepic.Picture1_MouseDown 2, 0, 0, 0
End If
Unload Me
End Sub

Private Sub chameleonButton2_Click()
    
    Lapepic.Draw_Plate_Line_KeyDown 13, 0
    
    If Lapepic.chameleonButton1.value = True Then

        ' X
        Lapepic.Line2.Visible = True
        Lapepic.Line2.x2 = Lapepic.Line2.X1 + ConvertData(Text1.Text, True)
        
        ' Y
        Lapepic.Line2.Visible = True
        Lapepic.Line2.y2 = Lapepic.Line2.Y1 + ConvertData(Text2.Text, True)

    ElseIf Lapepic.chameleonButton2.value = True Then
    
        Lapepic.Line2.x2 = ConvertData(Text1.Text, True)
        Lapepic.Line2.y2 = ConvertData(Text2.Text, True)

    End If
    
    Lapepic.Draw_Plate_Line_KeyDown 13, 0
    
    Text1.Text = 0
    Text2.Text = 0
    
End Sub

Private Sub Form_Load()
    If Lapepic.chameleonButton1.value = True Then
        Me.Caption = Lapepic.chameleonButton1.Caption
    ElseIf Lapepic.chameleonButton2.value = True Then
        Me.Caption = Lapepic.chameleonButton2.Caption
    End If
Lapepic.Line2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Lapepic.Draw_Plate_Line_KeyDown 27, 0
'Lapepic.chameleonButton1.Enabled = False
'Lapepic.chameleonButton1.value = False
'Lapepic.chameleonButton2.Enabled = False
'Lapepic.chameleonButton2.value = False
'Lapepic.chameleonButton1.value = False
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chameleonButton2_Click
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chameleonButton2_Click
    End If
End Sub
