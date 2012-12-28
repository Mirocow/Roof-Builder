VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form Move_and_change 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   Icon            =   "Change_size_tmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   840
      TabIndex        =   18
      Text            =   "100"
      Top             =   3360
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Apply"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add list"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808000&
      Caption         =   "Option2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808000&
      Caption         =   "Option1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      DisabledPicture =   "Change_size_tmp.frx":030A
      DownPicture     =   "Change_size_tmp.frx":074C
      DragIcon        =   "Change_size_tmp.frx":0B8E
      Height          =   495
      Left            =   120
      Picture         =   "Change_size_tmp.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Move Up"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      DisabledPicture =   "Change_size_tmp.frx":1412
      DownPicture     =   "Change_size_tmp.frx":1854
      DragIcon        =   "Change_size_tmp.frx":1C96
      Height          =   495
      Left            =   1800
      Picture         =   "Change_size_tmp.frx":20D8
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Move Down"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Text            =   "100"
      Top             =   2115
      Width           =   735
   End
   Begin VB.CommandButton Комманда3 
      DisabledPicture =   "Change_size_tmp.frx":251A
      DownPicture     =   "Change_size_tmp.frx":295C
      DragIcon        =   "Change_size_tmp.frx":2D9E
      Height          =   495
      Left            =   120
      Picture         =   "Change_size_tmp.frx":31E0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Move left"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Комманда2 
      DisabledPicture =   "Change_size_tmp.frx":3622
      DownPicture     =   "Change_size_tmp.frx":3A64
      DragIcon        =   "Change_size_tmp.frx":3EA6
      Height          =   495
      Left            =   1800
      Picture         =   "Change_size_tmp.frx":42E8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Move Right"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "100"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "100"
      Top             =   465
      Width           =   735
   End
   Begin VB.CommandButton Комманда4 
      DisabledPicture =   "Change_size_tmp.frx":472A
      DownPicture     =   "Change_size_tmp.frx":4B6C
      DragIcon        =   "Change_size_tmp.frx":4FAE
      Height          =   495
      Left            =   1800
      Picture         =   "Change_size_tmp.frx":53F0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Move Down"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Комманда1 
      DisabledPicture =   "Change_size_tmp.frx":5832
      DownPicture     =   "Change_size_tmp.frx":5C74
      DragIcon        =   "Change_size_tmp.frx":60B6
      Height          =   495
      Left            =   120
      Picture         =   "Change_size_tmp.frx":64F8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move Up"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dell list"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change length of list"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Movie list"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Movie list"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Move_and_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Sub ReleaseCapture Lib "user32" ()
'Private Const WM_NCLBUTTONDOWN = &HA1
'Private Const HTCAPTION = 2
'Dim ReturnValue As Long
'If Button = 1 Then
'Call ReleaseCapture
'ReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
'End If

Private Sub Combo1_Click()
Text4.Text = Combo1.Text
Command5_Click
End Sub

Private Sub Command1_Click()
On Error GoTo ERR
    Dim ars As Integer
    If SelectLists(0).list > 0 Then
    ars = UBound(SelectLists)
    Dim nL As Integer
    For nL = 0 To ars Step 1
        List_Properties_Cut(N_Slope, SelectLists(nL).list) = List_Properties_Cut(N_Slope, SelectLists(nL).list) + 1
        List_Properties_Length(N_Slope, SelectLists(nL).list, List_Properties_Cut(N_Slope, SelectLists(nL).list)) = Text3.Text
        List_Properties_PY(N_Slope, SelectLists(nL).list, List_Properties_Cut(N_Slope, SelectLists(nL).list)) = List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) + Text3.Text - Lapepic.SetProfilData(0).Text
        List_Properties_PX(N_Slope, SelectLists(nL).list, List_Properties_Cut(N_Slope, SelectLists(nL).list)) = List_Properties_PX(N_Slope, SelectLists(nL).list, 0)
    Next
    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Command2_Click()
On Error GoTo ERR
    Dim ars As Integer
    If SelectLists(0).list > 0 Then
    ars = UBound(SelectLists)
    Dim nL As Integer
        For nL = 0 To ars Step 1
            List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = 0
            List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = 0 'SlP(N_Slope).Pn_Red_lines
        Next
    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub



Private Sub Command3_Click()
On Error GoTo ERR
    Dim ars As Integer
    If SelectLists(0).list > 0 Then
        ars = UBound(SelectLists)
        For nL = 0 To ars Step 1
            If SelectLists(nL).list <> -1 Then
                If Option1.value Then
                    If List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) - Text3 > Lapepic.SetProfilData(3).Text Then
                        List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) - Text3
                        List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) - Text3
                    End If

                Else
                    List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) + Text3
                End If

            End If

        Next

        Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Command4_Click()
On Error GoTo ERR
    Dim ars As Integer
    If SelectLists(0).list > 0 Then
        ars = UBound(SelectLists)
        For nL = 0 To ars Step 1
            If SelectLists(nL).list <> -1 Then
                If Option1.value Then
                    List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) + Text3
                    List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) + Text3
                Else
                    If List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) - Text3 > Lapepic.SetProfilData(3).Text Then
                        List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_Length(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) - Text3
                    End If

                End If

            End If

        Next

        Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Command5_Click()
    If Text4 <> "" Then
        Dim len1 As Single
        If IsNumeric(Text4) = False Then Exit Sub
        On Error GoTo ERR
        Dim ars As Integer
        If SelectLists(0).list > 0 Then
        ars = UBound(SelectLists)
        Dim n As Integer
        For n = 0 To ars Step 1
            If Option1.value Then
                len1 = List_Properties_Length(N_Slope, SelectLists(n).list, SelectLists(n).Cut)
                List_Properties_Length(N_Slope, SelectLists(n).list, SelectLists(n).Cut) = CSng(Text4)
                List_Properties_PY(N_Slope, SelectLists(n).list, SelectLists(n).Cut) = List_Properties_PY(N_Slope, SelectLists(n).list, SelectLists(n).Cut) - (len1 - List_Properties_Length(N_Slope, SelectLists(n).list, SelectLists(n).Cut))
            Else
                len1 = List_Properties_Length(N_Slope, SelectLists(n).list, SelectLists(n).Cut)
                List_Properties_Length(N_Slope, SelectLists(n).list, SelectLists(n).Cut) = CSng(Text4)
                List_Properties_PY(N_Slope, SelectLists(n).list, SelectLists(n).Cut) = List_Properties_PY(N_Slope, SelectLists(n).list, SelectLists(n).Cut) '- (len1 - List_Properties_Length(N_Slope, SelectLists(n).List, SelectLists(n).Cut))
    
            End If
        Next
        Lapepic.Draw_Systems Lapepic.Picture1
        End If
    End If
    
ERR:
End Sub


Private Sub Form_Activate()
'If Lapepic.Check1.value Then
'
'Else
'Text4.Text = List_Properties_Length(N_Slope, SelectLists(N).list, SelectLists(N).Cut)
'End If
End Sub

Private Sub Form_Load()
    SetFont Me

    Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
    Me.Top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.Top)
    Me.Caption = lng.GetResIDstring(1409)

    Option1.Caption = lng.GetResIDstring(1498)
    Option2.Caption = lng.GetResIDstring(1497)

    Label1 = lng.GetResIDstring(1119)
    Label3 = lng.GetResIDstring(1120)
    Command2.Caption = lng.GetResIDstring(3009)
    Command1.Caption = lng.GetResIDstring(3018)
    Label2 = Label1

    Text2 = Lapepic.SetProfilData(1)
    
    If Lapepic.Check1.value Then
        Text4.Visible = False
        Text3.Enabled = False
        Command4.Enabled = False
        Command3.Enabled = False
        Option1.Enabled = False
        Option2.Enabled = False
        Dim i As Integer
        For i = 1 To Lapepic.txt_CL.ListCount - 1
            Combo1.AddItem Lapepic.txt_CL.list(i)
        Next
        Combo1.Visible = True
    Else
        Text4.Visible = True
        Combo1.Visible = False
        Text3.Enabled = True
        Command4.Enabled = True
        Command3.Enabled = True
        Option1.Enabled = True
        Option2.Enabled = True
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
SaveSetting App.ProductName, "Position", Me.name & "top", Me.Top
OfficeStart.menu_view_m(8).Checked = False
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Command5_Click
    End If
End Sub


Private Sub Комманда1_Click()
On Error GoTo ERR
    Dim ars As Integer
    If SelectLists(0).list > 0 Then
    ars = UBound(SelectLists)
    Dim nL As Integer
    For nL = 0 To ars Step 1
        If SelectLists(nL).list > -1 Then List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) + Val(Text2)
    Next

    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Комманда2_Click()
On Error GoTo ERR
    Dim ars As Integer
    If SelectLists(0).list > 0 Then
    ars = UBound(SelectLists)
    Dim nL As Integer
    For nL = 0 To ars Step 1
        If SelectLists(nL).list > -1 Then List_Properties_PX(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_PX(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) + Val(Text1)
    Next

    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Комманда3_Click()
On Error GoTo ERR
    Dim ars As Integer
    If SelectLists(0).list > 0 Then
    ars = UBound(SelectLists)
    Dim nL As Integer
    For nL = 0 To ars Step 1
        If SelectLists(nL).list > -1 Then List_Properties_PX(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_PX(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) - Val(Text1)
    Next

    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Комманда4_Click()
On Error GoTo ERR
    Dim ars As Integer
    If SelectLists(0).list > 0 Then
    ars = UBound(SelectLists)
    Dim nL As Integer
    For nL = 0 To ars Step 1
        If SelectLists(nL).list > -1 Then List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) = List_Properties_PY(N_Slope, SelectLists(nL).list, SelectLists(nL).Cut) - Val(Text2)
    Next

    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub

