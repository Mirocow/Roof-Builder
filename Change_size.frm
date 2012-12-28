VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form Move_and_change 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[:::]"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3465
   Icon            =   "Change_size.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin СhameleonButton.chameleonButton Command4 
      Height          =   405
      Left            =   2400
      TabIndex        =   17
      Top             =   2640
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   7
      TX              =   ""
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
      MICON           =   "Change_size.frx":030A
      PICN            =   "Change_size.frx":0326
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
   Begin СhameleonButton.chameleonButton Command3 
      Height          =   405
      Left            =   2985
      TabIndex        =   18
      Top             =   2640
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   7
      TX              =   ""
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
      MICON           =   "Change_size.frx":0778
      PICN            =   "Change_size.frx":0794
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
   Begin СhameleonButton.chameleonButton Комманда1 
      Height          =   405
      Left            =   2400
      TabIndex        =   10
      Top             =   480
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   7
      TX              =   ""
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
      MICON           =   "Change_size.frx":0BE6
      PICN            =   "Change_size.frx":0C02
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
   Begin СhameleonButton.chameleonButton Command5 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "Apply"
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
      MICON           =   "Change_size.frx":1054
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
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      Height          =   355
      Left            =   2400
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Text            =   "0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Text            =   "0"
      Top             =   1000
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin СhameleonButton.chameleonButton Command2 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "Apply"
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
      MICON           =   "Change_size.frx":1070
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
   Begin СhameleonButton.chameleonButton Command1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3660
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   7
      TX              =   "Apply"
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
      MICON           =   "Change_size.frx":108C
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
   Begin СhameleonButton.chameleonButton isButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   7
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
      MICON           =   "Change_size.frx":10A8
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
   Begin СhameleonButton.chameleonButton Комманда4 
      Height          =   405
      Left            =   2985
      TabIndex        =   14
      Top             =   480
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   7
      TX              =   ""
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
      MICON           =   "Change_size.frx":10C4
      PICN            =   "Change_size.frx":10E0
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
   Begin СhameleonButton.chameleonButton Комманда3 
      Height          =   405
      Left            =   2400
      TabIndex        =   15
      Top             =   1400
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   7
      TX              =   ""
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
      MICON           =   "Change_size.frx":1532
      PICN            =   "Change_size.frx":154E
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
   Begin СhameleonButton.chameleonButton Комманда2 
      Height          =   405
      Left            =   2985
      TabIndex        =   16
      Top             =   1400
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   7
      TX              =   ""
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
      MICON           =   "Change_size.frx":19A0
      PICN            =   "Change_size.frx":19BC
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change length of list"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Movie list"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Movie list"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Move_and_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ERR

    If SelectLists.Count > 0 Then
    
    ' Устанавливаем флаг изменения данных
    SetChange True
    
    Dim cul As cList
    For Each cul In SelectLists.Items
        
        SlP(N_Slope).CountSheets = SlP(N_Slope).CountSheets + 1
        List_Properties_Length(N_Slope, SlP(N_Slope).CountSheets) = ConvertData(Text4.Text, True)
        List_Properties_PX(N_Slope, SlP(N_Slope).CountSheets) = List_Properties_PX(N_Slope, cul.List)
        
        If Option1.value Then
            List_Properties_PY(N_Slope, SlP(N_Slope).CountSheets) = List_Properties_PY(N_Slope, cul.List) + ConvertData(Text4.Text, True) - Lapepic.SetProfilData(0).Tag
        Else
            List_Properties_PY(N_Slope, SlP(N_Slope).CountSheets) = List_Properties_PY(N_Slope, cul.List) - List_Properties_Length(N_Slope, cul.List) + Lapepic.SetProfilData(0).Tag
        End If
        
    Next
    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Command2_Click()
On Error GoTo ERR

    If SelectLists.Count > 0 Then
    
    ' Устанавливаем флаг изменения данных
    SetChange True
    
    Dim cul As cList
    For Each cul In SelectLists.Items
        List_Properties_Length(N_Slope, cul.List) = 0
        List_Properties_PY(N_Slope, cul.List) = 0 'SlP(N_Slope).Pn_Red_lines
    Next
    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub



Private Sub Command3_Click()
On Error GoTo ERR

    If SelectLists.Count > 0 Then
        
        Dim cul As cList
        For Each cul In SelectLists.Items
            If cul.List <> -1 Then
                
                If ConvertData(Text3, True) <> 0 Then
                    ' Устанавливаем флаг изменения данных
                    SetChange True
                    If Option1.value Then
                        
                        If List_Properties_Length(N_Slope, cul.List) - ConvertData(Text3, True) > Lapepic.SetProfilData(3).Tag Then
                            List_Properties_Length(N_Slope, cul.List) = List_Properties_Length(N_Slope, cul.List) - ConvertData(Text3, True)
                            List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) - ConvertData(Text3, True)
                        End If
    
                    Else
                        List_Properties_Length(N_Slope, cul.List) = List_Properties_Length(N_Slope, cul.List) + ConvertData(Text3, True)
                    End If
                End If

            End If

        Next

        Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Command4_Click()
On Error GoTo ERR

    If SelectLists.Count > 0 Then
        
        Dim cul As cList
        For Each cul In SelectLists.Items
            If cul.List <> -1 Then
            
                If ConvertData(Text3, True) <> 0 Then
                    ' Устанавливаем флаг изменения данных
                    SetChange True
                    If Option1.value Then
                        
                        List_Properties_Length(N_Slope, cul.List) = List_Properties_Length(N_Slope, cul.List) + ConvertData(Text3, True)
                        List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) + ConvertData(Text3, True)
                        
                    Else
                        
                        If List_Properties_Length(N_Slope, cul.List) - ConvertData(Text3, True) > Lapepic.SetProfilData(3).Tag Then
                            List_Properties_Length(N_Slope, cul.List) = List_Properties_Length(N_Slope, cul.List) - ConvertData(Text3, True)
                        End If
    
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
        
        On Error GoTo ERR
        If SelectLists.Count > 0 Then

        Dim cul As cList
        For Each cul In SelectLists.Items
            
            If ConvertData(Text4, True) <> 0 Then
            
                ' Устанавливаем флаг изменения данных
                SetChange True
            
                If Option1.value Then
                    
                    len1 = List_Properties_Length(N_Slope, cul.List)
                    List_Properties_Length(N_Slope, cul.List) = ConvertData(Text4, True)
                    List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) '- (len1 - List_Properties_Length(N_Slope, cul.List))
    
                    
                Else
                                    
                    len1 = List_Properties_Length(N_Slope, cul.List)
                    List_Properties_Length(N_Slope, cul.List) = ConvertData(Text4, True)
                    List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) - (len1 - List_Properties_Length(N_Slope, cul.List))
    
                End If
            End If
            
        Next
        
        Lapepic.Draw_Systems Lapepic.Picture1
        End If
    End If
ERR:
End Sub


Private Sub Form_Load()
    SetFont Me

    Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
    Me.Top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.Top)
    Me.Caption = lng.GetResIDstring(1409)

    Option1.Caption = lng.GetResIDstring(1497)
    Option2.Caption = lng.GetResIDstring(1498)

    Label1 = lng.GetResIDstring(1119)
    Label3 = lng.GetResIDstring(1120)
    Command2.Caption = lng.GetResIDstring(3009)
    Command1.Caption = lng.GetResIDstring(3018)
    Command5.Caption = lng.GetResIDstring(9378)
    Label2 = Label1

    Text2 = Lapepic.SetProfilData(1).Text
    
    If Lapepic.Check1.value Then
        Text3.Enabled = False
        Command4.Enabled = False
        Command3.Enabled = False
    Else
        Text3.Enabled = True
        Command4.Enabled = True
        Command3.Enabled = True
    End If
    
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
SaveSetting App.ProductName, "Position", Me.name & "top", Me.Top
OfficeStart.menu_view_m(8).Checked = False
End Sub


Private Sub isButton1_Click()
Unload Me
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    ' Устанавливаем флаг изменения данных
    SetChange True
    Command5_Click
End If
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ' Устанавливаем флаг изменения данных
    SetChange True
    Command5.value = True
End If
End Sub

Private Sub Комманда1_Click()
On Error GoTo ERR

    If SelectLists.Count > 0 Then

    Dim cul As cList
    For Each cul In SelectLists.Items
        If cul.List > -1 And ConvertData(Text2, True) <> 0 Then
            ' Устанавливаем флаг изменения данных
            SetChange True
            List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) + ConvertData(Text2, True)
        End If
    Next

    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Комманда2_Click()
On Error GoTo ERR

    If SelectLists.Count > 0 Then
    
    Dim cul As cList
    For Each cul In SelectLists.Items
        If cul.List > -1 And ConvertData(Text1, True) <> 0 Then
            ' Устанавливаем флаг изменения данных
            SetChange True
            List_Properties_PX(N_Slope, cul.List) = List_Properties_PX(N_Slope, cul.List) + ConvertData(Text1, True)
        End If
    Next

    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Комманда3_Click()
On Error GoTo ERR

    If SelectLists.Count > 0 Then
    
    Dim cul As cList
    For Each cul In SelectLists.Items
        If cul.List > -1 And ConvertData(Text1, True) <> 0 Then
            ' Устанавливаем флаг изменения данных
            SetChange True
            List_Properties_PX(N_Slope, cul.List) = List_Properties_PX(N_Slope, cul.List) - ConvertData(Text1, True)
        End If
    Next

    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


Private Sub Комманда4_Click()
On Error GoTo ERR

    If SelectLists.Count > 0 Then

    Dim cul As cList
    For Each cul In SelectLists.Items
        If cul.List > -1 And ConvertData(Text2, True) <> 0 Then
            ' Устанавливаем флаг изменения данных
            SetChange True
            List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) - ConvertData(Text2, True)
        End If
    Next

    Lapepic.Draw_Systems Lapepic.Picture1
    End If
ERR:
End Sub


