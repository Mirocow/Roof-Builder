VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D4055E62-5507-43CA-B528-924FB94C4FF4}#1.0#0"; "SplitterModern.ocx"
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form Lapemenu 
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10320
   ControlBox      =   0   'False
   Icon            =   "Lapemenu.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10320
   WindowState     =   2  'Maximized
   Begin SplitterHV.SplitHV SplitHV1 
      Height          =   6015
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   10610
      SplitLimit      =   8000
      Binding         =   1
      SplitWidth      =   80
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000C&
      Height          =   6975
      Left            =   4200
      TabIndex        =   4
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   2175
         Begin VB.CheckBox Check2 
            Caption         =   "Ñ ðàñêðîåì"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Value           =   2  'Grayed
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CheckBox ans 
            Caption         =   "Ïåðåçàïèñàòü áåç âîïðîñà"
            Enabled         =   0   'False
            Height          =   470
            Left            =   120
            TabIndex        =   6
            Top             =   1920
            Width           =   1935
         End
         Begin ÑhameleonButton.chameleonButton Check4 
            Height          =   615
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1085
            BTYPE           =   7
            TX              =   "Check4"
            ENAB            =   0   'False
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
            MICON           =   "Lapemenu.frx":030A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin ÑhameleonButton.chameleonButton Check3 
            Height          =   615
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1085
            BTYPE           =   7
            TX              =   "Check3"
            ENAB            =   0   'False
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
            MICON           =   "Lapemenu.frx":0326
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
            ICONS           =   16
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2535
         Left            =   720
         TabIndex        =   13
         Top             =   2640
         Width           =   2175
         Begin ÑhameleonButton.chameleonButton Command1 
            Height          =   495
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            BTYPE           =   7
            TX              =   "Calc"
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
            MICON           =   "Lapemenu.frx":0342
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
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            TabIndex        =   16
            Top             =   750
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton Check5 
            Caption         =   "2"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1935
         End
         Begin VB.OptionButton Check1 
            Caption         =   "1"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            TabIndex        =   14
            Top             =   1800
            Width           =   1935
         End
      End
      Begin MSComctlLib.ListView List2 
         Height          =   6015
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   32768
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Lapemenu.frx":035E
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView List1 
         Height          =   6015
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   32768
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Lapemenu.frx":04C0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   997
         EndProperty
      End
      Begin ÑhameleonButton.chameleonButton Command2 
         Height          =   495
         Left            =   720
         TabIndex        =   18
         Top             =   5280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BTYPE           =   7
         TX              =   "Óäàëèòü ðàñêðîé"
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
         MICON           =   "Lapemenu.frx":0622
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
      Begin ÑhameleonButton.chameleonButton Command5 
         Height          =   495
         Left            =   720
         TabIndex        =   19
         Top             =   5880
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BTYPE           =   7
         TX              =   "Óäàëèòü âñå"
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
         MICON           =   "Lapemenu.frx":063E
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
      Begin ÑhameleonButton.chameleonButton Command6 
         Height          =   495
         Left            =   720
         TabIndex        =   20
         Top             =   6480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BTYPE           =   7
         TX              =   "Îïèñàíèå"
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
         MICON           =   "Lapemenu.frx":065A
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
      Begin VB.Label Label1 
         Caption         =   "Ìåòêà1"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Ìåòêà2"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00DCFBFC&
      Height          =   6015
      Left            =   120
      MouseIcon       =   "Lapemenu.frx":0676
      ScaleHeight     =   5955
      ScaleWidth      =   3885
      TabIndex        =   3
      Top             =   240
      Width           =   3945
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   10215
      Begin VB.TextBox Label8 
         BackColor       =   &H00DCFBFC&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         MouseIcon       =   "Lapemenu.frx":07C8
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   120
         Width           =   10050
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Plain:"
      Height          =   195
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Lapemenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Select_Slope As Integer
Public option_p As String

Private Sub Check3_Click()
On Error Resume Next
    If Check3.value = False Then
'        Check3.Caption = lng.GetResIDstring(1150)
        Label8 = lng.GetResIDstring(1428)
    Else
'        Check3.Caption = lng.GetResIDstring(1151)
        Label8 = lng.GetResIDstring(1429)
    End If
End Sub


Private Sub Check4_Click()
On Error Resume Next
    If Check4.value = False Then
'        Check4.Caption = lng.GetResIDstring(1003)
        Label8 = lng.GetResIDstring(1428)
    Else
'        Check4.Caption = lng.GetResIDstring(1004)
        Label8 = lng.GetResIDstring(1429)
    End If
End Sub


Private Sub Command1_Click()
    Dim i As Integer
    Dim itmX As ListItem

    On Error GoTo ERR

    OfficeStart.MousePointer = 11
    OfficeStart.Enabled = False
'    Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(2).Picture

    If Me.Option1.value = True Then
        Plgs(LNC).Dll.UseAdditionalMethodCalc = False
    Else
        Plgs(LNC).Dll.UseAdditionalMethodCalc = True
    End If

    For i = 1 To Me.List1.ListItems.Count Step 1

        Set itmX = Me.List1.ListItems(i)
        If itmX.Selected = True Then
        
            Dim ProfilName As String
            If Gl.FileNameExtension = ".rbp" Then
                
                ProfilName = Trim(SlP(i).ProfilName)
                Factory_Name = Trim(SlP(i).Factory_Name)
                
            Else
            
                Factory_Name = Project.Label2.Caption
                ProfilName = Project.Label3.Caption
                
            End If
            
            If ProfilName <> "" And Factory_Name <> "" Then
                ' ÂÛßÂËÅÍÈÅ ÒÎ×ÅÊ ÏÎ ÂÅÐÒÈÊÀËÈ
                vert_hor_point i
                ' ÂÛßÂËÅÍÈÅ ÒÎ×ÅÊ ÏÎ ÃÎÐÈÇÎÍÒÀËÈ
                right_left_start_point i
                ' ÐÀÑ×ÅÒ
                mCalc i, ProfilName, Factory_Name
            End If

        End If

    Next

ExitCalc:
isSave = True

' Ïåðåðèñîâêà óæå ðàñ÷èòàííûõ
'fill_slope

'Module10.Draw_Systems Me.Picture3

'PleaseWait.CloseForm
'Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(1).Picture
OfficeStart.MousePointer = 99
OfficeStart.Enabled = True

Exit Sub
ERR:
'STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & Plgs(LNC).Dll.ERRDescription & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.17." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Sub


Private Sub Command2_Click()
    Dim slope_del As Integer
    On Error Resume Next
    
    For slope_del = 1 To List1.ListItems.Count
        If List1.ListItems(slope_del).Selected Then
            'Óäàëåíèå ðàñêðîÿ
            dell_snawing slope_del
            List1.ListItems(slope_del).Selected = False
        End If

    Next slope_del

    ' Ïåðåðèñîâêà óæå ðàñ÷èòàííûõ
    fill_slope
    Module10.Draw_Systems Me.Picture3
    
End Sub


Private Sub Command6_Click()
On Error Resume Next
    Teksti.Caption = lng.GetResIDstring(1074) & Me.List1.SelectedItem
    Teksti.Text1.Text = SlP(N_Slope).Describ
    Teksti.Show vbModal, OfficeStart
    SlP(N_Slope).Describ = Teksti.Text1.Text
    Unload Teksti
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Lapemenu.Visible = True Then
        Picture3.Height = Me.ScaleHeight - Frame2.Height - Label3.Height
        Frame2.Top = Picture3.Height + Label3.Height + 100
        frame1.Height = Frame2.Top
        Me.List1.Height = frame1.Height - 300 '- 650
        Me.List2.Height = frame1.Height - 300 '- 650
        SplitHV1.Height = Me.List1.Height '+ Label3.Height 'Picture3.Height + Label3.Height
        SplitHV1.ResizeControl

        Frame2.Width = Me.ScaleWidth
'        Shape2.Width = Frame2.Width - 2000
'        Label8.Width = Shape2.Width
        
            
'        Shape2.Width = Frame2.Width - 200 '- 3200
        Label8.Width = Frame2.Width - 200 '- 50
        
    End If
End Sub


Private Sub Label8_Click()
Load Teksti
Teksti.Text1 = Label8.Text
Teksti.Show vbModal, OfficeStart
Unload Teksti
End Sub


Private Sub List1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next

'    Command4.Enabled = True
    az = Item
    N_Slope = Item.Index
    OptionsS N_Slope
    Lapemenu.List2.ListItems.Clear
    fill_list False
End Sub


Private Sub OptionsS(n As Integer)
On Error Resume Next

    If SlP(n).CountSheets > 0 And SlP(n).CountOfPoints > 0 Then

        Check2.Enabled = True
        ans.Enabled = True
        Check5.Enabled = True
        Check1.Enabled = True
        Option1.Enabled = True

    ElseIf SlP(n).CountOfPoints > 0 Then
    
        Option1.Enabled = True
        Check5.Enabled = True
        Check1.Enabled = True

    Else
    
        Option1.Enabled = False
        Check1.Enabled = False
        Check5.Enabled = False
    
    End If

    If List2.ListItems.Count > 0 Then
        Check3.Enabled = True
        Check4.Enabled = True
        Check2.Enabled = True
        ans.Enabled = True
    Else
        Check3.Enabled = False
        Check4.Enabled = False
        Check2.Enabled = False
        ans.Enabled = False
    End If
  
End Sub




Private Sub List2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim l00BC As Single
    Dim N_Slope_Copy As Single
    Dim Ncopy As Integer
    Dim c
    Dim MIN As Single
    Dim MAX As Single
    Dim Centr As Single
    Dim n As Integer
    Dim N_list As Integer
    
    On Error Resume Next

      If option_p = "Copied" Then Exit Sub

      Label8 = ""

      If List1.SelectedItem.Index = 0 Then
          MsgBox (lng.GetResIDstring(1430)) '("Please selec lape for edit.")
          List1.SetFocus
          Exit Sub
      End If

      isSave = True

      If Asc(Item.Text) > 96 Then
          N_Slope_Copy = Asc(Item.Text) - 70
      Else
          N_Slope_Copy = Asc(Item.Text) - 64
      End If
  
      n = List1.ListItems.Count
      For Ncopy = 1 To n

          If List1.ListItems(Ncopy).Selected Then

'                N_Slope = Ncopy
              If SlP(Ncopy).CountOfPoints > 0 And ans.value = False Then
                  If l00BC <> 2 Then GoTo Msg
              End If

STC:

              If l00BC = 6 Or l00BC = 0 Then

                  ' Êîîðäèíàòû òî÷åê
                  SlP(Ncopy).CountOfPoints = SlP(N_Slope_Copy).CountOfPoints
                  For N_list = 1 To SlP(N_Slope_Copy).CountOfPoints Step 1
                      Lape_Points_X(Ncopy, N_list) = Lape_Points_X(N_Slope_Copy, N_list)
                      Lape_Points_Y(Ncopy, N_list) = Lape_Points_Y(N_Slope_Copy, N_list)
                  Next N_list

                  ' Ïðèâÿçêè
                  SlP(Ncopy).CountOfLines = SlP(N_Slope_Copy).CountOfLines
                  For N_list = 1 To SlP(N_Slope_Copy).CountOfLines Step 1
                      Lape_Lines(Ncopy, N_list, 0) = Lape_Lines(N_Slope_Copy, N_list, 0)
                      Lape_Lines(Ncopy, N_list, 1) = Lape_Lines(N_Slope_Copy, N_list, 1)
                  Next N_list

                  ' ÐÀÑÊÐÎÉ
                  If Check2.value Then
  
                      For N_list = 1 To SlP(N_Slope_Copy).CountSheets Step 1
                              List_Properties_PX(Ncopy, N_list) = List_Properties_PX(N_Slope_Copy, N_list)
                              List_Properties_PY(Ncopy, N_list) = List_Properties_PY(N_Slope_Copy, N_list)
                              List_Properties_Length(Ncopy, N_list) = List_Properties_Length(N_Slope_Copy, N_list)
                      Next N_list
  
                      SlP(Ncopy).Pn_StartLC = SlP(N_Slope_Copy).Pn_StartLC
                      SlP(Ncopy).PX_StartLC = SlP(N_Slope_Copy).PX_StartLC
                      
                  End If

                  '
  
                  If Check4.value Then  ' ÃÎÐÈÇÎÍÒÀËÜÍÎÅ ÊÎÏÈÐÎÂÀÍÈÅ

                      ' Çåðêàëüíûé ðèñóíîê ñêàòà
                      MIN = 99999
                      MAX = 0
  
                      For N_list = 1 To SlP(Ncopy).CountOfPoints Step 1
                          If Lape_Points_X(Ncopy, N_list) > MAX Then MAX = Lape_Points_X(Ncopy, N_list)
                          If Lape_Points_X(Ncopy, N_list) < MIN Then MIN = Lape_Points_X(Ncopy, N_list)
                      Next N_list
  
                      Centr = MIN + ((MAX - MIN) / 2)
                      For N_list = 1 To SlP(Ncopy).CountOfPoints Step 1
                          Lape_Points_X(Ncopy, N_list) = Centr - (Lape_Points_X(Ncopy, N_list) - Centr)
                      Next N_list
  
                      ' Çåðêàëüíûé ðàñêðîé
                      If Check2.value Then
                            
                          Dim ProfilName As String, FactoryName As String
                            
                          ProfilName = TrimNullChar(SlP(N_Slope_Copy).ProfilName)
                          FactoryName = TrimNullChar(SlP(N_Slope_Copy).Factory_Name)
                       
                          Dim CurrentPDataRS As Recordset
                          Set CurrentPDataRS = GetProfilData(ProfilName, GetFactoryID(FactoryName))
                          If Not CurrentPDataRS Is Nothing Then
                            
                          For N_list = 1 To SlP(N_Slope_Copy).CountSheets Step 1
                                  List_Properties_PX(Ncopy, N_list) = Centr - (List_Properties_PX(N_Slope_Copy, N_list) - _
                                  Centr + CurrentPDataRS![WORK_WIDTH]) 'Project.txtW)
                                  List_Properties_Length(Ncopy, N_list) = List_Properties_Length(N_Slope_Copy, N_list)
                          Next N_list
                          
                          End If
                          Set CurrentPDataRS = Nothing
  
                          SlP(Ncopy).PX_StartLC = SlP(N_Slope_Copy).PX_StartLC - MIN + MAX
                      End If
   
                  End If

                  If Check3.value Then  ' ÂÅÐÒÈÊÀËÜÍÎÅ ÊÎÏÈÐÎÂÀÍÈÅ

                      ' Çåðêàëüíûé ðèñóíîê ñêàòà
                      MIN = 99999
                      MAX = 0
  
                      For N_list = 1 To SlP(Ncopy).CountOfPoints Step 1
                          If Lape_Points_Y(Ncopy, N_list) > MAX Then MAX = Lape_Points_Y(Ncopy, N_list)
                          If Lape_Points_Y(Ncopy, N_list) < MIN Then MIN = Lape_Points_Y(Ncopy, N_list)
                      Next N_list
  
                      Centr = MIN + ((MAX - MIN) / 2)
                      For N_list = 1 To SlP(Ncopy).CountOfPoints Step 1
                          Lape_Points_Y(Ncopy, N_list) = Centr - Lape_Points_Y(Ncopy, N_list) + Centr
                      Next N_list
  
                      ' Çåðêàëüíûé ðàñêðîé
                      If Check2.value Then
                          For N_list = 1 To SlP(N_Slope_Copy).CountSheets Step 1
                              List_Properties_PY(Ncopy, N_list) = List_Properties_Length(N_Slope_Copy, N_list) + (Centr - List_Properties_PY(N_Slope_Copy, N_list) + Centr)
                          Next N_list
                      End If
  
                  End If

              End If
              
          SlP(Ncopy).Pn_Red_lines = SlP(N_Slope_Copy).Pn_Red_lines
          SlP(Ncopy).CountSheets = SlP(N_Slope_Copy).CountSheets
          SlP(Ncopy).Sf = SlP(N_Slope_Copy).Sf
          SlP(Ncopy).Sw = SlP(N_Slope_Copy).Sw
          SlP(Ncopy).ScaleHeightS = SlP(N_Slope_Copy).ScaleHeightS
          SlP(Ncopy).ScaleLeftS = SlP(N_Slope_Copy).ScaleLeftS
          SlP(Ncopy).ScaleTopS = SlP(N_Slope_Copy).ScaleTopS
          SlP(Ncopy).ScaleWidthS = SlP(N_Slope_Copy).ScaleWidthS
          SlP(Ncopy).ProfilName = SlP(N_Slope_Copy).ProfilName
          SlP(Ncopy).Factory_Name = SlP(N_Slope_Copy).Factory_Name

          End If

          List1.ListItems(Ncopy).Selected = False
      Next Ncopy
      
      Module10.Draw_Systems Me.Picture3

      Exit Sub
Msg:
      l00BC = MsgBox(lng.GetResIDstring(1431, "%LAPE%", List1.ListItems(Ncopy), "%LAPECOPY%", List2.SelectedItem.Text), vbYesNoCancel)
      GoTo STC
End Sub



Sub Command4_Click()

    Dim main_pic As Boolean
    Dim ans As Integer
    
    On Error GoTo err_size
    
        If KolvoScatov = 0 Then Exit Sub

        OfficeStart.menu_view.Visible = True

        main_pic = True
  
        If option_p = "" Or option_p = "Copy" Or option_p = "Mirrorcopy" Then
    
            If az = "" Then
                MsgBox (lng.GetResIDstring(1450))
                Exit Sub
            End If
        
        Else
            FlagDraw = 0
        End If
        
        If IsLoadForm("Lapepic") Then
            Unload Lapepic
            Load Lapepic
        Else
            Load Lapepic
        End If
  
        If SlP(N_Slope).Pn_Red_lines > 0 Then
            OptionDMM = "Msheet"
            Lapepic.sTabFx1.SelectTab 3
        Else
            
'            '#      Èíèöèàëèçàöèÿ îêíà
'            Lapepic.Picture1.ScaleLeft = -50 '0 '-50
'            Lapepic.Picture1.ScaleWidth = 1200 - 50
'            Lapepic.Picture1.ScaleTop = 1200 - 50 '625 '-50 '1200 '/ RatioH
'            Lapepic.Picture1.ScaleHeight = -50 '-600 '-625 '1200 '-Lapepic.Picture1.ScaleTop - 50
'             Module10.Change_scrol Lapepic.Picture1, Lapepic.HScroll1

'            Lapepic.Picture1.ScaleLeft = 0
'            Lapepic.Picture1.ScaleWidth = 1200
'            Lapepic.Picture1.ScaleTop = 600
'            Lapepic.Picture1.ScaleHeight = -600

'            Lapepic.Picture1.ScaleLeft = -625
'            Lapepic.Picture1.ScaleWidth = 625
'            Lapepic.Picture1.ScaleTop = 625
'            Lapepic.Picture1.ScaleHeight = -625

'            Lapepic.Picture1.ScaleLeft = -100
'            Lapepic.Picture1.ScaleWidth = 1100 '6400 'ConvertData(12, False)
'            Lapepic.Picture1.ScaleTop = 600 'ConvertData(6, False)
'            Lapepic.Picture1.ScaleHeight = -600 '-ConvertData(6, False)
'
'            Lapepic.SuperRuler1.MaxH = Lapepic.HScroll1.MAX '* 10
'            Lapepic.SuperRuler2.MaxV = Lapepic.HScroll1.MAX '* 10
        
            If SlP(N_Slope).CountOfLines > 0 Then
                Lapepic.sTabFx1.SelectTab 1
            Else
'                Lapepic.HScroll1.value = 1000
                Lapepic.sTabFx1.SelectTab 0
            End If

        End If

'        If Lapepic.Error = True Then Exit Sub
        
        Lapepic.Show
        Lapepic.Command10.Caption = az
        OfficeStart.TabStrip1.Tabs(4).Selected = True
        Unload Me
        Exit Sub
        
err_size:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.18." & ERR.Source & "]", ERR.Number, ERR.Description
        Resume Next
End Sub


Private Sub Command5_Click()
    Dim slope_del As Integer
    Dim i As Integer
    
    On Error Resume Next
    
        isSave = True

        For slope_del = 1 To List1.ListItems.Count
    
            If List1.ListItems(slope_del).Selected Then
        
                ' Óäàëåíèå ÷åðòåæà
                For i = 1 To SlP(slope_del).CountOfPoints Step 1
                    Lape_Points_X(slope_del, i) = 0
                    Lape_Points_Y(slope_del, i) = 0
                Next i

                SlP(slope_del).CountOfPoints = 0
        
                For i = 1 To SlP(slope_del).CountOfLines Step 1
                    Lape_Lines(slope_del, i, 0) = 0
                    Lape_Lines(slope_del, i, 1) = 0
                Next i

                SlP(slope_del).CountOfLines = 0
        
                '        Óäàëåíèå ïëîùàäè ôèãóðû
                SlP(slope_del).Sf = 0
                SlP(slope_del).Sw = 0
        
                '       Óäàëåíèå ðàñêðîÿ
                dell_snawing slope_del

                List1.ListItems(slope_del).Selected = False
                
                SlP(slope_del).Factory_Name = ""
                SlP(slope_del).ProfilName = ""
                SlP(slope_del).Describ = ""
                
                SlP(slope_del).ScaleLeftS = 0
                SlP(slope_del).ScaleWidthS = 0
                SlP(slope_del).ScaleTopS = 0
                SlP(slope_del).ScaleHeightS = 0
        
            End If
    
        Next slope_del

        ' Ïåðåðèñîâêà óæå ðàñ÷èòàííûõ
        fill_slope
        Module10.Draw_Systems Me.Picture3
End Sub


Sub dell_snawing(slope_del As Integer)
    On Error Resume Next

    Dim i As Integer

    For i = 1 To SlP(slope_del).CountSheets Step 1
            List_Properties_PY(slope_del, i) = 0
            List_Properties_PX(slope_del, i) = 0
            List_Properties_Length(slope_del, i) = 0
    Next i
    
    SlP(slope_del).CountSheets = 0
    SlP(slope_del).Pn_Red_lines = 0
    SlP(slope_del).ListLength = 0
    SlP(slope_del).Pn_StartLC = 0
    
'    SlP(slope_del).Describ = ""
'    SlP(slope_del).ScaleLeftS = 0
'    SlP(slope_del).ScaleWidthS = 0
'    SlP(slope_del).ScaleTopS = 0
'    SlP(slope_del).ScaleHeightS = 0

End Sub


Sub fill_slope()
    Dim i As Integer
    
    On Error Resume Next

    Lapemenu.List1.ListItems.Clear
    Lapemenu.List2.ListItems.Clear

    fill_list True
    
    ' Åñëè íåò îáîçíà÷åíèé ñêàòîâ
    If Lapemenu.List1.ListItems.Count = 0 Then
'        Command4.Enabled = False
        Command2.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
    Else
        Command6.Enabled = True
'        Command4.Enabled = True
'            Command1.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
        Label8.Text = lng.GetResIDstring(1009)
    End If

End Sub


Sub fill_list(is_fil_list1 As Boolean)
    Dim i As Integer
    Dim itmX As ListItem, itm1X As ListItem
    Dim letter As String
    
    On Error GoTo ERR
    
    For i = 1 To KolvoScatov Step 1
    
        If i > 26 Then
        
            letter = Chr$(i + 70)
            
            If is_fil_list1 Then Set itmX = List1.ListItems.Add(, , letter)
            
            If SlP(i).CountOfPoints > 0 Then Set itm1X = List2.ListItems.Add(, , letter)
        
            If SlP(i).CountSheets > 0 Then
                
                itm1X.ForeColor = vbBlue
                itm1X.SubItems(1) = lng.GetResIDstring(9639)
                itm1X.ListSubItems(1).ForeColor = vbBlue
                itm1X.SubItems(2) = Trim(SlP(i).ProfilName)
                itm1X.ListSubItems(2).ForeColor = vbBlue
                itm1X.SubItems(3) = Trim(SlP(i).Factory_Name)
                itm1X.ListSubItems(3).ForeColor = vbBlue
                
            Else
            
                If SlP(i).CountOfPoints > 0 Then
                    itm1X.ForeColor = vbRed
                    itm1X.SubItems(1) = lng.GetResIDstring(9638)
                    itm1X.ListSubItems(1).ForeColor = vbRed
                End If
                
            End If
    
        Else
        
            letter = Chr$(i + 64)
            
            If is_fil_list1 Then Set itmX = List1.ListItems.Add(, , letter)
        
            If SlP(i).CountOfPoints > 0 Then Set itm1X = List2.ListItems.Add(, , letter)
            
            If SlP(i).CountSheets > 0 Then
            
                itm1X.ForeColor = vbBlue
                itm1X.SubItems(1) = lng.GetResIDstring(9639)
                itm1X.ListSubItems(1).ForeColor = vbBlue
                itm1X.SubItems(2) = Trim(SlP(i).ProfilName)
                itm1X.ListSubItems(2).ForeColor = vbBlue
                itm1X.SubItems(3) = Trim(SlP(i).Factory_Name)
                itm1X.ListSubItems(3).ForeColor = vbBlue
                
            Else
            
                If SlP(i).CountOfPoints > 0 Then
                    itm1X.ForeColor = vbRed
                    itm1X.SubItems(1) = lng.GetResIDstring(9638)
                    itm1X.ListSubItems(1).ForeColor = vbRed
                End If
                
            End If
        
        End If
    
    Next i
    
    Set itmX = Nothing
    Set itm1X = Nothing
    Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.19." & ERR.Source & "]", ERR.Number, ERR.Description
    'Resume Next
End Sub


Private Sub Form_Load()

    On Error GoTo ERR

    SetFont Me

'    Command4.Caption = lng.GetResIDstring(9168)
'    Command3.Caption = lng.GetResIDstring(9169)
    ans.Caption = lng.GetResIDstring(9172)
    Check2.Caption = lng.GetResIDstring(9173)
    Command2.Caption = lng.GetResIDstring(9175)
    Command1.Caption = lng.GetResIDstring(9176)
    Command5.Caption = lng.GetResIDstring(9177)

    Lapemenu.Label1.Caption = lng.GetResIDstring(1001)
    Lapemenu.Label2.Caption = lng.GetResIDstring(1002)

    Check5.Caption = lng.GetResIDstring(9629)
    Check1.Caption = lng.GetResIDstring(9203)
    Option1.Caption = lng.GetResIDstring(9204)

    Check4.value = False
    Check3.value = False

    Command1.Caption = lng.GetResIDstring(9008)

'    Lapemenu.Command4.Caption = lng.GetResIDstring(1005)
    Lapemenu.Label3 = lng.GetResIDstring(1006)
    Lapemenu.Label8 = lng.GetResIDstring(1009)
    
'    Command4.Caption = lng.GetResIDstring(3001)
'    Command3.Caption = lng.GetResIDstring(3003)

    Label8.ToolTipText = lng.GetResIDstring(1123)

    Command6.Caption = lng.GetResIDstring(9572) 'Lng.GetResIDstring(1074)
    
    Picture3.BackColor = Setup.Command9.BackColor

    Check3.Caption = lng.GetResIDstring(1151)
    Check4.Caption = lng.GetResIDstring(1004)

    Set SplitHV1.obj1 = Picture3
    Set SplitHV1.obj1 = Label3
    Set SplitHV1.obj2 = frame1

    fill_slope

    Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.20." & ERR.Source & "]", ERR.Number, ERR.Description
    Resume Next
End Sub

Private Sub List1_DblClick()
    Command4_Click
End Sub


Private Sub List1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 27 Then Command3.value = True
    If KeyAscii = 13 Then Command4_Click
End Sub


Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Select_Slope <> 0 Then
    If Select_Slope > 26 Then
        az = Chr$(Select_Slope + 70)
    Else
        az = Chr$(Select_Slope + 64)
    End If
    
    N_Slope = Select_Slope

    Lapemenu.fill_slope
    'Lapemenu.List1.ListItems(Select_Slope).Selected = True
    Command4_Click

    Lapepic.SetDrawBorder
'    Lapepic.SavePolygon False

    Lapepic.Command5.value = True
End If
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If isActive.isFormFocus(Me.hwnd) Then
     isActive.SetFormFocus Picture3.hwnd
    End If

    Select_Slope = ROOFPIC.Find_lape_label(X, Y)
    If Select_Slope <> 0 Then
        Picture3.MousePointer = 99
    Else
        Picture3.MousePointer = 0
    End If

End Sub

Private Sub Picture3_Resize()
    On Error Resume Next
    Me.Picture3.ScaleLeft = ScaleLeft_Main 'ROOFPIC.Picture1.ScaleLeft
    Me.Picture3.ScaleTop = ScaleTop_Main 'ROOFPIC.Picture1.ScaleTop
    Me.Picture3.ScaleWidth = ScaleWidth_Main 'ROOFPIC.Picture1.ScaleWidth
    Me.Picture3.ScaleHeight = ScaleHeight_Main 'ROOFPIC.Picture1.ScaleHeight
    Module10.Draw_Systems Me.Picture3
End Sub

Private Sub SplitHV1_MoveEnd()
    On Error Resume Next
    List2.Width = frame1.Width - 3100
End Sub


Public Sub mCalc(CurrentSlope As Integer, ProfilName As String, FactoryName As String)
    Dim NERROR As Integer
    On Error GoTo ERR


    ' Ñîðòèðîâêà òî÷åê
    Dim isSort As Boolean
    Dim Lape_Lines_out() As Integer ' Ñâîéñòâà ëèíèé
    ReDim Lape_Lines_out(1 To MAXSLOPES, 1 To MAXSLOPELINE + 2, 1)
    Dim CountOfLines_out As Integer
    Dim CountOfPoints_out As Integer
    
    isSort = PointsSort(Lape_Lines, Lape_Lines_out, CurrentSlope, SlP(CurrentSlope).CountOfLines, SlP(CurrentSlope).CountOfPoints, CountOfLines_out, CountOfPoints_out)
    
    If isSort = False Then Exit Sub
    
    SlP(CurrentSlope).Sf = PolygonArea(CountOfLines_out, CurrentSlope, Lape_Lines_out, Lape_Points_X, Lape_Points_Y)

    ' Î÷èñòêà
    Dim L As Integer
    For L = 1 To MAXSLOPELISTS
        List_Properties_PY(CurrentSlope, L) = 0
        List_Properties_PX(CurrentSlope, L) = 0
        List_Properties_Length(CurrentSlope, L) = 0
    Next

    Dim arr_wd() As Long, arr_am() As Long
    
    If Check5.value Then
           
        'GetWarehouseLength ProfilName, False, arr_wd, arr_am
    
    ElseIf Check1.value = True Then

    Dim PDataRS As Recordset
    Set PDataRS = RequestSQL("select * from ProfilsWrongLength p where p.idname=" & GetProfilID(ProfilName) & " order by length")
    If Not PDataRS Is Nothing Then
        Do While Not PDataRS.EOF
            Dim ii As Integer
            ReDim Preserve arr_wd(ii)
            arr_wd(ii) = PDataRS.Fields(2)
            ReDim Preserve arr_am(ii)
            arr_am(ii) = 10
            PDataRS.MoveNext
            ii = ii + 1
        Loop
        PDataRS.Close
    End If

    Set PDataRS = Nothing
        
    End If

If Check5.value = True Or Check1.value = True Then
    On Error GoTo arrclear
    If UBound(arr_wd) Then
    Else
arrclear:

        If Check1.value = True Then
            MsgBox lng.GetResIDstring(1488, "%ProfilName%", ProfilName), vbCritical
        ElseIf Check5.value = True Then
            MsgBox lng.GetResIDstring(1489, "%ProfilName%", ProfilName), vbCritical
        End If

        Exit Sub
    End If

End If

    Plgs(Gl.LNC).Dll.InputWarehouseData arr_wd, arr_am
   
    Dim CurrentPDataRS As Recordset
    Set CurrentPDataRS = GetProfilData(ProfilName, GetFactoryID(FactoryName))

    If CurrentPDataRS!Width <> 0 Or CurrentPDataRS![WORK_WIDTH] <> 0 Then
    
    Dim ans As Integer
    ans = Plgs(LNC).Dll.SetMaterial _
    (CurrentPDataRS!Width, CurrentPDataRS![WORK_WIDTH], CurrentPDataRS!Step, CurrentPDataRS!Overlaping, _
    CurrentPDataRS!MIN_LENGTH, CurrentPDataRS!MAX_LENGTH, CurrentPDataRS!Height)
    Set CurrentPDataRS = Nothing
    
    NERROR = Plgs(LNC).Dll.calc _
    (Lape_Lines_out, CurrentSlope, CountOfLines_out, CountOfPoints_out, _
    Lape_Points_X, Lape_Points_Y, SlP(CurrentSlope).PX_StartLC, Lape_Points_Y(CurrentSlope, SlP(CurrentSlope).Pn_Red_lines), SlP(CurrentSlope).CountSheets, _
    List_Properties_PX, List_Properties_PY, 0, List_Properties_Length, 0)
    

    If SlP(CurrentSlope).CountSheets = 0 Then
        'SlP(CurrentSlope).ProfilName = ""
        'SlP(CurrentSlope).Factory_Name = ""
    Else
        SlP(CurrentSlope).ProfilName = ProfilName
        SlP(CurrentSlope).Factory_Name = FactoryName
    End If

    End If

    Exit Sub
ERR:
'    STRERR = STRERR & time & ". ( mcalc = " & NERROR & ") ... [ERROR] N " & ERR.Number & " (" & Plgs(LNC).Dll.ERRDescription & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.21." & ERR.Source & "]", ERR.Number, ERR.Description
    On Error Resume Next
End Sub

