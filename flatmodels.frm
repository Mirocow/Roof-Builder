VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form flatmodels 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flat models"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11205
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   4080
      TabIndex        =   23
      Top             =   5180
      Width           =   5295
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808000&
         Caption         =   "Показывать длины сторон"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   240
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "Перемещать точку"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Перемещать сторону"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   950
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   6375
      Width           =   5280
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Lape control"
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   9480
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
      Begin VB.TextBox y2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   21
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox x2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   16
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox y1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox x1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Измените эти данные до нужного размера, затем нажмите ENTER."
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Y2="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "X2="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Y1="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "X1="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   300
      Left            =   4080
      Max             =   10000
      Min             =   100
      SmallChange     =   200
      TabIndex        =   10
      Top             =   4970
      Value           =   1000
      Width           =   5295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DCFBFC&
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   4080
      ScaleHeight     =   4305
      ScaleWidth      =   5265
      TabIndex        =   9
      Top             =   600
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   " Base control "
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   9480
      TabIndex        =   3
      Top             =   600
      Width           =   1575
      Begin VB.CommandButton Command5 
         Caption         =   "Load Base"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Re Save Base"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Dell select"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Take out from Base "
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add to Base"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exit"
      Height          =   975
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   10610
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14482428
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "cod"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Additional information"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "N"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ScaleLeft"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ScaleWidth"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ScaleTop"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ScaleHeight"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   11359
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   975
      Left            =   4080
      Top             =   6360
      Width           =   5295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Manager of slopes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "flatmodels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type EmployeeRecord ' Тип для хранения временных шаблонов рисунков
    RN As Integer
    CN As Integer
    LPXf As Single
    LPXs As Single
    LPYf As Single
    LPYs As Single
    PXYSt As Single
    PXYFi As Single
    info As String * MAXL
    ScaleLeft As Single
    ScaleWidth As Single
    ScaleTop As Single
    ScaleHeight As Single
    Description As String * 500
End Type

'# models
Private Line_PX_M(100, MAXP + 2) As Single
Private Line_PY_M(100, MAXP + 2) As Single
Private ApConnect_M(100, MAXP + 2) As Integer
Private BpConnect_M(100, MAXP + 2) As Integer

Dim Findline As Integer

Private Sub Check1_Click()
Draw ListView1.SelectedItem.Index
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim l01B4 As Integer
Dim i As Integer
i = ListView1.SelectedItem - 1

If Not SlP(N_Slope).ApCount > 0 Then

For l01B4 = 0 To ListView1.ListItems.Item(ListView1.SelectedItem.Index).ListSubItems(2) + 1 Step 1
    
    If ApConnect_M(i, l01B4) <> 0 Then ApConnect(N_Slope, l01B4 + 1) = ApConnect_M(i, l01B4)
    If BpConnect_M(i, l01B4) <> 0 Then BpConnect(N_Slope, l01B4 + 1) = BpConnect_M(i, l01B4)
    
    If Line_PX_M(i, ApConnect_M(i, l01B4)) <> 0 And Line_PX_M(i, BpConnect_M(i, l01B4)) <> 0 Then
    Lape_Points_X(N_Slope, ApConnect(N_Slope, l01B4 + 1)) = Line_PX_M(i, ApConnect_M(i, l01B4))
    Lape_Points_Y(N_Slope, ApConnect(N_Slope, l01B4 + 1)) = Line_PY_M(i, ApConnect_M(i, l01B4))
    
    Lape_Points_X(N_Slope, BpConnect(N_Slope, l01B4 + 1)) = Line_PX_M(i, BpConnect_M(i, l01B4))
    Lape_Points_Y(N_Slope, BpConnect(N_Slope, l01B4 + 1)) = Line_PY_M(i, BpConnect_M(i, l01B4))
    
    End If
    
Next l01B4

    SlP(N_Slope).ApCount = ListView1.ListItems.Item(ListView1.SelectedItem.Index).ListSubItems(2)
    SlP(N_Slope).BpCount = ListView1.ListItems.Item(ListView1.SelectedItem.Index).ListSubItems(2)
     
    SlP(N_Slope).ScaleLeftS = ListView1.ListItems.Item(ListView1.SelectedItem.Index).ListSubItems(3)
    SlP(N_Slope).ScaleWidthS = ListView1.ListItems.Item(ListView1.SelectedItem.Index).ListSubItems(4)
    SlP(N_Slope).ScaleTopS = ListView1.ListItems.Item(ListView1.SelectedItem.Index).ListSubItems(5)
    SlP(N_Slope).ScaleHeightS = ListView1.ListItems.Item(ListView1.SelectedItem.Index).ListSubItems(6)
    
'Lapepic.Draw_Systems Lapepic
Lapepic.Command5.Value = True
'Lapepic.Option3.value = True
'Else
'MsgBox "It is impossible to insert a pattern as the edited slope is not empty."
End If
End Sub

Private Sub Command3_Click()
Dim l01B4 As Integer

For l01B4 = 1 To SlP(N_Slope).BpCount Step 1

    ApConnect_M(ListView1.ListItems.count, l01B4 - 1) = ApConnect(N_Slope, l01B4)
    BpConnect_M(ListView1.ListItems.count, l01B4 - 1) = BpConnect(N_Slope, l01B4)
    
    Line_PX_M(ListView1.ListItems.count, ApConnect_M(ListView1.ListItems.count, l01B4 - 1)) = Lape_Points_X(N_Slope, ApConnect(N_Slope, l01B4))
    Line_PY_M(ListView1.ListItems.count, ApConnect_M(ListView1.ListItems.count, l01B4 - 1)) = Lape_Points_Y(N_Slope, ApConnect(N_Slope, l01B4))
    
    Line_PX_M(ListView1.ListItems.count, BpConnect_M(ListView1.ListItems.count, l01B4 - 1)) = Lape_Points_X(N_Slope, BpConnect(N_Slope, l01B4))
    Line_PY_M(ListView1.ListItems.count, BpConnect_M(ListView1.ListItems.count, l01B4 - 1)) = Lape_Points_Y(N_Slope, BpConnect(N_Slope, l01B4))
    
Next l01B4

'Line_PX_M(ListView1.ListItems.count + 1, 0) = l01B4 - 1

Set itmX = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)
itmX.SubItems(1) = N_Slope & " from " & Gl.CurrentFile
itmX.SubItems(2) = SlP(N_Slope).BpCount

    itmX.SubItems(3) = SlP(N_Slope).ScaleLeftS
    itmX.SubItems(4) = SlP(N_Slope).ScaleWidthS
    itmX.SubItems(5) = SlP(N_Slope).ScaleTopS
    itmX.SubItems(6) = SlP(N_Slope).ScaleHeightS
    itmX.SubItems(7) = ""
    
Text2 = ""
If ListView1.ListItems.count <> 0 Then ListView1.Enabled = True
End Sub

Private Sub Command5_Click()
Dim MyRecord As EmployeeRecord
Dim recn As Integer
Dim i As Integer
Dim n As Integer
Dim FileN As Integer
Dim count As Integer
ListView1.ListItems.Clear
FileN = FreeFile
On Error GoTo ERR
Open App.Path & "\data\base.dat" For Random As #FileN Len = Len(MyRecord)
    Do While Not EOF(FileN)
        recn = recn + 1: 'i = i + 1
        
        Get #FileN, recn, MyRecord
'        If count = 0 Then count = MyRecord.RN
                For n = 1 To MyRecord.CN + 1 ' забивка линий
                
                    Line_PX_M(i, n) = MyRecord.LPXf
                    Line_PX_M(i, n) = MyRecord.LPXs
                    Line_PY_M(i, n) = MyRecord.LPYf
                    Line_PY_M(i, n) = MyRecord.LPYs
                    
                    ApConnect_M(i, n - 1) = MyRecord.PXYSt
                    BpConnect_M(i, n - 1) = MyRecord.PXYFi
                    
                    recn = recn + 1
                    Get #FileN, recn, MyRecord
                Next
                
        If MyRecord.RN <> 0 Then
        
'                    Line_PX_M(i, n) = MyRecord.LPXf
'                    Line_PX_M(i, n) = MyRecord.LPXs
'                    Line_PY_M(i, n) = MyRecord.LPYf
'                    Line_PY_M(i, n) = MyRecord.LPYs
                    
'                    ApConnect_M(i, n - 1) = MyRecord.PXYSt
'                    BpConnect_M(i, n - 1) = MyRecord.PXYFi
        
        Set itmX = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)
        
        itmX.SubItems(1) = MyRecord.info
        
        itmX.SubItems(2) = MyRecord.CN
        itmX.SubItems(3) = MyRecord.ScaleLeft
        itmX.SubItems(4) = MyRecord.ScaleWidth
        itmX.SubItems(5) = MyRecord.ScaleTop
        itmX.SubItems(6) = MyRecord.ScaleHeight
        
        itmX.SubItems(7) = MyRecord.Description

        End If
    i = i + 1
    Loop
'Close #FileN
If ListView1.ListItems.count <> 0 Then ListView1.Enabled = True
ERR:
STRERROR = STRERROR & Time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
Close #FileN
End Sub

Private Sub Command6_Click()
Dim i As Integer
Dim MyRecord As EmployeeRecord
Dim FileN As Integer
FileN = FreeFile

If dir(App.Path & "\data\base.dat") <> "" Then Kill App.Path & "\data\base.dat"
Open App.Path & "\data\base.dat" For Random As #FileN Len = Len(MyRecord)

For i = 0 To ListView1.ListItems.count - 1
    
    If Line_PX_M(i, 1) <> 0 Then
    
    For n = 0 To ListView1.ListItems.Item(i + 1).ListSubItems(2)
        
        MyRecord.PXYSt = ApConnect_M(i, n)
        MyRecord.PXYFi = BpConnect_M(i, n)
        
        MyRecord.RN = ListView1.ListItems.count ' количество фигур (записей)
        MyRecord.CN = ListView1.ListItems.Item(i + 1).ListSubItems(2)  ' количество линий (информация о точках)
        
        MyRecord.LPXf = Line_PX_M(i, n + 1)
        MyRecord.LPXs = Line_PX_M(i, n + 1)
        MyRecord.LPYf = Line_PY_M(i, n + 1)
        MyRecord.LPYs = Line_PY_M(i, n + 1)
               
    Put #FileN, , MyRecord

    
    Next
    
    MyRecord.info = ListView1.ListItems.Item(i + 1).ListSubItems(1)
    MyRecord.ScaleLeft = ListView1.ListItems.Item(i + 1).ListSubItems(3)
    MyRecord.ScaleWidth = ListView1.ListItems.Item(i + 1).ListSubItems(4)
    MyRecord.ScaleTop = ListView1.ListItems.Item(i + 1).ListSubItems(5)
    MyRecord.ScaleHeight = ListView1.ListItems.Item(i + 1).ListSubItems(6)
    MyRecord.Description = ListView1.ListItems.Item(i + 1).ListSubItems(7)
    
    Put #FileN, , MyRecord
'    Else
'    i = i - 1
    End If
    
Next
Close #FileN
'MyRecord = Nothing
End Sub

Private Sub Command7_Click()
Dim i As Integer
Dim l01B4 As Integer

If ListView1.ListItems.count = 0 Then Exit Sub

i = ListView1.SelectedItem.Index - 1

For l01B4 = 0 To ListView1.ListItems.Item(i + 1).ListSubItems(2) - 1 Step 1

Line_PX_M(i, ApConnect_M(i, l01B4)) = 0
Line_PX_M(i, BpConnect_M(i, l01B4)) = 0
Line_PY_M(i, ApConnect_M(i, l01B4)) = 0
Line_PY_M(i, BpConnect_M(i, l01B4)) = 0

BpConnect_M(i, l01B4) = 0
ApConnect_M(i, l01B4) = 0

Next l01B4


Command6.Value = True
ListView1.ListItems.Remove ListView1.SelectedItem.Index
Picture1.Cls

Command5.Value = True
End Sub

Private Sub Form_Activate()
If SlP(N_Slope).ApCount <> 0 Then
Command2.Enabled = False
Command3.Enabled = True
Else
If flagimg4 Then Command2.Enabled = True
Command3.Enabled = False
End If
If ListView1.ListItems.count = 0 Then
ListView1.Enabled = False
Picture1.Cls
'List1.Clear
End If
End Sub

Private Sub Form_Load()
Label3 = ResolveResstring(1322)
Label8 = ResolveResstring(1473)
Me.Caption = Label3
Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
Me.top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.top)
'Me.Width = GetSetting(App.ProductName, "Position", Me.name & "width", Me.Width)
'Me.Height = GetSetting(App.ProductName, "Position", Me.name & "height", Me.Height)
Picture1.FontSize = 10
Picture1.ForeColor = vbRed
Command5_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
SaveSetting App.ProductName, "Position", Me.name & "top", Me.top
'SaveSetting App.ProductName, "Position", Me.name & "width", Me.Width
'SaveSetting App.ProductName, "Position", Me.name & "height", Me.Height
Command6_Click
End Sub

Private Sub List1_Click()
Draw ListView1.SelectedItem
End Sub

Private Sub HScroll1_Change()
HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
Dim n As Integer
n = ListView1.SelectedItem.Index
Me.Picture1.ScaleLeft = ListView1.ListItems.Item(n).ListSubItems(3)
Me.Picture1.ScaleWidth = ListView1.ListItems.Item(n).ListSubItems(4)
Me.Picture1.ScaleTop = ListView1.ListItems.Item(n).ListSubItems(5)
Me.Picture1.ScaleHeight = ListView1.ListItems.Item(n).ListSubItems(6)
Module10.Change_scrol Me.Picture1, Me.HScroll1
Draw n
End Sub

Private Sub ListView1_DblClick()
ListView1.SelectedItem.ListSubItems(7) = InputBox("Set description", , ListView1.SelectedItem.ListSubItems(7))
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Me.List1.Clear
Label1 = ListView1.SelectedItem.ListSubItems(7)

'For l01B4 = 1 To Item.ListSubItems(2)
'    Me.List1.AddItem l01B4
'Next l01B4

'Draw Item

HScroll1.Value = HScroll1.Max / 3
HScroll1_Scroll
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NSlope As Integer

If Option1.Value = True Then
x1.Enabled = True
y1.Enabled = True
x2.Enabled = True
y2.Enabled = True
Else
Findline = 0
x1.Enabled = False
y1.Enabled = False
x2.Enabled = False
y2.Enabled = False
End If

Findline = DrawFindLine(X, Y) ' Поиск линии

If Findline <> 0 Then

NSlope = ListView1.SelectedItem.Index - 1
x1 = Line_PX_M(NSlope, ApConnect_M(NSlope, Findline))
y1 = Line_PY_M(NSlope, ApConnect_M(NSlope, Findline))
x2 = Line_PX_M(NSlope, BpConnect_M(NSlope, Findline))
y2 = Line_PY_M(NSlope, BpConnect_M(NSlope, Findline))
Draw ListView1.SelectedItem

ElseIf Option2.Value = True Then

' Поиск точки

End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim NSlope As Integer

If Option1.Value = True Then
NSlope = ListView1.SelectedItem.Index - 1
Line_PX_M(NSlope, ApConnect_M(NSlope, Findline)) = X
Line_PY_M(NSlope, ApConnect_M(NSlope, Findline)) = Y

x1 = X
y1 = Y

Line_PX_M(NSlope, BpConnect_M(NSlope, Findline)) = Line_PX_M(NSlope, BpConnect_M(NSlope, Findline)) + x2
Line_PY_M(NSlope, BpConnect_M(NSlope, Findline)) = Line_PY_M(NSlope, BpConnect_M(NSlope, Findline)) + y2

x2 = Line_PX_M(NSlope, BpConnect_M(NSlope, Findline))
y2 = Line_PY_M(NSlope, BpConnect_M(NSlope, Findline))

Draw ListView1.SelectedItem
Else

Draw ListView1.SelectedItem
End If
End If
End Sub

Private Sub Picture1_Paint()
'Lapepic.Draw_Systems Me
End Sub

Sub Draw(ByVal i As Integer)
Dim l01B4 As Integer
Dim n As Single
n = ListView1.SelectedItem.Index
Picture1.Cls
Picture1.DrawStyle = 6
Picture1.DrawMode = 13

If ListView1.ListItems.count < 1 Then Exit Sub
i = i - 1
For l01B4 = 0 To ListView1.ListItems.Item(n).ListSubItems(2) - 1 Step 1

If BpConnect_M(i, l01B4) <= ListView1.ListItems.Item(n).ListSubItems(2) And BpConnect_M(i, l01B4) <> 0 Then


If Findline = l01B4 Then
Me.Picture1.Line (Line_PX_M(i, ApConnect_M(i, l01B4)), Line_PY_M(i, ApConnect_M(i, l01B4)))-(Line_PX_M(i, BpConnect_M(i, l01B4)), Line_PY_M(i, BpConnect_M(i, l01B4))), vbRed
Else
Me.Picture1.Line (Line_PX_M(i, ApConnect_M(i, l01B4)), Line_PY_M(i, ApConnect_M(i, l01B4)))-(Line_PX_M(i, BpConnect_M(i, l01B4)), Line_PY_M(i, BpConnect_M(i, l01B4))), vbBlack
End If

If Check1.Value = 1 Then
Llen = Format(Sqr((Line_PX_M(i, BpConnect_M(i, l01B4)) - Line_PX_M(i, ApConnect_M(i, l01B4))) ^ 2 + _
  (Line_PY_M(i, BpConnect_M(i, l01B4)) - Line_PY_M(i, ApConnect_M(i, l01B4))) ^ 2), "###.0")
Picture1.PSet ((Line_PX_M(i, BpConnect_M(i, l01B4)) + Line_PX_M(i, ApConnect_M(i, l01B4))) / 2, _
(Line_PY_M(i, BpConnect_M(i, l01B4)) + Line_PY_M(i, ApConnect_M(i, l01B4))) / 2), "&H00C0C0C0"
Picture1.Print Llen

End If

End If

Next l01B4
End Sub


Private Sub Text2_LostFocus()
ListView1.SelectedItem.ListSubItems(7) = Text2.Text
End Sub

Function DrawFindLine(X, Y)
Dim i As Integer
Dim AC As Single
Dim BC As Single
Dim N_Slope As Integer
Dim n As Integer

N_Slope = ListView1.SelectedItem.Index - 1

For i = 0 To ListView1.SelectedItem.SubItems(2)

AC = Sqr((Line_PX_M(N_Slope, ApConnect_M(N_Slope, i)) - X) ^ 2 + (Line_PY_M(N_Slope, ApConnect_M(N_Slope, i)) - Y) ^ 2)
BC = Sqr((Line_PX_M(N_Slope, BpConnect_M(N_Slope, i)) - X) ^ 2 + (Line_PY_M(N_Slope, BpConnect_M(N_Slope, i)) - Y) ^ 2)

If CInt(AC + BC) = CInt(Sqr((Line_PX_M(N_Slope, BpConnect_M(N_Slope, i)) - Line_PX_M(N_Slope, ApConnect_M(N_Slope, i))) ^ 2 + _
  (Line_PY_M(N_Slope, BpConnect_M(N_Slope, i)) - Line_PY_M(N_Slope, ApConnect_M(N_Slope, i))) ^ 2)) Then
n = i
Exit For
End If
Next i
DrawFindLine = n
End Function

Private Sub x1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim N_Slope As Integer

N_Slope = ListView1.SelectedItem.Index - 1
Line_PX_M(N_Slope, ApConnect_M(N_Slope, Findline)) = x1
Draw ListView1.SelectedItem
End If
End Sub

Private Sub y1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim N_Slope As Integer

N_Slope = ListView1.SelectedItem.Index - 1
Line_PY_M(N_Slope, ApConnect_M(N_Slope, Findline)) = y1
Draw ListView1.SelectedItem
End If
End Sub

Private Sub x2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim N_Slope As Integer

N_Slope = ListView1.SelectedItem.Index - 1
Line_PX_M(N_Slope, BpConnect_M(N_Slope, Findline)) = x2
Draw ListView1.SelectedItem
End If
End Sub

Private Sub y2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim N_Slope As Integer

N_Slope = ListView1.SelectedItem.Index - 1
Line_PY_M(N_Slope, BpConnect_M(N_Slope, Findline)) = y2
Draw ListView1.SelectedItem
End If
End Sub
