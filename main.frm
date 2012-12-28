VERSION 5.00
Begin VB.Form main 
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   50
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Frame mainpfrm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00DCFBFC&
         Height          =   3855
         Left            =   120
         MouseIcon       =   "main.frx":030A
         ScaleHeight     =   3795
         ScaleWidth      =   4755
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Select_Slope As Integer

Private Sub Form_Load()
    Me.Caption = OfficeStart.menu_view_m.Item(10).Caption
    Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
    Me.Top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.Top)
    Me.Width = GetSetting(App.ProductName, "Position", Me.name & "width", Me.Width)
    Me.Height = GetSetting(App.ProductName, "Position", Me.name & "height", Me.Height)
    OfficeStart.menu_view_m(11).Checked = True
End Sub


Private Sub Form_Resize()
    On Error GoTo ERR
    mainpfrm.Width = Me.ScaleWidth - 120
    mainpfrm.Height = Me.ScaleHeight - 400
    Picture2.Width = mainpfrm.Width - 250
    Picture2.Height = mainpfrm.Height - 330
    HScroll1.Top = Me.Height - 750
    HScroll1.Width = mainpfrm.Width
    Module10.Draw_Systems Me.Picture2
    
    Exit Sub
ERR:
'    STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
'    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.41." & ERR.Source & "]", ERR.Number, ERR.Description
    Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
    SaveSetting App.ProductName, "Position", Me.name & "top", Me.Top
    SaveSetting App.ProductName, "Position", Me.name & "width", Me.Width
    SaveSetting App.ProductName, "Position", Me.name & "height", Me.Height
'    OfficeStart.menu_view_m(11).Checked = False
    Lapepic.Command10.value = False
End Sub


Private Sub HScroll1_Change()
On Error Resume Next
    PolygonRotate HScroll2.value
    Label12.Caption = Label6.Caption & HScroll2.value
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Select_Slope <> 0 Then
    If Select_Slope > 26 Then
        az = Chr$(Select_Slope + 70)
    Else
        az = Chr$(Select_Slope + 64)
    End If
    
    N_Slope = Select_Slope

    Lapemenu.fill_slope
    'Lapemenu.List1.ListItems(Select_Slope).Selected = True
    Lapemenu.Command4_Click

    Lapepic.SetDrawBorder
'    Lapepic.SavePolygon False

    Lapepic.Command5.value = True
End If
End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isActive.isFormFocus(Me.hwnd) Then
        isActive.SetFormFocus Picture2.hwnd
    End If
    Select_Slope = ROOFPIC.Find_lape_label(X, Y)
    If Select_Slope <> 0 Then
        Picture2.MousePointer = 99
    Else
        Picture2.MousePointer = 0
    End If
End Sub


Private Sub Picture2_Resize()
    On Error Resume Next
    Me.Picture2.ScaleLeft = ScaleLeft_Main 'ROOFPIC.Picture1.ScaleLeft
    Me.Picture2.ScaleTop = ScaleTop_Main 'ROOFPIC.Picture1.ScaleTop
    Me.Picture2.ScaleWidth = ScaleWidth_Main 'ROOFPIC.Picture1.ScaleWidth
    Me.Picture2.ScaleHeight = ScaleHeight_Main 'ROOFPIC.Picture1.ScaleHeight
End Sub


Private Function PolygonRotate(a As Integer)
    Dim i As Long
    Dim pPoint As POINT
    Dim pOrigin As POINT
    Dim pResult As POINT
    
    On Error Resume Next

    pOrigin.X = (XMin + XMax) / 2
    pOrigin.Y = (YMin + YMAx) / 2

    For i = 1# To SlP(N_Slope).CountOfPoints Step 1
    
        pPoint.X = SaveLape_Points_X(i)
        pPoint.Y = SaveLape_Points_Y(i)
    
        pResult = RotatePoint(pPoint, pOrigin, a)
    
        Lape_Points_X(N_Slope, i) = pResult.X
        Lape_Points_Y(N_Slope, i) = pResult.Y
    
    Next i

    Draw_Systems Me.Picture2

End Function
