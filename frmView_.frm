VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form frmView 
   Caption         =   "View"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   12180
   ControlBox      =   0   'False
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Форма1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   12180
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   7500
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   7980
      Left            =   2520
      ScaleHeight     =   7920
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   120
      Width           =   9675
      Begin VB.VScrollBar VScroll1 
         Height          =   7695
         LargeChange     =   5000
         Left            =   9360
         Max             =   22800
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   5000
         Left            =   0
         Max             =   7000
         TabIndex        =   5
         Top             =   7680
         Width           =   9375
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7455
         Left            =   120
         ScaleHeight     =   7425
         ScaleWidth      =   9090
         TabIndex        =   1
         Top             =   120
         Width           =   9120
         Begin VB.PictureBox pic_view 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   8475
            Index           =   0
            Left            =   720
            ScaleHeight     =   4755
            ScaleMode       =   0  'User
            ScaleWidth      =   5235
            TabIndex        =   2
            Top             =   480
            Width           =   10425
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Label1"
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
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu menu 
      Caption         =   "Сохранить как"
      Begin VB.Menu menusaveasgif 
         Caption         =   "Сохранить в формате gif"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu menusaveasjpg 
         Caption         =   "Сохранить в формате jpg"
      End
      Begin VB.Menu menusaveasbmp 
         Caption         =   "Сохранить в формате bmp"
      End
      Begin VB.Menu menusaveasxls 
         Caption         =   "Сохранить в формате xls"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menup 
      Caption         =   "Печать"
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private pOP As Integer ' Номер выбора печати
Public SSlope As Integer ' Номер ската для печати
Public pFunc As Integer

Dim fName
Dim RetVal As Long
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long

Private Sub PrintPic(Obj As Object)
    On Error GoTo ERR
    Select Case pOP
        Case 1
            MainPicture Obj
        Case 2
            If SSlope > 0 Then Print_ALL.Print_Scat SSlope, Obj
        Case 3
            Summary Obj, pFunc
        Case 4
    End Select
Exit Sub
ERR:
End Sub


Private Sub Form_Load()
    On Error Resume Next

    SetFont Me, 1

    Me.Caption = lng.GetResIDstring(9122)
    Command3.Caption = lng.GetResIDstring(9133)
    Command1.Caption = lng.GetResIDstring(9134)

    List1.AddItem Replace(lng.GetResIDstring(1310), "&", "")
    List1.AddItem Replace(lng.GetResIDstring(1311), "&", "")
    List1.AddItem Replace(lng.GetResIDstring(1312), "&", "")

    Label1 = OfficeStart.da.Caption

    pic_view(0).Width = 15500
    pic_view(0).Height = 29000
    
    VScroll1.MAX = pic_view(0).Height - Pic.Height
    
    pic_view(0).Top = 0
    pic_view(0).Left = 0

    'Me.WindowState = 2
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    SplitHV1.Height = Me.Height
    SplitHV1.ResizeControl
    Picture1.Height = Me.ScaleHeight - 100
    Picture1.Width = Me.ScaleWidth - 2500
    Pic.Height = Picture1.Height - 500
    Pic.Width = Picture1.Width - 500
    VScroll1.Left = Picture1.Width - VScroll1.Width - 80
    VScroll1.Height = Pic.Height + 150
    HScroll1.Top = Pic.Height + 160
    HScroll1.Width = Picture1.Width - 350
    List1.Height = Me.Height - 1100
End Sub


Private Sub HScroll1_Change()
    HScroll1_Scroll
End Sub


Private Sub HScroll1_Scroll()
    pic_view(0).Left = 0 - HScroll1.value
End Sub


Private Sub List1_Click()
   FlagPrinter = 0

   Dim l008C
 
  pic_view(0).Cls

  pic_view(0).ScaleMode = 1

  On Error Resume Next

  pOP = List1.ListIndex + 1
  Select Case List1.ListIndex + 1
      Case 1
          MainPicture pic_view(0)
      Case 2
      Dim l0098 As Integer
      Print_lape.lst_print.clear

  '    Print_lape.lst_print.style = 0

      For l0098 = 1 To KolvoScatov Step 1
          If SlP(l0098).CountSheets > 0 Then
              If l0098 > 26 Then
                  Print_lape.lst_print.AddItem Chr$(l0098 + 70) & "  " & SlP(l0098).ProfilName
              Else
                  Print_lape.lst_print.AddItem Chr$(l0098 + 64) & "  " & SlP(l0098).ProfilName
              End If

          End If

      Next l0098

      Print_lape.Show vbModal, OfficeStart: FlagPrinter = 0 '
      Unload Print_lape
  Case 3
      frmView.pFunc = -1
      Summary pic_view(0)
  Case 4
  End Select
End Sub


Private Sub menup_Click()
    On Error GoTo ERR
    OfficeStart.CommonDialog1.CancelError = True
    OfficeStart.CommonDialog1.ShowPrinter
    FlagPrinter = 1
    PrintPic Printer
Exit Sub
ERR:
End Sub


Private Sub menusaveasbmp_Click()
    Dim FILE As String
        On Error GoTo ERR

        Me.MousePointer = 11
        If dir(Gl.ProjectsDir & "\pic", vbDirectory) = "" Then MkDir Gl.ProjectsDir & "\pic"

        FILE = Dialog.GetFileName(Left(Gl.CurrentFile, Len(CurrentFile) - 4), "Bitmap file (*.bmp)|*.bmp", Gl.ProjectsDir & "\pic", False, Me.hwnd)

        If FILE = "" Then Exit Sub

        If FILE <> "" Then SavePicture pic_view(0).Image, FILE
        Me.MousePointer = 0

        Exit Sub
ERR:
        Me.MousePointer = 0
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.27." & ERR.Source & "]", ERR.Number, ERR.Description
        Resume Next
End Sub


Private Sub menusaveasjpg_Click()
    On Error Resume Next
    If dir(Gl.ProjectsDir & "\pic", vbDirectory) = "" Then MkDir Gl.ProjectsDir & "\pic"
    FILE = Dialog.GetFileName(Left(Gl.CurrentFile, Len(CurrentFile) - 4) & "_" & List1.list(List1.ListIndex), "Bitmap file (*.jpg)|*.jpg", Gl.ProjectsDir & "\pic", False, Me.hwnd)

    If FILE = "" Then Me.MousePointer = 0: Exit Sub
    ' инициализируем путь
    loadStr = FILE

    'создаем временно файл,требуемый для DIjpg.dll
    SavePicture pic_view(0).Image, "C:\tmp.bmp"

    'сохраняем в JPEG
    RetVal = DIWriteJpg(loadStr, Setup.Text4, 1)

    If RetVal = 1 Then  'Success

    Else  'An error occured
        MsgBox "DIWriteJpg did not succeed"
    End If

    Kill "C:\tmp.bmp"
End Sub


Private Sub pic_view_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        pic_view(0).MousePointer = 15

        Call ReleaseCapture
        Call SendMessage(pic_view(0).hwnd, &HA1, 2, 0&)

    Else
        '    PopupMenu mnuMenu
    End If
End Sub

Private Sub pic_view_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If isActive.isFormFocus(Me.hwnd) Then
    isActive.SetFormFocus Me.pic_view(0).hwnd
End If
    If Me.pic_view(0).MousePointer = 15 Then
End If
End Sub


Private Sub pic_view_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Me.pic_view(0).MousePointer = 15 Then Me.pic_view(0).MousePointer = 1
End Sub


Private Sub Picture1_Resize()
    On Error Resume Next
    Pic.Width = Picture1.Width - 480
    VScroll1.Left = Picture1.Width - VScroll1.Width - 60
    HScroll1.Width = Picture1.Width - 350
End Sub


Private Sub VScroll1_Change()
    VScroll1_Scroll
End Sub


Private Sub VScroll1_Scroll()
    pic_view(0).Top = 0 - VScroll1.value
End Sub
