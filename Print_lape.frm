VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form Print_lape 
   Caption         =   "Prin Lape"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2640
   Icon            =   "Print_lape.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin ÑhameleonButton.chameleonButton Command1 
      Height          =   495
      Left            =   -120
      TabIndex        =   2
      Top             =   4200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "&Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
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
      MICON           =   "Print_lape.frx":030A
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
   Begin VB.ListBox lst_print 
      CausesValidation=   0   'False
      Height          =   3660
      IntegralHeight  =   0   'False
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   510
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Print_lape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim i As Integer
        On Error GoTo ERR
        'Screen.MousePointer = 11
        
        frmView.Pages = -1
        frmView.CurentPage = 0
        
        For i = 0 To lst_print.ListCount - 1
            If lst_print.Selected(i) = True Or lst_print.ListIndex = i Then
                
                Dim CurrentSlope As Integer
                If Asc(Me.lst_print.List(i)) > 96 Then
                    CurrentSlope = Asc(Me.lst_print.List(i)) - 70
                Else
                    CurrentSlope = Asc(Me.lst_print.List(i)) - 64
                End If
            
                If FlagPrinter = 1 Then
                
                    Print_ALL.Print_Scat CurrentSlope, Printer
                    
                Else
                    
                    frmView.Pages = frmView.Pages + 1
                    frmView.Cls
                    frmView.SSlope = CurrentSlope
                    
                    If frmView.Pages > 0 Then
                        Load frmView.pic_view(frmView.Pages)
                        frmView.pic_view(frmView.Pages).Width = 15500
                        frmView.pic_view(frmView.Pages).Height = 29000
                        frmView.pic_view(frmView.Pages).Top = 0
                        frmView.pic_view(frmView.Pages).Left = 0
                        frmView.m_fovard.Enabled = True
                    End If
                    
                    On Error Resume Next
            
'                    frmView.pic_view(frmView.Pages).ScaleHeight = SlP(CurrentSlope).ScaleHeightS
'                    frmView.pic_view(frmView.Pages).ScaleLeft = SlP(CurrentSlope).ScaleLeftS
'                    frmView.pic_view(frmView.Pages).ScaleTop = SlP(CurrentSlope).ScaleTopS
'                    frmView.pic_view(frmView.Pages).ScaleWidth = SlP(CurrentSlope).ScaleWidthS
            
'                    Print_ALL.PicScaleWidth = frmView.pic_view.ScaleWidth * 1.5
'                    Print_ALL.PicScaleHeight = frmView.pic_view.ScaleHeight * 2.217 * 1.3
'                    Print_ALL.PicScaleLeft = frmView.pic_view.ScaleLeft - 0.25 * frmView.pic_view.ScaleWidth
'                    Print_ALL.PicScaleTop = frmView.pic_view.ScaleTop - 0.5 * frmView.pic_view.ScaleHeight
'
'                    Print_ALL.PicScaleWidth = frmView.pic_view.ScaleWidth * 1.3
'                    Print_ALL.PicScaleHeight = frmView.pic_view.ScaleHeight * 2.217 * 1.3
'                    Print_ALL.PicScaleLeft = frmView.pic_view.ScaleLeft - 0.25 * frmView.pic_view.ScaleWidth
'                    Print_ALL.PicScaleTop = frmView.pic_view.ScaleTop - 0.5 * frmView.pic_view.ScaleHeight
                    
'                    Print_ALL.PicScaleWidth = frmView.pic_view(frmView.Pages).ScaleWidth '* 1.8
'                    Print_ALL.PicScaleHeight = frmView.pic_view(frmView.Pages).ScaleHeight * 3 '* 4.7
'                    Print_ALL.PicScaleLeft = frmView.pic_view(frmView.Pages).ScaleLeft '- 0.001 * frmView.pic_view.ScaleWidth
'                    Print_ALL.PicScaleTop = frmView.pic_view(frmView.Pages).ScaleTop * 1.85 '- 0.5 * frmView.pic_view.ScaleHeight
            
                    Print_ALL.Print_Scat CurrentSlope, frmView.pic_view(frmView.Pages)
        
                End If
                
            End If
        Next i

        'Screen.MousePointer = 0
        Me.Hide

        Exit Sub

ERR:
        'Screen.MousePointer = 0
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.28." & ERR.Source & "]", ERR.Number, ERR.Description
        If MsgBox(lng.GetResIDstring(9141) & " (" & Printer.DeviceName & ")" & vbNewLine & Printer.Port, vbCritical, lng.GetResIDstring(1434)) = vbOK Then Exit Sub
End Sub


Private Sub Form_Load()
    SetFont Me
    Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
    Me.Top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.Top)
    Me.Width = GetSetting(App.ProductName, "Position", Me.name & "width", Me.Width)
    Me.Height = GetSetting(App.ProductName, "Position", Me.name & "height", Me.Height)
    Label1 = lng.GetResIDstring(1046)
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.Label1.Width = Me.Width
    Command1.Width = Me.Width '- 120
    Me.lst_print.Top = Me.Label1.Height
    Me.lst_print.Width = Me.ScaleWidth
    Me.lst_print.Height = Me.ScaleHeight - Me.Label1.Height - Command1.Height
    Command1.Top = Me.ScaleHeight - Command1.Height
    'Me.Text1.Left = (Me.ScaleWidth - Me.Text1.Width) / 2
    'Me.Text1.top = (Me.ScaleHeight - Me.Text1.Height) / 2
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
    SaveSetting App.ProductName, "Position", Me.name & "top", Me.Top
    SaveSetting App.ProductName, "Position", Me.name & "width", Me.Width
    SaveSetting App.ProductName, "Position", Me.name & "height", Me.Height
End Sub

