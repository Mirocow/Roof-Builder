VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form frmTip 
   Caption         =   "Tip of the Day"
   ClientHeight    =   3345
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Appearance      =   0  'Flat
      Caption         =   "&Show tips on startup"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2940
      Value           =   2  'Grayed
      Width           =   2775
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2505
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmTip.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblTipText 
      BackStyle       =   0  'Transparent
      Caption         =   "lblTipText"
      Height          =   1695
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "À Âû çíàåòå ..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   960
      X2              =   5880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   120
      Top             =   120
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   2655
      Left            =   120
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    ' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPOFDAY"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Select a tip at random.
    CurrentTip = Int((Tips.count * Rnd) + 1)
    Label2 = CurrentTip & "/" & Tips.count
    frmTip.DisplayCurrentTip
    
End Sub


Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
        ' Obtain the next free file descriptor.
        InFile = FreeFile
    
        ' Make sure a file is specified.
        If sFile = "" Then
            LoadTips = False
            Exit Function
        End If
    
        ' Make sure the file exists before trying to open it.
        If dir(sFile) = "" Then
            LoadTips = False
            Exit Function
        End If
    
        ' Read the collection from a text file.
        Open sFile For Input As InFile
        While Not EOF(InFile)
            Line Input #InFile, NextTip
        
            NextTip = Replace(NextTip, "[crlf]", vbCrLf)
            NextTip = Replace(NextTip, "%ProjectPath%", Gl.ProjectsDir)
        
            '        NextTip = Replace(NextTip, "[br]", chr$(13))
            '        NextTip = Replace(NextTip, "[sp]", Space(5))
                
                
            Tips.Add NextTip
        Wend

        Close InFile

        ' Display a tip at random.
        DoNextTip
    
        LoadTips = True
    
End Function


Private Sub chkLoadTipsAtStartup_Click()
    ' save whether or not this form should be displayed at startup
    SaveSetting App.ProductName, "Main", "Show Tips at Startup", chkLoadTipsAtStartup.value
End Sub


Private Sub cmdNextTip_Click()
    DoNextTip
End Sub


Private Sub cmdOK_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    On Error Resume Next

    'SetFont Me

    Label1 = lng.GetResIDstring(9121)
    chkLoadTipsAtStartup.Caption = lng.GetResIDstring(9118)
    cmdNextTip.Caption = lng.GetResIDstring(9119)
    Me.Caption = lng.GetResIDstring(9117)

    Dim ShowAtStartup As Long
    
        ' See if we should be shown at startup
       
        ' Set the checkbox, this will force the value to be written back out to the registry
        Me.chkLoadTipsAtStartup.value = vbChecked
    
        ' Seed Rnd
        Randomize
    
        ' Read in the tips file and display a tip at random.
        If LoadTips(App.Path & "\doc\" & TIP_FILE & lng.GetResIDstring(100) & ".TXT") = False Then
            cmdNextTip.Enabled = False
            '        lblTipText.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
'           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
'           "Then place it in the same directory as the application. "
            lblTipText.Caption = lng.GetResIDstring(9116)
        Else
            Label2.Visible = True
        End If

    
End Sub


Public Sub DisplayCurrentTip()
    If Tips.count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If

End Sub


Private Sub Picture1_Click()

End Sub


