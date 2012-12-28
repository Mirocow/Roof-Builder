VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form dump 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dump of slope`s data"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   9128
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "dump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    TreeView1.Visible = True
    TreeView1.Nodes.Clear

    ' разбивка по годам
    'DateCreated = FileDateTime(FILE)
    'strk = Year(DateCreated) & ":"
    '
    'If Not TreeView1.Nodes(strk).Key = strk Then
    '   TreeView1.Nodes.Add , tvwChild, strk, Year(DateCreated) ', 1, 1
    'End If
    '
    'TreeView1.Nodes.Add strk, tvwChild, strk & "." & MonthNames(Month(DateCreated) - 1), MonthNames(Month(DateCreated) - 1)
    '
    'Dim SomeNode As Node
    'Set SomeNode = TreeView1.Nodes.Add(strk & "." & MonthNames(Month(DateCreated) - 1), tvwChild, FILE, FileName)
    '
    'SomeNode.Text = FileName & ": " & open_file(FILE)
    
    On Error Resume Next
    
    ' Dump
    Dim nL As Integer
    Dim strk, strk1 As String
    
    
        For nL = 1 To MAXSLOPELISTS
    
            If List_Properties_PY(N_Slope, nL) = 0 Or List_Properties_PX(N_Slope, nL) = 0 Or List_Properties_Length(N_Slope, nL) = 0 Then Exit For
    
            strk = "List: " & nL
        
            If Not TreeView1.Nodes(strk).Key = strk Then
                TreeView1.Nodes.Add , tvwChild, strk, strk
            End If
       
             If List_Properties_PY(N_Slope, nL) = 0 Or List_Properties_PX(N_Slope, nL) = 0 Or List_Properties_Length(N_Slope, nL) = 0 Then Exit For
        
             strk1 = strk
        
             If Not TreeView1.Nodes(strk1).Key = strk1 Then
                 TreeView1.Nodes.Add strk, tvwChild, strk1, strk
             End If
        
             TreeView1.Nodes.Add strk1, tvwChild, strk1 & "X", "X = " & List_Properties_PX(N_Slope, nL)
             TreeView1.Nodes.Add strk1, tvwChild, strk1 & "Y", "Y = " & List_Properties_PY(N_Slope, nL)
             TreeView1.Nodes.Add strk1, tvwChild, strk1 & "H", "H = " & List_Properties_Length(N_Slope, nL)
        
        Next nL

End Sub

