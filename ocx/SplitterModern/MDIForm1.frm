VERSION 5.00
Object = "*\ASplitterModern.vbp"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin SplitterHV.SplitHV SplitHV1 
      Align           =   3  'Align Left
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   9869
      SplitLimit      =   5000
      LimitBorder     =   0
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long

Private Sub MDIForm_Load()
Me.Show
Load Form1
Load Form2
Call SetParent(Form1.hWnd, Me.hWnd)
Call SetParent(Form2.hWnd, Me.hWnd)
Set SplitHV1.obj1 = Form1
Set SplitHV1.obj2 = Form2
SplitHV1.ResizeControl
Form1.Show
Form2.Show
'Form1.Left = 0
'Form1.Top = 0
'Form1.Height = Me.Height
'Form1.Visible = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload Form1
Unload Form2
End Sub
