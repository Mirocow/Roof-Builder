VERSION 5.00
Object = "*\ASplitterModern.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin SplitterHV.SplitHV SplitHV2 
      Height          =   75
      Left            =   7920
      TabIndex        =   18
      Top             =   5280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   132
      SplitLimit      =   5000
      LimitBorder     =   0
      Style           =   0
   End
   Begin SplitterHV.SplitHV SplitHV1 
      Height          =   2415
      Left            =   5400
      TabIndex        =   17
      Top             =   2760
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   4260
      SplitLimit      =   5000
      LimitBorder     =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Text            =   "500"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "LeftTop"
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   6360
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "RightBottom"
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   3840
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   5415
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   6120
      ScaleHeight     =   4515
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      Begin VB.TextBox Text4 
         Height          =   1935
         Left            =   2160
         TabIndex        =   8
         Text            =   "Text4"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   2415
         Left            =   360
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   6480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
SplitHV1.isRollUp = True
SplitHV2.isRollUp = True
End Sub

Private Sub Command4_Click()
SplitHV1.SplitLimit = 0
SplitHV2.SplitLimit = 0
End Sub


Private Sub Form_Load()
Text1 = "rrrr"

Set SplitHV1.obj1 = Me.Text1
Set SplitHV1.obj1 = File1
Set SplitHV1.obj1 = Frame1
Set SplitHV1.obj1 = Combo1
Set SplitHV1.obj1 = Command1
Set SplitHV1.obj2 = Me.Picture1

Set SplitHV2.obj1 = Text1
Set SplitHV2.obj1 = File1
Set SplitHV2.obj1 = Frame1
Set SplitHV2.obj1 = Combo1
Set SplitHV2.obj1 = Command1
Set SplitHV2.obj1 = SplitHV1
Set SplitHV2.obj1 = Picture1
Set SplitHV2.obj2 = Text2

End Sub

Private Sub Form_Resize()
SplitHV1.ResizeControl
SplitHV2.ResizeControl
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Option1_Click()
'SplitHV1.SplitLimit = Val(Text5.Text): SplitHV2.SplitLimit = SplitHV1.SplitLimit
If Option1.Value Then SplitHV1.Binding = RightBottom: SplitHV2.Binding = RightBottom
End Sub

Private Sub Option2_Click()
'SplitHV1.SplitLimit = Val(Text5.Text): SplitHV2.SplitLimit = SplitHV1.SplitLimit
If Option2.Value Then SplitHV1.Binding = LeftTop: SplitHV2.Binding = LeftTop
End Sub

Private Sub Picture1_Resize()
Label1 = Picture1.Width
Label2 = Picture1.Height
End Sub

Private Sub SplitHV2_MoveEnd()
Text2 = Text2.Height
End Sub

