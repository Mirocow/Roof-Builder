VERSION 5.00
Object = "{A16B04F6-4B18-424D-9B66-598EBF3A90F9}#1.0#0"; "Geometry.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   7680
      Width           =   6855
      Begin VB.TextBox Text3 
         Height          =   1335
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Text            =   "0"
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove line"
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   120
      Top             =   120
   End
   Begin prjCanvas.Canvas Canvas1 
      Height          =   4695
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8281
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Point with Children by Key"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Text            =   "0"
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear All"
      Height          =   615
      Left            =   6120
      TabIndex        =   1
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Canvas1.Clear
End Sub

Private Sub Command2_Click()
Canvas1.RemovePoint Text1
Canvas1.Cls
Canvas1.ReDraw
End Sub

Private Sub Command3_Click()
Canvas1.RemoveLine Text2
Canvas1.Cls
Canvas1.ReDraw
End Sub

Private Sub Timer1_Timer()
Label1 = ""
Label2 = "X=" & Canvas1.X & ", Y=" & Canvas1.Y

Dim FP As CPoint
Set FP = Canvas1.FindPoint

Dim mLines As CLines
Set mLines = Canvas1.GetLines

If Not FP Is Nothing Then

    Label1 = Label1 & "Is Point: " & FP.isPoint & vbNewLine
    
    Label1 = Label1 & "FindPoint = " & FP.Key & " (X=" & FP.X & ", Y=" & FP.Y & ")" & vbNewLine

    Label1 = Label1 & "Amount of Edges = " & FP.Children & vbNewLine

'    If FP.Children > 0 Then

        Dim Childrens As Collection
        Set Childrens = mLines.GetPointsChildren(FP)
        Dim CL As CLine
        Dim str As String
        For Each CL In Childrens
            str = str & CL.Key & ","
        Next
        Label1 = Label1 & "Edges = (" & str & ")" & vbNewLine
        
'    End If
    
End If

Set mLines = Nothing

Dim FL As CLine
Set FL = Canvas1.FindLine
If Not FL Is Nothing Then
Label1 = Label1 & "Edge = " & FL.Key & ": (" & FL.BeginPoint.Key & " -> " & FL.EndPoint.Key & ")" & vbNewLine
End If

Label1 = Label1 & "Lines = " & Canvas1.LinesCount & vbNewLine
Label1 = Label1 & "Points = " & Canvas1.PointsCount & vbNewLine
End Sub
