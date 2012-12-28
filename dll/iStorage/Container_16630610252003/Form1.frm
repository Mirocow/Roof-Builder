VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check object"
      Height          =   435
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Direct"
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test List"
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Recordset"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtLoops 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "20000"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Hash"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Dictionary"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Array"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Collection"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   120
      Width           =   12495
   End
   Begin VB.Label Label1 
      Caption         =   "Loops:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' Form1
'--------------------------------------------------------------------------------
Option Explicit

' There are many different ways to store data in VB. This application tests a few of them.
' You shouldn't blindly choose the method that has the highest numbers!
' You should choose a method based on the usage pattern of your data.
' For instance, even though the timings say 'clsDirect' is fastest,
' for my application, a combination of clsHash and clsList proved to work 50% faster.
' If you don't know which method is best, try them all - they all implement interface
' IContainer, and swapping out one for another is a simple one-line code change.
' They're listed on the buttons in descending order of probable speed/usefulness.

' SPECIAL NOTE FOR 'clsList': THIS ROUTINE STORES NO DATA! It only manages lists of numbers.


Private Sub Display(ByVal s As String)
    ' Display text
    With txtOutput
        .SelStart = Len(.Text)
        .SelText = s & vbCrLf
        .SelStart = Len(.Text)
    End With
End Sub


Private Sub TestIt(ByVal obj As IContainer)
    ' Test a container
    Dim Timings As clsTimings
    Const s As String = "          "
    Dim Loops As Long
    
    If Check1.Value Then
        If obj.Test(obj) = False Then
            Display "Tests failed!"
            Exit Sub
        End If
    End If
    
    Loops = CLng(txtLoops.Text)
    If (obj.Name = "clsArray" Or obj.Name = "clsRecordset") And Loops > 2000 Then
        Loops = 2000
    End If
    Set Timings = obj.GetTimings(obj, Loops)
    Display "TESTING " & obj.Name & " (" & Loops & " loops) in loops per second"
    Display vbTab & "Bytes: " & vbTab & Right$(s & Timings.BytesPerLoop, 10)
    Display vbTab & "Add:   " & vbTab & Right$(s & Timings.AddsPerSecond, 10)
    If obj.Name <> "clsList" Then
        Display vbTab & "Get:   " & vbTab & Right$(s & Timings.GetsPerSecond, 10)
        Display vbTab & "Set:   " & vbTab & Right$(s & Timings.SetsPerSecond, 10)
        Display vbTab & "Exist: " & vbTab & Right$(s & Timings.ExistsPerSecond, 10)
        Display vbTab & "Lookup:" & vbTab & Right$(s & Timings.LookupsPerSecond, 10)
    End If
    Display vbTab & "Remove:" & vbTab & Right$(s & Timings.DeletesPerSecond, 10)
    Display vbTab & "Keys:  " & vbTab & Right$(s & Timings.KeysPerSecond, 10)
    Display ""
End Sub

Private Sub cmdClear_Click()
    ' Clear the text
    txtOutput.Text = ""
End Sub

Private Sub cmdTest_Click(Index As Integer)
    ' Perform the selected text
    Dim obj As IContainer
    Dim Test As String
    
    Test = cmdTest(Index).Caption
    
    If Test = "Test Collection" Then
        Set obj = New clsCollection
    ElseIf Test = "Test Array" Then
        Set obj = New clsArray
    ElseIf Test = "Test Dictionary" Then
        Set obj = New clsDictionary
    ElseIf Test = "Test Hash" Then
        Set obj = New clsHash
    ElseIf Test = "Test Recordset" Then
        Set obj = New clsRecordset
    ElseIf Test = "Test List" Then
        Set obj = New clsList
    ElseIf Test = "Test Direct" Then
        Set obj = New clsDirect
    Else
        MsgBox "Unknown type " & Test & "!"
    End If
    
    TestIt obj

'    Dim n As Long
'    For n = 0 To 1000
'        Dim Arr() As Integer
'        Dim i As Integer
'        For i = 0 To 1000
'            ReDim Arr(i)
'            Arr(i) = Rnd(1000)
'        Next
'
'        obj.Add Arr, n
'        Erase Arr
'        Display "Lookup variable " & n & " - " & obj.Lookup(n, Arr)
'        obj.Remove n
'        Display "Remove variable " & n & " - " & (obj.Exists(n) Or True)
'    Next

'    Dim A As Object
'    Set A = New Collection
'    obj.Add A, 5
''    Display obj.Item(1000)
'    Dim cData As Object
'    Display obj.Lookup(5, cData)
'    obj.Remove 5
'    Display obj.Exists(5) Or True

End Sub

Private Sub Form_Load()
    ' Display help
    Display "There's many different ways to store data in VB. This application tests a few of them."
    Display "You shouldn't blindly choose the method that has the highest numbers!"
    Display "You should choose a method based on the usage pattern of your data."
    Display "For instance, even though the timings say 'clsDirect' is fastest,"
    Display "for my application, a combination of clsHash and clsList proved to work 50% faster."
    Display "If you don't know which method is best, try them all - they all implement interface"
    Display "IContainer, and swapping out one for another is a simple one-line code change."
    Display "They're listed on the buttons in descending order of probable speed/usefulness."
    Display ""
    Display "Hash       - Custom hash class. Uses array and hashing algorithm to store data."
    Display "Array      - Built-in VB type. Uses brute force search on keys."
    Display "Direct     - Custom class, optimized for speed. Keys are automatically assigned."
    Display "List       - DOESN'T STORE DATA, only maintains lists of numbers. Unsafe - lets you add duplicates!"
    Display "Dictionary - Microsoft Scripting Runtime type. Ordered name/value pairs."
    Display "Collection - Built-in VB type. Keeps track of keys added in a separate array."
    Display "Recordset  - ADOR recordset, memory only (disconnected from database)."
    Display ""
    Display "NOTE: Timings are in loops per second. Bigger numbers are better."
    Display "      Values of -1 means it was too fast to measure - increase the number of loops to get an accurate count."
    Display ""
End Sub
