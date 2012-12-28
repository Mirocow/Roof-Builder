VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Lines As Collection
Public CurLine As CLine

Private LinesIndexes As Integer

'
' LINE
'

Public Function AddLine(ByVal NewLine As CLine) As CLine
    ' Поиск дублеров или a.key -> b.key или a.key <- b.key
    If isLine(NewLine) Then Exit Function
    ' Поиск точек
    If NewLine.BeginPoint.Key <> NewLine.EndPoint.Key Then
    NewLine.BeginPoint.Children = NewLine.BeginPoint.Children + 1
    NewLine.EndPoint.Children = NewLine.EndPoint.Children + 1
    End If
    LinesIndexes = LinesIndexes + 1
    NewLine.Key = LinesIndexes
    Lines.Add NewLine, NewLine.Key
    Set AddLine = NewLine
End Function

Public Function ReplaceLine(Key As String, ByVal NewLine As CLine) As CLine
    'LinesIndexes = LinesIndexes + 1
    NewLine.Key = LinesIndexes
    Lines.Remove Key
    Lines.Add NewLine, NewLine.Key
End Function

Public Function isLine(CheckLine As CLine) As Boolean
Dim L As CLine
For Each L In Lines
    If (L.BeginPoint.Key = CheckLine.BeginPoint.Key And L.EndPoint.Key = CheckLine.EndPoint.Key) Or _
    (L.BeginPoint.Key = CheckLine.EndPoint.Key And L.EndPoint.Key = CheckLine.BeginPoint.Key) Then
        isLine = True
        Exit For
    End If
Next
End Function

Public Sub RemoveLine(Key As String)
On Error GoTo ERR
    If Not Lines(Key) Is Nothing Then
        ' Удаление линии
        Lines.Remove Key
    End If
ERR:
End Sub

Public Function Find_Line(X As Single, Y As Single) As CLine
    Dim i As Integer
    Dim AC As Single
    Dim BC As Single
    Dim P1 As CPoint
    Dim P2 As CPoint
    Dim L As CLine
    
    On Error GoTo ERR
    
    For Each L In Lines
    
        Set P1 = L.BeginPoint
        Set P2 = L.EndPoint
        
        AC = Sqr((P1.X - X) ^ 2 + (P1.Y - Y) ^ 2)
        BC = Sqr((P2.X - X) ^ 2 + (P2.Y - Y) ^ 2)

        If Round(AC + BC, 0) = Round(Sqr((P2.X - P1.X) ^ 2 + (P2.Y - P1.Y) ^ 2), 0) Then
        Set Find_Line = L
        Exit Function
        End If
       
    Next

    Exit Function
ERR:
'    Find_Line = 0
End Function

Function Divide_Line(X As Single, Y As Single, Current_L As CLine) As CPoint
On Error Resume Next
    
    Dim pA As POINT
    Dim pB As POINT

    pA.X = Current_L.BeginPoint.X
    pA.Y = Current_L.BeginPoint.Y
    pB.X = Current_L.EndPoint.X
    pB.Y = Current_L.EndPoint.Y
    
    Dim cornerA As Integer
    Dim cornerC As Integer
        
    cornerA = GetGRD(pA, pB)
    
    Dim X1 As Single
    Dim Y1 As Single
    
    If cornerA = 0 Then ' 0
        X1 = X
        Y1 = pA.Y
    ElseIf cornerA < 90 Then ' 0-89
        X1 = X
        cornerC = 180 - 90 - cornerA
        Y1 = (((X - pA.X) * SinGrd(cornerA)) / SinGrd(cornerC)) + pA.Y
    ElseIf cornerA = 90 Then ' 90
        X1 = pA.X
        Y1 = Y
    ElseIf cornerA > 90 And cornerA < 180 Then ' 90-179
        X1 = X
        cornerA = 180 - cornerA
        cornerC = 180 - 90 - cornerA
        Y1 = Abs(((X - pA.X) * SinGrd(cornerA)) / SinGrd(cornerC)) + pA.Y
    ElseIf cornerA = 180 Then ' 180
        X1 = X
        Y1 = pA.Y
    ElseIf cornerA > 180 And cornerA < 270 Then
        X1 = X
        cornerA = cornerA - 180
        cornerC = 180 - 90 - cornerA
        Y1 = (((X - pA.X) * SinGrd(cornerA)) / SinGrd(cornerC)) + pA.Y
    ElseIf cornerA = 270 Then ' 180
        X1 = pA.X
        Y1 = Y
    ElseIf cornerA > 270 Then
        X1 = X
        cornerA = 360 - cornerA
        cornerC = 180 - 90 - cornerA
        Y1 = pA.Y - (((X - pA.X) * SinGrd(cornerA)) / SinGrd(cornerC))
    End If
    
    Dim DP As New CPoint
    Dim CL As New CLine
    
    CL.BeginPoint = Current_L.BeginPoint
    
    DP.X = X1
    DP.Y = Y1
    
    DP.Children = DP.Children + 2
    
    CL.EndPoint = DP
    
    AddLine CL
    
    Set CL = Nothing
    Set CL = New CLine
    
    CL.BeginPoint = DP
    CL.EndPoint = Current_L.EndPoint
    
    AddLine CL
    
    RemoveLine Current_L.Key
    
    
    Set CL = Nothing
    
    Set Divide_Line = DP
    
    Set DP = Nothing

End Function

Public Function LinesCount() As Integer
    Dim L As CLine
    Dim i As Integer
    Dim P1 As CPoint
    Dim P2 As CPoint
    
    On Error GoTo ERR
    
    i = 0
    For Each L In Lines
    
        Set P1 = L.BeginPoint
        Set P2 = L.EndPoint
        If P1.Key <> P2.Key Then
            i = i + 1
        End If
       
    Next
    
LinesCount = i
Exit Function
ERR:
End Function

Public Sub ClearLines()
   Set Lines = Nothing
   Set Lines = New Collection
   LinesIndexes = 0
End Sub

'
' POINT
'

Public Function FindPoint(X As Single, Y As Single, Optional Distance As Single) As CPoint
Dim L As CLine
Dim l00D2 As Single
Dim l00E2 As Single
Dim P1 As CPoint
Dim P2 As CPoint

l00D2 = 99999
For Each L In Lines
    Set P1 = L.BeginPoint
    Set P2 = L.EndPoint
    l00E2 = Sqr((X - P1.X) ^ 2 + (Y - P1.Y) ^ 2)
    If l00E2 < l00D2 Then
        l00D2 = l00E2
        If Distance = 0 Or Distance >= l00E2 Then
            Set FindPoint = P1
        End If
    End If
    l00E2 = Sqr((X - P2.X) ^ 2 + (Y - P2.Y) ^ 2)
    If l00E2 < l00D2 Then
        l00D2 = l00E2
        If Distance = 0 Or Distance >= l00E2 Then
            Set FindPoint = P2
        End If
    End If
Next
End Function

Public Function FindLineByKey(Pkey As String) As CLine
Dim L As CLine
Dim P1 As CPoint
Dim P2 As CPoint

For Each L In Lines
    Set P1 = L.BeginPoint
    Set P2 = L.EndPoint

Next
End Function


Public Function GetPointsChildren(P As CPoint) As Collection
    Dim P1 As CPoint
    Dim P2 As CPoint
    Dim L As CLine
    Dim mLines As New Collection
    
    On Error GoTo ERR
    For Each L In Lines  ' Проверяем каждую линию
    
        Set P1 = L.BeginPoint
        Set P2 = L.EndPoint
        If (P1.Key = P.Key Or P2.Key = P.Key) And P1.Key <> P2.Key Then
            mLines.Add L, L.Key
        End If
        
    Next
Set GetPointsChildren = mLines
Exit Function
ERR:
End Function

Public Function RemoveChild(Key As String) As Boolean
    Dim P1 As CPoint
    Dim P2 As CPoint
    Dim L As CLine
    
    On Error GoTo ERR
    For Each L In Lines  ' Проверяем каждую линию
    
        Set P1 = L.BeginPoint
        Set P2 = L.EndPoint
        
        If P1.Key = Key Then
            P1.Children = P1.Children - 1
        End If
        If P2.Key = Key Then
            P2.Children = P2.Children - 1
        End If
        
    Next
'Set GetPointsChildren = Lines
Exit Function
ERR:
End Function

Public Function PointsCount() As Integer
    Dim L As CLine
    Dim i As Integer
    Dim P1 As CPoint
    Dim P2 As CPoint
    
    On Error GoTo ERR
    
    i = 0
    For Each L In Lines
    
        Set P1 = L.BeginPoint
        Set P2 = L.EndPoint
        If P1.Key = P2.Key Then
            i = i + 1
        Else
            i = i + 2
        End If
       
    Next
    
PointsCount = i
Exit Function
ERR:
End Function

'
' PRIVATE
'

Private Sub Class_Initialize()
    ClearLines
End Sub

Private Sub Class_Terminate()
    Set Lines = Nothing
End Sub
