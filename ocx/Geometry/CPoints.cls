VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Collection wich stores the Pointes
Public Points As Collection
Private PointsIndexes As Integer

'' POINTS

'Adds a new Point to the polygon
Public Function AddPoint(X As Single, Y As Single, Optional Parent As CPoint, Optional AddChild As Boolean = True) As CPoint
    Dim pA As New CPoint
    PointsIndexes = PointsIndexes + 1
    pA.Key = PointsIndexes
    pA.X = X
    pA.Y = Y
    If Not Parent Is Nothing Then
        pA.AddParent Parent.Key
        If AddChild Then Points(Parent.Key).AddChild pA.Key
    End If
    Points.Add pA, pA.Key
    Set AddPoint = pA
    Set pA = Nothing
End Function


'Removes a Point from the polygon
Public Sub RemovePoint(Key As String)
On Error GoTo ERR
    If Not Points(Key) Is Nothing Then
        ' Удаление у всех дитя
        Dim P As CPoint
        For Each P In Points
            If P.ChildrenCount > 0 Then
            If P.FindChild(CInt(Key)) > -1 Then
                P.RemoveChild CInt(Key)
            End If
            End If
        Next
        ' Удаление у всех родителя
        For Each P In Points
            If P.ParentsCount > 0 Then
            If P.FindParent(CInt(Key)) > -1 Then
                P.RemoveParent CInt(Key)
            End If
            End If
        Next
        ' Удаление точки
        Points.Remove Key
    End If
ERR:
End Sub

Public Function FindPoint(X As Single, Y As Single, Optional Distance As Single) As CPoint
Dim P As CPoint
Dim l00D2 As Single
Dim l00E2 As Single

l00D2 = 99999
For Each P In Points
    l00E2 = Sqr((X - P.X) ^ 2 + (Y - P.Y) ^ 2)
    If l00E2 < l00D2 Then
        l00D2 = l00E2
        If Distance = 0 Or Distance >= l00E2 Then
            Set FindPoint = P
        End If
    End If
Next
End Function

Public Sub Connection(Key As String, Parent As CPoint)
If Parent.Key = Key Or Key = Points(Parent.Key).FindParent(CInt(Key), True) Or Parent.Key = Points(Key).FindParent(CInt(Parent.Key), True) Then Exit Sub

If Key > Parent.Key Then
Points(Key).AddParent CInt(Parent.Key)
Points(Parent.Key).AddChild CInt(Key)
Else
Points(Parent.Key).AddParent CInt(Key)
Points(Key).AddChild CInt(Parent.Key)
End If
End Sub

'Returns the number of points in the polygon
Public Function PointsCount() As Long
    PointsCount = Points.Count
End Function

'Clears all the points in the polygon
Public Sub ClearPoints()
   Set Points = Nothing
   Set Points = New Collection
   PointsIndexes = 0
End Sub

'' PRIVATE

' Удаление Родителя

' Удаление дитя

'' CLASS

Private Sub Class_Initialize()
    ClearPoints
End Sub

Private Sub Class_Terminate()
    Set Points = Nothing
End Sub
