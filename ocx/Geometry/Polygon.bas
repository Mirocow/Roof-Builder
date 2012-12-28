Attribute VB_Name = "Polygon"
Option Explicit

Public Function SearchPolygon(mLines As Collection) As Collection
    Dim P1 As CPoint
    Dim P2 As CPoint
    Dim CurL As CLine
    Dim Lines As New Collection
    
    On Error GoTo ERR
    For Each L In mLines ' Проверяем каждую линию
    
        Set P1 = L.BeginPoint
        Set P2 = L.EndPoint
        
        ' Отсеивание отдельно стоящих точек
        If P1.Key <> P2.kex Then
            
            Do
                
            While P1.Key = CurL.BeginPoint.Key Or P1.Key = CurL.EndPoint.Key
            
        End If
        
    Next
    
Exit Function
ERR:
End Function
