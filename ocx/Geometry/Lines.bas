Attribute VB_Name = "Lines"
Option Explicit

Public Function Find_Line(CP As CPoints, X As Single, Y As Single) As CLine
    Dim i As Integer
    Dim AC As Single
    Dim BC As Single
    Dim P As CPoint
    Dim Pindex As String
    
    On Error GoTo ERR

    ' 1  - Parent
    ' 0 - i (Child)
    
    For Each P In CP.Points
        
        If P.ParentsCount > 0 Then

            For i = 1 To P.ParentsCount

                If P.GetParent(i) > 0 Then
                
                    Pindex = P.GetParent(i)

                    AC = Sqr((P.X - X) ^ 2 + (P.Y - Y) ^ 2)
                    BC = Sqr((CP.Points(Pindex).X - X) ^ 2 + (CP.Points(Pindex).Y - Y) ^ 2)

                    If Round(AC + BC, 0) = Round(Sqr((CP.Points(Pindex).X - P.X) ^ 2 + _
                    (CP.Points(Pindex).Y - P.Y) ^ 2), 0) Then
                    Set Find_Line = New CLine
                    Find_Line.BeginPoint = CP.Points(Pindex)
                    Find_Line.EndPoint = P
                    Exit Function
                    End If

                End If

            Next
        End If
        
    Next

    Exit Function
ERR:
'    Find_Line = 0
End Function


Function Divide_Line(CP As CPoints, X As Single, Y As Single, Current_L As CLine) As CPoint
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
    
    Dim NewP As CPoint
    Set NewP = CP.AddPoint(X1, Y1, Current_L.BeginPoint, False)
    
    Dim A As CPoint
    Dim B As CPoint
    Set A = CP.Points(Current_L.BeginPoint.Key)
    Set B = CP.Points(Current_L.EndPoint.Key)
    
    A.ChangeChild CInt(B.Key), CInt(NewP.Key)
    B.AddChild CInt(NewP.Key)
    NewP.AddParent CInt(B.Key)
    
    B.RemoveParent CInt(A.Key)
    
    Set Divide_Line = NewP

End Function

