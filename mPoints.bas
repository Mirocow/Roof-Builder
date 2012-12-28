Attribute VB_Name = "mPoints"
Option Explicit

' Для поворачивания фигуры
'
'Public Type POINT
'    X As Single
'    Y As Single
'End Type
 
'Public P As POINT

Public tSinGrd(359) As Single
Public tCosGrd(359) As Single

Public CurrentDeg As Integer

Function D2R(ByVal Angle As Single) As Single
    D2R = Angle / 180 * PI
End Function


' pOrigin - точка в округ которой вращаем
Public Function RotatePoint(ByRef pPoint As POINT, ByRef pOrigin As POINT, ByVal Degrees As Single) As POINT
    RotatePoint.X = pOrigin.X + (CosGrd(Degrees) * (pPoint.X - pOrigin.X)) - (SinGrd(Degrees) * (pPoint.Y - pOrigin.Y))
    RotatePoint.Y = pOrigin.Y + (SinGrd(Degrees) * (pPoint.X - pOrigin.X)) + (CosGrd(Degrees) * (pPoint.Y - pOrigin.Y))
End Function


Public Sub LoadSinCosTables()
    Dim i, n As Long
    For i = 0 To 359 Step 1
        tSinGrd(i) = Sin(D2R(i))
        tCosGrd(i) = Cos(D2R(i))
    Next i
End Sub


'Function SinGrd(ByVal Angle As Single) As Single
'Dim A As Single
'
'A = Round(Angle, 1)
'
'Do
'    If A > 359 Then A = A - 360
'    If A < 0 Then A = A + 360
'Loop Until (A >= 0) And (A <= 359)
'
'SinGrd = tSinGrd(A * 10)
'End Function
'
'Function Cos(ByVal Angle As Single) As Single
'Dim A As Single
'
'A = Round(Angle, 1)
'
'Do
'    If A > 359 Then A = A - 360
'    If A < 0 Then A = A + 360
'Loop Until (A >= 0) And (A <= 359)
'
'CosGrd = tCos(A * 10)
'End Function

Public Function SinGrd(ByVal Angle As Single) As Single
    SinGrd = tSinGrd(Angle Mod 360)
End Function

 
Public Function CosGrd(ByVal Angle As Single) As Single
    CosGrd = tCosGrd(Angle Mod 360)
End Function

