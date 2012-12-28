Attribute VB_Name = "Grad"
Option Explicit

Public Const PI = 3.14159265358979

Public Type POINT
    X As Long
    Y As Long
End Type

Public tSinGrd(359) As Single
Public tCosGrd(359) As Single

Public CurrentDeg As Integer


' Взять градусы
Public Function GetGRD(ByRef pA As POINT, ByRef pB As POINT) As Integer
Dim A As Single
Dim c As Single
On Error Resume Next
c = pB.X - pA.X
A = pB.Y - pA.Y
GetGRD = Abs(Atn(A / c)) * 180 / PI
' FIX
If pB.Y > pA.Y And pB.X > pA.X Then ' 0-90
    Exit Function
ElseIf c = 0 And pB.Y > pA.Y Then ' 90
    GetGRD = 90
    Exit Function
ElseIf pB.Y > pA.Y And pA.X > pB.X Then ' 90-180
    GetGRD = 180 - GetGRD
    Exit Function
ElseIf A = 0 And pA.X > pB.X Then ' 180
    GetGRD = 180
    Exit Function
ElseIf pA.Y > pB.Y And pA.X > pB.X Then ' 180-270
    GetGRD = 180 + GetGRD
    Exit Function
ElseIf c = 0 And pB.Y < pA.Y Then ' 270
    GetGRD = 270
    Exit Function
ElseIf pA.Y > pB.Y And pB.X > pA.X Then ' 270-360
    GetGRD = 360 - GetGRD
    Exit Function
End If
End Function

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

Public Function SinGrd(ByVal Angle As Single) As Single
    If tSinGrd(1) = 0 And tCosGrd(1) = 0 Then LoadSinCosTables
    SinGrd = tSinGrd(Angle Mod 360)
End Function

 
Public Function CosGrd(ByVal Angle As Single) As Single
    CosGrd = tCosGrd(Angle Mod 360)
End Function
