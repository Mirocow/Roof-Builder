VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub ShowTemplate(i As Integer)
UserControl1
End Sub

Sub calc()
Dim i As Integer
Dim a2 As Single
Dim sbox As Single

'value(8).ForeColor = vbBlack: value(11).ForeColor = vbBlack

'For i = 0 To value.Count - 1
'    value(i).Text = Replace(value(i).Text, ".", ",")
'    value(i).ForeColor = vbBlack
'Next

On Error GoTo ERR

Select Case SlopeSampleModule.Text2
Case 0
    ' a = value(2), b = value(0), c = value(1)
    ' a1 = value(4), h = value(3)
    
    '      ___________________________
    'S = \/ p * ( p - a) * ( p - b ) * ( p - c )
    '
    '    Здесь p - это полупериметр, который вычисляется по формуле: p = ( a + b + c ) / 2.
    '
    If value(2) > 0 And value(0) > 0 And value(1) > 0 Then
        If ((CSng(value(2)) + CSng(value(0))) >= value(1) And (CSng(value(2)) + CSng(value(1))) >= value(0) And (CSng(value(0)) + CSng(value(1))) >= value(2)) Then
        'a, b, c - длины сторон заданного треугольника.
        
            Dim P As Single
            P = (CSng(value(2)) + CSng(value(0)) + CSng(value(1))) / 2
            Label2 = Round(Sqr(P * (P - value(2)) * (P - value(0)) * (P - value(1))), 2)
            
            ' S = 1/2*ah
            ' h = S/(a/2)
'            value(3).Enabled = False
            value(3) = Label2 / (value(2) / 2)
'            value(4).Enabled = False
            value(4) = Sqr(value(0) * value(0) - value(3) * value(3))
'            Command2.Enabled = False
            
            Else: Label2 = 0: MsgBox "Ошибка: Одна из сторон превышает сумму противоположных сторон", vbCritical
        End If
        
    ElseIf value(0) > 0 And value(1) > 0 And value(3) > 0 Then
    ' b,c,h
    
    Label2 = Round((value(1) / 2) * value(3), 2)
    value(4) = Round(Sqr(value(0) * value(0) - value(3) * value(3)), 2)
    value(4).Enabled = False
    value(2) = (Label2 / value(3)) * 2
'    value(2).Enabled = False
'    Command2.Enabled = False
    
    ElseIf value(0) > 0 And value(2) > 0 And value(3) Then
    ' b,a,h
    
    Label2 = Round((value(2) / 2) * value(3), 2)
    value(4) = Round(Sqr(value(0) * value(0) - value(3) * value(3)), 2)
'    value(4).Enabled = False
    value(1) = Sqr(((value(0) - value(4)) * (value(0) - value(4))) + CSng(value(3)) * value(3))
'    value(1).Enabled = False
'    Command2.Enabled = False
    
    ElseIf value(2) > 0 And value(1) > 0 And value(3) Then
    ' a,c,h
    
    Label2 = Round((value(1) / 2) * value(3), 2)
    a2 = Round(Sqr(value(1) * value(1) - value(3) * value(3)), 2)
    value(4) = value(2) - a2
'    value(4).Enabled = False
    value(0) = (Label2 / value(3)) * 2
'    value(0).Enabled = False
'    Command2.Enabled = False
    
    ElseIf value(4) > 0 And value(1) > 0 And value(2) Then
    ' a1, c, a
    
    a2 = value(2) - value(4)
    value(3) = Sqr(value(1) * value(1) - a2 * a2)
'    value(3).Enabled = False
    value(0) = Sqr(value(4) * value(4) + CSng(value(3)) * value(3))
'    value(0).Enabled = False
    Label2 = Round((value(1) / 2) * value(3), 2)
    
'    Command2.Enabled = False
    
    ElseIf value(4) > 0 And value(1) > 0 And value(0) Then
    ' a1, c, b
    
    value(3) = Sqr(value(0) * value(0) - value(4) * value(4))
'    value(3).Enabled = False
    value(2) = Sqr(value(1) * value(1) - value(3) * value(3)) + CSng(value(4))
'    value(2).Enabled = False
    Label2 = Round((value(1) / 2) * value(3), 2)
    
'    Command2.Enabled = False
    
    ElseIf value(2) > 0 And value(4) > 0 And value(3) Then
    'a, a1, h
    
    value(0) = Sqr(value(4) * value(4) + value(3) * value(3))
'    value(0).Enabled = False
    a2 = value(2) - value(4)
    value(1) = Sqr(a2 * a2 + value(3) * value(3))
'    value(1).Enabled = False
    Label2 = Round((value(1) / 2) * value(3), 2)
    
    End If
    
Case 1
    If value(6) > 0 And value(7) > 0 Then
    value(5) = Sqr(value(7) * value(7) - value(6) * value(6))
    If value(5) > 0 Then Label2 = Round((value(5) * value(6)) / 2, 2)
    ElseIf value(5) > 0 And value(6) > 0 Then
    value(7) = Sqr(value(5) * value(5) + value(6) * value(6))
    If value(7) > 0 Then Label2 = Round((value(5) * value(6)) / 2, 2)
    ElseIf value(7) > 0 And value(5) > 0 Then
    value(6) = Sqr(value(7) * value(7) - value(5) * value(5))
    If value(6) > 0 Then Label2 = Round((value(5) * value(6)) / 2, 2)
    End If
    
Case 2
    
    If CInt(value(8)) <= CInt(value(11)) Then value(8).ForeColor = vbRed: value(11).ForeColor = vbRed: Exit Sub
    
    If value(12) > 0 And value(9) > 0 Then
        value(10) = Sqr(value(12) * value(12) - value(9) * value(9))
'        value(10).Enabled = False
    End If
    
    If value(10) > 0 And value(9) > 0 Then
        value(12) = Sqr(value(10) * CSng(value(10)) + CSng(value(9)) * value(9))
'        value(12).Enabled = False
    End If
    
    If value(11) > 0 And value(9) > 0 And value(8) > 0 Then
        a2 = value(8) - (CSng(value(9)) + CSng(value(11)))
        Label3 = a2
    End If
    
    If a2 > 0 And (value(10) > 0 Or value(13) > 0) Then
        If value(13) > 0 Then value(10) = Sqr(value(13) * value(13) - a2 * a2) 'value(13).Enabled = False
        If value(10) > 0 Then value(13) = Sqr(value(10) * value(10) + a2 * a2) ': value(10).Enabled = False
        Label2 = Round(0.5 * (CSng(value(8)) + CSng(value(11))) * value(10), 2)
'        Exit Sub
    End If
    
Case 3
    If value(14) > 0 And value(15) > 0 Then
        Label2 = Round(value(14) * value(15), 2)
    End If

Case 4
    If value(17) > 0 And value(19) > 0 And value(16) > 0 Then
        value(21) = value(19) - value(16)
        sbox = value(21) * value(17)
        value(20) = Sqr(value(16) * value(16) + value(17) * value(17))
    ElseIf value(17) > 0 And value(19) > 0 And value(21) > 0 Then
        value(16) = value(19) - value(21)
        sbox = value(21) * value(17)
        value(20) = Sqr(value(16) * value(16) + value(17) * value(17))
    ElseIf value(20) > 0 And value(19) > 0 And value(21) > 0 Then
        value(16) = value(19) - value(21)
        value(17) = Sqr(value(20) * value(20) - value(16) * value(16))
        sbox = value(21) * value(17)
    ElseIf value(19) > 0 And value(16) > 0 And value(20) > 0 Then
        value(17) = Sqr(value(20) * value(20) - value(16) * value(16))
        value(21) = value(19) - value(16)
        sbox = value(21) * value(17)
    End If
    
        Label2 = Round(sbox + (value(16) * value(17)) / 2, 2)

Case 6
    If value(22) > 0 And value(23) > 0 Then
    value(18) = Sqr(value(22) * value(22) - value(23) * value(23))
    If value(18) > 0 Then Label2 = Round((value(18) * value(23)) / 2, 2)
    ElseIf value(18) > 0 And value(23) > 0 Then
    value(22) = Sqr(value(18) * value(18) + value(23) * value(23))
    If value(22) > 0 Then Label2 = Round((value(18) * value(23)) / 2, 2)
    ElseIf value(18) > 0 And value(22) > 0 Then
    value(23) = Sqr(value(22) * value(22) - value(18) * value(18))
    If value(23) > 0 Then Label2 = Round((value(18) * value(23)) / 2, 2)
    End If
    
Case 7
    ' a = value(2), b = value(0), c = value(1)
    ' a1 = value(4), h = value(3)
    
    '      ___________________________
    'S = \/ p * ( p - a) * ( p - b ) * ( p - c )
    '
    '    Здесь p - это полупериметр, который вычисляется по формуле: p = ( a + b + c ) / 2.
    '
    If value(34) > 0 And value(30) > 0 And value(31) > 0 Then
        If ((value(34) + value(30)) >= value(31) And (value(34) + value(31)) >= value(30) And (value(30) + value(31)) >= value(34)) Then
        'a, b, c - длины сторон заданного треугольника.
        
'            Dim P As Single
            P = (CSng(value(34)) + CSng(value(30)) + CSng(value(31))) / 2
            Label2 = Round(Sqr(P * (P - value(34)) * (P - value(30)) * (P - value(31))), 2)
            
            ' S = 1/2*ah
            ' h = S/(a/2)
'            value(33).Enabled = False
            value(33) = Label2 / (value(34) / 2)
'            value(32).Enabled = False
            value(32) = Sqr(value(30) * value(30) - value(33) * value(33))
'            Command2.Enabled = False
            
            Else: Label2 = 0: MsgBox "Ошибка: Одна из сторон превышает сумму противоположных сторон", vbCritical
        End If
        
    ElseIf value(30) > 0 And value(31) > 0 And value(33) > 0 Then
    ' b,c,h
    
    Label2 = Round((value(31) / 2) * value(33), 2)
    value(32) = Round(Sqr(value(30) * value(30) - value(33) * value(33)), 2)
'    value(32).Enabled = False
    value(34) = (Label2 / value(33)) * 2
'    value(34).Enabled = False
'    Command2.Enabled = False
    
    ElseIf value(30) > 0 And value(34) > 0 And value(33) Then
    ' b,a,h
    
    Label2 = Round((value(34) / 2) * value(33), 2)
    value(32) = Round(Sqr(value(30) * value(30) - value(33) * value(33)), 2)
'    value(32).Enabled = False
    value(31) = Sqr(((value(30) - value(32)) * (value(30) - value(32))) + value(33) * value(33))
'    value(31).Enabled = False
'    Command2.Enabled = False
    
    ElseIf value(34) > 0 And value(31) > 0 And value(33) Then
    ' a,c,h
    
    Label2 = Round((value(31) / 2) * value(33), 2)
    a2 = Round(Sqr(value(31) * value(31) - value(33) * value(33)), 2)
    value(32) = value(34) - a2
'    value(32).Enabled = False
    value(30) = (Label2 / value(33)) * 2
'    value(30).Enabled = False
'    Command2.Enabled = False
    
    ElseIf value(32) > 0 And value(31) > 0 And value(34) Then
    ' a1, c, a
    
    a2 = value(34) - value(32)
    value(33) = Sqr(value(31) * value(31) - a2 * a2)
'    value(33).Enabled = False
    value(30) = Sqr(value(32) * value(32) + value(33) * value(33))
'    value(30).Enabled = False
    Label2 = Round((value(31) / 2) * value(33), 2)
    
'    Command2.Enabled = False
    
    ElseIf value(32) > 0 And value(31) > 0 And value(30) Then
    ' a1, c, b
    
    value(33) = Sqr(value(30) * value(30) - value(32) * value(32))
'    value(33).Enabled = False
    value(34) = Sqr(value(31) * value(31) - value(33) * value(33)) + value(32)
'    value(34).Enabled = False
    Label2 = Round((value(31) / 2) * value(33), 2)
    
'    Command2.Enabled = False
    
    ElseIf value(34) > 0 And value(32) > 0 And value(33) Then
    'a, a1, h
    
    value(30) = Sqr(value(32) * value(32) + value(33) * value(33))
'    value(30).Enabled = False
    a2 = value(34) - value(32)
    value(31) = Sqr(a2 * a2 + value(33) * value(33))
'    value(31).Enabled = False
    Label2 = Round((value(31) / 2) * value(33), 2)
    
'    Command2.Enabled = False
    End If

Case 5
    If value(27) > 0 And value(28) > 0 Then
    value(29) = Sqr(value(28) * value(28) - value(27) * value(27))
    If value(29) > 0 Then Label2 = Round((value(29) * value(27)) / 2, 2)
    ElseIf value(28) > 0 And value(29) > 0 Then
    value(27) = Sqr(value(28) * value(28) - value(29) * value(29))
    If value(27) > 0 Then Label2 = Round((value(29) * value(27)) / 2, 2)
    ElseIf value(29) > 0 And value(27) > 0 Then
    value(28) = Sqr(value(29) * value(29) + value(27) * value(27))
    If value(28) > 0 Then Label2 = Round((value(29) * value(27)) / 2, 2)
    End If

Case 8
    If value(24) > 0 And value(25) > 0 Then
    value(26) = Sqr(value(25) * value(25) - value(24) * value(24))
    If value(26) > 0 Then Label2 = Round((value(24) * value(26)) / 2, 2)
    ElseIf value(25) > 0 And value(26) > 0 Then
    value(24) = Sqr(value(25) * value(25) - value(26) * value(26))
    If value(24) > 0 Then Label2 = Round((value(24) * value(26)) / 2, 2)
    ElseIf value(24) > 0 And value(26) > 0 Then
    value(25) = Sqr(value(24) * value(24) + value(26) * value(26))
    Label2 = Round((value(24) * value(26)) / 2, 2)
    End If

Case 9
    If CInt(value(38)) <= CInt(value(40)) Then value(38).ForeColor = vbRed: value(40).ForeColor = vbRed: Exit Sub
    
    If value(36) > 0 And value(37) > 0 Then
        value(35) = Sqr(value(36) * value(36) - value(37) * value(37))
'        value(35).Enabled = False
    End If
    
    If value(35) > 0 And value(37) > 0 Then
        value(36) = Sqr(value(35) * value(35) + value(37) * value(37))
'        value(36).Enabled = False
    End If
    
    If value(40) > 0 And value(37) > 0 And value(38) > 0 Then
        a2 = value(38) - (CSng(value(37)) + CSng(value(40)))
        Label3 = a2
    End If
    
    If a2 > 0 And (value(35) > 0 Or value(39) > 0) Then
        If value(39) > 0 Then value(35) = Sqr(value(39) * value(39) - a2 * a2) ': value(39).Enabled = False
        If value(35) > 0 Then value(39) = Sqr(value(35) * value(35) + a2 * a2) ': value(35).Enabled = False
        Label2 = Round(0.5 * (CSng(value(38)) + CSng(value(40))) * value(35), 2)
'        Exit Sub
    End If
    
Case 10
    If value(43) > 0 And value(44) > 0 And value(45) > 0 Then
        value(41) = value(43) - value(45)
        sbox = value(41) * value(44)
        value(42) = Sqr(value(45) * value(45) + value(44) * value(44))
    ElseIf value(43) > 0 And value(44) > 0 And value(41) > 0 Then
        value(45) = value(43) - value(44)
        sbox = value(41) * value(44)
        value(42) = Sqr(value(45) * value(45) + value(44) * value(44))
    ElseIf value(42) > 0 And value(43) > 0 And value(45) > 0 Then
        value(41) = value(43) - value(45)
        value(44) = Sqr(value(42) * value(42) - value(45) * value(45))
        sbox = value(44) * value(41)
'    ElseIf value(44) > 0 And value(45) > 0 And value(43) > 0 Then
'        value(43) = Sqr(value(42) * value(42) - value(45) * value(45))
'        value(41) = value(44) - value(45)
'        sbox = value(41) * value(43)
    End If
    
        Label2 = Round(sbox + (value(45) * value(43)) / 2, 2)
    
End Select

Label2 = Format(Label2 / 10000, "### ##0.00#") & " " & m2

Exit Sub
ERR:
MsgBox "ERR: " & ERR.Description, vbCritical, "Err"
Command4.value = True
End Sub

