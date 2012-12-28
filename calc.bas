Attribute VB_Name = "calc"
Option Explicit


Public Function PointsSort(ByRef Lape_Lines, ByRef Lape_Lines_out, N_Slope, Amount_PA, Amount_PB, Acount, Bcount) As Boolean
On Error GoTo ERR

    Dim ResultCurLine As Integer
    Dim Abegin As Integer
    Dim Bbegin As Integer
    Dim CurBPoint As Integer
    Dim NewArr_Lape_Lines() As Integer ' Свойства линий

    '
    ' Сортировка по привязкам
    '
    ReDim NewArr_Lape_Lines(1 To SlP(N_Slope).CountOfLines, 1)

    ResultCurLine = 1
    CurBPoint = 2
    Abegin = 0
    Bbegin = 0
    Dim P As Integer
    Dim CurA, CurB As String
    Dim check_str_array As String
    check_str_array = ""
    CurA = ""
    CurB = ""
    P = 0
    Do While (CurBPoint <> SlP(N_Slope).CountOfLines + 1)
    
continue:
        P = P + 1

        If Len(Lape_Lines(N_Slope, P, 0)) = 1 Then
            CurA = "0" & Lape_Lines(N_Slope, P, 0)
        Else
            CurA = Lape_Lines(N_Slope, P, 0)
        End If
        
        If Len(Lape_Lines(N_Slope, P, 1)) = 1 Then
            CurB = "0" & Lape_Lines(N_Slope, P, 1)
        Else
            CurB = Lape_Lines(N_Slope, P, 1)
        End If
        
        If Abegin = 0 Then
        If Lape_Lines(N_Slope, P, 0) <> Lape_Lines(N_Slope, P, 1) Then
            Abegin = Lape_Lines(N_Slope, P, 0)
            Bbegin = Lape_Lines(N_Slope, P, 1)
            NewArr_Lape_Lines(1, 0) = Abegin
            NewArr_Lape_Lines(1, 1) = Bbegin
            check_str_array = check_str_array & CurB & "," & CurA & " " & CurA & "," & CurB & " "
        Else
            GoTo continue
        End If
        End If

        If (Bbegin = Lape_Lines(N_Slope, P, 0) And (ResultCurLine <> P)) Then
    
            ' Отсеивание дубликатов
            ' отсев отдельно стоящих точек (1,1)
            If InStr(check_str_array, CurA & "," & CurB) = 0 And _
            InStr(check_str_array, CurB & "," & CurA) = 0 And _
            Lape_Lines(N_Slope, P, 0) <> Lape_Lines(N_Slope, P, 1) Then
        
                NewArr_Lape_Lines(CurBPoint, 0) = Lape_Lines(N_Slope, P, 0)
                NewArr_Lape_Lines(CurBPoint, 1) = Lape_Lines(N_Slope, P, 1)
                
                check_str_array = check_str_array & CurA & "," & CurB & " "
                
                If NewArr_Lape_Lines(CurBPoint, 1) = Abegin Then
                Exit Do
                End If
                                
                ResultCurLine = P
                Bbegin = NewArr_Lape_Lines(CurBPoint, 1)
                CurBPoint = CurBPoint + 1
                P = 0
                
            End If
                    
        ElseIf (Bbegin = Lape_Lines(N_Slope, P, 1) And ResultCurLine <> P) Then
    
            ' Отсеивание дубликатов
            ' отсев отдельно стоящих точек (1,1)
            If InStr(check_str_array, CurB & "," & CurA) = 0 And _
            InStr(check_str_array, CurA & "," & CurB) = 0 And _
            Lape_Lines(N_Slope, P, 0) <> Lape_Lines(N_Slope, P, 1) Then
            
                NewArr_Lape_Lines(CurBPoint, 0) = Lape_Lines(N_Slope, P, 1)
                NewArr_Lape_Lines(CurBPoint, 1) = Lape_Lines(N_Slope, P, 0)
                
                check_str_array = check_str_array & CurB & "," & CurA & " "
                
                If NewArr_Lape_Lines(CurBPoint, 1) = Abegin Then
                Exit Do
                End If
                
                ResultCurLine = P
                Bbegin = NewArr_Lape_Lines(CurBPoint, 1)
                CurBPoint = CurBPoint + 1
                P = 0
                
            End If
       
        End If

    Loop
    
    '
    ' Поиск фигуры
    '
    Dim PStart As Integer

    PStart = NewArr_Lape_Lines(1, 0)

    For P = 1 To ArraySize(NewArr_Lape_Lines) - 1 Step 1
        Lape_Lines_out(N_Slope, P, 0) = NewArr_Lape_Lines(P, 0): Acount = Acount + 1
        Lape_Lines_out(N_Slope, P, 1) = NewArr_Lape_Lines(P, 1): Bcount = Bcount + 1
        If PStart = NewArr_Lape_Lines(P, 1) Then
            PointsSort = True
            Exit For
        End If
    Next
    
Exit Function
ERR:
    PointsSort = False
End Function

    
Public Function ConvertData(ByVal value, Optional From_sm As Boolean = False) As String
    If value = 0 Then
        ConvertData = 0
        Exit Function
    End If
    If IsNumeric(value) = False Then
        value = Val(value)
    End If
    If From_sm Then
        Select Case Setup.Combo4.ListIndex
        Case 0 ' - см
            value = Round(value) ' Один к одному
        Case 2 ' - мм
    '        value = value * 1000
        Case 1 ' - метры
            value = value * 100
            value = Round(value)
        End Select
    Else
        Select Case Setup.Combo4.ListIndex
        Case 0 ' - см
            value = Round(value) ' Один к одному
        Case 2 ' - метры
    '        value = value / 1000
        Case 1 ' - метры
            value = value / 100
            value = Format$(Round(value, 2), "0.00")
        End Select
    End If
    ConvertData = value
End Function



Public Function PolygonArea(n, CurrentSlope, Lape_Lines, Lape_Points_X, Lape_Points_Y) As Double

    Dim Result As Double
    Dim i As Long
    Dim X() As Double
    Dim Y() As Double

        ReDim X(0# To n)
        ReDim Y(0# To n)
    
        For i = 1# To n Step 1
            X(i) = Lape_Points_X(CurrentSlope, Lape_Lines(CurrentSlope, i, 0))
            Y(i) = Lape_Points_Y(CurrentSlope, Lape_Lines(CurrentSlope, i, 0))
        Next i

        X(0#) = X(n)
        Y(0#) = Y(n)
        Result = 0#
        i = 0#
    
        Do
            Result = Result + (X(i) + X(i + 1#)) * (Y(i) - Y(i + 1#))
            i = i + 1#
        Loop Until Not i <= n - 1#
    
        Result = 0.5 * Abs(Result)
        PolygonArea = Result
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Положение точки относительно многоугольника.
'function IsPointInPolygon(
'    const   x   :   Real;
'    const   y   :   Real;
'    const   N   :   Integer;
'    const   XPO :   array [1..N] of Real;
'    const   YPO :   array [1..N] of Real):Boolean;
'
'Проверяет точку на принадлежность внутренней области многоугольника.
'На границе значение функции не определено.
'
'Параметры:
'    x,y - точка
'    XP, YP - массив вершин многоугольника.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Public Function IsPointInPolygon(ByRef x As Double, ByRef y As Double, ByRef n As Long, ByRef XPO() As Double, ByRef YPO() As Double) As Boolean
'    Dim Result As Boolean
'    Dim i As Long
'    Dim XPI() As Double
'    Dim YPI() As Double
'    Dim b1 As Boolean
'    Dim b2 As Boolean
'
'        ReDim XPI(0# To n)
'        ReDim YPI(0# To n)
'        For i = 1# To n Step 1
'            XPI(i) = XPO(i)
'            YPI(i) = YPO(i)
'        Next i
'
'        XPI(0#) = XPI(n)
'        YPI(0#) = YPI(n)
'        i = 0#
'        Result = False
'        Do
'            b1 = y > YPI(i)
'            b2 = y <= YPI(i + 1#)
'            If Not (b1 And Not b2 Or b2 And Not b1) Then
'                If x - XPI(i) < (y - YPI(i)) * (XPI(i + 1#) - XPI(i)) / (YPI(i + 1#) - YPI(i)) Then
'                    Result = Not Result
'                End If
'
'            End If
'
'            i = i + 1#
'        Loop Until Not i <= n - 1#
'
'        IsPointInPolygon = Result
'End Function


'Public Function isFindConnect(A As Integer, B As Integer) As Boolean
'    Dim i As Integer
'    For i = 1 To SlP(N_Slope).CountOfLines
'        If Lape_Lines(N_Slope, i, 0) = B Or Lape_Lines(N_Slope, i, 1) = B Then
'            Exit For
'        End If
'    Next
'End Function

'    Dim i As Integer
'    For i = 1 To SlP(N_Slope).CountOfLines
'        If POINT = Lape_Lines(N_Slope, i, 0) Then Find_Connect = i: Exit For
'    Next
'
'    For i = 1 To SlP(N_Slope).CountOfPoints
'        If POINT = Lape_Lines(N_Slope, i, 1) Then Find_Connect = i: Exit For
'    Next
'    For i = 1 To SlP(N_Slope).CountOfLines
'        If Lape_Lines(N_Slope, A, 0) = B Or Lape_Lines(N_Slope, A, 1) = B Then
'            isFindConnect = True
'            Exit For
'        End If
'    Next
'End Function


Public Function FindPointoint(X As Single, Y As Single, Optional Distance As Single) As Integer
    Dim l00D2 As Single
    Dim i As Integer
    Dim l00E2 As Single
    
    FindPointoint = 0
    l00D2 = 99999
    
    Dim lastP As Integer
    lastP = SlP(N_Slope).CountOfLines
    If SlP(N_Slope).CountOfPoints > lastP Then lastP = SlP(N_Slope).CountOfPoints
    
    For i = 1 To lastP Step 1
        l00E2 = Sqr((X - Lape_Points_X(N_Slope, i)) ^ 2 + (Y - Lape_Points_Y(N_Slope, i)) ^ 2)
        If l00E2 < l00D2 Then
            l00D2 = l00E2
            If Distance = 0 Or Distance >= l00E2 Then
                FindPointoint = i
            End If

        End If

    Next i

End Function


Public Function FindBindingPoints(a As Integer, b As Integer) As Boolean
    Dim i As Integer
    Dim maxpoints As Integer
    maxpoints = SlP(N_Slope).CountOfLines - 1
'    If SlP(N_Slope).CountOfPoints > maxpoints Then maxpoints = SlP(N_Slope).CountOfPoints
    
    For i = 1 To maxpoints
        If Lape_Lines(N_Slope, i, 0) = a And Lape_Lines(N_Slope, i, 1) = b Then
            FindBindingPoints = True
          Exit For
        End If
        If Lape_Lines(N_Slope, i, 0) = b And Lape_Lines(N_Slope, i, 1) = a Then
            FindBindingPoints = True
            Exit For
        End If
    Next
End Function

Public Function ArcCos(X As Double) As Double
    If X <> 1 Then
        ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    Else
        ArcCos = 0
    End If
End Function
