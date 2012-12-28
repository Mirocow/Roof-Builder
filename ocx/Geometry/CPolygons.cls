VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPolygons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GreatePolygon(m_Lines As CLines) As CPolygon
Dim CL As CLine
Dim A As CPoint
Dim B As CPoint
For Each CL In m_Lines.Lines
    Set A = CL.BeginPoint
    Set B = CL.EndPoint
Next
    Set A = Nothing
    Set B = Nothing
End Function

Public Function AddLine(ByVal NewLine As CLine) As CLine
    Lines.Add NewLine, NewLine.Key
    Set AddLine = NewLine
End Function

'Public Function PointsSort(ByRef Lape_Lines, ByRef Lape_Lines_out, N_Slope, Amount_PA, Amount_PB, Acount, Bcount) As Boolean
'On Error GoTo ERR
'
'    Dim ResultCurLine As Integer
'    Dim Abegin As Integer
'    Dim Bbegin As Integer
'    Dim CurBPoint As Integer
'    Dim NewArr_Lape_Lines() As Integer ' Свойства линий
'
'    '
'    ' Сортировка по привязкам
'    '
'    ReDim NewArr_Lape_Lines(1 To SlP(N_Slope).CountOfLines, 1)
'
'    ResultCurLine = 1
'    CurBPoint = 2
'    Abegin = 0
'    Bbegin = 0
'    Dim P As Integer
'    Dim CurA, CurB As String
'    Dim check_str_array As String
'    check_str_array = ""
'    CurA = ""
'    CurB = ""
'    P = 0
'    Do While (CurBPoint <> ArraySize(Lape_Lines) - 1)
'
'continue:
'        P = P + 1
'
'        If Len(Lape_Lines(N_Slope, P, 0)) = 1 Then
'            CurA = "0" & Lape_Lines(N_Slope, P, 0)
'        Else
'            CurA = Lape_Lines(N_Slope, P, 0)
'        End If
'
'        If Len(Lape_Lines(N_Slope, P, 1)) = 1 Then
'            CurB = "0" & Lape_Lines(N_Slope, P, 1)
'        Else
'            CurB = Lape_Lines(N_Slope, P, 1)
'        End If
'
'        If Abegin = 0 Then
'        If Lape_Lines(N_Slope, P, 0) <> Lape_Lines(N_Slope, P, 1) Then
'            Abegin = Lape_Lines(N_Slope, P, 0)
'            Bbegin = Lape_Lines(N_Slope, P, 1)
'            NewArr_Lape_Lines(1, 0) = Abegin
'            NewArr_Lape_Lines(1, 1) = Bbegin
'            check_str_array = check_str_array & CurB & "," & CurA & " " & CurA & "," & CurB & " "
'        Else
'            GoTo continue
'        End If
'        End If
'
'        If (Bbegin = Lape_Lines(N_Slope, P, 0) And (ResultCurLine <> P)) Then
'
'            ' Отсеивание дубликатов
'            ' отсев отдельно стоящих точек (1,1)
'            If InStr(check_str_array, CurA & "," & CurB) = 0 And _
'            InStr(check_str_array, CurB & "," & CurA) = 0 And _
'            Lape_Lines(N_Slope, P, 0) <> Lape_Lines(N_Slope, P, 1) Then
'
'                NewArr_Lape_Lines(CurBPoint, 0) = Lape_Lines(N_Slope, P, 0)
'                NewArr_Lape_Lines(CurBPoint, 1) = Lape_Lines(N_Slope, P, 1)
'
'                check_str_array = check_str_array & CurA & "," & CurB & " "
'
'                If NewArr_Lape_Lines(CurBPoint, 1) = Abegin Then
'                Exit Do
'                End If
'
'                ResultCurLine = P
'                Bbegin = NewArr_Lape_Lines(CurBPoint, 1)
'                CurBPoint = CurBPoint + 1
'                P = 0
'
'            End If
'
'        ElseIf (Bbegin = Lape_Lines(N_Slope, P, 1) And ResultCurLine <> P) Then
'
'            ' Отсеивание дубликатов
'            ' отсев отдельно стоящих точек (1,1)
'            If InStr(check_str_array, CurB & "," & CurA) = 0 And _
'            InStr(check_str_array, CurA & "," & CurB) = 0 And _
'            Lape_Lines(N_Slope, P, 0) <> Lape_Lines(N_Slope, P, 1) Then
'
'                NewArr_Lape_Lines(CurBPoint, 0) = Lape_Lines(N_Slope, P, 1)
'                NewArr_Lape_Lines(CurBPoint, 1) = Lape_Lines(N_Slope, P, 0)
'
'                check_str_array = check_str_array & CurB & "," & CurA & " "
'
'                If NewArr_Lape_Lines(CurBPoint, 1) = Abegin Then
'                Exit Do
'                End If
'
'                ResultCurLine = P
'                Bbegin = NewArr_Lape_Lines(CurBPoint, 1)
'                CurBPoint = CurBPoint + 1
'                P = 0
'
'            End If
'
'        End If
'
'    Loop
'
'    '
'    ' Поиск фигуры
'    '
'    Dim PStart As Integer
'
'    PStart = NewArr_Lape_Lines(1, 0)
'
'    For P = 1 To ArraySize(NewArr_Lape_Lines) - 1 Step 1
'        Lape_Lines_out(N_Slope, P, 0) = NewArr_Lape_Lines(P, 0): Acount = Acount + 1
'        Lape_Lines_out(N_Slope, P, 1) = NewArr_Lape_Lines(P, 1): Bcount = Bcount + 1
'        If PStart = NewArr_Lape_Lines(P, 1) Then
'            PointsSort = True
'            Exit For
'        End If
'    Next
'
'Exit Function
'ERR:
'    PointsSort = False
'End Function



