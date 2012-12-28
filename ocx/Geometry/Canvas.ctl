VERSION 5.00
Begin VB.UserControl Canvas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      Visible         =   0   'False
      X1              =   72
      X2              =   224
      Y1              =   176
      Y2              =   72
   End
End
Attribute VB_Name = "Canvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public X As Single ' Положение курсора
Public Y As Single

' Найденые
Public FindPoint As CPoint
Public FindLine As CLine

' Текущие
Public CurLine As CLine ' Линия которая рисуется в данный момент или последняя нарисованная
Public CurPoligon As CPolygon ' Класс содержащий последний созданный полигон

' Основные классы содержащие данные
Private m_Lines As CLines ' Класс содержащей линии и отдельно стоящие точки
Private m_Polygons As CPolygons ' Класс содержащий полигоны

Private m_CurrentPoint As CPoint
Private PointsIndexes As Integer

'''''''''''''''''''''''''''''''''
'''''''' Canvas
'''''''''''''''''''''''''''''''''

''''''''''''''''''''' PUBLIC
Public Sub Clear()
UserControl.Cls
m_Lines.ClearLines
Line1.Visible = False
Set FindLine = Nothing
PointsIndexes = 0
End Sub

Public Function LinesCount() As Integer
LinesCount = m_Lines.LinesCount
End Function

Public Function PointsCount() As Integer
PointsCount = m_Lines.PointsCount
End Function

Public Function GetLines() As CLines
Set GetLines = m_Lines
End Function


'''''''''''''''''''' PRIVATE

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim P As CPoint

' Поиск линии
If Button = 1 Then
    Set FindLine = m_Lines.Find_Line(X, Y) ' Поиск линий
Else
    Set FindLine = Nothing
End If

Set FindPoint = Nothing

If FindLine Is Nothing Then
    ' Поиск точки
    If Button = 1 Then
        Set FindPoint = m_Lines.FindPoint(X, Y, 20)
    Else
        Set FindPoint = m_Lines.FindPoint(X, Y, 50)
    End If
End If

If Line1.Visible = False Then ' рисование не начато
Set CurLine = New CLine

    If Button = 1 Then
        If Not FindLine Is Nothing Then
            
            ' Разделяем найденую линию и привязываемся к точке разделения
            Set FindPoint = m_Lines.Divide_Line(X, Y, FindLine)
            PointsIndexes = PointsIndexes + 1
            FindPoint.Key = PointsIndexes
            GoTo CONNECT1
            
        ElseIf FindPoint Is Nothing Then
            
            PointsIndexes = PointsIndexes + 1
            
            ' Добавление стартовой точки
            Set P = New CPoint
            P.X = X
            P.Y = Y
            P.Key = PointsIndexes
            CurLine.BeginPoint = P
            Line1.Visible = True
            Set P = Nothing
            
            Line1.Visible = True
            Line1.X1 = X
            Line1.Y1 = Y
            Line1.X2 = X
            Line1.Y2 = Y
            
        Else
         GoTo CONNECT1
        End If
    Else
        If Not FindPoint Is Nothing Then
CONNECT1:
            ' Привязываемся к точке
            CurLine.BeginPoint = FindPoint
            X = FindPoint.X
            Y = FindPoint.Y
            Set FindPoint = Nothing
            
            Line1.Visible = True
            Line1.X1 = X
            Line1.Y1 = Y
            Line1.X2 = X
            Line1.Y2 = Y
    
        End If
    End If
    
Else ' Выполняется рисование
    
    If Button = 1 Then
    
        If Not FindLine Is Nothing Then
            
            ' Разделяем найденую линию и привязываемся к точке разделения
            Set FindPoint = m_Lines.Divide_Line(X, Y, FindLine)
            PointsIndexes = PointsIndexes + 1
            FindPoint.Key = PointsIndexes
            GoTo CONNECT2
            
        ElseIf FindPoint Is Nothing Then
            
            Set P = New CPoint
            P.X = X
            P.Y = Y
            P.isPoint = True
            PointsIndexes = PointsIndexes + 1
            P.Key = PointsIndexes
            CurLine.EndPoint = P
            Set P = Nothing
            
            ' Проработать
            
            ' Если привязываемся к отдельно стоящей точке, то удаляем ее как линию
'            If CurLine.BeginPoint.isPoint Then
                ' Исправляем существующую точку
'                m_Lines.ReplaceLine CurLine.BeginPoint.Key, CurLine
'            Else
                ' Добавляем новую линию
                CurLine.BeginPoint.isPoint = True
                m_Lines.AddLine CurLine
'            End If
            
            Line1.Visible = False
        Else
            GoTo CONNECT2
        End If
        
    Else
        If Not FindPoint Is Nothing Then
CONNECT2:
            ' Выполняем поиск к чему привязаться когда линия уже создана
            CurLine.EndPoint = FindPoint
            If CurLine.BeginPoint.Key <> CurLine.EndPoint.Key Then
                m_Lines.AddLine CurLine
            End If
            Set FindPoint = Nothing
            
            ' Выполняем поиск точки привязки у уже имеющихся полиговнов
            
            ' Set CurPoligon = New CPolygon
        Else
            If CurLine.BeginPoint.Children = 0 Then
            CurLine.BeginPoint.isPoint = True
            ' Рисуем отдельно стоящую точку
            CurLine.EndPoint = CurLine.BeginPoint
            m_Lines.AddLine CurLine
            End If
            Set CurLine = Nothing
        End If
        Line1.Visible = False
    End If
    
'    Set CurLine = Nothing
    
End If

Draw
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.X = X
Me.Y = Y

If m_Lines.PointsCount > 0 Then
    Dim P As CPoint
    Set P = m_Lines.FindPoint(X, Y, 20)
    If Not P Is Nothing Then
        Set FindPoint = P
    Else

        Set FindLine = m_Lines.Find_Line(X, Y)
        Set FindPoint = Nothing

    End If
End If

If Line1.Visible = True Then
    Line1.X2 = X
    Line1.Y2 = Y
End If
End Sub

Private Sub DrawPoint()
Dim L As CLine
Dim P1 As CPoint
Dim P2 As CPoint
If Line1.Visible Then
    Set P1 = CurLine.BeginPoint
    If Not P1 Is Nothing Then UserControl.PSet (P1.X, P1.Y), vbBlack
End If
For Each L In m_Lines.Lines
    Set P1 = L.BeginPoint
    Set P2 = L.EndPoint
    UserControl.PSet (P1.X, P1.Y), vbBlack
    UserControl.PSet (P2.X, P2.Y), vbBlack
Next
    Set P1 = Nothing
    Set P2 = Nothing
End Sub

Private Sub DrawLines()
Dim i As Integer
Dim Llen As Single
Dim L As CLine
Dim P1 As CPoint
Dim P2 As CPoint
Dim coef As Single
coef = UserControl.ScaleWidth / 90

For Each L In m_Lines.Lines

    Set P1 = L.BeginPoint
    Set P2 = L.EndPoint
    
    If P1.Key <> P2.Key Then
        UserControl.Line (P2.X, P2.Y)-(P1.X, P1.Y), vbRed
        
        Llen = Format(Sqr((P1.X - P2.X) ^ 2 + (P1.Y - P2.Y) ^ 2), "###.0")
        
        If Llen > 0 Then
        UserControl.PSet ((P1.X + P2.X) / 2, ((P1.Y + P2.Y) / 2) - coef), UserControl.BackColor
            
        'UserControl.DrawWidth = 1
        'UserControl.FontSize = FontSize
        'UserControl.FontBold = True
        UserControl.Print Llen
        End If
    End If
    
Next
    Set P1 = Nothing
    Set P2 = Nothing
End Sub

Public Sub RemoveLine(Key As String)
m_Lines.RemoveLine Key
m_Lines.RemoveChild Key
End Sub

'Removes a Point from the polygon
Public Sub RemovePoint(Key As String)
On Error GoTo ERR
Dim CL As CLine
Dim A As CPoint
Dim B As CPoint
For Each CL In m_Lines.Lines
    
    ' поиск и удаление
    Set A = CL.BeginPoint
    Set B = CL.EndPoint
    
    ' Удаляем принадлежаще точкам линии
    If A.Key = Key Or B.Key = Key Then
        m_Lines.RemoveLine CL.Key
        m_Lines.RemoveChild CL.Key
    End If
    
Next
Set A = Nothing
Set B = Nothing

Exit Sub

ERR:
End Sub

Private Sub Draw()
UserControl.Cls
UserControl.DrawWidth = 3
DrawPoint
UserControl.DrawWidth = 1
DrawLines
End Sub

Public Sub ReDraw()
'UserControl.Cls
UserControl.DrawWidth = 3
DrawPoint
UserControl.DrawWidth = 1
DrawLines
End Sub

Public Sub Cls()
UserControl.Cls
End Sub

Private Sub UserControl_Initialize()
Set m_Lines = New CLines
End Sub

Private Sub UserControl_Terminate()
Set m_Lines = Nothing
Set m_CurrentPoint = Nothing
End Sub
