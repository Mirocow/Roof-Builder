Attribute VB_Name = "math"
''''''''''''''''''''''''''''''''''''''''''''''''''
'Этот модуль сгенерирован транслятором AlgoPascal.
''''''''''''''''''''''''''''''''''''''''''''''''''
'Эти подпрограммы должны быть определены программистом
' Function IsPointOnLine(ByVal x As Double, _
'         ByVal y As Double, _
'         ByVal z As Double, _
'         ByRef XL() As Double, _
'         ByRef YL() As Double, _
'         ByRef ZL() As Double, _
'         ByVal Epsilon As Double) As Boolean


'Подпрограммы
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Проверка принадлежит ли точка плоскости.
'
'function IsPointOnPlane(
'    x,y,z:real;
'    XP,YP,ZP:array [1..3] of real;
'    Epsilon:Real):Boolean
'
'(x,y,z)-coordinates of point
'(XP[i],YP[i],ZP[i])-coordinates of 3 points, which define the plane
'This function uses IsPointOnLine
'
'Epsilon - допустимая погрешность
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPointOnPlane(ByVal X As Double, _
         ByVal Y As Double, _
         ByVal z As Double, _
         ByRef xp() As Double, _
         ByRef yp() As Double, _
         ByRef ZP() As Double, _
         ByVal Epsilon As Double) As Boolean
    Dim Result As Boolean
    Dim i As Long
    Dim A As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim XL() As Double
    Dim YL() As Double
    Dim ZL() As Double

    ReDim XL(1 To 2)
    ReDim YL(1 To 2)
    ReDim ZL(1 To 2)
    If Abs(xp(1) - xp(2)) < Epsilon And Abs(yp(1) - yp(2)) < Epsilon And Abs(ZP(1) - ZP(2)) < Epsilon Or Abs(xp(1) - xp(3)) < Epsilon And Abs(yp(1) - yp(3)) < Epsilon And Abs(ZP(1) - ZP(3)) < Epsilon Or Abs(xp(2) - xp(3)) < Epsilon And Abs(yp(2) - yp(3)) < Epsilon And Abs(ZP(2) - ZP(3)) < Epsilon Then
        Result = False
    Else
        i = 1
        Do
            XL(i) = xp(i)
            YL(i) = yp(i)
            ZL(i) = ZP(i)
            i = i + 1
        Loop Until Not i <= 2
        If IsPointOnLine(xp(3), yp(3), ZP(3), XL, YL, ZL, Epsilon) Then
            Result = False
        Else
            A = (yp(2) - yp(1)) * (ZP(3) - ZP(1)) - (yp(3) - yp(1)) * (ZP(2) - ZP(1))
            b = (xp(3) - xp(1)) * (ZP(2) - ZP(1)) - (xp(2) - xp(1)) * (ZP(3) - ZP(1))
            c = (xp(2) - xp(1)) * (yp(3) - yp(1)) - (xp(3) - xp(1)) * (yp(2) - yp(1))
            d = A * xp(1) + b * yp(1) + c * ZP(1)
            Result = Abs(A * X + b * Y + c * z - d) < Epsilon
        End If
    End If

    IsPointOnPlane = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
'Этот модуль сгенерирован транслятором AlgoPascal.
''''''''''''''''''''''''''''''''''''''''''''''''''
'Подпрограммы
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Площадь многоугольника
'
'Входные параметры:
'    N   -   число точек
'    XO  -   массив абсцисс в порядке обхода
'            нумерация элементов [1..N]
'    YO  -   массив ординат в порядке обхода
'            нумерация элементов [1..N]
'
'Результат:
'    Площадь многоугольника, заданного указанными точками
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PolygonArea(ByVal n As Long, _
         ByRef XO() As Double, _
         ByRef YO() As Double) As Double
    Dim Result As Double
    Dim i As Long
    Dim X() As Double
    Dim Y() As Double

    ReDim X(0 To n)
    ReDim Y(0 To n)
    For i = 1 To n Step 1
        X(i) = XO(i)
        Y(i) = YO(i)
    Next i
    X(0) = X(n)
    Y(0) = Y(n)
    Result = 0
    i = 0
    Do
        Result = Result + (X(i) + X(i + 1)) * (Y(i) - Y(i + 1))
        i = i + 1
    Loop Until Not i <= n - 1
    Result = 0.5 * Abs(Result)

    PolygonArea = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
'Этот модуль сгенерирован транслятором AlgoPascal.
''''''''''''''''''''''''''''''''''''''''''''''''''
'Подпрограммы
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Проверка принадлежит ли точка прямой.
'
'function IsPointOnLine(
'    x,y,z:Real;
'    XL,YL,ZL:array[1..2] of Real;
'    Epsilon : Real):Boolean;
'
'(x,y,z)-координаты точки
'(XL[i],YL[i],ZL[i])-координаты двух точек прямой
'Epsilon - допустимая погрешность
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPointOnLine(ByVal X As Double, _
         ByVal Y As Double, _
         ByVal z As Double, _
         ByRef XL() As Double, _
         ByRef YL() As Double, _
         ByRef ZL() As Double, _
         ByVal Epsilon As Double) As Boolean
    Dim Result As Boolean
    Dim t As Double

    If Abs(XL(1) - XL(2)) < Epsilon And Abs(YL(1) - YL(2)) < Epsilon And Abs(ZL(1) - ZL(2)) < Epsilon Then
        Result = False
    Else
        If Abs(XL(1) - XL(2)) < Epsilon Then
            Result = Abs(X - XL(1)) < Epsilon
            If Abs(YL(1) - YL(2)) < Epsilon Then
                Result = Result And Abs(Y - YL(1)) < Epsilon
            Else
                Result = Result And Abs((z - ZL(1)) * (YL(2) - YL(1)) - (ZL(2) - ZL(1)) * (Y - YL(1))) < Epsilon
            End If
        Else
            t = (X - XL(1)) / (XL(2) - XL(1))
            Result = Abs(z - (ZL(1) + t * (ZL(2) - ZL(1)))) < Epsilon
            Result = Result And Abs(Y - (YL(1) + t * (YL(2) - YL(1)))) < Epsilon
        End If
    End If

    IsPointOnLine = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
'Этот модуль сгенерирован транслятором AlgoPascal.
''''''''''''''''''''''''''''''''''''''''''''''''''
'Подпрограммы
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
Public Function IsPointInPolygon(ByVal X As Double, _
         ByVal Y As Double, _
         ByVal n As Long, _
         ByRef XPO() As Double, _
         ByRef YPO() As Double) As Boolean
    Dim Result As Boolean
    Dim i As Long
    Dim XPI() As Double
    Dim YPI() As Double

    ReDim XPI(0 To n)
    ReDim YPI(0 To n)
    For i = 1 To n Step 1
        XPI(i) = XPO(i)
        YPI(i) = YPO(i)
    Next i
    XPI(0) = XPI(n)
    YPI(0) = YPI(n)
    i = 0
    Result = False
    Do
        If Not (Y > YPI(i) Xor Y <= YPI(i + 1)) Then
            If X - XPI(i) < (Y - YPI(i)) * (XPI(i + 1) - XPI(i)) / (YPI(i + 1) - YPI(i)) Then
                Result = Not Result
            End If
        End If
        i = i + 1
    Loop Until Not i <= n - 1

    IsPointInPolygon = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
'Этот модуль сгенерирован транслятором AlgoPascal.
''''''''''''''''''''''''''''''''''''''''''''''''''
'Подпрограммы
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Положение точки относительно выпуклого многоугольника.
'function IsPointInConvexPolygon(
'    const   xp  :   Real;
'    const   yp  :   Real;
'    const   N   :   Integer;
'    const   XO  :   array [1..N] Real;
'    const   YO  :   array [1..N] Real;
'    const   Epsilon : Real):Boolean;
'
'Проверяет точку на принадлежность внутренней области многоугольника.
'На границе значение функции не определено.
'
'Параметры:
'    xp,yp - точка
'    X0, Y0 - массив вершин многоугольника - с учетом порядка обхода.
'    Epsilon - погрешность
'     сравнения
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPointInConvexPolygon(ByVal xp As Double, _
         ByVal yp As Double, _
         ByVal n As Long, _
         ByRef XO() As Double, _
         ByRef YO() As Double, _
         ByVal Epsilon As Double) As Boolean
    Dim Result As Boolean
    Dim i As Long
    Dim cp As Long
    Dim CM As Long
    Dim d As Double
    Dim X() As Double
    Dim Y() As Double

    ReDim X(0 To n)
    ReDim Y(0 To n)
    For i = 1 To n Step 1
        X(i) = XO(i)
        Y(i) = YO(i)
    Next i
    X(0) = X(n)
    Y(0) = Y(n)
    i = 0
    cp = 0
    CM = 0
    Do
        d = X(i) * yp + xp * Y(i + 1) + X(i + 1) * Y(i) - X(i + 1) * yp - xp * Y(i) - X(i) * Y(i + 1)
        If d > 0 Then
            cp = cp + 1
        Else
            CM = CM + 1
        End If
        i = i + 1
    Loop Until Not (i < n And Abs(d) > Epsilon)
    Result = Abs(d) > Epsilon And (cp = n Or CM = n)

    IsPointInConvexPolygon = Result
End Function

'Быстрая проверка пересечения двух прямоугольников на плоскости
'где (первый прямоугольник n-ый)
'_x = (*).Location.X;
'_y = (*).Location.Y;
'_xs =(*).Size.Width;
'_ys =(*).Size.Height;
'
'где (второй прямоугольник)
'_obj_start
'
'
'if((_x>=_obj_start.Location.X-_xs)&(_x<=_obj_start.Location.X+_obj_start.Size.Width))
'if((_y<_obj_start.Location.Y+_obj_start.Size.Height)&(_y+_ys>_obj_start.Location.Y))
'{они пересеклись}


'DECLARE SUB CircleTestXY (xyd!(), Np%, x0!, y0!, kz%)
'DECLARE SUB CircleSquare (xyd!(), Np%, Square!)
'DefInt I-N
'**************************************************
'  Модуль XY_TESTC.BAS
'
' Процедуры:
' CircleTestXY - определение местоположения точки
' относительно фигуры-многоугольника
' CircleSquare - вычисление площади многоугольника
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''"""""""""""""""""""""""""""""""""
' тестовый пример использования функций
  Np = 6: Dim xyd(Np, 2)   ' массив для пятиугольника
  xyp(1, 1) = 10: xyp(2, 1) = 20
  xyp(1, 2) = 0: xyp(2, 2) = 10
  xyp(1, 3) = -10: xyp(2, 3) = 20
  xyp(1, 4) = -10: xyp(2, 4) = -20
  xyp(1, 5) = 10: xyp(2, 5) = -20
  xyp(1, Np) = xyp(1, 1): xyp(2, Np) = xyp(2, 1)
  ' вычисление площади многоугольника
  Call CircleSquare(xyp(), Np, Square)
  ' проверка - где находится заданная точка?
  x0 = 0: y0 = 0   ' координаты тестируемой точки
  Call CircleTestXY(xyp(), Np, x0, y0, kz)
  Print "kz, Square = "; kz; Square
'End

Sub CircleSquare(xyd(), Np, Square)
  ' Вычисление площади многоугольника
  '————————————————————————————————
  ' ВХОД:
  ' xyd() - массив координат углов многоугольника
  ' x = xyd(1,i), y = xyd(2,i) ; i = 1 to Np
  '  (Np-1) - количество узлов
  '  координаты 1-й точки = координатам N-й
  '
  ' ВЫХОД: Square - площадь многоугольника
  '''''''''''''''''''''''''''''''''''''''''''''''""""""""""""""""""""""""""""""""""
  Const pi = 3.141593
  Square = 0
  For k = 1 To Np  ' Np + 1
    X2 = xyd(1, k): Y2 = xyd(2, k)
    v2 = Sqr(X2 * X2 + Y2 * Y2)
    ay2 = Abs(Y2): ax2 = Abs(X2)
    If ax2 * 10000 > ay2 Then
      alfa2 = Atn(ay2 / ax2)
    Else: alfa2 = pi * 0.5
    End If
    If X2 < 0 Then alfa2 = pi - alfa2
    If Y2 < 0 Then alfa2 = -alfa2
    If k > 1 Then   ' проверка перехода
      Square = Square + 0.5 * Sin(alfa2 - alfa1) * v1 * v2
    End If
    X1 = X2: Y1 = Y2: v1 = v2: alfa1 = alfa2
  Next
End Sub

Sub CircleTestXY(xyd(), Np, x0, y0, kz)
  '
  ' Проверка местонахождения точки на плоскости
  ' относительно многоугольника - внутри или снаружи
  '————————————————————————-
  ' ВХОД:
  '  xyd() - массив координат углов многоугольника
  '  x = xyd(1,i), y = xyd(2,i) ; i = 1 to Np
  '  (Np-1) - количество узлов
  '  координаты 1-й точки = координатам N-й точки
  '  x0,y0  - координаты тестируемой точки
  '
  ' ВЫХОД:  положение тестируемой точки
  ' kz = 0  - вне
  '      = -100  - на границе
  '      = -4  - внутри (обход по часовой стрелке)
  '      =  4   - внутри (против часовой стрелки)
  ''''''''''''''''''''''''''
  kz = 0
  For k = 1 To Np   ' Np + 1
    ' IF l > Np THEN k = 1 ELSE k = l
    X2 = xyd(1, k) - x0: Y2 = xyd(2, k) - y0
    '
    ' проверка четверти плоскости
    kv2 = 0
    If X2 >= 0 And Y2 > 0 Then kv2 = 1
    If X2 < 0 And Y2 >= 0 Then kv2 = 2
    If X2 <= 0 And Y2 < 0 Then kv2 = 3
    If X2 > 0 And Y2 <= 0 Then kv2 = 4
    If kv2 = 0 Then kz = -100: Exit For
    '
    If k > 1 Then   ' проверка перехода
      If kv2 <> kv1 Then ' переход в другую четверть
        kv = kv2 - kv1
        If kv = 3 Then kv = -1
        If kv = -3 Then kv = 1
        If kv = 2 Or kv = -2 Then ' переход через две четверти
          If X1 = X2 Then kz = -100: Exit For
          yb = (Y2 * X1 - Y1 * X2) / (X1 - X2)
          If yb = 0 Then kz = -100: Exit For
          kv = kv * Sgn(yb)
          If kv1 = 2 Or kv1 = 4 Then kv = -kv
        End If
        kz = kz + kv
      End If
    End If
    X1 = X2: Y1 = Y2: kv1 = kv2
  Next
End Sub
