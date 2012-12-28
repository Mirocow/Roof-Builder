Attribute VB_Name = "Print_ALL"
'Option Explicit

'Public PicScaleWidth As Single
'Public PicScaleHeight As Single
'Public PicScaleLeft As Single
'Public PicScaleTop As Single

Type pl
    name As String
    Factory_Name As String
    len As Single
    prof As String
End Type
Public For_print_lapesn() As pl

Private SfigureAll As Double
Dim lsplit As String * 150
Public pitmanprintunload As Boolean
Public is_print As Boolean

Sub Print_Scat(ByVal N_Slope As Integer, print_ As Object)                ' печать ската

    Dim m1 As String
    Dim pcs As String
    Dim m2 As String
    Dim mm As String
    
    m1 = Setup.Text15 ' m1
    m2 = Setup.Text16 ' m2
    mm = Setup.Combo4.List(Setup.Combo4.ListIndex) ' mm
    pcs = Setup.Text18 ' pcs
    
    Dim str As String
    Dim str1 As String
    
    Dim space_ As String
    
    Dim l014E As String
    Dim nc As Integer
    
'    Dim l015A As String
    Dim ColsShow As Single ' > 3
    Dim RatioW As Single
    Dim RatioW1 As Single
    Dim RatioW2 As Single
    Dim RatioW3 As Single
    
    Dim Summ As Single
    Dim n As Integer
    Dim Slope As Integer
    
    Dim Len_of_list As Single
    
    Dim SaveScaleWidth As Integer
    Dim SaveScaleHeight As Integer
    Dim SaveScaleLeft As Integer
    Dim SaveScaleTop As Integer
    
    ' Переменные памяти
    Dim tWG As Single ' общая ширина
    Dim tW As Single ' рабочая ширина
    Dim tS As Integer ' длина волы, шаг
    Dim tOv As Integer ' нахлест
    Dim tMinl As Integer ' минимальная длина
    Dim tMaxl As Integer ' максимальная длина
    Dim tH As Integer ' высота панели
    Dim mMAXLIST As Integer ' Колво листов
    Dim tMAXSLOPES As Integer
    
    Dim YMAx As Single
    Dim l0186 As Single
    Dim l0188 As Single
    Dim l018A As Single
    Dim l018C As Single
    Dim l018E As Single

    On Error GoTo ERR
    
    is_print = True

    Dim RS As Recordset
    Set RS = GetProfilData(SlP(N_Slope).ProfilName, GetFactoryID(SlP(N_Slope).Factory_Name))
    If RS.RecordCount Then
        tWG = RS.Fields(3)
        tW = RS.Fields(2)
        tS = RS.Fields(4)
        tOv = RS.Fields(5)
        tMinl = RS.Fields(6)
        tMAXSLOPES = RS.Fields(7)
        tH = RS.Fields(8)
        Set RS = Nothing
    End If
  
    If NumCopies <> 0 Then

        Top print_

        print_.PSet (0, 500), RGB(255, 255, 255)

        For nc = 1 To NumCopies Step 1

            If N_Slope > 26 Then
                Slope = N_Slope + 70
            Else
                Slope = N_Slope + 64
            End If

            print_.FontBold = True
            print_.Print Space(8) & Setup.Text10.Text & Space(4) & lng.GetResIDstring(1453) & Chr$(Slope) ' Firm Name
            print_.FontBold = False

            Data print_
  
            Bottom print_
  
            ' Предпросмотр
            If FlagPrinter <> 1 Then

                ' Автокоррекция
                SetCenter N_Slope, SlP(N_Slope).ScaleLeftS, SlP(N_Slope).ScaleWidthS, SlP(N_Slope).ScaleTopS, SlP(N_Slope).ScaleHeightS

                print_.ScaleWidth = SlP(N_Slope).ScaleWidthS * 1.3
                print_.ScaleHeight = SlP(N_Slope).ScaleHeightS * 2.217 '* 1.3
                print_.ScaleLeft = SlP(N_Slope).ScaleLeftS - 0.13 * print_.ScaleWidth
                print_.ScaleTop = SlP(N_Slope).ScaleTopS - 0.18 * print_.ScaleHeight

'                print_.ScaleWidth = SlP(N_Slope).ScaleWidthS
'                print_.ScaleHeight = SlP(N_Slope).ScaleHeightS
'                print_.ScaleLeft = SlP(N_Slope).ScaleLeftS
'                print_.ScaleTop = SlP(N_Slope).ScaleTopS
                
            Else
            ' Принтер
                
                If SlP(N_Slope).ScaleWidthS = 0 And _
                SlP(N_Slope).ScaleHeightS = 0 And _
                SlP(N_Slope).ScaleLeftS = 0 And _
                SlP(N_Slope).ScaleTopS = 0 Then
                
                    SlP(N_Slope).ScaleLeftS = -100
                    SlP(N_Slope).ScaleWidthS = 1100
                    SlP(N_Slope).ScaleTopS = 600
                    SlP(N_Slope).ScaleHeightS = -600
                
                End If
                
                ' Автокоррекция
                SetCenter N_Slope, SlP(N_Slope).ScaleLeftS, SlP(N_Slope).ScaleWidthS, SlP(N_Slope).ScaleTopS, SlP(N_Slope).ScaleHeightS
                
                print_.ScaleWidth = SlP(N_Slope).ScaleWidthS * 1.3
                print_.ScaleHeight = SlP(N_Slope).ScaleHeightS * 2.217 * 1.3
                print_.ScaleLeft = SlP(N_Slope).ScaleLeftS - 0.13 * print_.ScaleWidth
                print_.ScaleTop = SlP(N_Slope).ScaleTopS - 0.28 * print_.ScaleHeight

            End If
    
            RatioW = print_.ScaleWidth / 90
            RatioW1 = 2 * RatioW
            print_.DrawStyle = 0
            If FlagPrinter Then print_.DrawWidth = Setup.Text3.Text
            If SlP(N_Slope).Pn_StartLC = 0 Then Exit Sub
            RatioW = print_.ScaleWidth / 100
            RatioW1 = 2 * RatioW
            RatioW2 = 3 * RatioW
            YMAx = -99999

  
            For n = 1 To SlP(N_Slope).CountSheets Step 1
                If List_Properties_PY(N_Slope, n) > YMAx Then YMAx = List_Properties_PY(N_Slope, n)
            Next n
  
            l0186 = YMAx - print_.ScaleHeight / 15
            If Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC) < SlP(N_Slope).PX_StartLC Then
                l0188 = Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC)
                l018A = SlP(N_Slope).PX_StartLC
            Else
                l0188 = SlP(N_Slope).PX_StartLC
                l018A = Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC)
            End If

            If l018A - l0188 < 2 * RatioW2 Then
                l018C = 2 * RatioW2
                l018E = l018A + RatioW1
            Else
                l018C = RatioW1
                l018E = ((SlP(N_Slope).PX_StartLC + Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC)) / 2) - RatioW1
            End If
  
            ' Шрифт
            Dim size As Single
            size = print_.FontSize
            print_.FontSize = 8
          
            print_.DrawStyle = 2
            If FlagPrinter Then print_.DrawWidth = 1
            print_.Line (Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_StartLC))-(Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), l0186 - RatioW1)
            print_.Line (SlP(N_Slope).PX_StartLC, YMAx)-(SlP(N_Slope).PX_StartLC, l0186 - RatioW1)
            print_.DrawStyle = 0
            
            If FlagPrinter Then print_.DrawWidth = Setup.Text3.Text
            print_.Line (l0188 - RatioW1, l0186 - RatioW2)-(l018A + l018C, l0186 - RatioW2)
            print_.Line (SlP(N_Slope).PX_StartLC, l0186 - RatioW)-(SlP(N_Slope).PX_StartLC, l0186 - 2 * RatioW2)
            print_.Line (Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), l0186 - RatioW)-(Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), l0186 - 2 * RatioW2)
            print_.Line (Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC) - RatioW, l0186 - 2 * RatioW1)-(Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC) + RatioW, l0186 - RatioW1)
            print_.Line (SlP(N_Slope).PX_StartLC - RatioW, l0186 - 2 * RatioW1)-(SlP(N_Slope).PX_StartLC + RatioW, l0186 - RatioW1)
            print_.PSet (l018E, l0186 - 0.5 * RatioW), RGB(255, 255, 255)
            print_.Print Format$(SlP(N_Slope).PX_StartLC - Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), "####")
            print_.PSet (SlP(N_Slope).PX_StartLC + RatioW, l0186 - 4 * RatioW), RGB(255, 255, 255)
        
            print_.Print lng.GetResIDstring(1454)
        
            RatioW3 = 0
            
            If RatioW3 = 0 Then
                If print_.ScaleWidth > 4000 Then RatioW3 = 1.5 * RatioW
            Else
                RatioW3 = 0
            End If
            
            For n = 1 To SlP(N_Slope).CountSheets Step 1
                If List_Properties_Length(N_Slope, n) > 0 Then
                    
                    ' вывод номера листов
                     If Setup.Check14.value Then
                         print_.PSet (List_Properties_PX(N_Slope, n) + 5, List_Properties_PY(N_Slope, n) - List_Properties_Length(N_Slope, n) - RatioW3 - 2), RGB(255, 255, 255)
                          Dim marker As String
                         marker = Format$(n, "00")
                         print_.Print marker
                     End If
    
                     ' Отображение шага волны
                     If Setup.Check2.value = 1 Then
                         If tS > 1 Then
                          Dim i As Integer
                          
                          If Setup.Option2.value Then
                              For i = List_Properties_PY(N_Slope, n) - List_Properties_Length(N_Slope, n) To List_Properties_PY(N_Slope, n) Step tS
                                  print_.Line (List_Properties_PX(N_Slope, n), i)-(List_Properties_PX(N_Slope, n) + tW, i), &H8000000F  ' Прорисовка шага волны
                              Next i
                          ElseIf Setup.Option1.value Then
                              For i = List_Properties_PY(N_Slope, n) To List_Properties_PY(N_Slope, n) - List_Properties_Length(N_Slope, n) Step -tS
                                  print_.Line (List_Properties_PX(N_Slope, n), i)-(List_Properties_PX(N_Slope, n) + tW, i), &H8000000F ' Прорисовка шага волны
                              Next i
                          End If
                          
                         End If
                     End If
                     
                     ' Отображение длин листов
                     If Setup.Check8.value = 1 Then
                         Dim nStr As Integer
                         
                         str = Format$(ConvertData(List_Properties_Length(N_Slope, n)), "0.00")
                         
                         For nStr = 1 To Len(str)
                             print_.CurrentX = List_Properties_PX(N_Slope, n) + ((tW) / 2) - (print_.TextWidth("0") / 2)
                             print_.CurrentY = List_Properties_PY(N_Slope, n) - (List_Properties_Length(N_Slope, n) / 2) - _
                             ((Len(str) * print_.TextHeight("0")) / 2) + ((nStr - 1) * print_.TextHeight("0"))
                             print_.Print mID$(str, nStr, 1)
                         Next
                         
                     End If
                     
                End If
            Next n
          
            For n = 1 To SlP(N_Slope).CountSheets Step 1
                If List_Properties_Length(N_Slope, n) > 0 Then
                     
                     ' рисование листов
                     print_.Line (List_Properties_PX(N_Slope, n), List_Properties_PY(N_Slope, n))-(List_Properties_PX(N_Slope, n) + tW, List_Properties_PY(N_Slope, n) - List_Properties_Length(N_Slope, n)), , B
             
                End If
            Next n
        
            print_.FontSize = size
                  
            print_.DrawStyle = 2
            If FlagPrinter Then print_.DrawWidth = Setup.Text3.Text - 2
                  
            For n = 1 To SlP(N_Slope).CountOfLines Step 1 ' Рисование самой фигуры
            
                print_.Line (Lape_Points_X(N_Slope, Lape_Lines(N_Slope, n, 0)), Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, n, 0)))-(Lape_Points_X(N_Slope, Lape_Lines(N_Slope, n, 1)), Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, n, 1))), RGB(0, 0, 0)
                If Setup.Check13 Then
                    Dim Llen As Long
                    Llen = Format(Sqr((Lape_Points_X(N_Slope, Lape_Lines(N_Slope, n, 1)) - Lape_Points_X(N_Slope, Lape_Lines(N_Slope, n, 0))) ^ 2 + _
                    (Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, n, 1)) - Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, n, 0))) ^ 2), "###.0")
                    MarkLine print_, N_Slope, -RatioW, n, Llen, 8, False
                End If
        
            Next n
                  
            ' print_.ScaleHeight / 6
            ' print_.PSet (0, 10300), RGB(255, 255, 255) ' Установка положения вывода текстовых данных
            print_.PSet (print_.ScaleLeft, print_.ScaleHeight * 0.17), RGB(255, 255, 255) ' Установка положения вывода текстовых данных
            
            print_.Print lng.GetResIDstring(1125) & Trim(SlP(N_Slope).Factory_Name) & Space(5) & Trim(SlP(N_Slope).ProfilName) & _
            " (" & ConvertData(tW) & "; " & ConvertData(tS) & "; " & ConvertData(tOv) & ")" & vbNewLine
            
            Dim ListLen As Single
            
            For n = 1 To SlP(N_Slope).CountSheets Step 1 ' Набивка длин по листам
                
                    If List_Properties_Length(N_Slope, n) > 0 Then
                        
                        ListLen = List_Properties_Length(N_Slope, n)
                        
                        str = Format$(ConvertData(ListLen), "0.00")
                        
                        If N_Slope > 26 Then
                            space_ = space_ & Space(18) & AddSpaceRight$(5, Chr$(N_Slope + 70) & Format$(n, "0")) & AddSpaceLeft$(5, str) & " " & mm
                        Else
                            space_ = space_ & Space(18) & AddSpaceRight$(5, Chr$(N_Slope + 64) & Format$(n, "0")) & AddSpaceLeft$(5, str) & " " & mm
                        End If
                        
                        ColsShow = ColsShow + 1
                        Summ = Summ + ListLen
                        
                    End If
            
                    ' Разбивка вывода на колонки
                    If ColsShow = Setup.Text19 Or n = SlP(N_Slope).CountSheets Then
                        print_.CurrentX = print_.ScaleLeft
                        print_.Print space_
                        space_ = ""
                        ColsShow = 0
                    End If
        
            Next n
              
            print_.CurrentX = print_.ScaleLeft
            print_.Print Space(18) & lsplit
            
            Len_of_list = Summ / 100
            
            str = Space(18) & Format$(Len_of_list, "# ##0.00") & m1
            str = str & ", " & lng.GetResIDstring(1034) & Format$((tW * Len_of_list / 100), "# ##0.00") & m2
            
            If tWG <> 0 Then
                str = str & ", " & lng.GetResIDstring(1035) & Format$((tWG * Len_of_list / 100), "# ##0.00") & m2
            End If
                  
                  
            If SlP(N_Slope).Sf > 0 And SlP(N_Slope).ListLength > 0 Then
            
                Dim SFigure As Single

                SFigure = SlP(N_Slope).Sf / 10000
                str = str & ", " & lng.GetResIDstring(1062) & Format(SFigure, "# ##0.00") & " " & m2
                
                If Setup.Check15.value And SlP(N_Slope).Sf > 0 And SlP(N_Slope).ListLength > 0 Then
    
                    Dim prc As Integer
                    prc = 100 - (SFigure / SlP(N_Slope).Sw) * 100
                    str = str & ", " & lng.GetResIDstring(1060) & prc & " %"
                
                End If
                
            End If
                  
            print_.CurrentX = print_.ScaleLeft
            print_.Print str
                  
            ' Вывод описания ската
            If Replace(SlP(N_Slope).Describ, Chr(0), "") <> "" Then
            
                str = ""
                print_.Print
                print_.CurrentX = print_.ScaleLeft
                For n = 1 To Len(SlP(N_Slope).Describ) Step 1
                    str1 = mID$(SlP(N_Slope).Describ, n, 1)
                    str$ = str$ + str1
                    If str1 = Chr$(10) Then str$ = str$ + Space$(20)
                Next n
                
                print_.Print Space(10) & str$
                
            End If
                  
            If FlagPrinter Then
                print_.NewPage
                print_.EndDoc
            End If
        
        Next
        
    End If
    
    is_print = False

Exit Sub

Exit Sub
ERR:
'STRERR = STRERR & time & ". (modulprint.Print_Scat) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.49." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Sub


Sub Summary(print_ As Object, Optional v As Integer)             'Суммарная
    
    Dim m1 As String
    Dim pcs As String
    Dim m2 As String
    Dim mm As String
    
    m1 = Setup.Text15 ' m1
    m2 = Setup.Text16 ' m2
    mm = Setup.Combo4.List(Setup.Combo4.ListIndex) ' mm
    pcs = Setup.Text18 ' pcs
    
    Dim l014E As String
    Dim l019C As Single
    Dim l01A2 As String

    Dim ipcs As Single
    Dim Summ_m1 As Single
    Dim ProfilSumm As Single
    Dim l01C0 As String
    
    Dim space1 As String
    Dim space2 As String
    
    Dim l01C8 As String
    Dim Len_of_list As Single
 
    Dim i As Integer
    Dim nc As Integer
    Dim n As Integer
    Dim str As String

    Dim Sw As Single
    Dim SWg As Single
    Dim tSW As Single
    
    Dim test_len As Single
    Dim lapes_names As String
    
    Dim len_of_lapes_names As Integer
    
    len_of_lapes_names = Setup.Text20

    On Error GoTo ERR
    
    is_print = True

    'Settings print_
  
    If NumCopies <> 0 Then
    
'        print_.Cls

        For nc = 1 To NumCopies Step 1

          Top print_

          print_.PSet (0, 500), RGB(255, 255, 255)

          str = Setup.Text10.Text ' Firm Name

          print_.FontBold = True
          print_.Print Space(8) & str & Space(4) & lng.GetResIDstring(1455) '& Gl.KolvoScatov
          print_.FontBold = False

          Data print_

          str = " "
          
          Load Pitmani

          Pitmani.Combo1.ListIndex = 0

          n = Pitmani.List1.ListCount

          ' Если нет данных не выводить список
          If n = 0 Then
              MsgBox lng.GetResIDstring(1098)
              If FlagPrinter Then print_.KillDoc
              pitmanprintunload = False
              Exit Sub
          End If
            
          ' заполняется SfigureAll
          If frmView.pFunc > -1 Then
              prepair frmView.pFunc
          Else
              Pitmani.Show vbModal, OfficeStart ' КОРЕКТИРОВКА ДЛИН
          End If

          If pitmanprintunload = True And FlagPrinter And frmView.pFunc = -1 Then
              Unload Pitmani
              print_.KillDoc
              Exit Sub
          End If

          Unload Pitmani

          ipcs = 1
          Summ_m1 = 0
          ProfilSumm = 0
    
          Dim profcheck As String
          Dim Factory_Name As String
          Dim sumstr As String
          
          sumstr = ""
    
          ' Печать наименований профиля с линией
          profcheck = For_print_lapesn(1).prof
          Factory_Name = For_print_lapesn(1).Factory_Name
          
          print_.ForeColor = vbRed
          print_.Print Space(13) & lsplit & vbNewLine & Space(13) & Factory_Name & Space(5) & profcheck & vbNewLine
          print_.ForeColor = vbBlack
        
          ' Предворительное заполнение данных одного листа
          test_len = For_print_lapesn(1).len
          str = Format$(test_len, "0.00")
          space1 = Space$(Abs(3 - Len(pcs)))
          space2 = Space$(Abs(5 - Len(str)))
          str = Space$(15) & str & space2 & mm & " " & "1" & space1 & pcs
          
          If n > 1 Then
          
            ' Если листов более 1
            lapes_names = " (" & For_print_lapesn(1).name & ","
            
          Else
            
            ' Если скат содержит всего 1 лист
            lapes_names = " (" & For_print_lapesn(1).name & ")"
            print_.Print str & mID$(lapes_names, 1, len_of_lapes_names)
            Summ_m1 = For_print_lapesn(1).len

            Dim RS As Recordset
            Set RS = GetProfilData(profcheck, GetFactoryID(Factory_Name))
            If RS.RecordCount Then
                tWG = ConvertData(RS.Fields(3), True)
                tW = ConvertData(RS.Fields(2), True)
                Set RS = Nothing
            End If

            If tW <> 0 Then
                tSW = tW * For_print_lapesn(1).len / 1000 / 100
                Sw = tSW
            End If

            If tWG <> 0 Then
                tSW = tWG * For_print_lapesn(1).len / 1000 / 100
                SWg = tSW
            End If

          End If
                    
          ' Если больше 1 листа на скате 1
          For i = 2 To n Step 1
          
              If test_len = For_print_lapesn(i).len And profcheck = For_print_lapesn(i).prof Then
                  
                  lapes_names = lapes_names + For_print_lapesn(i).name & ","
                  ipcs = ipcs + 1
                
                  If Len(str) > 0 Then
                      str = Format$(test_len, "0.00")
                      space1 = Space$(Abs(3 - Len(pcs)))
                      space2 = Space$(Abs(5 - Len(str)))
                      str = Space$(15) & str & space2 & mm & "  " & CStr(ipcs) & space1 & pcs
                  End If
                
                  Summ_m1 = Summ_m1 + ConvertData(For_print_lapesn(i).len, True)
                  ProfilSumm = ProfilSumm + ConvertData(For_print_lapesn(i).len, True)
                
              Else
              
                  Summ_m1 = Summ_m1 + ConvertData(For_print_lapesn(i - 1).len, True)
                  ProfilSumm = ProfilSumm + ConvertData(For_print_lapesn(i - 1).len, True)
                  
                  ' Убирание запятой и закрытие списка обозначений )
                  lapes_names = Left$(lapes_names, Len(lapes_names) - 1)
                  lapes_names = lapes_names & ")"
                  
                  l019C = 1
                  While l019C < Len(lapes_names)
                      print_.Print str & mID$(lapes_names, l019C, len_of_lapes_names)
                      str = Space$(47)
                      l019C = l019C + len_of_lapes_names
                  Wend
                
                  lapes_names = " (" & For_print_lapesn(i).name & ","
                  test_len = For_print_lapesn(i).len
                  ipcs = 1
                  space1 = Space$(Abs(3 - Len(pcs)))
                  space2 = Space$(Abs(5 - Len(test_len)))
                  str = Space$(15) & test_len & space2 & mm & " " & CStr(ipcs) & space1 & pcs
                
              End If
              
              If i = n Then ' Заканчиваем вывод
                
                  Summ_m1 = Summ_m1 + ConvertData(For_print_lapesn(n).len, True)
                  ProfilSumm = ProfilSumm + ConvertData(For_print_lapesn(n).len, True)
                  
                  ' Убирание запятой и закрытие списка обозначений )
                  If lapes_names <> "" Then lapes_names = Left$(lapes_names, Len(lapes_names) - 1)
                  lapes_names = lapes_names & ")"
                
                  ' Вывод строки обозначений
                  l019C = 1
                  While l019C < Len(lapes_names)
                      print_.Print str & mID$(lapes_names, l019C, len_of_lapes_names)
                      str = Space$(47)
                      l019C = l019C + len_of_lapes_names
                  Wend
        
              End If
              
              ' Обработка смены профиля
              If i = n Or profcheck <> For_print_lapesn(i).prof Then
                
                  print_.Print
        
                  Set RS = GetProfilData(profcheck, GetFactoryID(Factory_Name))
                  If Not RS Is Nothing Then
                  If RS.RecordCount Then
                      tWG = ConvertData(RS.Fields(3))
                      tW = ConvertData(RS.Fields(2))
                      Set RS = Nothing
                  End If
                  End If

                  Len_of_list = ProfilSumm
                  ProfilSumm = 0
        
                  If Gl.FileNameExtension = ".rbp" Then
                    sumstr = Format$(ConvertData(Len_of_list), "# ##0.00") + m1$
                  End If
                  
                  If tW <> 0 Then
                      tSW = ConvertData(tW, True) * Len_of_list
                      Sw = Sw + tSW
                      If Gl.FileNameExtension = ".rbp" Then
                        sumstr = sumstr & ", " & lng.GetResIDstring(1034) & Format$(tSW / 10000, "# ##0.00") + m2$
                      End If
                  End If
                  
                  If tWG <> 0 Then
                      tSW = ConvertData(tWG, True) * Len_of_list
                      SWg = SWg + tSW
                      If Gl.FileNameExtension = ".rbp" Then
                        sumstr = sumstr & ", " & lng.GetResIDstring(1035) & Format$(tSW / 10000, "# ##0.00") + m2$
                      End If
                  End If
                  
                  If Gl.FileNameExtension = ".rbp" Then print_.Print Space(13) & sumstr
        
                  ' Печать наименований профиля следующего
                  If i <> n Then
                      print_.ForeColor = vbRed
                      print_.Print Space(13) & lsplit & vbNewLine & Space(13) & Factory_Name & Space(5) & profcheck & vbNewLine
                      print_.ForeColor = vbBlack
                      profcheck = For_print_lapesn(i).prof
                      Factory_Name = For_print_lapesn(i).Factory_Name
                  End If
                  
                  ' Задание следующего профиля
                  profcheck = For_print_lapesn(i).prof
                  Factory_Name = For_print_lapesn(i).Factory_Name
                  
              End If
        
          Next i
                
          Len_of_list = Summ_m1 / 100
              
          print_.ForeColor = vbRed
          print_.Print Space(10) & "_" & lsplit
          print_.ForeColor = vbBlack
          
          sumstr = lng.GetResIDstring(1141) & Format$(Len_of_list, "# ##0.00") + m1$
              
          If tW > 0 And Sw > 0 Then
            sumstr = sumstr & "; " & lng.GetResIDstring(1034) & Format$(Sw / 10000, "# ##0.00") + m2$
          End If
              
          If tWG > 0 And SWg > 0 Then
            sumstr = sumstr & "; " & lng.GetResIDstring(1035) & Format$(SWg / 10000, "# ##0.00") + m2$
          End If
              
          If SfigureAll > 0 And Sw > 0 Then
                
                SfigureAll = SfigureAll / 10000
                
                ' Площадь ската
                sumstr = sumstr & vbNewLine & Space(10) & lng.GetResIDstring(1062) & Format(SfigureAll, "# ##0.00") & " " & m2$
            
                Dim prc As Integer
                prc = 100 - ((SfigureAll / (Sw / 10000)) * 100)
                sumstr = sumstr & "; " & lng.GetResIDstring(1060) & Format(prc, "00") & " %"
        
          End If
            
          print_.Print Space(10) & sumstr
            
        '      If Setup.Text8 <> 0 Then
        '        print_.Print Space$(13) & "Округление до " & Setup.Text8 & " после запятой."
        '      End If
            
         Bottom print_
            
        If FlagPrinter Then
          print_.NewPage
          print_.EndDoc
        End If
      
     Next nc
    
    End If
    
    is_print = False

Exit Sub

ERR:
'STRERR = STRERR & time & ". (modulprint.Summary) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.50." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Sub


Sub Data(print_ As Object)
    Dim str As String
    Dim l014E As String
    Dim l0150 As Single
    Dim l0154 As Single
    Dim l01A2 As String
    Dim l01A6 As String
    Dim l019C As Integer
    Dim i As Integer

    On Error GoTo ERR

    print_.Print

    Dim Description As String
    Description$ = vbNewLine & Setup.Text11.Text & vbNewLine ' Firm R
    l0150 = InStr(Description$, vbNewLine)
    While l0150 > 0
        l0154 = l0150 + 2
        l0150 = InStr(l0154, Description$, vbNewLine, vbBinaryCompare)
        If l0150 > 0 Then
            str$ = mID$(Description$, l0154, l0150 - l0154)
            Dim strt As String
            l019C = 1
            While l019C <= Len(str$)
                print_.Print Space(11) & strt$ & mID$(str$, l019C, 120)
                l019C = l019C + 120
            Wend
    End If
    Wend
    
    print_.Print
    
    Description$ = vbNewLine & Setup.Text17.Text & vbNewLine ' Firm C
    l0150 = InStr(Description$, vbNewLine)
    While l0150 > 0
        l0154 = l0150 + 2
        l0150 = InStr(l0154, Description$, vbNewLine, vbBinaryCompare)
        If l0150 > 0 Then
            str$ = mID$(Description$, l0154, l0150 - l0154)
'            Dim strt As String
            l019C = 1
            While l019C <= Len(str$)
                print_.Print Space(11) & strt$ & mID$(str$, l019C, 120)
                l019C = l019C + 120
            Wend
    End If
    Wend

    print_.Print
    
    str$ = lng.GetResIDstring(1073)
    
    For i = 1 To Len(UserCreatProject) Step 1
        l01A6$ = mID$(UserCreatProject, i, 1)
        str$ = str$ + l01A6$
        If l01A6$ = Chr$(10) Then str$ = str$ + Space$(20)
    Next i
    
    print_.Print str$
    
    print_.Print
    
    str$ = lng.GetResIDstring(1074)
    
    print_.Print str$
    
    str$ = Space$(10)
    For i = 1 To Len(Project.Text3) Step 1
        l01A6$ = mID$(Project.Text3, i, 1)
        str$ = str$ + l01A6$
        If l01A6$ = Chr$(10) Then str$ = str$ + Space$(10)
    Next i
    
    print_.Print str$
      
    str$ = lng.GetResIDstring(1075)
    print_.Print
    print_.Print str$ & " " & width1 & " " & "-" & " " & cover & " " & "-" & " " & ColorRoof  'Project.Label8.text
      
Exit Sub

ERR:
'STRERR = STRERR & time & ". (modulprint.data) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.51." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Sub


Sub MainPicture(print_ As Object) 'Общий вид
    Dim str As String
    Dim i As Integer
    Dim n As Integer
    Dim strt As String
    Dim l0124 As Single
    Dim l0126 As Single

    On Error GoTo ERR
    
    is_print = True

    'Settings print_
  
    If NumCopies <> 0 Then

        For i = 1 To NumCopies Step 1

            Top print_
            
            SetCenter N_Slope, SlP(N_Slope).ScaleLeftS, SlP(N_Slope).ScaleWidthS, SlP(N_Slope).ScaleTopS, SlP(N_Slope).ScaleHeightS

            print_.PSet (0, 500), RGB(255, 255, 255)


            str$ = Setup.Text10.Text ' Firm Name
            print_.FontBold = True
            print_.Print Space(8) & str$ & Space(4) & lng.GetResIDstring(1456) '"Общий вид  "
            print_.FontBold = False

            Data print_

            Bottom print_

            str$ = "                    "
            For n = 1 To Len(MainDescrib) Step 1
                strt$ = mID$(MainDescrib, n, 1)
                str$ = str$ + strt$
                If strt$ = Chr$(10) Then str$ = str$ + Space$(20)
            Next n
            print_.Print str$

            If FlagPrinter <> 1 Then
                print_.ScaleWidth = ScaleWidth_Main * 1.8
                print_.ScaleHeight = ScaleHeight_Main * 4.7
                print_.ScaleLeft = ScaleLeft_Main - 0.001 * print_.ScaleWidth
                print_.ScaleTop = ScaleTop_Main - 0.25 * print_.ScaleHeight
            Else
                print_.ScaleWidth = ScaleWidth_Main * 1.1
                print_.ScaleHeight = ScaleHeight_Main * 2.1
                print_.ScaleLeft = ScaleLeft_Main - 0.001 * print_.ScaleWidth
                print_.ScaleTop = ScaleTop_Main - 0.25 * print_.ScaleHeight
            End If

            print_.DrawStyle = 0
            
            If FlagPrinter Then print_.DrawWidth = Setup.Text3.Text
            For n = 1 To MainCountOfLines Step 1
                print_.Line (Main_Points_X(Points_m_A(n)), Main_Points_Y(Points_m_A(n)))-(Main_Points_X(Points_m_B(n)), Main_Points_Y(Points_m_B(n))), RGB(0, 0, 0)
            Next n

            If FlagPrinter Then print_.DrawWidth = 1

            l0124 = print_.ScaleWidth / 90
            l0126 = 2 * l0124
            For n = 1 To KolvoScatov Step 1
                l0124 = print_.ScaleWidth / 80
                l0126 = 0.9 * l0124
                
'                print_.Line (Label_X(n) - l0126, Label_Y(n) + l0126)-(Label_X(n) + l0126, Label_Y(n) - l0126), , B
                
                print_.FontBold = True
                print_.PSet (Label_X(n) - l0124 * 0.5, Label_Y(n) + l0124 * 0.8), RGB(255, 255, 255)
                If n > 26 Then
                    print_.Print Chr$(n + 70)
                Else
                    print_.Print Chr$(n + 64)
                End If
                print_.FontBold = False
  
            Next n

            If FlagPrinter Then
                print_.NewPage
                print_.EndDoc
            End If

        Next

    End If
    
    is_print = False

    Exit Sub

ERR:
'STRERR = STRERR & time & ". (modulprint.MainPicture) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.52." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Sub


Function prepair(P As Integer) As Integer
Dim str As String
Dim n As Integer
Dim var As Single
Dim N_list As Integer
Dim c As Integer
Dim N_Slope As Integer

  On Error GoTo ERR

  ReDim For_print_lapesn(0)

  n = 0
  SfigureAll = 0

  ' Подготовка массива к сортировке
  For N_Slope = 1 To KolvoScatov Step 1
      For N_list = 1 To SlP(N_Slope).CountSheets Step 1
              
              If List_Properties_Length(N_Slope, N_list) > 0 Then
                  n = n + 1
    
                  ReDim Preserve For_print_lapesn(n)
                  
                  For_print_lapesn(n).len = Format$(ConvertData(List_Properties_Length(N_Slope, N_list)), "0.00")
  
                  If N_Slope > 26 Then
                      For_print_lapesn(n).name = Chr$(70 + N_Slope) & Format$(N_list, "00")
                  Else
                      For_print_lapesn(n).name = Chr$(64 + N_Slope) & Format$(N_list, "00")
                  End If
  
                  For_print_lapesn(n).prof = Trim(SlP(N_Slope).ProfilName)
                  For_print_lapesn(n).Factory_Name = Trim(SlP(N_Slope).Factory_Name)
  
              End If

      Next N_list

      If SlP(N_Slope).CountSheets > 0 Then SfigureAll = SfigureAll + SlP(N_Slope).Sf

  Next N_Slope
  
  '
  ' СОРТИРОВКА И ГРУПИРОВКА
  '

  If Gl.FileNameExtension = ".rbp" And P <= 1 Then
      ' Сортировка по наименованию профиля
      For N_list = 1 To n - 1 Step 1
          For c = N_list + 1 To n Step 1
  
              If For_print_lapesn(c).prof > For_print_lapesn(N_list).prof Then
      
                  ' Поменять местами длины
                  var = For_print_lapesn(N_list).len
                  For_print_lapesn(N_list).len = For_print_lapesn(c).len
                  For_print_lapesn(c).len = var
      
                  ' Поменять местами обозначения
                  str = For_print_lapesn(N_list).name
                  For_print_lapesn(N_list).name = For_print_lapesn(c).name
                  For_print_lapesn(c).name = str
      
                  ' Поменять местами имя профиля
                  str = For_print_lapesn(N_list).prof
                  For_print_lapesn(N_list).prof = For_print_lapesn(c).prof
                  For_print_lapesn(c).prof = str
                  
                  ' Поменять местами имя производителя
                  str = For_print_lapesn(N_list).Factory_Name
                  For_print_lapesn(N_list).Factory_Name = For_print_lapesn(c).Factory_Name
                  For_print_lapesn(c).Factory_Name = str
      
              End If
    
          Next c
          
      Next N_list
  End If

  ' МЕТОДЫ СОРТИРОВКИ
  If P = 0 Then
  
      ' Сортировка по убыванию
      For N_list = 1 To n Step 1
          For c = N_list + 1 To n Step 1

              If For_print_lapesn(c).prof = For_print_lapesn(N_list).prof And For_print_lapesn(c).len > For_print_lapesn(N_list).len Then
      
                  ' Поменять местами длины
                  var = For_print_lapesn(N_list).len
                  For_print_lapesn(N_list).len = For_print_lapesn(c).len
                  For_print_lapesn(c).len = var

                  ' Поменять местами обозначения
                  str = For_print_lapesn(N_list).name
                  For_print_lapesn(N_list).name = For_print_lapesn(c).name
                  For_print_lapesn(c).name = str
      
                  str = For_print_lapesn(N_list).prof
                  For_print_lapesn(N_list).prof = For_print_lapesn(c).prof
                  For_print_lapesn(c).prof = str
                  
                  ' Поменять местами имя производителя
                  str = For_print_lapesn(N_list).Factory_Name
                  For_print_lapesn(N_list).Factory_Name = For_print_lapesn(c).Factory_Name
                  For_print_lapesn(c).Factory_Name = str
      
              End If

          Next c
          
      Next N_list

  ElseIf P = 1 Then
  
      ' Сортировка по возростанию
      For N_list = 1 To n Step 1
          For c = N_list + 1 To n Step 1

              If For_print_lapesn(c).prof = For_print_lapesn(N_list).prof And For_print_lapesn(c).len < For_print_lapesn(N_list).len Then
      
                  ' Поменять местами длины
                  var = For_print_lapesn(N_list).len
                  For_print_lapesn(N_list).len = For_print_lapesn(c).len
                  For_print_lapesn(c).len = var

                  ' Поменять местами обозначения
                  str = For_print_lapesn(N_list).name
                  For_print_lapesn(N_list).name = For_print_lapesn(c).name
                  For_print_lapesn(c).name = str
      
                  str = For_print_lapesn(N_list).prof
                  For_print_lapesn(N_list).prof = For_print_lapesn(c).prof
                  For_print_lapesn(c).prof = str
                  
                  ' Поменять местами имя производителя
                  str = For_print_lapesn(N_list).Factory_Name
                  For_print_lapesn(N_list).Factory_Name = For_print_lapesn(c).Factory_Name
                  For_print_lapesn(c).Factory_Name = str
      
              End If

          Next c
          
      Next N_list

  End If

  ' Сортировка по наименованию скатов
  N_list = 1
  While N_list < n
      c = N_list + 1
      While For_print_lapesn(N_list).len = For_print_lapesn(c).len And For_print_lapesn(N_list).prof = For_print_lapesn(c).prof And c < n
          If For_print_lapesn(c).name < For_print_lapesn(N_list).name Then

              str = For_print_lapesn(N_list).name
              For_print_lapesn(N_list).name = For_print_lapesn(c).name
              For_print_lapesn(c).name = str

          End If
          c = c + 1
      Wend
      N_list = N_list + 1
  Wend

  prepair = n
  Exit Function

ERR:
'STRERR = STRERR & time & ". (modulprint.prepair) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.53." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Function


Private Sub SetCenter(N_Slope As Integer, ByRef ScaleLeft As Single, ByRef ScaleWidth As Single, ByRef ScaleTop As Single, ByRef ScaleHeight As Single)
    Dim XMax As Single
    Dim XMin As Single
    Dim YMAx As Single
    Dim YMin As Single

    Dim FigureLenght As Single
    Dim FigureHeight As Single

    Dim l00A6 As Single
    Dim l00AE As Single

    Dim l004A As Single
    Dim l004C As Single
    
    On Error Resume Next

      If SlP(N_Slope).CountOfLines < 1 Then
          GoTo L96EE
      End If

      Dim P As Integer
      XMax = -99999: XMin = 99999: YMAx = -99999: YMin = 99999
      For P = 1 To SlP(N_Slope).CountOfPoints Step 1
          If Lape_Points_X(N_Slope, P) < XMin Then XMin = Lape_Points_X(N_Slope, P)
          If Lape_Points_X(N_Slope, P) > XMax Then XMax = Lape_Points_X(N_Slope, P)
          If Lape_Points_Y(N_Slope, P) < YMin Then YMin = Lape_Points_Y(N_Slope, P)
          If Lape_Points_Y(N_Slope, P) > YMAx Then YMAx = Lape_Points_Y(N_Slope, P)
      Next P

      FigureLenght = XMax - XMin
      FigureHeight = YMAx - YMin

      If FigureHeight <= 0 Then FigureHeight = 1
      If FigureLenght <= 0 Then FigureLenght = 1
      If FigureHeight = 0 Then FigureHeight = 1
      
      Dim RatioW As Single
      Dim RatioH As Single
      
      RatioW = 1.51
      RatioH = 1.2
      
      If RatioW * FigureLenght >= 6400 Then GoTo L96EE
      
      If FigureLenght / FigureHeight > RatioH Then

          l00A6 = RatioH * FigureLenght
          l00AE = l00A6 / RatioW

          If l00A6 > 6400 Then
              GoTo L96EE
          End If

          ScaleLeft = XMin - 0.1 * FigureLenght
          ScaleWidth = l00A6
          ScaleHeight = -l00AE
          ScaleTop = (YMAx - (FigureHeight / 2)) + (l00AE / 2)

      Else

          l00AE = -RatioH * FigureHeight
          l00A6 = l00AE * -RatioW

          If l00A6 > 6400 Then
              GoTo L96EE
          End If

          ScaleLeft = (XMax - (FigureLenght / 2)) - (l00A6 / 2)
          ScaleWidth = l00A6
          ScaleHeight = l00AE
          ScaleTop = YMAx + 0.1 * FigureHeight

      End If

L96EE:
End Sub

Sub Top(print_ As Object)
    Dim Style As Integer
    On Error Resume Next
    Style = print_.DrawStyle
    print_.ScaleMode = 1
    print_.DrawStyle = 0
    print_.FontSize = 6 'Gl.PrintFontSize - 2
    print_.PSet (0, 150), RGB(255, 255, 255)
    print_.Print Space(2) & "`" & App.ProductName & "` " & Gl.Ver & Space(2) & ", E-mail: support@roof-builder.ru, " & lng.GetResIDstring(9652) & " Computer: " & Gl.UserName & Space(10) & CurrentFile & Space(10) & Date$
    print_.FontSize = Gl.PrintFontSize
    print_.Line (10, 370)-(print_.Width - 10, 370)
    print_.DrawStyle = Style
    lsplit = String$(100, "_")
End Sub


Sub Bottom(print_ As Object)
    On Error GoTo ERR
    If Setup.GetIDData(8) Then Exit Sub
ERR:
    Dim fs As Integer
    fs = print_.FontSize
    print_.FontSize = 14
    print_.FontBold = True
    print_.PSet (print_.ScaleWidth / 2.5, print_.ScaleHeight / 3), RGB(255, 255, 255)
    print_.Print "Roof Builder"
    print_.PSet (print_.ScaleWidth / 2.5 + 100, print_.CurrentY + 100), RGB(255, 255, 255)
    print_.Print "DEMO VERSION"
    print_.PSet (print_.ScaleWidth / 2.5, print_.CurrentY + 100), RGB(255, 255, 255)
    print_.Print "www.roof-builder.ru (RUS)"
    print_.PSet (print_.ScaleWidth / 2.5, print_.CurrentY + 100), RGB(255, 255, 255)
    print_.Print "www.roofbuilder.net (ENG)"
    print_.FontBold = False
    print_.FontSize = fs
End Sub
