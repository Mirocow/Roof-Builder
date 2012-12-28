Attribute VB_Name = "Module10"
' Module10
Option Explicit
'Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Private Type savet
'    name As String * 50
'    trash As String * 255
'    datereg As Date
'    value_reg As String * 1024
'    value_lic_n As String * 1024
'    Ver As String * 6
'End Type

Function AddSpaceLeft(LenField As Integer, ByVal p002C As String) As String
    Dim l002E As String
        l002E$ = p002C
        While Len(l002E$) < LenField
            l002E$ = " " + l002E$
        Wend

        AddSpaceLeft$ = l002E$
End Function


Function AddSpaceRight(LenField As Integer, ByVal p0112 As String) As String
    Dim l0114 As String
        l0114$ = p0112
        While Len(l0114$) < LenField
            l0114$ = l0114$ + " "
        Wend

        AddSpaceRight$ = l0114$
End Function


Sub Draw_lape_label(P As PictureBox, n As Integer, X As Single, Y As Single, Optional color As OLE_COLOR = vbBlack)                                                            ' KV DRAW
    Dim l010A As Single
    Static az As Integer
    Dim tmp_color As OLE_COLOR

        If Not P.name = "Picture2" Then
            
            l010A = P.ScaleWidth / 100
            P.PSet (X - l010A * 0.4, Y + l010A * 0.8), "&H00DCFBFC"
            
            P.FontBold = True
            
            tmp_color = P.ForeColor
            P.ForeColor = color
            If n > 26 Then
                P.Print Chr$(n + 70)
            Else
                P.Print Chr$(n + 64)
            End If
            P.ForeColor = tmp_color
            
            P.FontBold = False

            'If Not P.name = "Picture3" Then
            '    P.Line (x - l010A, Y + l010A)-(x + l010A, Y - l010A), , B
            'End If

        Else
            
            l010A = P.ScaleWidth / 49
            P.PSet (X - l010A, Y + l010A), "&H00DCFBFC"
            
            P.FontBold = True
            
            tmp_color = P.ForeColor
            P.ForeColor = color
            If n > 26 Then
                P.Print Chr$(n + 70)
            Else
                P.Print Chr$(n + 64)
            End If
            P.ForeColor = tmp_color
            
            P.FontBold = False

        End If

End Sub


Sub Clear_lape_label(LapeName As Integer) 'lape CLEAR
    Dim l007C As Single
    Dim l0084 As Single
    Dim l0088 As Single
        l007C = Label_X(LapeName)
        l0084 = Label_Y(LapeName)
        
        l0088 = ROOFPIC.Picture1.ScaleWidth / 100
        ROOFPIC.Picture1.ForeColor = "&H00DCFBFC"
        
        ROOFPIC.Picture1.PSet (l007C - l0088 * 0.4, l0084 + l0088 * 0.8), "&H00DCFBFC"
        
        ROOFPIC.Picture1.FontBold = True
        
        If LapeName > 26 Then
            ROOFPIC.Picture1.Print Chr$(LapeName + 70)
        Else
            ROOFPIC.Picture1.Print Chr$(LapeName + 64)
        End If
        
        ROOFPIC.Picture1.FontBold = False

        'ROOFPIC.Picture1.Line (l007C - l0088, l0084 + l0088)-(l007C + l0088, l0084 - l0088), "&H00DCFBFC", B
        'ROOFPIC.Picture1.ForeColor = RGB(0, 0, 0)
End Sub


Function Close_Project(Optional CloseOnlyProject As Boolean, Optional exit_dialog_show As Boolean, Optional save_dialog_show As Boolean) As Boolean
'запрос на закрытие проекта
If CurrentFile <> "" Then
On Error GoTo ERR

    If isSave = True Then
    
        If IsLic Then
        
            Select Case MsgBox(lng.GetResIDstring(1420, "%DIR%", ProjectsDir, "%FILENAME%", LCase$(CurrentFile)), vbInformation + vbYesNoCancel)
            Case vbYes
                fpt_ ProjectsDir, CurrentFile
                Close_Project = True
            Case vbNo
                Close_Project = True
            Case vbCancel
                Close_Project = False
                Exit Function
            End Select
            
        Else
        
            Select Case MsgBox(lng.GetResIDstring(1419, "%PROJECTNAME%", LCase$(CurrentFile)), vbInformation + vbYesNo)
            Case vbYes
                Close_Project = True
            Case vbNo
                Close_Project = False
                Exit Function
            End Select
            
        End If

    Else
    
        Close_Project = True
        CurrentProjectDir = ""
        
    End If
    
ERR:

'    If IsLic And FileExists(Gl.TempDir & "~$" & CurrentFile) Then
'        On Error Resume Next
'        Kill Gl.TempDir & "~$" & CurrentFile
'    End If
    
    ' Закрытие проекта
    If CloseOnlyProject Then
    
        CurrentFile = ""
        
        Dim F As Form
        For Each F In Forms
            Unload F
            Set F = Nothing
        Next F
        
        OfficeStart.StatusBar.Panels(2).Text = ""
        isSave = False
        Profil_Name = ""
        CurrentProjectDir = ""

    End If
    
Else

    Close_Project = True
    
End If

'OfficeStart.Clear_project True

End Function

Sub withoutl()
    About.Show vbModal, OfficeStart
End Sub


'Function CircleSquare(Slope)
'Dim I As Integer
'Const pi = 3.141593
'Dim Square As Single
'Dim X2 As Single
'Dim Y2 As Single
'Dim v2 As Single
'Dim ay2 As Single
'Dim ax2 As Single
'Dim alfa2 As Single
'Dim alfa1 As Single
'Dim v1
'Dim x1
'Dim y1
'  ' Вычисление площади многоугольника
'  '————————————————————————————————
'  ' ВХОД:
'  ' xyd() - массив координат углов многоугольника
'  ' x = xyd(1,i), y = xyd(2,i) ; i = 1 to Np
'  '  (Np-1) - количество узлов
'  '  координаты 1-й точки = координатам N-й
'  '
'  ' ВЫХОД: Square - площадь многоугольника
'  '''''''''''''''''''''''''''''''''''''''''''''''
'
''Lape_Points_XY(Slope, ApConnect(Slope, i)).X
''Lape_Points_XY(Slope, ApConnect(Slope, i)).Y
''
''Lape_Points_XY(Slope, BpConnect(Slope, i)).X
''Lape_Points_XY(Slope, BpConnect(Slope, i)).Y
'
'  Square = 0
'  For I = 1 To SlP.CountOfLines(Slope)  ' Np + 1
'
'    X2 = Lape_Points_XY(Slope, I).X
'    Y2 = Lape_Points_XY(Slope, I).Y
'
'    v2 = Sqr(X2 ^ 2 + Y2 ^ 2)
'    ay2 = Abs(Y2): ax2 = Abs(X2)
'    If ax2 * 10000 > ay2 Then
'      alfa2 = Atn(ay2 / ax2)
'    Else: alfa2 = pi * 0.5
'    End If
'    If X2 < 0 Then alfa2 = pi - alfa2
'    If Y2 < 0 Then alfa2 = -alfa2
'    If I > 1 Then   ' проверка перехода
'      Square = Square + 0.5 * sin(alfa2 - alfa1) * v1 * v2
'    End If
'    x1 = X2: y1 = Y2: v1 = v2: alfa1 = alfa2
'  Next
'  CircleSquare = Square / 1000
'End Function


'Public Function myRound(Number, Digits As Long)
'Dim decStr As String
'decStr = Replace(Space(Digits), " ", "0")
'myRound = CDec(FChange_scrolormat(Number, "0." & decStr))
'End Function


Public Sub Change_scrol(Pic As Object, scroll As Object)
On Error Resume Next

Dim FigureLength As Single
Dim l00C4 As Single
Dim l00C6 As Single
Dim l00C8 As Single
Dim l00CA As Single

  FigureLength = Pic.ScaleLeft + Pic.ScaleWidth
  l00C4 = (Pic.ScaleLeft + FigureLength) / 2
  l00C6 = Pic.ScaleTop + (Pic.ScaleHeight / 2)
  l00C8 = scroll.value
  l00CA = l00C8 / RatioW '1.33 '(l00C8 / (Abs(Pic.ScaleLeft) + Pic.ScaleWidth)) 'RatioW
  
  Pic.ScaleTop = l00C6 + (l00CA / 2) '+ 100
  Pic.ScaleHeight = -l00CA '+ 100
  Pic.ScaleWidth = l00C8 '+ 100
  Pic.ScaleLeft = l00C4 - (l00C8 / 2) '+ 100

End Sub


Sub right_left_start_point(i As Integer)
    Dim P As Integer
    Dim XMin As Single: Dim PXmin As Single
    Dim XMax As Integer: Dim PXMax As Single
        XMin = 9999
        XMax = -9999
    
        On Error GoTo ERR
    
        For P = 1 To SlP(i).CountOfPoints Step 1
            If XMin >= Lape_Points_X(i, P) Then XMin = Lape_Points_X(i, P): PXmin = P
            If XMax <= Lape_Points_X(i, P) Then XMax = Lape_Points_X(i, P): PXMax = P
        Next P
      
        If Setup.Option3.value Then
            SlP(i).Pn_StartLC = PXmin  '- самая левая точка
            Plgs(LNC).Dll.HorizontalDirection = 1
        Else
            SlP(i).Pn_StartLC = PXMax  '- самая правая точка
            Plgs(LNC).Dll.HorizontalDirection = 0
        End If
    
        SlP(i).PX_StartLC = Int(Lape_Points_X(i, SlP(i).Pn_StartLC))
        SlP(i).PX_StartLC = IIf(SlP(i).PX_StartLC = 0, 1, SlP(i).PX_StartLC)
    
        Exit Sub
ERR:
End Sub



Sub vert_hor_point(i As Integer)
    Dim YMin As Integer: Dim PYmin As Integer
    Dim YMAx As Integer: Dim PYMax As Integer
    Dim P As Integer
    
        On Error GoTo ERR
    
        YMin = 9999
        YMAx = -9999
    
        For P = 1 To SlP(i).CountOfPoints Step 1
            If YMin >= Lape_Points_Y(i, P) Then YMin = Lape_Points_Y(i, P): PYmin = P
            If YMAx <= Lape_Points_Y(i, P) Then YMAx = Lape_Points_Y(i, P): PYMax = P
        Next P
    
        If Setup.Option2.value Then
            SlP(i).Pn_Red_lines = PYmin '- самая нижняя точка
            Plgs(LNC).Dll.VerticalDirection = 1
        Else
            SlP(i).Pn_Red_lines = PYMax '- самая верхняя точка
            Plgs(LNC).Dll.VerticalDirection = 0
        End If

        Exit Sub
ERR:
End Sub

Public Sub MarkLine(ByRef Obj, N_Slope, coef, NLINE, Msg, FontSize As Integer, b As Boolean)
    
    If Lape_Lines(N_Slope, NLINE, 1) > 0 Then
    
        Obj.PSet ( _
(Lape_Points_X(N_Slope, Lape_Lines(N_Slope, NLINE, 1)) + Lape_Points_X(N_Slope, Lape_Lines(N_Slope, NLINE, 0))) / 2, _
((Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, NLINE, 1)) + Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, NLINE, 0))) / 2) - coef _
), "&H00C0C0C0"
    
        Obj.DrawWidth = 1
        Obj.FontSize = FontSize
        Obj.FontBold = b
'        Msg = ConvertData(Msg)
        Obj.Print ConvertData(Msg)
    
    End If
    
End Sub



Public Function GetHiWord(dw As Long) As Integer
    If dw& And &H80000000 Then
        GetHiWord% = (dw& / 65535) - 1
        Else: GetHiWord% = dw& / 65535
    End If

End Function


Public Function GetLoWord(dw As Long) As Integer
    If dw& And &H8000& Then
        GetLoWord% = &H8000 Or (dw& And &H7FFF&)
        Else: GetLoWord% = dw& And &HFFFF&
    End If

End Function


Public Sub Navigate(frm As Form, ByVal NavTo As String)
    Const SW_SHOW = 5
    Dim hBrowse As Long
    hBrowse = ShellExecute(frm.hwnd, "open", NavTo, "", "", SW_SHOW)
End Sub


Public Sub SendTo(frm As Form, MailTo As String, Subject As String, Message As String)
    Const SW_SHOW = 5
'    ShellExecute frm.hwnd, _
'        "open", _
'        "mailto:" & MailTo & _
'        "?cc=" & "" & _
'        "&bcc=" & "" & _
'        "&subject=" & Subject & _
'        "&body=" & Replace(Message, vbCrLf, _
'            "%0D%0A"), _
'        vbNullString, vbNullString, _
'        SW_SHOW

    ShellExecute frm.hwnd, _
        "open", _
        "mailto:" & MailTo & _
        "?subject=" & Subject & _
        "&body=" & Replace(Message, vbCrLf, _
            "%0D%0A"), _
        vbNullString, vbNullString, _
        SW_SHOW
End Sub


Public Sub SetFont(pfrmIn As Form, Optional use_prt As Byte)
    On Error Resume Next
    Dim FormControl As Control
    For Each FormControl In pfrmIn
        FormControl.Font.Charset = lng.LngCharset ' 0 2 128 255 204 186 172
        If use_prt Then
            FormControl.Font = Gl.PrintFont
            FormControl.size = Gl.PrintFontSize
        Else
            FormControl.Font = Gl.WindowsFont
            FormControl.size = Gl.WindowsFontSize
        End If
    Next
End Sub



Public Sub SetCenter(Obj, HScroll, N_Slope As Integer)
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
      
      If RatioW * FigureLenght >= HScroll.MAX Then GoTo L96EE
      If FigureLenght / FigureHeight > RatioH Then

          l00A6 = RatioH * FigureLenght
          l00AE = l00A6 / RatioW

          If l00A6 > HScroll.MAX Then
              HScroll.value = HScroll.MAX
              GoTo L96EE
          End If

          Obj.ScaleLeft = XMin - 0.1 * FigureLenght
          Obj.ScaleWidth = l00A6
          Obj.ScaleHeight = -l00AE
          Obj.ScaleTop = (YMAx - (FigureHeight / 2)) + (l00AE / 2)

'            Obj.ScaleLeft = XMin - 0.1 * FigureLenght
'            Obj.ScaleWidth = RatioW * FigureLenght
'            l004A = Obj.ScaleWidth
'            l004C = Obj.ScaleWidth / RatioH
'            Obj.ScaleHeight = -l004C
'            Obj.ScaleTop = (YMAx - (FigureHeight / 2)) + (l004C / 2)

      Else

          l00AE = -RatioH * FigureHeight
          l00A6 = l00AE * -RatioW

          If l00A6 > HScroll.MAX Then
              HScroll.value = HScroll.MAX
              GoTo L96EE
          End If

          Obj.ScaleLeft = (XMax - (FigureLenght / 2)) - (l00A6 / 2)
          Obj.ScaleWidth = l00A6
          Obj.ScaleHeight = l00AE
          Obj.ScaleTop = YMAx + 0.1 * FigureHeight

'            Obj.ScaleTop = YMAx + 0.1 * FigureHeight
'            Obj.ScaleHeight = -RatioW * FigureHeight
'            l004A = Obj.ScaleHeight * -RatioH
'            Obj.ScaleWidth = l004A
'            Obj.ScaleLeft = (XMax - (FigureLenght / 2)) - (l004A / 2)

      End If

      If HScroll.MIN > l00A6 Then
          HScroll.value = 1 'HScroll.MIN
      Else
          ' Не центрировать если уже все выровнено
          If HScroll.value <> Round(l00A6) Then
              HScroll.value = l00A6
          End If
      End If

L96EE:
End Sub


Sub Draw_Systems(MainPic As PictureBox, Optional TXTERR As String, Optional lshift As Integer)
    Dim l0138 As Single
    Dim l013A As Single
    Dim l013C As Single
    Dim l013E As Single
    Dim l0140
    Dim i As Integer
    Dim l0144 As Integer

    On Error GoTo ERR

    ''''''''''''''''''''''''''
    MainPic.Cls
    OfficeStart.StatusBar.Panels(4).Picture = Pic
    ''''''''''''''''''''''''''
    l0138 = MainPic.ScaleWidth / 90
    l013A = 2 * l0138
    l013C = MainPic.ScaleLeft + 0.95 * MainPic.ScaleWidth
    l013E = MainPic.ScaleTop + 0.95 * MainPic.ScaleHeight

    If TXTERR <> "" Then
ERR:
        'OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.23." & ERR.Source & "]", ERR.Number, ERR.Description
        Dim size As Single
        size = MainPic.FontSize
        MainPic.FontSize = 12
        MainPic.Cls
        MainPic.PSet (MainPic.ScaleLeft + 0.22 * MainPic.ScaleWidth, MainPic.ScaleTop + 0.42 * MainPic.ScaleHeight), "&H00DCFBFC"
        MainPic.Print TXTERR
        MainPic.FontSize = size
        GoTo FEXIT
        
    End If
    
    If MainCountOfLines > 0 Then
        MainPic.DrawWidth = Setup.Text3.Text
        For i = 1 To MainCountOfLines Step 1
            MainPic.Line (Main_Points_X(Points_m_A(i)), Main_Points_Y(Points_m_A(i)))-(Main_Points_X(Points_m_B(i)), Main_Points_Y(Points_m_B(i))), Setup.Command10.BackColor
        Next i
    End If
    
    On Error GoTo ERR


    If OptionDMM = "Msel" And IsLoadForm("ROOFPIC") Then
        MainPic.DrawWidth = 1
        MainPic.ForeColor = vbRed
        MainPic.FontBold = True
        If P_A > 0 Then
            MainPic.PSet (Main_Points_X(P_A) - l0138, Main_Points_Y(P_A) + l013A), MainPic.BackColor
            MainPic.Print "A"
        End If

        If P_B > 0 Then
            MainPic.PSet (Main_Points_X(P_B) - l0138, Main_Points_Y(P_B) + l013A), MainPic.BackColor
            MainPic.Print "B"
        End If
        MainPic.FontBold = False
        MainPic.ForeColor = vbBlack
    End If
  
    If MainPic.name <> "ShowMP" Then
        MainPic.DrawWidth = 1
        For i = 1 To KolvoScatov Step 1
            
            If SlP(i).CountSheets > 0 Then
                Module10.Draw_lape_label MainPic, i, Label_X(i), Label_Y(i), vbBlue
            ElseIf SlP(i).CountOfLines > 0 Then
                Module10.Draw_lape_label MainPic, i, Label_X(i), Label_Y(i), vbRed
            Else
                Module10.Draw_lape_label MainPic, i, Label_X(i), Label_Y(i), "&H808080"
            End If
            
        Next i
    End If

FEXIT:
    OfficeStart.StatusBar.Panels(4).Picture = Nothing
End Sub
