Attribute VB_Name = "Files"
Const MAX_PATH = 255

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private TemparyPosition() As Long

Private Const MAXBACKSSPACE = 4000


Public Function GetTempDir() As String
    Dim sRet As String, lngLen As Long
    
    'create buffer
    sRet = String(MAX_PATH, 0)

    lngLen = GetTempPath(MAX_PATH, sRet)
    If lngLen = 0 Then ERR.Raise ERR.LastDllError
    GetTempDir = Left$(sRet, lngLen)
End Function

Public Function CheckWriteDir(folder As String) As Boolean
    Dim lngAttr As Long
    
    lngAttr = GetAttr(folder)
    If lngAttr < 33 Then
        If lngAttr = (vbDirectory Or vbNormal) Or _
            lngAttr = (vbArchive Or vbDirectory Or vbNormal) Then
            CheckWriteDir = True
        ElseIf lngAttr = (vbDirectory Or vbReadOnly) Or _
            lngAttr = (vbArchive Or vbDirectory Or vbReadOnly) Then
            CheckWriteDir = False
        End If
    Else
        If dir(folder, vbDirectory) <> "" Then
            CheckWriteDir = True
        Else
            CheckWriteDir = False
        End If
    End If
    
End Function


'
' ‘ункци€ записи временных точек со смещением
'
Public Function SaveF(fpath As String, Slope As Integer, Optional position As Integer = 0) As Long
  
    Dim i As Integer
    Dim cf As FileMan.clsFile
    
    On Error GoTo ERR
    
    Set cf = New clsFile
    
    If cf.FOpen(fpath, aRead_Write) Then ' 100 шагов + 100 байт под служебную инфу
    
        If cf.FN = 0 Then GoTo ERR
    
        ' /* —Ћ”∆≈ЅЌјя „ј—“№
        ' ReadSystemInfo cf
        
        If ArraySize(TemparyPosition) > 0 Then
            If Positions > position Then
                Positions = position
                ReDim Preserve TemparyPosition(position)
            End If
            cf.fseek TemparyPosition(position)
        Else
            ReDim Preserve TemparyPosition(0)
            TemparyPosition(0) = MAXBACKSSPACE
            cf.fseek MAXBACKSSPACE
        End If
        ' —Ћ”∆≈ЅЌјя „ј—“№ */
        
        If IsLoadForm("Lapepic") Then ' ќкно расчета
            
            ' ”правл€ющие флаги
            cf.FWriteString OptionDMM, vSingle
            cf.FWrite Lapepic.sTabFx1.CurrentTab
            
            ' —ам чертеж
            cf.FWrite SlP(Slope).CountOfPoints
            For i = 1 To SlP(Slope).CountOfPoints Step 1
                cf.FWrite Lape_Points_X(Slope, i)
                cf.FWrite Lape_Points_Y(Slope, i)
            Next i
            
            cf.FWrite SlP(Slope).CountOfLines
            For i = 1 To SlP(Slope).CountOfLines Step 1
                cf.FWrite Lape_Lines(Slope, i, 0)
                cf.FWrite Lape_Lines(Slope, i, 1)
            Next i
            
            ' ’аррактеристики
            cf.FWrite SlP(Slope).CountOfLines '  оличество линий на чертеже
            cf.FWrite SlP(Slope).CountOfPoints '  оличество точек на чертеже
            ' –азмеры чертежа
            cf.FWrite SlP(Slope).ScaleLeftS
            cf.FWrite SlP(Slope).ScaleWidthS
            cf.FWrite SlP(Slope).ScaleTopS
            cf.FWrite SlP(Slope).ScaleHeightS
            ' ƒанные расчета
            cf.FWrite SlP(Slope).Pn_Red_lines ' номер точки через которую проходит красна€ лини€
            cf.FWrite SlP(Slope).Pn_StartLC ' номер точки через которую проходит  лини€ начала разлиновки
            cf.FWrite SlP(Slope).PX_StartLC ' координаты по X линии начала разлиновки
            cf.FWrite SlP(Slope).CountSheets ' количество полос
            cf.FWrite SlP(Slope).ListLength ' ƒлина листов ската  (ѕагонна€ длина)
            cf.FWrite SlP(Slope).Sf ' ѕлощадь плоскости
            cf.FWrite SlP(Slope).Sw ' ѕлощадь покрыти€ по рабочей ?
            
            ' –аскрой
            cf.FWrite SlP(Slope).CountSheets
            For i = 1 To SlP(Slope).CountSheets Step 1
                cf.FWrite List_Properties_PY(Slope, i)
                cf.FWrite List_Properties_PX(Slope, i)
                cf.FWrite List_Properties_Length(Slope, i)
            Next i
        
        ElseIf IsLoadForm("ROOFPIC") Then
        
            ' ”правл€ющие переменные
            cf.FWriteString OptionDMM, vSingle
            cf.FWrite P_A
            cf.FWrite P_B
            cf.FWrite ROOFPIC.sTabFx1.CurrentTab
        
            '# PICTURE WIDTH HEIGHT
            cf.FWrite ScaleHeight_Main
            cf.FWrite ScaleWidth_Main
            cf.FWrite ScaleLeft_Main
            cf.FWrite ScaleTop_Main
            
            '# MAIN
            cf.FWrite KolvoScatov
            ' ѕеременные окна RoofPic
            cf.FWrite LapeName
            
            ' количество точек
            cf.FWrite MainCountOfPoints
            For i = 0 To MainCountOfPoints Step 1
                cf.FWrite Main_Points_X(i)
                cf.FWrite Main_Points_Y(i)
            Next
            
            ' количество расставленых обозначений
            cf.FWrite MainCountOfLines
            For i = 1 To MainCountOfLines Step 1
                cf.FWrite Label_X(i)
                cf.FWrite Label_Y(i)
                '
                cf.FWrite Points_m_A(i)
                cf.FWrite Points_m_B(i)
            Next
        
        End If

        ' /* —Ћ”∆≈ЅЌјя „ј—“№
        position = position + 1
        ReDim Preserve TemparyPosition(position)
        TemparyPosition(position) = cf.fseek
        SaveF = position
        
        SaveSystemInfo cf
        ' —Ћ”∆≈ЅЌјя „ј—“№ */
        
        cf.FClose
        
    End If
    Set cf = Nothing

Exit Function
ERR:
SaveF = 0
End Function


Public Function ReadF(fpath As String, Slope As Integer, position As Integer) As Long

    Dim i As Integer
    Dim cf As FileMan.clsFile
    Set cf = New clsFile
    
    If cf.FOpen(fpath, aRead) Then
    
        If cf.FN = 0 Then GoTo ERR
        If cf.FLOF() = 0 Then GoTo ERR
    
        ' /* —Ћ”∆≈ЅЌјя „ј—“№
        ReadSystemInfo cf
        
        If ArraySize(TemparyPosition) > 0 Then
            cf.fseek TemparyPosition(position - 1)
        Else
            cf.fseek MAXBACKSSPACE
        End If
        ' —Ћ”∆≈ЅЌјя „ј—“№ */
        
        If IsLoadForm("Lapepic") Then ' ќкно расчета
        
            ' ”правл€ющие флаги
            cf.FReadString OptionDMM, vSingle
            cf.FRead i
            Lapepic.sTabFx1.SelectTab i
            
            ' —ам чертеж
            cf.FRead SlP(Slope).CountOfPoints
    '        If SlP(Slope).CountOfPoints = 0 Then Exit Function
            For i = 1 To SlP(Slope).CountOfPoints Step 1
                cf.FRead Lape_Points_X(Slope, i)
                cf.FRead Lape_Points_Y(Slope, i)
            Next i
            
            cf.FRead SlP(Slope).CountOfLines
    '        If SlP(Slope).CountOfLines = 0 Then Exit Function
            For i = 1 To SlP(Slope).CountOfLines Step 1
                cf.FRead Lape_Lines(Slope, i, 0)
                cf.FRead Lape_Lines(Slope, i, 1)
            Next i
            
            ' ’аррактеристики
            cf.FRead SlP(Slope).CountOfLines '  оличество линий на чертеже
            cf.FRead SlP(Slope).CountOfPoints '  оличество точек на чертеже
            ' –азмеры чертежа
            cf.FRead SlP(Slope).ScaleLeftS
            cf.FRead SlP(Slope).ScaleWidthS
            cf.FRead SlP(Slope).ScaleTopS
            cf.FRead SlP(Slope).ScaleHeightS
            ' ƒанные расчета
            cf.FRead SlP(Slope).Pn_Red_lines ' номер точки через которую проходит красна€ лини€
            cf.FRead SlP(Slope).Pn_StartLC ' номер точки через которую проходит  лини€ начала разлиновки
            cf.FRead SlP(Slope).PX_StartLC ' координаты по X линии начала разлиновки
            cf.FRead SlP(Slope).CountSheets ' количество полос
            cf.FRead SlP(Slope).ListLength ' ƒлина листов ската  (ѕагонна€ длина)
            cf.FRead SlP(Slope).Sf ' ѕлощадь плоскости
            cf.FRead SlP(Slope).Sw ' ѕлощадь покрыти€ по рабочей ?
            
            ' –аскрой
            cf.FRead SlP(Slope).CountSheets
            For i = 1 To SlP(Slope).CountSheets Step 1
                cf.FRead List_Properties_PY(Slope, i)
                cf.FRead List_Properties_PX(Slope, i)
                cf.FRead List_Properties_Length(Slope, i)
            Next i
            
        ElseIf IsLoadForm("ROOFPIC") Then
        
            ' ”правл€ющие переменные
            cf.FReadString OptionDMM, vSingle
            cf.FRead P_A
            cf.FRead P_B
            cf.FRead i
            ROOFPIC.sTabFx1.SelectTab i
        
            '# PICTURE WIDTH HEIGHT
            cf.FRead ScaleHeight_Main
            cf.FRead ScaleWidth_Main
            cf.FRead ScaleLeft_Main
            cf.FRead ScaleTop_Main
            
            '# MAIN
            cf.FRead KolvoScatov
            ' ѕеременные окна RoofPic
            cf.FRead LapeName
            
            ' количество точек
            cf.FRead MainCountOfPoints
            For i = 0 To MainCountOfPoints Step 1
                cf.FRead Main_Points_X(i)
                cf.FRead Main_Points_Y(i)
            Next
            
            ' количество расставленых обозначений
            cf.FRead MainCountOfLines
            For i = 1 To MainCountOfLines Step 1
                cf.FRead Label_X(i)
                cf.FRead Label_Y(i)
                '
                cf.FRead Points_m_A(i)
                cf.FRead Points_m_B(i)
            Next
        
        End If
        
        ' /* —Ћ”∆≈ЅЌјя „ј—“№
        
        ' —Ћ”∆≈ЅЌјя „ј—“№ */
        
        cf.FClose
        ReadF = position
    
    End If
    Set cf = Nothing
    
Exit Function
ERR:
ReadF = 0
End Function


Public Sub HistoryClear(save As Boolean)
On Error GoTo ERR
    
    ' ќбнуление механизма Ўј√ назад/вперед
    OfficeStart.Toolbar1.Buttons(7).Enabled = False
    OfficeStart.Toolbar1.Buttons(6).Enabled = False
    OfficeStart.menuRedo.Enabled = False
    OfficeStart.menuUndo.Enabled = False
    
    If TemporaryFileName <> "" Then
        If dir(Gl.TemporaryFileName, vbNormal) <> "" Then _
        Kill TemporaryFileName
        If save Then
            'Positions = SaveF(TemporaryFileName, N_Slope)
            SetChange True
            CurentPosition = Positions = 1
        End If
    End If
    
Exit Sub
ERR:
    SetChange False
End Sub


''
' PRIVATE
''

Private Function ReadSystemInfo(cf As clsFile)

        Erase TemparyPosition
        cf.fseek 1
        cf.FReadArray TemparyPosition
        
End Function


Private Function SaveSystemInfo(cf As clsFile)
        
        cf.fseek 1
        cf.FWriteArray TemparyPosition
        
End Function
