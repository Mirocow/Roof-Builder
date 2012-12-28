Attribute VB_Name = "FileOperation"
Public FileVersion As Byte

Function Ld_(PrjDir As String, ProjecFile As String)
Call VarPtr("VMProtect begin")
     On Error GoTo ERR
     
     If IsLic = False Then GoTo NOLIC
     
'     OfficeStart.MousePointer = 11

     If Right(PrjDir, 1) <> "\" Then PrjDir = PrjDir & "\"

     If Right(ProjecFile, 4) = ".rfd" Then
         Gl.FileNameExtension = ".rfd"
     ElseIf Right(ProjecFile, 4) = ".rbp" Then
         Gl.FileNameExtension = ".rbp"
     Else
         MsgBox lng.GetResIDstring(1475), vbCritical
'         Screen.MousePointer = 0
         Exit Function
     End If

     OfficeStart.StatusBar.Panels(2) = "Try open: " & PrjDir & ProjecFile
     
     ' Dim cf As FileMan.clsFile
     ' Set cf = New clsFile
     
     Dim cf As Object
     Set cf = Setup.ws_Getdata(True)
     If cf Is Nothing Then Exit Function
     
     If cf.FOpen(PrjDir & ProjecFile, 0) Then

        If cf.FN = 0 Then GoTo ERR
        If cf.FLOF() = 0 Then GoTo ERR
         
         OfficeStart.StatusBar.Panels(3) = Format(cf.FLOF() / 1024, "0.00") & " kb"
    
         If Left(PrjDir, 2) <> "\\" And Left(ProjecFile, 2) <> "~$" Then
             OfficeStart.AddAtmel LCase(PrjDir & ProjecFile)
         End If
    
         GetFileData cf ' Чтение данных
    
         cf.FClose
         
     End If
     Set cf = Nothing

     OfficeStart.StatusBar.Panels(3) = OfficeStart.StatusBar.Panels(3) & "/v: " & FileVersion

     N_Slope = 1
     isSave = False
     Ld_ = ProjecFile

     OfficeStart.StatusBar.Panels(2) = PrjDir & ProjecFile & " Load [OK]"
'     OfficeStart.MousePointer = 0
     Exit Function

ERR:

     If Not cf Is Nothing Then cf.FClose
     Set cf = Nothing
'     OfficeStart.MousePointer = 0
     isSave = False

     OfficeStart.StatusBar.Panels(2) = PrjDir & ProjecFile & " Load [ERROR]"
     If IsLic Then OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.13." & ERR.Source & "]", ERR.Number, ERR.Description & " FILE: " & PrjDir & ProjecFile

     If Left(PrjDir, 2) <> "\\" Then
         MsgBox lng.GetResIDstring(1483, "%CATALOGUE%", PrjDir, "%FILE%", ProjecFile), vbCritical, lng.GetResIDstring(1413)
     Else
         MsgBox lng.GetResIDstring(1484, "%CATALOGUE%", PrjDir, "%FILE%", ProjecFile), vbCritical, lng.GetResIDstring(1413)
     End If

     OfficeStart.Clear_project
     OfficeStart.Enabled = True
     Exit Function
        
NOLIC:
      Module10.withoutl
      OfficeStart.Enabled = True
      
Call VarPtr("VMProtect end")
End Function


Public Function GetFileData(ByRef cf As Object) As Boolean

    Dim len_of_data As Single
    Dim istr As Integer
    Dim str As String
    Dim dstr As Double
    Dim i As Integer
    Dim n As Integer

    On Error GoTo ERR

    ' Начало загрузки данных проекта
    GetDataProject cf, False

    ' Начало загрузки переменных главного рисунка
    GetMainData cf

    cf.FRead KolvoScatov
    If KolvoScatov > MAXSLOPES Then GoTo ERR
    For i = 1 To KolvoScatov Step 1
        cf.FRead Label_X(i)
        cf.FRead Label_Y(i)
    Next i
    ' Конец

    cf.FReadString MainDescrib, 1
    
    Dim ProfilName As String

    ' Чтение поскатно
    For N_Slope = 1 To KolvoScatov Step 1

        ' Чтение данных профиля
        If FileVersion >= 4 And Gl.FileNameExtension = ".rbp" Then
            
            cf.FReadString SlP(N_Slope).ProfilName, 1 ' Имя профиля проекта
            
            If TrimNullChar(SlP(N_Slope).ProfilName) = "" Then
                Exit Function
            End If
            
            If FileVersion >= 10 Then
                ' later
                AddProfilToBD cf, SlP(N_Slope).ProfilName, SlP(N_Slope).Factory_Name, False
            Else
                ' ver 9
                AddProfilToBD_9 cf, SlP(N_Slope).ProfilName, SlP(N_Slope).Factory_Name, False
            End If
            
        End If

        cf.FRead SlP(N_Slope).ScaleLeftS
        cf.FRead SlP(N_Slope).ScaleWidthS
        cf.FRead SlP(N_Slope).ScaleTopS
        cf.FRead SlP(N_Slope).ScaleHeightS
        cf.FRead SlP(N_Slope).CountOfPoints

        For i = 1 To SlP(N_Slope).CountOfPoints Step 1
            cf.FRead Lape_Points_X(N_Slope, i)
            cf.FRead Lape_Points_Y(N_Slope, i)
        Next i

        cf.FRead SlP(N_Slope).CountOfLines
        For i = 1 To SlP(N_Slope).CountOfLines Step 1
            cf.FRead Lape_Lines(N_Slope, i, 0)
            cf.FRead Lape_Lines(N_Slope, i, 1)
        Next i

        cf.FRead SlP(N_Slope).Pn_Red_lines
        cf.FRead SlP(N_Slope).PX_StartLC
        cf.FRead SlP(N_Slope).Pn_StartLC

        cf.FRead SlP(N_Slope).CountSheets
        
        If FileVersion > 0 Then
            If FileVersion = 2 Or FileVersion = 3 Then cf.FRead istr ' Ver 2,3
            If FileVersion >= 5 Then
                cf.FRead SlP(N_Slope).Sf
                cf.FRead SlP(N_Slope).Sw
            End If
        Else
            istr = 2
        End If

        Dim nlist As Integer
        nlist = 0
        If SlP(N_Slope).CountSheets > 0 Then nlist = 1
        
        ' Чтение по листам
        For i = 1 To SlP(N_Slope).CountSheets Step 1

            If FileVersion >= 4 And FileVersion < 8 Then
                cf.FRead istr
            End If

            Dim Data As Single
            For n = 0 To istr Step 1
                cf.FRead Data
                If Data <> 0 Then
                    List_Properties_PY(N_Slope, nlist) = Data
                End If
                cf.FRead Data
                If Data <> 0 Then
                    List_Properties_PX(N_Slope, nlist) = Data
                End If
                cf.FRead Data
                If Data > 0 Then
                    List_Properties_Length(N_Slope, nlist) = Data
                    SlP(N_Slope).ListLength = SlP(N_Slope).ListLength + List_Properties_Length(N_Slope, nlist)
                    nlist = nlist + 1
                End If
            Next n

        Next i
        
        If SlP(N_Slope).CountSheets > 0 Then SlP(N_Slope).CountSheets = nlist

        cf.FReadString SlP(N_Slope).Describ, 1
        If FileVersion <= 3 Then cf.FRead len_of_data

    Next N_Slope

    GetFileData = True
    Exit Function
ERR:
    GetFileData = False
End Function


Public Function GetDataProject(ByRef cf As Object, isPreload As Boolean) As Boolean
    Dim str As String
    Dim len_of_data As Single
    Dim dstr As Double
    Dim Result As Boolean

    On Error GoTo ERR

    cf.FRead FileVersion
    
    ' 5 байтов резев
    cf.fseek cf.fseek + 4
    
    If FileVersion > 8 Then
        Dim CalcType As Byte
        cf.FRead Chr$(1)
    Else
        cf.fseek cf.fseek + 1
    End If
    
    ' ПОДПИСЬ
    cf.FRead str, , 2
    
    If str = "N" & Chr$(9) Then
        cf.FRead dstr
        cf.FReadString PrjDescrib, 1
    Else
        cf.fseek 1
        cf.FReadString PrjDescrib, 1
        If len_of_data = 0 Then PrjDescrib = ""
    End If

    cf.FReadString UserCreatProject, 1
    cf.FReadString Profil_Name, 1 ' Имя профиля проекта
    
    Profil_Name = TrimNullChar(Profil_Name)

    If Gl.FileNameExtension = ".rfd" Then
    
        If FileVersion >= 10 Then
            
            ' later
            Result = AddProfilToBD(cf, Profil_Name, Factory_Name, isPreload)
            
        ElseIf FileVersion >= 4 Then
            
            ' ver 9
            Result = AddProfilToBD_9(cf, Profil_Name, Factory_Name, isPreload)
            
        End If
    
    End If
    
    cf.FReadString width1, 1
    cf.FReadString cover, 1
    cf.FReadString ColorRoof, 1

    GetDataProject = True
    Exit Function
ERR:
    GetDataProject = False
End Function


Public Function GetMainData(ByRef cf As Object) As Boolean
    Dim i As Integer
    Dim len_of_data As Single

    On Error GoTo ERR

    cf.FRead ScaleLeft_Main
    cf.FRead ScaleWidth_Main
    cf.FRead ScaleTop_Main
    cf.FRead ScaleHeight_Main
    
    cf.FRead MainCountOfPoints
    For i = 1 To MainCountOfPoints Step 1
        cf.FRead Main_Points_X(i)
        cf.FRead Main_Points_Y(i)
    Next i

    cf.FRead MainCountOfLines
    For i = 1 To MainCountOfLines Step 1
        cf.FRead Points_m_A(i)
        cf.FRead Points_m_B(i)
    Next i

    GetMainData = True
    Exit Function
ERR:
    GetMainData = False
End Function


Function AddProfilToBD(ByRef cf As Object, ProfilName As String, Factory_Name As String, isPreload As Boolean) As Boolean
    Dim GeneralWidth As Single  ' общая ширина
    Dim Width As Single ' рабочая ширина
    Dim Step As Single ' длина волы, шаг
    Dim Overlaping As Single ' нахлест
    Dim MinLength As Single ' минимальная длина
    Dim MAXSlopeLength As Single ' максимальная длина
    Dim Heigth As Single ' высота панели
    Dim L1 As Single
    Dim L2 As Single
    Dim IDGRP As Integer
    Dim FactoryID As Integer

    ' Чтение данных
    cf.FRead GeneralWidth '       "WORK_WIDTH"
    cf.FRead Width '       "WIDTH"
    cf.FRead Step '       "STEP"
    cf.FRead Overlaping '       "OVERLAPING"
    cf.FRead MinLength '       "MIN_LENGTH"
    cf.FRead MAXSlopeLength '       "MAX_LENGTH"
    cf.FRead Heigth '       "HEIGHT
    cf.FRead L1 '      "L1"
    cf.FRead L2 '      "L2"
    cf.FRead IDGRP '      "IDGROUP"
    cf.FRead FactoryID '      "IDFACTORY" '  не используется
    
    If isPreload = False Then
    
        If FileVersion >= 7 Then
            cf.FReadString Factory_Name, 0
            Factory_Name = TrimNullChar(Factory_Name)
        Else
            Factory_Name = ""
        End If
        
        FactoryID = 0
        
        If Factory_Name <> "" Then
            ' Проверка на наличие завода в бд
            FactoryID = GetFactoryID(Factory_Name)
            If FactoryID = 0 Then
                
                Connect Gl.FileName, -Val(iData(0))  ' -0 false (ACTIVE), -1 (DEACTIVE)
                ' Добавление в базу нового производителя
                FactoryID = SaveFactoryData(Factory_Name)
                Connect Gl.FileName, True
                
            End If
        End If
        
        ProfilName = TrimNullChar(ProfilName)
    
        ' Если нет подобных заносим их в текущую бд
        Set RS = GetProfilData(ProfilName, FactoryID)
        If RS Is Nothing Then
            
            Connect Gl.FileName, -Val(iData(0))  ' -0 false (ACTIVE), -1 (DEACTIVE)
            ' Добавление в базу нового профиля
            
            Dim idProfil As Integer
            idProfil = 0
            idProfil = SaveProfilData(ProfilName, GeneralWidth, Width, Step, Overlaping, MinLength, MAXSlopeLength, Heigth, L1, L2, 0, IDGRP, FactoryID)
            If idProfil = 0 Then
                ' Сообщение об ошибке
                AddProfilToBD = False
            Else
                AddProfilToBD = True
            End If
            Connect Gl.FileName, True
            
        Else
            RS.Close
            Set RS = Nothing
            AddProfilToBD = False
        End If

    End If

End Function

Function AddProfilToBD_9(ByRef cf As Object, ProfilName As String, Factory_Name As String, isPreload As Boolean) As Boolean
    Dim GeneralWidth As Integer  ' общая ширина
    Dim Width As Integer ' рабочая ширина
    Dim Step As Integer ' длина волы, шаг
    Dim Overlaping As Integer ' нахлест
    Dim MinLength As Integer ' минимальная длина
    Dim MAXSlopeLength As Integer ' максимальная длина
    Dim Heigth As Integer ' высота панели
    Dim L1 As Integer
    Dim L2 As Integer
    Dim IDGRP As Integer
    Dim IDFACT As Integer

    ' Чтение данных
    cf.FRead GeneralWidth '       "WORK_WIDTH"
    cf.FRead Width '       "WIDTH"
    cf.FRead Step '       "STEP"
    cf.FRead Overlaping '       "OVERLAPING"
    cf.FRead MinLength '       "MIN_LENGTH"
    cf.FRead MAXSlopeLength '       "MAX_LENGTH"
    cf.FRead Heigth '       "HEIGHT
    cf.FRead L1 '      "L1"
    cf.FRead L2 '      "L2"
    cf.FRead IDGRP '      "IDGROUP"
    cf.FRead IDFACT '      "IDFACTORY" '  не используется
    
    If isPreload = False Then
    
        If FileVersion >= 7 Then
            cf.FReadString Factory_Name, 0
            Factory_Name = TrimNullChar(Factory_Name)
        Else
            Factory_Name = ""
        End If
        
        If ProfilName = "" Then
            Exit Function
        End If
        
        IDFACT = 0
        
        If Factory_Name <> "" Then
            ' Проверка на наличие завода в бд
            Set RS = RequestSQL("select FirmFactory.id from FirmFactory where FirmFactory.Name=" & "'" & Factory_Name & "'")
            If RS Is Nothing Then
                
                Connect Gl.FileName, -Val(iData(0))  ' -0 false (ACTIVE), -1 (DEACTIVE)
                ' Добавление в базу нового производителя
                IDFACT = SaveFactoryData(Factory_Name)
                'If Not  Then
                'End If
                Connect Gl.FileName, True
                
            Else
                IDFACT = RS!id
            End If
            'RS.Close
            Set RS = Nothing
            ' Занесение завода в бд
        End If
        
        ProfilName = TrimNullChar(ProfilName)
    
        ' Если нет подобных заносим их в текущую бд
        Set RS = RequestSQL("select ProfiName.id from ProfiName where ProfiName.Name=" & "'" & ProfilName & "' and ProfiName.IDFACTORY=" & IDFACT)
        If RS Is Nothing Then
            
            Connect Gl.FileName, -Val(iData(0))  ' -0 false (ACTIVE), -1 (DEACTIVE)
            ' Добавление в базу нового профиля
            
            Dim idProfil As Integer
            idProfil = 0
            idProfil = SaveProfilData(ProfilName, CSng(GeneralWidth), CSng(Width), CSng(Step), CSng(Overlaping), CSng(MinLength), CSng(MAXSlopeLength), CSng(Heigth), CSng(L1), CSng(L2), 0, IDGRP, IDFACT)
            If idProfil = 0 Then
                ' Сообщение об ошибке
                AddProfilToBD_9 = False
            Else
                AddProfilToBD_9 = True
            End If
            Connect Gl.FileName, True
            
        Else
            RS.Close
            Set RS = Nothing
            AddProfilToBD_9 = False
        End If

    End If

End Function


Public Function FileExists(ByRef FileName As String) As Boolean

    If FileName = "" Or Right(FileName, 1) = "\" Then
        FileExists = False
        Exit Function
    End If

    FileExists = (dir(FileName) <> "")
End Function

Public Function DirectoryExists(ByRef TheDirectory As String) As Boolean

    Dim sDummy As String
    On Error Resume Next


    If Right$(TheDirectory, 1) <> "\" Then
        TheDirectory = TheDirectory & "\"
    End If

    sDummy = dir$(TheDirectory & "*.*", vbDirectory)
    DirectoryExists = Not (sDummy = "")
End Function


Public Sub fpt_(Config_Dir, FILE)
Call VarPtr("VMProtect begin")
On Error GoTo ERR

If IsLic = False Or FEXIT Then Exit Sub


'        Dim cf As FileMan.clsFile
'        Set cf = New FileMan.clsFile

'        cf.FOpen App.Path & "\test.bin", 0
'        cf.FOpen App.Path & "\test.bin", 1
'        cf.FOpen App.Path & "\test.bin", 2
'        cf.FWrite CLng(1)
'        cf.FRead lng
'        cf.fseek
'        cf.fseek 5
'        cf.FWriteList CLng(2), CStr("ferhgrtjh fdgdfh")
'        array = cf.FReadList(2)
'        cf.FWriteString str, 0-long
'        cf.FWriteString str, 1-single
'        cf.FReadString f8, 1
'        cf.FClose
        
        Dim cf As Object
        Set cf = Setup.ws_Getdata(True)
        If cf Is Nothing Then Exit Sub
    
        Dim RS As Recordset
        Dim len_of_data As Single
        Dim str As String
        Dim i As Integer
        
        Dim ProfilID As Integer
        Dim ProfilName As String

'        Screen.MousePointer = 11
        If Right(Config_Dir, 1) <> "\" Then Config_Dir = Config_Dir & "\"

        If dir(Config_Dir, vbDirectory) = "" Then Exit Sub

        If dir(Config_Dir & FILE) <> "" Then Kill Config_Dir & FILE

        If cf.FOpen(Config_Dir & FILE, 1) Then

            If cf.FN = 0 Then GoTo ERR
    
            cf.FWrite CByte(FILEVER) ' Версия
            
            ' Зарезирвированное место 4 байта
            cf.fseek cf.fseek + 4
            
            cf.FWrite CByte(1) 'Setup.Combo4.ListIndex) ' В чем будем считать см или мм
            
            '  Подпись "N" & Chr(9) файла проекта
            cf.FWrite CByte(78)
            cf.FWrite CByte(9)
            
            ' Зарезирвированное место 8 байт
            cf.fseek cf.fseek + 8
    
            cf.FWriteString Project.Label2.Caption, 1
            cf.FWriteString UserCreatProject, 1
            cf.FWriteString Project.Label3.Caption, 1
    
            If FileNameExtension = ".rfd" Then
            ' Запись данных профиля
            
                ProfilName = Profil_Name
    
                ProfilID = GetProfilID(ProfilName, Project.Label2.Tag)
                Set RS = RequestSQL("select profils.*,ProfiName.idgroup,ProfiName.idfactory from profils,ProfiName where profils.id=ProfiName.ID and ProfiName.ID=" & ProfilID)
                If Not RS Is Nothing Then
                If RS.RecordCount Then
                    
                    cf.FWrite CSng(RS.Fields(CInt(iData(2)))) '       "WORK_WIDTH"
                    cf.FWrite CSng(RS.Fields(CInt(iData(3)))) '       "WIDTH"
                    cf.FWrite CSng(RS.Fields(CInt(iData(4)))) '       "STEP"
                    cf.FWrite CSng(RS.Fields(CInt(iData(5)))) '       "OVERLAPING"
                    cf.FWrite CSng(RS.Fields(CInt(iData(6)))) '       "MIN_LENGTH"
                    cf.FWrite CSng(RS.Fields(CInt(iData(7)))) '       "MAX_LENGTH"
                    cf.FWrite CSng(RS.Fields(CInt(iData(8)))) '       "HEIGHT"
                    cf.FWrite CSng(RS.Fields(CInt(iData(9)))) '       "L1"
                    cf.FWrite CSng(RS.Fields(CInt(iData(10)))) '      "L2"
                    ' WL
                    cf.FWrite CInt(RS.Fields(CInt(iData(12)))) '      "IDGROUP"
                    cf.FWrite CInt(RS.Fields(CInt(iData(13)))) '      "IDFACTORY" 12
                    
                    cf.FWriteString Factory_Name, 0
                    
                    RS.Close
                    Set RS = Nothing
                    
                End If
                Else
                
                    ' Пустышки так как профиль не задан
                    cf.FWrite CSng(0) '       "WORK_WIDTH"
                    cf.FWrite CSng(0) '       "WIDTH"
                    cf.FWrite CSng(0) '       "STEP"
                    cf.FWrite CSng(0) '       "OVERLAPING"
                    cf.FWrite CSng(0) '       "MIN_LENGTH"
                    cf.FWrite CSng(0) '       "MAX_LENGTH"
                    cf.FWrite CSng(0) '       "HEIGHT"
                    cf.FWrite CSng(0) '       "L1"
                    cf.FWrite CSng(0) '      "L2"
                    ' WL
                    cf.FWrite CInt(0) '      "IDGROUP"
                    cf.FWrite CInt(0) '      "IDFACTORY" 12
                    
                    cf.FWriteString "", 0
                        
                End If
    
            End If
    
            '
            ' MAIN
            '
            cf.FWriteString width1, 1
            cf.FWriteString cover, 1
            cf.FWriteString ColorRoof, 1
    
            ' Размеры рисунка крыши
            cf.FWrite ScaleLeft_Main
            cf.FWrite ScaleWidth_Main
            cf.FWrite ScaleTop_Main
            cf.FWrite ScaleHeight_Main
    
            cf.FWrite MainCountOfPoints
            For i = 1 To MainCountOfPoints Step 1
                cf.FWrite Main_Points_X(i)
                cf.FWrite Main_Points_Y(i)
            Next i
    
            cf.FWrite MainCountOfLines
            For i = 1 To MainCountOfLines Step 1
                cf.FWrite Points_m_A(i)
                cf.FWrite Points_m_B(i)
            Next i
    
            cf.FWrite KolvoScatov
            For i = 1 To KolvoScatov Step 1
                cf.FWrite Label_X(i)
                cf.FWrite Label_Y(i)
            Next i
    
            cf.FWriteString MainDescrib
            '
            ' MAIN END
            '
            
            Dim FactoryName As String
            Dim Slope As Integer
    
            ' Данные по скатам
            For Slope = 1 To KolvoScatov Step 1
    
                If FileNameExtension = ".rbp" Then
                
                    ProfilName = Trim(SlP(Slope).ProfilName)
                    
                    ' Запись данных профиля
                    cf.FWriteString ProfilName, 1
                    
                    FactoryName = TrimNullChar(SlP(Slope).Factory_Name)
    
                    If FactoryName <> "" Then
                        ProfilID = GetProfilID(ProfilName, GetFactoryID(FactoryName))
                    Else
                        ProfilID = 0
                    End If
                    
                    Set RS = RequestSQL("select profils.*,ProfiName.idgroup,ProfiName.idfactory from profils,ProfiName where profils.id=ProfiName.ID and ProfiName.ID=" & ProfilID)
                    If Not RS Is Nothing Then
                    If RS.RecordCount Then
                        
                        cf.FWrite CSng(RS.Fields(CInt(iData(2)))) '       "WORK_WIDTH"
                        cf.FWrite CSng(RS.Fields(CInt(iData(3)))) '       "WIDTH"
                        cf.FWrite CSng(RS.Fields(CInt(iData(4)))) '       "STEP"
                        cf.FWrite CSng(RS.Fields(CInt(iData(5)))) '       "OVERLAPING"
                        cf.FWrite CSng(RS.Fields(CInt(iData(6)))) '       "MIN_LENGTH"
                        cf.FWrite CSng(RS.Fields(CInt(iData(7)))) '       "MAX_LENGTH"
                        cf.FWrite CSng(RS.Fields(CInt(iData(8)))) '       "HEIGHT"
                        cf.FWrite CSng(RS.Fields(CInt(iData(9)))) '       "L1"
                        cf.FWrite CSng(RS.Fields(CInt(iData(10)))) '      "L2"
                        ' WL
                        cf.FWrite CInt(RS.Fields(CInt(iData(12)))) '      "IDGROUP"
                        cf.FWrite CInt(RS.Fields(CInt(iData(13)))) '      "IDFACTORY" 12
                        
                        RS.Close
                        Set RS = Nothing
                        
                        cf.FWriteString FactoryName, 0
                        
                    End If
                    Else
                        
                        ' Пустышки так как профиль не задан
                        cf.FWrite CSng(0) '       "WORK_WIDTH"
                        cf.FWrite CSng(0) '       "WIDTH"
                        cf.FWrite CSng(0) '       "STEP"
                        cf.FWrite CSng(0) '       "OVERLAPING"
                        cf.FWrite CSng(0) '       "MIN_LENGTH"
                        cf.FWrite CSng(0) '       "MAX_LENGTH"
                        cf.FWrite CSng(0) '       "HEIGHT"
                        cf.FWrite CSng(0) '       "L1"
                        cf.FWrite CSng(0) '      "L2"
                        ' WL
                        cf.FWrite CInt(0) '      "IDGROUP"
                        cf.FWrite CInt(0) '      "IDFACTORY" 12
                        
                        cf.FWriteString "", 0
                        
                    End If
    
                End If
                
                '
                ' SLOPE
                '
                
                ' Размеры чертежа
                cf.FWrite SlP(Slope).ScaleLeftS
                cf.FWrite SlP(Slope).ScaleWidthS
                cf.FWrite SlP(Slope).ScaleTopS
                cf.FWrite SlP(Slope).ScaleHeightS
    
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
    
                cf.FWrite SlP(Slope).Pn_Red_lines
                cf.FWrite SlP(Slope).PX_StartLC
                cf.FWrite SlP(Slope).Pn_StartLC
                cf.FWrite SlP(Slope).CountSheets
    
                cf.FWrite SlP(Slope).Sf ' Площадь фигуры
                cf.FWrite SlP(Slope).Sw ' Площадь покрытия
                
                '
                ' Раскрой
                '
                For i = 1 To SlP(Slope).CountSheets Step 1
                    cf.FWrite List_Properties_PY(Slope, i)
                    cf.FWrite List_Properties_PX(Slope, i)
                    cf.FWrite List_Properties_Length(Slope, i)
                Next i
    
                ' Описание каждой плоскости
                cf.FWriteString SlP(Slope).Describ, 1
                
                ' Вспомогательные линии (направляющие)
                
                '
                ' END SLOPE
                '
    
            Next Slope
    
            cf.FWriteString App.ProductName & " " & Ver$ & "  " & Date$
            OfficeStart.StatusBar.Panels(3) = Format(cf.FLOF / 1024, "0.00") & " kb"
            cf.FClose
        
        End If
        Set cf = Nothing

        isSave = False
        If Left(Config_Dir, 2) <> "\\" And Left(FILE, 2) <> "~$" Then
            OfficeStart.AddAtmel LCase(Config_Dir & FILE)
        End If

'        OfficeStart.StatusBar.Panels(2) = Config_Dir & FILE
        OfficeStart.StatusBar.Panels(2) = Config_Dir & FILE & " Save [OK]"
'        Screen.MousePointer = 0
        Exit Sub

ERR:
        If Not cf Is Nothing Then cf.FClose
        Set cf = Nothing
'        Screen.MousePointer = 0
        OfficeStart.StatusBar.Panels(2) = Config_Dir & FILE & " Save [ERROR]"
        If IsLic Then OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.14." & ERR.Source & "]", ERR.Number, ERR.Description & " FILE: " & Config_Dir & FILE
        
Call VarPtr("VMProtect end")
End Sub


Public Function GetFileExtension(FileName As String)
GetFileExtension = Right(FileName, Len(FileName) - InStrRev(FileName, ".", -1) + 1)
End Function
