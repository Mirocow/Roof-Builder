Attribute VB_Name = "MainBaseFunction"
Public isMDAC As Boolean
Public verMdac As String
Public DBConnect As Boolean
Private DB As dao.Database

'===========================================================
' Function to check the version of MDAC and istall if < 2.8
'===========================================================
Public Function checkMDAC() As Boolean
    
    On Error Resume Next
    
    Dim strFile, strDir, strMDAC
    Dim varEnvironment
    Dim objFSO
    
    Dim strVersion
    Dim objShell

    Set objShell = CreateObject("wscript.shell")

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set varEnvironment = objShell.Environment("Process")
    
    'Clean up
    Set objShell = Nothing
    WScript.Quit

    
    strFile = "MSDASQL.DLL"
    strDir = varEnvironment("COMMONPROGRAMFILES") & "\system\ole db\"

    strMDAC = objFSO.GetFileVersion(strDir & strFile)
    
    Select Case strMDAC
'        Case "2.10.4202.0"
'            ver = "MDAC 2.1 SP2"
'        Case "2.50.4403.6"
'            ver = "MDAC 2.5"
'        Case "2.51.5303.2"
'            ver = "MDAC 2.5 SP1"
'        Case "2.52.6019.0"
'            ver = "MDAC 2.5 SP2"
'        Case "2.53.6200.0"
'            ver = "MDAC 2.5 SP3"
'        Case "2.60.6526.0"
'            ver = "MDAC 2.6 RTM"
'        Case "2.61.7326.0"
'            ver = "MDAC 2.6 SP1"
'        Case "2.62.7926.0"
'            ver = "MDAC 2.6 SP2"
'        Case "2.62.7400.0"
'            ver = "MDAC 2.6 SP2 Refresh"
'        Case "2.70.7713.0"
'            ver = "MDAC 2.7 RTM"
'        Case "2.70.9001.0"
'            ver = "MDAC 2.7 Refresh"
'        Case "2.71.9030.0"
'            ver = "MDAC 2.7 SP1"
'        Case " 2.71.9040.2"
'            ver = "MDAC 2.7 SP1 on Windows XP SP1"

        Case "2.80.1022.0"
            verMdac = "MDAC 2.8 RTM"
        Case "2.81.1117.0"
            verMdac = "MDAC 2.8 SP1 on Windows XP SP2"
        Case "2.82.1830.0"
            verMdac = "MDAC 2.8 SP2 on Windows Server 2003 SP1"
        Case Else
            verMdac = strMDAC
    End Select
    
    If Ver = "" Then
        checkMDAC = False
    Else
        checkMDAC = True
    End If
    
    isMDAC = checkMDAC

    Set objFSO = Nothing
    Set varEnvironment = Nothing
    
End Function

Function SeekBase(PathtoBase As String) As String
On Error GoTo ERR
    
    isMDAC = checkMDAC
    If isMDAC = False Then
        Exit Function
    End If

    If Left(PathtoBase, 2) <> "\\" Then
        If dir(PathtoBase, vbNormal) = "" Then
            ' Нет такой бд
            PathtoBase = ""
        End If
    End If

    ' Ищем в папке по уполчанию
    If PathtoBase = "" Then PathtoBase = Gl.ConfigDir & "\materials.mdb"
    If dir(PathtoBase, vbNormal) = "" Then GoTo DBERR

    CloseDB
    ' Проверка на доступность базы данных
    If Connect(PathtoBase, True) Then
        SeekBase = PathtoBase
    Else
DBERR:
        MsgBox lng.GetResIDstring(20003, "%PathtoBase%", PathtoBase), vbCritical 'PathtoBase
    End If

Exit Function
ERR:
'STRERR = STRERR & time & ". (BaseFunction) ... [ERROR] N " & ERR.Number & " " & ERR.Description &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.42." & ERR.Source & "]", ERR.Number, ERR.Description
End Function


Function Connect(BaseName As String, Optional isReadOnly As Boolean, Optional login As String, Optional pass As String) As Boolean
    On Error GoTo ERR
    
    isMDAC = checkMDAC
    If isMDAC = False Then
        Exit Function
    End If

    If Not DB Is Nothing Then
        DB.Close
        Set DB = Nothing
    End If
    
    ' And (GetAttr(BaseName) = vbArchive And vbReadOnly)
    Set DB = OpenDatabase(BaseName, False, isReadOnly, "MS Access;PWD=")
        
    If Not DB Is Nothing Then
        DBConnect = True
        Connect = True
    End If
    
Exit Function

ERR:

    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.43." & ERR.Source & "]", ERR.Number, ERR.Description
    If ERR.Number = 429 Then
        MsgBox "Connect:" & vbNewLine & ERR.Number & " " & ERR.Description & vbNewLine & lng.GetResIDstring(20004), vbCritical
    Else
'        MsgBox "Connect:" & vbNewLine & ERR.Number & " " & ERR.Description, vbCritical
        Connect = False
        DBConnect = False
    End If

End Function


Function CloseDB() As Boolean
    Set DB = Nothing
End Function


Function RequestSQL(sql As String) As Recordset
    On Error GoTo ERR
    If DBConnect = False Then Exit Function
    
    Dim RS As Recordset
    Set RS = DB.OpenRecordset(sql, dbOpenDynaset) ', dbReadOnly)

    If RS.EOF Then
        Set RequestSQL = Nothing
    Else
        Set RequestSQL = RS
    End If

    Exit Function
ERR:
    Set RS = Nothing
'        STRERR = STRERR & time & ". (REQSQL: " & sql & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.44." & ERR.Source & "]", ERR.Number, ERR.Description & " in SQL: " & sql
End Function


Function GetProfilData(ProfilName As String, Optional idFactory As Integer) As Recordset
    Dim RS As Recordset
    If DBConnect = False Then Exit Function
    
    On Error Resume Next
        If Factory_Name <> "" Or idFactory > 0 Then
            If idFactory = 0 Then idFactory = GetFactoryID(Factory_Name)
            If idFactory = 0 Then GoTo NOFACTORY
            Set RS = RequestSQL("select * from profils p where p.id=" & GetProfilID(ProfilName, idFactory))
        Else
NOFACTORY:
            Set RS = RequestSQL("select TOP 1 *  from profils p where p.id=" & GetProfilID(ProfilName))
        End If
        If Not RS Is Nothing Then
            Set GetProfilData = RS
        Else
            Beep
            Set GetProfilData = Nothing
        End If
        Set RS = Nothing
End Function


Function GetProfilID(ProfilName As String, Optional idFactory As Integer) As Integer
    Dim RS As Recordset
    If DBConnect = False Then Exit Function
    
    If Factory_Name <> "" Or idFactory > 0 Then
        If idFactory = 0 Then idFactory = GetFactoryID(Factory_Name)
        If idFactory = 0 Then GoTo NOFACTORY
        Set RS = RequestSQL("select ProfiName.ID from ProfiName where ProfiName.Name=" & "'" & Trim(ProfilName) & "' and ProfiName.IDFACTORY=" & idFactory)
    Else
NOFACTORY:
        Set RS = RequestSQL("select TOP 1 ProfiName.ID  from ProfiName where ProfiName.Name=" & "'" & Trim(ProfilName) & "'")
    End If
    If Not RS Is Nothing Then GetProfilID = RS!id
    Set RS = Nothing
End Function


Function GetProfilName(ProfilID As Integer) As String
    Dim RS As Recordset
    If DBConnect = False Then Exit Function
    
    Set RS = RequestSQL("select TOP 1 ProfiName.NAME  from ProfiName where ProfiName.ID=" & ProfilID)
    If Not RS Is Nothing Then GetProfilName = RS!name
    Set RS = Nothing
End Function


Function GetGroupName(GroupID As Integer) As String
    Dim RS As Recordset
    If DBConnect = False Then Exit Function
    
    Set RS = RequestSQL("select GroupName.NAME  from GroupName where GroupName.ID=" & GroupID)
    If Not RS Is Nothing Then GetGroupName = RS!name
    Set RS = Nothing
End Function


Function GetFactoryIDFromProfilID(ProfilID As Integer) As Integer
    Dim RS As Recordset
    Set RS = RequestSQL("select TOP 1 ProfiName.IDFACTORY from ProfiName where ProfiName.id=" & ProfilID)
    If Not RS Is Nothing Then GetFactoryIDFromProfilID = RS!idFactory
    Set RS = Nothing
End Function


Function SaveProfilData(ProfilName As String, b As Single, c As Single, d As Single, e As Single, F As Single, _
                       g As Single, h As Single, i As Single, j As Single, _
                       Lcount As Integer, IDGROUP As Integer, idFactory As Integer) As Integer
On Error GoTo ERR
If DBConnect = False Then Exit Function

Dim RS As Recordset
Dim id As Integer
id = GetProfilID(ProfilName, idFactory)
Set RS = RequestSQL("select * from profils p where p.id=" & id)

If Not RS Is Nothing Then
    
    ' Профиль есть переходим к редактированию
    RS.Edit
    
    ' Запись и данные есть
    ' Обновление только данные
    RS!WORK_WIDTH = b
    RS!Width = c
    RS!Step = d
    RS!Overlaping = e
    RS!MIN_LENGTH = F
    RS!MAX_LENGTH = g
    RS!Height = h
    RS!L1 = i
    RS!L2 = j
    RS!wl = Lcount
    
    DB.Execute "update ProfiName p SET p.IDGROUP=" & IDGROUP & " where p.ID = " & RS!id

    SaveProfilData = RS!id
    RS.UpDate ' Обновляем записи в бд
    RS.Close
    
Else
    
    ' Получаем ID
    Set RS = RequestSQL("select max(id) from ProfiName")
    Dim count As Integer
    count = IIf(IsNull(RS.Fields(0)), 1, RS.Fields(0) + 1)
    RS.Close
    
    ' Профиля нет, добавляем профиль
    DB.Execute "insert into `ProfiName` (ID,Name,IDGROUP,IDFACTORY) values " & _
    "('" & count & "','" & ProfilName & "','" & IDGROUP & "','" & idFactory & "')"
        
    ' Добавляем данные
    DB.Execute "insert into `Profils` (ID,IDNAME,WORK_WIDTH,WIDTH,STEP,OVERLAPING,MIN_LENGTH,MAX_LENGTH,HEIGHT,L1,L2,WL) values " & _
    "('" & count & "','" & count & "','" & b & "','" & c & "','" & d & "','" & e & "','" & F & "','" & g & "','" & h & "','" & i & "','" & j & "','" & Lcount & "')"

    SaveProfilData = count
    
End If
Set RS = Nothing
Exit Function
ERR:
End Function


Function GetFactoryID(FactoryName As String) As Integer
    Dim RS As Recordset
    If DBConnect = False Then Exit Function
    
    On Error Resume Next
        Set RS = RequestSQL("select FirmFactory.ID from FirmFactory where FirmFactory.Name = " & "'" & Trim(FactoryName) & "'")
        If Not RS Is Nothing Then
            GetFactoryID = RS!id
        Else
            Beep
            GetFactoryID = -1
        End If

        Set RS = Nothing
End Function


Function GetFactoryName(id As Integer) As String
    Dim RS As Recordset
    If DBConnect = False Then Exit Function
    On Error Resume Next
        Set RS = RequestSQL("select FirmFactory.NAME from FirmFactory where FirmFactory.ID = " & id)
        If Not RS Is Nothing Then
            GetFactoryName = RS!name
        Else
            Beep
            GetFactoryName = ""
        End If

        Set RS = Nothing
End Function


Function SaveFactoryData(name As String, Optional fid As Integer) As Integer
On Error GoTo ERR
If DBConnect = False Then Exit Function

Dim RS As Recordset
Set RS = RequestSQL("select * from FirmFactory where FirmFactory.ID = " & fid)

If Not RS Is Nothing Then

    SaveFactoryData = RS!id
    
    RS.Edit
    RS!name = name
    RS.UpDate
    RS.Close
    
Else
    
    ' Получаем ID
    Set RS = RequestSQL("select max(id) from FirmFactory")
    Dim count As Integer
    count = IIf(IsNull(RS.Fields(0)), 1, RS.Fields(0) + 1)
    RS.Close
    
    ' Производителя нет, добавляем производителя
    DB.Execute "insert into `FirmFactory` (ID,Name,URL) values " & _
    "('" & count & "','" & name & "','')"
         
    SaveFactoryData = count
    
End If
Set RS = Nothing
Exit Function
ERR:
End Function


Function SetProfilStandartLength(id As Integer, L As Integer)
On Error GoTo ERR
If DBConnect = False Then Exit Function

Dim RS As Recordset
Set RS = RequestSQL("select * from ProfilsWrongLength where idname=" & id & " and LENGTH=" & L)

If Not RS Is Nothing Then
    ' Длина есть переходим к редактированию
    '    RS.Edit
Else

    Dim count As Integer
    ' Профиля нет, надо добавить
    Set RS = RequestSQL("select max(id) from ProfilsWrongLength")
    count = IIf(IsNull(RS.Fields(0)), 1, RS.Fields(0) + 1)
    RS.Close

    ' Добавить данные
    DB.Execute "insert into `ProfilsWrongLength` (ID,IDNAME,LENGTH) values " & _
    "('" & count & "','" & id & "','" & L & "')"

End If

Exit Function
ERR:
'    STRERR = STRERR & time & ". (REQSQL: " & sql & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.45." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Function


Function SetProfilWrongLength(id As Integer, L1 As Integer, L2 As Integer, on_off As Byte)
On Error GoTo ERR
If DBConnect = False Then Exit Function

Dim RS As Recordset
Set RS = RequestSQL("select * from ProfilsWLength where idname=" & id & " and LENGTH1=" & L1)

If Not RS Is Nothing Then
    ' Длина есть переходим к редактированию
    RS.Edit
    RS.Fields(2) = L1
    RS.Fields(3) = L2
    RS.Fields(4) = on_off
    RS.UpDate
    RS.Close
Else

    Dim count As Integer
    ' Профиля нет, надо добавить
    Set RS = RequestSQL("select max(id) from ProfilsWLength")
    count = IIf(IsNull(RS.Fields(0)), 1, RS.Fields(0) + 1)
    RS.Close
    
    ' Добавить данные
    DB.Execute "insert into `ProfilsWLength` (ID,IDNAME,LENGTH1,LENGTH2,ONOFF) values " & _
    "('" & count & "','" & id & "','" & L1 & "','" & L2 & "','" & on_off & "')"

End If

Exit Function
ERR:
'    STRERR = STRERR & time & ". (REQSQL: " & sql & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.46." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Function



Function SetWarehouseLength(id As Integer, L As Integer, a As Integer, on_off As Byte) As Byte
On Error GoTo ERR
If DBConnect = False Then Exit Function

Dim RS As Recordset
Set RS = RequestSQL("select * from Warehouse_profils where idname=" & id & " and LENGTH=" & L)

If Not RS Is Nothing Then
    ' Длина есть переходим к редактированию
    RS.Edit
    RS.Fields(3) = a
    RS.Fields(4) = on_off
    RS.UpDate
    RS.Close
    SetWarehouseLength = 1
Else
    Dim count As Integer
    ' Профиля нет, надо добавить
    Set RS = RequestSQL("select max(id) from Warehouse_profils")
    count = IIf(IsNull(RS.Fields(0)), 1, RS.Fields(0) + 1)
    RS.Close

    ' Добавить данные
    DB.Execute "insert into `Warehouse_profils` (ID,IDNAME,LENGTH,AMOUNT,ONOFF) values " & _
    "('" & count & "','" & id & "','" & L & "','" & a & "','" & on_off & "')"

    SetWarehouseLength = 2

End If

Exit Function
ERR:
'    STRERR = STRERR & time & ". (REQSQL: " & sql & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.47." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Function


Sub LoadProfilAdditionalData(table As String, Obj As ComboBox, idname As Integer, Optional LANG As String)
Dim RS As Recordset
If DBConnect = False Then Exit Sub

On Error Resume Next
    LANG = IIf(IsNumeric(LANG), "ALL", LANG)
    Set RS = RequestSQL("select NAME from " & table & " where (LNG='ALL' or LNG='" & LANG & "') and (IDNAME=" & idname & "  OR IsNull(IDNAME)) group by NAME")
    If Not RS Is Nothing Then
        RS.Sort = "id"
        Do While Not RS.EOF
            Obj.AddItem CheckNull(RS!name)
            RS.MoveNext
        Loop
        RS.Close
    End If
End Sub


Function SetProfilAdditionalData(table As String, value, idname As Integer, Optional LANG As String)
On Error GoTo ERR
If DBConnect = False Then Exit Function

LANG = IIf(IsNumeric(LANG), "ALL", LANG)

Dim RS As Recordset
Set RS = RequestSQL("select * from " & table & " where LNG='" & LANG & "' and NAME='" & value & "' and IDNAME=" & idname)

If RS Is Nothing Then

    Dim count As Integer
    Set RS = RequestSQL("select max(id) from " & table)
    count = IIf(IsNull(RS.Fields(0)), 1, RS.Fields(0) + 1)
    RS.Close

    DB.Execute "insert into `" & table & "` (ID,NAME,LNG,IDNAME) values " & _
    "('" & count & "','" & value & "','" & LANG & "','" & idname & "')"

End If

Exit Function
ERR:
'    STRERR = STRERR & time & ". (REQSQL: " & sql & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.48." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Function



Public Sub GetWarehouseLength(ProfilName As String, FactoryID As Integer, ListAdd As Boolean, Optional arr1, Optional arr2, Optional MIN As Integer, Optional MAX As Integer)
  On Error Resume Next
  If DBConnect = False Then Exit Sub
  
  Dim RS As Recordset
  Set RS = RequestSQL("select * from Warehouse_profils where idname=" & GetProfilID(ProfilName, FactoryID)) '"(select ProfiName.ID from ProfiName where ProfiName.Name='" & ProfilName & "')order by length")
  If Not RS Is Nothing Then
  Dim c As Integer
  If ListAdd Then Lapepic.txt_CL.AddItem 0
  ReDim arr1(0)
  ReDim arr2(0)
  Do While Not RS.EOF
      If RS.Fields(4) Then

          If (MIN > 0 And RS.Fields(2) >= MIN) Or (MIN = 0) Then
              If (MAX > 0 And RS.Fields(2) <= MAX) Or (MAX = 0) Then

                  If ListAdd Then
                      Lapepic.txt_CL.AddItem ConvertData(RS.Fields(2))
                  End If

                  ReDim Preserve arr1(c): arr1(c) = RS.Fields(2)
                  ReDim Preserve arr2(c): arr2(c) = RS.Fields(3)
                  c = c + 1

              End If

          End If

      End If

      RS.MoveNext
  Loop

      RS.Close
  End If

  Set RS = Nothing
End Sub


'
' SIMPLE
'

Public Function CheckNull(vol, Optional todb As Boolean)
On Error Resume Next

If IsNumeric(vol) Then
    CheckNull = IIf(IsNull(vol), 0, vol)
Else
    If todb Then ' Если в базу то Null а не ""
        CheckNull = IIf(vol = "", Null, Trim(vol))
    Else
        CheckNull = IIf(IsNull(vol), "", Trim(vol))
    End If
End If
End Function


Public Function CheckNullNomber(vol, Optional todb As Boolean)
On Error Resume Next

If todb Then ' Если в базу то Null а не 0
    CheckNullNomber = IIf(vol = 0, Null, Trim(vol))
Else
    CheckNullNomber = IIf(IsNull(vol), 0, Trim(vol))
End If

End Function


Public Sub Execute(sql As String)
On Error Resume Next

If DBConnect = False Then Exit Sub
DB.Execute sql
End Sub


Function DelBaseData(sql As String)
Dim RS As Recordset
On Error Resume Next
If DBConnect = False Then Exit Function

    Set RS = RequestSQL(sql)
    If Not RS Is Nothing Then
        Do While Not RS.EOF
            RS.Delete
            RS.MoveNext
        Loop
        RS.Close
        Set RS = Nothing
    End If
End Function


Function GetSQLResult(sql As String) As Variant
    Dim RS As Recordset
    If DBConnect = False Then Exit Function
    
    On Error Resume Next
        Set RS = RequestSQL(sql)
        If Not RS Is Nothing Then
            GetSQLResult = RS(0)
        Else
            Beep
            GetSQLResult = vbNull
        End If
        Set RS = Nothing
End Function

