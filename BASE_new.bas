Attribute VB_Name = "BASE"
Dim db As DAO.Database 'ќбъ€вл€ем базу данных
Dim rs As DAO.Recordset 'ќбъ€вл€ем рекордсет

Function ConnectTo(ByVal strConn As String) As Boolean
On Error GoTo ErrorHandler
    If Right(FileName, 3) = "mdb" Then
        Set db = DAO.OpenDatabase(strConn)
    End If

ConnectTo = True
Exit Function

ErrorHandler:
End Function

Function DisConnect() As Boolean
On Error GoTo ERRDB
    rs.Close
    db.Close
    
    Set rs = Nothing
    Set db = Nothing

DisConnect = True
Exit Function
ERRDB:
STRERROR = STRERROR & "I can do it: " & sSQL & vbCrLf & " ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
End Function

Function SendSQLRequest(sSQL As String)
On Error GoTo ERRDB

Set rs = db.OpenRecordset(sSQL)

'rs.MoveLast
'ReDim data(rs.RecordCount, rs.Fields.count)
'
'With rs
'    .MoveFirst 'ѕеремещаемс€ к первой записи
'    Do While Not .EOF '¬ыполн€ть пока есть записи
'
'    For i = 0 To rs.Fields.count - 1
'    data(rs.AbsolutePosition, i) = Trim$(.Fields(i))
'    Next
'
'    .MoveNext 'ѕеремещаемс€ к следующей записи
'    Loop
'End With

'rs.Close
'db.Close
'
'Set rs = Nothing
'Set db = Nothing

'SendSQLRequest = data

SendSQLRequest = rs
Exit Function
ERRDB:
SendSQLRequest = False
STRERROR = STRERROR & "I can do it: " & sSQL & vbCrLf & " ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
End Function

Function Add_to_Data_DB(FileName As String, sSQL As String, ByRef data, row As Boolean, Optional dontadd As Boolean) As Boolean
Dim db As DAO.Database 'ќбъ€вл€ем базу данных
Dim rs As DAO.Recordset 'ќбъ€вл€ем рекордсет
Dim i As Integer
'Dim sSQL As String 'ѕеременна€, где будет размещЄн SQL запрос
On Error GoTo ERRDB

Set db = DAO.OpenDatabase(FileName)
sSQL = "SELECT * FROM " & sSQL & ";"
Set rs = db.OpenRecordset(sSQL)

If dontadd = True Then
rs.Edit
Else
'rs.Delete
rs.AddNew
End If

For i = 0 To UBound(data)
    If row Then
        If i <> 0 Then
        data(i) = Replace(data(i), vbCrLf, "&Vbcrlf&")
        rs.Fields(i) = data(i)
        End If
    Else
        rs.AddNew
        rs.Fields(1) = data(i)
        rs.Update
    End If
Next i

If row Then rs.Update

Add_to_Data_DB = True
Exit Function
ERRDB:
MsgBox "I can do it: " & sSQL & vbCrLf & " ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
Add_to_Data_DB = False
End Function

Function Dell_All_Data_in_DB(FileName As String, sSQL As String) As Boolean
Dim db As DAO.Database 'ќбъ€вл€ем базу данных
'Dim rs As DAO.Recordset 'ќбъ€вл€ем рекордсет
'Dim sSQL As String 'ѕеременна€, где будет размещЄн SQL запрос

Set db = DAO.OpenDatabase(FileName)

sSQL = "DELETE FROM " & sSQL & ";"
'db.
db.Execute sSQL
db.Close

'Set rs = Nothing
Set db = Nothing
End Function

Function Dell_Item_in_DB(FileName As String, sSQL As String, sItem As String) As Boolean
Dim db As DAO.Database 'ќбъ€вл€ем базу данных
Dim rs As DAO.Recordset 'ќбъ€вл€ем рекордсет
'Dim sSQL As String 'ѕеременна€, где будет размещЄн SQL запрос

Set db = DAO.OpenDatabase(FileName)

sSQL = "DELETE FROM " & sSQL & "WHERE " & sItem & ";"
db.Execute sSQL
db.Close

Set rs = Nothing
Set db = Nothing
End Function

'Function Update_Iten_Data_in_DB(FileName As String, sSQL As String, ByRef data) As Boolean
'Dim db As DAO.Database 'ќбъ€вл€ем базу данных
'Dim rs As DAO.Recordset 'ќбъ€вл€ем рекордсет
'Dim i As Integer
''Dim sSQL As String 'ѕеременна€, где будет размещЄн SQL запрос
'On Error GoTo ERRDB
'
'Set db = DAO.OpenDatabase(FileName)
'sSQL = "SELECT * FROM " & sSQL & ";"
'Set rs = db.OpenRecordset(sSQL)
'
'rs.AddNew
'
'For i = 1 To UBound(data)
'rs.Fields(i) = data(i)
'rs.Update
'Next i
'
'Add_to_Data_DB = True
'Exit Function
'ERRDB:
'Add_to_Data_DB = False
'End Function
