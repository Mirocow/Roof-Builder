Attribute VB_Name = "BASE"
Function TestConnection(ByVal strConn As String) As String
Dim db As DAO.Database 'ќбъ€вл€ем базу данных
On Error GoTo ErrorHandler

Set db = DAO.OpenDatabase(strConn)
db.Close
Set db = Nothing


Exit Function

ErrorHandler:
Set db = Nothing
TestConnection = ERR.Description
End Function

Function SendSQLRequest(FileName As String, sSQL As String) As Variant()
On Error GoTo ERRDB

If Right(FileName, 3) = "mdb" Then

Dim data()
Dim i As Integer

Dim db As DAO.Database 'ќбъ€вл€ем базу данных
Dim rs As DAO.Recordset 'ќбъ€вл€ем рекордсет

Set db = DAO.OpenDatabase(FileName)
Set rs = db.OpenRecordset(sSQL)

rs.MoveLast
ReDim data(rs.RecordCount, rs.Fields.count)

With rs
    .MoveFirst 'ѕеремещаемс€ к первой записи
    Do While Not .EOF '¬ыполн€ть пока есть записи
        For i = 0 To rs.Fields.count - 1
        data(rs.AbsolutePosition, i) = Trim$(.Fields(i))
        Next
    .MoveNext 'ѕеремещаемс€ к следующей записи
    Loop
End With

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

SendSQLRequest = data
End If
Exit Function

ERRDB:
STRERROR = STRERROR & "I can do it: " & sSQL & vbCrLf & " ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
Exit Function
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
