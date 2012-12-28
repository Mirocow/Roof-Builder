Attribute VB_Name = "vbs"
Option Explicit

Private sc As ScriptControl
'Private WithEvents ws As UniSock

Public Sub InitSC()
Set sc = New ScriptControl
sc.Language = "VBScript"
sc.AllowUI = True
sc.UseSafeSubset = True

'Dim RS As Recordset
'RegisterObjectSC "RS", RS
RegisterObjectSC "App", App
End Sub

Public Sub UnloadSC()
Set sc = Nothing
End Sub

Public Function RegisterObjectSC(name As String, Obj As Object)
sc.AddObject name, Obj
End Function

Public Function RegisterModuleSC(name As String, data As String) As Boolean
On Error GoTo ERR
sc.Reset
sc.Modules.Add name
sc.Modules(name).AddCode data
RegisterModuleSC = True
Exit Function
ERR:
RegisterModuleSC = False
End Function

Public Function RunFunctionsSC(mName As String, fName As String)
sc.Modules(mName).Run fName
End Function

'Public Function RunFuncSC(Name As String, data As String)
'sc.AddCode data
'sc.Run Name
'End Function

'Public Function RegisterSC(name As String, data As String)
'sc.Modules.Add
'End Function

'Public Function AddObjectSC(name As String, obj As Object)
'sc.AddObject name, obj
'End Function

'Public Function RunSC(CallName As String, args)
'On Error GoTo ERR
'
''script_1.ExecuteStatement (TrueTrim(txt_script.Text))
''    script_1.Modules
'
''sc.AddObject "clsFileAccess", fa
''sc.AddObject "App", App
'sc.AddCode TrueTrim(data)
''sc.Run "Main"
'
'ERR:
'End Function

