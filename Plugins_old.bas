Attribute VB_Name = "Plugins"
Option Explicit

Private Declare Function CreateThread Lib "kernel32" (anyThread As Any, _
                                                      ByVal lngSize As Long, _
                                                      ByVal lngStart As Long, _
                                                      ByVal lngValue As Long, _
                                                      ByVal lngFlags As Long, _
                                                      lngThread As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
                                                           lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal strName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal lngModule As Long, _
                                                        ByVal strName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal lngModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal lngHandle As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal lngHandle As Long, _
                                                             ByVal lngTime As Long) As Long

Private Enum RETURN_RESULT
    UnknownError = 0
    LoadLibraryError = 1
    NotActiveX = 2
    ThreadError = 3
    Success = 4
    Failure = 5
End Enum

Public Type P
    Dll As Object
    About As String
    Pname As String
    ERR As String
    Ver As String
    GUID As String
    ERRDescription As String
End Type

'Public clc As New CalcRoofList

Public Plgs() As P

Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Private Function RegLib(ByVal strReg As String, ByVal lngLoad As Long) As RETURN_RESULT
Dim hLib As Long, hProc As Long, hTr As Long, tID As Long, succeed As Long, xc As Long
On Error GoTo ERR

        hLib = LoadLibraryA(strReg & vbNullString)
        If hLib = 0 Then RegLib = UnknownError: Exit Function
    
        hProc = GetProcAddress(hLib, IIf(lngLoad, "DllRegisterServer", "DllUnregisterServer"))
        If hProc = 0 Then RegLib = NotActiveX: Exit Function
    
        hTr = CreateThread(ByVal 0&, 0&, ByVal hProc, ByVal 0, 0, tID)
        If hTr = 0 Then RegLib = ThreadError: Exit Function
    
        succeed = (WaitForSingleObject(hTr, 10000) = 0)
        If succeed Then
            CloseHandle hTr
            RegLib = Success
        Else
            GetExitCodeThread hTr, xc
            ExitThread xc
            RegLib = Failure
        End If

        FreeLibrary hLib
        
Exit Function
ERR:
RegLib = Failure
End Function


Sub GetPlugins()
    Dim i As Integer
    Dim plg As Object
    Dim Ver As String, Description As String

        On Error GoTo ERR 'Resume Next
            
    Dim strPlugin As String
        strPlugin = dir(App.Path & "\plugins\*.dll")

        'SetCurrentDirectory App.Path

        While strPlugin <> ""
 
            Set plg = LoadPlugin(App.Path & "\plugins\", strPlugin, "CalcRoofList")
    
            If Not plg Is Nothing Then
                ReDim Preserve Plgs(i)
                Set Plgs(i).Dll = plg
                Set plg = Nothing
                        
                Plgs(i).Pname = strPlugin
                Plgs(i).About = Plgs(i).Dll.About
                Plgs(i).ERRDescription = Plgs(i).Dll.ERRDescription
        
                If Plgs(i).ERR = "" Then
            
                    Ver = ""
                    Ver = Plgs(i).Dll.RBLibVer
                    If Ver <> "" Then Plgs(i).Ver = Ver
            
                    Plgs(i).Dll.Load_Library MAXSLOPELISTS + 10, 0, App.ProductName
                    Setup.Combo3.AddItem strPlugin
            
                Else
        
                    Setup.Combo3.AddItem strPlugin
            
                End If
        
                i = i + 1

            End If
    
            strPlugin = dir()
        Wend

        Exit Sub
ERR:
'        Plgs(i).ERR = ERR.Description
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.16." & ERR.Source & "]", ERR.Number, ERR.Description
        Resume Next
End Sub


Public Function LoadPlugin(dllpath As String, pluginname As String, ByVal strClassName As String) As Object
    On Error GoTo Handler
    Dim Plgs() As String
        'Dim plugin_name As String
 
        SetCurrentDirectory dllpath
        '   Plgs = Split(pluginname, "\")
        '   plugin_name = Plgs(UBound(Plgs))
        Set LoadPlugin = CreateObject(Left(pluginname, Len(pluginname) - 4) & "." & strClassName)

        Exit Function
Handler:
    Dim ans As RETURN_RESULT
    Dim TestPlugin As Object

        ' Регистрация библиотек
        ans = RegLib(pluginname, True)
    
        If ans = Success Then
            Set TestPlugin = CreateObject(Left(pluginname, Len(pluginname) - 4) & "." & strClassName)
            Set LoadPlugin = TestPlugin
        Else
            Set LoadPlugin = Nothing
        End If

        Set TestPlugin = Nothing
        Resume Next
        
End Function



Public Function UnregPlugins()
    Dim ans As RETURN_RESULT
    Dim i As Integer
        On Error Resume Next

        For i = 0 To UBound(Plgs)
            ans = RegLib(App.Path & "\Plugins\" & Plgs(i).Pname, False)
        Next

End Function


Public Function Path() As String
    If Right(App.Path, 1) = "\" Then Path = App.Path Else Path = App.Path & "\"
End Function


