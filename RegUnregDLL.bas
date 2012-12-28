Attribute VB_Name = "RegUnregDLL"
Rem Автор Беляев Данила [outen@mail.ru]

Option Explicit
Private Declare Function CreateThread Lib "kernel32" (anyThread As Any, ByVal lngSize As Long, ByVal lngStart As Long, ByVal lngValue As Long, ByVal lngFlags As Long, lngThread As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal strName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal lngModule As Long, ByVal strName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal lngModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal lngHandle As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal lngHandle As Long, ByVal lngTime As Long) As Long

Public Enum RETURN_RESULT
    UnknownError = 0
    LoadLibraryError = 1
    NotActiveX = 2
    ThreadError = 3
    Success = 4
    Failure = 5
End Enum


Public Function Reg(ByVal strReg As String, ByVal lngLoad As Long) As RETURN_RESULT
Dim hLib As Long, hProc As Long, hTr As Long, tID As Long, succeed As Long, xc As Long

    hLib = LoadLibraryA(strReg & vbNullString)
    If hLib = 0 Then Reg = UnknownError: Exit Function
    
    hProc = GetProcAddress(hLib, IIf(lngLoad, "DllRegisterServer", "DllUnregisterServer"))
    If hProc = 0 Then Reg = NotActiveX: Exit Function
    
    hTr = CreateThread(ByVal 0&, 0&, ByVal hProc, ByVal 0, 0, tID)
    If hTr = 0 Then Reg = ThreadError: Exit Function
    
    succeed = (WaitForSingleObject(hTr, 10000) = 0)
    If succeed Then
        CloseHandle hTr
        Reg = Success
'        Exit Function
    Else
        GetExitCodeThread hTr, xc
        ExitThread xc
        Reg = Failure
'        exit function
    End If

    FreeLibrary hLib
End Function
