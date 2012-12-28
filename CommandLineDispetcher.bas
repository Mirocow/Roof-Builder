Attribute VB_Name = "TwoClickFile"
Type WNDCLASSEX
    cbSize As Long
    Style As Long
    lpfnWndProc As Long
    cbClsextra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type

'Type POINTAPI
'        x As Long
'        y As Long
'End Type
Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const MEM_DECOMMIT = &H4000
Public Const MEM_COMMIT = &H1000&
Public Const PAGE_READWRITE = 4&
Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                                    ByVal bInheritHandle As Long, _
                                                    ByVal dwProcessId As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, _
                                                           ByVal lpBaseAddress As Long, _
                                                           lpBuffer As Any, _
                                                           ByVal nSize As Long, _
                                                           lpNumberOfBytesWritten As Long) As Long

Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, _
                                                       ByVal lpAddress As Long, _
                                                       ByVal dwSize As Long, _
                                                       ByVal flAllocationType As Long, _
                                                       ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, _
                                                     ByVal dwSize As Long, _
                                                     ByVal dwFreeType As Long) As Long

Declare Function RtlMoveMemory Lib "ntdll" (ByVal dst As Long, _
 _
                                            ByVal src As Long, _
 _
                                            ByVal size As Long) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, _
                                                               lpdwProcessId As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                     ByVal lpWindowName As String) As Long

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, _
 _
                                                                ByVal wMsg As Long, _
 _
                                                                ByVal wParam As Long, _
 _
                                                                ByVal lParam As Long) As Long

Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, _
 _
                                                                    ByVal wMsg As Long, _
 _
                                                                    ByVal wParam As Long, _
 _
                                                                    ByVal lParam As Long) As Long

Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, _
 _
                                                              ByVal hWnd As Long, _
 _
                                                              ByVal wMsgFilterMin As Long, _
 _
                                                              ByVal wMsgFilterMax As Long) As Long

Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long

Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long

Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)

Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
 _
                                                                      ByVal lpClassName As String, _
 _
                                                                      ByVal lpWindowName As String, _
 _
                                                                      ByVal dwStyle As Long, _
 _
                                                                      ByVal x As Long, _
 _
                                                                      ByVal y As Long, _
 _
                                                                      ByVal nWidth As Long, _
 _
                                                                      ByVal nHeight As Long, _
 _
                                                                      ByVal hWndParent As Long, _
 _
                                                                      ByVal hMenu As Long, _
 _
                                                                      ByVal hInstance As Long, _
 _
                                                                      lpParam As Any) As Long

Const CLASSNAME = "Rascal_Traffic"
Const WINDOWNAME = "Traffic"
Const MY_MSG = 66646
Public tHwnd As Long
Public strCommand As String
Public strCommandInp(1023) As Byte

Private Const WM_DESTROY = &H2

Public Function FindTrafficWindow() As Boolean ' ищем окно и шлем сообщение
    Dim ret As Long
    Dim dwProcId As Long
    Dim hProcess As Long
    Dim lpMem As Long
    
    tHwnd = FindWindow(CLASSNAME, WINDOWNAME) 'с обственно, поиск окна
    If tHwnd = 0 Then Exit Function
    GetWindowThreadProcessId tHwnd, dwProcId ' получаем Id процесса нашего окна
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, dwProcId) ' открываем процесс, получаем хэндл
    lpMem = VirtualAllocEx(hProcess, 0, 1024, MEM_COMMIT, PAGE_READWRITE) ' выделяем память под строку
    
    Dim S() As Byte
    S = StrConv(strCommand, vbFromUnicode)
    If Len(strCommand) = 0 Then Exit Function
    WriteProcessMemory hProcess, lpMem, ByVal VarPtr(S(0)), Len(strCommand), ret 'пишем строку в память процесса
    PostMessage tHwnd, 66646, lpMem, 0 ' посылаем сообщение, в параметре - указатель на область памяти
    CloseHandle hProcess
    FindTrafficWindow = True
End Function


Public Function CreateTrafficWindow() ' тут мы создаем окно
    Dim wc As WNDCLASSEX
    Dim message As Msg
    
        wc.cbSize = Len(wc)
        wc.Style = 0
        wc.lpfnWndProc = GetFuncPtr(AddressOf WindowProc)
        wc.cbClsextra = 0&
        wc.cbWndExtra = 0&
        wc.hInstance = App.hInstance
        wc.hIcon = 0
        wc.hCursor = 0
        wc.hbrBackground = 0
        wc.lpszMenuName = 0&
        wc.lpszClassName = CLASSNAME
        wc.hIconSm = 0
    
        RegisterClassEx wc
    
        tHwnd = CreateWindowEx(0&, _
CLASSNAME, _
WINDOWNAME, _
0, _
0, _
0, _
0, _
0, _
0&, _
0&, _
App.hInstance, _
0&)

End Function


Public Function WindowProc(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case message
    Case WM_DESTROY
        PostQuitMessage 0&
        Exit Function
    Case 66646 ' получили сообщение, копируем строку в переменную, обрабатываем
        RtlMoveMemory VarPtr(strCommandInp(0)), wParam, 1024
        VirtualFree wParam, 1024, MEM_DECOMMIT
        '  MsgBox StrConv(strCommandInp, vbUnicode)
        SetForegroundWindow (OfficeStart.hWnd)
        Dim fName As String
        fName = Replace(StrConv(strCommandInp, vbUnicode), Chr(0), "")
        If fName <> "" And ProjectsDir & FileName <> fName Then
            OfficeStart.OpenFilePreload fName
        End If
    Case Else
    WindowProc = DefWindowProc(hWnd, message, wParam, lParam)
End Select
End Function

Function GetFuncPtr(ByVal lngFnPtr As Long) As Long
    GetFuncPtr = lngFnPtr
End Function

