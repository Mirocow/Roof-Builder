Attribute VB_Name = "PM"
#Const PMHard = 1
#Const PMDebug = 1

' NEW
'Private Declare Sub OutputDebugString Lib "kernel32.dll" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
'Private Declare Function ContinueDebugEvent Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long

Private Declare Sub MDFile Lib "aamd532" (ByVal f As String, ByVal r As String)
'Private Declare Function MapFileAndCheckSumA Lib "Imagehlp.dll" (ByVal FileName As String, HeaderSum As Long, CheckSum As Long) As Long
'Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

'Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Foo1 As Long, ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long

''Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
''Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
Public Declare Function CreateFileNS Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
''Public Declare Function WriteFileNO Lib "kernel32" Alias "WriteFile" (ByVal hfile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
''
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const EXCEPTION_ACCESS_VIOLATION = &HC0000005

'Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
'Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
'Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
'
'Private Const TH32CS_SNAPPROCESS As Long = 2&
'
'Private Type PROCESSENTRY32
'    dwSize As Long
'    cntUsage As Long
'    th32ProcessID As Long
'    th32DefaultHeapID As Long
'    th32ModuleID As Long
'    cntThreads As Long
'    th32ParentProcessID As Long
'    pcPriClassBase As Long
'    dwFlags As Long
'    szexeFile As String * 260
'End Type

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal strName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal lngModule As Long, ByVal strName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam%, ByVal lParam%) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal lngModule As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function PutMem2 Lib "msvbvm60" (ByVal pDst As Long, ByVal NewValue As Long) As Long
'Private Declare Function PutMem4 Lib "msvbvm60" (ByVal pDst As Long, ByVal NewValue As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

'Private Const GMEM_FIXED As Long = &H0
'Private Const MAX_PARAMS As Long = 10

Public Function SP() As Boolean
Dim n As Integer
Randomize
n = Int((254 * Rnd) + 1)

'#If PMHard = 1 Then
'If CStr(Project.Text1 <> "" Or Gl.CurrentFile <> "" And _
''(DebugersSoftwareLoaded(n) = n Or CallIsDebuggerPresent(n) = n Or SpyDebuggers(n) = n)) Then 'Or SpyDebuggers(n) <> n
'#Else
If Project.Text1 <> "" Or Gl.CurrentFile <> "" Then
'#End If
    
    If Gl.CurrentFile <> "" Then Project.Text1 = Gl.CurrentFile
    If Gl.Catalogue_files <> "" Then Project.Label11 = Gl.Catalogue_files
    
    OfficeStart.da.Enabled = True
    OfficeStart.menu_print_valinta.Enabled = True
    OfficeStart.menu_upd.Enabled = True
    
    OfficeStart.menu_save_as.Enabled = True
    OfficeStart.menu_save.Enabled = True
    
    OfficeStart.Toolbar1.Buttons(3).Enabled = True
'    OfficeStart.Toolbar1.Buttons(4).Enabled = True
    OfficeStart.Toolbar1.Buttons(9).Enabled = True
    
    OfficeStart.TabStrip1.Enabled = True
    OfficeStart.mOpWp.Enabled = True
    
    Project.Command4.Enabled = OfficeStart.TabStrip1.Enabled
    
    Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(1).Picture
    
    OfficeStart.StatusBar.Panels(1).Text = Processing & OptionDMM

Else

    N_Slope = 1: az = "a"
    
    OfficeStart.Toolbar1.Buttons(3).Enabled = False
'    OfficeStart.Toolbar1.Buttons(4).Enabled = False
    OfficeStart.Toolbar1.Buttons(9).Enabled = False
    
'    Project.Text1 = ""
    OfficeStart.menu_upd.Enabled = False
'    OfficeStart.StatusBar.Panels(2).Text = ""
    OfficeStart.StatusBar.Panels(3).Text = ""
'    OfficeStart.StatusBar.Panels(4).Text = ""
    OfficeStart.menu_save.Enabled = False
    OfficeStart.menu_save_as.Enabled = False
    OfficeStart.da.Enabled = False
    OfficeStart.menu_print_valinta.Enabled = False
'    OfficeStart.menu_xls.Enabled = False
    OfficeStart.mOpWp.Enabled = False
    OfficeStart.TabStrip1.Enabled = False
    Project.Command4 = False
    Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(3).Picture
    
End If
End Function

'Private Function CRC() As Boolean
'Open AppExe For Binary As #1
' Get #1, FPos - 7, FileCRC   'read the last 8 bytes of the file
'Close #1
'End Function

'Private Function ProcessDetection() As Boolean
'Dim myProcess As PROCESSENTRY32
'Dim mySnapshot As Long
'Dim Processes As String
'Processes = "filemon.exe regmon.exe PROCEXP.EXE"
'
'myProcess.dwSize = Len(myProcess)
'
''create snapshot
'mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
'
'xx = 0
''get first process
'ProcessFirst mySnapshot, myProcess
''Debug.Print Left(myProcess.szexeFile, InStr(1, myProcess.szexeFile, chr$(0)) - 1) ' set exe name
'If InStr(1, Processes, Left(myProcess.szexeFile, InStr(1, myProcess.szexeFile, chr$(0)) - 1), vbTextCompare) Then ProcessDetection = True: Exit Function
''ProcessID(xx) = myProcess.th32ProcessID ' set PID
'
''while there are more processes
'While ProcessNext(mySnapshot, myProcess)
'xx = xx + 1
''Debug.Print Left(myProcess.szexeFile, InStr(1, myProcess.szexeFile, chr$(0)) - 1) ' set exe name
'If InStr(1, Processes, Left(myProcess.szexeFile, InStr(1, myProcess.szexeFile, chr$(0)) - 1), vbTextCompare) Then ProcessDetection = True: Exit Function
''ProcessID(xx) = myProcess.th32ProcessID ' set PID
'Wend
'End Function

Private Function DebugersSoftwareLoaded(ByVal t As Integer) As Integer
Dim n As Integer
Dim Debugers() As String
On Error GoTo ERR
Randomize
Debugers = Split("\\.\SIWDEBUG, \\.\SIWVID, \\.\NTICE, \.\SuperBPMDev0, \\.\SICE, \\.\TRW, \\.\FILEVXD, \\.\FILEMON, \\.\REGVXD, \\.\REGMON", ", ", , vbTextCompare)
For n = 0 To UBound(Debugers)
If Not DebugersIsLoaded(Debugers(n)) Then DebugersSoftwareLoaded = t Else DebugersSoftwareLoaded = Int((99999999 * Rnd) + 9999999)
Next
Exit Function
ERR:
#If PMDebug Then
STRERROR = STRERROR & Time & ". ( PM_IL ) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
#End If
End Function

Private Function DebugersIsLoaded(ds As String) As Boolean
Dim hfile As Long, retval As Long
    hfile = CreateFileNS(ds, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hfile <> -1 Then
        ' Debugers is detected.
        retval = CloseHandle(hfile) ' Close the file handle
        RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
        DebugersIsLoaded = True
'        MsgBox ds, vbCritical
    Else
    ' Debugers is not found.
    DebugersIsLoaded = False
    End If
End Function


Public Function CallIsDebuggerPresent(ByVal t As Byte) As Byte
On Error GoTo ERR
Dim lb As Long
'TimerStart = Timer
Randomize
CallIsDebuggerPresent = CByte((254 * Rnd) + 1)
lb = LoadLibraryA(ByVal "kernel32")
If lb = 0 Then GoTo ERR
mlngAddress = GetProcAddress(lb, "IsDebuggerPresent")
If mlngAddress = 0 Then FreeLibrary lb: GoTo ERR
If CallWindowProc(mlngAddress, OfficeStart.hWnd, ByVal 0&, ByVal 0&, ByVal 0&) = 0 Then _
CallIsDebuggerPresent = t Else RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0: GoTo ERR  'Else MsgBox "IsDebuggerPresent", vbCritical
FreeLibrary lb
Exit Function
ERR:
#If PMDebug Then
STRERROR = STRERROR & Time & ". ( PM_CD ) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
#End If
'CallIsDebuggerPresent = t / 2
End Function


Public Function SpyDebuggers(ByVal t As Byte) As Byte
On Error GoTo ERR
Dim Debugers() As String, n As Integer
'TimerStart = Timer
Randomize
SpyDebuggers = Int((254 * Rnd) + 1)
Debugers = Split("thread, ACPU, Shadow, SND, , CPU, OLLYDBG, OllyDbg, TIdaWindow, 18467-41, HexWorks, OWL_Window, NMSCMW, UltraEdit-32, ACPUDUMP,  ACPUREG,  ACPUSTACK", ", ", , vbTextCompare)
For n = 0 To UBound(Debugers) - 1
If FindWindow(CStr(Debugers(n)), vbNullString) = 0 Then _
SpyDebuggers = t Else RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0: GoTo ERR  'MsgBox Debugers(n), vbCritical: Exit For
Next
Exit Function
ERR:
#If PMDebug Then
STRERROR = STRERROR & Time & ". ( PM_SD ) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
#End If
'SpyDebuggers = t / 2
End Function


''Check for processes and wipe from 200000 to N amount of bytes in steps of 48
''(to aggressively screw with the code)
'Public Sub HAMMERPROCESS(PID As Long, hammertop)
'    If Not InitProcess(PID) Then MsgBox "Failed shutdown"
'    Dim addr As Long
'    For P = 20000 To hammertop Step 48
'    addr = CLng(Val(Trim(P)))
'
'    Call WriteProcessMemory(myHandle, addr, "6", 1, L)
'    Next P
'End Sub
'
'Function InitProcess(PID As Long)
'pHandle = OpenProcess(&H1F0FFF, False, PID)
'If (pHandle = 0) Then
'    InitProcess = False
'Else
'    InitProcess = True
'End If
'End Function
'
''CFP ("OLLYDBG.exe"), 2000000
'Public Sub CFP(procname$, hammerrange)
'    For xx = 0 To 256
'    If LCase(procname$) = LCase(ProcessName$(xx)) Then HAMMERPROCESS CLng(ProcessID(xx)), hammerrange
'    Next xx
'End Sub


Function StrDecode(ByVal str As String) As String
On Error GoTo ERR
Dim tmp As String
Dim str1() As Byte
tmp = ""
str1 = StrConv(str, vbFromUnicode)
Dim flag1 As Byte
Dim CheckStrCode As Integer
For i = 0 To UBound(str1)
If str1(i) > 0 Then
'    flag1 = str1(i)
    str1(i) = str1(i) Xor 1
'    If flag1 Mod 2 = 1 Then _
'        If flag1 > str1(i) Then CheckStrCode = CheckStrCode + 1
'    If flag1 Mod 2 = 0 Then _
'        If flag1 < str1(i) Then CheckStrCode = CheckStrCode - 1
End If
Next
StrDecode = StrConv(str1, vbUnicode)
Exit Function
ERR:
#If PMDebug Then
STRERROR = STRERROR & Time & ". ( PM_STD:" & str & " ) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
#End If
StrDecode = ""
End Function


Public Function ChekLicence(ByVal ProductID As String, ByVal nlicence As String) As String
On Error GoTo ERR
Dim LicRnd As Integer
'DebugersSoftwareLoaded (Left(nlicence, 1))
'Dim Tmpl As String
'Dim lRet&, crcH&, crcC&
'Randomize
'LicRnd = Int((260 * Rnd()) + 260)
'Tmpl = String$(DebugersSoftwareLoaded(LicRnd), " ")
'Файл нашей программы
'GetModuleFileName 0, Tmpl, Len(Tmpl)
'Узнаем CRC
'lRet = MapFileAndCheckSumA(App.Path & "\roof.exe", 0, crcC)
'Если фактический crcH <> crcC подсчитанному
'при компиляции тогда завершаем прогу
templic = CallByName(OfficeStart, PM.GetFunctionName(ProductID, PM.StrDecode(Gl.Regnumber))(Pig) & CBool(InStr(MD5, OfficeStart.Picture2.ToolTipText)), VbMethod, nlicence)
'Tmpl = ProductID
ChekLicence = templic
Exit Function
ERR:
#If PMDebug Then
STRERROR = STRERROR & Time & ". ( PM_CL ) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
#End If
'ChekLicence = templic
End Function

Public Function MD5File(f As String) As String
On Error GoTo ERR
' compute MD5 digest on o given file, returning the result
'If dir(f) = "" Then Exit Function
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
Exit Function
ERR:
MsgBox ERR.Description, vbCritical
End Function

Function GetFunctionName(s As String, s1 As String) As String()
Dim i As Integer
Dim lenstr As Integer
Dim templic As String
On Error GoTo ERR
lenstr = Len(s1)
templic = ""
For i = Pig To lenstr Step 3
templic = templic & Chr$(Mid(s, i, 3) Xor Mid(s1, i, 3))
Next
GetFunctionName = Split(templic, ",")
Exit Function
ERR:
#If PMDebug Then
STRERROR = STRERROR & Time & ". ( PM_GFN " & Pig & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
#End If
End Function

'Function FExecuteCode(stCode As String, Optional fCheckOnly As Boolean) As Boolean
'FExecuteCode = EbExecuteLine(StrPtr(stCode), 0&, 0&, Abs(fCheckOnly)) = 0
'End Function

'Public Sub CTV(appid$)
''SICE
''NTICE
''SIWDEBUG
''SIWVID
''Check threats vxd
'    If CF("\\.\" & appid$, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0) <> -1 Then
'    retval = CloseHandle(hfile) ' Close the file handle
''    End
'    End If
'End Sub

'Public Function CallFunction(ByVal FuncPointer As Long, _
'ParamArray p()) As Long
'  Dim i As Long
'  Dim hGlobal As Long, hGlobalOffset As Long
'
'  'Учтём совпадение числа параметров:
'  If UBound(p) - LBound(p) + 1 = 4 Then
'    CallFunction = CallWindowProc(FuncPointer, _
' CLng(p(0)), CLng(p(1)), CLng(p(2)), CLng(p(3)))
'  Else
'    hGlobal = GlobalAlloc(GMEM_FIXED, 5 * MAX_PARAMS + _
'5 + 3 + 1)   'Заполняем всё подряд, ZEROINIT не нуно.
'    If hGlobal = 0 Then ERR.Raise 7 'insuff. memory
'    hGlobalOffset = hGlobal
'
'    For i = LBound(p) To UBound(p)
''если параметров нет, то ubound<lbound, и цикл не выполнится вообще
'      PutMem2 hGlobalOffset, &H68 'asmPUSH_imm32
'      hGlobalOffset = hGlobalOffset + 1
'      PutMem4 hGlobalOffset, CLng(p(i))
'      hGlobalOffset = hGlobalOffset + 4
'    Next
'
'    'Добавляем вызов функции
'    PutMem2 hGlobalOffset, &HE8 ' asmCALL_rel32
'    hGlobalOffset = hGlobalOffset + 1
'    PutMem4 hGlobalOffset, FuncPointer - hGlobalOffset - 4
'    hGlobalOffset = hGlobalOffset + 4
'
'    PutMem4 hGlobalOffset, &H10C2&        'ret 0x0010
'
'    CallFunction = CallWindowProc(hGlobal, 0, 0, 0, 0)
'
'    GlobalFree hGlobal
'  End If
'End Function

'Private Declare Function WNetGetUserA Lib "mpr" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
'
'Private Function GetUser() As String
'   Dim sUserNameBuff As String * 255
'   sUserNameBuff = Space(255)
'   Call WNetGetUserA(vbNullString, sUserNameBuff, 255&)
'   GetUser = Left$(sUserNameBuff, InStr(sUserNameBuff, vbNullChar) - 1)
'End Function
'
'Private Sub Form_Load()
'   Me.Caption = GetUser
'End Sub


'crc уже прошита в экзешник,
'данная функция читает его из экзешника
'Public Function GetPEHeadCrc(ByVal Path As String) As Long
'  Dim tmp As String
'  Dim lRet&, crcH&
'  lRet = MapFileAndCheckSumA(Path, crcH, 0)
'  GetPEHeadCrc = crcH
'End Function
