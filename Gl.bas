Attribute VB_Name = "Gl"
Option Explicit

Public Const FILEVER = 10

Public TemporaryFileName As String
Public Positions As Integer
Public CurentPosition As Integer

' LIC DATA
Public iData() As String

Public Const PI = 3.14159265358979

Public TimerStart As Single
Public Ver As String ' App.ver
Public LNC As Integer ' Выбранная библиотека для рассчета
Public bDebug As Boolean

' Locale
Public lng As Lng_class
Public CurrentLocale As String

'--------------------------------------
Public PV As String * 6

Public IsLic As Boolean ' Переменная косвенно сигнализирующая о DEMO MODE (от лохов)

Public UserName As String
Public Uname As String
Public UserCreatProject As String

Public CountRicentlyFiles As Integer ' Максимальное количество ячеек
Public RecentlyFiles() As String ' Ранее открытые файлы

'Public Leftbuttons(10) As Integer ' Запоминание интерфейса с левой стороны

'
' Базовые установки
'
Public m2 As String * 5

Public MAINMAXSLOPELINE As Integer ' количество допустимых линий на главном рисунке 200
Public MAXSLOPELINE As Integer ' количество допустимых точек 20
Public MAXSLOPES As Integer ' количество допустимым расчитываемых поверхностей 52
Public MAXSLOPELISTS As Integer ' количество допустимым листов на скате

'Public MAXSL As Integer

Public Pic As Object ' Вывод флага работы в OfficeStart
'Public STRERR As String

'# FILE SYSTEM
Public isSave As Boolean
Public isChange As Boolean
Public ConfigDir As String

Public ProjectsDir As String
Public TempDir As String
Public CurrentProjectDir As String

Public FileName As String ' mdb
Public FileNameExtension As String
Public CurrentFile As String ' Текущий открытый файл
Public file_name_size As Integer ' Максимальная длина имени файда

'# MATERIAL
Public Profil_Name As String
Public Factory_Name As String

Public width1 As String
Public cover As String
Public ColorRoof As String

Public FlagPrinter As Integer

Public PrintFont As String
Public PrintFontSize As Single

Public WindowsFont As String
Public WindowsFontSize As Single

Public NumCopies As Integer
'Public PX As Integer
'Public Py As Integer
'Public drawwidth As Integer

'# DRAW - Режимы работы программы
Public OptionDMM As String
'Public Draw As String

'# Project
Public PrjDescrib As String

'# PICTURE WIDTH HEIGHT
Public ScaleHeight_Main As Single
Public ScaleWidth_Main As Single
Public ScaleLeft_Main As Single
Public ScaleTop_Main As Single

'# MAIN
Public KolvoScatov As Integer
' Переменные окна RoofPic
Public LapeName As Integer

' количество точек
Public MainCountOfPoints As Integer
Public Main_Points_X() As Single
Public Main_Points_Y() As Single

' количество расставленых обозначений
Public MainCountOfLines As Integer
Public Label_X() As Single ' координаты
Public Label_Y() As Single

'
Public Points_m_A() As Integer
Public Points_m_B() As Integer
' Описание
Public MainDescrib As String

'#
'# LAPE
'#

Public Const RatioW = 1.51 '1 ' 1.51
Public Const RatioH = 1.2 '2.5 '1.2 ' 1.2

Public N_Slope As Integer ' текущий скат,рисунок,расчет итд

' Выбранная точка при редактировании
Public P_A As Integer
Public P_B As Integer

'
' SLP - SLOPE PROPERTY
'
Public Type SlopeProperty
    CountOfLines As Integer ' Количество линий на чертеже
    CountOfPoints As Integer ' Количество точек на чертеже
    ' Разменры SCALE для LAPEPIC
    ScaleLeftS As Single
    ScaleWidthS As Single
    ScaleTopS As Single
    ScaleHeightS As Single
    ' Данные расчета
    Pn_Red_lines As Integer ' номер точки через которую проходит красная линия
    Pn_StartLC As Integer ' номер точки через которую проходит  линия начала разлиновки
    PX_StartLC As Single ' координаты по X линии начала разлиновки
    Factory_Name As String * 50 ' имя производителя
    ProfilName As String * 50 ' имя профиля для ската
    CountSheets As Integer ' количество полос
    ListLength As Single ' Длина листов ската  (Пагонная длина)
    Sf As Double ' Площадь плоскости
    Sw As Double ' Площадь покрытия по рабочей ?
    Describ As String * 255
    
End Type
Public SlP() As SlopeProperty ' Массив скатов MAXSLOPES

Public az As String * 1

' Линии прорисовки скатов
Public Lape_Lines() As Integer ' Свойства линий
Public Lape_Points_X() As Single ' Координаты линий по X
Public Lape_Points_Y() As Single ' Координаты линий по Y

' Для механизма поворота фигуры
Public SaveLape_Points_X() As Single ' Координаты линий по X
Public SaveLape_Points_Y() As Single ' Координаты линий по Y

'Public Const GUIDINGS = 20
'
'Public Type tGuid
'    X As Single
'    Y As Single
'End Type
'
'Public Lape_Guidings() As tGuid ' Массив с координатами вспомогательных линий

' Раскрой (ЛИСТЫ)
Public List_Properties_Length() As Single ' длина полосы
Public List_Properties_PX() As Single  ' Координаты по X (НАЧАЛО ПРОРИСОВКИ)
Public List_Properties_PY() As Single  ' Координаты по Y (НАЧАЛО ПРОРИСОВКИ)

Public SelectLines() As Integer

'Public SaveList_Properties_PX() As Single  ' Координаты по X (НАЧАЛО ПРОРИСОВКИ)
'Public SaveList_Properties_PY() As Single  ' Координаты по Y (НАЧАЛО ПРОРИСОВКИ)

'
'# ALL
'
Public FlagDraw As Integer
'Public Txt_to_Lape_or_mainp As String
'Public FindPoints As Integer

Public FEXIT As Boolean ' Флаг нажатия кнопки выход

' Выделеные листы
Public SelectLists As cCollection

' Невыполнимые длины
Public Type WrongL
    MIN As Integer
    MAX As Integer
End Type
Public WrongLs() As WrongL

'Public Help_Points(MAXSLOPES, MAXSLOPES, 1) As Single
'Public help_N_points As Integer
'Public help_Lape_Lines(MAXSLOPES, MAXSLOPES, 1) As Single
'Public help_All_lines As Integer

'Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
'Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const HKEY_CURRENT_USER = &H80000001
'Private Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const REG_SZ = 1
'Private Const KEY_QUERY_VALUE = &H1
'Private Const KEY_SET_VALUE = &H2
'Private Const KEY_CREATE_SUB_KEY = &H4
'Private Const KEY_ENUMERATE_SUB_KEYS = &H8
'Private Const KEY_NOTIFY = &H10
'Private Const KEY_CREATE_LINK = &H20
'Private Const SYNCHRONIZE = &H100000
'Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const MF_BYPOSITION = &H400&

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
'Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long
                                                 
Public Declare Function GetInputState Lib "user32" () As Long

'Public ReloadRB As Boolean

Public Function mGetComputerName() As String
    Dim compuname As String
        On Error GoTo ERR
        compuname = String$(255, " ")
        GetComputerName compuname, 255
        mGetComputerName = Trim_Null(compuname)
        Exit Function
ERR:
        mGetComputerName = ""
End Function

Public Function mGetUserName() As String
    Dim Uname As String
        On Error GoTo ERR
        Uname = String$(255, " ")
        GetUserName Uname, 255
        mGetUserName = Trim_Null(Uname)
        Exit Function
ERR:
        mGetUserName = ""
End Function


Public Function Trim_Null(ByVal t As String) As String
    Trim_Null = Left$(t, InStr(t, Chr$(0)) - 1)
End Function


Public Sub AssociateFile(ftype As String, fExtension As String, dtype As String, tDefault As String, tDescription As String, aPath As String, nIcon As Integer)
'    Dim RetVal As Long, kWnd As Long, dCommand As String, shell As String
'    dCommand = "\shell\open\command"
'    shell = "\shell\open\"

    '    retval = RegCreateKey(HKEY_CLASSES_ROOT, ftype, kWnd)
    '    retval = RegCloseKey(kWnd)
    '    retval = RegCreateKey(HKEY_CLASSES_ROOT, ftype & shell, kWnd)
    '    retval = RegCloseKey(kWnd)
    '    retval = RegCreateKey(HKEY_CLASSES_ROOT, ftype & shell & dType, kWnd)
    '    retval = RegCloseKey(kWnd)
    '    retval = RegCreateKey(HKEY_CLASSES_ROOT, ftype & shell & dType & "\command", kWnd)
    '    retval = RegSetValue(HKEY_CLASSES_ROOT, ftype & shell & dType & "\command", REG_SZ, aPath & " %1", Len(aPath) + 3)
    '    retval = RegCloseKey(kWnd)

'    RetVal = RegCreateKey(HKEY_CLASSES_ROOT, fExtension, kWnd)
'    RetVal = RegSetValue(HKEY_CLASSES_ROOT, fExtension, REG_SZ, dtype, Len(dtype))
'    RetVal = RegCloseKey(kWnd)
'    RetVal = RegCreateKey(HKEY_CLASSES_ROOT, dtype & dCommand, kWnd)
'    RetVal = RegSetValue(HKEY_CLASSES_ROOT, dtype, REG_SZ, tDescription, Len(tDescription))
'    RetVal = RegSetValue(HKEY_CLASSES_ROOT, dtype & shell, REG_SZ, tDefault, Len(tDefault))
'    RetVal = RegSetValue(HKEY_CLASSES_ROOT, dtype & dCommand, REG_SZ, aPath & " %1", Len(aPath) + 3)
'    RetVal = RegCloseKey(kWnd)
'    RetVal = RegCreateKey(HKEY_CLASSES_ROOT, dtype & "\DefaultIcon", kWnd)
'    RetVal = RegSetValue(HKEY_CLASSES_ROOT, dtype & "\DefaultIcon", REG_SZ, aPath & "," & nIcon, Len(aPath) + Len(CStr(nIcon)) + 1)
'    RetVal = RegCloseKey(kWnd)
End Sub


Public Sub DellAssociateFile(dtype As String, fExtension As String)
'    Dim RetVal As Long, kWnd As Long, dCommand As String
'    dCommand = "\shell\open\command"
'    RetVal = RegDeleteKey(HKEY_CLASSES_ROOT, dtype & "\DefaultIcon")
'    RetVal = RegDeleteKey(HKEY_CLASSES_ROOT, dtype & dCommand)
'    RetVal = RegDeleteKey(HKEY_CLASSES_ROOT, fExtension)
End Sub

' Cint
Public Function mCint(ByRef v) As Integer
On Error GoTo ERR
If IsNull(v) Then mCint = 0: Exit Function
mCint = v / 1
Exit Function
ERR:
mCint = CInt(v)
End Function

' Взять градусы
Public Function GetGRD(ByRef pA As POINT, ByRef pB As POINT) As Integer
Dim a As Single
Dim c As Single
On Error Resume Next
c = pB.X - pA.X
a = pB.Y - pA.Y
GetGRD = Abs(Atn(a / c)) * 180 / PI
' FIX
If pB.Y > pA.Y And pB.X > pA.X Then ' 0-90
    Exit Function
ElseIf c = 0 And pB.Y > pA.Y Then ' 90
    GetGRD = 90
    Exit Function
ElseIf pB.Y > pA.Y And pA.X > pB.X Then ' 90-180
    GetGRD = 180 - GetGRD
    Exit Function
ElseIf a = 0 And pA.X > pB.X Then ' 180
    GetGRD = 180
    Exit Function
ElseIf pA.Y > pB.Y And pA.X > pB.X Then ' 180-270
    GetGRD = 180 + GetGRD
    Exit Function
ElseIf c = 0 And pB.Y < pA.Y Then ' 270
    GetGRD = 270
    Exit Function
ElseIf pA.Y > pB.Y And pB.X > pA.X Then ' 270-360
    GetGRD = 360 - GetGRD
    Exit Function
End If
' Вычисление градусов
'GetGRD = Atn(a / c) * 180 / PI
'If RetReal Then Exit Function
'If X <= 0 And y >= 0 Then ' 90-180
'GetGRD = 180 - GetGRD * -1
'ElseIf X >= 0 And y >= 0 Then ' 0-90
'GetGRD = GetGRD
'ElseIf X >= 0 And y <= 0 Then ' 270-360
'GetGRD = 360 - GetGRD * -1
'ElseIf X <= 0 And y <= 0 Then ' 180-270
'GetGRD = GetGRD + 180
'End If
'Dim b As Single
'b = Sqr(c ^ 2 + a ^ 2)
'If b = a Then
'Else
'GetGRD = ArcCos((b ^ 2 + c ^ 2 - a ^ 2) / (2 * b * c)) * 180 / PI
'End If
'Exit Function
'ERR:
'GetGRD = 0
End Function

' степень 2
'Public Function Raise2(ByRef v)
'Raise2 = v * v
'End Function

Public Function IsLoadForm(ByRef name As String) As Boolean
Dim frmloaded As Form
IsLoadForm = False
If name = "" Then Exit Function
For Each frmloaded In Forms
If frmloaded.name = name Then
IsLoadForm = True
Exit Function
End If
Next
End Function

Sub loadFactory(FactoryName As String)
Dim i As Integer
    If Gl.FileNameExtension = ".rfd" Then
        For i = 1 To MAXSLOPES - 1
            SlP(i).Factory_Name = FactoryName
        Next
    End If
End Sub


Public Sub loadProfil(ProfilName As String)
Dim i As Integer
    If Gl.FileNameExtension = ".rfd" Then
        For i = 1 To MAXSLOPES - 1
            SlP(i).ProfilName = ProfilName
        Next
    End If
End Sub

Public Function SetChange(vVal As Boolean)
    ' Устанавливаем флаг изменения данных
    isChange = vVal
    isSave = vVal
End Function

