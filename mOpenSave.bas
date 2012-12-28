Attribute VB_Name = "Dialog"
Public Type FILTERINFO
    Extension As String
    name As String
    Path As String
    RegPath As String
End Type

Public ExportFilters() As FILTERINFO

Private Type OPENFILENAME 'Open & Save Dialog
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Public Const MAX_PATH = 260
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

'
' Открытие диалога выбора папки
'
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Enum BrowseType
    BrowseForFolders = &H1
    BrowseForComputers = &H1000
    BrowseForPrinters = &H2000
    BrowseForEverything = &H4000
End Enum

Public Enum FolderType
    CSIDL_BITBUCKET = 10
    CSIDL_CONTROLS = 3
    CSIDL_DESKTOP = 0
    CSIDL_DRIVES = 17
    CSIDL_FONTS = 20
    CSIDL_NETHOOD = 18
    CSIDL_NETWORK = 19
    CSIDL_PERSONAL = 5
    CSIDL_PRINTERS = 4
    CSIDL_PROGRAMS = 2
    CSIDL_RECENT = 8
    CSIDL_SENDTO = 9
    CSIDL_STARTMENU = 11
End Enum

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                      ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, _
                                                                ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, _
                                                                       ByVal nFolder As Long, _
                                                                       ListId As Long) As Long
                                                                       
Private Const BIF_RETURNONLYFSDIRS  As Long = 1
'Private Const CSIDL_DRIVES          As Long = &H11
Private Const WM_USER               As Long = &H400
Private Const MAX_PATH              As Long = 260 ' Is it a bad thing that I memorized this value?

'\\ message fromGetFileAttributes browser
Private Const BFFM_INITIALIZED     As Long = 1
Private Const BFFM_SELCHANGED      As Long = 2
Private Const BFFM_VALIDATEFAILEDA As Long = 3 '\\ lParam:szPath ret:1(cont),0(EndDialog)
Private Const BFFM_VALIDATEFAILEDW As Long = 4 '\\ lParam:wzPath ret:1(cont),0(EndDialog)
Private Const BFFM_IUNKNOWN        As Long = 5 '\\ provides IUnknown to client. lParam: IUnknown*

'\\ messages to browser
Private Const BFFM_SETSTATUSTEXTA   As Long = WM_USER + 100
Private Const BFFM_ENABLEOK         As Long = WM_USER + 101
Private Const BFFM_SETSELECTIONA    As Long = WM_USER + 102
Private Const BFFM_SETSELECTIONW    As Long = WM_USER + 103
Private Const BFFM_SETSTATUSTEXTW   As Long = WM_USER + 104
Private Const BFFM_SETOKTEXT        As Long = WM_USER + 105 '\\ Unicode only
Private Const BFFM_SETEXPANDED      As Long = WM_USER + 106 '\\ Unicode only
                                                                       
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
                                                                       
                                                                       
'\ Открытие диалога выбора папки

Public Function GetFileName(sTemplate As String, Optional ByVal sFilter As String, _
                            Optional ByVal initdir As String, Optional bOpen As Boolean = True, _
                            Optional hwnd As Long) As String
                            
On Error Resume Next

Dim OFN As OPENFILENAME
Dim ret As Long
Dim sExt As String

With OFN
    .lStructSize = Len(OFN)
    For i = 1 To Len(sFilter)
        If mID(sFilter, i, 1) = "|" Then
            Mid(sFilter, i, 1) = vbNullChar
        End If

    Next

    sFilter = sFilter & String$(2, 0)

    .hWndOwner = hwnd

    .lpstrFilter = sFilter
    .lpstrInitialDir = IIf(initdir = "", App.Path, initdir)
    .hInstance = App.hInstance

    .lpstrFile = sTemplate & String$(MAX_PATH - Len(sTemplate), 0)
    .lpstrFileTitle = String$(MAX_PATH, 0)
    .nMaxFile = MAX_PATH
     
End With

If bOpen Then
    ret = GetOpenFileName(OFN)
Else
    ret = GetSaveFileName(OFN)
End If

If ret Then GetFileName = LCase(TrimNull(OFN.lpstrFile))

 Dim Filtr() As String
 ReDim Filtr(0)
 Filtr = Split(sFilter, Chr(0))
 If OFN.nFilterIndex = 1 Then
     sExt = Right(Filtr(OFN.nFilterIndex), Len(Filtr(OFN.nFilterIndex)) - 1)
 Else
     sExt = Right(Filtr(OFN.nFilterIndex + 1), Len(Filtr(OFN.nFilterIndex + 1)) - 1)
 End If
 If GetFileName <> "" And Right(GetFileName, 4) <> sExt Then GetFileName = GetFileName + sExt
 
End Function


Public Function TrimNull(startstr As String) As String
    Dim pos As Integer
        pos = InStr(startstr, Chr$(0))
        If pos Then
            TrimNull = Left$(startstr, pos - 1)
            Exit Function
        End If

        TrimNull = startstr
End Function



Public Function BrowseFolders(hWndOwner As Long, sMessage As String, Browse As BrowseType, _
                              ByVal RootFolder As FolderType, Optional sPath As String) As String
Dim Nullpos As Integer
Dim lpIDList As Long
Dim Res As Long
Dim BInfo As BrowseInfo
Dim RootID As Long

'следующие вызовы функции сработали нормально
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_DESKTOP) '+весь компьютер
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_DRIVES) '+только устройства
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_NETHOOD) '+только сеть
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_PROGRAMS) '+папка Программы
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_STARTMENU) '+Главное меню
'
'результат действия следующих кодов вызвал недоумение...
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_BITBUCKET) '-корзина
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_CONTROLS) '-панель управления
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_FONTS) '-папка со шрифтами
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_NETWORK) '-NetHood
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_PERSONAL) '-Мои документы
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_PRINTERS) '-Принтеры
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_RECENT) '-RECENT
'MsgBox BrowseFolders(hWnd, "Select a Folder", BrowseForFolders, CSIDL_SENDTO) '-SENDTO

'        BInfo.hWndOwner = Screen.ActiveForm.hwnd

SHGetSpecialFolderLocation hWndOwner, RootFolder, RootID

BInfo.hWndOwner = hWndOwner
BInfo.lpszTitle = lstrcat(sMessage, "")
BInfo.ulFlags = Browse

BInfo.pIDLRoot = RootID
'        bi.pszDisplayName = VarPtr(B(0))
'        bi.lpszTitle = sDialogTitle
'        bi.ulFlags = BIF_RETURNONLYFSDIRS

If DirectoryExists(sPath) Then BInfo.lpfnCallback = PtrToFunction(AddressOf BFFCallback)
BInfo.lParam = StrPtr(sPath)

If RootID <> 0 Then BInfo.pIDLRoot = RootID
lpIDList = SHBrowseForFolder(BInfo)
If lpIDList <> 0 Then
    sPath = String$(MAX_PATH, 0)
    Res = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    Nullpos = InStr(sPath, vbNullChar)
    If Nullpos <> 0 Then
        sPath = Left(sPath, Nullpos - 1)
    End If

End If

BrowseFolders = sPath
End Function

Private Function DirectoryExists(ByVal sDirectory As String) As Long
  If LenB(sDirectory) Then
    If GetFileAttributes(sDirectory) >= vbNormal Then
      DirectoryExists = True
    End If
  End If
End Function

Private Function PtrToFunction(ByVal lFcnPtr As Long) As Long
  PtrToFunction = lFcnPtr
End Function

' typedef int (CALLBACK* BFFCALLBACK)(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData);
Public Function BFFCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal sData As String) As Long
  If uMsg = BFFM_INITIALIZED Then
    SendMessage hwnd, BFFM_SETSELECTIONA, True, ByVal sData
  End If
End Function
