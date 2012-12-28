VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

    'MsgBox "This will test the registry functions." +  vbNewLine  + _
'       "Start Regedit in order to see the changes."
    '
    'A.CreateKey HKEY_LOCAL_MACHINE, "Software\Test\å"
    'MsgBox "Created key HKEY_LOCAL_MACHINE\Software\Test"
    '
    'A.SetStringValue HKEY_LOCAL_MACHINE, "Software\Test\å", "Testing String Values", "Test OK"
    'MsgBox "Created string key Testing String Values with the value Test OK"
    '
    'A.SetLongValue HKEY_LOCAL_MACHINE, "Software\Test\å", "Testing Long Values", 999
    'MsgBox "Created dword key Testing Long Values with the value 999"
    '
    'A.SetBooleanValue HKEY_LOCAL_MACHINE, "Software\Test\å", "Testing Boolean Values", True
    'MsgBox "Created dword key Testing Boolean Values with the value True (1)"
    '
    'MsgBox A.GetStringValue(HKEY_LOCAL_MACHINE, "Software\Test\å", "Testing String Values", "error"), , "String Value retreived from registry"
    'MsgBox A.GetLongValue(HKEY_LOCAL_MACHINE, "Software\Test\å", "Testing Long Values", -1), , "Long Value retreived from registry"
    'MsgBox A.GetBooleanValue(HKEY_LOCAL_MACHINE, "Software\Test\å", "Testing Boolean Values", False), , "Boolean Value retreived from registry"
    '
    'A.DeleteKey HKEY_LOCAL_MACHINE, "Software\Test"
    'MsgBox "Key Software\Test deleted."
    '
    'A.CreateKey HKEY_LOCAL_MACHINE, "Software\Test"
    'MsgBox "Empty key HKEY_LOCAL_MACHINE\Software\Test created"
    '
    'A.DeleteEmptyKey HKEY_LOCAL_MACHINE, "Software\Test"
    'MsgBox "Empty key HKEY_LOCAL_MACHINE\Software\Test deleted"


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
                                                                                    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, _
                                                                                    ByVal samDesired As Long, lpSecurityAttributes As Any, _
                                                                                    phkResult As Long, _
                                                                                    lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, lpData As Any, _
                                                                                  ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

Private Const REG_NONE = 0                       ' No value type
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0&

Public Enum REG_HKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Public Enum REG_TYPE
    NONE = 0                       ' No value type
    SZ = 1                         ' Unicode nul terminated string
    EXPAND_SZ = 2                  ' Unicode nul terminated string
    Binary = 3                     ' Free form binary
    DWORD = 4                      ' 32-bit number
    DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
    DWORD_BIG_ENDIAN = 5           ' 32-bit number
    LINK = 6                       ' Symbolic Link (unicode)
    MULTI_SZ = 7                   ' Multiple Unicode strings
End Enum

Private Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If

    StripNulls = OriginalStr
End Function


Public Function CreateKey(hKey As REG_HKEY, _
                          KeyName As String) As Boolean
    Dim RegKeyHandle As Long
    Dim ret As Long
        ret = RegCreateKeyEx(CLng(hKey), KeyName, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, RegKeyHandle, ByVal 0&)
        If ret <> ERROR_SUCCESS Then
            RegCloseKey RegKeyHandle
            CreateKey = False
        Else
            RegCloseKey RegKeyHandle
            CreateKey = True
        End If

End Function


Public Function DeleteEmptyKey(hKey As REG_HKEY, _
                               KeyName As String) As Boolean
    Dim ret As Long
        ret = RegDeleteKey(CLng(hKey), KeyName)
        If ret <> ERROR_SUCCESS Then
            DeleteEmptyKey = False
        Else
            DeleteEmptyKey = True
        End If

End Function


Public Function DeleteKey(hKey As REG_HKEY, KeyName As String) As Boolean
    Dim ret As Long
        ret = SHDeleteKey(CLng(hKey), KeyName)
        If ret <> ERROR_SUCCESS Then
            DeleteKey = False
        Else
            DeleteKey = True
        End If

End Function


Public Function SetStringValue(hKey As REG_HKEY, KeyName As String, ValueName As String, _
                               ByVal value As String) As Boolean
    Dim ret As Long
    Dim RegKeyHandle As Long
        ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle)
        If ret <> ERROR_SUCCESS Then
            RegCloseKey RegKeyHandle
            SetStringValue = False
            Exit Function
        End If

        ret = RegSetValueEx(RegKeyHandle, ValueName, 0, REG_SZ, ByVal value, Len(value))
        RegCloseKey RegKeyHandle
        SetStringValue = (ret = ERROR_SUCCESS)
End Function


Public Function SetLongValue(hKey As REG_HKEY, _
                             KeyName As String, _
                             ValueName As String, _
                             value As Long) As Boolean
    Dim ret As Long
    Dim RegKeyHandle As Long
        ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle)
        If ret <> ERROR_SUCCESS Then
            RegCloseKey RegKeyHandle
            SetLongValue = False
            Exit Function
        End If

        ret = RegSetValueEx(RegKeyHandle, ValueName, ByVal 0&, REG_DWORD, CLng(value), Len(value))
        RegCloseKey RegKeyHandle
        SetLongValue = (ret = ERROR_SUCCESS)
End Function



Public Function GetStringValue(hKey As REG_HKEY, KeyName As String, KeyValue As String, Optional Default As String = "") As String
    Dim ret As Long
    Dim RegKeyHandle As Long
    Dim Buffer(0 To 16299) As Byte
    Dim BufferLen As Long
    Dim s As String
    Dim i As Long
        ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle)
        If ret <> ERROR_SUCCESS Then
            RegCloseKey RegKeyHandle
            GetStringValue = Default
            Exit Function
        End If

        BufferLen = 16300
        ret = RegQueryValueEx(RegKeyHandle, KeyValue, ByVal 0&, REG_SZ, Buffer(0), BufferLen)
        If ret <> ERROR_SUCCESS Then
            GetStringValue = Default
            Exit Function
        End If

        s = ""
        If BufferLen >= 1 Then
            For i = 0 To BufferLen - 1 ' win32 - 2, win64 -1
                s = s + Chr(Buffer(i))
            Next i

        End If

        GetStringValue = s
End Function


Public Function GetLongValue(hKey As REG_HKEY, _
                             KeyName As String, _
                             KeyValue As String, _
                             Optional Default As Long = -1) As Long
    Dim ret As Long
    Dim RegKeyHandle As Long
    Dim Buffer As Long
    Dim BufferLen As Long

    Dim i As Long
        ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle)
        If ret <> ERROR_SUCCESS Then
            RegCloseKey RegKeyHandle
            GetLongValue = Default
            Exit Function
        End If

        BufferLen = Len(Buffer)
        ret = RegQueryValueEx(RegKeyHandle, KeyValue, ByVal 0&, REG_SZ, Buffer, BufferLen)
        If ret <> ERROR_SUCCESS Then
            GetLongValue = Default
            Exit Function
        End If

        GetLongValue = Buffer
End Function


Public Function SetBooleanValue(hKey As REG_HKEY, _
                                KeyName As String, _
                                KeyValue As String, _
                                value As Boolean) As Boolean
    SetBooleanValue = SetLongValue(hKey, KeyName, KeyValue, IIf(value = False, 0, 1))
End Function


Public Function GetBooleanValue(hKey As REG_HKEY, _
                                KeyName As String, _
                                KeyValue As String, _
                                Optional Default As Boolean = False) As Boolean
    Dim v As Long
    Dim ret As Long
        v = GetLongValue(hKey, KeyName, KeyValue, -1)

        If v <> -1 Then
            GetBooleanValue = IIf(v = 0, False, True)
            Exit Function
        End If

        GetBooleanValue = Default
End Function


