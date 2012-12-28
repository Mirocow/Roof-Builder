Attribute VB_Name = "SYS"
Option Explicit
Private Const ANYSIZE_ARRAY As Long = 0
Private Const MAX_MESSAGE_LEN As Long = 200
Private Const MAX_NAME_LEN As Long = 250
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const LANG_NEUTRAL As Long = &H0
Private Const TOKEN_QUERY As Long = &H8
Private Const TokenGroups As Integer = 2
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Private Const SECURITY_BUILTIN_DOMAIN_RID As Long = &H20
Private Const DOMAIN_ALIAS_RID_ADMINS As Long = &H220
Private Const ERROR_NONE_MAPPED As Long = 1332&
Private Type SID_IDENTIFIER_AUTHORITY
value(5) As Byte
End Type
Private Type SID_AND_ATTRIBUTES
pSid As Long
Attributes As Long
End Type
Private Type TOKEN_GROUPS
GroupCount As Long
Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type
Private Declare Function GetCurrentProcess Lib "kernel32" ( _
) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
ByVal ProcessHandle As Long, _
ByVal DesiredAccess As Long, _
ByRef TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" ( _
ByVal TokenHandle As Long, _
ByVal TokenInformationClass As Integer, _
ByRef TokenInformation As Any, _
ByVal TokenInformationLength As Long, _
ByRef ReturnLength As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" ( _
ByRef pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, _
ByVal nSubAuthorityCount As Byte, _
ByVal nSubAuthority0 As Long, _
ByVal nSubAuthority1 As Long, _
ByVal nSubAuthority2 As Long, _
ByVal nSubAuthority3 As Long, _
ByVal nSubAuthority4 As Long, _
ByVal nSubAuthority5 As Long, _
ByVal nSubAuthority6 As Long, _
ByVal nSubAuthority7 As Long, _
ByRef lpPSid As Long) As Long
Private Declare Function EqualSid Lib "advapi32.dll" ( _
ByVal pSid1 As Long, _
ByVal pSid2 As Long) As Long
Private Declare Function IsValidSid Lib "advapi32.dll" ( _
ByVal pSid As Long) As Long
Private Declare Function ConvertSidToStringSid Lib "advapi32.dll" Alias "ConvertSidToStringSidA" ( _
ByVal pSid As Long, ByRef str1 As Long) As Long
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" ( _
ByVal lpSystemName As String, _
ByVal pSid As Long, _
ByVal name As String, _
ByRef cbName As Long, _
ByVal ReferencedDomainName As String, _
ByRef cbReferencedDomainName As Long, _
ByRef peUse As Integer) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
ByVal dwFlags As Long, lpSource As Any, _
ByVal dwMessageId As Long, _
ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, _
ByVal nSize As Long, _
ByRef Arguments As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" ( _
ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Sub FreeSid Lib "advapi32.dll" ( _
ByVal pSid As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Function IsAdmin() As Boolean
Dim nErr As Long, i As Long

'Получили токен
Dim hToken As Long: hToken = 0
If OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, hToken) = 0 Then
nErr = ERR.LastDllError
GoTo Errors
End If

'Выделяем память под список групп
Dim pGroupInfo() As TOKEN_GROUPS
Dim nbSize As Long: nbSize = 0
nErr = 0
If GetTokenInformation(hToken, TokenGroups, ByVal 0&, nbSize, nbSize) = 0 Then
nErr = ERR.LastDllError
If nErr <> ERROR_INSUFFICIENT_BUFFER Then
GoTo Errors
End If
End If
'Вот такой извращённый способ застолбить память
' (чтобы не было System Crush)
Dim nIndex As Long, nGroups As Long
ReDim pGroupInfo(0)
nIndex = nbSize / LenB(pGroupInfo(0))
ReDim pGroupInfo(nIndex)

'Получили список групп
If GetTokenInformation(hToken, TokenGroups, pGroupInfo(0), nbSize, nbSize) = 0 Then
nErr = ERR.LastDllError
GoTo Errors
End If

'Чтобы избежать ошибки "Subscript is out of range"
nGroups = pGroupInfo(0).GroupCount
Dim aGroups() As SID_AND_ATTRIBUTES
ReDim aGroups(nGroups - 1)
nbSize = LenB(aGroups(0)) * nGroups
CopyMemory aGroups(0), pGroupInfo(0).Groups(0), nbSize

Dim SIDAuth As SID_IDENTIFIER_AUTHORITY
SIDAuth.value(0) = 0
SIDAuth.value(1) = 0
SIDAuth.value(2) = 0
SIDAuth.value(3) = 0
SIDAuth.value(4) = 0
SIDAuth.value(5) = 5
Dim pSidAdmin As Long: pSidAdmin = 0
If AllocateAndInitializeSid(SIDAuth, 2, _
SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, _
0, 0, 0, 0, 0, 0, pSidAdmin) = 0 Then
nErr = ERR.LastDllError
GoTo Errors
End If

'Это для того, чтобы убедится, что СИД корректный
'Все вызовы функции SidToStr после отладки надо удалить
Dim sTest As String, sTest2 As String
'sTest = SidToStr(pSidAdmin)

For i = 0 To nGroups - 1
If EqualSid(aGroups(i).pSid, pSidAdmin) > 0 Then
IsAdmin = True
'sTest2 = SidToStr(aGroups(i).pSid)
'sTest = AccName(aGroups(i).pSid)
'MsgBox sTest2 & vbNewLine & sTest
Exit For
End If
Next

If pSidAdmin > 0 Then FreeSid (pSidAdmin)
Exit Function
Errors:
'обработка ошибок кривая - "на скорую руку"
Dim sDescr As String
sDescr = String(MAX_MESSAGE_LEN, 0)
FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, nErr, LANG_NEUTRAL, sDescr, MAX_MESSAGE_LEN, ByVal 0&
sDescr = Left$(sDescr, InStr(sDescr, Chr(0)) - 1)
MsgBox "Ошибка: " & nErr & vbNewLine & sDescr
End Function

Private Function SidToStr(ByVal pSid As Long)
If IsValidSid(pSid) = 0 Then
Exit Function
End If
Dim bstrTemp As String
Dim pstrTemp As Long
If ConvertSidToStringSid(pSid, pstrTemp) = 0 Then
Exit Function
End If
Dim nLen As Long
nLen = lstrlen(pstrTemp)
bstrTemp = String(nLen, 0)
Call lstrcpy(bstrTemp, pstrTemp)
SidToStr = bstrTemp
End Function

Private Function AccName(ByVal pSid As Long) As String
Dim nErr As Long
Dim sAccountName As String, sDomainName As String
Dim nAccountName As Long, nDomainName As Long
nAccountName = MAX_NAME_LEN
sAccountName = String(nAccountName, 0)
nDomainName = MAX_NAME_LEN
sDomainName = String(nDomainName, 0)
Dim eSidType As Integer
If LookupAccountSid(vbNullString, pSid, sAccountName, _
nAccountName, sDomainName, nDomainName, eSidType) = 0 Then
nErr = ERR.LastDllError
If nErr <> ERROR_NONE_MAPPED Then
GoTo Errors
End If
sAccountName = "NONE_MAPPED"
Else
sDomainName = Left$(sDomainName, InStr(sDomainName, Chr(0)) - 1)
sAccountName = Left$(sAccountName, InStr(sAccountName, Chr(0)) - 1)
End If
AccName = sDomainName & "\" & sAccountName
Exit Function
Errors:
'обработка ошибок кривая - "на скорую руку"
Dim sDescr As String
sDescr = String(MAX_MESSAGE_LEN, 0)
FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, nErr, LANG_NEUTRAL, sDescr, MAX_MESSAGE_LEN, ByVal 0&
sDescr = Left$(sDescr, InStr(sDescr, Chr(0)) - 1)
MsgBox "Ошибка: " & nErr & vbNewLine & sDescr
End Function

Public Function ParsProfilData(var As String, Optional part As Index = 1) As String
Dim a() As String
a = Split(var, Space(100))
If part = 1 Then
ParsProfilData = a(0)
Else
ParsProfilData = a(1)
End If
End Function


