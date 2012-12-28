Attribute VB_Name = "Lng"
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function LngGetIDsInfo Lib "rb_loc.dll" (ids As Long, ByVal count As Long) As Long
Public Declare Function LngSwitchLanguage Lib "rb_loc.dll" (ByVal ID As Long, ByVal lpString As String) As Long
Public Declare Function LngGetLanguageID Lib "rb_loc.dll" () As Long
'Public Declare Function LngSwitchLanguage Lib "rb_loc.dll" Alias "LngGetLanguage" () As Long
'Public Declare Function LngSwitchLanguage Lib "rb_loc.dll" (ByVal id As Long) As Long
Public Declare Function LngGetString Lib "rb_loc.dll" (ByVal ID As Long, ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function LngGetVersion Lib "rb_loc.dll" () As Long
Public Declare Function LngGetAuthor Lib "rb_loc.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function LngGetFileName Lib "rb_loc.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function LngGetDescription Lib "rb_loc.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function LngGetLastSaved Lib "rb_loc.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function LngGetURL Lib "rb_loc.dll" (ByVal lpBuffer As String, nSize As Long) As Long

Public CountLanguages As Long
Public CurrentIdlanguage As Long
Public Idlanguage As Long
Public LngCharset As Integer

Public LIDs(10) As Long

Private Type ActivateT
c As String * 20
End Type

Public Function GetLastSaved() As String
Dim str As String
Dim strlen As Long
str = String$(255, " ")
strlen = LngGetLastSaved(str, 255)
If strlen <> 0 Then GetLastSaved = Left$(str, InStr(str, Chr$(0)) - 1)
End Function

Public Function GetURL() As String
Dim str As String
Dim strlen As Long
str = String$(255, " ")
strlen = LngGetURL(str, 255)
If strlen <> 0 Then GetURL = Left$(str, InStr(str, Chr$(0)) - 1)
End Function

Public Function GetDllFileName() As String
Dim str As String
Dim strlen As Long
str = String$(1024, " ")
strlen = LngGetFileName(str, 1024)
If strlen <> 0 Then GetDllFileName = Left$(str, InStr(str, Chr$(0)) - 1)
End Function

Public Function GetDescription() As String
Dim str As String
Dim strlen As Long
str = String$(1024, " ")
strlen = LngGetDescription(str, 1024)
If strlen <> 0 Then GetDescription = Left$(str, InStr(str, Chr$(0)) - 1)
End Function


Public Function GetAuthor() As String
Dim str As String
Dim strlen As Long
str = String$(255, " ")
strlen = LngGetAuthor(str, 255)
If strlen <> 0 Then GetAuthor = Left$(str, InStr(str, Chr$(0)) - 1)
End Function

Public Function GetResIDstring(ByVal resID As Integer, ParamArray varReplacements()) As String
Dim intMacro As Integer
Dim strResString As String
    
Dim str As String
Dim strlen As Long

On Error GoTo MismatchedPairs
str = String$(4096, " ")
strlen = LngGetString(resID, str, 4096)
'On Error Resume Next

If strlen <> 0 Then
strResString = Left(str, strlen)

'strResString = LoadResstring(resID)
'If strResString = "" Then GoTo MismatchedPairs
        
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        If UBound(varReplacements) > 0 Then
        strValue = varReplacements(intMacro + 1)
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
        End If
        
    Next intMacro
    
    If InStr(strResString, "|1") > 0 Then strResString = Replace(strResString, "|1", vbCrLf)
    GetResIDstring = strResString
Else
    GetResIDstring = "[" & resID & "]"
End If

Exit Function
MismatchedPairs:
'STRERROR = STRERROR & Time & ". (modul10) ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
GetResIDstring = "[" & resID & "]" '"String not found."
End Function

Public Function GetLngCode(c As Long) As String
Dim FileN As Integer
Dim FileName As String

Dim Activate As ActivateT
On Error Resume Next
FileN = FreeFile
FileName = App.Path & "\LNG\lng" & c & ".key"
If dir(FileName, vbNormal) <> "" Then
Open FileName For Random As #FileN Len = Len(Activate)
Get #FileN, 1, Activate
Close #FileN
GetLngCode = Trim(Activate.c)
End If
End Function

Public Sub SetLngCode(c As Long, lc As String)
Dim FileN As Integer
Dim Activate As ActivateT
On Error Resume Next
Activate.c = lc
FileN = FreeFile
Open App.Path & "\LNG\lng" & c & ".key" For Random As #FileN Len = Len(Activate)
Put #FileN, , Activate
Close #FileN
End Sub
