VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lng_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function LngGetIDsInfo Lib "rb_loc" (ids As Long, ByVal count As Long) As Long
Private Declare Function LngSwitchLanguage Lib "rb_loc" (ByVal id As Long, ByVal lpString As String) As Long
Private Declare Function LngGetLanguageID Lib "rb_loc" () As Long
Private Declare Function LngGetString Lib "rb_loc" (ByVal id As Long, ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function LngGetVersion Lib "rb_loc" () As Long
Private Declare Function LngGetAuthor Lib "rb_loc" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LngGetFileName Lib "rb_loc" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LngGetDescription Lib "rb_loc" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LngGetLastSaved Lib "rb_loc" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LngGetURL Lib "rb_loc" (ByVal lpBuffer As String, nSize As Long) As Long

Public CountLanguages As Long
Public CurrentCodelanguage As Long
Public Idlanguage As Long
Public LngCharset As Integer

Private LIDs(10) As Long

Private Type ActivateT
    c As String * 20
End Type

Public Function GetVersion() As String
On Error GoTo ERR
    GetVersion = GetLoWord(LngGetVersion) & "." & GetHiWord(LngGetVersion)
ERR:
End Function


Public Function GetLanguageID() As Long
On Error GoTo ERR
    GetLanguageID = LngGetLanguageID
ERR:
End Function


Public Function GetIDsInfo() As Long
On Error GoTo ERR
    GetIDsInfo = LngGetIDsInfo(LIDs(0), 10)
ERR:
End Function


Public Function GetIDIfo(id As Long) As Long
On Error GoTo ERR
    GetIDIfo = LIDs(id)
ERR:
End Function


Public Function GetLocale(id As Long) As String
On Error GoTo ERR
    Dim str As String
    str = String$(100, " ")
    GetLocaleInfo id, 2, str, 100
    GetLocale = Trim(str)
ERR:
End Function


Public Function GetLastSaved() As String
On Error GoTo ERR
    Dim str As String
    Dim strlen As Long
    str = String$(255, " ")
    strlen = LngGetLastSaved(str, 255)
    If strlen <> 0 Then GetLastSaved = Left$(str, InStr(str, Chr$(0)) - 1)
ERR:
End Function


Public Function GetURL() As String
On Error GoTo ERR
    Dim str As String
    Dim strlen As Long
    str = String$(255, " ")
    strlen = LngGetURL(str, 255)
    If strlen <> 0 Then GetURL = Left$(str, InStr(str, Chr$(0)) - 1)
ERR:
End Function


Public Function GetDllFileName() As String
On Error GoTo ERR
    Dim str As String
    Dim strlen As Long
    str = String$(1024, " ")
    strlen = LngGetFileName(str, 1024)
    If strlen <> 0 Then GetDllFileName = Left$(str, InStr(str, Chr$(0)) - 1)
ERR:
End Function


Public Function GetDescription() As String
On Error GoTo ERR
    Dim str As String
    Dim strlen As Long
    str = String$(1024, " ")
    strlen = LngGetDescription(str, 1024)
    If strlen <> 0 Then GetDescription = Left$(str, InStr(str, Chr$(0)) - 1)
ERR:
End Function

Public Function GetAuthor() As String
On Error GoTo ERR
    Dim str As String
    Dim strlen As Long
    str = String$(255, " ")
    strlen = LngGetAuthor(str, 255)
    If strlen <> 0 Then GetAuthor = Left$(str, InStr(str, Chr$(0)) - 1)
ERR:
End Function


Public Function GetResIDstring(ByVal resID As Long, ParamArray varReplacements()) As String
    Dim intMacro As Integer
    Dim strResString As String
    
    Dim str As String
    Dim strlen As Long

    On Error GoTo MismatchedPairs

    str = String$(4096, " ")
    strlen = LngGetString(resID, str, 4096)

    If strlen <> 0 Then
    strResString = Left(str, strlen)

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
    
'
' дополнительная обработка
'
If InStr(strResString, "|1") > 0 Then strResString = Replace(strResString, "|1", vbNewLine)
    
GetResIDstring = strResString
Else
    GetResIDstring = resID
End If

Exit Function
MismatchedPairs:
GetResIDstring = resID
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


Public Sub SetLngCode(c As Long, _
                      lC As String)
    Dim FileN As Integer
    Dim Activate As ActivateT
        On Error Resume Next
        If SwitchLanguage(c, lC) = 0 Then
            Activate.c = lC
            FileN = FreeFile
            Open App.Path & "\LNG\lng" & c & ".key" For Random As #FileN Len = Len(Activate)
            Put #FileN, , Activate
            Close #FileN
        End If

End Sub


Public Function SwitchLanguage(CurrentCodelanguage As Long, _
                               Optional code As String) As Long
    If code <> "" Then
        SwitchLanguage = LngSwitchLanguage(CurrentCodelanguage, code)
    Else
        SwitchLanguage = LngSwitchLanguage(CurrentCodelanguage, GetLngCode(CurrentCodelanguage))
    End If

End Function



Private Sub Class_Initialize()
    CountLanguages = GetIDsInfo
End Sub


