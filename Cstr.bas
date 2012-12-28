Attribute VB_Name = "Functions"
Option Explicit
Const NONE = 0
Const STRINGTYPE = 1
Const INTEGERTYPE = 2
Const LONGTYPE = 3
Const FLOATTYPE = 4
Const CHARPERCENT = 5

Public Function ArraySize(ByRef a) As Long
    On Error GoTo ERR
    ArraySize = UBound(a) + 1
    Exit Function
ERR:
    ArraySize = 0
    ERR.clear
End Function

'Public Function SprintF(sFormats As String, ParamArray aArguments() As Variant) As String
'
'    Dim nCurrentFlag As Integer
'    Dim nPos As Integer
'    Dim sWork As String
'    Dim nCurVal As Integer
'    Dim nMaxArg As Integer
'    Dim sCurFormat As String
'    Dim nArgCount As Integer
'    Dim nxIndex As Integer
'    Dim bFound As Boolean
'    Dim nType As Integer
'    Dim sCurParm As String
'    Dim nLenFlags As Integer
'    Dim nWidth As Integer
'    ' *** Get the number of arguments
'    nMaxArg = UBound(aArguments) + 1
'    ' *** Length of the flags
'    nLenFlags = Len(sFormats)
'    ' *** Initialize some variables
'    nCurrentFlag = 1
'    nCurVal = 0
'    nArgCount = 0
'    ' *** Get the first flag
'    nPos = InStr(nCurrentFlag, sFormats, "%")
'    ' *** Verify if the number of flags is t
'    '     he same as the number of argument
'
'
'    Do While (nPos > 0)
'
'
'        If mID$(sFormats, nPos + 1, 1) <> "%" Then ' *** Don't count %%, will be converted To % later
'            nArgCount = nArgCount + 1
'            nCurrentFlag = nPos + 1
'        Else
'            nCurrentFlag = nPos + 2
'        End If
'
'
'        ' *** Get next flag
'        nPos = InStr(nCurrentFlag, sFormats, "%")
'    Loop
'
'    ' *** Compare the number of flags agains
'    '     t the number of arguments
'    If nArgCount <> nMaxArg Then ERR.Raise 450, , "Mismatch of parameters For String " & sFormats & ". Expected " & nArgCount & " but received " & nMaxArg & "."
'
'    ' *** Initialize some variables
'    nCurrentFlag = 1
'    nCurVal = 0
'    nArgCount = 0
'    sWork = ""
'    ' *** Get the first flag
'    nPos = InStr(nCurrentFlag, sFormats, "%")
'
'
'    Do While (nPos > 0)
'        ' *** First, get the variable identifier
'        '     .
'        ' *** Scan from nCurrentFlag (the %) to
'        '     EOL looking for the
'        ' *** first occurance of s, d, l, or f
'        bFound = False
'        nType = NONE
'        nxIndex = nPos + 1
'
'
'        Do While (bFound = False) And (nxIndex <= nLenFlags)
'
'
'            If Not bFound Then
'                sCurParm = mID$(sFormats, nxIndex, 1)
'
'
'                Select Case mID$(sFormats, nxIndex, 1)
'                    Case "%"
'                    nType = CHARPERCENT
'                    bFound = True
'                    nPos = nPos + 1
'                    nCurVal = nxIndex + 2
'                    Case "s", "S"
'                    nType = STRINGTYPE
'                    bFound = True
'                    nCurVal = nxIndex + 1
'                    Case "d", "i", "u"
'                    nType = INTEGERTYPE
'                    bFound = True
'                    nCurVal = nxIndex + 1
'                    Case "l"
'
'
'                    If mID$(sFormats, nxIndex + 1, 1) = "d" Then
'                        nType = LONGTYPE
'                        bFound = True
'                        nCurVal = nxIndex + 2
'                    Else
'                        ERR.Raise 93, , "Unrecognized pattern " & mID$(sFormats, nxIndex - 1, 3) & " in " & sFormats
'                    End If
'
'                    Case "f", "e", "E", "g", "G"
'                    nType = FLOATTYPE
'                    bFound = True
'                    nCurVal = nxIndex + 1
'                End Select
'
'        End If
'
'
'        If Not bFound Then nxIndex = nxIndex + 1
'
'    Loop
'
'
'    ' *** Not found, raise an error
'    If Not bFound Then ERR.Raise 93, , "Invalid % format in <" & sFormats & ">"
'
'    ' *** Get the complete flag
'    sCurParm = mID$(sFormats, nPos, nCurVal - nPos)
'
'    ' *** Different case if Percent or other
'    '
'
'    If nType = CHARPERCENT Then
'        sWork = sWork & mID$(sFormats, nCurrentFlag, nPos - nCurrentFlag)
'        nCurVal = nCurVal - 1
'    Else
'        sCurFormat = BuildFormat(sCurParm, nType, aArguments(nArgCount))
'        sWork = sWork & TreatBackSlash(mID$(sFormats, nCurrentFlag, nPos - nCurrentFlag)) & sCurFormat
'        nArgCount = nArgCount + 1
'    End If
'    nCurrentFlag = nCurVal
'    ' *** Get next flag
'    nPos = InStr(nCurrentFlag, sFormats, "%")
'
'Loop
'
''    If nType = CHARPERCENT Then
''        sWork = sWork & Mid$(sFormats, nCurrentFlag, nPos - nCurrentFlag)
''        nCurVal = nCurVal - 1
''    Else
''        sCurFormat = BuildFormat(sCurParm, nType, aArguments(nArgCount))
''        sWork = sWork & Mid$(sFormats, nCurrentFlag, nPos - nCurrentFlag) & sCurFormat
''        nArgCount = nArgCount + 1
''    End If
''
''    nCurrentFlag = nCurVal
''    ' *** Get next flag
''    nPos = InStr(nCurrentFlag, sFormats, "%")
''Loop
'
'
'SprintF = sWork
'
''SprintF = TreatBackSlash(sWork)
'End Function
'
'
'
'Function BuildFormat(sFormat As String, nDataType As Integer, vCurrentValue As Variant) As String
'
'    ' *** Build the format
'
'    Dim sPrefix As String
'    Dim sFlag As String
'    Dim nWidth As Long
'    Dim bAlignLeft As Boolean
'    Dim bSign As Boolean
'    Dim sPad As String * 1
'    Dim bBlank As Boolean
'    Dim bDecimal As Boolean
'    Dim nI As Integer
'    Dim sTmp As String
'    Dim sWidth As String
'    Dim nPrecision As Integer
'    Dim nPlaces As Integer
'
'
'    If (Len(sFormat) < 2) Then
'        BuildFormat = ""
'        Exit Function
'    End If
'
'    ' *** Get the flag
'    sFlag = ""
'    bAlignLeft = False
'    bSign = False
'    sPad = "@"
'    bBlank = False
'    bDecimal = False
'
'
'    Select Case mID$(sFormat, 2, 1)
'        Case "-":
'        bAlignLeft = True
'        sFormat = mID$(sFormat, 3)
'
'        Case "+":
'        bSign = True
'        sFormat = mID$(sFormat, 3)
'
'        Case "0":
'        sPad = "0"
'        sFormat = mID$(sFormat, 3)
'
'        Case " ":
'        bBlank = True
'        sFormat = mID$(sFormat, 3)
'
'        Case "#":
'        bDecimal = True
'        sFormat = mID$(sFormat, 3)
'
'        Case Else
'        sFormat = mID$(sFormat, 2)
'
'    End Select
'
'' *** Get the width
'
'
'If nDataType = LONGTYPE Then
'    sPrefix = mID$(sFormat, 1, Len(sFormat) - 2)
'Else
'    sPrefix = mID$(sFormat, 1, Len(sFormat) - 1)
'End If
'
'' *** Get the width
'sWidth = ""
'nI = 1
'sTmp = mID$(sPrefix, nI, 1)
'
'
'Do While IsNumeric(sTmp)
'    sWidth = sWidth & sTmp
'
'    nI = nI + 1
'    sTmp = mID$(sPrefix, nI, 1)
'Loop
'
'' *** Check the precision
'nPrecision = InStr(sPrefix, ".")
'
'
'If (nPrecision = 0) Then
'    ' *** No precision, only width (eventual
'    '     ly)
'    If (Trim(sWidth) = "") Then sWidth = "0"
'    nWidth = CLng(sWidth)
'
'
'    If (bAlignLeft = False) Then
'        sFormat = String(nWidth, sPad)
'    Else
'        If (Len(CStr(vCurrentValue)) > nWidth) Then nWidth = Len(CStr(vCurrentValue))
'        sFormat = String(Len(CStr(vCurrentValue)), sPad) & String(nWidth - Len(CStr(vCurrentValue)), " ")
'    End If
'
'Else
'    sTmp = "0"
'    nI = nPrecision + 1
'
'
'    Do While IsNumeric(mID(sPrefix, nI, 1))
'        sTmp = sTmp & mID(sPrefix, nI, 1)
'        nI = nI + 1
'    Loop
'
'
'    nPlaces = CLng(sTmp)
'
'
'
'    Select Case nDataType
'        Case INTEGERTYPE, LONGTYPE:
'        sFormat = String(nPlaces, "0")
'        Case FLOATTYPE:
'        sFormat = "#." & String(nPlaces, "0")
'    End Select
'
'End If
'
'BuildFormat = Format$(vCurrentValue, sFormat)
'End Function
'
'
'
'Private Function TreatBackSlash(sLine As String) As String
'    ' *** Treat all the backslach
'    Dim nI As Integer
'    Dim nPos As Integer
'    Dim sChar As String * 1
'    ' *** Get the first backslash
'    nI = 1
'    nPos = InStr(nI, sLine, "\")
'
'
'    Do While (nPos > 0)
'        ' *** First, get the char after
'        sChar = mID$(sLine, nPos + 1, 1)
'
'
'        Select Case sChar
'            Case "n"
'            sLine = Left(sLine, nPos - 1) & Chr$(13) & Chr$(10) & Right(sLine, Len(sLine) - nPos - 1)
'            nPos = nPos + 1
'            Case "r"
'            sLine = Left(sLine, nPos - 1) & Chr$(13) & Right(sLine, Len(sLine) - nPos - 1)
'            nPos = nPos + 1
'            Case "t"
'            sLine = Left(sLine, nPos - 1) & Chr$(9) & Right(sLine, Len(sLine) - nPos - 1)
'            nPos = nPos + 1
'            Case "\"
'            sLine = Left(sLine, nPos - 1) & "\" & Right(sLine, Len(sLine) - nPos - 1)
'            nPos = nPos + 1
'            Case Else
'            ERR.Raise 93, , "Invalid escape sequence: \" & sChar
'        End Select
'
'
'    nPos = InStr(nPos, sLine, "\")
'Loop
'
'TreatBackSlash = sLine
'End Function


Public Function TimeStamp() As String
    Dim StartDate As String
    Dim EndTime As String
    Dim StartTime As String
    Dim EndDate As String
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim DateTimeStart As Date
    Dim DateTimeEnd As Date
    Dim TotalHrs As String
    StartDate = "1/1/1970"
    StartTime = "00:00:00"
    
    EndDate = CStr(Date)
    EndTime = CStr(Time)
    
    DateTimeStart = FormatDateTime(StartDate & " " & StartTime)
    DateTimeEnd = FormatDateTime(EndDate & " " & EndTime)
    TimeStamp = DateDiff("s", DateTimeStart, DateTimeEnd, vbUseSystemDayOfWeek, _
    vbUseSystem)
End Function


Public Function RC4(ByRef ByteArray() As Byte, ByVal Password As String) As String
Call VarPtr("VMProtect begin")
On Error Resume Next
Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, Key() As Byte, temp As Byte
If Len(Password) = 0 Then
    Exit Function
End If
If ArraySize(ByteArray) = 0 Then
    Exit Function
End If
If Len(Password) > 256 Then
    Key() = StrConv(Left$(Password, 256), vbFromUnicode)
Else
    Key() = StrConv(Password, vbFromUnicode)
End If
For X = 0 To 255
    RB(X) = X
Next X
X = 0
Y = 0
Z = 0
For X = 0 To 255
    Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
    temp = RB(X)
    RB(X) = RB(Y)
    RB(Y) = temp
Next X
X = 0
Y = 0
Z = 0
For X = 0 To ArraySize(ByteArray) - 1
    Y = (Y + 1) Mod 256
    Z = (Z + RB(Y)) Mod 256
    temp = RB(Y)
    RB(Y) = RB(Z)
    RB(Z) = temp
    ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
Next X
RC4 = StrConv(ByteArray, vbUnicode)
Call VarPtr("VMProtect end")
End Function

Function IP2Long(IP As String) As Long
Call VarPtr("VMProtect begin")
On Error GoTo ERR
Dim d() As String
Dim long_ip As Long

d = Split(IP, ".")

long_ip = long_ip Or d(0)
long_ip = ShiftToLeft(long_ip, 8)

long_ip = long_ip Or d(1)
long_ip = ShiftToLeft(long_ip, 8)

long_ip = long_ip Or d(2)
long_ip = ShiftToLeft(long_ip, 8)

long_ip = long_ip Or d(3)

ERR:
IP2Long = long_ip
Call VarPtr("VMProtect end")
End Function


Public Function ShiftToLeft(ByVal value As Long, ByVal Shift As Integer) As Long
    Dim OnBits(0 To 31) As Long
    MakeOnBits OnBits
    If (value And (2 ^ (31 - Shift))) Then GoTo OverFlow
    ShiftToLeft = ((value And OnBits(31 - Shift)) * (2 ^ Shift))
    Exit Function
OverFlow:
    ShiftToLeft = ((value And OnBits(31 - (Shift + 1))) * (2 ^ (Shift))) Or &H80000000
End Function

Public Function ShiftToRight(ByVal value As Long, ByVal Shift As Integer) As Long
    Dim OnBits(0 To 31) As Long
    Dim hi As Long
    MakeOnBits OnBits
    If (value And &H80000000) Then hi = &H40000000
    ShiftToRight = (value And &H7FFFFFFE) \ (2 ^ Shift)
    ShiftToRight = (ShiftToRight Or (hi \ (2 ^ (Shift - 1))))
End Function

Private Sub MakeOnBits(ByRef OnBits)
    Dim j As Integer, v As Long
    For j = 0 To 30
        v = v + (2 ^ j)
        OnBits(j) = v
    Next j
    OnBits(j) = v + &H80000000
End Sub


Public Function HextoDec(HexNum As String) As Byte
'Call VarPtr("VMProtect begin")
On Error GoTo ERR
    'converts a hexadecimal value to a decim
    '     al value
    'You can use the characters a-f but also
    '     A-F (in capitals)
    'for example: label1.caption = HextoDec(
    '     "Ab789Ff")
    'returns as the labels caption: 17980057
    '     5
    'an error handling is included
    Dim xx%, yy%

    For xx = 1 To Len(HexNum)
        If Asc(mID(HexNum, xx, 1)) < 48 Then GoTo ERR
        If Asc(mID(HexNum, xx, 1)) > 57 And Asc(mID(HexNum, xx, 1)) < 65 Then GoTo ERR
        If Asc(mID(HexNum, xx, 1)) > 70 And Asc(mID(HexNum, xx, 1)) < 97 Then GoTo ERR
        If Asc(mID(HexNum, xx, 1)) > 102 Then GoTo ERR
    Next xx

    HextoDec = "&h" & HexNum
ERR:
'Call VarPtr("VMProtect end")
End Function

Public Function FormaString(v As String, nLen As Integer)
Dim sFormat As String
sFormat = String(nLen, "0")
FormaString = Format$(v, sFormat)
End Function


Public Function TrimNullChar(var As String) As String
TrimNullChar = Trim(Replace(var, Chr(0), ""))
End Function

'Public Function GrayScale(coloredColor As Long) As Long
'  'Convertir un Long RGB a un LongRGB grayscale
'  'For this task Imported previously developed functions from my project Opticlops
'  '' Desc: Convert a RGB color to long
'  'Private Function RGBToLong(RGBColor As RGB) As Long
'  '    RGBToLong = RGBColor.Blue + RGBColor.Green * 265 + RGBColor.Red * 65536
'  'End Function
'  '
'  '' Desc Convert a long into a RGB structure
'  'Private Function LongToRGB(lcolor As Long) As RGB
'  '    LongToRGB.Red = lcolor And &HFF
'  '    LongToRGB.Green = (lcolor \ &H100) And &HFF
'  '    LongToRGB.Blue = (lcolor \ &H10000) And &HFF
'  'End Function
'  On Error GoTo Error
'
'  Dim R As Long, g As Long, b As Long
'  Dim neutral As Long
'  'Splitt into RGB values
'  b = coloredColor And &HFF
'  g = (coloredColor \ &H100) And &HFF
'  R = (coloredColor \ &H10000) And &HFF
'  'Obtener el promedio
'  neutral = (R / 3 + g / 3 + b / 3)
'  'Build Long
'  GrayScale = RGB(neutral, neutral, neutral)
'  Exit Function
'
'Error:
'End Function

'Public Function hiByte(ByVal w As Integer) As Byte
'    If w And &H8000 Then
'        hiByte = &H80 Or ((w And &H7FFF) / &HFF)
'    Else
'        hiByte = w / 256
'    End If
'End Function
'
'
'Public Function HiWord(dw As Long) As Integer
'    If dw And &H80000000 Then
'        HiWord = (dw / 65535) - 1
'    Else
'        HiWord = dw / 65535
'    End If
'End Function
'
'
'Public Function LoByte(w As Integer) As Byte
'    LoByte = w And &HFF
'End Function
'
'
'Public Function LoWord(dw As Long) As Integer
'    If dw And &H8000& Then
'        LoWord = &H8000 Or (dw And &H7FFF&)
'    Else
'        LoWord = dw And &HFFFF&
'    End If
'End Function
'
'
'Public Function MakeInt(ByVal LoByte As Byte, ByVal hiByte As Byte) As Integer
'    MakeInt = ((hiByte * &H100) + LoByte)
'End Function
'
'
'Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
'    MakeLong = ((HiWord * &H10000) + LoWord)
'End Function
'


