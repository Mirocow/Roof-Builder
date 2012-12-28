VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const cVer = 1.2

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Public Enum FAccess
    aRead
    aWrite
    aRead_Write
End Enum

Public Enum vFormat
    vLong
    vSingle
End Enum

Private FileN As Integer
Private FileStatus As Integer
'Private SystemInfo As Integer


Public Property Get Ver() As Single
    Ver = cVer
End Property

'Public Property Get Sys() As Integer
'    Sys = SystemInfo
'End Property

Public Function FOpen(ByRef Path As String, Optional ByVal Access As FAccess = aRead, Optional ByVal position As Long = 1, Optional ByVal Wipe As Boolean = False) As Boolean
Dim signature As String
On Error GoTo ERR
FileN = FreeFile
Select Case Access
Case 0
    
    If FileExists(Path) Then
        Open Path For Binary Access Read As #FileN
        FileStatus = 1
'        Dim mcVer As Single
'        If FSeek = 1 Then
'            FRead mcVer
'            FReadString signature, vSingle
'            SystemInfo = FSeek
'        Else
'            FSeek 1
'            FReadString signature, vSingle
'            FSeek position
'        End If
'        If cVer <> mcVer Then GoTo ERR
    Else
        GoTo ERR
    End If

Case 1
    
    If FileExists(Path) And Wipe Then
        Kill Path
    End If
    Open Path For Binary Access Write As #FileN
    FileStatus = 2
    ' Пишем шапку нашего формата
'    If FSeek = 1 Then
'        FWrite CSng(cVer)
'        FWriteString "iStorage by MDinc @2010", vSingle
'        SystemInfo = FSeek
'    End If
        
Case 2
    
    If position > 1 Then FSeek position
    Open Path For Binary Access Read Write As #FileN
    FileStatus = 3
    ' Пишем шапку нашего формата
'    If FSeek = 1 Then
'        FWrite CSng(cVer)
'        FWriteString "iStorage by MDinc @2010", vSingle
'        SystemInfo = FSeek
'    End If
        
End Select


FOpen = True
Exit Function
ERR:
FileN = 0
FileStatus = 0
FOpen = False
End Function

Public Function FWrite(data As Variant, Optional recnumber As Integer) As Boolean
On Error GoTo ERR
If recnumber > 0 Then
    Select Case VarType(data)
    Case vbInteger
        Put #FileN, recnumber, CInt(data)
    Case vbLong
        Put #FileN, recnumber, CLng(data)
    Case vbSingle
        Put #FileN, recnumber, CSng(data)
    Case vbDouble
        Put #FileN, recnumber, CDbl(data)
    Case vbString
        Put #FileN, recnumber, CStr(data)
    Case vbBoolean
        Put #FileN, recnumber, CBool(data)
    Case vbByte
        Put #FileN, recnumber, CByte(data)
    Case vbCurrency
        Put #FileN, recnumber, CCur(data)
    End Select
Else
    Select Case VarType(data)
    Case vbInteger
        Put #FileN, , CInt(data)
    Case vbLong
        Put #FileN, , CLng(data)
    Case vbSingle
        Put #FileN, , CSng(data)
    Case vbDouble
        Put #FileN, , CDbl(data)
    Case vbString
        Put #FileN, , CStr(data)
    Case vbBoolean
        Put #FileN, , CBool(data)
    Case vbByte
        Put #FileN, , CByte(data)
    Case vbCurrency
        Put #FileN, , CCur(data)
    End Select
End If
FWrite = True
Exit Function
ERR:
FWrite = False
End Function


Public Function FRead(ByRef data As Variant, Optional recnumber As Integer, Optional buflen As Long = 0) As Boolean
On Error GoTo ERR

Dim iData As Integer
Dim lData As Long
Dim sdata As Single
Dim dData As Double
Dim stData As String
Dim bData As Boolean
Dim byData As Byte
Dim cData As Currency

If recnumber > 0 Then
    Select Case VarType(data)
    Case vbInteger
        Get #FileN, recnumber, iData
        data = CInt(iData)
    Case vbLong
        Get #FileN, recnumber, lData
        data = CLng(lData)
    Case vbSingle
        Get #FileN, recnumber, sdata
        data = CSng(sdata)
    Case vbDouble
        Get #FileN, recnumber, dData
        data = CDbl(dData)
    Case vbString
        If buflen > 0 Then
            stData = Space$(buflen)
        Else
            stData = Space$(Len(data))
        End If
        Get #FileN, recnumber, stData
        data = stData
    Case vbBoolean
        Get #FileN, recnumber, bData
        data = CBool(bData)
    Case vbByte
        Get #FileN, recnumber, byData
        data = CByte(byData)
    Case vbCurrency
        Get #FileN, recnumber, cData
        data = CCur(cData)
    End Select
Else
    Select Case VarType(data)
    Case vbInteger
        Get #FileN, , iData
        data = CInt(iData)
    Case vbLong
        Get #FileN, , lData
        data = CLng(lData)
    Case vbSingle
        Get #FileN, , sdata
        data = CSng(sdata)
    Case vbDouble
        Get #FileN, , dData
        data = CDbl(dData)
    Case vbString
        If buflen > 0 Then
        stData = Space$(buflen)
        Else
        stData = Space$(Len(data))
        End If
        Get #FileN, , stData
        data = stData
    Case vbBoolean
        Get #FileN, , bData
        data = CBool(bData)
    Case vbByte
        Get #FileN, , byData
        data = CByte(byData)
    Case vbCurrency
        Get #FileN, , cData
        data = CCur(cData)
    End Select
    
    If data = "-1,#IND" Then
        GoTo ERR
    End If
    
End If
FRead = True
Exit Function
ERR:
data = 0
FRead = False
End Function

''''''''''''''''''''''''''''''

Public Function FReadList(Count As Integer) As Variant()
On Error GoTo ERR

Dim i As Integer
Dim flen As Long
Dim buf As String
FRead i, FSeek
FRead flen, FSeek
buf = Space$(flen)
FRead buf, FSeek

Dim Section As zabArray
Set Section = New zabArray
Section.SetSerial buf

buf = ""

Dim n As Integer

Dim arrBuf() As Variant

For n = 0 To Count - 1 Step 1
    ReDim Preserve arrBuf(n)
    arrBuf(n) = Section.GetValue("c" & n)
Next

Set Section = Nothing

FReadList = arrBuf
Exit Function
ERR:
End Function

Public Function FWriteList(ParamArray datalist() As Variant) As Boolean
Dim Param As Variant
Dim Section As zabArray
On Error GoTo ERR
Set Section = New zabArray
Dim i As Integer
For Each Param In datalist
    Section.SetValue "c" & i, Param
    i = i + 1
Next Param
Dim s As String
s = Section.GetSerial
FWrite CInt(FSeek), FSeek
FWrite CLng(Len(s)), FSeek
FWrite CStr(s), FSeek
Set Section = Nothing
FWriteList = True
Exit Function
ERR:
FWriteList = False
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''
'
'   FWriteArray
'
'   Dim MyArray(1 To 10, 5 To 15, 10 To 20)  'объявляем массив
'   Dim AnyArray(10)
'   Upper = UBound(MyArray, 1) 'возвращает 10
'   Upper = UBound(MyArray, 3) 'возвращает 20
'   Upper = UBound(AnyArray) 'возвращает 10
'
'   ArraySize = 3
'
''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FWriteArray(ByRef data, Optional ArraySize As Long = 1)
Dim i As Long, n As Long
Dim ii As Long, iii As Long, iiii As Long
Dim Size As Integer
Dim cType As Long
      
    Select Case ArraySize
    Case 0: Case 1
        
        Size = LBound(data)
        FWrite Size, FSeek
        Size = UBound(data)
        FWrite Size, FSeek

        For i = LBound(data) To UBound(data)
            
            cType = VarType(data(i))
            FWrite cType
            
            If cType <> vbString Then
                FWrite data(i), FSeek
            Else
                FWriteString data(i), vLong
            End If
            
        Next
    
    Case 2
    
        Size = LBound(data, 1)
        FWrite Size, FSeek
        Size = UBound(data, 1)
        FWrite Size, FSeek
        
        For i = LBound(data, 1) To UBound(data, 1)
        
            Size = LBound(data, 2)
            FWrite Size, FSeek
            Size = UBound(data, 2)
            FWrite Size, FSeek

            For ii = LBound(data, 2) To UBound(data, 2)
                    
                cType = VarType(data(i, ii))
                FWrite cType
                
                If cType <> vbString Then
                    FWrite data(i, ii), FSeek
                Else
                    FWriteString data(i, ii), vLong
                End If
                    
            Next
            
        Next
        
    Case 3
    
        Size = LBound(data, 1)
        FWrite Size, FSeek
        Size = UBound(data, 1)
        FWrite Size, FSeek
    
        For i = LBound(data, 1) To UBound(data, 1)
        
            Size = LBound(data, 2)
            FWrite Size, FSeek
            Size = UBound(data, 2)
            FWrite Size, FSeek
        
            For ii = LBound(data, 2) To UBound(data, 2)
                    
                Size = LBound(data, 3)
                FWrite Size, FSeek
                Size = UBound(data, 3)
                FWrite Size, FSeek
    
                For iii = LBound(data, 3) To UBound(data, 3)
                
                    cType = VarType(data(i, ii, iii))
                    FWrite cType
                        
                    If cType <> vbString Then
                        FWrite data(i, ii, iii), FSeek
                    Else
                        FWriteString data(i, ii, iii), vLong
                    End If
                        
                Next
                
            Next
            
        Next
        
    Case 4
        
        Size = LBound(data, 1)
        FWrite Size, FSeek
        Size = UBound(data, 1)
        FWrite Size, FSeek
    
        For i = LBound(data, 1) To UBound(data, 1)
        
            Size = LBound(data, 2)
            FWrite Size, FSeek
            Size = UBound(data, 2)
            FWrite Size, FSeek
        
            For ii = LBound(data, 2) To UBound(data, 2)
                    
                Size = LBound(data, 3)
                FWrite Size, FSeek
                Size = UBound(data, 3)
                FWrite Size, FSeek
    
                For iii = LBound(data, 3) To UBound(data, 3)
                        
                    Size = LBound(data, 4)
                    FWrite Size, FSeek
                    Size = UBound(data, 4)
                    FWrite Size, FSeek
        
                    For iiii = LBound(data, 4) To UBound(data, 4)
                    
                        cType = VarType(data(i, ii, iii, iiii))
                        FWrite cType
                            
                        If cType <> vbString Then
                            FWrite data(i, ii, iii, iiii), FSeek
                        Else
                            FWriteString data(i, ii, iii, iiii), vLong
                        End If
                            
                    Next
                        
                Next
                
            Next
            
        Next
        
    End Select

End Function

''''

Public Function FReadArray(ByRef data, Optional ArraySize As Long = 1)
Dim i As Long, n As Long
Dim ii As Long, iii As Long, iiii As Long

Dim from As Integer, toend As Integer
Dim from1 As Integer, toend1 As Integer
Dim from2 As Integer, toend2 As Integer
Dim from3 As Integer, toend3 As Integer
Dim sTmp As String

Dim cType As Long

If FLOF() = 0 Then Exit Function

    Select Case ArraySize
    Case 0: Case 1
    
        FRead from, FSeek
        FRead toend, FSeek

        ReDim Preserve data(from To toend)

        For i = from To toend
            
            FRead cType
            data(i) = FReadType(cType)
'            FRead Data(i), FSeek

        Next
    
    Case 2
    
        FRead from, FSeek
        FRead toend, FSeek
        
        For i = from To toend
        
            FRead from1, FSeek
            FRead toend1, FSeek
            
            If GetArraySize(data) = 0 Then
                ReDim Preserve data(from To toend, from1 To toend1)
            End If

            For ii = from1 To toend1
                    
                FRead cType
                data(i, ii) = FReadType(cType)
'                FRead Data(i, ii), FSeek
                    
            Next
            
        Next
        
    Case 3
    
        FRead from, FSeek
        FRead toend, FSeek
    
        For i = from To toend
        
            FRead from1, FSeek
            FRead toend1, FSeek
        
            For ii = from1 To toend1
            
                FRead from2, FSeek
                FRead toend2, FSeek
                
                If GetArraySize(data) = 0 Then
                    ReDim Preserve data(from To toend, from1 To toend1, from2 To toend2)
                End If
    
                For iii = from2 To toend2
                        
                    FRead cType
                    data(i, ii, iii) = FReadType(cType)
'                    FRead Data(i, ii, iii), FSeek

                Next
                
            Next
            
        Next
        
    Case 4
    
        FRead from, FSeek
        FRead toend, FSeek
    
        For i = from To toend
        
            FRead from1, FSeek
            FRead toend1, FSeek
        
            For ii = from1 To toend1
            
                FRead from2, FSeek
                FRead toend2, FSeek
    
                For iii = from2 To toend2
                        
                    FRead from3, FSeek
                    FRead toend3, FSeek
            
                    If GetArraySize(data) = 0 Then
                        ReDim Preserve data(from To toend, from1 To toend1, from2 To toend2, from3 To toend3)
                    End If
            
                    For iiii = from To toend
            
                        FRead cType
                        data(i, ii, iii, iiii) = FReadType(cType)
'                        FRead Data(i, ii, iii, iiii), FSeek
            
                    Next

                Next
                
            Next
            
        Next
    End Select

End Function

Private Function FReadType(cType As Long) As Variant
Dim cI As Integer
Dim cL As Long
Dim cSi As Single
Dim cS As String
Dim cD As Double
Dim cbl As Boolean
Dim cB As Byte

Select Case cType
    Case vbInteger
        FRead cI, FSeek
        FReadType = cI
    Case vbLong
        FRead cL, FSeek
        FReadType = cL
    Case vbSingle
        FRead cSi, FSeek
        FReadType = cSi
    Case vbDouble
        FRead cD, FSeek
        FReadType = cD
    Case vbCurrency
'        GoTo ERR
    Case vbString
        FReadString cS, vLong
        FReadType = cS
    Case vbBoolean
        FRead cbl, FSeek
        FReadType = cbl
    Case vbByte
        FRead cB, FSeek
        FReadType = cB
    Case vbDate
'        GoTo ERR
End Select
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetArraySize(ByRef a) As Integer
    On Error GoTo ERR
    GetArraySize = UBound(a) + 1
    Exit Function
ERR:
    GetArraySize = 0
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FReadAll() As String
Dim i As Long
On Error GoTo ERR
i = FLOF()
If i > 0 Then
    FReadAll = Space$(i)
    Get #FileN, , FReadAll
End If
Exit Function
ERR:
End Function


Public Function FBReadAll() As Byte()
Dim buf() As Byte
Dim i As Long
On Error GoTo ERR
i = FLOF()
If i > 0 Then
    ReDim buf(0 To i - 1)
    Get #FileN, , buf()
End If
'Exit Function
ERR:
FBReadAll = buf
End Function


'''''''''''''''''''''''''''''''''

Public Function FWriteString(ByVal data As String, Optional ByVal vf As vFormat = vSingle) As Boolean
On Error GoTo ERR
Dim sdata As String
    Select Case VarType(data)
    Case vbString
        sdata = data
    Case Else
        sdata = CStr(data)
    End Select
    Select Case vf
    Case vSingle
        FWrite CSng(Len(sdata)), FSeek
    Case vLong
        FWrite CLng(Len(sdata)), FSeek
    End Select
    FWrite sdata, FSeek
FWriteString = True
Exit Function
ERR:
FWriteString = False
End Function


Public Function FReadString(ByRef sbuf As String, Optional ByVal vf As vFormat = vSingle) As Boolean
Dim slen1 As Long
Dim slen2 As Single
On Error GoTo ERR
Select Case vf
Case vSingle
    FRead slen2, FSeek
    sbuf = Space$(slen2)
Case vLong
    FRead slen1, FSeek
    sbuf = Space$(slen1)
End Select
FRead sbuf, FSeek, Len(sbuf)
FReadString = True
Exit Function
ERR:
FReadString = False
End Function


''''''''''''''''''''''''''''''''''''''''


Public Function FSeek(Optional position As Long = -1) As Long
On Error Resume Next
If position = 0 Then position = 1
If position > -1 Then
' Устанавливаем
    Seek #FileN, position
    FSeek = Seek(FileN)
Else
' Берем данные
    FSeek = Seek(FileN)
End If
End Function

Public Function FLOF() As Long
On Error Resume Next
FLOF = LOF(FileN)
End Function


Public Function FEOF() As Integer
On Error Resume Next
FEOF = EOF(FileN)
End Function

Public Function FClose(Optional FN As Integer) As Boolean
On Error Resume Next
If FileStatus = 1 Or FileStatus = 2 Then
FWrite CSng(cVer)
FWriteString "iStorage by MDinc @2011", vSingle
'SystemInfo = FSeek
End If
If FN = 0 Then FN = FileN
Close FN
FileStatus = 0
FileN = 0
End Function

Public Function FileExists(ByRef Path As String) As Boolean
On Error Resume Next
FileExists = Not (GetFileAttributes(Path) = -1)
End Function

Public Function FileAttributes(ByRef strPath As String) As VbFileAttribute
On Error Resume Next
FileAttributes = GetFileAttributes(strPath)

'attr = GetFileAttributes(strPath)
'If (attr And &H1) = &H1 Then
'Val = " 'Только для чтения',"
'End If
'If (attr And &H2) = &H2 Then
'Val = Val & " 'Скрытый',"
'End If
'If (attr And &H4) = &H4 Then
'Val = Val & " 'Системный',"
'End If
'If (attr And &H20) = &H20 Then
'Val = Val & " 'Архивный',"
'End If
'If (attr And &H80) = &H80 Then
'Val = Val & " 'Нормальный',"
'End If
'If (attr And &H800) = &H800 Then
'Val = Val & " 'Компрессированный',"
'End If
'Val = Left(Val, Len(Val) - 1)
'If (attr And &H10) = &H10 Then
'MsgBox "Данная папка имеет атрибуты " & Val
'Else
'MsgBox "Данный файл имеет атрибуты " & Val
'End If

End Function

Public Property Get FN() As Integer
    FN = FileN
End Property

Public Property Get FS() As Integer
    FS = FileStatus
End Property

'
'    Binary File Manager
'
'    Dim cf As FileMan.clsFile
'    Set cf = New clsFile
'
'    cf.FOpen App.Path & "\test.bin", aWrite
'    cf.FWriteString "Test of data string save for binary file", vSingle
'    cf.FWrite CSng(13)
'    cf.FWriteString "", 1
'    cf.FWrite ScaleLeft_Main
'    cf.FWrite ScaleWidth_Main
'    cf.FWrite ScaleTop_Main
'    cf.FWrite ScaleHeight_Main
'    cf.FWrite CLng(1564455)
'    cf.FWrite CInt(1564)
'    cf.FWrite CDbl(451864641867#)
'    cf.FClose
'
'    Dim f0 As Byte
'    Dim f1 As Long
'    Dim f2 As Integer
'    Dim f3 As Single, f7 As Single
'    Dim f5 As Double
'    Dim f6 As Boolean
'    Dim f8 As String ', f8 As String
'
'    Dim flist(), flist1() As Variant
'
'    cf.FOpen App.Path & "\test.bin", aRead
'    cf.FReadString f8, vSingle
'    cf.FRead f3
'    cf.FReadString f8, 1
'    cf.FRead ScaleLeft_Main
'    cf.FRead ScaleWidth_Main
'    cf.FRead ScaleTop_Main
'    cf.FRead ScaleHeight_Main
'    cf.FRead f1
'    cf.FRead f2
'    cf.FRead f5
'    cf.FClose
