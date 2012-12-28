Attribute VB_Name = "Main"
'Declare Function VarPtr Lib "msvbvm50.dll" (Var As Any) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal ByteLen As Long)

Public Function SaveDataIntoFile(cm As IContainer, FileName As String) As Boolean
    Dim cf As clsFile
    Dim i As Integer
    Dim cItem As Variant
    
    On Error GoTo ERR
    
    Set cf = New clsFile
    If cf.FOpen(FileName, aWrite, , True) Then
    
        ' Сохраняем число элементов
        Dim Count As Long
        Count = cm.Count
        cf.FWrite Count
        
        For i = 1 To Count
            cItem = cm.Item(i)
            
            ' Сохраняем вначале тип переменной
            Dim cType As Long
            cType = VarType(cItem)
            cf.FWrite cType
            
            ' Сохраняем данные
            Select Case cType
                Case vbInteger
                    cf.FWrite CInt(cItem)
                Case vbLong
                    cf.FWrite CLng(cItem)
                Case vbSingle
                    cf.FWrite CSng(cItem)
                Case vbDouble
                    cf.FWrite CDbl(cItem)
                Case vbCurrency
                    GoTo ERR
                Case vbString
                    cf.FWriteString cItem, vLong
                Case vbBoolean
                    cf.FWrite CBool(cItem)
                Case vbByte
                    cf.FWrite CByte(cItem)
                Case vbDate
                    GoTo ERR
                Case vbEmpty
                    GoTo ERR
                Case vbObject
                    ' Это объект
                    GoTo ERR
                Case Is >= vbArray
                    ' Это массив
                    Dim ArrDimension As Long
                    ArrDimension = GetArrayDimension(cItem)
'                    cf.FWrite GetArrayType(cItem)
                    cf.FWrite ArrDimension
                    cf.FWriteArray cItem, ArrDimension
                Case Else
        
            End Select
        Next
        cf.FClose
        
    Else
        GoTo ERR
    End If
    
    Set cf = Nothing
    SaveDataIntoFile = True
    Exit Function
ERR:
SaveDataIntoFile = False
cf.FClose
Set cf = Nothing
End Function

Public Function LoadDataFromFile(cm As IContainer, FileName As String) As Boolean
    Dim cf As clsFile
    
    Dim cI As Integer
    Dim cL As Long
    Dim cSi As Single
    Dim cS As String
    Dim cD As Double
    Dim cbl As Boolean
    Dim cB As Byte
    
    On Error GoTo ERR
    
    Set cf = New clsFile
    
    If cf.FOpen(FileName, aRead) Then
    
        ' Читаем число элементов
        Dim Count As Long
        cf.FRead Count
        
        For i = 1 To Count
            
            ' Читаем вначале тип переменной
            Dim cType As Long
            cf.FRead cType
            
            ' Сохраняем данные
            Select Case cType
                Case vbInteger
                    cf.FRead cI
                    cm.Add cI, i
                Case vbLong
                    cf.FRead cL
                    cm.Add cL, i
                Case vbSingle
                    cf.FRead cSi
                    cm.Add cSi, i
                Case vbDouble
                    cf.FRead cD
                    cm.Add cD, i
                Case vbCurrency
                    GoTo ERR
                Case vbString
                    cf.FReadString cS, vLong
                    cm.Add cS, i
                Case vbBoolean
                    cf.FRead cbl
                    cm.Add cbl, i
                Case vbByte
                    cf.FRead cB
                    cm.Add cB, i
                Case vbDate
                    GoTo ERR
                Case vbEmpty
                    GoTo ERR
                Case vbObject
                    ' Это объект
                    GoTo ERR
                Case Is >= vbArray
                    ' Это массив
                    Dim cArr() As Variant
                    Dim ArrDimension As Long
                    Erase cArr
                    cf.FRead ArrDimension
                    cf.FReadArray cArr, ArrDimension
                    cm.Add cArr, i
                Case Else
        
            End Select
        Next

    cf.FClose
    End If
    
    Set cf = Nothing
    LoadDataFromFile = True
    Exit Function
ERR:
LoadDataFromFile = False
cf.FClose
Set cf = Nothing
End Function


'UBound(A, 1) 100
'UBound(A, 2) 3
'UBound(A, 3) 4
Public Function GetArrayDimension(arr As Variant) As Integer
On Error GoTo ERR
Dim i As Integer
Dim Size As Long
For i = 1 To 4
    Size = UBound(arr, i)
    GetArrayDimension = GetArrayDimension + 1
Next
Exit Function
ERR:
End Function

' lSize = UBound(a) - LBound(a) + 1

'Private Function GetArrayType(Arr As Variant) As Long
'GetArrayType = VarType(Arr)
'End Function

'Private Function InitArrayType(cType) As Variant()
'Select Case cType
'    Case vbInteger
'        Dim InitArrayType As Integer
'    Case vbLong
'        Dim InitArrayType As Long
'    Case vbSingle
'        Dim InitArrayType As Single
'    Case vbDouble
'        Dim InitArrayType As Double
'    Case vbCurrency
'        GoTo ERR
'    Case vbString
'        Dim InitArrayType As String
'    Case vbBoolean
'        Dim InitArrayType As Boolean
'    Case vbByte
'        Dim InitArrayType As Byte
'    Case vbDate
'        GoTo ERR
'End Select
'End Function

'Private Function Array_IsEqual(ByRef sArr1() As String, ByRef sArr2() As String) As Boolean
''Check if arrays are equal.
'   'Input: a dimensioned array to reference.
'   'Input: a dimensioned array to check.
'   'Return: true if the arrays are the same.
'
'   Dim lCnt1 As Long, lCnt2 As Long
'
'   'Set the return to true and prove otherwise.
'      Array_IsEqual = True
'
'   'Check if the boundries are the same.
'      If UBound(sArr1) - LBound(sArr1) = UBound(sArr2) - LBound(sArr2) Then
'         'Inititalize the counters.
'            lCnt1 = LBound(sArr1)
'            lCnt2 = LBound(sArr2)
'
'         'Loop through the arrays and compare elements.
'            Do
'               If sArr1(lCnt1) <> sArr2(lCnt2) Then
'                  'Element is not equal.
'                     Array_IsEqual = False
'                     Exit Do
'
'               End If
'
'               'Increment the counters.
'                  lCnt1 = lCnt1 + 1
'                  lCnt2 = lCnt2 + 1
'
'            Loop Until lCnt1 > UBound(sArr1) Or lCnt2 > UBound(sArr2)
'
'      Else
'         'Arrays are not the same.
'         Array_IsEqual = False
'         Exit Function
'      End If
'
'End Function

'Function ArrayStart(arr As Variant) As Long
'    Dim ptr As Long
'    Dim VType As Integer
'
'    Const VT_BYREF = &H4000&
'
'    ' get the real VarType of the argument
'    ' this is similar to VarType(), but returns also the VT_BYREF bit
'    CopyMemory VType, arr, 2
'
'    ' exit if not an array
'    If (VType And vbArray) = 0 Then Exit Function
'
'    ' get the address of the SAFEARRAY descriptor
'    ' this is stored in the second half of the
'    ' Variant parameter that has received the array
'    CopyMemory ptr, ByVal VarPtr(arr) + 8, 4
'
'    ' see whether the routine was passed a Variant
'    ' that contains an array, rather than directly an array
'    ' in the former case ptr already points to the SA structure.
'    ' Thanks to Monte Hansen for this fix
'
'    If (VType And VT_BYREF) Then
'        ' ptr is a pointer to a pointer
'        CopyMemory ptr, ByVal ptr, 4
'    End If
'
'    ' get the address of the SAFEARRAY structure
'    ' this is stored in the descriptor
'
'    ' get the first word of the SAFEARRAY structure
'    ' which holds the number of dimensions
'    ' ...but first check that saAddr is non-zero, otherwise
'    ' this routine bombs when the array is uninitialized
'    ' (Thanks to VB2TheMax aficionado Thomas Eyde for
'    '  suggesting this edit to the original routine.)
'    If ptr Then
'        CopyMemory ArrayStart, ByVal ptr + 12, 4   '***
'    End If
'End Function

'Public Sub CopyArray(AArray As Variant, BArray As Variant)
''copies B into A,  assumes A is larger than B
''also,  they are 1-based and have same # of dimensions,  can generalize in the future
'
'    Dim J As Integer
'    Dim D As Integer
'    Dim t As Long
'    Dim memLength As Long
'
'    D = GetDimensions(BArray)
'    For J = 1 To D
'        memLength = memLength + UBound(BArray, J)
'    Next J
'
'
'    'checks to see what type of array is being passed in
'    'and sets t = the length of that data type
'    'got the lengths by doing lenb(vbDouble) = 8
'    ' lenb(vbsingle) = lenb(vblong) = 4
'    'lenb(vbinteger) = 2
'    'the rest I dont want to deal with
'    t = VarType(BArray) - 8192
'    If t = 5 Then
'        t = 8
'    ElseIf t = 4 Or t = 3 Then
'        t = 4
'    ElseIf t = 2 Then
'    Else
'        MsgBox "Error,  Must be an array of numbers", , "Invalid Array to Copy"
'        Exit Sub
'    End If
'
'
'    memLength = memLength * t
'
'    CopyMemory VarPtr(AArray), ByVal VarPtr(BArray), memLength
'
'End Sub

'Option Explicit
''
''  SafeArray details lookup (Dr Memory)
''
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'  (Destination As Any, Source As Any, ByVal Length As Long)
'
'Private Type SafeArray
'   nDims    As Integer
'   junk1    As Integer
'   Size     As Long
'   junk2    As Long
'   DataPtr  As Long
'   Bounds(20) As Long   '  pairs of [N, LB], N = count,
'   End Type             '    LB = LBound, UBound = N + LB - 1
'
'Public Sub CheckArray(vArray As Variant)
'
'    Dim vTYpe As Integer, vInfo As SafeArray
'
'    vTYpe = VarType(vArray)
'    If vTYpe < vbArray Then
'       Debug.Print "Bad argument, type is " & vTYpe
'       Exit Sub
'       End If
'
'    Debug.Print "Argument is ";
'    Select Case vTYpe - vbArray
'         Case vbInteger:   Debug.Print "Integer array"
'         Case vbLong:      Debug.Print "Long array"
'         Case vbSingle:    Debug.Print "Single array"
'         Case vbByte::     Debug.Print "Byte array"
'         Case vbDouble:    Debug.Print "Double array"
'         Case vbString:    Debug.Print "String array"
'         Case Else:        Debug.Print "array of type " & vTYpe - vArray
'                    Exit Sub
'         End Select
'
'    Dim Bdescriptor As Long, Bsize As Long
'    Dim i&, j&
'    '
'    ' Two ways to get the SafeArray pointer
'    '   first, this is cute (saw it on the web)
'    '
'    CopyMemory vTYpe, vArray, 2   ' get the data (which is pointer to SafeArray)
'    CopyMemory vArray, vbLong, 2  ' by coercing the Variant to type Long
'    Bdescriptor = vArray          ' got it!
'    CopyMemory vArray, vTYpe, 2   ' restore the original type
'    If Bdescriptor = 0 Then Debug.Print "Unallocated array!": Exit Sub
'    '
'    ' But if you look at the pointer values you discover it's really
'    '   a lot easier than that:
'    '
'    Bdescriptor = VarPtr(vArray) + 16   ' Snap!
'
'    CopyMemory Bdescriptor, ByVal Bdescriptor, 4   ' de-reference it
'    CopyMemory vInfo, ByVal Bdescriptor, 2         ' get # of dims
'    CopyMemory vInfo, ByVal Bdescriptor, 16 + 8 * vInfo.nDims ' get # of dims
'
'    Debug.Print "Element size = "; vInfo.Size
'    Debug.Print "ArrayPtr     = "; vInfo.DataPtr
'    Debug.Print "Dim's        = " & vInfo.nDims
'
'    j = 0: Bsize = 1
'    For i = 1 To vInfo.nDims
'       Debug.Print "Dim "; i; " is  ("; vInfo.Bounds(j + 1); " to "; vInfo.Bounds(j) + vInfo.Bounds(j + 1) - 1; ")"
'       Bsize = Bsize * vInfo.Bounds(j)
'       j = j + 2
'       Next
'    Debug.Print "Total elements " & Bsize
'    Debug.Print "Total bytes  = " & Bsize * vInfo.Size
'End Sub

'Public Sub CopyArray(vArray As Variant, toArray As Variant)
'   '
'   ' by Dr Memory  (May 2005)
'   '
'    Dim vTYpe As Integer, vInfo As SafeArray
'
'    vTYpe = VarType(vArray)
'    If vTYpe < vbArray Then
'       Debug.Print "Bad argument, type is " & vTYpe
'       Exit Sub
'       End If
'
'
'    Dim Bdescriptor As Long, Bsize As Long
'    Dim i&, j&
'
'    Bdescriptor = VarPtr(vArray) + 16   ' Snap!
'
'    CopyMemory Bdescriptor, ByVal Bdescriptor, 4   ' de-reference it
'    CopyMemory vInfo, ByVal Bdescriptor, 2         ' get # of dims
'    CopyMemory vInfo, ByVal Bdescriptor, 16 + 8 * vInfo.nDims ' get # of dims
'
'    j = 0: Bsize = 1
'    For i = 1 To vInfo.nDims
'       Bsize = Bsize * vInfo.Bounds(j)
'       j = j + 2
'       Next
'
'    ReDim toArray(Bsize - 1) As Byte
'    CopyMemory toArray(0), ByVal vInfo.DataPtr, Bsize ' that's it!
'    End Sub
