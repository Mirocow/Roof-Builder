Attribute VB_Name = "mChrArr"
Option Explicit

'Memory manipulating APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Heap manipulating APIs
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

'API to get the pointer of an array
Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (lpObject() As Any) As Long

Public Sub OpenChrArr(ByRef intChars() As Integer, ByRef strCopy As String)
    '------------------------------------------------------------------
    'Purpose:   To create an array which references the same data
    '           as the specified string which allows you to view
    '           all the character codes of each character in the string
    '           as well as modify them. This is done by allocating
    '           memory to create our own safe array header containing
    '           the properties of the string and make the array point
    '           to it.
    '
    'Params:
    '           intChars():     Reference to array to modify
    '
    '           strCopy:        Reference to string make array point to
    '                           If it is a null string, then the array
    '                           does not point to anything (will have
    '                           to be changed with UpdateChrArr)
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    'Format of any safe array header (each x = 1 byte)
    '
    ' x x           Dimensions
    ' x x           Flags (default is 128)
    ' x x x x       Numbers of times array has been locked without being unlocked
    ' x x x x       Pointer to array data
    '
    ' x x x x       Number of elements of first dimension
    ' x x x x       LBound of first dimension
    '
    ' *** Number of blocks like this depends on the number of dimensions
    '     however in this case, we will ever only have one dim in the array
    '
    '
    'Format of our safe array header
    '
    ' 1   0
    ' 128 0
    ' 0   0    0    0
    ' StrPtr of strCopy
    ' Len of strCopy
    ' 0   0    0    0
    '------------------------------------------------------------------

44:    Dim lngArrPtr       As Long
    
    'Allocate 24 bytes of memory for safe array header (the 8 in
    'the dwFlags param makes all bytes have the intial value 0)
48:    lngArrPtr = HeapAlloc(GetProcessHeap, 8, 24)
    
    'Make the array header pointer in intChars point to our allocated memory
51:    CopyMemory ByVal ArrPtr(intChars), lngArrPtr, 4
    
    'Number of dimensions (1)
54:    FillMemory ByVal lngArrPtr, 1, 1
    
    'Flags for array (default is 128 for integer arrays produced by VB)
57:    FillMemory ByVal lngArrPtr + 2, 1, 128
    
    'Size in bytes of each array element (2)
60:    FillMemory ByVal lngArrPtr + 4, 1, 2
    
    'Pointer to array data (in other words, pointer to string characters)
63:    CopyMemory ByVal lngArrPtr + 12, StrPtr(strCopy), 4
    
    'Number of elements in array (ie number of characters in string)
66:    CopyMemory ByVal lngArrPtr + 16, CLng(Len(strCopy)), 4
End Sub

Public Sub UpdateChrArr(ByRef intChars() As Integer, ByRef strCopy As String)
    '------------------------------------------------------------------
    'Purpose:   To change an array already pointing to string data so
    '           that it points to a new string (array in question
    '           should have been prepared via the OpenChrArr method)
    '
    'Params:
    '           intChars():     Array to modify
    '
    '           strCopy:        Reference to new string to point to
    '                           If it is a null string, then it will
    '                           make the array point to nothing
    '------------------------------------------------------------------

14:    Dim lngArrPtr       As Long
    
    'Get pointer to array header
17:    CopyMemory lngArrPtr, ByVal ArrPtr(intChars), 4
    
    'Update reference to data area (ie new StrPtr so replace old)
20:    CopyMemory ByVal lngArrPtr + 12, StrPtr(strCopy), 4
    
    'Update number of elements in array (ie new Len so replace old)
23:    CopyMemory ByVal lngArrPtr + 16, CLng(Len(strCopy)), 4
End Sub

Public Sub CloseChrArr(ByRef intChars() As Integer)
    '------------------------------------------------------------------
    'Purpose:   To do the clean up work for arrays created via the
    '           OpenChrArr method. It deletes the arrays reference to
    '           the allocated memory header and frees the memory itself
    '
    'Params:
    '           intChars():     The array to clean up
    '------------------------------------------------------------------

10:    Dim lngAlloc        As Long
11:    Dim lngArrPtr       As Long

    'Get pointer to pointer of array header (our allocated data)
14:    lngArrPtr = ArrPtr(intChars)

    'Get pointer to array header (allocated memory)
17:    CopyMemory lngAlloc, ByVal lngArrPtr, 4
    
    'Clear reference to allocated memory (effectively putting the
    'array back in the state we found it in)
21:    ZeroMemory ByVal lngArrPtr, 4
    
    'Free the allocated memory
24:    HeapFree GetProcessHeap, 0, ByVal lngAlloc
End Sub
