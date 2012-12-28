Attribute VB_Name = "Safe"
'**************************************
' Name: Safe UBound and LBound
' Description:Ever wanted to use LBound and UBound to get arrays boundaries without jumping over error message when the array is empty? These functions will replace the ordinary LBound and UBound procedures so you don’t need to worry about errors. I've also included a way to get the dimensions of an array. Just paste the following code into a module, and the problem is solved.
' By: Kristian S. Stangeland
'
' Inputs:SafeUBound and SafeLBound: [Address to the array], [What dimension you want to obtain]
'ArrayDims: [Address to the array]
'
' Returns:As expected from the ordinary functions, except that they will return -1 when the array is empty.
'
' Assumes:You obtain the address to an array by passing it to the VarPtrArray API call. So if you want to get the boundaries of an array called aTmp, you need to call the functions like this:
'lLowBound = SafeLBound(VarPtrArray(aTmp))
'lHighBound = SafeUBound(VarPtrArray(aTmp))
'lDimensions = ArrayDims(VarPtrArray(aTmp))
'When dealing with string arrays that isn't allocated at design time, you *must* add the value 4 to the lpArray-paramenter:
'lLowBound = SafeLBound(VarPtrArray(aString) + 4)
'
' Side Effects:Since the return value is minus when the array is empty it's a big chance you will get problems with minus dimensioned arrays, but who use them anyway?
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=55074&lngWId=1'for details.'**************************************

Option Explicit
Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function SafeUBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long
Dim lAddress&, cElements&, lLbound&, cDims%
If Dimension < 1 Then
SafeUBound = -1
Exit Function
End If
CopyMemory lAddress, ByVal lpArray, 4
If lAddress = 0 Then
' The array isn't initilized
SafeUBound = -1
Exit Function
End If
' Calculate the dimensions
CopyMemory cDims, ByVal lAddress, 2
Dimension = cDims - Dimension + 1
' Obtain the needed data
CopyMemory cElements, ByVal (lAddress + 16 + ((Dimension - 1) * 8)), 4
CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4
SafeUBound = cElements + lLbound - 1
End Function

Public Function SafeLBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long
Dim lAddress&, cElements&, lLbound&, cDims%
If Dimension < 1 Then
SafeLBound = -1
Exit Function
End If
CopyMemory lAddress, ByVal lpArray, 4
If lAddress = 0 Then
' The array isn't initilized
SafeLBound = -1
Exit Function
End If
' Calculate the dimensions
CopyMemory cDims, ByVal lAddress, 2
Dimension = cDims - Dimension + 1
' Obtain the needed data
CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4
SafeLBound = lLbound
End Function

Public Function ArrayDims(ByVal lpArray As Long) As Integer
Dim lAddress As Long
CopyMemory lAddress, ByVal lpArray, 4
If lAddress = 0 Then
' The array isn't initilized
ArrayDims = -1
Exit Function
End If
CopyMemory ArrayDims, ByVal lAddress, 2
End Function
