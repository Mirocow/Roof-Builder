Attribute VB_Name = "CycleShift"
Option Explicit

Public Enum dcShiftDirection
    Left = -1
    Right = 0
End Enum

'=======================================
'     ==========
'Public Function Shift(ByVal lValue As L
'     ong, ByVal lNumberOfBitsToShift As Long,
'     ByVal lDirectionToShift As dcShiftDirect
'     ion) As Long
'Author: Donald Moore (MindRape)
'E-mail: moore@futureone.com
' Date: 06/16/99
'Enters:
' lValue as Long
' lNumberOfBitsToShift as Long
' lDirectionToShift as Long
'Returns:
' Long - shifted value
'Purpose:
' Shift the given value by the given num
'     ber of bits to shift in the given direct
'     ion.
' Shifting bits to the left acts as a mu
'     ltiplier and to the right divides.

Public Function Shift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long, ByVal lDirectionToShift As dcShiftDirection) As Long

    Const ksCallname As String = "Shift"
    On Error GoTo Procedure_Error
    Dim LShift As Long

    If lDirectionToShift Then 'shift left
        LShift = lValue * (2 ^ lNumberOfBitsToShift)
    Else 'shift right
        LShift = lValue \ (2 ^ lNumberOfBitsToShift)
    End If

    
Procedure_Exit:
    Shift = LShift
    Exit Function
    
Procedure_Error:
    ERR.Raise ERR.Number, ksCallname, ERR.Description, ERR.HelpFile, ERR.HelpContext
    Resume Procedure_Exit
End Function

'=======================================
'     ==========
'Public Function LShift(ByVal lValue As
'     Long, ByVal lNumberOfBitsToShift As Long
'     ) As Long
'Author: Donald Moore (MindRape)
'E-mail: moore@futureone.com
' Date: 06/16/99
'Enters:
' lValue as Long
' lNumberOfBitsToShift as Long
'Returns:
' Long - shifted value
'Purpose:
' Shift the given value by the given num
'     ber of bits left

Public Function LShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long

    Const ksCallname As String = "LShift"
    On Error GoTo Procedure_Error
    LShift = Shift(lValue, lNumberOfBitsToShift, Left)
    
Procedure_Exit:
    Exit Function
    
Procedure_Error:
    ERR.Raise ERR.Number, ksCallname, ERR.Description, ERR.HelpFile, ERR.HelpContext
    Resume Procedure_Exit
End Function

'=======================================
'     ==========
'Public Function RShift(ByVal lValue As
'     Long, ByVal lNumberOfBitsToShift As Long
'     ) As Long
'Author: Donald Moore (MindRape)
'E-mail: moore@futureone.com
' Date: 06/16/99
'Enters:
' lValue as Long
' lNumberOfBitsToShift as Long
'Returns:
' Long - shifted value
'Purpose:
' Shift the given value by the given num
'     ber of bits right

Public Function RShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long

    Const ksCallname As String = "RShift"
    On Error GoTo Procedure_Error
    RShift = Shift(lValue, lNumberOfBitsToShift, Right)
    
Procedure_Exit:
    Exit Function
    
Procedure_Error:
    ERR.Raise ERR.Number, ksCallname, ERR.Description, ERR.HelpFile, ERR.HelpContext
    Resume Procedure_Exit
End Function

' #define ROL(x, n) (((x) << (n)) | ((x) >> (32-(n))))
Public Function CycleShiftLeft(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long
CycleShiftLeft = LShift(lValue, lNumberOfBitsToShift) Or RShift(lValue, 32 - lNumberOfBitsToShift)
End Function

'#define ROL(x, n) (((x) << (n)) | ((x) >> (32-(n))))
Public Function CycleShiftRight(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long
CycleShiftLeft = RShift(lValue, lNumberOfBitsToShift) Or LShift(lValue, 32 - lNumberOfBitsToShift)
End Function
