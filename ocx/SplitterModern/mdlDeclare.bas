Attribute VB_Name = "mdlDeclare"
'***********************************
'�  MDinc dmms@narod.ru
'***********************************
Option Explicit

'Public Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

'Public Declare Function SetRect Lib "user32" _
'    (lpRect As RECT, _
'    ByVal X1 As Long, _
'    ByVal Y1 As Long, _
'    ByVal X2 As Long, _
'    ByVal Y2 As Long) _
'    As Long
    
Public Enum enoperation
    Shift = 1 ' ��������� ��������� �� �����
    Push = 2  ' �������� � ����� � ����
    Clear = 3 ' �������� ����
End Enum

Public Function Amount(ByRef arr, Optional operation As enoperation) As Integer
On Error GoTo Handler
On Error Resume Next
Select Case operation
Case Shift
    If UBound(arr) > 0 Then ReDim Preserve arr(UBound(arr) - 1)
Case Push
    If UBound(arr) = -1 Then
       ReDim arr(0)
    Else
       ReDim Preserve arr(UBound(arr) + 1)
    End If
Case Clear
    ReDim arr(0)
End Select

Amount = UBound(arr)

Exit Function
Handler:
    MsgBox UBound(arr)
End Function
