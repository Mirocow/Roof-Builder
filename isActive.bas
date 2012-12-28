Attribute VB_Name = "isActive"
Option Explicit

' This program demonstrates a simple method of detecting if the mouse pointer is over
' the active part of custom buttons.  The method is more reliable than using the mousemove
' event and is an improvement on usual mouseover methods in that it allows irregular shaped
' buttons to be used (detection does not fire outside of the button shape.
'
' Very effective custom buttons can be developed this way and by using Images rather than pictures
' you should save picture resources.
'
' Author: Phobos - 17/12/2001

' Type Declarations
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long


Public Function isFormFocus(hwnd As Long) As Boolean
Dim MOUSE As POINTAPI
GetCursorPos MOUSE
Dim lhWnd As Long
lhWnd = WindowFromPoint(MOUSE.x, MOUSE.y)
If hwnd = GetParent(lhWnd) Then
    isFormFocus = True
End If
End Function

Public Sub SetFormFocus(hwnd As Long)
    SetFocus hwnd
End Sub

'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'Public Function isFormFocus(obj_hwid As Long) As Boolean
'Dim test_hwnd As Long
'
'test_hwnd = GetFocus
'Do While test_hwnd <> 0
'    'Check if the window isn't a child
'    If GetParent(test_hwnd) = obj_hwid Then
'        'Get the window's thread
'        isFormFocus = True
'        Exit Do
'    End If
'    'retrieve the next window
'    test_hwnd = GetWindow(test_hwnd, 5)
'Loop
'End Function

Public Function IsAPPMouseOver() As Boolean
    Dim MOUSE As POINTAPI, TopLeft As POINTAPI
    Dim Right As Long
    Dim Bottom As Long

    ' Identify the position (in pixels) of the first available screen coordinate on the form.
    TopLeft.x = 0:   TopLeft.y = 0
    Call ClientToScreen(OfficeStart.hwnd, TopLeft)
    
    Right = TopLeft.x + OfficeStart.Width
    Bottom = TopLeft.y + OfficeStart.Height
            
    ' Now we know where the image is we need to find out if the mouse pointer is in the mask area.
    Call GetCursorPos(MOUSE)
    
    ' TOP
    If MOUSE.x > TopLeft.x And MOUSE.x < Right Then
        If MOUSE.y > TopLeft.y And MOUSE.y < Bottom Then
            IsAPPMouseOver = True
        End If
    End If
End Function
