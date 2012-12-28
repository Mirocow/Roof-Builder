Attribute VB_Name = "ModSubClass"
Option Explicit

Private Const WM_MOUSEMOVE = &H200
Private Const WM_MOUSELEAVE = &H2A3
Private Const WM_MOUSEDOWN = &H201
Private Const WM_MOUSEUP = &H202
Private Const WM_MOUSEDBL = &H203
Private Const GWL_WNDPROC = (-4)

Public Const WM_UP_SCROLL = 0
Public Const WM_DOWN_SCROLL = 1
Public Const BAR_MIN_HEIGHT = 135

'Public Up_ As Object
'Public UpHot_ As Object
'Public UpPres_ As Object
'Public Down_ As Object
'Public DownHot_ As Object
'Public DownPres_ As Object

'Public Bar_ As Object
'Public BarHot_ As Object
'Public BarPres_ As Object

'Public BarMidColor
Public MouseOver As Boolean

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TRACKMOUSEEVENT Lib "comctl32.dll" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Enum TrackMouseEventFlags
    TME_HOVER = 1&
    TME_LEAVE = 2&
    TME_NONCLIENT = &H10&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As TrackMouseEventFlags
    dwHoverTime As Long
End Type

Private Mstyle As Integer

Public Sub Subclass(Pic As PictureBox, PropName As Long, style As Integer)
Dim PrevProc As Long, hWnd&

    Mstyle = style

    hWnd = Pic.hWnd

    If GetProp(hWnd, "ExWndProcPtr") <> 0 Then Exit Sub

    PrevProc = GetWindowLong(hWnd, GWL_WNDPROC)
    SetProp hWnd, "ExWndProcPtr", PrevProc
    SetProp hWnd, "ExObjPtr", ObjPtr(Pic)
    SetProp hWnd, "Tracking", 0
    SetProp hWnd, "Name", PropName

    SetWindowLong hWnd, GWL_WNDPROC, AddressOf PicWndProc
End Sub

Public Function PicWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Pic As PictureBox
Dim lProc As Long, pObject&
Dim lPtr As Long
Dim RC As RECT, PT As POINTAPI
Dim ET As TRACKMOUSEEVENT
    
    lProc = GetProp(hWnd, "ExWndProcPtr")

    If lProc = 0 Then Exit Function
    
    
    

    Select Case uMsg
    Case WM_MOUSEMOVE
        'If wParam = 1 Then
        '    GetCursorPos PT
        '    ScreenToClient hWnd, PT
        '    If PT.X < 0 Or PT.Y < 0 Or PT.X > 15 Or PT.Y > 15 Then
        '        uMsg = WM_MOUSEUP
        '        MouseOver = False
        '        GoTo MOUSEOUT
        '    Else
        '        'uMsg = WM_MOUSEDOWN
        '        MouseOver = True
        '        GoTo MOUSEIN
        '    End If
        'End If
        If GetProp(hWnd, "Tracking") = 0 Then
            ET.cbSize = Len(ET)
            ET.hwndTrack = hWnd
            ET.dwFlags = TrackMouseEventFlags.TME_LEAVE Or TrackMouseEventFlags.TME_HOVER
            TRACKMOUSEEVENT ET
            Call SetProp(hWnd, "Tracking", 1)
            pObject = GetProp(hWnd, "ExObjPtr")
            
            If pObject <> 0 Then
                CopyMemory Pic, pObject, 4
                
                Select Case GetProp(hWnd, "Name")
                Case 0
                
                    If Mstyle = 1 Then
                        Set Pic.Picture = LoadResPicture(102, 0)
                    Else
                        Set Pic.Picture = LoadResPicture(104, 0)
                    End If
                
                    Pic.MousePointer = 99
                    Pic.MouseIcon = LoadResPicture(101, 2)
                    
                End Select
                    
                CopyMemory Pic, 0&, 4
            End If
        End If
    Case WM_MOUSELEAVE
'MOUSEOUT:
        pObject = GetProp(hWnd, "ExObjPtr")
        Call SetProp(hWnd, "Tracking", 0)

        If pObject Then
            CopyMemory Pic, pObject, 4

            Select Case GetProp(hWnd, "Name")
            Case 0
            
                If Mstyle = 1 Then
                    Set Pic.Picture = LoadResPicture(101, 0)
                Else
                    Set Pic.Picture = LoadResPicture(103, 0)
                End If

                Pic.MousePointer = 0
                
            End Select
            
            CopyMemory Pic, 0&, 4
        End If
'    Case WM_MOUSEDOWN, WM_MOUSEDBL
''MOUSEIN:
'        pObject = GetProp(hWnd, "ExObjPtr")
'
'        If pObject Then
'            CopyMemory Pic, pObject, 4
'
'            Select Case GetProp(hWnd, "Name")
'            Case 0
''                Set Pic.Picture = BarPres_
'            End Select
'
'            CopyMemory Pic, 0&, 4
'        End If
    Case WM_MOUSEUP
        pObject = GetProp(hWnd, "ExObjPtr")
        Call SetProp(hWnd, "Tracking", 0)
    Case Else
        'If uMsg <> 15 And uMsg <> 20 And uMsg <> 132 And uMsg <> 3 And uMsg <> 5 And _
            uMsg <> 70 And uMsg <> 71 And uMsg <> 131 And uMsg <> 133 And uMsg <> 32 And _
            uMsg <> 673 And uMsg <> 4110 Then Form1.List1.AddItem uMsg
    End Select
    
    PicWndProc = CallWindowProc(lProc, hWnd, uMsg, wParam, lParam)
End Function

Public Sub UnSubclass(Pic As PictureBox)
Dim ET As TRACKMOUSEEVENT
Dim lProc As Long, hWnd&, lretTrack&
      
    hWnd = Pic.hWnd
    
    lProc = GetProp(hWnd, "ExWndProcPtr")

    If lProc = 0 Then Exit Sub
    
    ET.cbSize = Len(ET)
    ET.hwndTrack = hWnd
    
    ET.dwFlags = TME_LEAVE Or TME_QUERY
    lretTrack = TRACKMOUSEEVENT(ET)

    If (ET.dwFlags And TME_LEAVE) = TME_LEAVE Then
        ET.dwFlags = TrackMouseEventFlags.TME_LEAVE Or TrackMouseEventFlags.TME_CANCEL
        lretTrack = TRACKMOUSEEVENT(ET)
    End If

    ET.dwFlags = TrackMouseEventFlags.TME_HOVER Or TrackMouseEventFlags.TME_QUERY
    lretTrack = TRACKMOUSEEVENT(ET)

    If (ET.dwFlags And TME_HOVER) = TME_HOVER Then
        ET.dwFlags = TrackMouseEventFlags.TME_HOVER Or TrackMouseEventFlags.TME_CANCEL
        lretTrack = TRACKMOUSEEVENT(ET)
    End If

    SetWindowLong hWnd, GWL_WNDPROC, lProc
    RemoveProp hWnd, "ExWndProcPtr"
    RemoveProp hWnd, "ExTendWndPtr"
    RemoveProp hWnd, "Tracking"
    RemoveProp hWnd, "Name"
End Sub

'Paint middle bar
'Public Function PaintBar(Bar As PictureBox)
'Dim X As Integer, Y As Integer
'Dim Colr As Long
'Dim StrtPt As Integer
'
'    Bar.Cls
'
'    If Bar.Height > BAR_MIN_HEIGHT * 2 Then
'        StrtPt = (Bar.Height / 2) - 60
'
'       For Y = StrtPt To StrtPt + 105 Step 30
'            For X = 60 To 135 Step 15
'                Bar.Line (X, Y)-(X, Y), vbWhite, BF
'            Next X
'        Next Y
'        For Y = StrtPt + 15 To StrtPt + 120 Step 30
'            For X = 75 To 150 Step 15
'                Colr = BarMidColor((X - 75) / 15)
'                Bar.Line (X, Y)-(X, Y), Colr, BF
'            Next X
'        Next Y
'    End If
'End Function
