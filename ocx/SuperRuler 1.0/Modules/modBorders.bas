Attribute VB_Name = "modBorders"
Option Explicit

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long


Private Const SED_OLDPROC = "SED_OLDPROC"
Private Const SED_OLDGWLSTYLE = "SED_OLDGWLSTYLE"
Private Const SED_OLDGWLEXSTYLE = "SED_OLDGWLEXSTYLE"
Private Const SED_BORDERS = "SED_BORDERS"

Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

Private Const WS_THICKFRAME = &H40000
Private Const WS_BORDER = &H800000
Private Const WS_EX_WINDOWEDGE = &H100&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000

Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const WM_NCPAINT = &H85

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Private Enum sedBorderWidth
    sbwNone
    sbwSingle
    sbwDouble
End Enum

Public Sub SetBorderStyle(ByVal hWnd As Long, ByVal eBorderStyle As enuBorderStyle)
    Dim lRet As Long
    
    lRet = GetProp(hWnd, SED_OLDPROC)

    If lRet <> 0 Then
        SetWindowLong hWnd, GWL_WNDPROC, lRet
    Else
        SetProp hWnd, SED_OLDGWLSTYLE, GetWindowLong(hWnd, GWL_STYLE)
        SetProp hWnd, SED_OLDGWLEXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE)
    End If
    
    pSetBorder hWnd, eBorderStyle
    
    lRet = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf pWindowProc)

    SetProp hWnd, SED_OLDPROC, lRet
    SetProp hWnd, SED_BORDERS, CLng(eBorderStyle)

    SetWindowPos hWnd, 0, 0, 0, 0, 0, _
        SWP_NOMOVE Or _
        SWP_NOSIZE Or _
        SWP_NOOWNERZORDER Or _
        SWP_NOZORDER Or _
        SWP_FRAMECHANGED

End Sub

Private Sub pSetBorder(ByVal hWnd As Long, ByVal eBorderStyle As enuBorderStyle)

    Dim pWidth As sedBorderWidth
    
    Select Case eBorderStyle
        Case bsNoBorder
            pWidth = sbwNone
        Case bsRaised
            pWidth = sbwDouble
        Case bsRaisedInner
            pWidth = sbwSingle
        Case bsSunken
            pWidth = sbwDouble
        Case bsSunkenOuter
            pWidth = sbwSingle
        Case bsEtched
            pWidth = sbwDouble
        Case bsBump
            pWidth = sbwDouble
    End Select
    
    Select Case pWidth
        Case sbwNone
            pWinStyleNeg hWnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hWnd, GWL_EXSTYLE, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
        Case sbwSingle
            pWinStyleNeg hWnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hWnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
            pWinStyleAdd hWnd, GWL_EXSTYLE, WS_EX_STATICEDGE
        Case sbwDouble
            pWinStyleNeg hWnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hWnd, GWL_EXSTYLE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
            pWinStyleAdd hWnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE
    End Select
    
    SetWindowPos hWnd, 0, 0, 0, 0, 0, _
        SWP_NOMOVE Or _
        SWP_NOSIZE Or _
        SWP_NOOWNERZORDER Or _
        SWP_NOZORDER Or _
        SWP_FRAMECHANGED
        
End Sub

Private Function pWindowProc( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
    Select Case uMsg
    
        Case WM_NCPAINT
            
            pDrawBorder hWnd, wParam, GetProp(hWnd, SED_BORDERS)
        
        Case Else
            pWindowProc = CallWindowProc(GetProp(hWnd, SED_OLDPROC), hWnd, uMsg, wParam, lParam)
            
    End Select
    
End Function

Private Sub pWinStyleAdd(ByVal hWnd As Long, ByVal lStyle As Long, ByVal lFlags As Long)
    
    SetWindowLong hWnd, lStyle, GetWindowLong(hWnd, lStyle) Or lFlags
    
End Sub

Private Sub pWinStyleNeg(ByVal hWnd As Long, ByVal lStyle As Long, ByVal lFlags As Long)
    
    SetWindowLong hWnd, lStyle, GetWindowLong(hWnd, lStyle) And Not lFlags
    
End Sub

Private Sub pDrawBorder(ByVal hWnd As Long, ByVal wParam As Long, ByVal lBorderType As enuBorderStyle)

    Dim lRet As Long
    Dim lMode As Long
    Dim hDC As Long
    Dim Rec As RECT
    
    If lBorderType = bsNoBorder Then Exit Sub
    
    hDC = GetWindowDC(hWnd)
    
    lRet = GetWindowRect(hWnd, Rec)
    
    Rec.Right = Rec.Right - Rec.Left
    Rec.Bottom = Rec.Bottom - Rec.Top
    Rec.Left = 0
    Rec.Top = 0

    lMode = 0
    Select Case lBorderType
        Case bsRaised
            lMode = BDR_RAISED
        Case bsRaisedInner
            lMode = BDR_RAISEDINNER
        Case bsSunken
            lMode = BDR_SUNKEN
        Case bsSunkenOuter
            lMode = BDR_SUNKENOUTER
        Case bsEtched
            lMode = BDR_SUNKENOUTER Or BDR_RAISEDINNER
        Case bsBump
            lMode = BDR_SUNKENINNER Or BDR_RAISEDOUTER
    End Select
    
    lRet = DrawEdge(hDC, Rec, lMode, BF_RECT)
    
    lRet = ReleaseDC(hWnd, hDC)

End Sub
