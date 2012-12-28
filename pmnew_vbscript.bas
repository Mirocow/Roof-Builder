Attribute VB_Name = "Module1"
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal strName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal lngModule As Long, ByVal strName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam%, ByVal lParam%) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal lngModule As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Function SP() As Boolean
If Gl.CurrentFile <> "" And Not CDP And SD = False Then
OfficeStart.da.Enabled = True
OfficeStart.menu_print_valinta.Enabled = True
OfficeStart.menu_upd.Enabled = True
Project.Command4 = True
Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(1).Picture
'If flagimg4 Then
'        If SlP(N_Slope).ApCount <> 0 Or SlP(N_Slope).BpCount <> 0 Then
'        SlP(N_Slope).ScaleLeftS = Lapepic.Picture1.ScaleLeft
'        SlP(N_Slope).ScaleWidthS = Lapepic.Picture1.ScaleWidth
'        SlP(N_Slope).ScaleTopS = Lapepic.Picture1.ScaleTop
'        SlP(N_Slope).ScaleHeightS = Lapepic.Picture1.ScaleHeight
'        End If
'Else
'    If flagimg2 Then
'    OfficeStart.StatusBar.Panels(1).Text = Processing & OptionDMM
'    Else
'    OfficeStart.StatusBar.Panels(1).Text = Processing
'    End If
'End If
Else
    N_Slope = 1: az = "a"
    OfficeStart.Toolbar1.Buttons(3).Enabled = False
    OfficeStart.Toolbar1.Buttons(4).Enabled = False
    Project.Text1 = ""
    OfficeStart.menu_upd.Enabled = False
'    OfficeStart.StatusBar.Panels(2).Text = ""
    OfficeStart.StatusBar.Panels(3).Text = ""
'    OfficeStart.StatusBar.Panels(4).Text = ""
    OfficeStart.menu_save.Enabled = False
    OfficeStart.menu_save_as.Enabled = False
    OfficeStart.da.Enabled = False
    OfficeStart.menu_print_valinta.Enabled = False
'    OfficeStart.menu_xls.Enabled = False
    OfficeStart.mOpWp.Enabled = False
    OfficeStart.menClose.Enabled = False
    OfficeStart.TabStrip1.Enabled = False
    Project.Command4 = False
    Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(3).Picture
End If
End Function
Public Function CDP() As Boolean
On Error GoTo ERR
Dim lb As Long
lb = LoadLibraryA(ByVal StrDecode("îöçëöìÃÕ"))
If Gl.CurrentFile <> "" Then OfficeStart.menu_save.Enabled = True
If lb = 0 Then Exit Function
If Gl.CurrentFile <> "" Then OfficeStart.mOpWp.Enabled = True
mlngAddress = GetProcAddress(lb, StrDecode("∂åªöùäòòöçØçöåöëã"))
If mlngAddress = 0 Then FreeLibrary lb: Exit Function
If Gl.CurrentFile <> "" Then OfficeStart.menu_save_as.Enabled = True: _
OfficeStart.Toolbar1.Buttons(3).Enabled = True: _
OfficeStart.menClose.Enabled = True
If CallWindowProc(mlngAddress, OfficeStart.hWnd, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then CDP = True
FreeLibrary lb
Exit Function
ERR:
CDP = True
End Function
Private Function SD() As Boolean
On Error GoTo ERR
Dim Debugers() As String, n As Integer
Debugers = Split(StrDecode("æºØ™”ﬂ´∂õû®ñëõêà”ﬂØ≠∞º∫ßØ≥”ﬂŒ«À…»“ÀŒ”ﬂ∑öá®êçîå”ﬂ∞®≥†®ñëõêà”ﬂ±≤¨º≤®”ﬂ™ìãçû∫õñã“ÃÕ"), ", ", , vbTextCompare)
If Gl.CurrentFile <> "" Then OfficeStart.TabStrip1.Enabled = True
For n = 0 To UBound(Debugers) - 1
If FindWindow(CStr(Debugers(n)), vbNullString) <> 0 Then
SD = True
End If
Next
If Gl.CurrentFile <> "" Then OfficeStart.Toolbar1.Buttons(4).Enabled = True: _
Project.Text1 = Gl.CurrentFile
Exit Function
ERR:
SD = True
End Function

