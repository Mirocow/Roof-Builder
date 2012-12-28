Attribute VB_Name = "Draw"
Public Type POINT
  X As Long
  Y As Long
End Type

Public Type LOGPEN
  lopnStyle As Long
  lopnWidth As Long
  lopnColor As Long
End Type

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function Set_Pixel Lib "gdi32" Alias "SetPixel" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal color As Long) As Long



' full version of APILine
Public Sub APILineEx(Pic As PictureBox, _
                      p1 As POINT, _
                      p2 As POINT, _
                      L As LOGPEN)
  'Use the API LineTo for Fast Drawing
  On Error GoTo APILineEx_Error
  
  'Pic.ScaleMode = vbPixels
  
  Dim mX1 As Long
  Dim mY1 As Long
  Dim mX2 As Long
  Dim mY2 As Long

  mX1 = Pic.ScaleX(p1.X, Pic.ScaleMode, vbPixels) - Pic.ScaleX(Pic.ScaleLeft, Pic.ScaleMode, vbPixels)
  mY1 = Pic.ScaleY(p1.Y, Pic.ScaleMode, vbPixels) - Pic.ScaleY(Pic.ScaleTop, Pic.ScaleMode, vbPixels)
  mX2 = Pic.ScaleX(p2.X, Pic.ScaleMode, vbPixels) - Pic.ScaleX(Pic.ScaleLeft, Pic.ScaleMode, vbPixels)
  mY2 = Pic.ScaleY(p2.Y, Pic.ScaleMode, vbPixels) - Pic.ScaleY(Pic.ScaleTop, Pic.ScaleMode, vbPixels)

  Dim P As POINT
  P.X = 0 'CLng(mX1)
  P.Y = 0 'CLng(mY1)
  
  Dim hPen As Long, hPenOld As Long
  hPen = CreatePen(L.lopnStyle, L.lopnWidth, L.lopnColor)
  hPenOld = SelectObject(Pic.hdc, hPen)
  
  MoveToEx Pic.hdc, CLng(mX1), CLng(mY1), P
  LineTo Pic.hdc, CLng(mX2), CLng(mY2)
  
  SelectObject Pic.hdc, hPenOld
  DeleteObject hPen
  
  Exit Sub

APILineEx_Error:
'OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.10." & ERR.Source & "]", ERR.Number, ERR.Description
End Sub

Public Sub SetPixel(hdc As Long, X As Long, Y As Long, color As Long)
Set_Pixel hdc, X, Y, color
End Sub

