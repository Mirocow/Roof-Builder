Attribute VB_Name = "DopOWN"
'Public Function OpenOWN(OpenFile As String)
'Dim FileN As Integer
'Dim Shifr As String
'
' FileN = FreeFile
'  Open OpenFile For Binary As #FileN
'  If LOF(FileN) = 0 Then
'  MsgBox ("File" + OpenFile + " is corrupted")
'  End If
'
'Shifr = Input$(LOF(FileN), FileN)
'Close #FileN
'roofcalc_ownDeshifrator Shifr
'
'End Function
'Dim PVer As Long
'Dim Properties As String * 2

Public Function roofcalc_ownDeshifrator(PreobrazShifr As String, ver As Long) As String
Dim l01D2 As Long
Dim l01D4
Dim l01D6 As Single
Dim l01D8 As String
Dim i As Integer
Dim n As Integer
'Dim temp As String
  
'  About.process = About.process & vbcrlf & "Получение служебной информации; "
  n = 1
'  ver = Val(Mid$(PreobrazShifr$, 3, 1)): n = 6
  If Left$(PreobrazShifr$, 2) = "RB" Then ver = Val(Mid$(PreobrazShifr$, 6, 1)): n = 6 '6
  
  For l01D2 = n To Len(PreobrazShifr$) - 2 Step 3

    For l01D4 = 0 To 2 Step 1
      l01D6 = Asc(Mid$(PreobrazShifr$, l01D2 + l01D4, 1)) - (73 + l01D4)
      If l01D6 < 0 Then l01D6 = l01D6 + 256
      
'     temp = temp & l01D6
      
      Mid$(PreobrazShifr$, l01D2 + l01D4, 1) = Chr$(l01D6)
    Next l01D4
    l01D8$ = Mid$(PreobrazShifr$, l01D2 + 2, 1)
    Mid$(PreobrazShifr$, l01D2 + 2, 1) = Mid$(PreobrazShifr$, l01D2 + 1, 1)
    Mid$(PreobrazShifr$, l01D2 + 1, 1) = Mid$(PreobrazShifr$, l01D2, 1)
    Mid$(PreobrazShifr$, l01D2, 1) = l01D8$
  Next l01D2

'About.process = About.process & vbcrlf & PreobrazShifr$
roofcalc_ownDeshifrator = Right(PreobrazShifr$, Len(PreobrazShifr$) - (n - 1))
End Function


'Public Function Shifr(txtFileEDIT$) As String
'Dim l01D6 As Single
'Dim l01D2 As Long
'Dim l01D4 As Integer
'Dim l01D8 As String
'PVer = 2.1
'
'Screen.MousePointer = 11
'
''Debug.Print txtFileEDIT$
'
'On Error GoTo err1
' For l01D2 = 1 To Len(txtFileEDIT$) - 2 Step 3
'     l01D8$ = Mid$(txtFileEDIT$, l01D2 + 2, 1)
'     Mid$(txtFileEDIT$, l01D2 + 2, 1) = Mid$(txtFileEDIT$, l01D2, 1)
'    Mid$(txtFileEDIT$, l01D2, 1) = Mid$(txtFileEDIT$, l01D2 + 1, 1)
'    Mid$(txtFileEDIT$, l01D2 + 1, 1) = l01D8$
'
'    For l01D4 = 0 To 2 Step 1
'      l01D6 = Asc(Mid$(txtFileEDIT$, l01D2 + l01D4, 1)) + (73 + l01D4)
'      If l01D6 > 254 Then l01D6 = l01D6 - 256
'      If l01D6 = -1 Then
'      l01D6 = 1
'      End If
'      Mid$(txtFileEDIT$, l01D2 + l01D4, 1) = Chr$(l01D6)
'    Next l01D4
'
'Next
'
''Debug.Print txtFileEDIT$
''Shifr = txtFileEDIT$
'Shifr = "RB" & PVer & Properties & txtFileEDIT$ '& " This file is specification of " & App.ProductName & " " & Gl.VER & ". Copyright © 2000-2003 MDinc. All Right Reserved."
''Debug.Print Shifr
'Screen.MousePointer = 0
'Exit Function
'err1:
'Screen.MousePointer = 0
'End Function

