VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form ExportM 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Manager"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   6120
      Width           =   5655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "Operations with lapes"
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   5655
      Begin VB.CheckBox Check7 
         BackColor       =   &H00808000&
         Caption         =   "Add the slopes to current project"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00808000&
         Caption         =   "Mirror copy"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00808000&
         Caption         =   "Rewrite"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Export >>"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808000&
         Caption         =   "With cutting"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "As well as the first operation exports one of slopes in current the project. As it is direct so mirror."
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1440
         TabIndex        =   8
         Top             =   2040
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Operations with main picture"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4920
         TabIndex        =   14
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808000&
         Caption         =   "Rewrite"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808000&
         Caption         =   "With slopes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export mirror >>"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Export >>"
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   2475
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         Caption         =   "Left:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Operation with main picture The first button is direct Export to current project. And the second is mirror export. "
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   2760
         TabIndex        =   7
         Top             =   1920
         Width           =   2775
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Last modification at 07.10.04"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7680
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "The status of loading"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5920
      Width           =   2535
   End
End
Attribute VB_Name = "ExportM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileXport As String
Public FlagExport As Boolean

' Данные рисунка кровли
Dim EXpScaleLeft_Main As Single
Dim EXpScaleWidth_Main As Single
Dim EXpScaleTop_Main As Single
Dim EXpScaleHeight_Main As Single

Dim EXpN_Points_Main As Integer
Dim EXpLine_PX_Main(MAXSLOPESIN) As Single
Dim EXpLine_PY_Main(MAXSLOPESIN) As Single
Dim EXpAmount_Lines_Main As Integer
Dim EXpPoints_m_A(1 To MAXSLOPESIN) As Integer
Dim EXpPoints_m_B(1 To MAXSLOPESIN) As Integer
Dim EXpKolvoScatov As Integer
Dim EXpLabel_X(1 To MAXSLOPES) As Single
Dim EXpLabel_Y(1 To MAXSLOPES) As Single

' Данные рисунка скатов
Dim EXpP_Sheets(MAXSLOPES) As String * 50

Dim EXpScaleLeftS(1 To MAXSLOPES) As Single
Dim EXpScaleWidthS(1 To MAXSLOPES) As Single
Dim EXpScaleTopS(1 To MAXSLOPES) As Single
Dim EXpScaleHeightS(1 To MAXSLOPES) As Single

Dim EXpAmount_PB(MAXSLOPES) As Integer
Dim EXpAmount_PA(MAXSLOPES) As Integer
Dim EXpLape_Points_X(1 To MAXSLOPES, MAXSLOPESINE + 2) As Single
Dim EXpLape_Points_Y(1 To MAXSLOPES, MAXSLOPESINE + 2) As Single

Dim EXpPoints_A(1 To MAXSLOPES, 1 To MAXSLOPESINE + 2) As Integer
Dim EXpPoints_B(1 To MAXSLOPES, 1 To MAXSLOPESINE + 2) As Integer

Dim EXpPn_Red_lines(1 To MAXSLOPES) As Integer
Dim EXpPX_StartLC(1 To MAXSLOPES) As Single
Dim EXpPn_StartLC(1 To MAXSLOPES) As Integer
Dim EXpN_Sheets(1 To MAXSLOPESIST) As Integer

Dim EXpLapesDescrib(1 To MAXSLOPES) As String

Dim EXpcuts(1 To MAXSLOPES) As Integer

Dim EXpList_Properties_Length(1 To MAXSLOPES, MAXSLOPESIST, MAXC)   As Single ' длина полосы
Dim EXpList_Properties_PX(1 To MAXSLOPES, MAXSLOPESIST, MAXC) As Single   ' Координаты по X (НАЧАЛО ПРОРИСОВКИ)
Dim EXpList_Properties_PY(1 To MAXSLOPES, MAXSLOPESIST, MAXC) As Single  ' Координаты по Y (НАЧАЛО ПРОРИСОВКИ)


Function Load_f()
Dim newformat As Boolean

Dim l0034 As String

Dim strn As Single
Dim str As String
Dim dstr As Double

Dim l0040 As Single
Dim i As Single
Dim l00E0 As Single
Dim FileN As Integer

On Error GoTo ERR

  FileN = FreeFile
  Open FileXport For Binary Access Read As #FileN
  
  If LOF(FileN) = 0 Then
    If MsgBox(ResolveResstring(1447), vbCritical, ResolveResstring(1413)) = vbOK Then
    Close #FileN
    FileXport = ""
    Exit Function
    End If
  End If
  
  ' RB NEW
  str = string$(8, " ")
  Get #FileN, , str
  If Right(str, 4) = chr$(2) & chr$(0) & "N" & chr$(9) Then
  newformat = True
  Get #FileN, , dstr
  Get #FileN, , strn
  str = Input$(strn, FileN)
  Else
  Seek FileN, 1
  Get #FileN, , strn
  str = Trim$(Input$(strn, FileN))
  End If

  Get #FileN, , strn
  If Not strn = 0 Then
  str = Input$(strn, FileN)
  Else
  End If
  Get #FileN, , strn
  
  str = Trim$(Input$(strn, FileN))
  
  For i = 0 To MAXSLOPES
  EXpP_Sheets(i) = profil
  Next
  
  Get #FileN, , strn
  str = Input$(strn, FileN)
  Get #FileN, , strn
  str = Input$(strn, FileN)
  Get #FileN, , strn
  str = Input$(strn, FileN)

  
  ' Начало загрузки переменных главного рисунка
  Text2 = "Загрузка рисунка кровли..." & vbCrLf
  
  Get #FileN, , EXpScaleLeft_Main
  Text2 = Text2 & EXpScaleLeft_Main & vbCrLf
  Get #FileN, , EXpScaleWidth_Main
  Text2 = Text2 & EXpScaleWidth_Main & vbCrLf
  Get #FileN, , EXpScaleTop_Main
  Text2 = Text2 & EXpScaleTop_Main & vbCrLf
  Get #FileN, , EXpScaleHeight_Main
  Text2 = Text2 & EXpScaleHeight_Main & vbCrLf
  
  Text2 = Text2 & "Загрузка координат точек..." & vbCrLf
  Get #FileN, , EXpN_Points_Main
  Text2 = Text2 & EXpN_Points_Main & vbCrLf
  For i = 1 To EXpN_Points_Main Step 1
    Get #FileN, , EXpLine_PX_Main(i)
    Text2 = Text2 & EXpLine_PX_Main(i) & vbCrLf
    Get #FileN, , EXpLine_PY_Main(i)
    Text2 = Text2 & EXpLine_PY_Main(i) & vbCrLf
  Next i
  
  Get #FileN, , EXpAmount_Lines_Main
  For i = 1 To EXpAmount_Lines_Main Step 1
    Get #FileN, , EXpPoints_m_A(i)
    Get #FileN, , EXpPoints_m_B(i)
  Next i
  
  Text2 = Text2 & "Загрузка координат обозначений..." & vbCrLf
  Get #FileN, , EXpKolvoScatov
  Text2 = Text2 & EXpKolvoScatov & vbCrLf
  For i = 1 To EXpKolvoScatov Step 1
    Get #FileN, , EXpLabel_X(i)
    Text2 = Text2 & EXpLabel_X(i) & vbCrLf
    Get #FileN, , EXpLabel_Y(i)
    Text2 = Text2 & EXpLabel_Y(i) & vbCrLf
  Next i
  ' Конец
  
  Get #FileN, , strn
  str = Input$(strn, FileN)

Text2 = Text2 & "Загрузка свойств скатов..." & vbCrLf
For N_Slope = 1 To EXpKolvoScatov Step 1

    If Right(FileXport, 3) = "rbp" Then Get #FileN, , EXpP_Sheets(N_Slope)

    Get #FileN, , EXpScaleLeftS(N_Slope)
    Get #FileN, , EXpScaleWidthS(N_Slope)
    Get #FileN, , EXpScaleTopS(N_Slope)
    Get #FileN, , EXpScaleHeightS(N_Slope)

    Get #FileN, , EXpAmount_PB(N_Slope)
    For i = 1 To EXpAmount_PB(N_Slope) Step 1
      Get #FileN, , EXpLape_Points_X(N_Slope, i)
      Get #FileN, , EXpLape_Points_Y(N_Slope, i)
    Next i

    Get #FileN, , EXpAmount_PA(N_Slope)
    For i = 1 To EXpAmount_PA(N_Slope) Step 1
      Get #FileN, , EXpPoints_A(N_Slope, i)
      Get #FileN, , EXpPoints_B(N_Slope, i)
    Next i

    Get #FileN, , EXpPn_Red_lines(N_Slope)
    Get #FileN, , EXpPX_StartLC(N_Slope)
    Get #FileN, , EXpPn_StartLC(N_Slope)
    Get #FileN, , EXpN_Sheets(N_Slope)

    If newformat = True Then
     Get #FileN, , EXpcuts(N_Slope)
'     If Expcuts(N_Slope) > Gl.MAXC Then Expcuts(N_Slope) = Gl.MAXC
    Else
     EXpcuts(N_Slope) = 2
    End If

    Text2 = Text2 & "Свойства плоскости [" & N_Slope & "] загружены" & vbCrLf

    Text2 = Text2 & "Загрузка параметров раскроя..." & vbCrLf
    For i = 1 To EXpN_Sheets(N_Slope) Step 1
'    Expcutsp(N_Slope, i) = EXpcuts(N_Slope)
      For l00E0 = 0 To EXpcuts(N_Slope) Step 1
        Get #FileN, , EXpList_Properties_PY(N_Slope, i, l00E0)
        Get #FileN, , EXpList_Properties_PX(N_Slope, i, l00E0)
        Get #FileN, , EXpList_Properties_Length(N_Slope, i, l00E0)
      Next l00E0
    Next i
    Text2 = Text2 & "Параметры раскроя загружены." & vbCrLf

    Get #FileN, , strn
    str = Input$(strn, FileN)
    Get #FileN, , strn

Next N_Slope

  Close #FileN
  Text2 = Text2 & "[Загрузка завершена успешно]" & vbCrLf
  Exit Function
ERR:
    STRERROR = STRERROR & Time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbCrLf
   Close #FileN
   Text2 = Text2 & "[Произошла ошибка в процессе загрузки]" & vbCrLf
'   MsgBox ResolveResstring(1452, "%CATALOGUE%", "", "%FILE%", FileXport), vbCritical, ResolveResstring(1413) '"Can`t read this file." & vbcrlf & Gl.Catalogue_files & Gl.CurrentFile & vbcrlf & "Maybe this file is corrupted.", vbCritical, "System error"
'   FileXport = ""
End Function

Private Sub Check2_Click()
MsgBox "Не риализовано"
Exit Sub
End Sub

Private Sub Check3_Click()
MsgBox "Не риализовано"
Exit Sub
End Sub

Private Sub Check5_Click()
MsgBox "Не риализовано"
Exit Sub
End Sub

Private Sub Command1_Click()
ExportMain False
FlagExport = False
'Unload Me
End Sub

Private Sub Command2_Click()
ExportMain True
FlagExport = False
'Unload Me
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim icopy As Integer
Dim l00BC As Single
Dim c
Dim XMin As Single
Dim XMax As Single
Dim xCentr As Single
Dim n As Integer
Dim N_list As Integer

If Check7.Value = 1 Then icopy = List2.ListCount + 1

For i = 1 To List1.ListCount - 1

If Check7.Value = 1 Then
icopy = icopy + 1
Else
icopy = i
End If

If List1.Selected(i - 1) = True Then

Gl.P_Sheets(icopy) = EXpP_Sheets(i)

If Check3.Value = 1 Then
' Перезапись или запись
Gl.Pn_Red_lines(icopy) = EXpPn_Red_lines(i)
Gl.Cuts(icopy) = EXpcuts(i)
Pn_StartLC(icopy) = EXpPn_StartLC(i)
End If

  ' Запись
  SlP(i).BpCount = EXpAmount_PB(i)
  For N_list = 1 To SlP(i).BpCount Step 1
    Lape_Points_X(icopy, N_list) = EXpLape_Points_X(i, N_list)
    Lape_Points_Y(icopy, N_list) = EXpLape_Points_Y(i, N_list)
  Next N_list
  
  SlP(icopy).ApCount = EXpAmount_PA(i)
  For N_list = 1 To SlP(i).ApCount Step 1
    Points_A(icopy, N_list) = EXpPoints_A(i, N_list)
    Points_B(icopy, N_list) = EXpPoints_B(i, N_list)
  Next N_list
  
  If Check3.Value = 1 Then
  N_Sheets(icopy) = EXpN_Sheets(i)
  
  For N_list = 0 To N_Sheets(i) Step 1
    For c = 0 To Cuts(i) Step 1
        List_Properties_PX(icopy, N_list, c) = EXpList_Properties_PX(i, N_list, c)
        List_Properties_PY(icopy, N_list, c) = EXpList_Properties_PY(i, N_list, c)
        List_Properties_Length(icopy, N_list, c) = EXpList_Properties_Length(i, N_list, c)
    Next c
  Next N_list
  End If
  
  ScaleLeftS(icopy) = EXpScaleLeftS(i)
  ScaleWidthS(icopy) = EXpScaleWidthS(i)
  ScaleTopS(icopy) = EXpScaleTopS(i)
  ScaleHeightS(icopy) = EXpScaleHeightS(i)
  LapesDescrib(icopy) = EXpLapesDescrib(i)
  
If Check6.Value = 1 Then
' Обработка данных после заполнения

    ' Зеркальный рисунок ската
    XMin = 99999
    XMax = 0
    
    For N_list = 1 To SlP(i).BpCount Step 1
      If Lape_Points_X(i, N_list) > XMax Then XMax = Lape_Points_X(i, N_list)
      If Lape_Points_X(i, N_list) < XMin Then XMin = Lape_Points_X(i, N_list)
    Next N_list
    
    xCentr = XMin + ((XMax - XMin) / 2)
    For N_list = 1 To SlP(i).BpCount Step 1
      Lape_Points_X(i, N_list) = xCentr - (Lape_Points_X(i, N_list) - xCentr)
    Next N_list
    
    If Check3.Value = 1 Then
    ' Зеркальный раскрой
'    N_Sheets(i) = N_Sheets(i)
    For N_list = 1 To N_Sheets(i) Step 1
        For c = 0 To Gl.Cuts(i) Step 1
            List_Properties_PY(i, N_list, c) = List_Properties_PY(i, N_list, c)
            List_Properties_PX(i, N_list, c) = xCentr - (List_Properties_PX(i, N_list, c) - xCentr + Project.ListView1.ListItems(0)) 'Project.txtW)
            List_Properties_Length(i, N_list, c) = List_Properties_Length(i, N_list, c)
        Next c
    Next N_list
    End If
     
End If

End If
Next i


Me.List2.Clear
For i = 1 To icopy + 1 Step 1
    If i > 26 Then
    Me.List2.AddItem chr$(i + 70) & Space(5) & P_Sheets(i)
    Else
    Me.List2.AddItem chr$(i + 64) & Space(5) & P_Sheets(i)
    End If
Next i

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

  Load_f

  Me.Picture1.ScaleLeft = EXpScaleLeft_Main
  Me.Picture1.ScaleWidth = EXpScaleWidth_Main
  Me.Picture1.ScaleTop = EXpScaleTop_Main
  Me.Picture1.ScaleHeight = EXpScaleHeight_Main

  Me.Picture1.Cls
  ''''''''''''''''''''''''''
  l0138 = EXpScaleWidth_Main / 90
  l013A = 2 * l0138
  l013C = EXpScaleLeft_Main + 0.95 * EXpScaleWidth_Main
  l013E = EXpScaleTop_Main + 0.95 * EXpScaleHeight_Main
  
'меню скатов
'  Me.Picture1.Line (l013C - 100, l013E)-(l013C, l013E)
'  Me.Picture1.Line (l013C - 100, l013E - l0138)-(l013C - 100, l013E + l0138)
'  Me.Picture1.Line (l013C, l013E - l0138)-(l013C, l013E + l0138)
'  Me.Picture1.PSet (l013C - 15 - l013A, l013E + 1.5 * l013A), "&H00DCFBFC"
'  Me.Picture1.Print "1 x " & Format(Picture1.ScaleWidth / 100, "###")

Dim i As Integer

Len_AB_X = EXpScaleWidth_Main / 350
For i = 1 To EXpN_Points_Main Step 1
  Me.Picture1.Circle (EXpLine_PX_Main(i), EXpLine_PY_Main(i)), Len_AB_X, vbBlack
Next i
  
For i = 1 To EXpAmount_Lines_Main Step 1
    Me.Picture1.Line (EXpLine_PX_Main(EXpPoints_m_A(i)), EXpLine_PY_Main(EXpPoints_m_A(i)))-(EXpLine_PX_Main(EXpPoints_m_B(i)), EXpLine_PY_Main(EXpPoints_m_B(i))), RGB(0, 0, 0)
Next i

For i = 1 To EXpKolvoScatov Step 1
    If i > 26 Then
    Me.List1.AddItem chr$(i + 70) & Space(5) & EXpP_Sheets(i)
    Else
    Me.List1.AddItem chr$(i + 64) & Space(5) & EXpP_Sheets(i)
    End If
Next i

For i = 1 To KolvoScatov Step 1
    If i > 26 Then
    Me.List2.AddItem chr$(i + 70) & Space(5) & P_Sheets(i)
    Else
    Me.List2.AddItem chr$(i + 64) & Space(5) & P_Sheets(i)
    End If
Next i
  
End Sub

Sub ExportMain(MirrorDirect As Boolean)

If KolvoScatov = 0 Or Check4.Value = 1 Then

ScaleLeft_Main = EXpScaleLeft_Main
ScaleWidth_Main = EXpScaleWidth_Main
ScaleTop_Main = EXpScaleTop_Main
ScaleHeight_Main = EXpScaleHeight_Main
  
N_Points_Main = EXpN_Points_Main
For i = 1 To EXpN_Points_Main Step 1
  Line_PX_Main(i) = EXpLine_PX_Main(i)
  Line_PY_Main(i) = EXpLine_PY_Main(i)
Next i
  
Amount_Lines_Main = EXpAmount_Lines_Main
For i = 1 To EXpAmount_Lines_Main Step 1
  Points_m_A(i) = EXpPoints_m_A(i)
  Points_m_B(i) = EXpPoints_m_B(i)
Next i
  
If Check1.Value = 1 Then
KolvoScatov = EXpKolvoScatov
For i = 1 To EXpKolvoScatov Step 1
  Label_X(i) = EXpLabel_X(i)
  Label_Y(i) = EXpLabel_Y(i)
Next i
End If

If MirrorDirect = True Then

    XMin = 99999
    XMax = 0
    
    For i = 1 To N_Points_Main Step 1
      If Line_PX_Main(i) > XMax Then XMax = Line_PX_Main(i)
      If Line_PX_Main(i) < XMin Then XMin = Line_PX_Main(i)
    Next i
    
    xCentr = XMin + ((XMax - XMin) / 2)
    For i = 1 To N_Points_Main Step 1
      Line_PX_Main(i) = xCentr - (Line_PX_Main(i) - xCentr)
    Next i
    
    If Check1.Value = 1 Then
    XMin = 99999
    XMax = -99999
    
    For i = 1 To KolvoScatov Step 1
      If Label_X(i) > XMax Then XMax = Label_X(i)
      If Label_X(i) < XMin Then XMin = Label_X(i)
    Next i
    
    xCentr = XMin + ((XMax - XMin) / 2)
    For i = 1 To KolvoScatov Step 1
        Label_X(i) = xCentr - (Label_X(i) - xCentr) - Val(Text1)
    Next i
    End If

End If

End If

ROOFPIC.Command5.Value = True
End Sub

Private Sub List2_Click()
MsgBox "Не риализовано"
Exit Sub
End Sub
