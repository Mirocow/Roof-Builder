VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form ReportTEXT 
   Caption         =   "Форма1"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   11880
   LinkTopic       =   "Форма1"
   ScaleHeight     =   7230
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Комманда2 
      Caption         =   "Комманда2"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Комманда1 
      Caption         =   "Комманда1"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.VScrollBar ВСкролинг1 
      Height          =   6255
      Left            =   11520
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.HScrollBar ГСкролинг1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   11415
   End
   Begin VB.PictureBox Рисунок1 
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6195
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   600
      Width           =   11415
      Begin VB.PictureBox ReportView 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   16726
         Left            =   120
         ScaleHeight     =   16665
         ScaleWidth      =   11850
         TabIndex        =   1
         Top             =   120
         Width           =   11907
      End
   End
End
Attribute VB_Name = "ReportTEXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Call Shapka
Call BezVodostoka
End Sub



Private Sub ВСкролинг1_Change()
ReportView.top = -ВСкролинг1.Value
End Sub

Private Sub ГСкролинг1_Scroll()
ReportView.Left = -ГСкролинг1.Value
End Sub

Public Function preobraz(MYstring As String)
preobraz = ReportView.TextWidth(MYstring) + " * " + ReportView.TextHeight(MYstring)
End Function

Public Sub DrawBox(KolStroc As Integer, ByVal gor As Single, flag As Boolean)
Static gor1 As Single
Static gor2 As Single
gor1 = gor
Dim i As Integer
For i = 0 To KolStroc
If Not n1.Kolvo(i).Text = 0 Then

If i = 1 Then
flag = False
End If



'***Рамка строки
 'горизонтальные
     ReportView.Line (Cm(1), Cm(gor1))-(Cm(19), Cm(gor1))
     gor2 = gor1 + 0.65
     ReportView.Line (Cm(1), Cm(gor2))-(Cm(19), Cm(gor2))
 'вертикальные
     ReportView.Line (Cm(1), Cm(gor1))-(Cm(1), Cm(gor2))
     ReportView.Line (Cm(8), Cm(gor1))-(Cm(8), Cm(gor2))
     ReportView.Line (Cm(10), Cm(gor1))-(Cm(10), Cm(gor2))
     ReportView.Line (Cm(12.5), Cm(gor1))-(Cm(12.5), Cm(gor2))
     ReportView.Line (Cm(15), Cm(gor1))-(Cm(15), Cm(gor2))
     ReportView.Line (Cm(19), Cm(gor1))-(Cm(19), Cm(gor2))


If flag = True Then
gor1 = gor1 + 0.65 * 2
End If
gor1 = gor1 + 0.65

'***
End If
Next

End Sub

Public Sub Shapka()
ReportView.Left = 0
ReportView.top = 0
ReportTEXT.Left = (Screen.Width - ReportTEXT.Width) / 2

'выставление скролов
ВСкролинг1.MAX = ReportView.Height - Рисунок1.Height
ГСкролинг1.MAX = ReportView.Width - Рисунок1.Width





Dim Otstup As String
Dim Month As String
Otstup = "   "
        
        '***
        ReportView.FontBold = n1Setup.Проверка1.Value
        ReportView.FontItalic = n1Setup.Проверка2.Value
        ReportView.FontName = n1Setup.FontNameMY.Text
        ReportView.FontSize = Val(n1Setup.FFontSize.Caption)
        
        ReportView.CurrentX = Cm(1)
        ReportView.CurrentY = Cm(1)
        Select Case Format(Date, "m")
      Case 1: Month = "Января"
      Case 2: Month = "Февраля"
      Case 3: Month = "Марта"
      Case 4: Month = "Апреля"
      Case 5: Month = "Майя"
      Case 6: Month = "Июня"
      Case 7: Month = "Июля"
      Case 8: Month = "Августа"
      Case 9: Month = "Сентября"
      Case 10: Month = "Октября"
      Case 11: Month = "Ноября"
      Case 12: Month = "Декабря"
   End Select
        ReportView.Print "Дата:" + Otstup + Format(Date, "d") + " " + Month + " " + Format(Date, "yyyy")
        ReportView.CurrentX = Cm(1)
        ReportView.Print "Кому:" + Otstup + CStr(stZakazchic.txtFamely + " " + stZakazchic.txtInitial)
        ReportView.CurrentX = Cm(1)
        ReportView.Print "Профиль кровли:" + Otstup + CStr(n1.RoofProf.Text)
        ReportView.CurrentX = Cm(1)
        ReportView.Print "Цветовой код кровли:" + Otstup + CStr(RR)
        '     Шапка
        ReportView.FontBold = True
        ReportView.CurrentX = Cm(1.2)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "НАИМЕНОВАНИЕ ТОВАРА"
        ReportView.CurrentX = Cm(8.2)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "ЕД-ЦА"
        ReportView.CurrentX = Cm(8.2)
        ReportView.Print "Изм."
        ReportView.CurrentX = Cm(10.2)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "КОЛ-ВО"
        ReportView.CurrentX = Cm(12.7)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "ЦЕНА ЗА"
        ReportView.CurrentX = Cm(12.7)
        ReportView.Print "ЕДЕНИЦУ"
        ReportView.CurrentX = Cm(15.2)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "СУММА"
        ReportView.FontBold = False
        ReportView.CurrentY = Cm(1)
        ReportView.CurrentX = Cm(15.5)
        Nschet = stZakazchic.txtLoc
        ReportView.Print "Счет №:" + Otstup + CStr(NScet)
        ReportView.Print
        ReportView.CurrentX = Cm(10)
        ReportView.Print "Вид покрытия:" + Otstup + CStr(n1.RoofMat.Text)
        '***
'               ReportView.TextWidth()
'        ReportView.TextHeight()
        
        'горизонтальные
        ReportView.DrawWidth = 2
        ReportView.Line (Cm(1), Cm(3.5))-(Cm(19), Cm(3.5))
        ReportView.Line (Cm(1), Cm(5))-(Cm(19), Cm(5))
        'вертикальные
        ReportView.Line (Cm(1), Cm(3.5))-(Cm(1), Cm(5))
        ReportView.Line (Cm(8), Cm(3.5))-(Cm(8), Cm(5))
        ReportView.Line (Cm(10), Cm(3.5))-(Cm(10), Cm(5))
        ReportView.Line (Cm(12.5), Cm(3.5))-(Cm(12.5), Cm(5))
        ReportView.Line (Cm(15), Cm(3.5))-(Cm(15), Cm(5))
        ReportView.Line (Cm(19), Cm(3.5))-(Cm(19), Cm(5))
        '     Шапка end

End Sub

Public Sub BezVodostoka()
Dim i As Integer
Dim LINEgor1 As Single
Static TEXTtemp As Single
Dim flagZapisiGrafic As Integer

' управляющие переменные
Dim UPR As Single
Dim UPR1 As Single
UPR = 5.1 '+ 2
'UPR1 = Cm(0.65) * 3

LINEgor1 = UPR      '=5.1
TEXTtemp = UPR1       '=0


For i = 0 To 48
If Not n1.Kolvo(i).Text = 0 Then

If i = 1 Then
flagZapisiGrafic = 1
ElseIf i = 2 Then
flagZapisiGrafic = 0
End If
'
'If i = 31 Then
'flagZapisiGrafic = 1
'ElseIf i = 32 Then
'flagZapisiGrafic = 0
'End If



        ReportView.CurrentX = Cm(1.1)
        ReportView.CurrentY = Cm(5.15) + TEXTtemp
Select Case i
Case 0
        ReportView.Print "Кровельный материал"
        ReportView.CurrentX = Cm(8.2)
        ReportView.CurrentY = Cm(5.15) + TEXTtemp
        ReportView.Print "M"
        ReportView.CurrentX = Cm(8.55)
        ReportView.CurrentY = Cm(5.1) + TEXTtemp
        ReportView.FontSize = ReportView.FontSize - 1
        ReportView.Print "2"
        ReportView.FontSize = ReportView.FontSize + 1
Case 1
        ReportView.Print "Гладкий лист в рулоне"
Case 2
        ReportView.Print "LHP Коньковая планка полукруглая"
Case 3
        ReportView.Print "LH3 Коньковая планка"
Case 4
        ReportView.Print "LPT 50x50 Торцевая планка"
Case 5
        ReportView.Print "LPT 250 Торцевая планка "
Case 6
        ReportView.Print "LR 200 Карнизная планка"
Case 7
        ReportView.Print "LL 416 Планка для швов (пристенная)"
Case 8
        ReportView.Print " LSPL 310 Пл. для вн. стыков (яндовая)"
Case 9
        ReportView.Print "LNU 3 Планка  для наружных углов"
Case 10
        ReportView.Print "LNS 3 Планка для внутренних углов"
Case 11
        ReportView.Print "LHPK Конец на коньковую планку"
Case 12
        ReportView.Print "LAPK Конец для шатровой крыши"
Case 13
        ReportView.Print "LYN Y - образная планка конька"
Case 14
        ReportView.Print "LTH T- образная планка конька"
Case 15
        ReportView.Print "RA 4,9х27 Шуруп цветной"
Case 16
        ReportView.Print "Краска спрей колор"
Case 17
        ReportView.Print "ТН Филер"
Case 18
        ReportView.Print "Soydaband Липкая лента 0,15х10 м"
Case 19
        ReportView.Print "Гидроизоляция-рулон 1,3х50 м"
Case 20
        ReportView.Print "Силикон"
Case 21
        ReportView.Print "Спец. упаковка"
Case 22
        ReportView.Print "Лестница на крышу"
Case 23
        ReportView.Print "Лестница на стену"
Case 24
        ReportView.Print "Мостик  3 м"
Case 25
        ReportView.Print "Пожарный люк 400х400"
Case 26
        ReportView.Print "Пожарный люк 600x600"
Case 27
        ReportView.Print "Антенный выход, VH32"
Case 28
        ReportView.Print "Снегозадержатель " + "vlen"
Case 29
        ReportView.Print "Снегозадержатель " + "vleр"
Case 30
        ReportView.Print "Снегозадержатель " + "LE-310"
Case 31
        ReportView.Print " Вентиляционная труба " + "VPE"

End Select

       




        ReportView.CurrentX = Cm(10.2)
        ReportView.CurrentY = Cm(5.15) + TEXTtemp
        ReportView.Print CStr(n1.Kolvo(i).Text)
        
        
        ReportView.CurrentX = Cm(12.7)
        ReportView.CurrentY = Cm(5.15) + TEXTtemp
        ReportView.Print CStr(n1.CenaEd(i).Caption)
        ReportView.CurrentX = Cm(15.5)
        ReportView.CurrentY = Cm(5.15) + TEXTtemp
        ReportView.Print CStr(n1.Summ(i).Caption)
        
If flagZapisiGrafic = 1 Then
TEXTtemp = TEXTtemp + (Cm(0.65)) * 2
             ReportView.CurrentX = Cm(1.1)
             ReportView.CurrentY = Cm(4.5) + TEXTtemp
             ReportView.Print "Сумма"
             ReportView.CurrentX = Cm(15.5)
             ReportView.CurrentY = Cm(4.5) + TEXTtemp
      '***?
        
'        ReportView.Print CStr(n1.obsh.Caption)
      '***?

        
Call DrawBox(i, LINEgor1, False)
Else
TEXTtemp = TEXTtemp + Cm(0.65)
Call DrawBox(i - 1, LINEgor1, True)
End If
        
             

End If
Next
'        ReportView.CurrentX = Cm(15.5)
'        ReportView.CurrentY = Cm(5.15) + TEXTtemp
'        ReportView.Print CStr(n1.obsh1.Text)
End Sub

Function Cm(ByVal dCentimetres As Double) As Long
   ' *** Convertit les centimиtres en coordonnйes pour la printer ***
   
   On Error Resume Next
   
   Cm = Int(dCentimetres * 567)

End Function

Private Sub Комманда1_Click()

End Sub

Private Sub Комманда2_Click()

End Sub
