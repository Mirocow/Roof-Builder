VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "�hameleonButton.ocx"

Begin VB.Form ReportTEXT 
   Caption         =   "�����1"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   11880
   LinkTopic       =   "�����1"
   ScaleHeight     =   7230
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ��������2 
      Caption         =   "��������2"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton ��������1 
      Caption         =   "��������1"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.VScrollBar ���������1 
      Height          =   6255
      Left            =   11520
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.HScrollBar ���������1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   11415
   End
   Begin VB.PictureBox �������1 
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



Private Sub ���������1_Change()
ReportView.top = -���������1.Value
End Sub

Private Sub ���������1_Scroll()
ReportView.Left = -���������1.Value
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



'***����� ������
 '��������������
     ReportView.Line (Cm(1), Cm(gor1))-(Cm(19), Cm(gor1))
     gor2 = gor1 + 0.65
     ReportView.Line (Cm(1), Cm(gor2))-(Cm(19), Cm(gor2))
 '������������
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

'����������� �������
���������1.MAX = ReportView.Height - �������1.Height
���������1.MAX = ReportView.Width - �������1.Width





Dim Otstup As String
Dim Month As String
Otstup = "   "
        
        '***
        ReportView.FontBold = n1Setup.��������1.Value
        ReportView.FontItalic = n1Setup.��������2.Value
        ReportView.FontName = n1Setup.FontNameMY.Text
        ReportView.FontSize = Val(n1Setup.FFontSize.Caption)
        
        ReportView.CurrentX = Cm(1)
        ReportView.CurrentY = Cm(1)
        Select Case Format(Date, "m")
      Case 1: Month = "������"
      Case 2: Month = "�������"
      Case 3: Month = "�����"
      Case 4: Month = "������"
      Case 5: Month = "����"
      Case 6: Month = "����"
      Case 7: Month = "����"
      Case 8: Month = "�������"
      Case 9: Month = "��������"
      Case 10: Month = "�������"
      Case 11: Month = "������"
      Case 12: Month = "�������"
   End Select
        ReportView.Print "����:" + Otstup + Format(Date, "d") + " " + Month + " " + Format(Date, "yyyy")
        ReportView.CurrentX = Cm(1)
        ReportView.Print "����:" + Otstup + CStr(stZakazchic.txtFamely + " " + stZakazchic.txtInitial)
        ReportView.CurrentX = Cm(1)
        ReportView.Print "������� ������:" + Otstup + CStr(n1.RoofProf.Text)
        ReportView.CurrentX = Cm(1)
        ReportView.Print "�������� ��� ������:" + Otstup + CStr(RR)
        '     �����
        ReportView.FontBold = True
        ReportView.CurrentX = Cm(1.2)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "������������ ������"
        ReportView.CurrentX = Cm(8.2)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "��-��"
        ReportView.CurrentX = Cm(8.2)
        ReportView.Print "���."
        ReportView.CurrentX = Cm(10.2)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "���-��"
        ReportView.CurrentX = Cm(12.7)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "���� ��"
        ReportView.CurrentX = Cm(12.7)
        ReportView.Print "�������"
        ReportView.CurrentX = Cm(15.2)
        ReportView.CurrentY = Cm(3.8)
        ReportView.Print "�����"
        ReportView.FontBold = False
        ReportView.CurrentY = Cm(1)
        ReportView.CurrentX = Cm(15.5)
        Nschet = stZakazchic.txtLoc
        ReportView.Print "���� �:" + Otstup + CStr(NScet)
        ReportView.Print
        ReportView.CurrentX = Cm(10)
        ReportView.Print "��� ��������:" + Otstup + CStr(n1.RoofMat.Text)
        '***
'               ReportView.TextWidth()
'        ReportView.TextHeight()
        
        '��������������
        ReportView.DrawWidth = 2
        ReportView.Line (Cm(1), Cm(3.5))-(Cm(19), Cm(3.5))
        ReportView.Line (Cm(1), Cm(5))-(Cm(19), Cm(5))
        '������������
        ReportView.Line (Cm(1), Cm(3.5))-(Cm(1), Cm(5))
        ReportView.Line (Cm(8), Cm(3.5))-(Cm(8), Cm(5))
        ReportView.Line (Cm(10), Cm(3.5))-(Cm(10), Cm(5))
        ReportView.Line (Cm(12.5), Cm(3.5))-(Cm(12.5), Cm(5))
        ReportView.Line (Cm(15), Cm(3.5))-(Cm(15), Cm(5))
        ReportView.Line (Cm(19), Cm(3.5))-(Cm(19), Cm(5))
        '     ����� end

End Sub

Public Sub BezVodostoka()
Dim i As Integer
Dim LINEgor1 As Single
Static TEXTtemp As Single
Dim flagZapisiGrafic As Integer

' ����������� ����������
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
        ReportView.Print "���������� ��������"
        ReportView.CurrentX = Cm(8.2)
        ReportView.CurrentY = Cm(5.15) + TEXTtemp
        ReportView.Print "M"
        ReportView.CurrentX = Cm(8.55)
        ReportView.CurrentY = Cm(5.1) + TEXTtemp
        ReportView.FontSize = ReportView.FontSize - 1
        ReportView.Print "2"
        ReportView.FontSize = ReportView.FontSize + 1
Case 1
        ReportView.Print "������� ���� � ������"
Case 2
        ReportView.Print "LHP ��������� ������ �����������"
Case 3
        ReportView.Print "LH3 ��������� ������"
Case 4
        ReportView.Print "LPT 50x50 �������� ������"
Case 5
        ReportView.Print "LPT 250 �������� ������ "
Case 6
        ReportView.Print "LR 200 ��������� ������"
Case 7
        ReportView.Print "LL 416 ������ ��� ���� (����������)"
Case 8
        ReportView.Print " LSPL 310 ��. ��� ��. ������ (�������)"
Case 9
        ReportView.Print "LNU 3 ������  ��� �������� �����"
Case 10
        ReportView.Print "LNS 3 ������ ��� ���������� �����"
Case 11
        ReportView.Print "LHPK ����� �� ��������� ������"
Case 12
        ReportView.Print "LAPK ����� ��� �������� �����"
Case 13
        ReportView.Print "LYN Y - �������� ������ ������"
Case 14
        ReportView.Print "LTH T- �������� ������ ������"
Case 15
        ReportView.Print "RA 4,9�27 ����� �������"
Case 16
        ReportView.Print "������ ����� �����"
Case 17
        ReportView.Print "�� �����"
Case 18
        ReportView.Print "Soydaband ������ ����� 0,15�10 �"
Case 19
        ReportView.Print "�������������-����� 1,3�50 �"
Case 20
        ReportView.Print "�������"
Case 21
        ReportView.Print "����. ��������"
Case 22
        ReportView.Print "�������� �� �����"
Case 23
        ReportView.Print "�������� �� �����"
Case 24
        ReportView.Print "������  3 �"
Case 25
        ReportView.Print "�������� ��� 400�400"
Case 26
        ReportView.Print "�������� ��� 600x600"
Case 27
        ReportView.Print "�������� �����, VH32"
Case 28
        ReportView.Print "���������������� " + "vlen"
Case 29
        ReportView.Print "���������������� " + "vle�"
Case 30
        ReportView.Print "���������������� " + "LE-310"
Case 31
        ReportView.Print " �������������� ����� " + "VPE"

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
             ReportView.Print "�����"
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
   ' *** Convertit les centim�tres en coordonn�es pour la printer ***
   
   On Error Resume Next
   
   Cm = Int(dCentimetres * 567)

End Function

Private Sub ��������1_Click()

End Sub

Private Sub ��������2_Click()

End Sub
