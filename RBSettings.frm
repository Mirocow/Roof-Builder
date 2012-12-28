VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form Form1 
   Caption         =   "Roof Builder settings"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "RBSettings.frx":0000
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function LngGetIDsInfo Lib "rb_loc.dll" (ids As Long, ByVal count As Long) As Long

'Private getlicn As New LI1a0.GetLicence

Public Function mGetComputerName() As String
Dim compuname As String
compuname = String$(255, " ")
GetComputerName compuname, 255
mGetComputerName = Replace$(Trim$(compuname), Chr$(0), "")
End Function


Private Sub Form_Load()
Dim Path As String
On Error GoTo ERR

Text1 = ""

Text1 = Text1 & "Registered information:" & vbNewLine
Text1 = Text1 & "Name: " & GetSetting("Roof Builder", "Main", "username", "") & vbNewLine
Text1 = Text1 & "Licence: " & GetSetting("Roof Builder", "Main", "licence", "") & vbNewLine
Text1 = Text1 & "Reg N: " & GetSetting("Roof Builder", "REG", "Regnumber", "") & vbNewLine
Text1 = Text1 & "Ver: " & GetSetting("Roof Builder", "REG", "Ver", "") & vbNewLine & vbNewLine

chekpv = GetSetting("Roof Builder", "REG", "Ver", "")

' Проверка совпадения ID продукта
'Dim i As Integer
'Dim templic As String
'Dim temp As String
'Dim lenstr As Integer
'Dim mvarCOMPUTERNAME As String
'mvarCOMPUTERNAME = mGetComputerName & chr$(0)
'lenstr = Len(mvarCOMPUTERNAME + "Roof Builder" + chekpv)
'temp = mvarCOMPUTERNAME + "Roof Builder" + chekpv
'templic = ""
'For i = 1 To lenstr
'templic = templic & CStr(Asc(Mid(temp, i, 1)) Xor Asc(Mid(temp, lenstr - i + 1, 1)) And Asc(Mid(temp, lenstr - i + 1, 1)))
'Next

Text1 = Text1 & "Test information:" & vbNewLine
Text1 = Text1 & "Name: " & mGetComputerName & vbNewLine
'Text1 = Text1 & "Licence: " & templic & vbNewLine & vbNewLine

'Text1 = Text1 & "Test lic: " & lic & vbNewLine & vbNewLine

Text1 = Text1 & "Config file: " & GetSetting("Roof Builder", "Main", "mdbcofigfile", "") & vbNewLine
Text1 = Text1 & "Projects` directory: " & GetSetting("Roof Builder", "Main", "url_work", "") & vbNewLine & vbNewLine

Path = GetSetting("Roof Builder", "Main", "mdbcofigfile", "")
Path = Replace(Path, "cfg\materials.mdb", "")

strPlugin = Dir(Path & "\Plugins\*.dll")
'Text1 = Text1 & "Libraries` directory:" & vbNewLine
Text1 = Text1 & "Libraries:" & vbNewLine
 While strPlugin <> ""
  Text1 = Text1 & strPlugin & vbNewLine
  strPlugin = Dir()
 Wend
 
Text1 = Text1 & vbNewLine
 
MySettings = GetAllSettings(appname:="Roof Builder", section:="Main")
For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
   Text1 = Text1 & MySettings(intSettings, 0) & " = " & MySettings(intSettings, 1) & vbNewLine
Next intSettings

Text1 = Text1 & vbNewLine

MySettings = GetAllSettings(appname:="Roof Builder", section:="REG")
For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
   Text1 = Text1 & MySettings(intSettings, 0) & " = " & MySettings(intSettings, 1) & vbNewLine
Next intSettings

Text1 = Text1 & vbNewLine

MySettings = GetAllSettings(appname:="Roof Builder", section:="Position")
For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
   Text1 = Text1 & MySettings(intSettings, 0) & " = " & MySettings(intSettings, 1) & vbNewLine
Next intSettings

Text1 = Text1 & vbNewLine

MySettings = GetAllSettings(appname:="Roof Builder", section:="RecentlyFiles")
For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
   Text1 = Text1 & MySettings(intSettings, 0) & " = " & MySettings(intSettings, 1) & vbNewLine
Next intSettings

Text1 = Text1 & vbNewLine

MySettings = GetAllSettings(appname:="Roof Builder", section:="CalcOption")
For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
   Text1 = Text1 & MySettings(intSettings, 0) & " = " & MySettings(intSettings, 1) & vbNewLine
Next intSettings

Text1 = Text1 & vbNewLine

MySettings = GetAllSettings(appname:="Roof Builder", section:="RunUp")
For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
   Text1 = Text1 & MySettings(intSettings, 0) & " = " & MySettings(intSettings, 1) & vbNewLine
Next intSettings

'Text1 = Text1 & vbNewLine
'
'MySettings = GetAllSettings(appname:="Roof Builder", section:="plgs")
'For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
'   Text1 = Text1 & MySettings(intSettings, 0) & " = " & MySettings(intSettings, 1) & vbNewLine
'Next intSettings

Exit Sub
ERR:
Text1 = Text1 & ERR.Number & " " & ERR.Description & vbNewLine
Resume Next
End Sub

Private Sub Form_Resize()
Text1.Top = Me.ScaleTop
Text1.Left = Me.ScaleLeft
Text1.Width = Me.ScaleWidth
Text1.Height = Me.ScaleHeight
End Sub

