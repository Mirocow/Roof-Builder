VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Complect 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   14482428
      BackColorBkg    =   14482428
      Appearance      =   0
   End
End
Attribute VB_Name = "Complect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MSFlexGrid1.Clear
Dim RSC As Recordset
Set RSC = MainBaseFunction.RequestSQL("select c.code, c.name, c.length, c.devide, c.overcloak  from Completes c order by c.id")

If Not RSC Is Nothing Then

MSFlexGrid1.Cols = RSC.Fields.Count
'MSFlexGrid1.Rows = rs.RecordCount + 1

MSFlexGrid1.ColWidth(0) = 1200
MSFlexGrid1.ColWidth(1) = 5000

'Dim c As Integer
'For c = 0 To 2
'MSFlexGrid1.TextMatrix(0, 0) = "ID"
'MSFlexGrid1.TextMatrix(0, 1) = "Дата создания"
'MSFlexGrid1.TextMatrix(0, 2) = "Имя клиента"
'Next

Dim r As Integer
r = 1
Do While Not RSC.EOF
    For c = 0 To RSC.Fields.Count - 1
    MSFlexGrid1.Rows = r + 1
    MSFlexGrid1.TextMatrix(r, c) = CheckNull(RSC.Fields(c))
    Next
    r = r + 1
    RSC.MoveNext
Loop

RSC.Close
End If
End Sub

Private Sub MSFlexGrid1_Click()

End Sub
