VERSION 5.00
Begin VB.Form lic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "lice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Accept"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "No"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "lice.frx":030A
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "lic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    SaveSetting App.ProductName, "Main", "lic_YES_NO", "No"
    Dim Button As MSComctlLib.Button
    Set Button = OfficeStart.Toolbar1.Buttons(10)
    OfficeStart.Toolbar1_ButtonClick Button
End Sub


Private Sub Command2_Click()
    SaveSetting App.ProductName, "Main", "lic_YES_NO", "Yes"
    Unload Me
End Sub


Private Sub Form_Load()
    SetFont Me
    Command1.Caption = lng.GetResIDstring(3005)
    Command2.Caption = lng.GetResIDstring(3004)
End Sub


