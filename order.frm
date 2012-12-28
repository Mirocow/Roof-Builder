VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form order 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Для оформления заказа пошлите ниже следующий текст на адрес roofbuilder@narod.ru"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Roof Builder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5760
      Picture         =   "order.frx":0000
      Top             =   3240
      Width           =   480
   End
End
Attribute VB_Name = "order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
About.Show vbModal, OfficeStart
End Sub

Private Sub Form_Load()
Text1 = "Форма заказа программы Roof Builder" & vbNewLine

Text1 = Text1 & "Компания/ч.л.: <Вставте название Вашей компании>" & vbNewLine & " желает приобрести лицензии версии <Вставте желаемую версию Prof & Light>"
Text1 = Text1 & vbNewLine
Text1 = Text1 & "Имя компьютера: " & Gl.UserName & vbNewLine
Text1 = Text1 & "Индификационный номер: " & Gl.ProductID & vbNewLine & vbNewLine
Text1 = Text1 & "Реквизиты: <Вставте Ваши реквизиты для выставления счета>"
End Sub

