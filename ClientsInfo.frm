VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ClientsInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client`s info"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "ClientsInfo.frx":0000
      Top             =   2040
      Width           =   5535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Text            =   "12.12.2004"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "00"
      Top             =   720
      Width           =   1095
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5640
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5640
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Done projects"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   5535
   End
   Begin VB.Label Label6 
      Caption         =   "Contact`s information:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label5 
      Caption         =   "Addres:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Customer`s Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Date of addition:"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Client info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Customer`s ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "ClientsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
