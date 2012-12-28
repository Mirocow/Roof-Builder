VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SetMaterial 
   BackColor       =   &H00808000&
   Caption         =   "Set material"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   4440
      TabIndex        =   30
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chanel"
      Height          =   495
      Left            =   5760
      TabIndex        =   29
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Height          =   315
      Index           =   0
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   1140
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Height          =   315
      Index           =   1
      Left            =   3000
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   1140
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Height          =   315
      Index           =   2
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1140
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Height          =   315
      Index           =   3
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox txtW 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "0"
      Top             =   1575
      Width           =   735
   End
   Begin VB.TextBox txtS 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Text            =   "0"
      Top             =   2295
      Width           =   735
   End
   Begin VB.TextBox txtWG 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "0"
      Top             =   1935
      Width           =   735
   End
   Begin VB.TextBox txtO 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "0"
      Top             =   2655
      Width           =   735
   End
   Begin VB.TextBox txtMinl 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Text            =   "0"
      Top             =   1575
      Width           =   735
   End
   Begin VB.TextBox txtMaxl 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Text            =   "0"
      Top             =   1935
      Width           =   735
   End
   Begin VB.ComboBox lstprof 
      BackColor       =   &H00DCFBFC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   330
      Width           =   6735
   End
   Begin VB.TextBox txtH 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "0"
      Top             =   3015
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Text            =   "0"
      Top             =   3375
      Width           =   615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   5040
      TabIndex        =   0
      Top             =   2175
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   192
      BackColor       =   14482428
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Length"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "quantity"
         Object.Width           =   1147
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Метка1"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   855
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Метка1"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      TabIndex        =   27
      Top             =   855
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Метка1"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   855
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Метка2"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   855
      Width           =   1215
   End
   Begin VB.Label Метка1 
      Caption         =   "Working width:"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   1575
      Width           =   1695
   End
   Begin VB.Label Метка2 
      Caption         =   "Step:"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   2295
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "General width:"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   1935
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "Overlapping"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   2655
      Width           =   1575
   End
   Begin VB.Label Label22 
      Caption         =   "Min lenght"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   1575
      Width           =   1335
   End
   Begin VB.Label Label23 
      Caption         =   "Max lenght"
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   1935
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Select profil"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label9 
      Caption         =   "Generl Height"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3015
      Width           =   1575
   End
   Begin VB.Label Label21 
      Caption         =   "Cost"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3375
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   3120
      Left            =   240
      Top             =   735
      Width           =   6735
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DCFBFC&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4560
      TabIndex        =   15
      Top             =   3375
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Label20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   5040
      TabIndex        =   14
      Top             =   1575
      Width           =   1815
   End
End
Attribute VB_Name = "SetMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
