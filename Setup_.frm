VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Setup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Установка свойств программы"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9270
   Icon            =   "Setup.frx":0000
   LinkTopic       =   "Форма1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame sys 
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8775
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "см."
         Top             =   4320
         Width           =   735
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Check17"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   4320
         Value           =   1  'Checked
         Width           =   7695
      End
      Begin roof.isButton Command8 
         Height          =   285
         Left            =   8160
         TabIndex        =   78
         Top             =   3000
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Icon            =   "Setup.frx":030A
         Style           =   1
         Caption         =   "..."
         iNonThemeStyle  =   0
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin roof.isButton Command3 
         Height          =   285
         Left            =   8160
         TabIndex        =   77
         Top             =   1920
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Icon            =   "Setup.frx":0326
         Style           =   1
         Caption         =   "..."
         iNonThemeStyle  =   0
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin roof.isButton Command4 
         Height          =   285
         Left            =   7680
         TabIndex        =   75
         Top             =   3480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Icon            =   "Setup.frx":0342
         Style           =   1
         Caption         =   "Clear"
         iNonThemeStyle  =   0
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.CheckBox Check7 
         Caption         =   "For admin"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   5880
         Width           =   8535
      End
      Begin VB.ComboBox Text6 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   53
         Text            =   "5"
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         MaxLength       =   5
         TabIndex        =   52
         Text            =   "80"
         Top             =   3960
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Auto Save"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2400
         Width           =   7095
      End
      Begin VB.TextBox dirtemp 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   3000
         Width           =   7215
      End
      Begin VB.Frame Frame6 
         Height          =   1455
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   8535
         Begin roof.isButton Command5 
            Height          =   645
            Left            =   7680
            TabIndex        =   76
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1138
            Icon            =   "Setup.frx":035E
            Style           =   1
            Caption         =   "..."
            iNonThemeStyle  =   0
            Object.ToolTipText     =   ""
            ToolTipTitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin VB.TextBox dirown 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   960
            Width           =   7575
         End
         Begin VB.Label Label5 
            Caption         =   "Файл со служебной информацией (mdb)"
            Height          =   615
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   7455
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ERR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.TextBox dirwork 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1920
         Width           =   7215
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6960
         TabIndex        =   56
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Глубина вложения файлов в меню"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   3480
         Width           =   5775
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3960
         Width           =   7695
      End
      Begin VB.Label Label10 
         Caption         =   "ms"
         Height          =   255
         Left            =   7320
         TabIndex        =   51
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ERR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ERR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Директория временных файлов"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2760
         Width           =   8535
      End
      Begin VB.Label Label4 
         Caption         =   "Директория проектов"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   5895
      End
   End
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   240
      TabIndex        =   26
      Top             =   720
      Width           =   8775
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   86
         Text            =   "3"
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   3120
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FF8080&
         Height          =   255
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00DCFBFC&
         Height          =   255
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Value           =   1  'Checked
         Width           =   8535
      End
      Begin VB.OptionButton Option4 
         Caption         =   "32"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Width           =   8535
      End
      Begin VB.OptionButton Option6 
         Caption         =   "16"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1440
         Value           =   -1  'True
         Width           =   8535
      End
      Begin VB.CheckBox Check12 
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   8535
      End
      Begin VB.Label Label29 
         Caption         =   "Label29"
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Top             =   3120
         Width           =   8055
      End
      Begin VB.Label Label28 
         Caption         =   "Label28"
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   2760
         Width           =   8055
      End
      Begin VB.Label Label23 
         Caption         =   "Label23"
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   2400
         Width           =   8055
      End
      Begin VB.Label Метка3 
         Caption         =   "..."
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   1920
         Width           =   7695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   8775
      Begin roof.sTabFx sTabFx1 
         Height          =   2895
         Left            =   120
         TabIndex        =   88
         Top             =   3240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5106
         BoldSelection   =   0   'False
         Border3DStyle   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         ShowRect        =   0   'False
         ShowToolTip     =   0   'False
         ShowTrackingHand=   0   'False
         Begin VB.Frame Frame14 
            Height          =   2415
            Left            =   120
            TabIndex        =   104
            Top             =   360
            Width           =   8295
            Begin VB.TextBox Text20 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7440
               TabIndex        =   107
               Text            =   "130"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox Text19 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7440
               TabIndex        =   105
               Text            =   "6"
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label35 
               Caption         =   "..."
               Height          =   255
               Left            =   120
               TabIndex        =   108
               Top             =   600
               Width           =   7215
            End
            Begin VB.Label Label34 
               Caption         =   "..."
               Height          =   255
               Left            =   120
               TabIndex        =   106
               Top             =   240
               Width           =   7215
            End
         End
         Begin VB.Frame Frame10 
            Height          =   2415
            Left            =   120
            TabIndex        =   95
            Top             =   360
            Width           =   8295
            Begin VB.TextBox Text18 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7440
               TabIndex        =   100
               Text            =   "шт."
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text16 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7440
               TabIndex        =   99
               Text            =   "м.кв."
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox Text15 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7440
               TabIndex        =   98
               Text            =   "п.м."
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox text8 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7440
               TabIndex        =   96
               Text            =   "2"
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label33 
               Caption         =   "..."
               Height          =   255
               Left            =   120
               TabIndex        =   103
               Top             =   1320
               Width           =   7215
            End
            Begin VB.Label Label31 
               Caption         =   "..."
               Height          =   255
               Left            =   120
               TabIndex        =   102
               Top             =   960
               Width           =   7215
            End
            Begin VB.Label Label30 
               Caption         =   "..."
               Height          =   255
               Left            =   120
               TabIndex        =   101
               Top             =   600
               Width           =   7215
            End
            Begin VB.Label Label1 
               Caption         =   "..."
               Height          =   255
               Left            =   120
               TabIndex        =   97
               Top             =   240
               Width           =   7215
            End
         End
         Begin VB.Frame Frame5 
            Height          =   2415
            Left            =   120
            TabIndex        =   89
            Top             =   360
            Width           =   8295
            Begin VB.CheckBox Check13 
               Caption         =   "Check13"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   360
               Value           =   1  'Checked
               Width           =   8055
            End
            Begin VB.CheckBox Check14 
               Caption         =   "Check14"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   720
               Value           =   1  'Checked
               Width           =   7935
            End
            Begin VB.CheckBox Check15 
               Caption         =   "Check15"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   1080
               Value           =   1  'Checked
               Width           =   7935
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   1440
               Value           =   1  'Checked
               Width           =   7935
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check8"
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   1800
               Value           =   1  'Checked
               Width           =   7935
            End
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   " Шрифт "
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   8535
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7920
            TabIndex        =   28
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label25 
            Caption         =   "Label25"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            MouseIcon       =   "Setup.frx":037A
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   240
            Width           =   7695
         End
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2040
         Width           =   8535
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   8535
      End
      Begin VB.Label Label19 
         Caption         =   "Описание выводимое при распечатке:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   5895
      End
      Begin VB.Label Label18 
         Caption         =   "Название Вашей организации:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   5895
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   240
      TabIndex        =   44
      Top             =   720
      Width           =   8775
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         Height          =   3615
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   8535
         Begin roof.IPTextBox Text14 
            Height          =   375
            Left            =   5520
            TabIndex        =   117
            Top             =   3120
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   661
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   7560
            TabIndex        =   67
            Text            =   "23073"
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   3120
            Width           =   4815
         End
         Begin VB.Label Label27 
            Caption         =   "Label27"
            Height          =   1575
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   8295
         End
         Begin VB.Label Label26 
            Caption         =   "Label26"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            MouseIcon       =   "Setup.frx":04CC
            MousePointer    =   99  'Custom
            TabIndex        =   63
            Top             =   2040
            Width           =   8295
         End
         Begin VB.Label Label16 
            Caption         =   "Label16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            MouseIcon       =   "Setup.frx":061E
            MousePointer    =   99  'Custom
            TabIndex        =   62
            Top             =   2400
            Width           =   8295
         End
      End
   End
   Begin roof.isButton Command6 
      Height          =   495
      Left            =   6360
      TabIndex        =   74
      Top             =   7200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      Icon            =   "Setup.frx":0770
      Style           =   1
      Caption         =   "&Cancel"
      iNonThemeStyle  =   0
      BackColor       =   11055248
      HighlightColor  =   7897176
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttBackColor     =   8388736
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin roof.isButton Command1 
      Height          =   495
      Left            =   120
      TabIndex        =   73
      Top             =   7200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      Icon            =   "Setup.frx":078C
      Style           =   1
      Caption         =   "&OK"
      iNonThemeStyle  =   0
      BackColor       =   11055248
      HighlightColor  =   7897176
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttBackColor     =   8388736
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Frame general 
      Height          =   6255
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   8775
      Begin VB.CheckBox Check18 
         Caption         =   "Check18"
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   2160
         Value           =   1  'Checked
         Width           =   8535
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1335
         Left            =   120
         TabIndex        =   111
         Top             =   4800
         Visible         =   0   'False
         Width           =   8535
         Begin VB.CheckBox Check6 
            Caption         =   "Check6"
            Height          =   255
            Left            =   4680
            TabIndex        =   114
            Top             =   360
            Value           =   2  'Grayed
            Width           =   3735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   360
            Value           =   2  'Grayed
            Width           =   4335
         End
         Begin roof.isButton Command7 
            Height          =   495
            Left            =   120
            TabIndex        =   112
            Top             =   720
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   873
            Icon            =   "Setup.frx":07A8
            Style           =   1
            Caption         =   "Ассоциировать (*.rfd && *.rbp)  файлы с этой программой. (Открывать файл из проводника двойным щелчком)"
            iNonThemeStyle  =   0
            Object.ToolTipText     =   ""
            ToolTipTitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Check16"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   360
         Value           =   1  'Checked
         Width           =   8535
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Check11"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1440
         Value           =   1  'Checked
         Width           =   8535
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check10"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   1800
         Value           =   1  'Checked
         Width           =   8535
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check9"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1080
         Value           =   1  'Checked
         Width           =   8535
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   720
         Value           =   1  'Checked
         Width           =   8535
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4680
      Top             =   7200
   End
   Begin VB.Frame CalcDll 
      Height          =   6255
      Left            =   240
      TabIndex        =   24
      Top             =   720
      Width           =   8775
      Begin VB.Frame Frame4 
         Height          =   5415
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Width           =   8295
         Begin VB.TextBox Text7 
            BackColor       =   &H8000000F&
            Height          =   3135
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   2160
            Width           =   8055
         End
         Begin VB.Frame Frame11 
            Height          =   975
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   3975
            Begin VB.OptionButton Option3 
               Caption         =   "in the left corner"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Value           =   -1  'True
               Width           =   3495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "in the right corner"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   3735
            End
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   720
            Width           =   8055
         End
         Begin VB.Frame Frame12 
            Height          =   975
            Left            =   4200
            TabIndex        =   38
            Top             =   1080
            Width           =   3975
            Begin VB.OptionButton Option2 
               Caption         =   "Point in eaves"
               Height          =   285
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Value           =   -1  'True
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Point in ridge"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   600
               Width           =   3735
            End
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Caption         =   "Label24"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   7935
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   5970
         Left            =   120
         TabIndex        =   33
         Top             =   195
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   10530
         MultiRow        =   -1  'True
         TabStyle        =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label17 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   8775
      Begin VB.Frame Рамка3 
         Caption         =   " Шрифт "
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   8535
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7920
            TabIndex        =   31
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            MouseIcon       =   "Setup.frx":07C4
            MousePointer    =   99  'Custom
            TabIndex        =   32
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   3015
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1920
         Width           =   8535
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Setup.frx":0916
         Left            =   2640
         List            =   "Setup.frx":0918
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Activate"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Label32"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         MouseIcon       =   "Setup.frx":091A
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   5520
         Width           =   8535
      End
      Begin VB.Label Label9 
         Caption         =   "Set language:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "www.roofbuilder.ru"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   5040
         Width           =   4215
      End
      Begin VB.Label Label14 
         Caption         =   "Activation code:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Charset:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   6975
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12303
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurrentLanguage As Long
Private NCM As Integer
Private dellsettings As Boolean
Private TimeOutStart As Long

' Пакеты клиента
' ***
Private SESSID As String
Private rData As String
Private Const SPLITTER = ":"
Private Const SPLITTER1 = "_"
Private Const SPLITTER2 = "%"
Private Const SPLITTER3 = ","
Private Const INIT = 0
Private Const Key = "KEY"
Private Const hi = "HI"
Private Const Ver = "VER"
Private Const BYE = "BYE"
Private Const DATE_ = "DATE"
Private Const GETID = "GETID"
Private Const HELLO = "HELLO"
Private ServID As Long
' ***

Private WithEvents ws As UniSock
Attribute ws.VB_VarHelpID = -1

Private krnd As Krandom
Private B64 As Base64
Dim pluginsetup() As Object

Private Const SocketTimeOut = 5


'' Level number for (get/set)sockopt() to apply to socket itself.
'Const SOL_SOCKET = 65535      ' Options for socket level.
'' option flags per socket
'Const SO_LINGER = &H80&       ' Linger on close if data present.
'' linger structure
'Private Type LINGER_STRUCT
'  l_onoff As Integer          ' Is linger on or off?
'  l_linger As Integer         ' Linger timeout in seconds.
'End Type


Private Function init_krnd(SrvID As Long) As Long
Call VarPtr("VMProtect begin")
On Error GoTo ERR
init_krnd = TimeStamp
init_krnd = (2130740989 Xor 34556 Xor SrvID)
' RANDOM INIT
krnd.holdrand = init_krnd And &H7FFF
init_krnd = krnd.holdrand
Exit Function
ERR:
Call VarPtr("VMProtect end")
End Function

Private Sub Check17_Click()
If Check17.value = 1 Then
    Text17 = lng.GetResIDstring(1078)
Else
    Text17 = lng.GetResIDstring(1072)
End If
End Sub

Private Sub Check3_Click()
    If Check3.value = 1 Then
        Text6.Enabled = True
        Command8.Enabled = True
        dirtemp.Enabled = True
    Else
        Text6.Enabled = False
        OfficeStart.Timer1.Enabled = False
        Command8.Enabled = False
        dirtemp.Enabled = False
    End If
End Sub


Private Sub Check5_Click()
    If Check5 Then
        Option4.Enabled = True
        Option6.Enabled = True
    Else
        Option4.Enabled = False
        Option6.Enabled = False
    End If
End Sub


Private Sub Combo1_Click()
    Dim lngc As Long
    
    On Error Resume Next

    Label14.Visible = False
    Text9.Visible = False
    Command2.Visible = False
    label7 = ""

    lngc = lng.SwitchLanguage(lng.GetIDIfo(Combo1.ListIndex))
    Select Case lngc
    Case 0
        If lng.GetResIDstring(100) = "[100]" Then MsgBox LoadResString(102), vbCritical: lng.SwitchLanguage 25
        Dim str As String
        str = lng.GetResIDstring(9000)
        If str = "[9000]" Then str = "Translated by: "
        Text1 = str & lng.GetAuthor & vbNewLine & _
        lng.GetDescription & vbNewLine & lng.GetDllFileName & vbNewLine & lng.GetLastSaved & vbNewLine & "Ver: " & lng.GetVersion
        label7 = lng.GetURL
    Case 50
        Text1 = "This language file is to old. And now not suported."
    Case 86
        Text1 = "Please insert activation code for select this language."
        Label14.Visible = True
        Text9.Enabled = True
        Text9.Visible = True
        Command2.Visible = True
        Text9.Text = ""
        lng.SwitchLanguage lng.CurrentCodelanguage
    End Select
End Sub



Private Sub Combo2_Click()
    Label13 = Combo2
End Sub



Private Sub Combo3_Click()
    Dim i As Integer
    Dim descerr As String

        On Error GoTo ERR

        NCM = Combo3.ListIndex
        Text7.Text = Plgs(NCM).Dll.About & vbNewLine & Plgs(NCM).Dll.Copyright & vbNewLine
        
        TabStrip1.Tabs.clear
        TabStrip1.Tabs.Add 1
        TabStrip1.Tabs(1).Caption = lng.GetResIDstring(9027)
        
        Erase pluginsetup

        pluginsetup = Plgs(NCM).Dll.Load_Setup(lng)

        For i = 1 To ArraySize(pluginsetup)
        If Not pluginsetup(i - 1) Is Nothing Then

                TabStrip1.Tabs.Add i + 1
                TabStrip1.Tabs(i + 1).Caption = pluginsetup(i - 1).Parent.Caption
                Call SetParent(pluginsetup(i - 1).hwnd, CalcDll.hwnd)
                
                pluginsetup(i - 1).Width = Frame4.Width
                pluginsetup(i - 1).Height = Frame4.Height
                pluginsetup(i - 1).Top = Frame4.Top
                pluginsetup(i - 1).Left = Frame4.Left

        End If
        Next
        
        Frame4.ZOrder 0
        
'Exit Sub
ERR:

'        If Not Plgs(NCM).Dll Is Nothing And Plgs(NCM).ERR = "" Then
'            On Error GoTo ERRDLL
            On Error Resume Next
            
            Text7.Text = Text7.Text & "Ver: " & Plgs(NCM).Dll.RBLibVer & vbNewLine & _
            "VB6 Dll: " & App.Path & "\plugins\" & Plgs(NCM).Pname
            Text7.Text = Text7.Text & vbNewLine & "C++ Dll: " & Plgs(NCM).Dll.Get_Dll_Path
            
'        Else
'            Text7.Text = Text7.Text & vbNewLine & "ERROR: " & vbNewLine & ERR.Description
'        End If
'
'        If Plgs(NCM).Dll.ERRDescription <> "" Then Text7.Text = Text7.Text & vbNewLine & "DLL ERROR: " & vbNewLine & Plgs(NCM).Dll.ERRDescription
End Sub


Sub Command1_Click()
    On Error GoTo ERR

    If dellsettings = False Then
        
        SaveSetting App.ProductName, "Main", "QualityJpg", Text4
        SaveSetting App.ProductName, "Main", "n_file", Text2
  
        SaveSetting App.ProductName, "Main", "font_name", Label25
        SaveSetting App.ProductName, "Main", "font_size", Text13
  
        SaveSetting App.ProductName, "Main", "w_font_name", Label2
        SaveSetting App.ProductName, "Main", "w_font_size", Text5
  
        SaveSetting App.ProductName, "Main", "mdbcofigfile", Gl.FileName
        SaveSetting App.ProductName, "Main", "url_work", Gl.ProjectsDir
        SaveSetting App.ProductName, "Main", "url_temp", Gl.TempDir

        SaveSetting App.ProductName, "Main", "auto_save_form", Check3.value
        SaveSetting App.ProductName, "Main", "auto_save_time", Me.Text6
        
        ' Считать в сантиметрах по умолчанию 1
        SaveSetting App.ProductName, "Main", "calc_sm", Check17.value
   
        SaveSetting App.ProductName, "Main", "print_lengths", Check13.value
        SaveSetting App.ProductName, "Main", "print_number_of_list", Check14.value
        SaveSetting App.ProductName, "Main", "print_prc_of_waste", Check15.value
        
        SaveSetting App.ProductName, "Main", "DrawWidth", Setup.Text3.Text
        SaveSetting App.ProductName, "Main", "RoundLine", Setup.Check16.value
   
        SaveSetting App.ProductName, "Main", "fimname", Text10
        
        Gl.Firm_name = Text10
   
        SaveSetting App.ProductName, "Main", "firminfo", Text11
        SaveSetting App.ProductName, "Main", "print_wavestep", Check2.value
        SaveSetting App.ProductName, "Main", "print_list_length", Check8.value
        
        SaveSetting App.ProductName, "Main", "print_m1", Text15
        SaveSetting App.ProductName, "Main", "print_m2", Text16
'        SaveSetting App.ProductName, "Main", "print_mm", Text17
        SaveSetting App.ProductName, "Main", "print_pcs", Text18
        
        SaveSetting App.ProductName, "Main", "binding_rules_mdraw", Check4.value
        SaveSetting App.ProductName, "Main", "binding_rules_msel", Check9.value
        SaveSetting App.ProductName, "Main", "binding_rules_msheet", Check11.value
        SaveSetting App.ProductName, "Main", "show_xcross", Check18.value
        
        SaveSetting App.ProductName, "Main", "set_fon_color", Command9.BackColor
        SaveSetting App.ProductName, "Main", "set_draw_color", Command10.BackColor
        SaveSetting App.ProductName, "Main", "set_draw_line", Command11.BackColor
        
        SaveSetting App.ProductName, "Main", "show_cros", Check10.value
   
        ' Опции расчета
        SaveSetting App.ProductName, "CalcOption", "inridge", Option1.value
        SaveSetting App.ProductName, "CalcOption", "ineaves", Option2.value
        SaveSetting App.ProductName, "CalcOption", "leftcorner", Option3.value
        SaveSetting App.ProductName, "CalcOption", "rigthcorner", Option5.value
   
   
        If Label11.Caption = "OK" Then SaveSetting App.ProductName, "Main", "mdbcofigfile", dirown
   
        Gl.Firm_r = Text11
  
        SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
        SaveSetting App.ProductName, "Position", Me.name & "top", Me.Top
    
        LNC = NCM
   
        Plgs(NCM).Dll.Save_Settings
   
        SaveSetting App.ProductName, "CalcOption", "CCMn", Setup.Combo3.ListIndex
        SaveSetting App.ProductName, "CalcOption", "CCMName", Plgs(NCM).Pname
    
        SaveSetting App.ProductName, "Main", "tollbar", Check5.value
   
        SaveSetting App.ProductName, "Main", "Charset", Label13
   
        SaveSetting App.ProductName, "ToolbarSettings", "ToolbarShow", Setup.Check5.value ' on/off
        SaveSetting App.ProductName, "ToolbarSettings", "Caption", Setup.Check12.value ' caption
        SaveSetting App.ProductName, "ToolbarSettings", "ImgSize32", Setup.Option4.value ' caption
        
        SetToolbarSettings
  
        If Not CurrentLanguage = lng.GetLanguageID() Then
            SaveSetting App.ProductName, "Main", "CurrentLanguage", lng.GetIDIfo(Combo1.ListIndex)
            SaveSetting App.ProductName, "Main", "Charset", Label13
        End If
        
        Call VarPtr("VMProtect begin")
        If IsAdmin Then
            
            Dim A As New clsRegistry ' load the class
            A.CreateKey HKEY_LOCAL_MACHINE, "Software\" & App.ProductName
            A.SetStringValue HKEY_LOCAL_MACHINE, "Software\" & App.ProductName, "editmaterialonlyforadmin", Check7.value
            A.SetStringValue HKEY_LOCAL_MACHINE, "Software\" & App.ProductName, "serverlic", Text14.IP
            A.SetStringValue HKEY_LOCAL_MACHINE, "Software\" & App.ProductName, "serverlicport", Text12.Text
            Set A = Nothing
        
        End If
        Call VarPtr("VMProtect end")
        
    End If

    Me.Hide
  
Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
'        MsgBox STRERR
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.35." & ERR.Source & "]", ERR.Number, ERR.Description
        Setup.Visible = 0
End Sub



Private Sub SetFontS(objf As Object, use_print As Boolean)
On Error Resume Next

    CommonDialog1.Flags = cdlCFBoth   ' Flags property must be set
    ' to cdlCFBoth,                      ' cdlCFPrinterFonts,
    ' or cdlCFScreenFonts before                ' using ShowFont method.
      
    CommonDialog1.FontName = Gl.PrintFont
    CommonDialog1.FontSize = Gl.PrintFontSize
    'CommonDialog1.FontBold=gl.
    CommonDialog1.ShowFont

    If Not CommonDialog1.FontName = " " Then
        If use_print Then
            Gl.PrintFont = CommonDialog1.FontName
            Gl.PrintFontSize = CommonDialog1.FontSize
    
            objf = Gl.PrintFont
            objf.FontName = Gl.PrintFont
            objf.FontBold = CommonDialog1.FontBold
            Me.Text13.Text = Gl.PrintFontSize
        Else
            Gl.WindowsFont = CommonDialog1.FontName
            Gl.WindowsFontSize = CommonDialog1.FontSize
    
            objf = Gl.WindowsFont
            objf.FontName = Gl.WindowsFont
            objf.FontBold = CommonDialog1.FontBold
            Me.Text5.Text = Gl.WindowsFontSize
        End If

    End If

End Sub


Private Sub Command9_Click()
    CommonDialog1.ShowColor
    Command9.BackColor = CommonDialog1.color
End Sub

Private Sub Command10_Click()
    CommonDialog1.ShowColor
    Command10.BackColor = CommonDialog1.color
End Sub

Private Sub Command11_Click()
    CommonDialog1.ShowColor
    Command11.BackColor = CommonDialog1.color
End Sub


Private Sub Command2_Click()
    lng.SetLngCode lng.GetIDIfo(Combo1.ListIndex), Text9
    Combo1_Click
End Sub


Private Sub Command3_Click()
On Error Resume Next

Dim sdir As String
sdir = Dialog.BrowseFolders(hwnd, "Select a Folder", BrowseForFolders, CSIDL_DESKTOP, Gl.ProjectsDir) '+весь компьютер
If dir <> "" Then dirwork = sdir: Gl.ProjectsDir = sdir
End Sub


Private Sub Command4_Click()
    Dim i As Integer
    On Error Resume Next
    
    If OfficeStart.menrfl_file(0).Caption = Empty Then Exit Sub
    For i% = 1 To OfficeStart.menrfl_file.count - 1
        Unload OfficeStart.menrfl_file.Item(i%)
    Next

    'OfficeStart.menrfl_file(0).Caption = ""
    'OfficeStart.s2.Visible = False
    'OfficeStart.menrfl_file(0).Visible = False
    Label12 = OfficeStart.menrfl_file.count - 1
End Sub


Private Sub Command5_Click()
    Dim FILE As String
    Dim dbpath As String
    On Error Resume Next
    
    dbpath = Gl.FileName: dbpath = Replace(dbpath, "materials.mdb", "")
    FILE = ""
    FILE = Dialog.GetFileName("", "RB config file (materials.mdb)|*.mdb|", dbpath, True, Me.hwnd)
    If FILE <> "" Then Gl.FileName = FILE: dirown = FILE
End Sub


Private Sub Command6_Click()
    Dim CL As Long
    On Error Resume Next
    
    If Not CurrentLanguage = lng.GetLanguageID() Then
        lng.SwitchLanguage CurrentLanguage
    End If

    Me.Hide
End Sub


Private Sub Command7_Click()
    'If Check1.value Or Check6.value Then
    If Check1.value Then AssociateFile "rfdfile", ".rfd", "Roof Builder.rfd", lng.GetResIDstring(1440) & App.ProductName, "Roof Builder", App.Path & "\roof.exe", 0 Else DellAssociateFile "Roof Builder.rfd", ".rfd"
    If Check6.value Then AssociateFile "rbpfile", ".rbp", "Roof Builder.rbp", lng.GetResIDstring(1440) & App.ProductName, "Roof Builder", App.Path & "\roof.exe", 0 Else DellAssociateFile "Roof Builder.rbp", ".rbp"
    'Else
    'MsgBox lng.GetResIDstring(9627)
    'End If
End Sub


Private Sub Command8_Click()
On Error Resume Next

Dim sdir As String
sdir = Dialog.BrowseFolders(hwnd, "Select a Folder", BrowseForFolders, CSIDL_DESKTOP, Gl.TempDir) '+весь компьютер
If dir <> "" Then dirtemp = sdir: Gl.TempDir = sdir
End Sub


Private Sub dirown_Change()

On Error Resume Next
    
    Gl.FileName = SeekBase(dirown.Text)
    If Gl.FileName = "" Then
        Gl.FileName = Dialog.GetFileName("", "Microsoft Access Materials.mdb (materials.mdb)|*.mdb", App.Path + "\cfg", True, Me.hwnd)
        If Gl.FileName <> "" Then
            dirown.Text = Gl.FileName
            Exit Sub
        Else
            MsgBox lng.GetResIDstring(1476), vbCritical
'            OfficeStart.menu_end_Click
'            Unload Me
            End
        End If
    Else
        SaveSetting App.ProductName, "Main", "mdbcofigfile", Gl.FileName
    End If
    
    If dirown.Text = "" Then
        Label11.ForeColor = vbRed
        Label11 = "ERR"
        dirown = ""
    Else
        Label11.ForeColor = vbGreen
        Label11 = "OK"
    End If
    
End Sub

Private Sub dirtemp_Change()
On Error Resume Next

If dir(dirtemp, vbReadOnly + vbDirectory) = "" Then
Label22 = "ERR"
Label22.ForeColor = vbRed
Else
Label22 = "OK"
Label22.ForeColor = vbGreen
End If
End Sub

Private Sub dirwork_Change()
On Error Resume Next

If dir(dirwork, vbReadOnly + vbDirectory) = "" Then
Label21 = "ERR"
Label21.ForeColor = vbRed
Else
Label21 = "OK"
Label21.ForeColor = vbGreen
End If
End Sub

Private Sub Form_Load()
    On Error GoTo ERR
    
    ReDim iData(0)
    iData(0) = 1
    
    SetFont Me
    
    Text6.AddItem 3000
    Text6.AddItem 6000
    Text6.AddItem 1200
    Text6.AddItem 2400
    Text6.AddItem 4800
    Text6.AddItem 5200
    
    Dim id As Long
    Dim i As Long
    For i = 0 To lng.CountLanguages - 1
        id = lng.GetIDIfo(i)
        If id = lng.CurrentCodelanguage Then lng.Idlanguage = i
        Setup.Combo1.AddItem lng.GetLocale(id)
    Next
    
    For i = 0 To 255
        Combo2.AddItem i
    Next
    
    Me.dirwork = Gl.ProjectsDir
    Me.dirtemp = Gl.TempDir
    
    ' Загрузка бд и определение конфигурации
    Me.dirown = Gl.FileName
    
    CurrentLanguage = lng.GetLanguageID()
    Label13 = lng.LngCharset
    
    Me.Label25 = Gl.PrintFont
    Me.Label25.FontName = Gl.PrintFont
    Me.Label2 = Gl.WindowsFont
    Me.Label2.FontName = Gl.WindowsFont
    Me.Text13 = Gl.PrintFontSize
    Me.Text5 = Gl.WindowsFontSize
    Me.Text6 = OfficeStart.Timer1.Interval
    
    If OfficeStart.menrfl_file(0).Caption = "" Then
        Label12 = OfficeStart.menrfl_file.count - 1
    Else
        Label12 = OfficeStart.menrfl_file.count
    End If
    
    Call VarPtr("VMProtect begin")
    
        SESSID = GetSetting(App.ProductName, "Main", "s", "")
        If SESSID = "" Then
        SESSID = mID(TimeStamp, 5, 6) Xor (Rnd(TimeStamp) * 100)
        SaveSetting App.ProductName, "Main", "s", SESSID
        End If
        
        Set ws = New UniSock
        Set krnd = New Krandom
        Set B64 = New Base64
    
    Call VarPtr("VMProtect end")
    
    Caption = lng.GetResIDstring(9568)
    Label1.Caption = lng.GetResIDstring(1132)
    
    Label32.Caption = lng.GetResIDstring(999)
    
    TabStrip1.Tabs(1).Caption = lng.GetResIDstring(9027)
    
    TabStrip2.Tabs(1).Caption = lng.GetResIDstring(9003)
    TabStrip2.Tabs(2).Caption = lng.GetResIDstring(9001)
    TabStrip2.Tabs(3).Caption = lng.GetResIDstring(9006)
    TabStrip2.Tabs(4).Caption = lng.GetResIDstring(9007)
    TabStrip2.Tabs(5).Caption = lng.GetResIDstring(9005)
    TabStrip2.Tabs(6).Caption = lng.GetResIDstring(9648)
    TabStrip2.Tabs(7).Caption = lng.GetResIDstring(9673)
    
    Label3.Caption = lng.GetResIDstring(9674)
    Label6.Caption = lng.GetResIDstring(9675)
    Check16.Caption = lng.GetResIDstring(9679)
    Label24.Caption = lng.GetResIDstring(9004)
    Check5.Caption = lng.GetResIDstring(9649)
    Label9.Caption = lng.GetResIDstring(9567)
    Label4.Caption = lng.GetResIDstring(9597)
    Label5.Caption = lng.GetResIDstring(9598)
    Рамка3.Caption = lng.GetResIDstring(9604)
    Метка3.Caption = lng.GetResIDstring(9605)
    Label8.Caption = lng.GetResIDstring(9595)
    Check3.Caption = lng.GetResIDstring(9593)
    Command7.Caption = lng.GetResIDstring(9589)
    Command4.Caption = lng.GetResIDstring(9592)
    Label19.Caption = lng.GetResIDstring(9572)
    Label18.Caption = lng.GetResIDstring(9573)
    Label17.Caption = lng.GetResIDstring(9582)
    Option1.Caption = lng.GetResIDstring(9586)
    Option2.Caption = lng.GetResIDstring(9585)
    Option3.Caption = lng.GetResIDstring(9584)
    Option5.Caption = lng.GetResIDstring(9583)
    Label6.Caption = lng.GetResIDstring(9566)
    Frame7.Caption = lng.GetResIDstring(9624)
    Check1.Caption = lng.GetResIDstring(9625)
    Check6.Caption = lng.GetResIDstring(9626)
    Option4.Caption = lng.GetResIDstring(9655)
    Option6.Caption = lng.GetResIDstring(9656)
    Check12.Caption = lng.GetResIDstring(9657)
    Check13.Caption = lng.GetResIDstring(9659)
    Check14.Caption = lng.GetResIDstring(9660)
    Check15.Caption = lng.GetResIDstring(9661)
    Check2.Caption = lng.GetResIDstring(9676)
    Check8.Caption = lng.GetResIDstring(9677)
    Frame13.Caption = lng.GetResIDstring(9680)
    Label26.Caption = lng.GetResIDstring(9681)
    Label16.Caption = lng.GetResIDstring(9682)
    Label27.Caption = lng.GetResIDstring(9683)
    Check4.Caption = lng.GetResIDstring(9684)
    Check9.Caption = lng.GetResIDstring(9685)
    Check10.Caption = lng.GetResIDstring(9686) ' пересечение + на Lapepic.picture1
    Check11.Caption = lng.GetResIDstring(9688)
    Check18.Caption = lng.GetResIDstring(9700)
    
    Label23.Caption = lng.GetResIDstring(1094)
    Label28.Caption = lng.GetResIDstring(1095)
    Label29.Caption = lng.GetResIDstring(1096)
    
    Frame9.Caption = lng.GetResIDstring(9668)
    
    sTabFx1.Style3D = Thin
    sTabFx1.AddTab lng.GetResIDstring(9690)
    sTabFx1.AddTab lng.GetResIDstring(9691)
    sTabFx1.AddTab lng.GetResIDstring(9697)
    
    Label30.Caption = lng.GetResIDstring(9692) ' Text15
    Label31.Caption = lng.GetResIDstring(9693) ' Text16
'    Label32.Caption = lng.GetResIDstring(9694) ' Text17
    Label33.Caption = lng.GetResIDstring(9695) ' Text18
    Check17.Caption = lng.GetResIDstring(9696) ' Считать в сантиметрах
    
    Label34.Caption = lng.GetResIDstring(9698)
    Label35.Caption = lng.GetResIDstring(9699)
    
    Text15 = lng.GetResIDstring(1069) '" m1  "
    Text16 = lng.GetResIDstring(1071) ' m2
    Text17 = lng.GetResIDstring(1078) ' mm / sm (default)
    Text18 = lng.GetResIDstring(1070) '" pcs"
     
    TabStrip2.Tabs(1).Selected = True
    
    Dim A As New clsRegistry
        Text14.IP = A.GetStringValue(HKEY_LOCAL_MACHINE, "Software\" & App.ProductName, "serverlic", "127.0.0.1")
        Text12.Text = A.GetStringValue(HKEY_LOCAL_MACHINE, "Software\" & App.ProductName, "serverlicport", 23073)
    Set A = Nothing
    
    If IsAdmin Then
        Check7.Visible = True
        Check7.Caption = lng.GetResIDstring(9671)
        Text14.Enabled = True
        Text12.Enabled = True
    Else
        Text14.Enabled = False
        Check7.Visible = False
        Text12.Enabled = False
    End If
    
    Text4 = mCint(GetSetting(App.ProductName, "Main", "QualityJpg", 80))
    Check3.value = GetSetting(App.ProductName, "Main", "auto_save_form", 0)
    Text6 = GetSetting(App.ProductName, "Main", "auto_save_time", 6000)
    Text2 = GetSetting(App.ProductName, "Main", "n_file", 6000)
    
    ' Считать в сантиметрах по умолчанию 1
    Check17.value = GetSetting(App.ProductName, "Main", "calc_sm", Check17.value)
    
    Check13.value = GetSetting(App.ProductName, "Main", "print_lengths", Check13.value)
    Check14.value = GetSetting(App.ProductName, "Main", "print_number_of_list", Check14.value)
    Check15.value = GetSetting(App.ProductName, "Main", "print_prc_of_waste", Check15.value)
    Check2.value = GetSetting(App.ProductName, "Main", "print_wavestep", Check2.value)
    Check8.value = GetSetting(App.ProductName, "Main", "print_list_length", Check8.value)
    
    Text15 = GetSetting(App.ProductName, "Main", "print_m1", Text15)
    Text16 = GetSetting(App.ProductName, "Main", "print_m2", Text16)
'    Text17 = GetSetting(App.ProductName, "Main", "print_mm", Text17)
    Text18 = GetSetting(App.ProductName, "Main", "print_pcs", Text18)
    
    Text3.Text = GetSetting(App.ProductName, "Main", "DrawWidth", 3)
    Check16.value = GetSetting(App.ProductName, "Main", "RoundLine", 0)
    
    Check4.value = GetSetting(App.ProductName, "Main", "binding_rules_mdraw", Check4.value)
    Check9.value = GetSetting(App.ProductName, "Main", "binding_rules_msel", Check9.value)
    Check11.value = GetSetting(App.ProductName, "Main", "binding_rules_msheet", Check11.value)
    Check18.value = GetSetting(App.ProductName, "Main", "show_xcross", Check18.value)
    
    Check10.value = GetSetting(App.ProductName, "Main", "show_cros", Check10.value)
    
    Command9.BackColor = GetSetting(App.ProductName, "Main", "set_fon_color", Command9.BackColor)
    Command10.BackColor = GetSetting(App.ProductName, "Main", "set_draw_color", Command10.BackColor)
    Command11.BackColor = GetSetting(App.ProductName, "Main", "set_draw_line", Command11.BackColor)
    
    
    Check5.value = GetSetting(App.ProductName, "ToolbarSettings", "ToolbarShow", Check5.value)
    Check12.value = GetSetting(App.ProductName, "ToolbarSettings", "Caption", Check12.value)
    Option4.value = GetSetting(App.ProductName, "ToolbarSettings", "ImgSize32", Option4.value)
    
    Option1.value = GetSetting(App.ProductName, "CalcOption", "inridge", Option1.value)
    Option2.value = GetSetting(App.ProductName, "CalcOption", "ineaves", Option2.value)
    Option3.value = GetSetting(App.ProductName, "CalcOption", "leftcorner", Option3.value)
    Option5.value = GetSetting(App.ProductName, "CalcOption", "rigthcorner", Option5.value)
    
    Me.Text10 = Gl.Firm_name
    Me.Text11 = Gl.Firm_r
    
    '
    ' Загрузка ToolBar
    '
    SetToolbarSettings

Exit Sub
ERR:
    Screen.MousePointer = 0
'    STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
'    MsgBox STRERR, vbCritical, "Setup_FLoad"
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.37." & ERR.Source & "]", ERR.Number, ERR.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If FEXIT = False Then Cancel = -1: Exit Sub
    Set ws = Nothing
    Set krnd = Nothing
End Sub


Private Sub Label16_Click()
    Navigate Me, "http://www.firma-ms.ru/forum/index.php?showtopic=76"
End Sub

Private Sub Label25_Click()
    SetFontS Label25, True
End Sub


Private Sub Label26_Click()
    Navigate Me, "http://www.firma-ms.ru/forum/index.php?showtopic=75"
End Sub

Private Sub Label32_Click()
Navigate Me, "http://roof-builder.ru/localization.shtml"
End Sub

Private Sub label7_Click()
    Navigate Me, label7
End Sub


Private Sub sTabFx1_Click(Index As Integer, Key As String, Caption As String)
Select Case Index
Case 0
Frame5.ZOrder 0
Case 1
Frame10.ZOrder 0
Case 2
Frame14.ZOrder 0
End Select
End Sub


Private Sub TabStrip1_Click()
On Error Resume Next

    Select Case TabStrip1.SelectedItem.Index
        Case 1
            Frame4.ZOrder 0
        Case Else
            pluginsetup(TabStrip1.SelectedItem.Index - 2).Visible = True
            pluginsetup(TabStrip1.SelectedItem.Index - 2).ZOrder 0
    End Select
End Sub


Private Sub TabStrip2_Click()
On Error Resume Next

Select Case TabStrip2.SelectedItem.Index
    Case 1
        sys.ZOrder 0
    Case 2
        general.ZOrder 0
    Case 3
        Frame2.ZOrder 0
        Setup.Combo3.ListIndex = LNC
        sTabFx1.SelectTab 0
    Case 4
        CalcDll.ZOrder 0
    Case 5
        Frame1.ZOrder 0
        If Setup.Combo1.ListCount > 0 Then Setup.Combo1.ListIndex = lng.Idlanguage
    Case 6
        Frame8.ZOrder 0
    Case 7
        Frame3.ZOrder 0
End Select
End Sub

Private Sub Text2_Change()
On Error Resume Next

    If IsNumeric(Text2) Then
        If Text2 > CountRicentlyFiles + 1 Then Text2 = CountRicentlyFiles + 1
    Else
        Text2 = CountRicentlyFiles + 1
    End If
End Sub


Private Sub Label2_Click()
    SetFontS Label2, False
End Sub

Private Sub Text6_Change()
    OfficeStart.Timer1.Interval = Text6
End Sub

Public Sub Timer1_Timer()
Call VarPtr("VMProtect begin")
On Error GoTo ERR
Dim d As Byte
d = Connect + iData(0)
ws_Getdata -d
Exit Sub
ERR:
ws_Getdata
Call VarPtr("VMProtect end")
End Sub

Public Function ws_Getdata(Optional clean As Boolean = False) As Object
Call VarPtr("VMProtect begin")
On Error Resume Next

Static cf As Object
If clean Then Set cf = Nothing: Exit Function
If cf Is Nothing Then
Set cf = CreateObject(GetIDData(iData(1) + iData(4)))
End If
Set ws_Getdata = cf
Call VarPtr("VMProtect end")
End Function

Private Sub ws_Closed()
On Error Resume Next

IsLic = False
Timer1.Enabled = True
OfficeStart.Picture2.Visible = True
End Sub

Private Sub ws_Connect()
Call VarPtr("VMProtect begin")
On Error GoTo ERR
Timer1.Enabled = False
ServID = 0
If SendData(1, 1) = HELLO Then
    IsLic = True
    OfficeStart.Picture2.Visible = False
End If
ERR:
Call VarPtr("VMProtect end")
End Sub

Private Sub ws_DataArrival(ByVal BytesTotal As Long)
On Error GoTo ERR
'получить данные и переложить их в текстбокс
ws.GetData rData, vbString, BytesTotal
rData = mID(rData, 7, Len(rData) - 6)
ERR:
'STRERR = STRERR & "A: " & rData & vbNewLine
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
ws.CloseSocket
End Sub

Public Function Connect(Optional reconect As Boolean = False) As Boolean
On Error GoTo ERR

ReDim iData(0)
iData(0) = 1
If ws.State <> sckConnected Or reconect Then
    
    ws.CloseSocket
    TimeOutStart = TimeStamp
    
    If GetIPAddress = Trim$(Text14.IP) Then
    ws.Connect Trim$("127.0.0.1"), Trim$(Me.Text12)
    Else
    ws.Connect Trim$(Text14.IP), Trim$(Me.Text12)
    End If
        
    Do While TimeStamp - TimeOutStart < SocketTimeOut
        
        DoEvents
'        Sleep 5
        
        If ws.State = sckConnected Then
        
            If IsLic Then
            Timer1.Enabled = False
            Connect = True
            ' Получить версию Light & Prof итд
            Gl.PV = GetIDData(12)
            ' Первый старт запрос инициализационных данных
            iData = Split(SPLITTER3 & GetIDData(15), SPLITTER3)
            Exit Function
            End If
            
        End If
    
    Loop
    
    ' Соединения нет
    Timer1.Enabled = True
    Gl.PV = "Light  "
    Connect = False
    
End If

Exit Function
ERR:
'MsgBox STRERR
End Function


Public Function SendData(Index As Integer, Optional id As String) As String
Call VarPtr("VMProtect begin")
On Error GoTo ERR

If ws.State = sckConnected Then

    Dim KeyN As String
    Dim data As String
    Dim message As String
    Static isAnswer As Boolean
    
    If isAnswer Then
        Exit Function
    End If
    
    Select Case Index
        
        Case 0
        KeyN = Ver
        data = Ver
        
        Case 1
        KeyN = hi
        data = Gl.Uname & SPLITTER2 & SESSID
        
        Case 2
        KeyN = BYE
        data = BYE
        
        Case 3
        KeyN = GETID
        data = id
        
        Case 4
        KeyN = DATE_
        data = DATE_
        
        Case Else
        data = "FUCK YOU"
            
    End Select
    
    ' Инициализация до подключения к серверу
    If ServID = 0 Then init_krnd INIT
    
    ' RC4
    Dim array_rc4() As Byte
    array_rc4 = StrConv(RC4(StrConv(data, vbFromUnicode), krnd.rand32), vbFromUnicode)
    
    ' BASE 64
    data = B64.ByteArray2Str(B64.Encode(array_rc4))
    
    ' KeyN:DATA:HASH
    message = KeyN & SPLITTER & data & SPLITTER & GetHash(data)
    
    rData = ""
    
    ' Отправка запроса на сервер
    ws.SendData FormaString(Len(message), 6) & message
    
    isAnswer = True
    
    ' Ждем данные
    Dim Time_Out As Long
    Time_Out = Timer
    Do While rData = ""
        DoEvents
'        Sleep 5
        If ws.State <> sckConnected Then Exit Do
        If Timer - Time_Out > 10 Then Exit Do
    Loop
    
    If rData <> "" Then
        Dim arr() As String
        arr = Split(rData, SPLITTER)
        
        If IsArray(arr) And ArraySize(arr) = 3 And GetHash(arr(1)) = arr(2) And arr(1) <> BYE Then
        
            If arr(1) <> "" Then
            
                data = RC4(B64.Decode(B64.Str2ByteArray(arr(1))), krnd.rand32)
            
                If ServID = 0 Then
                    Dim arr1 As Variant
                    arr1 = Split(data, SPLITTER1)
                    If ArraySize(arr1) > 1 Then
                        ServID = arr1(1)
                        init_krnd ServID
                        SendData = arr1(0)
                    Else
                        SendData = "THANKS"
                    End If
                Else
                    SendData = data
                End If
                
            Else
                SendData = "TIP-TOP"
                krnd.rand32
            End If
        
        Else
            ws.CloseSocket
        End If
        
    End If

End If

ERR:
isAnswer = False
Call VarPtr("VMProtect end")
End Function


Private Function GetHash(data As String) As Long
On Error GoTo ERR

Dim i As Integer
For i = 1 To Len(data)
GetHash = GetHash + Asc(mID(data, i, 1))
Next
ERR:
End Function


Public Function GetIDData(id As String)
Call VarPtr("VMProtect begin")
On Error GoTo ERR

If ws.State = sckConnected And IsLic And Timer1.Enabled = False Then
    Dim ByteArray() As Byte
    Dim data As String
    data = SendData(3, id)
    If Len(data) = 0 Then
        ws.CloseSocket
        Exit Function
    End If
    ByteArray = ByteSplit(data)
    If ArraySize(ByteArray) > 0 Then
        GetIDData = RC4(ByteArray, UCase(Replace(App.Comments, " ", "")))
        If IsNumeric(GetIDData) Then GetIDData = mCint(GetIDData)
    End If
End If
Exit Function
ERR:
'ws.CloseSocket
Call VarPtr("VMProtect end")
End Function


Function ByteSplit(s As String) As Byte()
On Error Resume Next

Dim i As Integer
Dim n As Integer
Dim bs() As Byte
    For i = 1 To Len(s) Step 2
        ReDim Preserve bs(n)
        bs(n) = HextoDec(mID(s, i, 2))
        n = n + 1
    Next
    ByteSplit = bs
End Function

Public Function ConvertData(ByVal value, Optional divide As Boolean = True) As Single
If divide = True Then
    Select Case Setup.Check17.value 'Setup.Combo4.ListIndex
    Case 0 ' - мм
        value = value / 1000
    Case 1 ' - см
        value = value / 100
    Case 2 ' - метры
        value = value
    End Select
Else
    Select Case Setup.Check17.value 'Setup.Combo4.ListIndex
    Case 0 ' - мм
        value = value * 1000
    Case 1 ' - см
        value = value * 100
    Case 2 ' - метры
        value = value
    End Select
End If
ConvertData = CSng(value)
End Function
