VERSION 5.00
Object = "{D4055E62-5507-43CA-B528-924FB94C4FF4}#1.0#0"; "SplitterModern.ocx"
Object = "{433FA3B6-0E64-4034-BAA7-A8B879261CD5}#1.0#0"; "acSR.ocx"
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form Lapepic 
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   17745
   ControlBox      =   0   'False
   FillStyle       =   4  'Upward Diagonal
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "Lapepic.frx":0000
   LinkTopic       =   "Форма1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   17745
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin СhameleonButton.chameleonButton Check11 
      Height          =   375
      Left            =   0
      TabIndex        =   70
      Top             =   360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Lapepic.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
      ICONS           =   16
   End
   Begin acSR.SuperRuler SuperRuler2 
      Height          =   4410
      Left            =   0
      Top             =   720
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   7779
      Orientation     =   1
      MouseTrackingOn =   -1  'True
      BackColor       =   -2147483633
   End
   Begin acSR.SuperRuler SuperRuler1 
      Height          =   375
      Left            =   485
      Top             =   360
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   661
      MouseTrackingOn =   -1  'True
      BackColor       =   -2147483633
   End
   Begin SplitterHV.SplitHV SplitHV1 
      Height          =   5175
      Left            =   12840
      TabIndex        =   9
      Top             =   0
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   9128
      SplitLimit      =   4880
   End
   Begin SplitterHV.SplitHV SplitHV2 
      Height          =   75
      Left            =   0
      TabIndex        =   6
      Top             =   5160
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   132
      SplitLimit      =   3100
      Style           =   0
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   4935
   End
   Begin VB.Frame cutfrm 
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   12960
      TabIndex        =   3
      Top             =   360
      Width           =   4935
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Разрез листов в шахматном порядке"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   50
         TabIndex        =   23
         Top             =   5760
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5415
         Left            =   50
         TabIndex        =   11
         Top             =   0
         Width           =   4815
         Begin VB.CheckBox IsSetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   4440
            TabIndex        =   89
            Top             =   2340
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox IsSetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   88
            Top             =   2020
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox IsSetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   4440
            TabIndex        =   87
            Top             =   1710
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox IsSetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   86
            Top             =   1400
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox IsSetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   4440
            TabIndex        =   85
            Top             =   1080
            Value           =   1  'Checked
            Width           =   255
         End
         Begin СhameleonButton.chameleonButton chameleonButton3 
            Height          =   855
            Left            =   4080
            TabIndex        =   84
            Top             =   120
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1508
            BTYPE           =   7
            TX              =   "..."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":0326
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton cange_max_size 
            Height          =   300
            Index           =   0
            Left            =   3000
            TabIndex        =   71
            Top             =   2340
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            BTYPE           =   7
            TX              =   "<"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":0342
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            Caption         =   "Check7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   3000
            Width           =   4455
         End
         Begin VB.OptionButton calc_type 
            Caption         =   "Option7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   33
            Top             =   3360
            Width           =   4695
         End
         Begin VB.OptionButton calc_type 
            Caption         =   "Option6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   32
            Top             =   2715
            Value           =   -1  'True
            Width           =   4695
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   1245
            Left            =   0
            TabIndex        =   25
            Top             =   3480
            Width           =   4720
            Begin VB.ComboBox Combo1 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3840
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Tag             =   "0"
               Top             =   600
               Width           =   855
            End
            Begin VB.CheckBox Check6 
               Appearance      =   0  'Flat
               Caption         =   "Check6"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   240
               TabIndex        =   29
               Top             =   960
               Width           =   4455
            End
            Begin VB.ComboBox txt_CL 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3840
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Tag             =   "0"
               Top             =   200
               Width           =   855
            End
            Begin VB.OptionButton Check1 
               Caption         =   "Check1"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   0
               TabIndex        =   27
               Top             =   120
               Width           =   3735
            End
            Begin VB.OptionButton Check5 
               Caption         =   "Check5"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   0
               TabIndex        =   26
               Top             =   480
               Width           =   2895
            End
            Begin VB.Label Label1 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3000
               TabIndex        =   31
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.TextBox SetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3720
            TabIndex        =   16
            Text            =   "0"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox SetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3720
            TabIndex        =   15
            Text            =   "0"
            Top             =   1400
            Width           =   615
         End
         Begin VB.TextBox SetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   3720
            TabIndex        =   14
            Text            =   "0"
            Top             =   1710
            Width           =   615
         End
         Begin VB.TextBox SetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   3720
            TabIndex        =   13
            Text            =   "0"
            Top             =   2340
            Width           =   615
         End
         Begin VB.TextBox SetProfilData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   3720
            TabIndex        =   12
            Text            =   "0"
            Top             =   2020
            Width           =   615
         End
         Begin VB.PictureBox Picture2 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   72
            Top             =   0
            Width           =   0
         End
         Begin СhameleonButton.chameleonButton cange_max_size 
            Height          =   300
            Index           =   1
            Left            =   3315
            TabIndex        =   73
            Top             =   2340
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            BTYPE           =   7
            TX              =   ">"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":035E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   0
            TabIndex        =   83
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Нахлест"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Шаг волны"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   1400
            Width           =   3615
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Высота панели"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   1680
            Width           =   3615
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Мин длина"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Макс длина"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   2340
            Width           =   2895
         End
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Text            =   "0"
         Top             =   6120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "Разрез листов в шахматном порядке"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   50
         TabIndex        =   4
         Top             =   6120
         Visible         =   0   'False
         Width           =   3690
      End
   End
   Begin VB.TextBox label7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      MouseIcon       =   "Lapepic.frx":037A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   0
      Width           =   12855
   End
   Begin VB.Frame frame1 
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   0
      TabIndex        =   1
      Top             =   5160
      Width           =   12915
      Begin VB.TextBox Label2 
         BackColor       =   &H00DCFBFC&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   100
         Locked          =   -1  'True
         MouseIcon       =   "Lapepic.frx":04CC
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2280
         Width           =   12735
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   4920
         TabIndex        =   10
         Top             =   0
         Width           =   1455
         Begin СhameleonButton.chameleonButton Command5 
            Height          =   405
            Left            =   520
            TabIndex        =   68
            Top             =   600
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":061E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Комманда4 
            Height          =   405
            Left            =   525
            TabIndex        =   64
            Top             =   1000
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":063A
            PICN            =   "Lapepic.frx":0656
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   300
            LargeChange     =   100
            Left            =   20
            Max             =   6400
            Min             =   100
            SmallChange     =   100
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1800
            Value           =   1000
            Width           =   1365
         End
         Begin СhameleonButton.chameleonButton Комманда3 
            Height          =   405
            Left            =   120
            TabIndex        =   65
            Top             =   600
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":0AA8
            PICN            =   "Lapepic.frx":0AC4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Комманда1 
            Height          =   405
            Left            =   525
            TabIndex        =   66
            Top             =   200
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":0F16
            PICN            =   "Lapepic.frx":0F32
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Комманда2 
            Height          =   405
            Left            =   930
            TabIndex        =   67
            Top             =   600
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BTYPE           =   7
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":1384
            PICN            =   "Lapepic.frx":13A0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin VB.Label Label11 
            Height          =   495
            Left            =   0
            TabIndex        =   81
            Top             =   1680
            Width           =   1455
         End
      End
      Begin roof.sTabFx sTabFx1 
         Height          =   2055
         Left            =   6480
         TabIndex        =   49
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3625
         BoldSelection   =   0   'False
         Border3DStyle   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         ShowRect        =   0   'False
         ShowToolTip     =   0   'False
         ShowTrackingHand=   0   'False
         Begin VB.Frame Frame11 
            Height          =   1575
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   6255
            Begin VB.HScrollBar HScroll2 
               Height          =   300
               Left            =   120
               Max             =   360
               TabIndex        =   90
               Top             =   960
               Value           =   1
               Width           =   4215
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   600
               Width           =   4215
            End
         End
         Begin VB.Frame Frame8 
            Height          =   1575
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   6255
            Begin VB.OptionButton Check3 
               Caption         =   "Points A-B"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Value           =   -1  'True
               Width           =   3255
            End
            Begin VB.OptionButton Check9 
               Caption         =   "Line A-B"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   480
               Width           =   3255
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Point"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   720
               Width           =   3255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check8"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   1200
               Width           =   6015
            End
         End
         Begin VB.Frame Frame9 
            Height          =   1575
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Visible         =   0   'False
            Width           =   6255
            Begin VB.CheckBox Check10 
               Caption         =   "Check10"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   1200
               Width           =   6015
            End
            Begin VB.CheckBox Check13 
               Caption         =   "Check13"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   960
               Width           =   5295
            End
            Begin СhameleonButton.chameleonButton isButton1 
               Height          =   315
               Left            =   2760
               TabIndex        =   62
               Top             =   240
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               BTYPE           =   7
               TX              =   "Command13"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Lapepic.frx":17F2
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
               ICONS           =   16
            End
            Begin СhameleonButton.chameleonButton Command1 
               Height          =   315
               Left            =   2760
               TabIndex        =   63
               Top             =   600
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               BTYPE           =   7
               TX              =   "Command13"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Lapepic.frx":180E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
               ICONS           =   16
            End
            Begin СhameleonButton.chameleonButton chameleonButton1 
               Height          =   315
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               BTYPE           =   7
               TX              =   "F5"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Lapepic.frx":182A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   -1  'True
               VALUE           =   0   'False
               ICONS           =   16
            End
            Begin СhameleonButton.chameleonButton chameleonButton2 
               Height          =   315
               Left            =   120
               TabIndex        =   80
               Top             =   600
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               BTYPE           =   7
               TX              =   "F6"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Lapepic.frx":1846
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   -1  'True
               VALUE           =   0   'False
               ICONS           =   16
            End
         End
         Begin VB.Frame Frame7 
            Height          =   1575
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   6255
            Begin СhameleonButton.chameleonButton Command2 
               Height          =   315
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               BTYPE           =   2
               TX              =   ""
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Lapepic.frx":1862
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
               ICONS           =   16
            End
            Begin СhameleonButton.chameleonButton Command13 
               Height          =   315
               Left            =   120
               TabIndex        =   61
               Top             =   600
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               BTYPE           =   7
               TX              =   "Command2"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Lapepic.frx":187E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
               ICONS           =   16
            End
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2175
         Left            =   120
         TabIndex        =   35
         Top             =   0
         Width           =   4815
         Begin СhameleonButton.chameleonButton Command7 
            Height          =   300
            Left            =   3880
            TabIndex        =   74
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            BTYPE           =   7
            TX              =   "*-1"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":189A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Command10 
            Height          =   300
            Left            =   1560
            TabIndex        =   69
            Top             =   1320
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   529
            BTYPE           =   7
            TX              =   "A"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":18B6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   300
            Left            =   1560
            Max             =   360
            TabIndex        =   59
            Top             =   840
            Value           =   1
            Width           =   2325
         End
         Begin VB.TextBox txt_step 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   41
            Text            =   "1"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2880
            TabIndex        =   40
            Text            =   "0"
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   480
            TabIndex        =   39
            Text            =   "0"
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Метка2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   480
            TabIndex        =   38
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Label8 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "0"
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2880
            TabIndex        =   36
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin СhameleonButton.chameleonButton Command9 
            Height          =   300
            Left            =   4280
            TabIndex        =   75
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            BTYPE           =   7
            TX              =   "0"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":18D2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Command3 
            Height          =   300
            Left            =   1480
            TabIndex        =   76
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            BTYPE           =   7
            TX              =   "*-1"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":18EE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Command8 
            Height          =   300
            Left            =   1880
            TabIndex        =   77
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            BTYPE           =   7
            TX              =   "0"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":190A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin СhameleonButton.chameleonButton Command12 
            Height          =   300
            Left            =   1480
            TabIndex        =   78
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            BTYPE           =   7
            TX              =   "*-1"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Lapepic.frx":1926
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            ICONS           =   16
         End
         Begin VB.Label Label9 
            Caption         =   "Метка2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   1335
            Width           =   1455
         End
         Begin VB.Label Метка4 
            Caption         =   "Y ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   47
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "Град"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   46
            Top             =   860
            Width           =   375
         End
         Begin VB.Label Метка1 
            Caption         =   "X ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Шаг"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   44
            Top             =   1755
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "A - B"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "(A - B)/2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   42
            Top             =   480
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DCFBFC&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   480
      ScaleHeight     =   4395
      ScaleMode       =   0  'User
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   720
      Width           =   12375
      Begin VB.Line line_guidings 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         Index           =   0
         Visible         =   0   'False
         X1              =   2040
         X2              =   6840
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   5640
         X2              =   8520
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   6960
         X2              =   6960
         Y1              =   720
         Y2              =   3000
      End
      Begin VB.Line Line2 
         BorderStyle     =   2  'Dash
         Visible         =   0   'False
         X1              =   840
         X2              =   4200
         Y1              =   2520
         Y2              =   2040
      End
      Begin VB.Image Image1 
         Height          =   195
         Left            =   6240
         Top             =   5880
         Width           =   210
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   960
         X2              =   3720
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line14 
         BorderColor     =   &H0073E13E&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   960
         X2              =   3720
         Y1              =   1320
         Y2              =   1320
      End
   End
End
Attribute VB_Name = "Lapepic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const colorstr = vbBlack

Private Const TIE_DISTANCE = 20
Public ListSelected As Boolean
'Public Error As Boolean
Private Opt As String

' Данные конфигурационного файла
Private CurrentPDataRS As Recordset

Dim Xtemp As Single
Dim Ytemp As Single

Private SinA As Single

Private Llen As Single

Dim Y_cut As Single

'**************************
'*** УПРАВЛЯЮЩИЕ
'**************************
Private FindPoint As Integer ' Найденная точка
Private Current_P As Integer ' Выбранная точка
Private Current_L As Integer ' Выбранная линия


Public Cut_N As Integer

Dim NERROR As Integer

Private ResizePic As Boolean
Private isPushCalc As Boolean

Private L1 As Integer
Private L2 As Integer
Private wl As Integer
Private StandartLen() As Long

Dim XMin As Single
Dim XMax As Single
Dim YMin As Single
Dim YMAx As Single

Const PI = 3.14159265358979

Private ab As Single

Private isClickMouse As Boolean

Private Sub calc_type_Click(Index As Integer)
On Error Resume Next
    
    If Index = 1 Then
    
        Check1.Enabled = True
        Check5.Enabled = True
        Check6.Enabled = True
        Check7.Enabled = False
        
    Else
    
        Check1.Enabled = False
        Check5.Enabled = False
        txt_CL.Enabled = False
        Combo1.Enabled = False
        Check7.Enabled = True
        
        Check1.value = 0
        Check5.value = 0
        Check6.value = 0
        Check6.Enabled = False
        
    End If
End Sub

Private Sub cange_max_size_Click(Index As Integer)
Dim waves As Integer
On Error Resume Next

If Val(SetProfilData(1).Tag) = 0 Then Exit Sub


' Устанавливаем флаг изменения данных
SetChange True

If Index = 0 Then

    If mCint(SetProfilData(4).Tag) - mCint(SetProfilData(1).Tag) - SetProfilData(0).Tag < CurrentPDataRS!MIN_LENGTH Then
            waves = CurrentPDataRS!MIN_LENGTH \ SetProfilData(1).Tag
            SetProfilData(4).Tag = waves * SetProfilData(1).Tag + SetProfilData(1).Tag + SetProfilData(0).Tag
    Else
    SetProfilData(4).Tag = mCint(SetProfilData(4).Tag) - SetProfilData(1).Tag
    End If

Else

    If mCint(SetProfilData(4).Tag) + mCint(SetProfilData(1).Tag) > CurrentPDataRS!MAX_LENGTH Then
        If mCint(SetProfilData(4).Tag) > 0 Then
            waves = CurrentPDataRS!MAX_LENGTH \ SetProfilData(1).Tag
            SetProfilData(4).Tag = waves * SetProfilData(1).Tag + mCint(SetProfilData(0).Tag)
        End If
    Else
    SetProfilData(4).Tag = mCint(SetProfilData(4).Tag) + SetProfilData(1).Tag
    End If
    
End If

isPushCalc = True
    
Command2_Click
End Sub

Private Sub chameleonButton1_Click()
On Error Resume Next

If chameleonButton1.value = False Then
    Unload SetPoint
Else
    SetPoint.Show vbModal, OfficeStart
End If
End Sub

Private Sub chameleonButton2_Click()
On Error Resume Next

If chameleonButton2.value = False Then
    Unload SetPoint
Else
    SetPoint.Show vbModal, OfficeStart
End If
End Sub

Private Sub chameleonButton3_Click()

Load ChangeProfil

ChangeProfil.lstprof.ListIndex = 0
ChangeProfil.lstprof_Click
ChangeProfil.Show vbModal, OfficeStart

    If ChangeProfil.Label24.Tag <> 0 Then
    ' Factory
    Label5.Caption = ChangeProfil.Label24.Caption
    Label5.Tag = ChangeProfil.Label24.Tag
    End If
    
    If ChangeProfil.ListView1.ListItems.Count > 0 Then
    ' Profil
    Label4.Caption = ChangeProfil.ComboBox1.List(ChangeProfil.ComboBox1.ListIndex)
    Label4.Tag = ChangeProfil.ListView1.ListItems(1).SubItems(1)
    End If

Unload ChangeProfil

If Label4.Tag <> "" Or Label5.Tag <> "" Then
SetChange True
SwitchProfile
HistoryClear True
End If
End Sub

Private Sub Check1_Click()
    Dim arr_wd() As Long, arr_am() As Long
    
    On Error Resume Next
    
    Combo1.Enabled = False
    txt_CL.Clear
    GetWarehouseLength Label4.Caption, GetFactoryID(Label5), True, arr_wd, arr_am

    If Check1.value Then
        On Error GoTo arrclear
        If UBound(arr_wd) >= 0 Then
            If arr_wd(0) > 0 Then
                Check6.Enabled = True
                Check5.value = 0
                txt_CL.Enabled = True
                Check6.Enabled = True
                Plgs(Gl.LNC).Dll.InputWarehouseData arr_wd, arr_am
            Else
                GoTo arrclear
            End If

        Else
arrclear:
            MsgBox lng.GetResIDstring(1489, "%Profil%", Label4.Caption), vbCritical
            Check1.value = 0
            txt_CL.Enabled = False
            Check6.Enabled = False
        End If

    Else
        Check6.value = 0
        Check6.Enabled = False
        txt_CL.Enabled = False
    End If

End Sub

Private Sub Check2_Click()
On Error Resume Next

    If Check2.value = 0 Then Check4.Enabled = False: Text6.Enabled = False Else Check4.Enabled = True: Text6.Enabled = True
End Sub


Private Sub Check3_Click()
On Error Resume Next

    ValueFieldsOnOff False

    If Check3.value Then
        P_A = 0
        P_B = 0
        Draw_Systems Me.Picture1
    End If
End Sub


Private Sub Check5_Click()
On Error Resume Next

Dim PDataRS As Recordset
txt_CL.Enabled = False
Combo1.Clear
Combo1.AddItem 0
Dim arr_am() As Long
Set PDataRS = RequestSQL("select * from ProfilsWrongLength p where p.idname=" & GetProfilID(Label4.Caption, GetFactoryID(Label5)) & " order by length")
If Not PDataRS Is Nothing Then
    
    Do While Not PDataRS.EOF
        Dim i As Integer
        ReDim Preserve StandartLen(i)
        StandartLen(i) = PDataRS.Fields(2)
        Combo1.AddItem ConvertData(StandartLen(i))
        ReDim Preserve arr_am(i)
        arr_am(i) = 100
        PDataRS.MoveNext
        i = i + 1
    Loop

    PDataRS.Close
End If

Set PDataRS = Nothing
    
If Check5.value Then
    On Error GoTo arrclear
    If UBound(StandartLen) >= 0 Then
        If StandartLen(0) > 0 Then
            '    Check6.Enabled = True
            '    Check1.value = 0
            Combo1.Enabled = True
            Plgs(Gl.LNC).Dll.InputWarehouseData StandartLen, arr_am
        Else
            GoTo arrclear
        End If

    Else
arrclear:
        MsgBox lng.GetResIDstring(1488, "%Profil%", Label4.Caption), vbCritical
        Check5.value = 0
        Combo1.Enabled = False
    End If

Else
    'Check6.value = 0
    'Check6.Enabled = False
    Combo1.Enabled = False
End If

End Sub


Private Sub Check6_Click()
    Draw_Systems Me.Picture1
End Sub

Private Sub Check7_Click()
On Error Resume Next

Dim PDataRS As Recordset
If Check7.value = 1 Then
    Set PDataRS = RequestSQL("select * from ProfilsWLength p where p.idname=" & GetProfilID(Label4.Caption) & " order by length1")
    Dim i As Integer
    i = -1
    If Not PDataRS Is Nothing Then
        Do While Not PDataRS.EOF
            If CheckNullNomber(PDataRS.Fields(4), False) Then
                i = i + 1
                ReDim Preserve WrongLs(i)
                WrongLs(i).MIN = PDataRS.Fields(2)
                WrongLs(i).MAX = PDataRS.Fields(3)
            End If

            PDataRS.MoveNext
        Loop

        PDataRS.Close
Else
    Check7.value = 0
End If

If i = -1 Then MsgBox lng.GetResIDstring(1501, "%Profil%", Label4.Caption), vbCritical
End If
End Sub


Private Sub Check9_Click()

    ValueFieldsOnOff False
    
    If Check9.value Then
        P_A = 0
        P_B = 0
    End If
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    On Error Resume Next
    
    ' Удаление чертежа
    If SlP(N_Slope).CountOfLines > SlP(N_Slope).CountOfPoints Then SlP(N_Slope).CountOfPoints = SlP(N_Slope).CountOfLines
    For i = 1 To SlP(N_Slope).CountOfPoints Step 1
        Lape_Points_X(N_Slope, i) = 0
        Lape_Points_Y(N_Slope, i) = 0
        Lape_Lines(N_Slope, i, 0) = 0
        Lape_Lines(N_Slope, i, 1) = 0
    Next i
    SlP(N_Slope).CountOfPoints = 0
    SlP(N_Slope).CountOfLines = 0

    SlP(N_Slope).Pn_Red_lines = 0
'    Option1.value = 1
    sTabFx1.SelectTab 0
    Draw_Systems Me.Picture1
End Sub


Private Sub Command10_Click()
On Error Resume Next

If Command10.value = False Then
    Unload main
Else
    main.Show vbModeless, OfficeStart
End If
End Sub


Private Sub Command12_Click()
On Error Resume Next
    Метка2.Text = Метка2.Text * (-1)
    Метка2_KeyPress 13
End Sub

Private Sub Command13_Click()
Option2_Click
End Sub

Private Sub Command3_Click()
On Error Resume Next
    Text5.Text = Text5.Text * (-1)
    Text5_KeyPress 13
End Sub


Sub Command5_Click()
    On Error Resume Next
    SetCenter Picture1, HScroll1, N_Slope
    RollesRefresh
    Draw_Systems Me.Picture1
End Sub


Private Sub Command7_Click()
On Error Resume Next
    Text4.Text = Text4.Text * (-1)
    Text4_KeyPress 13
End Sub


Private Sub Command8_Click()
    Text5.Text = 0
    Text5_KeyPress 13
End Sub


Private Sub Command9_Click()
    Text4.Text = 0
    Text4_KeyPress 13
End Sub

Sub Form_Load()
    On Error GoTo ERR
    SetFont Me
'    SelectLists.Clear
    Picture1.ScaleMode = 0
    
    OfficeStart.menuprofmanager.Enabled = False
    
    OfficeStart.Toolbar1.Buttons(7).Enabled = False
    OfficeStart.Toolbar1.Buttons(6).Enabled = False
    OfficeStart.menuRedo.Enabled = False
    OfficeStart.menuUndo.Enabled = False
    
    Positions = CurentPosition = 0
    
    If Setup.Combo4.ListIndex = 1 Then
        SuperRuler1.Measure = m
        SuperRuler2.Measure = m
    Else
        SuperRuler1.Measure = sm
        SuperRuler2.Measure = sm
    End If
    Current_P = 0
    
    sTabFx1.Style3D = Thin
    sTabFx1.AddTab lng.GetResIDstring(1016)
    sTabFx1.AddTab lng.GetResIDstring(1017)
    sTabFx1.AddTab lng.GetResIDstring(9189)
    sTabFx1.AddTab lng.GetResIDstring(1018)
    
    'sTabFx1.TabDisabled(1) = True
    'sTabFx1.TabDisabled(2) = True
    Command13.Caption = lng.GetResIDstring(1117)
    Picture1.BackColor = Setup.Command9.BackColor
    Picture1.MouseIcon = LoadResPicture(102, 2)
    Command2.Caption = lng.GetResIDstring(9196)
    Label14.Caption = lng.GetResIDstring(1010)
    Label3.Caption = lng.GetResIDstring(1500)
    Label9.Caption = lng.GetResIDstring(1013)
    Label6.Caption = lng.GetResIDstring(1015)
    '    Command4.Caption = lng.GetResIDstring(3003)
    Label2 = lng.GetResIDstring(1123)
    Text5.ToolTipText = lng.GetResIDstring(1124)
    Text4.ToolTipText = Text5.ToolTipText
    Check4.Caption = lng.GetResIDstring(9181)
    Label36.Caption = lng.GetResIDstring(1055)
    Check5.Caption = lng.GetResIDstring(9203)
    Check7.Caption = lng.GetResIDstring(9205)
    Check1.Caption = lng.GetResIDstring(9629)
    Check6.Caption = lng.GetResIDstring(9640)
'    Label16.Caption = lng.GetResIDstring(9687)
    Label18.Caption = lng.GetResIDstring(1056)
    Label19.Caption = lng.GetResIDstring(1057)
    Command1.Caption = lng.GetResIDstring(9177)
    Label15.Caption = lng.GetResIDstring(1079)
    Check10.Caption = lng.GetResIDstring(9684)
    Check13.Caption = lng.GetResIDstring(9689)
    calc_type(0).Caption = lng.GetResIDstring(9204)
    calc_type(1).Caption = lng.GetResIDstring(9206)
    Check3.Caption = lng.GetResIDstring(96700)
    Check9.Caption = lng.GetResIDstring(96701)
    Option4.Caption = lng.GetResIDstring(96702)
    Check8.Caption = lng.GetResIDstring(9685)
    chameleonButton1.Caption = lng.GetResIDstring(1504)
    chameleonButton2.Caption = lng.GetResIDstring(1505)
    isButton1.Caption = lng.GetResIDstring(9564)
    
    Label12.Caption = Label6.Caption & 0
    
    Dim FactoryName As String
    Dim ProfileName As String
    Dim FactoryID As Integer
    
    ProfileName = TrimNullChar(SlP(N_Slope).ProfilName)
    FactoryName = TrimNullChar(SlP(N_Slope).Factory_Name)
    
    Set SelectLists = New cCollection
        
    If Gl.FileNameExtension = ".rfd" Then
    
        '
        ' RFD
        '
        If SlP(N_Slope).CountOfPoints > 0 And SlP(N_Slope).Pn_Red_lines > 0 Then

        Dim ans As Integer

            ' Проверяем сменили ли мы профиль
            If (Project.Label3.Caption <> "" And Project.Label3.Caption <> ProfileName) Or (FactoryName <> Project.Label2.Caption) Then

                ans = MsgBox(lng.GetResIDstring(1490), vbInformation + vbYesNo)
                If ans = 6 Then

                    SlP(N_Slope).Pn_Red_lines = 0
                    SlP(N_Slope).Pn_StartLC = 0
                    SlP(N_Slope).CountSheets = 0
                    sTabFx1.SelectTab 0

                    Label5.Caption = Project.Label2.Caption
                    Label4.Caption = Project.Label3.Caption

                ElseIf ans = 7 Then
                
                    Label5.Caption = FactoryName
                    Label4.Caption = ProfileName
                    
                End If

            Else
            
                ' Открываю уже расчитанный
                Label5.Caption = Project.Label2.Caption
                Label4.Caption = Project.Label3.Caption
                
            End If

        Else
        
            ' Еще не расчитан
            Label5.Caption = Project.Label2.Caption
            Label4.Caption = Project.Label3.Caption
            
        End If
        
        chameleonButton3.Enabled = False
        
        Factory_Name = Label5.Caption
        
        ' Загрузка профиля
        SwitchProfile
        
    Else
    
        '
        ' RBP
        '
        Label5.Caption = FactoryName
        Label4.Caption = ProfileName
        Factory_Name = Label5.Caption
        
        chameleonButton3.Enabled = True
        
        ' Загрузка профиля
        SwitchProfile
        
    End If
    
    If Not SlP(N_Slope).CountOfLines = 0 Then
            Lapepic.Picture1.ScaleLeft = SlP(N_Slope).ScaleLeftS
            If SlP(N_Slope).ScaleWidthS <> 0 Then Lapepic.Picture1.ScaleWidth = SlP(N_Slope).ScaleWidthS
            Lapepic.Picture1.ScaleTop = SlP(N_Slope).ScaleTopS
            If SlP(N_Slope).ScaleHeightS <> 0 Then Lapepic.Picture1.ScaleHeight = SlP(N_Slope).ScaleHeightS
    Else
            Lapepic.Picture1.ScaleLeft = -100
            Lapepic.Picture1.ScaleWidth = 1100 '6400 'ConvertData(12, False)
            Lapepic.Picture1.ScaleTop = 600 'ConvertData(6, False)
            Lapepic.Picture1.ScaleHeight = -600 '-ConvertData(6, False)
    
            Lapepic.SuperRuler1.MaxH = Lapepic.HScroll1.MAX '* 10
            Lapepic.SuperRuler2.MaxV = Lapepic.HScroll1.MAX '* 10
    End If
    
    ' Привязка к линейки рисвание
    If Setup.Check4.value Then
        Check10.value = 1
    Else
        Check10.value = 0
    End If
    
    ' Привязка к линейки корректировка
    If Setup.Check9.value Then
        Check8.value = 1
    Else
        Check8.value = 0
    End If
    
    ' Рисование перекрестия
    If Setup.Check11.value Then
        Check11.value = 1
    Else
        Check11.value = 0
    End If
    
    Set SplitHV2.obj1 = Picture1
    Set SplitHV2.obj1 = cutfrm
    Set SplitHV2.obj1 = SplitHV1
    Set SplitHV2.obj2 = Frame1
    
    Set SplitHV1.obj1 = label7
    Set SplitHV1.obj1 = Picture1
    Set SplitHV1.obj2 = Text7
    Set SplitHV1.obj2 = cutfrm
    
    LoadSinCosTables
    
    HistoryClear True
    
    OfficeStart.HistoryWorking = False
    
    Exit Sub
ERR:
    '        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.2." & ERR.Source & "]", ERR.Number, ERR.Description
    Resume Next
End Sub


Private Sub Form_Resize()
On Error Resume Next

    If Me.Width > 5700 And Me.Height > 5100 Then

        SplitHV1.Height = Picture1.Height
        SplitHV2.Width = Me.ScaleWidth
        SplitHV2.ResizeControl
        SplitHV1.ResizeControl
        
        Frame1.Width = Me.ScaleWidth
        Label2.Width = Frame1.Width - 200
        sTabFx1.Width = Me.Width - 6700
        
    End If

End Sub


Private Sub Functions_Options(X As Single, Y As Single, Shift As Integer, Button As Integer)
    Dim Opt As String
    Dim pA As POINT
    Dim pB As POINT

    On Error GoTo ERR

    Select Case OptionDMM ' РЕЖИМ РИСОВАНИЯ
        
        Case "Mdraw"
        
            Text7.Text = "P(" & FindPointoint(X, Y) & ") X=" & ConvertData(X) & ", y=" & ConvertData(Y)
        
            If Current_P > 0 Then
            
'              Text7.text = Text7 & "; " & Current_P & " (" & Lape_Points_X(N_Slope, Current_P) & "," & Lape_Points_Y(N_Slope, Current_P) & ")"
              
              Lapepic.Text5 = X - Lape_Points_X(N_Slope, Current_P)
              Lapepic.Text4 = Y - Lape_Points_Y(N_Slope, Current_P)
              
              Lapepic.Text5 = ConvertData(Lapepic.Text5)
              Lapepic.Text4 = ConvertData(Lapepic.Text4)

              pA.X = Lape_Points_X(N_Slope, Current_P)
              pA.Y = Lape_Points_Y(N_Slope, Current_P)
              pB.X = X
              pB.Y = Y
                    
              Label8.Text = GetGRD(pA, pB)
              Lapepic.HScroll3.value = Label8.Text
              
'              Me.Метка2 = Format$(Sqr((X - Lape_Points_X(N_Slope, Current_P)) ^ 2 + (Y - Lape_Points_Y(N_Slope, Current_P)) ^ 2), "0.0")
            
            End If
            
        Case "Msel" ' РЕЖИМ КОРРЕКТИРОВКИ

            Text7.Text = "X=" & Format$(ConvertData(X), "0.00") & ", y=" & Format$(ConvertData(Y), "0.00")
            If Option4.value And P_A > 0 Then Text7.Text = Text7 & "; " & P_A & " (" & Lape_Points_X(N_Slope, P_A) & "," & Lape_Points_Y(N_Slope, P_A) & ")"
            Text7.Refresh

            If Check3.value Or Check9.value Then
            
                If P_A > 0 And P_B Then
                
                    Dim Len_AB_X As Single
                    Dim Len_AB_Y As Single
                
                    Len_AB_X = Lape_Points_X(N_Slope, P_B) - Lape_Points_X(N_Slope, P_A)
                    Len_AB_Y = Lape_Points_Y(N_Slope, P_B) - Lape_Points_Y(N_Slope, P_A)
                    
                    Lapepic.Text5 = ConvertData(Len_AB_X)
                    Lapepic.Text4 = ConvertData(Len_AB_Y)
                    
                    pA.X = Lape_Points_X(N_Slope, P_A)
                    pA.Y = Lape_Points_Y(N_Slope, P_A)
                    pB.X = Lape_Points_X(N_Slope, P_B)
                    pB.Y = Lape_Points_Y(N_Slope, P_B)
                    
                    Label8.Text = GetGRD(pA, pB)
                    Lapepic.HScroll3.value = Label8.Text
                    
  
                    ab = Sqr((Len_AB_X) ^ 2 + (Len_AB_Y) ^ 2)
  
                    Метка2.Text = ConvertData(ab)
  
                Else
                    
                    Lapepic.Text5 = ConvertData(X)
                    Lapepic.Text4 = ConvertData(Y)
                    
                End If
      
            ElseIf Option4.value Then

                If Button = 1 And P_A > 0 Then
                    If Option4.value Then
                        Lape_Points_Y(N_Slope, P_A) = Y
                        Lape_Points_X(N_Slope, P_A) = X
                        Draw_Systems Me.Picture1
                    End If
                End If

            End If

        Case "Msheet" ' РЕЖИМ НАЧЕРТАНИЯ ЛИСТОВ

            If Shift = 4 And Lapepic.Picture1.MousePointer = 7 Then ' Режим резки листов
                
            Label2 = lng.GetResIDstring(1401) '"Задайте высоту для разреза листа с нахлестом. Начинать резать с низу!"
  
  
            Dim cl As Integer
            cl = Find_list(X, Y)
  
            Dim YCut As Single
            YCut = Abs(Format$((List_Properties_PY(N_Slope, cl) - List_Properties_Length(N_Slope, cl)) - Y, "0.00"))  '& " см"'- SetProfilData(1).Tag, "0000.00")) '& " см"
            
            Text7.Text = ConvertData(YCut)
            If SetProfilData(1).Tag <> 0 Then
                Text7.Text = Text7.Text & " (" & Round(YCut / SetProfilData(1).Tag) & ") "
            End If
            Text7.Text = Text7.Text & " (x=" & ConvertData(X) & ", y=" & ConvertData(Y) & ")"
                        
                        
            Dim F As Boolean
            Dim i As Integer

            If Check5.value = True Then
            
                If CSng(YCut) > Val(SetProfilData(3).Tag) Then
    
                    For i = 0 To ArraySize(StandartLen)
                        If (StandartLen(i) + (SetProfilData(1).Tag - L1 - L2) < CSng(YCut) And CSng(YCut) < StandartLen(i) + (SetProfilData(1) - L1 - L2) + (L1 + L2)) Then
                            F = True: Label1 = StandartLen(i - 1): Exit For
                        Else
                            F = False
                        End If
    
                    Next
    
                Else
                    F = True
                End If
            
                If F Then
                    Line13.BorderColor = vbRed
                Else
                    Line13.BorderColor = vbGreen
                End If
            
            ElseIf Check7.value = 1 Then
            
                If CSng(YCut) > SetProfilData(3).Tag Then
    
                    For i = 0 To UBound(WrongLs)
                        If WrongLs(i).MIN < CSng(YCut) And CSng(YCut) < WrongLs(i).MAX Then
                            F = True: Label1 = WrongLs(i).MIN & ", " & WrongLs(i).MAX: Exit For
                        Else
                            F = False
                        End If
    
                    Next
    
                Else
                    F = True
                End If
          
                If F Then
                    Line13.BorderColor = vbRed
                Else
                    Line13.BorderColor = vbGreen
                End If
            
            Else
          
                If YCut < SetProfilData(3).Tag Then
                    Line13.BorderColor = vbRed
                Else
                    Line13.BorderColor = vbGreen
                End If
            
            End If
          
            Line13.Visible = True ' Начертание линии разреза
            Line13.X1 = Picture1.ScaleLeft
            Line13.Y1 = Y '+ SetProfilData(1).Tag
            Line13.x2 = Picture1.ScaleLeft + Picture1.ScaleWidth  ' Picture1.Width +
            Line13.y2 = Y '+ SetProfilData(1).Tag
    
            If Y_cut <> 0 Then  ' Последний разрез
                Line14.Visible = True
                Line14.X1 = Picture1.ScaleLeft
                Line14.Y1 = Y_cut + SetProfilData(1).Tag
                Line14.x2 = Picture1.ScaleLeft + Picture1.ScaleWidth 'Picture1.Width
                Line14.y2 = Y_cut + SetProfilData(1).Tag
            End If

    End If
    
'    If SelectLists.Count > 0 Then
'        Lapepic.Text5 = ConvertData(X)
'        Lapepic.Text4 = ConvertData(Y)
'        Text7.Text = " X=" & ConvertData(X) & ", y=" & ConvertData(Y)
'    Else
'        Lapepic.Text5 = ConvertData(CurrentPDataRS![WORK_WIDTH])
'        Lapepic.Text4 = ConvertData(List_Properties_Length(N_Slope, SelectLists.Item(0).List))
'        If SelectLists.count > 0 Then
'            Text7.Text = ""
'        Else
'            Text7.Text = List_Properties_Length(N_Slope, SelectLists.Item(0).list) '& " " & setup.Combo4.list(setup.Combo4.ListIndex)
'        End If
'    End If

End Select

Exit Sub
ERR:
'STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.3." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    OfficeStart.menuprofmanager.Enabled = True
    Unload Move_and_change
    Unload main
    
    HistoryClear False
    
    Set SelectLists = Nothing
    
    HistoryClear False
    
End Sub


Private Sub HScroll1_Change()
     HScroll1_Scroll
End Sub


Private Sub HScroll1_Scroll()
On Error Resume Next

'    If SuperRuler1.UserScale > 1000 And SlP(N_Slope).CountSheets > 1000 Then Exit Sub
    
    Module10.Change_scrol Picture1, HScroll1
    
    Text7.Text = HScroll1.MAX & Setup.Combo4.List(Setup.Combo4.ListIndex) & " / " & HScroll1.value '& Project.Label3.Caption
    
    RollesRefresh
    
'    If ResizePic = False Then
    If Lapepic.Visible Then Draw_Systems Me.Picture1
    
    Label11.ToolTipText = HScroll1.value

End Sub


Sub RollesRefresh()
On Error Resume Next

Dim Ratio As Long
Ratio = 1

Ratio = Round_to_big(HScroll1.value / (Screen.TwipsPerPixelX * 100)) * 100

SuperRuler1.UserScale = Ratio
SuperRuler2.UserScale = Ratio

SuperRuler1.Width = Picture1.Width
SuperRuler2.Height = Picture1.Height

SuperRuler1.ScaleLeft = Picture1.ScaleLeft
SuperRuler1.ScaleWidth = Picture1.ScaleWidth
SuperRuler1.Refresh

SuperRuler2.ScaleTop = Picture1.ScaleTop
SuperRuler2.ScaleHeight = Picture1.ScaleHeight
SuperRuler2.Refresh

End Sub

Function Round_to_big(Number)
On Error Resume Next

  Round_to_big = Number
  If Number > Int(Number) Then Round_to_big = Abs(Int(Number)) + 1
End Function

Private Sub HScroll2_Change()
On Error Resume Next
PolygonRotate HScroll2.value
Label12.Caption = Label6.Caption & HScroll2.value
SetChange True
End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub

Private Sub HScroll3_Change()
If Not Label8.Text = HScroll3.value Then

    SetChange True
    Label8.Text = HScroll3.value
    Label8_KeyPress 13

End If
End Sub

Private Sub isButton1_Click()
OfficeStart.menuslope_Click
End Sub

Private Sub IsSetProfilData_Click(Index As Integer)
If IsSetProfilData(Index).value = 1 Then
    SetProfilData(Index).Enabled = True
    SetProfilData(Index).BackColor = &H80000005
    
    SetProfilData(Index).Text = IsSetProfilData(Index).Tag
    
    If Index = 4 Then
        cange_max_size(0).Enabled = True
        cange_max_size(1).Enabled = True
    End If
    
Else
    SetProfilData(Index).Enabled = False
    SetProfilData(Index).BackColor = &HC8D0D4
    
    IsSetProfilData(Index).Tag = SetProfilData(Index).Text
    SetProfilData(Index).Text = 0
    
    If Index = 4 Then
        cange_max_size(0).Enabled = False
        cange_max_size(1).Enabled = False
    End If
    
End If
End Sub

Private Sub Label2_Click()
Load Teksti
Teksti.Text1 = Label2.Text
Teksti.Show vbModal, OfficeStart
Unload Teksti
End Sub

Private Sub label7_Click()
Load Teksti
Teksti.Text1 = label7
Teksti.Show vbModal, OfficeStart
Unload Teksti
End Sub

Private Sub Label8_KeyPress(KeyAscii As Integer)
    
    On Error Resume Next
    
    If KeyAscii <> 13 Then Exit Sub
    
    If P_A = 0 Or P_B = 0 Then Exit Sub

    Dim pPoint As POINT
    Dim pOrigin As POINT
    Dim pResult As POINT
    
    Dim Click2X As Single
    Dim Click2Y As Single
    
    Click2X = Lape_Points_X(N_Slope, P_A) + CSng(Sqr((Lape_Points_X(N_Slope, P_B) - Lape_Points_X(N_Slope, P_A)) ^ 2 + (Lape_Points_Y(N_Slope, P_B) - Lape_Points_Y(N_Slope, P_A)) ^ 2))
    Click2Y = Lape_Points_Y(N_Slope, P_A)

    pOrigin.X = Lape_Points_X(N_Slope, P_A)
    pOrigin.Y = Lape_Points_Y(N_Slope, P_A)

    pPoint.X = Click2X
    pPoint.Y = Click2Y

    pResult = RotatePoint(pPoint, pOrigin, Abs(Val(Label8.Text)))

    Lape_Points_X(N_Slope, P_B) = pResult.X
    Lape_Points_Y(N_Slope, P_B) = pResult.Y

    Lapepic.Text5 = ConvertData(Format$(Lape_Points_X(N_Slope, P_B) - Lape_Points_X(N_Slope, P_A), "0.0"))
    Lapepic.Text4 = ConvertData(Format$(Lape_Points_Y(N_Slope, P_B) - Lape_Points_Y(N_Slope, P_A), "0.0"))

    Метка2.Text = Format$(Sqr((Lapepic.Text5) ^ 2 + (Lapepic.Text4) ^ 2), "0.0")

    Draw_Systems Me.Picture1
    
End Sub


Private Sub Option4_Click()
On Error Resume Next

ValueFieldsOnOff False

If Option4.value Then
    P_A = 0
    P_B = 0
    Draw_Systems Me.Picture1
End If
End Sub


Function Find_list(X As Single, Y As Single) As Single
    Dim l017E As Single
    Dim N_list As Integer
    
    On Error Resume Next

    l017E = 0
    For N_list = 1 To SlP(N_Slope).CountSheets Step 1
            If X > List_Properties_PX(N_Slope, N_list) Then
                If X < List_Properties_PX(N_Slope, N_list) + CurrentPDataRS![WORK_WIDTH] Then
                    If Y < List_Properties_PY(N_Slope, N_list) Then
                        If Y > List_Properties_PY(N_Slope, N_list) - List_Properties_Length(N_Slope, N_list) Then
                            l017E = N_list
                            Exit For
                        End If

                    End If

                End If

            End If
        If l017E > 0 Then Exit For
    Next N_list

    Find_list = l017E
End Function


Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 112
    If FlagDraw = 0 Then sTabFx1.SelectTab 0 ' F1
    Exit Sub
Case 113
    If FlagDraw = 0 Then sTabFx1.SelectTab 1 ' F2
    Exit Sub
Case 114
    If FlagDraw = 0 Then sTabFx1.SelectTab 3 ' F3
    Exit Sub
Case 116 ' F5
    If FlagDraw = -1 Then
    If chameleonButton1.value = False Then
        chameleonButton1.value = True
    Else
        chameleonButton1.value = False
    End If
    End If
    Exit Sub
Case 117 ' F6
    If FlagDraw = -1 Then
    If chameleonButton2.value = False Then
        chameleonButton2.value = True
    Else
        chameleonButton2.value = False
    End If
    End If
    Exit Sub
Case 118 ' F7
    Exit Sub
End Select
Draw_Plate_Line_KeyDown KeyCode, Shift
End Sub


Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ERR
    
    ' Отмена рисования Esc
    If OptionDMM = "Mdraw" And KeyCode = 27 And FlagDraw = -1 Then
        Picture1_MouseDown 2, 0, 0, 0
    End If

    If OptionDMM = "Mdel" And (KeyCode = 46 Or KeyCode = vbKeyControl Or KeyCode = vbKeyShift) Then
        OptionDMM = "Mdraw"
        Picture1.MousePointer = 2
    End If
  
    If OptionDMM = "Msel" And KeyCode = 45 Then
        OptionDMM = "Msel"
        Picture1.MousePointer = 99
    End If

    If OptionDMM = "Msheet" And KeyCode = 46 Then
        Picture1.MousePointer = 99
    End If

    If OptionDMM = "Msheet" And KeyCode = vbKeyControl Then
        Picture1.MousePointer = 99
    End If
  
    If OptionDMM = "Msheet" And (KeyCode = 45 Or KeyCode = 18) Then
        Picture1.MousePointer = 99
        Line13.Visible = False
        Line14.Visible = False
        Text7.Text = ""
    End If

Exit Sub
ERR:
'    STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.7." & ERR.Source & "]", ERR.Number, ERR.Description
    Resume Next
End Sub



Private Sub Picture1_LostFocus()
    On Error Resume Next
    Line2.Visible = False
'    SlP(N_Slope).ScaleLeftS = Me.Picture1.ScaleLeft
'    SlP(N_Slope).ScaleWidthS = Me.Picture1.ScaleWidth
'    SlP(N_Slope).ScaleTopS = Me.Picture1.ScaleTop
'    SlP(N_Slope).ScaleHeightS = Me.Picture1.ScaleHeight
    Line3.Visible = False
    Line4.Visible = False
End Sub

Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
isClickMouse = True

If SlP(N_Slope).CountOfPoints = 0 And Setup.Check17.value = 1 Then
Draw_Plate_Line_MouseDown Button, Shift, 0, 0
Else
Draw_Plate_Line_MouseDown Button, Shift, X, Y
End If

isClickMouse = False
End Sub

Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If isActive.isFormFocus(Me.hwnd) Then
 isActive.SetFormFocus Picture1.hwnd
End If
Draw_Plate_Line_MouseMove Button, Shift, X, Y
End Sub


Sub WriteError(Optional modulename As String, Optional Error As String)
On Error Resume Next
    Beep
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR." & modulename & "]", ERR.Number, Error
    MsgBox lng.GetResIDstring(1472, "%PROFIL%", Label4.Caption, "%LIB%", modulename) & vbNewLine & Error, vbCritical   ' &  vbNewLine  & Lng.GetResIDstring(1445), vbCritical '&  vbNewLine  & str & IIf(ERROR <> "", ": [" + ERROR + "]", ""), vbCritical '&  vbNewLine  & Lng.GetResIDstring(1445), vbCritical
End Sub


'*********************************************************************************************
' POINT
'

Sub Draw_Point(ByVal Np As Integer, color As String, Optional size As Integer)
    Static tcolor As Long
    Dim l01CA As Single
    
    On Error Resume Next
    
    l01CA = Picture1.ScaleWidth / IIf(size > 0, size, 300)

    tcolor = Picture1.FillColor
    Picture1.FillColor = color
    Picture1.FillStyle = 0
    Picture1.Circle (Lape_Points_X(N_Slope, Np), Lape_Points_Y(N_Slope, Np)), l01CA, vbBlack
    
'    aspect = ScaleY(1, vbUser, vbPixels) / ScaleX(1, vbUser, vbPixels)
'    Circle (0, 0), 1, , , , asp

    Picture1.FillColor = tcolor
    Picture1.FillStyle = 1
End Sub


Private Function Add_Point(marray) As Boolean
On Error Resume Next

    If marray > MAXSLOPELINE Then
        MsgBox lng.GetResIDstring(1414, "%CURR%", marray, "%MAX%", Gl.MAXSLOPELINE), vbCritical
        Add_Point = True
        Exit Function
    Else
        marray = marray + 1
    End If
    Add_Point = False
End Function


Private Function Undo_Point(Optional marray) As Boolean
On Error Resume Next

    If marray > 0 Then
        Undo_Point = True
        marray = marray - 1
    End If
    Undo_Point = False
End Function


Sub Dell_Point(FindPoint As Integer)
    Dim i As Integer
    Dim l01D0 As Single
    Dim n As Single
        
    On Error Resume Next

        If SlP(N_Slope).CountOfPoints <= 1 Then
    
            If FindPoint <= 0 Then
                MsgBox lng.GetResIDstring(9198), vbInformation, ""
                OptionDMM = "Mdraw"
                Picture1.MousePointer = 2
                Exit Sub
            End If
    
        End If

        For i = FindPoint To SlP(N_Slope).CountOfPoints - 1 Step 1
            Lape_Points_X(N_Slope, i) = Lape_Points_X(N_Slope, i + 1)
            Lape_Points_Y(N_Slope, i) = Lape_Points_Y(N_Slope, i + 1)
        Next i

        Lape_Points_X(N_Slope, SlP(N_Slope).CountOfPoints) = 0: Lape_Points_Y(N_Slope, SlP(N_Slope).CountOfPoints) = 0
        SlP(N_Slope).CountOfPoints = SlP(N_Slope).CountOfPoints - 1
        Current_P = SlP(N_Slope).CountOfLines
        l01D0 = -1
        
'        Dim fclean As Boolean
'        fclean = False
        
        Do While l01D0 = -1
            l01D0 = 0
            For i = 1 To SlP(N_Slope).CountOfLines Step 1
                If Lape_Lines(N_Slope, i, 0) = FindPoint Or Lape_Lines(N_Slope, i, 1) = FindPoint Then
'                    If fclean = False Then
                    GoSub LD856
'                    fclean = True
                    Exit For
                Else
'                    If fclean Then
'                    Lape_Lines(N_Slope, i, 0) = 0
'                    Lape_Lines(N_Slope, i, 1) = 0
'                    End If
                End If

            Next i

        Loop

        For i = 1 To SlP(N_Slope).CountOfLines Step 1
            If Lape_Lines(N_Slope, i, 0) > FindPoint Then Lape_Lines(N_Slope, i, 0) = Lape_Lines(N_Slope, i, 0) - 1
            If Lape_Lines(N_Slope, i, 1) > FindPoint Then Lape_Lines(N_Slope, i, 1) = Lape_Lines(N_Slope, i, 1) - 1
        Next i

        Draw_Systems Me.Picture1
        Exit Sub

LD856:
        l01D0 = -1
        For n = i To SlP(N_Slope).CountOfLines Step 1
            Lape_Lines(N_Slope, n, 0) = Lape_Lines(N_Slope, n + 1, 0)
            Lape_Lines(N_Slope, n, 1) = Lape_Lines(N_Slope, n + 1, 1)
        Next n

        SlP(N_Slope).CountOfLines = SlP(N_Slope).CountOfLines - 1
        Current_P = SlP(N_Slope).CountOfLines
        Return

End Sub

'*****************************************************************************************

'*****************************************************************************************
' LINE
'
Function Draw_Plate_Line_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim P_N
    Dim l019C As Double
    Dim l019E As Double
    Dim i As Integer
    Dim CUTSN As Integer
    Dim find As Integer

    On Error GoTo ERR
    
'    If Setup.Check16.value Then
    X = Round(X, 2)
    Y = Round(Y, 2)
'    End If
    
    '
    ' Mdraw Mdel Msel
    '

    If OptionDMM = "Mdraw" Or (OptionDMM = "Mset" And isClickMouse = False) Then
        
        If Button = 1 Then
        
             If Check10.value Then
                ' Рисование с использованием привязки к линейкам
                X = SuperRuler1.CalculateValue(X, Y)
                Y = SuperRuler2.CalculateValue(X, Y)
             End If
         
             If FlagDraw = 0 Then
             
                sTabFx1.TabDisabled(1) = True
                sTabFx1.TabDisabled(2) = True
                chameleonButton1.Enabled = True
                chameleonButton2.Enabled = True
                 
                 If SlP(N_Slope).CountOfPoints > 0 Then
                 
                    'Current_P = 0
                 
                    ' Поиск точки
                    If Picture1.MousePointer = 2 Then Current_P = FindPointoint(X, Y, TIE_DISTANCE)
                    
                    If Current_P = 0 Then
                        ' Поиск линии
                        Dim cl As Integer
                        cl = Find_Line(X, Y)
                        If cl > 0 And cl = Current_L Then
                            Current_L = 0
                            Draw_Systems Me.Picture1
                            Exit Function
                        Else
                            Current_L = cl
                        End If
                    Else
                        Current_L = 0
                    End If
                    
                 Else
                    Current_L = 0
                    Current_P = 0
                 End If
                 
                 If Current_P > 0 Then
                    
                    GoTo TIEA ' переходим к привязке если есть точка где произв лев клик мыши
                    
                 ElseIf Current_L > 0 Then
                
                    '
                    ' Деление линии и начало новой
                    '
                    Current_P = Divide_Line(X, Y, Current_L)
                    If Current_P > 0 Then
                        Draw_Systems Me.Picture1
                        Draw_Point Current_P, vbGreen, 270  ' прорисовка точки
                        GoTo TIEA
                    End If
                    Exit Function
                    
                 ElseIf Picture1.MousePointer = 2 Then
                 
                    ' Устанавливаем флаг изменения данных
'                    SetChange true
                 
                    ' Добавление точки 1
                    If Add_Point(SlP(N_Slope).CountOfPoints) Then Exit Function ' Points
                    
                    ' Добавление линии
                    If Add_Point(SlP(N_Slope).CountOfLines) Then Exit Function ' Lines
        
                    ' забивка точек X,Y
                    Lape_Points_X(N_Slope, SlP(N_Slope).CountOfPoints) = X
                    Lape_Points_Y(N_Slope, SlP(N_Slope).CountOfPoints) = Y
                    
                    Current_P = SlP(N_Slope).CountOfPoints
                    
                    ' выполнение привязки 1 точки
                    Lape_Lines(N_Slope, SlP(N_Slope).CountOfLines, 0) = SlP(N_Slope).CountOfPoints
        
                    Draw_Point Current_P, vbGreen, 270  ' прорисовка точки
                    FlagDraw = -1
                    
                    Draw_Line Current_P, X, Y, True ' прорисовка линии (для того чтоб стереть предыдущие координаты)
                    Line2.Visible = True
                    
                    Exit Function
                     
                 End If
    
             ElseIf Picture1.MousePointer = 2 Then ' Рисование уже начато (черн линия)
             
                chameleonButton1.Enabled = False
                chameleonButton2.Enabled = False
                sTabFx1.TabDisabled(1) = False
                sTabFx1.TabDisabled(2) = False
                
                '
                ' Рисование линий с прязкой к углам (рисование прямых)
                '
                If OfficeStart.menu_view_m(0).Checked Or (OptionDMM = "Mdraw" And Shift = 1) Then
     
                    ' первая строка 0 - 90 и 270 - 360
                    If (0.75 <= SinA And SinA <= 1) And (1 >= SinA And SinA >= 0.8) Or _
                    (-0.8 >= SinA And SinA >= -1) And (-0.8 >= SinA And SinA >= -1) Then
                        Y = Lape_Points_Y(N_Slope, Current_P)
                    ' вторая строка 90 - 180 и 180 - 270
                    ElseIf (-0.93 <= SinA And SinA >= 0) And (0 <= SinA And SinA <= 0.8) Or _
                    (0 >= SinA And SinA >= -0.8) And (-0.8 <= SinA And SinA <= 0) Then
                        X = Lape_Points_X(N_Slope, Current_P)
                    End If
      
                End If
     
                ' Поиск точки
                If Picture1.MousePointer = 2 Then FindPoint = FindPointoint(X, Y, TIE_DISTANCE)
                
                If FindPoint = 0 Then
                    ' Поиск линии
                    Current_L = Find_Line(X, Y)
                Else
                    Current_L = 0
                End If
                 
                 If FindPoint > 0 Then
                     
                     GoTo TIEB ' Привязка к точке B (уже существует)
                     
                ElseIf Current_L > 0 Then
                    
                    ' Устанавливаем флаг изменения данных
                    SetChange True
                    
                    ' Привязка к точке, разделяющей линюю
                    Dim LastLine As Integer
                    LastLine = SlP(N_Slope).CountOfLines
                    Current_P = Divide_Line(X, Y, Current_L)
                    If Current_P > 0 Then
                        Lape_Lines(N_Slope, LastLine, 1) = Current_P
                        FlagDraw = 0
                        Line2.Visible = False
                        Draw_Systems Me.Picture1
                    End If
                    
                    Exit Function
                    
                 Else
                 
                     ' Добавление точки 2
                     If Add_Point(SlP(N_Slope).CountOfPoints) Then Exit Function 'B
                     
                     ' Выполнение привязки 2 точки к точке 1
                     Lape_Lines(N_Slope, SlP(N_Slope).CountOfLines, 1) = SlP(N_Slope).CountOfPoints
                     
                     ' Забивка точек X,Y
                     Lape_Points_X(N_Slope, SlP(N_Slope).CountOfPoints) = X
                     Lape_Points_Y(N_Slope, SlP(N_Slope).CountOfPoints) = Y

                     Current_P = SlP(N_Slope).CountOfPoints
                     
                    ' Устанавливаем флаг изменения данных
                    SetChange True
                     
                     If Check13.value Then
                        ' Рисовать неотрывно
                        Draw_Systems Me.Picture1
                        GoTo TIEA
                     Else
                        ' Обычный метод рисования
                        FlagDraw = 0
                        Line2.Visible = False
                        Draw_Systems Me.Picture1
                     End If
                     
                     Exit Function
                     
                 End If
                 
             End If
    
        ElseIf Button = 2 And Picture1.MousePointer = 2 Then
            
            If FlagDraw = 0 Then
            
                sTabFx1.TabDisabled(1) = True
                sTabFx1.TabDisabled(2) = True
                chameleonButton1.Enabled = True
                chameleonButton2.Enabled = True
            
                If SlP(N_Slope).CountOfPoints = 0 Then Exit Function
            
                Current_P = FindPointoint(X, Y)
TIEA:
                ' Устанавливаем флаг изменения данных
'                SetChange true
                
                If Add_Point(SlP(N_Slope).CountOfLines) Then Exit Function 'A

                Lape_Lines(N_Slope, SlP(N_Slope).CountOfLines, 0) = Current_P
                
                FlagDraw = -1
                Line2.Visible = True
                Draw_Line Current_P, X, Y, True
                
                Exit Function
                
            Else
            
                chameleonButton1.Enabled = False
                chameleonButton2.Enabled = False
                sTabFx1.TabDisabled(1) = False
                sTabFx1.TabDisabled(2) = False
                
                If X = 0 And Y = 0 Then
                    FindPoint = Current_P
                Else
                    FindPoint = FindPointoint(X, Y) ' Поиск точки около которой был произведен клик
                End If
TIEB:
                
                ' Устанавливаем флаг изменения данных
                SetChange True
                
                If Current_P <> FindPoint Then
                    If FindBindingPoints(Current_P, FindPoint) Then
                        If Undo_Point(SlP(N_Slope).CountOfLines) Then Exit Function ' Такая линия существует (удаляем колво линий)
                    Else
                        Lape_Lines(N_Slope, SlP(N_Slope).CountOfLines, 1) = FindPoint ' Выполняем привязку к найденой точки
                    End If
                Else
                    If Undo_Point(SlP(N_Slope).CountOfLines) Then Exit Function 'A удаление привязки (привязка прошла неудачна, не найдено точек привязки)
                End If
                
                FlagDraw = 0
                Line2.Visible = False
                Draw_Systems Me.Picture1
                
                Exit Function
    
            End If

        End If
    
    ElseIf OptionDMM = "Mdel" And Picture1.MousePointer = 10 Then
    
      ' Устанавливаем флаг изменения данных
      SetChange True
    
      If Shift = 2 Then
      
        ' Удаление точки
        Dell_Point FindPointoint(X, Y)
        
      Else
        
        ' Удаление линии
        If Del_Line(Find_Line(X, Y)) Then Draw_Systems Me.Picture1
        
      End If
      
      Current_P = 0
    
    ElseIf OptionDMM = "Msel" Then
    
'    If Setup.Check16.value Then
'        X = Round(X, 2)
'        Y = Round(Y, 2)
'    End If
    
        ' Выделение точек A & B
        If Check3.value Then
            
            If Button = 1 Then
            
                P_A = FindPointoint(X, Y) ' A
            
            Else
        
                If P_A > 0 Then P_B = FindPointoint(X, Y) ' B
            
            End If
            
        ' Выделение линии и установки точек A & B
        ElseIf Check9.value Then
        
            Dim nn As Integer
            nn = Find_Line(X, Y)
            If nn > 0 Then
                P_A = Lape_Lines(N_Slope, nn, 0)
                P_B = Lape_Lines(N_Slope, nn, 1)
            End If
            
        ' Выделение точки
        ElseIf Option4.value Then
        
            P_A = FindPointoint(X, Y)
                
        End If
    

        If (P_A > 0 And P_B > 0) And (P_A <> P_B) And Option4.value = False Then
            ValueFieldsOnOff True
        Else
            ValueFieldsOnOff False
        End If
        
        Draw_Systems Me.Picture1
        
        Exit Function
    
    '
    ' Msheet
    '
    ElseIf OptionDMM = "Msheet" Then
    
      If SlP(N_Slope).Pn_Red_lines = 0 Then ' Проверка выполнялся расчет или нет
          
          SlP(N_Slope).Pn_Red_lines = FindPointoint(X, Y)
          Label2 = lng.GetResIDstring(1407) '"Выберите точку, откуда будет браться старт измерения"
          Draw_Systems Me.Picture1
          Exit Function
          
      Else
      
          If SlP(N_Slope).Pn_StartLC = 0 Then ' Проверка есть ли точка начала старта отсчета
              
              SlP(N_Slope).Pn_StartLC = FindPointoint(X, Y)
              SlP(N_Slope).PX_StartLC = Int(Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC))
              SlP(N_Slope).CountSheets = 0
              Label2 = lng.GetResIDstring(1408) '"Выберите стартовую линию"
              ' Устанавливаем флаг изменения данных
              SetChange True
              Draw_Systems Me.Picture1
              Exit Function
              
          Else
          
              If SlP(N_Slope).CountSheets = 0 Then ' Проверка производился ли расчет
                  
                  SlP(N_Slope).PX_StartLC = Int(X)
                  calc
                  ' Устанавливаем флаг изменения данных
                  SetChange True
                  Draw_Systems Me.Picture1
              
              Else ' РЕЖИМ РЕДАКТИРОВАНИЯ ЛИСТОВ, ПОЛОС *** ЕСЛИ РАСКРОЙ ИМЕЕТСЯ ***
                  
                  Select Case Picture1.MousePointer ' Выбор действия в зависимости от вида мыши
                      
                      Case 0, 1, 99
                          
                          On Error Resume Next
                          If Button = 1 Then
                            SelectList X, Y, Shift, Button
                            If SelectLists.Count > 0 Then
                              ValueFieldY_OnOff True
                            Else
                              ValueFieldY_OnOff False
                            End If
                          Else
'                            If SelectLists.Count > 0 Then
'                                Lapepic.ListEdit True
'                            Else
'                                Lapepic.ListEdit False
'                            End If
                            If Move_and_change.Visible = True Then
                                Lapepic.ListEdit False
                            Else
                                Lapepic.ListEdit True
                            End If
                          End If
                          On Error GoTo ERR
                          
                      Case 7
                      
                          '---------------------------------------------------------------
                          ' РАЗРЕЗ ЛИСТОВ
                          '---------------------------------------------------------------
                          
                          isSave = True
                
DIV:
                        
                          If Shift = 6 Then
                              Y = Line13.Y1
                          End If
                          
                          Dim cul As cList
                          For Each cul In SelectLists.Items
                  
                              If Shift = 6 Then ' разрез вдоль линии
                  
                                  ' Выполнение подбора высоты разреза при наличии невыполнимых длин (WL - кол-во длин)
                                  If Line13.BorderColor = vbRed And Check5.value = True Then _
                                  Y = List_Properties_PY(N_Slope, cul.List) - _
                                  List_Properties_Length(N_Slope, cul.List) + mCint(Label1) '+ SetProfilData(1).Tag
                      
                                  ' Ограничение по высоте
                                  If List_Properties_PY(N_Slope, cul.List) > Y And _
                                  List_Properties_PY(N_Slope, cul.List) - _
                                  List_Properties_Length(N_Slope, cul.List) < Y Then
                                      
                                        CutList Button, Y, cul.List ' Разрез выбранных листов
                                      
                                  End If
    
                    
                              ElseIf Check7.value = 1 Then ' невыполнимые длины
                  
                                  If Line13.BorderColor = vbRed Then _
                                  Y = List_Properties_PY(N_Slope, SelectLists.Item(cul.List).List) - _
                                  List_Properties_Length(N_Slope, SelectLists.Item(cul.List).List) + Label1
                                   
                                  If List_Properties_PY(N_Slope, SelectLists.Item(cul.List).List) > Y And _
                                  List_Properties_PY(N_Slope, SelectLists.Item(cul.List).List) - _
                                  List_Properties_Length(N_Slope, SelectLists.Item(cul.List).List) < Y Then
                                      
                                        CutList Button, Y, SelectLists.Item(cul.List).List ' Разрез выбранных листов
                                      
                                  End If
                    
                              Else ' Разрез выбранных листов листов
                  
                                  ' Выполнение подбора высоты разреза при наличии невыполнимых длин (WL - кол-во длин)
                                  If Line13.BorderColor = vbRed And Check5.value = True Then
                                      Y = List_Properties_PY(N_Slope, SelectLists.Item(cul.List).List) - _
                                      List_Properties_Length(N_Slope, SelectLists.Item(cul.List).List) + Label1 '+ SetProfilData(1).Tag
                                  ElseIf Line13.BorderColor = vbRed Then
                                      Y = List_Properties_PY(N_Slope, SelectLists.Item(cul.List).List) - _
                                      List_Properties_Length(N_Slope, SelectLists.Item(cul.List).List) + SetProfilData(3).Tag
                                  End If
      
                                  ' Ограничение по высоте
                                  If List_Properties_PY(N_Slope, SelectLists.Item(cul.List).List) > Y And _
                                  List_Properties_PY(N_Slope, SelectLists.Item(cul.List).List) - _
                                  List_Properties_Length(N_Slope, SelectLists.Item(cul.List).List) < Y Then
                      
                                        CutList Button, Y, SelectLists.Item(cul.List).List ' Разрез выбранных листов
                      
                                  End If
                     
                              End If
                  
                          Next
                          
                          ' Устанавливаем флаг изменения данных
                          SetChange True
                
                          Me.Picture1.MousePointer = 99
                
                  End Select
                  Draw_Systems Me.Picture1 ' Прорисовка
                  Exit Function
              End If
    
          End If
    
      End If
    
    End If
    
    Exit Function
ERR:
    'STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.8." & ERR.Source & "]", ERR.Number, ERR.Description
    Resume Next
End Function

Function Draw_Plate_Line_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ERR
    
   '
   ' Не занята клавиша Shift (когда не начато рисование флаг FlagDraw = 0)
   '
    
'    If Setup.Check16.value Then
        X = Round(X, 2)
        Y = Round(Y, 2)
'    End If

'    SuperRuler1.RenderTrackLine X, 0
'    SuperRuler2.RenderTrackLine 0, Y
    
    If (OptionDMM = "Mdraw" And Check10.value) Or _
    (OptionDMM = "Msel" And Check8.value) Then
    
        ' Рисование с использованием привязки к линейкам
        X = SuperRuler1.CalculateValue(X, Y)
        Y = SuperRuler2.CalculateValue(X, Y)

    End If
    
    SuperRuler1.RenderTrackLine X, 0
    SuperRuler2.RenderTrackLine 0, Y
    
    If Current_P > 0 And OptionDMM = "Mdraw" And FlagDraw = -1 Then
    
        ' Рисование прямых
        If Shift = 1 Or OfficeStart.menu_view_m(0).Checked Then
            
            SinA = (X - Lape_Points_X(N_Slope, Current_P)) / Sqr((X - Lape_Points_X(N_Slope, Current_P)) ^ 2 + _
            (Y - Lape_Points_Y(N_Slope, Current_P)) ^ 2)
    
            ' первая строка 0 - 90 и 270 - 360
            If (0.75 <= SinA And SinA <= 1) And (1 >= SinA And SinA >= 0.8) Or _
                (-0.8 >= SinA And SinA >= -1) And (-0.8 >= SinA And SinA >= -1) Then
                    Y = Lape_Points_Y(N_Slope, Current_P)
            ' вторая строка 90 - 180 и 180 - 270
            ElseIf (-0.93 <= SinA And SinA >= 0) And (0 <= SinA And SinA <= 0.8) Or _
                (0 >= SinA And SinA >= -0.8) And (-0.8 <= SinA And SinA <= 0) Then
                    X = Lape_Points_X(N_Slope, Current_P)
            End If
            
        End If
    
        ' Рисование линии
        If Picture1.MousePointer = 2 Then
            Call Draw_Line(Current_P, X, Y) ' рисование
        End If
    
    End If
    
    If (OptionDMM = "Mdraw" Or OptionDMM = "Msel") And Check11.value Then
        Line3.X1 = X
        Line3.x2 = X
        Line3.Y1 = -Lapepic.HScroll1.MAX '0 'Picture1.ScaleTop
        Line3.y2 = Lapepic.HScroll1.MAX 'Picture1.ScaleHeight * 3
        Line3.Visible = True
        Line3.Refresh

        Line4.X1 = -Lapepic.HScroll1.MAX 'Picture1.ScaleLeft
        Line4.x2 = Lapepic.HScroll1.MAX 'Picture1.ScaleWidth * 3
        Line4.Y1 = Y
        Line4.y2 = Y
        Line4.Visible = True
        Line4.Refresh
    Else
        Line3.Visible = False
        Line4.Visible = False
    End If

    Call Functions_Options(X, Y, Shift, Button) ' обработка действий
    
    Exit Function
ERR:
'    STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
'    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.9." & ERR.Source & "]", ERR.Number, ERR.Description
    Resume Next
End Function


Function Draw_Plate_Line_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Step As Single
    Dim l0188 As Single
    Dim l018A
    Dim cul As cList

    On Error GoTo ERR

    If OptionDMM = "Mdraw" Or OptionDMM = "Mdel" Or OptionDMM = "Mset" Then
        
        ' Ввод с ипользованием кл Enter
        If KeyCode = 13 Then
            
            isClickMouse = False
            ValueFieldsOnOff True
            
            If Picture1.MousePointer = 2 Then
                
                If OptionDMM = "Mdraw" Then
                    
                    ' Начало рисование
                    OptionDMM = "Mset"
                    
                    HScroll3.Enabled = False
                    Метка2.Enabled = False
                    Command12.Enabled = False
                    
                    If FlagDraw = 0 Then
                    If SlP(N_Slope).CountOfPoints = 0 Then
                        Draw_Plate_Line_MouseDown 1, 0, SuperRuler1.GetCurrentPos, SuperRuler2.GetCurrentPos
                    Else
                        Draw_Plate_Line_MouseDown 1, 0, Lape_Points_X(N_Slope, SlP(N_Slope).CountOfPoints), Lape_Points_Y(N_Slope, SlP(N_Slope).CountOfPoints)
                    End If
                    End If
                    Me.label7.ForeColor = vbRed
                    Me.label7.Text = lng.GetResIDstring(9678)
                    
                Else
                    
                    ' Конец рисования
                    OptionDMM = "Mdraw"
                    Me.label7.Text = ""
                    Me.label7.ForeColor = vbBlack
                    ' или чертит иль отменяет
                    Draw_Plate_Line_MouseDown 1, 0, Line2.x2, Line2.y2
                    Line2.Visible = False
                    
                    ' Проверка на прямой ввод если ввод был был с окна расчета
                    ' Включить продолжение
                    If Not IsLoadForm("SetPoint") Then
                        Command5.value = True
                        Draw_Plate_Line_KeyDown 13, 0
                    End If
                    
                End If
                
            End If
            
        ElseIf KeyCode = 27 Then
        
            If OptionDMM = "Mset" Then
                    ValueFieldsOnOff False
                    ' Конец рисования
                    OptionDMM = "Mdraw"
                    Me.label7.Text = ""
                    Me.label7.ForeColor = vbBlack
                    Line2.Visible = False
                    Command5.value = True
            ElseIf OptionDMM = "Mdraw" Then
                Lapepic.Picture1_MouseDown 2, 0, 0, 0
            End If
            
        ElseIf ((KeyCode = vbKeyDelete Or KeyCode = vbKeyControl) Or (KeyCode = 16 And Shift = 3)) And FlagDraw = 0 Then
        
            Me.Picture1.MousePointer = 10
            OptionDMM = "Mdel"
            
        End If
    
    ElseIf OptionDMM = "Msel" Then
    
        If KeyCode = 46 Then
        If Del_Line(Find_Line_by_AB(P_A, P_B)) Then
            ' Устанавливаем флаг изменения данных
            SetChange True
            Draw_Systems Me.Picture1
        End If
        Exit Function
        End If

        If Check3.value Or Check9.value Then
            
            Step = 1
            
            '  Изменение шага
            If Shift = 1 Then
                Step = 10
            Else
                If Shift = 2 Then
                    Step = 100
                Else
                    If Check10.value Then
                        If Check8.value Then
                            Step = SuperRuler1.UserScale / 10
                        End If
                    End If
                End If
            End If
  
            Me.txt_step.Text = Step
            Select Case KeyCode
                Case 38
                    Lape_Points_Y(N_Slope, P_B) = Lape_Points_Y(N_Slope, P_B) + Step
                Case 40
                    Lape_Points_Y(N_Slope, P_B) = Lape_Points_Y(N_Slope, P_B) - Step
                Case 37
                    Lape_Points_X(N_Slope, P_B) = Lape_Points_X(N_Slope, P_B) - Step
                Case 39
                    Lape_Points_X(N_Slope, P_B) = Lape_Points_X(N_Slope, P_B) + Step
            End Select
  
            If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 37 Or KeyCode = 39 Then
                ' Устанавливаем флаг изменения данных
                SetChange True
                Draw_Systems Me.Picture1
                Call Functions_Options(0, 0, 0, 0)
            End If
  
        End If
        
    ElseIf OptionDMM = "Msheet" Then

            ' Удаление листов
            If KeyCode = 46 Then
                
                For Each cul In SelectLists.Items
                    List_Properties_Length(N_Slope, cul.List) = 0
                Next
        
                SelectLists.Clear
                
                ' Устанавливаем флаг изменения данных
                SetChange True
                
                Draw_Systems Me.Picture1
                Exit Function
            End If
        
            If KeyCode = 45 Or KeyCode = 18 Then Picture1.MousePointer = 7  ' Разрез ската
    
           Step = 1
    
           ' Изменение шага
           If Shift = 1 Then
               Step = 10
           ElseIf Shift = 3 Then
               Step = SetProfilData(1).Tag
           Else
               If Shift = 2 Then
                   Step = 100
               ElseIf Shift = 0 Then
                    Step = 1
               End If
           End If
           Me.txt_step.Text = Step
       
           ' При смещении вправо или лево ПЕРЕРАСКРОЙ
           If KeyCode = 37 Or KeyCode = 39 Then
           
'               SelectLists.Clear
               Select Case KeyCode
                   Case 37 ' <--
                       SlP(N_Slope).PX_StartLC = SlP(N_Slope).PX_StartLC - Step
                       isPushCalc = False
                   Case 39 ' -->
                       SlP(N_Slope).PX_StartLC = SlP(N_Slope).PX_StartLC + Step
                       isPushCalc = False
               End Select
             
               calc
               ' Устанавливаем флаг изменения данных
               SetChange True
               
               Draw_Systems Me.Picture1
               Exit Function
             
          End If
    
          ' Операции с листами
          If SelectLists.Count > 0 And Picture1.MousePointer = 99 Then
              
          ' Подготовка к груповой обработке
          Dim ars As Integer
          Dim n As Integer
          Dim dosize As Boolean
          Dim L As Integer
              
          On Error Resume Next
          Dim L3 As Integer
          L3 = SetProfilData(1).Tag - L1 - L2
            
'          ars = SelectLists.Count
            
          Select Case KeyCode
              Case 40
                  
                  For Each cul In SelectLists.Items ' Уменьшение листов
                        
                      ' Расчет минимально допустимой длины
                      If List_Properties_Length(N_Slope, cul.List) - Step < SetProfilData(3).Tag Then
                          
                          L = List_Properties_PY(N_Slope, cul.List) - List_Properties_Length(N_Slope, cul.List)
                          List_Properties_Length(N_Slope, cul.List) = SetProfilData(3).Tag
                          List_Properties_PY(N_Slope, cul.List) = SetProfilData(3).Tag + L
                          
                          dosize = False
                      Else
                          dosize = True
                      End If
                  
                      If dosize Then
                
                          L = List_Properties_PY(N_Slope, cul.List) - List_Properties_Length(N_Slope, cul.List)
                          If txt_CL.ListCount - 1 > 1 And (Check5.value Or Check1.value) Then
                
                              For n = 1 To txt_CL.ListCount - 1
                                  If List_Properties_Length(N_Slope, cul.List) >= txt_CL.List(n) And _
                                  List_Properties_Length(N_Slope, cul.List) <= txt_CL.List(n + 1) Then
                                      List_Properties_Length(N_Slope, cul.List) = txt_CL.List(n)
                                      List_Properties_PY(N_Slope, cul.List) = txt_CL.List(n) + L
                                      Exit For
                                  End If
      
                              Next
                  
                          Else
                
                              '
                              ' С учетом шага
                              '
                              If Move_and_change.Option1.value Then
                                    
                                  If Shift = 3 Then
                                    Step = List_Properties_Length(N_Slope, cul.List) \ CSng(SetProfilData(1).Tag)
                                    Step = (Step * SetProfilData(1).Tag) - SetProfilData(1).Tag
                                    Step = List_Properties_Length(N_Slope, cul.List) - Step
                                  End If
                              
                                  List_Properties_Length(N_Slope, cul.List) = List_Properties_Length(N_Slope, cul.List) - Step
                                  List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) - Step
                                  
                              Else
                                  List_Properties_Length(N_Slope, cul.List) = List_Properties_Length(N_Slope, cul.List) + Step
                                  
                              End If
                  
                          End If
                
                      End If
      
                  Next
                
              Case 38 ' увеличение листов
                  
                    
                For Each cul In SelectLists.Items
      
                      L = List_Properties_PY(N_Slope, cul.List) - List_Properties_Length(N_Slope, cul.List)
                      If txt_CL.ListCount - 1 > 1 And (Check5.value Or Check1.value) Then
                
                          For n = 1 To txt_CL.ListCount - 1
                              If List_Properties_Length(N_Slope, cul.List) < txt_CL.List(n) And _
                              List_Properties_Length(N_Slope, cul.List) >= txt_CL.List(n - 1) Then
                                  List_Properties_Length(N_Slope, cul.List) = txt_CL.List(n)
                                  List_Properties_PY(N_Slope, cul.List) = txt_CL.List(n) + L
                                  Exit For
                              End If
      
                          Next
                  
                      Else
                
                          '
                          ' С учетом шага
                          '
                          If Move_and_change.Option1.value Then
                          
                          
                              If Shift = 3 Then
                                Step = List_Properties_Length(N_Slope, cul.List) \ CSng(SetProfilData(1).Tag)
                                Step = (Step * SetProfilData(1).Tag) + SetProfilData(1).Tag
                                Step = Step - List_Properties_Length(N_Slope, cul.List)
                              End If
                          
                              List_Properties_Length(N_Slope, cul.List) = List_Properties_Length(N_Slope, cul.List) + Step
                              List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) + Step
                              
                          Else
                          
                              List_Properties_Length(N_Slope, cul.List) = List_Properties_Length(N_Slope, cul.List) - Step
                              
                          End If
      
                      End If
      
                  Next
                
              Case 33
                  
                  List_Properties_Length(N_Slope, SelectLists.Item(0).List) = List_Properties_Length(N_Slope, SelectLists.Item(0).List) + Me.txt_step.Text
                  List_Properties_PY(N_Slope, SelectLists.Item(0).List) = List_Properties_PY(N_Slope, SelectLists.Item(0).List) + Me.txt_step.Text
                
              Case 34
                  
                  If Step < List_Properties_Length(N_Slope, SelectLists.Item(0).List) Then
                      List_Properties_Length(N_Slope, SelectLists.Item(0).List) = List_Properties_Length(N_Slope, SelectLists.Item(0).List) - Me.txt_step.Text
                      List_Properties_PY(N_Slope, SelectLists.Item(0).List) = List_Properties_PY(N_Slope, SelectLists.Item(0).List) - Me.txt_step.Text
                  End If
                
          End Select
      
          If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then
              Draw_Systems Me.Picture1
              
              ' Устанавливаем флаг изменения данных
              SetChange True
              
              Call Functions_Options(0, 0, 0, 0)
          End If
      
   End If

End If
  
Exit Function
ERR:
'STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.6." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next
End Function


Function Find_Line(X As Single, Y As Single) As Integer
    Dim i As Integer
    Dim AC As Single
    Dim BC As Single
    
    On Error GoTo ERR

'    Dim maxpoints As Integer
'    maxpoints = SlP(N_Slope).CountOfLines
'    If SlP(N_Slope).CountOfPoints > maxpoints Then maxpoints = SlP(N_Slope).CountOfPoints

    For i = 1 To SlP(N_Slope).CountOfLines
    
        If Lape_Lines(N_Slope, i, 0) > 0 Or Lape_Lines(N_Slope, i, 1) > 0 Then

        AC = Sqr((Lape_Points_X(N_Slope, Lape_Lines(N_Slope, i, 0)) - X) ^ 2 + (Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, i, 0)) - Y) ^ 2)
        BC = Sqr((Lape_Points_X(N_Slope, Lape_Lines(N_Slope, i, 1)) - X) ^ 2 + (Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, i, 1)) - Y) ^ 2)

        If Round(AC + BC, 0) = Round(Sqr((Lape_Points_X(N_Slope, Lape_Lines(N_Slope, i, 1)) - Lape_Points_X(N_Slope, Lape_Lines(N_Slope, i, 0))) ^ 2 + _
        (Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, i, 1)) - Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, i, 0))) ^ 2), 0) Then
            Find_Line = i
            Exit For
        End If
        
        End If

    Next i

    Exit Function
ERR:
    Find_Line = 0
End Function


Function Find_Line_by_AB(a As Integer, b As Integer) As Integer
On Error Resume Next

If a = 0 Or b = 0 Then Exit Function
Dim i As Integer
For i = 1 To SlP(N_Slope).CountOfLines Step 1
    If (Lape_Lines(N_Slope, i, 0) = a And Lape_Lines(N_Slope, i, 1) = b) Or (Lape_Lines(N_Slope, i, 1) = a And Lape_Lines(N_Slope, i, 0) = b) Then
        Find_Line_by_AB = i
        Exit For
    End If
Next
End Function


Function Del_Line(n As Integer) As Boolean
On Error Resume Next

If n > 0 Then
Dim i As Integer
For i = n To SlP(N_Slope).CountOfLines Step 1
    Lape_Lines(N_Slope, i, 0) = Lape_Lines(N_Slope, i + 1, 0)
    Lape_Lines(N_Slope, i, 1) = Lape_Lines(N_Slope, i + 1, 1)
Next
SlP(N_Slope).CountOfLines = SlP(N_Slope).CountOfLines - 1
Del_Line = True
End If
End Function

Function Divide_Line(X As Single, Y As Single, Current_L As Integer) As Integer
On Error Resume Next

    ' Добавление точки 1
    If Add_Point(SlP(N_Slope).CountOfPoints) Then Exit Function ' Points
    
    Divide_Line = SlP(N_Slope).CountOfPoints
    
    Dim a As Integer
    Dim b As Integer
    a = Lape_Lines(N_Slope, Current_L, 0)
    b = Lape_Lines(N_Slope, Current_L, 1)
    
    Dim pA As POINT
    Dim pB As POINT
    
    pA.X = Lape_Points_X(N_Slope, a)
    pA.Y = Lape_Points_Y(N_Slope, a)
    pB.X = Lape_Points_X(N_Slope, b)
    pB.Y = Lape_Points_Y(N_Slope, b)
    
    Dim cornerA As Integer
    Dim cornerC As Integer
        
    cornerA = GetGRD(pA, pB)
    
    Dim X1 As Single
    Dim Y1 As Single
    
    If cornerA = 0 Then ' 0
        X1 = X
        Y1 = pA.Y
    ElseIf cornerA < 90 Then ' 0-89
        X1 = X
        cornerC = 180 - 90 - cornerA
        Y1 = (((X - pA.X) * SinGrd(cornerA)) / SinGrd(cornerC)) + pA.Y
    ElseIf cornerA = 90 Then ' 90
        X1 = pA.X
        Y1 = Y
    ElseIf cornerA > 90 And cornerA < 180 Then ' 90-179
        X1 = X
        cornerA = 180 - cornerA
        cornerC = 180 - 90 - cornerA
        Y1 = Abs(((X - pA.X) * SinGrd(cornerA)) / SinGrd(cornerC)) + pA.Y
    ElseIf cornerA = 180 Then ' 180
        X1 = X
        Y1 = pA.Y
    ElseIf cornerA > 180 And cornerA < 270 Then
        X1 = X
        cornerA = cornerA - 180
        cornerC = 180 - 90 - cornerA
        Y1 = (((X - pA.X) * SinGrd(cornerA)) / SinGrd(cornerC)) + pA.Y
    ElseIf cornerA = 270 Then ' 180
        X1 = pA.X
        Y1 = Y
    ElseIf cornerA > 270 Then
        X1 = X
        cornerA = 360 - cornerA
        cornerC = 180 - 90 - cornerA
        Y1 = pA.Y - (((X - pA.X) * SinGrd(cornerA)) / SinGrd(cornerC))
    End If
    
    ' забивка точек X,Y
    Lape_Points_X(N_Slope, SlP(N_Slope).CountOfPoints) = X1
    Lape_Points_Y(N_Slope, SlP(N_Slope).CountOfPoints) = Y1
    
    Lape_Lines(N_Slope, Current_L, 0) = a
    Lape_Lines(N_Slope, Current_L, 1) = Divide_Line
    
    ' Добавление линии
    If Add_Point(SlP(N_Slope).CountOfLines) Then Exit Function ' Lines
    
    Lape_Lines(N_Slope, SlP(N_Slope).CountOfLines, 0) = Divide_Line
    Lape_Lines(N_Slope, SlP(N_Slope).CountOfLines, 1) = b
End Function


Sub Draw_Line(ByVal Np As Integer, ByVal X As Single, ByVal Y As Single, Optional Clear As Boolean)
    On Error Resume Next
    Line2.BorderColor = Setup.Command10.BackColor
    Line2.X1 = Lape_Points_X(N_Slope, Np)
    Line2.x2 = X
    Line2.Y1 = Lape_Points_Y(N_Slope, Np)
    Line2.y2 = Y
    Xtemp = X
    Ytemp = Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And P_A > 0 Then
    If Option4.value Then
        SetChange True
    End If
End If
End Sub

'*****************************************************************************************************

Private Sub Picture1_Resize()
    If Lapepic.Visible Then Command5.value = True
End Sub

Private Sub SetProfilData_Change(Index As Integer)
Dim Data

On Error Resume Next

'If SetProfilData(Index).Text = "" Then Exit Sub

Data = Val(Replace(SetProfilData(Index).Text, ",", "."))
Select Case Index
Case 0
    If IsNumeric(Data) = False Then Data = ConvertData(CurrentPDataRS!Overlaping)
    If Val(Data) < 0 Then Data = 0
Case 1
'    If Data = 0 Then
'    cange_max_size(0).Enabled = False
'    cange_max_size(1).Enabled = False
'    Else
'    cange_max_size(0).Enabled = True
'    cange_max_size(1).Enabled = True
'    End If
Case 2
Case 3
    If IsNumeric(Data) = False Then Data = ConvertData(CurrentPDataRS!MIN_LENGTH)
    If Val(Data) < 0 Then Data = 0
Case 4
End Select

SetProfilData(Index).Tag = ConvertData(Data, True)

End Sub

Private Sub SetProfilData_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Command2_Click
End Sub

Private Sub SplitHV1_MoveEnd()
    On Error Resume Next
    ResizePic = False
'    SuperRuler1.Width = Me.Picture1.Width
'    SuperRuler1.ScaleLeft = Picture1.ScaleLeft
'    SuperRuler1.ScaleWidth = Picture1.ScaleWidth
'    SuperRuler1.Refresh
End Sub


Private Sub SplitHV2_MoveEnd()
    On Error Resume Next
    ResizePic = False
End Sub


Private Sub sTabFx1_Click(Index As Integer, Key As String, Caption As String)
On Error Resume Next

'If OfficeStart.HistoryWorking = False Then SetChange True

ValueFieldsOnOff False

Select Case Index
Case 0
    Frame9.ZOrder 0
    Frame9.Visible = True
    Select_Option_Draw
Case 1
    Frame8.ZOrder 0
    Frame8.Visible = True
    Select_Option_Correction
Case 3
    Frame7.ZOrder 0
    Frame7.Visible = True
    Select_Option_Calculate
Case 2
    Frame11.ZOrder 0
    Frame11.Visible = True
    Select_Option_Rotate
End Select
End Sub


Private Sub sTabFx1_Resize()
On Error Resume Next
Frame9.Width = sTabFx1.Width - 200
Frame8.Width = Frame9.Width
Frame7.Width = Frame9.Width
Frame11.Width = Frame9.Width
End Sub

'
' SuperRuler1
'
Private Sub SuperRuler1_MouseDown(Button As Integer, Shift As Integer, value As Single)
SuperRuler1_Move 1, 0, SuperRuler1.GetCurrentPos, 0
SuperRuler1.MousePointer = 9
End Sub

Private Sub SuperRuler1_MouseUp(Button As Integer, Shift As Integer, value As Single)
Lapepic.Draw_Systems Lapepic.Picture1
SuperRuler1.MousePointer = 0
End Sub

Private Sub SuperRuler1_Move(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And SuperRuler1.MousePointer = 9 Then
Lapepic.Picture1.ScaleLeft = Lapepic.Picture1.ScaleLeft - X
SuperRuler1.ScaleLeft = Picture1.ScaleLeft
SuperRuler1.Refresh
End If
End Sub

'
' SuperRuler2
'
Private Sub SuperRuler2_MouseDown(Button As Integer, Shift As Integer, value As Single)
SuperRuler2_Move 1, 0, 0, SuperRuler2.GetCurrentPos
SuperRuler2.MousePointer = 7
End Sub

Private Sub SuperRuler2_MouseUp(Button As Integer, Shift As Integer, value As Single)
Lapepic.Draw_Systems Lapepic.Picture1
SuperRuler2.MousePointer = 0
End Sub

Private Sub SuperRuler2_Move(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And SuperRuler2.MousePointer = 7 Then
Lapepic.Picture1.ScaleTop = Lapepic.Picture1.ScaleTop - Y
SuperRuler2.ScaleTop = Picture1.ScaleTop
SuperRuler2.Refresh
End If
End Sub

' Y
Private Sub Text4_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If Text4 <> "" And KeyAscii = 13 Then
        
        If OptionDMM = "Mdraw" Or OptionDMM = "Msel" Then
            
            SetChange True
            Lape_Points_Y(N_Slope, P_B) = Lape_Points_Y(N_Slope, P_A) + ConvertData(Text4, True)
            Draw_Systems Me.Picture1

        ElseIf OptionDMM = "Mset" Then
        
           Me.Line2.Visible = True
           Me.Line2.y2 = Me.Line2.Y1 + ConvertData(Text4.Text, True)
           Picture1.SetFocus
        
        ElseIf OptionDMM = "Msheet" Then
        
            Dim len1 As Single
            
            If IsNumeric(Text4) = False Then Exit Sub
            
            SetChange True
            Dim cul As cList
            For Each cul In SelectLists.Items
                len1 = List_Properties_Length(N_Slope, cul.List)
                List_Properties_Length(N_Slope, cul.List) = ConvertData(Text4, True)
                List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) - (len1 - List_Properties_Length(N_Slope, cul.List))
            Next
               
            Lapepic.Draw_Systems Lapepic.Picture1
            
        End If
        
    End If

End Sub

' X
Private Sub Text5_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If Text5 <> "" And KeyAscii = 13 Then

        If OptionDMM = "Msel" Then
            
            SetChange True
            Lape_Points_X(N_Slope, P_B) = Lape_Points_X(N_Slope, P_A) + ConvertData(Text5, True)
            Draw_Systems Me.Picture1
            
        ElseIf OptionDMM = "Mset" Then
            
            Me.Line2.Visible = True
            Me.Line2.x2 = Me.Line2.X1 + ConvertData(Text5.Text, True)
            Picture1.SetFocus
            
        End If

    End If

End Sub


Private Sub txt_CL_Click()
    On Error Resume Next
    
    If txt_CL <> 0 Then
        Dim len1 As Single
    
        Dim cul As cList
        For Each cul In SelectLists.Items
            len1 = List_Properties_Length(N_Slope, cul.List)
            List_Properties_Length(N_Slope, cul.List) = ConvertData(txt_CL.Text, True)
            List_Properties_PY(N_Slope, cul.List) = List_Properties_PY(N_Slope, cul.List) - (len1 - List_Properties_Length(N_Slope, cul.List))
        Next

    Else
        If Check5.value = 1 Then
            Check5_Click
        Else
            Check1_Click
        End If

    End If

    Lapepic.Draw_Systems Lapepic.Picture1
End Sub


Private Sub Command2_Click()
    isPushCalc = False
    calc True
    SetChange True
    Draw_Systems Me.Picture1
End Sub


Sub calc(Optional use_start_points As Boolean)

        Dim ans As Integer
        Dim ProfilName As String
        Dim FactoryName As String
        Dim mp As Integer
        
        On Error Resume Next
        
        mp = OfficeStart.MousePointer

        ProfilName = Label4.Caption
        If ProfilName = "" Then
            If FileNameExtension = ".rbp" Then MsgBox lng.GetResIDstring(1480), vbCritical: Exit Sub
        End If
        FactoryName = Label5.Caption

        OfficeStart.MousePointer = 11
        OfficeStart.Enabled = False
        
        ' Очистка расчета ската
        Dim L As Integer
        For L = 1 To MAXSLOPELISTS
            List_Properties_PY(N_Slope, L) = 0
            List_Properties_PX(N_Slope, L) = 0
            List_Properties_Length(N_Slope, L) = 0
        Next
        SlP(N_Slope).CountSheets = 0
        
        ' Кеш вычислений крайних точек и начала расчета
        If isPushCalc Or (SlP(N_Slope).Pn_Red_lines = 0 Or SlP(N_Slope).Pn_StartLC = 0) Then
            ' ВЫЯВЛЕНИЕ ТОЧЕК ПО ВЕРТИКАЛИ
            If SlP(N_Slope).Pn_Red_lines = 0 Or use_start_points Then vert_hor_point N_Slope
            ' ВЫЯВЛЕНИЕ ТОЧЕК ПО ГОРИЗОНТАЛИ
            If SlP(N_Slope).Pn_StartLC = 0 Or use_start_points Then right_left_start_point N_Slope
            isPushCalc = False
            If SlP(N_Slope).Pn_Red_lines = 0 Or SlP(N_Slope).Pn_StartLC = 0 Then GoTo ERR
        End If

        On Error GoTo ERR

        isSave = True
'        Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(2).Picture
        
        Set CurrentPDataRS = GetProfilData(ProfilName, GetFactoryID(FactoryName))
        If CurrentPDataRS Is Nothing Then Exit Sub
    
        ' Сортировка точек
        Dim isSort As Boolean
        Dim Lape_Lines_out() As Integer ' Свойства линий
        ReDim Lape_Lines_out(1 To MAXSLOPES, 1 To MAXSLOPELINE, 1)
        Dim CountOfLines_out As Integer
        Dim CountOfPoints_out As Integer
    
        ' Сортировка и упорядочевание точек
        isSort = PointsSort(Lape_Lines, Lape_Lines_out, N_Slope, SlP(N_Slope).CountOfLines, SlP(N_Slope).CountOfPoints, CountOfLines_out, CountOfPoints_out)
        If isSort = False Then GoTo ERR
    
        ' Расчет площади фигуры
        SlP(N_Slope).Sf = PolygonArea(CountOfLines_out, N_Slope, Lape_Lines_out, Lape_Points_X, Lape_Points_Y)
        If SlP(N_Slope).Sf = 0 Then GoTo ERR
        
        ' Определение стаической высоты для расчета
        Dim ModuleHeight As Integer
        If calc_type(1).value Then
            If Check1.value Or Check5.value Then
                ModuleHeight = ConvertData(txt_CL, True) ' Данные из комбобокса
            Else
                ModuleHeight = mCint(CurrentPDataRS!Height)
            End If
        Else
            If SetProfilData(2).Tag > 0 Then
                ModuleHeight = SetProfilData(2).Tag
            Else
                ModuleHeight = SetProfilData(4).Tag
            End If
        End If
        
        ' Установка свойств материала
        ans = Plgs(LNC).Dll.SetMaterial _
        (CurrentPDataRS!Width, CurrentPDataRS![WORK_WIDTH], SetProfilData(1).Tag, SetProfilData(0).Tag, _
        SetProfilData(3).Tag, SetProfilData(4).Tag, _
        ModuleHeight)

        ' Установка параметров расчета
        If Setup.Option1.value = True Then Plgs(LNC).Dll.VerticalDirection = 0 Else: Plgs(LNC).Dll.VerticalDirection = 1
        If Setup.Option3.value = True Then Plgs(LNC).Dll.HorizontalDirection = 1 Else: Plgs(LNC).Dll.HorizontalDirection = 0
    
        ' Установка фиксов
        If calc_type(0).value Then
            Dim Fixes() As Long
            
            ' PLATE_MIN_HEIGHT_CUT - 0
            ' PLATE_NORMAL_HEIGHT_ROWS - 3
            ' PLATE_MAX_HEIGHT_ROWS_MIN_CUTS - 1
            
            If IsSetProfilData(3).value = 1 Then ReDim Fixes(ArraySize(Fixes)): Fixes(ArraySize(Fixes) - 1) = 0
            If IsSetProfilData(4).value = 1 Then ReDim Fixes(ArraySize(Fixes)): Fixes(ArraySize(Fixes) - 1) = 1  ' 3
            
            Plgs(LNC).Dll.FixesCount = ArraySize(Fixes)
            If Plgs(LNC).Dll.FixesCount > 0 Then Plgs(LNC).Dll.Fixes Fixes
            
        Else
            Plgs(LNC).Dll.FixesCount = -1
        End If

        ' Установка метода расчета
        If Check1.value Or Check5.value Or SetProfilData(2).Tag > 0 Then
            Plgs(LNC).Dll.UseAdditionalMethodCalc = True
        Else
            Plgs(LNC).Dll.UseAdditionalMethodCalc = False
        End If

        ' Расчет
        NERROR = Plgs(LNC).Dll.calc _
        (Lape_Lines_out, N_Slope, CountOfLines_out, CountOfPoints_out, _
        Lape_Points_X, Lape_Points_Y, SlP(N_Slope).PX_StartLC, Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_Red_lines), SlP(N_Slope).CountSheets, _
        List_Properties_PX, List_Properties_PY, 0, List_Properties_Length, 0)
    
        If NERROR < 1 Then GoTo ERR
        If SlP(N_Slope).CountSheets = 0 Then GoTo ERR
            
        SlP(N_Slope).ProfilName = ProfilName
        SlP(N_Slope).Factory_Name = IIf(Label5 <> "", Label5, Factory_Name)
    
        If IsLoadForm("Lapepic") Then
'            Lapepic.sTabFx1.SelectTab 3
            Frame7.ZOrder 0
            Frame7.Visible = True
            Select_Option_Calculate
        End If
    
'        Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(1).Picture

        OptionDMM = "Msheet"
        OfficeStart.MousePointer = 99
        OfficeStart.Enabled = True
        Exit Sub
ERR:
        
        On Error Resume Next
        OfficeStart.MousePointer = mp
        OfficeStart.Enabled = True
        
        If SlP(N_Slope).Pn_Red_lines = 0 Or SlP(N_Slope).Pn_StartLC = 0 Then
            WriteError "Start points", lng.GetResIDstring(1502) & vbNewLine & "ErrorN: " & NERROR
        ElseIf isSort = False Then
            WriteError "Operation SORT", lng.GetResIDstring(1492) & vbNewLine & "ErrorN: " & NERROR
            SlP(N_Slope).Pn_Red_lines = 0: SlP(N_Slope).CountSheets = 0
            Lapepic.sTabFx1.SelectTab 0
        ElseIf SlP(N_Slope).Sf = 0 Then
            WriteError "Area", "" & vbNewLine & "ErrorN: " & NERROR
        Else
            ' TODO:// Писать дамп SlP,
            WriteError "Plugin." & Plgs(LNC).Pname, IIf(Plgs(LNC).ERR <> "", Plgs(LNC).ERR, "") & vbNewLine & "ErrorN: " & NERROR
            '& IIf(Plgs(LNC).Dll.ERRDescription <> "", vbNewLine & Plgs(LNC).Dll.ERRDescription, "")
        End If
        
        ERR.Clear
End Sub


'
' РАЗРЕЗ ЛИСТА
'

Private Sub CutList(Button As Integer, ByVal Y As Single, N_list As Integer)
        Dim coefficient As Single
        On Error GoTo ERR

        SlP(N_Slope).CountSheets = SlP(N_Slope).CountSheets + 1
            
        List_Properties_PX(N_Slope, SlP(N_Slope).CountSheets) = List_Properties_PX(N_Slope, N_list)
        List_Properties_PY(N_Slope, SlP(N_Slope).CountSheets) = List_Properties_PY(N_Slope, N_list)
        List_Properties_Length(N_Slope, SlP(N_Slope).CountSheets) = List_Properties_Length(N_Slope, N_list)

        If Button = 1 Then Y_cut = 0

        Y = Y - SetProfilData(1).Tag

        If Not Button = 2 Then
    
            If N_list Mod 2 And Check4.value = 1 Then
                ' Разрез в шахматном порядке с заданным шагом
                List_Properties_PY(N_Slope, N_list) = Format(Y + Text6.Text, "###.0#")
            Else
                ' Разрез стандартным способом
                List_Properties_PY(N_Slope, N_list) = Format(Y, "0.00")
            End If
    
        Else
            List_Properties_PY(N_Slope, N_list) = Y_cut ' Востановление запомненого разреза
            Y = Y_cut
        End If

        ' предыдущий лист (режущ) нижний
        List_Properties_Length(N_Slope, N_list) = List_Properties_Length(N_Slope, N_list) - (List_Properties_PY(N_Slope, SlP(N_Slope).CountSheets) - List_Properties_PY(N_Slope, N_list))

        If SetProfilData(1).Tag <> 0 Then ' Разрез с учетом шага профиля (шаг волны)
            ' Проверка (на разрез) невыполнимых длин
            If Check5.value = True Then
                Y = Y - SetProfilData(1).Tag
                coefficient = SetProfilData(1).Tag - SetProfilData(0).Tag
            Else
                coefficient = SetProfilData(1).Tag - ((Y - Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_Red_lines)) Mod SetProfilData(1).Tag) ' коэф
            End If
        Else
            coefficient = SetProfilData(0).Tag * -1
        End If

        List_Properties_Length(N_Slope, SlP(N_Slope).CountSheets) = List_Properties_PY(N_Slope, SlP(N_Slope).CountSheets) - (List_Properties_PY(N_Slope, N_list) + coefficient) ' следующий (добавл)
        List_Properties_Length(N_Slope, N_list) = List_Properties_Length(N_Slope, N_list) + coefficient + SetProfilData(0).Tag
        List_Properties_PY(N_Slope, N_list) = List_Properties_PY(N_Slope, N_list) + coefficient + SetProfilData(0).Tag

        If Not Button = 2 Then Y_cut = Format(Y, "0.00") ' Запоминание разреза (зеленая линия)

        Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
        OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.12." & ERR.Source & "]", ERR.Number, ERR.Description
        Resume Next
End Sub


Private Sub Комманда1_Click()
On Error Resume Next

    Lapepic.Picture1.ScaleTop = Lapepic.Picture1.ScaleTop + Lapepic.Picture1.ScaleHeight * 0.01
    Lapepic.Draw_Systems Lapepic.Picture1
    SuperRuler2.ScaleTop = Picture1.ScaleTop
'    SuperRuler2.ScaleHeight = Picture1.ScaleHeight
    SuperRuler2.Refresh
End Sub


Private Sub Комманда2_Click()
On Error Resume Next

    Lapepic.Picture1.ScaleLeft = Lapepic.Picture1.ScaleLeft - Lapepic.Picture1.ScaleWidth * 0.01
    Lapepic.Draw_Systems Lapepic.Picture1
    SuperRuler1.ScaleLeft = Picture1.ScaleLeft
    SuperRuler1.Refresh
End Sub


Private Sub Комманда3_Click()
On Error Resume Next

    Lapepic.Picture1.ScaleLeft = Lapepic.Picture1.ScaleLeft + Lapepic.Picture1.ScaleWidth * 0.01
    Lapepic.Draw_Systems Lapepic.Picture1
    SuperRuler1.ScaleLeft = Picture1.ScaleLeft
    SuperRuler1.Refresh
End Sub


Private Sub Комманда4_Click()
On Error Resume Next

    Lapepic.Picture1.ScaleTop = Lapepic.Picture1.ScaleTop - Lapepic.Picture1.ScaleHeight * 0.01
    Lapepic.Draw_Systems Lapepic.Picture1
    SuperRuler2.ScaleTop = Picture1.ScaleTop
'    SuperRuler2.ScaleHeight = Picture1.ScaleHeight
    SuperRuler2.Refresh
End Sub



Private Sub Метка2_Change()
On Error Resume Next
Text1 = Метка2 / 2
End Sub


Private Sub Метка2_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    Dim X As Single, Y As Single
    Dim x2 As Single
    Dim y2 As Single
    
    On Error Resume Next
    
    If OptionDMM = "Mdraw" Or OptionDMM = "Msel" Then
    
        SetChange True
        
        x2 = ConvertData(Метка2, True)
        
        ' Координаты точки B относительно точки A
        Y = x2 * (ConvertData(Lapepic.Text4, True) / ab)
        If Lapepic.Text5 <= 0 And Lapepic.Text4 >= 0 Then
            Lape_Points_X(N_Slope, P_B) = Lape_Points_X(N_Slope, P_A) - Sqr(x2 * x2 - Y * Y)
            Lape_Points_Y(N_Slope, P_B) = Lape_Points_Y(N_Slope, P_A) + Y
        ElseIf Lapepic.Text5 <= 0 And Lapepic.Text4 <= 0 Then
            Lape_Points_X(N_Slope, P_B) = Lape_Points_X(N_Slope, P_A) - Sqr(x2 * x2 - Y * Y)
            Lape_Points_Y(N_Slope, P_B) = Lape_Points_Y(N_Slope, P_A) - Abs(Y)
        ElseIf Lapepic.Text5 >= 0 And Lapepic.Text4 <= 0 Then
            Lape_Points_X(N_Slope, P_B) = Lape_Points_X(N_Slope, P_A) + Sqr(x2 * x2 - Y * Y)
            Lape_Points_Y(N_Slope, P_B) = Lape_Points_Y(N_Slope, P_A) - Abs(Y)
        ElseIf Lapepic.Text5 >= 0 And Lapepic.Text4 >= 0 Then
            Lape_Points_X(N_Slope, P_B) = Lape_Points_X(N_Slope, P_A) + Sqr(x2 * x2 - Y * Y)
            Lape_Points_Y(N_Slope, P_B) = Lape_Points_Y(N_Slope, P_A) + Y
        End If
        
        Draw_Systems Me.Picture1
        
    ElseIf OptionDMM = "Mset" Then
    
        SetChange True
    
        x2 = ConvertData(Метка2, True)
        
        ' Координаты точки B относительно точки A
        Dim ab_ As Single
        ab_ = Sqr((Lapepic.Text5) ^ 2 + (Lapepic.Text4) ^ 2)
        Y = ConvertData(Метка2 * (Lapepic.Text4 / ab_), True)
        If Lapepic.Text5 <= 0 And Lapepic.Text4 >= 0 Then
            Me.Line2.x2 = Me.Line2.X1 - Sqr(x2 * x2 - Y * Y)
            Me.Line2.y2 = Me.Line2.Y1 + Y
        ElseIf Lapepic.Text5 <= 0 And Lapepic.Text4 <= 0 Then
            Me.Line2.x2 = Me.Line2.X1 - Sqr(x2 * x2 - Y * Y)
            Me.Line2.y2 = Me.Line2.Y1 - Abs(Y)
        ElseIf Lapepic.Text5 >= 0 And Lapepic.Text4 <= 0 Then
            Me.Line2.x2 = Me.Line2.X1 + Sqr(x2 * x2 - Y * Y)
            Me.Line2.y2 = Me.Line2.Y1 - Abs(Y)
        ElseIf Lapepic.Text5 >= 0 And Lapepic.Text4 >= 0 Then
            Me.Line2.x2 = Me.Line2.X1 + Sqr(x2 * x2 - Y * Y)
            Me.Line2.y2 = Me.Line2.Y1 + Y
        End If
        
        Picture1.SetFocus
        
    End If
    
End Sub


Sub SelectList(X As Single, Y As Single, Shift As Integer, Button As Integer)
Dim cl As Integer
Dim List As New cList

On Error Resume Next

cl = Find_list(X, Y)

If cl = 0 Then
    SelectLists.Clear
    Exit Sub
End If

If Button = 2 Then Exit Sub

If Shift = 2 Then
    
    If Not SelectLists.Item(cl) Is Nothing Then
    
       SelectLists.Remove cl
    
    Else
    
        List.List = cl
        SelectLists.Add List, cl
    
    End If

Else

    SelectLists.Clear
    List.List = cl
    SelectLists.Add List, cl

End If

End Sub


Public Function SetDrawBorder()
    ' нахождение крайней левой и правой точки по X
    On Error Resume Next
    
    Dim P As Integer
        XMin = 999999
        XMax = -999999
        For P = 1 To SlP(N_Slope).CountOfLines Step 1
            If XMin >= Lape_Points_X(N_Slope, P) Then XMin = Lape_Points_X(N_Slope, P)
            If XMax <= Lape_Points_X(N_Slope, P) Then XMax = Lape_Points_X(N_Slope, P)
        Next P

        ' Верхние и нижнии границы
        YMin = 999999
        YMAx = -999999
        For P = 1 To SlP(N_Slope).CountOfLines Step 1
            If YMin >= Lape_Points_Y(N_Slope, P) Then YMin = Lape_Points_Y(N_Slope, P)
            If YMAx <= Lape_Points_Y(N_Slope, P) Then YMAx = Lape_Points_Y(N_Slope, P)
        Next P
End Function


Public Function SavePolygon(Optional back As Boolean)
On Error Resume Next

    Dim i As Integer
'    If back = False Then
        ReDim SaveLape_Points_X(SlP(N_Slope).CountOfPoints)
        ReDim SaveLape_Points_Y(SlP(N_Slope).CountOfPoints)
'    End If

    For i = 1# To SlP(N_Slope).CountOfPoints Step 1

'        If back Then
'            Lape_Points_X(N_Slope, i) = SaveLape_Points_X(i)
'            Lape_Points_Y(N_Slope, i) = SaveLape_Points_Y(i)
'        Else
            SaveLape_Points_X(i) = Lape_Points_X(N_Slope, i)
            SaveLape_Points_Y(i) = Lape_Points_Y(N_Slope, i)
'        End If

    Next

End Function

Public Function PolygonRotate(a As Integer)
    Dim i As Long
    Dim pPoint As POINT
    Dim pOrigin As POINT
    Dim pResult As POINT
    
    On Error Resume Next

    pOrigin.X = (XMin + XMax) / 2
    pOrigin.Y = (YMin + YMAx) / 2

    For i = 1# To SlP(N_Slope).CountOfPoints Step 1
    
        pPoint.X = SaveLape_Points_X(i)
        pPoint.Y = SaveLape_Points_Y(i)
    
        pResult = RotatePoint(pPoint, pOrigin, a)
    
        Lape_Points_X(N_Slope, i) = pResult.X
        Lape_Points_Y(N_Slope, i) = pResult.Y
    
    Next i

    Draw_Systems Me.Picture1

End Function

Public Function ListEdit(ShowForm As Boolean)
On Error Resume Next

    If ShowForm Then
        If OptionDMM = "Msheet" And SlP(N_Slope).Pn_StartLC <> 0 And SlP(N_Slope).Pn_StartLC <> 0 And SlP(N_Slope).CountSheets <> 0 Then
            Move_and_change.Show vbModeless, OfficeStart
        End If
        OfficeStart.menu_view_m(8).Checked = True
    Else
        Unload Move_and_change
        OfficeStart.menu_view_m(8).Checked = False
    End If
End Function


Sub Draw_Systems(MainPic As PictureBox)
      If FEXIT Then Exit Sub
      
      Dim isFigure As Boolean
      Dim l01AA As Single
      Dim l01AC As Single
      Dim Px As Single
      Dim Py As Single
      Dim l01B6 As Single
      Dim l01B8 As Single
      Dim l01BA As Single
      Dim l01BC As Single
      Dim l01C0 As Single

      Dim i As Integer
      Dim Summ As Single
      
      Dim p1 As POINT
      Dim p2 As POINT
      Dim L As LOGPEN
      
      On Error GoTo ERR

      'Error = False

      MainPic.Cls
'      Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(2).Picture
      
      ' Строим координатную сетку
      If Setup.Check18.value = 1 Then
            Dim Y As Long
            Dim X As Long

'            For Y = -SuperRuler2.MaxV To SuperRuler2.MaxV Step SuperRuler2.UserScale
'                For X = -SuperRuler1.MaxH To SuperRuler1.MaxH Step SuperRuler1.UserScale
'                    MainPic.PSet (X, Y), vbBlack
'                Next X
'            Next Y

            For Y = -SuperRuler2.MaxV To SuperRuler2.MaxV Step SuperRuler2.UserScale
                For X = -SuperRuler1.MaxH To SuperRuler1.MaxH Step SuperRuler1.UserScale
                    SetPixel Picture1.hdc, Picture1.ScaleX(X, Picture1.ScaleMode, vbPixels), Picture1.ScaleY(Y, Picture1.ScaleMode, vbPixels), vbBlack
                Next X
            Next Y

      End If

      l01AA = MainPic.ScaleWidth / 90
      l01AC = 2 * l01AA
      Px = MainPic.ScaleLeft + MainPic.ScaleWidth * 0.99
      Py = MainPic.ScaleTop + MainPic.ScaleHeight * 0.96

      '---------------------------------------------------------------------- 'построитель линии в низу
'      MainPic.Line (Px - 125, Py)-(Px - 25, Py) '
'      MainPic.Line (Px - 125, Py - l01AA)-(Px - 125, Py + l01AA)
'      MainPic.Line (Px - 25, Py - l01AA)-(Px - 25, Py + l01AA) '
'      MainPic.PSet (MainPic.ScaleLeft + MainPic.ScaleWidth * 0.91 - 62, MainPic.ScaleTop + MainPic.ScaleHeight * 0.91), RGB(255, 255, 255) '
'      MainPic.Print "(1 x " & Format(ConvertData(MainPic.ScaleWidth, True), "###") & ")m"
      '-------------------------------------------------------------------end
      
      If OptionDMM = "Msheet" And SlP(N_Slope).Pn_StartLC > 0 Then

          l01AA = MainPic.ScaleWidth / 100
          l01AC = 2 * l01AA
          l01B6 = 3 * l01AA
          l01B8 = MainPic.ScaleTop
  
          If Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC) < SlP(N_Slope).PX_StartLC Then
              l01BA = Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC)
              l01BC = SlP(N_Slope).PX_StartLC
          Else
              l01BA = SlP(N_Slope).PX_StartLC
              l01BC = Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC)
          End If

          If l01BC - l01BA < 2 * l01B6 Then
              Summ = 2 * l01B6
              l01C0 = l01BC + l01AC
          Else
              Summ = l01AC
              l01C0 = ((SlP(N_Slope).PX_StartLC + Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC)) / 2) - l01AC
          End If
  
          MainPic.DrawStyle = 2
          MainPic.Line (Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_StartLC))-(Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), l01B8 - l01AC)
          MainPic.DrawStyle = 0
          MainPic.Line (l01BA - l01AC, l01B8 - l01B6)-(l01BC + Summ, l01B8 - l01B6)
          MainPic.Line (SlP(N_Slope).PX_StartLC, l01B8 - l01AA)-(SlP(N_Slope).PX_StartLC, l01B8 - 2 * l01B6)
          MainPic.Line (Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), l01B8 - l01AA)-(Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC), l01B8 - 2 * l01B6)
          MainPic.Line (Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC) - l01AA, l01B8 - 2 * l01AC)-(Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC) + l01AA, l01B8 - l01AC)
          MainPic.Line (SlP(N_Slope).PX_StartLC - l01AA, l01B8 - 2 * l01AC)-(SlP(N_Slope).PX_StartLC + l01AA, l01B8 - l01AC)
          MainPic.PSet (l01C0, l01B8 - 0.5 * l01AA), RGB(255, 255, 255)
          MainPic.Print Format$(ConvertData(SlP(N_Slope).PX_StartLC - Lape_Points_X(N_Slope, SlP(N_Slope).Pn_StartLC)), "####")
  
      End If
      
      '# ----
      '# ----  - RED LINE Линия отсчета по  вертикалм
      '# ----
      If OptionDMM = "Msheet" And SlP(N_Slope).Pn_Red_lines > 0 Then
            MainPic.Line (MainPic.ScaleLeft, Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_Red_lines))-(MainPic.ScaleLeft + MainPic.ScaleWidth, Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_Red_lines)), vbRed
      End If

      '''''''''''''''''''''''''''
      ' Прорисовка чертежа
      '''''''''''''''''''''''''''
      
DRAWFIGURE:

      If SlP(N_Slope).CountOfLines > 0 Or SlP(N_Slope).CountOfPoints > 0 Then

        isFigure = False
    
        MainPic.DrawStyle = 0
        MainPic.FontSize = 8
    
        Dim LineColor As Integer
        For i = 1 To SlP(N_Slope).CountOfLines Step 1
        
          If Lape_Lines(N_Slope, i, 1) > 0 Then
              
            ' Подсчет погонных метров
            Llen = (Lape_Points_X(N_Slope, Lape_Lines(N_Slope, i, 1)) - Lape_Points_X(N_Slope, Lape_Lines(N_Slope, i, 0))) ^ 2 + _
            (Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, i, 1)) - Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, i, 0))) ^ 2
            
            If Llen <> -1 Then
            
                Llen = Sqr(Llen)

                p1.X = Round(Lape_Points_X(N_Slope, Lape_Lines(N_Slope, i, 0)))
                p1.Y = Round(Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, i, 0)))
                p2.X = Round(Lape_Points_X(N_Slope, Lape_Lines(N_Slope, i, 1)))
                p2.Y = Round(Lape_Points_Y(N_Slope, Lape_Lines(N_Slope, i, 1)))
                L.lopnColor = Setup.Command10.BackColor
                L.lopnStyle = 0
                L.lopnWidth = 3
                APILineEx MainPic, p1, p2, L
                
            End If

          End If
        
          If OfficeStart.menu_view_m(1).Checked And Llen > 0 Then ' SHOW LENGTH
              MarkLine MainPic, N_Slope, l01AA, i, Llen, 8, False
          End If
        
        Next i
        
        For i = 1 To SlP(N_Slope).CountOfPoints Step 1
            Draw_Point i, vbRed, 270
        Next i

  Else
  
      isFigure = True
      
  End If

  ' Прорисовка букв корректировки
  If Opt <> "P_Adjust" Then
  
      If OptionDMM = "Msel" Then
      Dim size As Single
      size = MainPic.FontSize
      MainPic.FontSize = 10
      MainPic.DrawStyle = 0
      MainPic.ForeColor = vbRed
      MainPic.FontBold = True

      '
      ' A-B
      '
      If P_A > 0 Then
          If Check3.value Or Check9.value Then
              MainPic.PSet (Lape_Points_X(N_Slope, P_A) - l01AA, Lape_Points_Y(N_Slope, P_A) + l01AC), MainPic.BackColor  'RGB(0, 0, 255)
              MainPic.Print "A"
          ElseIf Option4.value Then
              Draw_Point P_A, vbGreen, 270
          End If
      End If

      If P_B > 0 And (Check3.value Or Check9.value) Then
          MainPic.PSet (Lape_Points_X(N_Slope, P_B) - l01AA, Lape_Points_Y(N_Slope, P_B) + l01AC), MainPic.BackColor 'RGB(0, 0, 255)
          MainPic.Print "B"
      End If
 
      MainPic.FontBold = False
      MainPic.ForeColor = vbBlack
      MainPic.FontSize = size
      
  End If

  '''''''''''''''''''''''''''
  ' Прорисовка раскроя
  '''''''''''''''''''''''''''
  If OptionDMM = "Msheet" Then

    If OfficeStart.menu_view_m(2).Checked Then
  
        If SlP(N_Slope).CountSheets > Gl.MAXSLOPELISTS Then GoTo ERR
        If SetProfilData.Count < 5 Then GoTo ERR
        
        Dim cList As Long
        SlP(N_Slope).ListLength = 0
        
        Dim ars As Integer
'        ars = SelectLists.count
        
        MainPic.DrawMode = 9
        For i = 1 To SlP(N_Slope).CountSheets Step 1
        
          If List_Properties_Length(N_Slope, i) <> 0 Then
        
              SlP(N_Slope).ListLength = SlP(N_Slope).ListLength + List_Properties_Length(N_Slope, i)
        
              ' Отображение длин со склада txt_CL
              If Check6.value Then
                  Dim n As Integer
                  If txt_CL.ListCount > 0 Then
                      For n = 0 To txt_CL.ListCount - 1 Step 1
                          If txt_CL.List(n) = Format$(List_Properties_Length(N_Slope, i), "0") Then MainPic.Line (List_Properties_PX(N_Slope, i), List_Properties_PY(N_Slope, i))-(List_Properties_PX(N_Slope, i) + CurrentPDataRS![WORK_WIDTH], List_Properties_PY(N_Slope, i) - List_Properties_Length(N_Slope, i)), "&H00C7EEF5", BF  ' Прорисовка каждого листа
                      Next
                  End If
              End If
        
             ' Отоброжение листов
             If List_Properties_Length(N_Slope, i) > 0 Then
                 
                 If OfficeStart.menu_view_m(3).Checked Then MainPic.FillStyle = 5 ' перечеркнутом виде
                 
                 If Not SelectLists.Item(i) Is Nothing Then
                 
                    ' Отображение выбранных листов
                    MainPic.Line ( _
                     Round(List_Properties_PX(N_Slope, i)), _
                     Round(List_Properties_PY(N_Slope, i)) _
                     )-( _
                     Round(List_Properties_PX(N_Slope, i) + CurrentPDataRS![WORK_WIDTH]), _
                     Round(List_Properties_PY(N_Slope, i) - List_Properties_Length(N_Slope, i))), _
                      &H89DEE2, BF
                 
                 End If
                 
                ' Прорисовка каждого листа
                MainPic.Line ( _
                Round(List_Properties_PX(N_Slope, i)), _
                Round(List_Properties_PY(N_Slope, i)) _
                )-( _
                Round(List_Properties_PX(N_Slope, i) + CurrentPDataRS![WORK_WIDTH]), _
                Round(List_Properties_PY(N_Slope, i) - List_Properties_Length(N_Slope, i))), _
                , B

                 
                MainPic.FillStyle = 1
                 
             End If
         
             ' Отображение длин листов
             If OfficeStart.menu_view_m(6).Checked Then
                 MainPic.ForeColor = vbMagenta
                 Dim strListLen As String
                 Dim nStr As Integer
                 strListLen = ConvertData(List_Properties_Length(N_Slope, i))
                 For nStr = 1 To Len(strListLen)
                     MainPic.CurrentX = List_Properties_PX(N_Slope, i) + (CurrentPDataRS![WORK_WIDTH] / 2) - (MainPic.TextWidth("0") / 2)
                     MainPic.CurrentY = List_Properties_PY(N_Slope, i) - (List_Properties_Length(N_Slope, i) / 2) - _
                     ((Len(strListLen) * MainPic.TextHeight("0")) / 2) + ((nStr - 1) * MainPic.TextHeight("0"))
                     MainPic.Print mID$(strListLen, nStr, 1)
                 Next
                 MainPic.ForeColor = vbBlack
              End If
        
              ' Отображение шага волны
              If OfficeStart.menu_view_m(5).Checked Then
               Dim WaveStep As Integer
               WaveStep = SetProfilData(1).Tag
               If WaveStep > 5 Then
                Dim StartVaweDraw As Integer
                
                If Setup.Option2.value Then
                
                    L.lopnColor = &HCCCCCC
                    L.lopnStyle = 0
                    L.lopnWidth = 1
                
                    StartVaweDraw = List_Properties_PY(N_Slope, i) - List_Properties_Length(N_Slope, i)
                    For n = StartVaweDraw To List_Properties_PY(N_Slope, i) Step WaveStep
                    
                       p1.X = List_Properties_PX(N_Slope, i)
                       p1.Y = n
                       p2.X = List_Properties_PX(N_Slope, i) + CurrentPDataRS![WORK_WIDTH]
                       p2.Y = n
                       APILineEx MainPic, p1, p2, L
                                             
                    Next n
                    
                ElseIf Setup.Option1.value Then
                
                    L.lopnColor = &HCCCCCC
                    L.lopnStyle = 0
                    L.lopnWidth = 1
                    
                    StartVaweDraw = List_Properties_PY(N_Slope, i)
                    
'                    StartVaweDraw = List_Properties_Length(N_Slope, i) - (Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_Red_lines) - (List_Properties_PY(N_Slope, i) - List_Properties_Length(N_Slope, i)))
'                    If StartVaweDraw > 0 Then
'                        StartVaweDraw = (StartVaweDraw \ WaveStep) * WaveStep
'                        StartVaweDraw = Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_Red_lines) + StartVaweDraw
'                    Else
'                        StartVaweDraw = (StartVaweDraw \ WaveStep) * WaveStep
'                        StartVaweDraw = Lape_Points_Y(N_Slope, SlP(N_Slope).Pn_Red_lines) + StartVaweDraw - WaveStep
'                    End If
                    
                    For n = StartVaweDraw To List_Properties_PY(N_Slope, i) - List_Properties_Length(N_Slope, i) Step -WaveStep
                        
                       p1.X = List_Properties_PX(N_Slope, i)
                       p1.Y = n
                       p2.X = List_Properties_PX(N_Slope, i) + CurrentPDataRS![WORK_WIDTH]
                       p2.Y = n
                       APILineEx MainPic, p1, p2, L
                                                 
                    Next n
                    
                End If
               End If
              End If
        
              ' Отображение обозначений листов
              If OfficeStart.menu_view_m(4).Checked Then
                MainPic.PSet (List_Properties_PX(N_Slope, i), List_Properties_PY(N_Slope, i) - List_Properties_Length(N_Slope, i)), RGB(255, 255, 255)
                MainPic.ForeColor = "&H00808080"
                Dim str As String
                str = Format$(i, "00")
                MainPic.Print str
                cList = cList + 1
                MainPic.ForeColor = vbBlack
              End If
                            
          End If
        
        Next i
            
        MainPic.DrawMode = 13
         
        ' Вывод справочной информации
        label7 = Command10.Caption & ": "
        
        Dim ListLen As Single
        ListLen = SlP(N_Slope).ListLength
         
        ListLen = ListLen / 100
        label7 = label7 & Format(ListLen, "# ##0.00") & Setup.Text15 & "; "
        
        If CurrentPDataRS![WORK_WIDTH] > 0 Then
          
          ' SW
          SlP(N_Slope).Sw = ListLen * CurrentPDataRS![WORK_WIDTH] / 100 ' Забивка площади покрываемых листов по рабочей ширине
          label7 = label7 & lng.GetResIDstring(1066) & Format(SlP(N_Slope).Sw, "# ##0.00") & Setup.Text16 & "; "
          
        End If
        
        If CurrentPDataRS!Width > 0 Then
            label7 = label7 & lng.GetResIDstring(1067) & Format((ListLen * CurrentPDataRS!Width) / 100, "# ##0.00") & Setup.Text16 & "; "
        End If
        
        If SlP(N_Slope).Sf > 0 And SlP(N_Slope).ListLength > 0 Then
        
            Dim SFigure As Single
            SFigure = SlP(N_Slope).Sf / 10000
            label7 = label7 & lng.GetResIDstring(1062) & Format(SFigure, "# ##0.00") & " " & Setup.Text16
            Dim prc As Integer
            prc = 100 - (SFigure / SlP(N_Slope).Sw) * 100
            label7 = label7 & lng.GetResIDstring(1060) & prc & " %"
        
        End If

    End If
  End If
End If

'Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(1).Picture
Exit Sub

ERR:
'Me.MousePointer = 99
MainPic.ForeColor = vbBlack
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.10." & ERR.Source & "]", ERR.Number, ERR.Description
'Resume Next
End Sub

Sub ValueFieldsOnOff(Optional op As Boolean = False)

' X
ValueFieldX_OnOff op

' Y
ValueFieldY_OnOff op

' A-B
Метка2.Enabled = op
Command12.Enabled = op

' Градусы
Label8.Enabled = op
HScroll3.Enabled = op

End Sub

Sub ValueFieldY_OnOff(Optional op As Boolean = False)
Text4.Enabled = op
Command7.Enabled = op
Command9.Enabled = op
End Sub

Sub ValueFieldX_OnOff(Optional op As Boolean = False)
Text5.Enabled = op
Command3.Enabled = op
Command8.Enabled = op
End Sub


' ***********************
'
' ***********************

Sub Select_Option_Draw()
    Dim l0154 As Single
    
    On Error GoTo ERR

    If (SlP(N_Slope).Pn_Red_lines > 0 Or SlP(N_Slope).CountSheets > 0) And OfficeStart.HistoryWorking = False Then
        l0154 = MsgBox(lng.GetResIDstring(1402), 4)
        If l0154 = 6 Then
            SetChange True
            SelectLists.Clear
        Else
            Lapepic.sTabFx1.SelectTab 3
            Exit Sub
        End If
    End If
    
    label7.Text = ""
    Text7.Text = ""
    Current_L = 0

    OfficeStart.menu_view_m(0).Enabled = True
    OfficeStart.menu_view_m(2).Enabled = False
    OfficeStart.menu_view_m(3).Enabled = False
    OfficeStart.menu_view_m(4).Enabled = False
    OfficeStart.menu_view_m(5).Enabled = False
    OfficeStart.menu_view_m(6).Enabled = False
    OfficeStart.menu_view_m(8).Enabled = False
    
'    cange_max_size(0).Enabled = False
'    cange_max_size(1).Enabled = False
    
    calc_type(0).Enabled = False
    calc_type(1).Enabled = False

    If SlP(N_Slope).CountOfLines > 1 Or SlP(N_Slope).CountOfPoints > 1 Then
        sTabFx1.TabDisabled(2) = False
    Else
        sTabFx1.TabDisabled(2) = True
    End If
    
    SlP(N_Slope).Pn_Red_lines = 0
    SlP(N_Slope).CountSheets = 0
    SlP(N_Slope).Pn_StartLC = 0
    SlP(N_Slope).Sw = 0
    
    Check3.Enabled = False
    Check9.Enabled = False
    Option4.Enabled = False
  
    OptionDMM = "Mdraw"
    Label2 = lng.GetResIDstring(1434)
    P_A = 0
    P_B = 0
    SelectLists.Clear
    Cut_N = 0
    Picture1.MousePointer = 2

    If IsLoadForm("Lapepic") And SlP(N_Slope).CountOfPoints > 0 Then
        Draw_Systems Me.Picture1
    End If

    Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.4." & ERR.Source & "]", ERR.Number, ERR.Description
    Resume Next
End Sub


Private Sub Option2_Click()
    Dim l0154 As Integer
    On Error GoTo ERR
    
    label7.Text = ""
    Text7.Text = ""
    Current_L = 0

    OptionDMM = "Msheet"
    Picture1.MousePointer = 99
    
    OfficeStart.menu_view_m(2).Enabled = True
    OfficeStart.menu_view_m(3).Enabled = True
    OfficeStart.menu_view_m(3).Enabled = True
    OfficeStart.menu_view_m(4).Enabled = True
    OfficeStart.menu_view_m(5).Enabled = True
    OfficeStart.menu_view_m(6).Enabled = True
    OfficeStart.menu_view_m(8).Enabled = True
    
    SlP(N_Slope).Pn_Red_lines = 0
    SlP(N_Slope).Pn_StartLC = 0

    isPushCalc = True

    calc
    SetChange True

    Exit Sub
ERR:
'        STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.5." & ERR.Source & "]", ERR.Number, ERR.Description
    WriteError Plgs(LNC).Pname, ERR.Description
End Sub


Sub Select_Option_Correction()
    Dim l0168 As Single
    
    On Error Resume Next

    If (SlP(N_Slope).Pn_Red_lines > 0 Or SlP(N_Slope).CountSheets > 0) And OfficeStart.HistoryWorking = False Then
        l0168 = MsgBox(lng.GetResIDstring(1402), 4)
        If l0168 = 6 Then
            SetChange True
        Else
            Lapepic.sTabFx1.SelectTab 3
            Exit Sub
        End If

    End If
    
    SetDrawBorder
    Line13.Visible = False
    Line14.Visible = False
    
    label7.Text = ""
    Text7.Text = ""
    Current_L = 0

    OfficeStart.menu_view_m(0).Enabled = False
    OfficeStart.menu_view_m(2).Enabled = False
    OfficeStart.menu_view_m(3).Enabled = False
    OfficeStart.menu_view_m(4).Enabled = False
    OfficeStart.menu_view_m(5).Enabled = False
    OfficeStart.menu_view_m(6).Enabled = False
    OfficeStart.menu_view_m(8).Enabled = False
    
'    cange_max_size(0).Enabled = False
'    cange_max_size(1).Enabled = False
    
    calc_type(0).Enabled = False
    calc_type(1).Enabled = False

    If SlP(N_Slope).CountOfLines > 1 Or SlP(N_Slope).CountOfPoints > 1 Then
        sTabFx1.TabDisabled(2) = False
        Check3.Enabled = True
        Check9.Enabled = True
        Option4.Enabled = True
    Else
        sTabFx1.TabDisabled(2) = True
        Check3.Enabled = False
        Check9.Enabled = False
        Option4.Enabled = False
    End If
    
    SlP(N_Slope).Pn_Red_lines = 0
    SlP(N_Slope).CountSheets = 0
    SlP(N_Slope).Pn_StartLC = 0
    SlP(N_Slope).Sw = 0

    OptionDMM = "Msel"
    Label2 = lng.GetResIDstring(1435)
    P_A = 0: P_B = 0
    SelectLists.Clear
    Cut_N = 0
    Picture1.MousePointer = 1

    Lapepic.Picture1.MousePointer = 99
    Lapepic.Picture1.MouseIcon = LoadResPicture(101, 2)

    If IsLoadForm("Lapepic") Then
        Draw_Systems Me.Picture1
    End If

End Sub

Sub Select_Option_Calculate()

On Error Resume Next

    If Error = True Then Exit Sub

    label7.Text = ""
    Text7.Text = ""
    Current_L = 0

    Picture1.MousePointer = 99
    OfficeStart.menu_view_m(2).Enabled = True
    OfficeStart.menu_view_m(3).Enabled = True
    OfficeStart.menu_view_m(3).Enabled = True
    OfficeStart.menu_view_m(4).Enabled = True
    OfficeStart.menu_view_m(5).Enabled = True
    OfficeStart.menu_view_m(6).Enabled = True
    OfficeStart.menu_view_m(8).Enabled = True

    Check3.Enabled = False
    Check9.Enabled = False
    Option4.Enabled = False
    
'    cange_max_size(0).Enabled = True
'    cange_max_size(1).Enabled = True
    
    calc_type(0).Enabled = True
    calc_type(1).Enabled = True

    Dim l0154 As Integer
    If SlP(N_Slope).Pn_Red_lines > 0 And SlP(N_Slope).CountSheets > 0 Then
        If IsLoadForm("Lapepic") Then
'            SetChange True
            Check3.Enabled = False
            Draw_Systems Me.Picture1
        End If
        Label2 = lng.GetResIDstring(1410)
        Exit Sub
    End If

    OptionDMM = "Msheet"
    OfficeStart.menu_view_m(0).Enabled = False

    If SlP(N_Slope).Pn_Red_lines = 0 Then
        Label2 = lng.GetResIDstring(1405) '"Выберите стартовую точку карниза крыши"
    Else
        Label2 = lng.GetResIDstring(1410)
    End If

End Sub


Sub Select_Option_Rotate()
    Dim l0168 As Single
    
    On Error Resume Next

    If (SlP(N_Slope).Pn_Red_lines > 0 Or SlP(N_Slope).CountSheets > 0) And OfficeStart.HistoryWorking = False Then
        l0168 = MsgBox(lng.GetResIDstring(1402), 4)
        If l0168 = 6 Then
            SetChange True
        Else
            Lapepic.sTabFx1.SelectTab 3
            Exit Sub
        End If

    End If
    
    SetDrawBorder
    Line13.Visible = False
    Line14.Visible = False
    
    label7.Text = ""
    Text7.Text = ""
    Current_L = 0

    OfficeStart.menu_view_m(0).Enabled = False
    OfficeStart.menu_view_m(2).Enabled = False
    OfficeStart.menu_view_m(3).Enabled = False
    OfficeStart.menu_view_m(4).Enabled = False
    OfficeStart.menu_view_m(5).Enabled = False
    OfficeStart.menu_view_m(6).Enabled = False
    OfficeStart.menu_view_m(8).Enabled = False
    
'    cange_max_size(0).Enabled = False
'    cange_max_size(1).Enabled = False
    
    calc_type(0).Enabled = False
    calc_type(1).Enabled = False

    If SlP(N_Slope).CountOfLines > 1 Or SlP(N_Slope).CountOfPoints > 1 Then
        sTabFx1.TabDisabled(2) = False
        Check3.Enabled = True
        Check9.Enabled = True
        Option4.Enabled = True
    Else
        sTabFx1.TabDisabled(2) = True
        Check3.Enabled = False
        Check9.Enabled = False
        Option4.Enabled = False
    End If
    
    SlP(N_Slope).Pn_Red_lines = 0
    SlP(N_Slope).CountSheets = 0
    SlP(N_Slope).Pn_StartLC = 0
    SlP(N_Slope).Sw = 0

'    isSave = True

    OptionDMM = "Msel"
    Label2 = lng.GetResIDstring(1435)
    P_A = 0: P_B = 0
    SelectLists.Clear
    Cut_N = 0
    Picture1.MousePointer = 1

    Lapepic.Picture1.MousePointer = 99
    Lapepic.Picture1.MouseIcon = LoadResPicture(101, 2)

    SavePolygon False
    
    HScroll2.value = 0

    If IsLoadForm("Lapepic") Then
        Draw_Systems Me.Picture1
    End If

End Sub

Private Sub SwitchProfile()
    On Error GoTo ERR

    If Trim(Label4.Caption) = "" Then Exit Sub

'    Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(2).Picture
    
    Set CurrentPDataRS = GetProfilData(Label4.Caption, GetFactoryID(Label5.Caption))

    If CurrentPDataRS Is Nothing Then GoTo ERR

    Label4 = Trim(Label4.Caption)
    
    SetProfilData(1).Text = -1
    SetProfilData(0).Text = -1
    SetProfilData(2).Text = -1
    SetProfilData(3).Text = -1
    SetProfilData(4).Text = -1
    
    ' Заполнение
    SetProfilData(1).Text = ConvertData(CurrentPDataRS!Step)
    SetProfilData(0).Text = ConvertData(CurrentPDataRS!Overlaping)
    SetProfilData(2).Text = ConvertData(CurrentPDataRS!Height)
    SetProfilData(3).Text = ConvertData(CurrentPDataRS!MIN_LENGTH)
    SetProfilData(4).Text = ConvertData(CurrentPDataRS!MAX_LENGTH)
    
    L1 = CheckNullNomber(CurrentPDataRS!L1)
    L2 = CheckNullNomber(CurrentPDataRS!L2)
    wl = CheckNullNomber(CurrentPDataRS!wl)

Exit Sub
ERR:
'   STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.1." & ERR.Source & "]", ERR.Number, ERR.Description
    WriteError Plgs(LNC).Pname & ": " & ERR.Description
    Command2_Click
End Sub
