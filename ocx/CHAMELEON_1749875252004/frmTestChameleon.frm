VERSION 5.00
Begin VB.Form frmTestChameleon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chameleon Button Test"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "frmTestChameleon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmTestChameleon.frx":0BC2
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton chameleonButton1 
      Height          =   615
      Left            =   5640
      TabIndex        =   42
      Top             =   3000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      BTYPE           =   7
      TX              =   ""
      ENAB            =   -1  'True
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
      MICON           =   "frmTestChameleon.frx":3D00
      PICN            =   "frmTestChameleon.frx":3D1C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton KDE2 
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   38
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "KDE 2"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16576
      BCOLO           =   16576
      FCOL            =   65535
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":416E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton High3D 
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   6480
      Width           =   1575
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   12
      TX              =   "3D Hover"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":418A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   671
      TabIndex        =   27
      Top             =   0
      Width           =   10065
      Begin VB.ComboBox btnType 
         Height          =   315
         ItemData        =   "frmTestChameleon.frx":41A6
         Left            =   3600
         List            =   "frmTestChameleon.frx":41D5
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   0
         Width           =   1215
      End
      Begin prjChameleon.chameleonButton toolB 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "New"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTestChameleon.frx":4244
         PICN            =   "frmTestChameleon.frx":4260
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton toolB 
         Height          =   360
         Index           =   1
         Left            =   480
         TabIndex        =   29
         ToolTipText     =   "Open"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTestChameleon.frx":4372
         PICN            =   "frmTestChameleon.frx":438E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton toolB 
         Height          =   360
         Index           =   2
         Left            =   840
         TabIndex        =   30
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTestChameleon.frx":44A0
         PICN            =   "frmTestChameleon.frx":44BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton toolB 
         Height          =   360
         Index           =   3
         Left            =   1320
         TabIndex        =   31
         ToolTipText     =   "Cut"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTestChameleon.frx":45CE
         PICN            =   "frmTestChameleon.frx":45EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton toolB 
         Height          =   360
         Index           =   4
         Left            =   1680
         TabIndex        =   33
         ToolTipText     =   "Copy"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTestChameleon.frx":46FC
         PICN            =   "frmTestChameleon.frx":4718
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton toolB 
         Height          =   360
         Index           =   5
         Left            =   2040
         TabIndex        =   34
         ToolTipText     =   "Paste"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTestChameleon.frx":482A
         PICN            =   "frmTestChameleon.frx":4846
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   35
         Top             =   60
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   163
         X2              =   163
         Y1              =   0
         Y2              =   24
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   164
         X2              =   164
         Y1              =   0
         Y2              =   24
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   84
         X2              =   84
         Y1              =   0
         Y2              =   24
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   83
         X2              =   83
         Y1              =   0
         Y2              =   24
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   330
         Y1              =   25
         Y2              =   25
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   0
         X2              =   330
         Y1              =   26
         Y2              =   26
      End
   End
   Begin prjChameleon.chameleonButton cbOXP 
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "Office XP"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":4958
      PICN            =   "frmTestChameleon.frx":4974
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbWinXPcool 
      Height          =   495
      Left            =   3480
      TabIndex        =   18
      ToolTipText     =   "This is another checkbox"
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "WinXP Custom"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   12640511
      FCOL            =   0
      FCOLO           =   128
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":4DC6
      PICN            =   "frmTestChameleon.frx":4DE2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbWinXP 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Windows &XP"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":4DFE
      PICN            =   "frmTestChameleon.frx":4E1A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbFlatCool 
      Height          =   495
      Left            =   3480
      TabIndex        =   22
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "Flat Custom"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421376
      BCOLO           =   16576
      FCOL            =   0
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":513C
      PICN            =   "frmTestChameleon.frx":5158
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbN6Cool 
      Height          =   495
      Left            =   3480
      TabIndex        =   21
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "N6 Custom"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16512
      BCOLO           =   16512
      FCOL            =   14737632
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5174
      PICN            =   "frmTestChameleon.frx":5190
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   -1  'True
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbJavaCool 
      Height          =   495
      Left            =   3480
      TabIndex        =   20
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "Java Custom"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   32896
      BCOLO           =   32896
      FCOL            =   4194304
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":52A2
      PICN            =   "frmTestChameleon.frx":52BE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbMacCool 
      Height          =   495
      Left            =   3480
      TabIndex        =   19
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Mac Custom"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   65280
      FCOLO           =   49344
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":52DA
      PICN            =   "frmTestChameleon.frx":52F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbWin32Cool 
      Height          =   495
      Left            =   3480
      TabIndex        =   17
      ToolTipText     =   "This is a Checkbox"
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Win32 Custom"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   32768
      BCOLO           =   49152
      FCOL            =   8438015
      FCOLO           =   33023
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5312
      PICN            =   "frmTestChameleon.frx":532E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   -1  'True
   End
   Begin prjChameleon.chameleonButton cbWin16Cool 
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      ToolTipText     =   "Pressing here will show the about box"
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   1
      TX              =   "Win16 Custom"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   192
      BCOLO           =   16576
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":534A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbFlatDis 
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "Flat disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5366
      PICN            =   "frmTestChameleon.frx":5382
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbFlat 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "Flat"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5494
      PICN            =   "frmTestChameleon.frx":54B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbNetscapeDis 
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "Netscape disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":54CC
      PICN            =   "frmTestChameleon.frx":54E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbMacDis 
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Mac disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5504
      PICN            =   "frmTestChameleon.frx":5520
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbWinXPDis 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "WinXP disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5872
      PICN            =   "frmTestChameleon.frx":588E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbWin16Dis 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   1
      TX              =   "Win16 disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5B80
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbWin32Dis 
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Win32 disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5B9C
      PICN            =   "frmTestChameleon.frx":5BB8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbJavaDis 
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "Java disabled"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5BD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbNetscape 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "Netscape"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5BF0
      PICN            =   "frmTestChameleon.frx":5C0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbJava 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "&Java"
      ENAB            =   -1  'True
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5C28
      PICN            =   "frmTestChameleon.frx":5C44
      PICH            =   "frmTestChameleon.frx":5D56
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbMac 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Mac"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":5E68
      PICN            =   "frmTestChameleon.frx":5E84
      PICH            =   "frmTestChameleon.frx":62D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbWin32 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Windows 32-bit"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":6728
      PICN            =   "frmTestChameleon.frx":6744
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbWin16 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   1
      TX              =   "Windows 16-bit"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":6760
      PICN            =   "frmTestChameleon.frx":677C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbFlat2 
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "Flat Hover"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":6798
      PICN            =   "frmTestChameleon.frx":67B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbFlatDis2 
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "Hover disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":67D0
      PICN            =   "frmTestChameleon.frx":67EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbFlat2cool 
      Height          =   495
      Left            =   3480
      TabIndex        =   23
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "Flat Hover Custom"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16744576
      FCOL            =   12582912
      FCOLO           =   49344
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":6808
      PICN            =   "frmTestChameleon.frx":6824
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbOXPdis 
      Height          =   495
      Left            =   1800
      TabIndex        =   25
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "Office XP disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":6840
      PICN            =   "frmTestChameleon.frx":685C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cbOXPcool 
      Height          =   495
      Left            =   3480
      TabIndex        =   26
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "Office XP Custom"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   33023
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":6CAE
      PICN            =   "frmTestChameleon.frx":6CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton High3Ddis 
      Height          =   495
      Left            =   1800
      TabIndex        =   37
      Top             =   6480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   13
      TX              =   "Oval Button"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   128
      BCOLO           =   16576
      FCOL            =   33023
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":7344
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton TransCustom 
      Height          =   495
      Left            =   3480
      TabIndex        =   39
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   11
      TX              =   "Transp. Custom"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   65535
      FCOLO           =   65280
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":7360
      PICN            =   "frmTestChameleon.frx":737C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton Trans 
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   5880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   11
      TX              =   "Transparent"
      ENAB            =   -1  'True
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":748E
      PICN            =   "frmTestChameleon.frx":74AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton TransDis 
      Height          =   495
      Left            =   1800
      TabIndex        =   41
      Top             =   5880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   11
      TX              =   "Transparent Disabled"
      ENAB            =   0   'False
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTestChameleon.frx":75B4
      PICN            =   "frmTestChameleon.frx":75D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmTestChameleon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnType_Click()
Dim i As Long

For i = toolB.LBound To toolB.UBound
    toolB(i).ButtonType = btnType.ItemData(btnType.ListIndex)
Next
End Sub

'just to test if it works fine
Private Sub cbFlat_Click()
    MsgBox "Flat clicked!!!"
End Sub

Private Sub cbFlat2_Click()
    MsgBox "Flat Hover clicked!!!"
End Sub

Private Sub cbJava_Click()
    MsgBox "Java clicked!!!"
End Sub

Private Sub cbMac_Click()
    MsgBox "Mac clicked!!!"
End Sub

Private Sub cbNetscape_Click()
    MsgBox "Netscape clicked!!!"
End Sub

Private Sub cbWin16_Click()
    MsgBox "Win16 clicked!!!"
End Sub

Private Sub cbWin32_Click()
    MsgBox "Win32 clicked!!!"
End Sub

Private Sub cbWin32Cool_Click()
    'this was to test the double clicking to raise two consecutive single clicks
    Me.Caption = Me.Caption & "c"
End Sub

Private Sub cbWinXP_Click()
    MsgBox "WinXP clicked!!!"
End Sub

Private Sub Form_Initialize()
    If GetSetting(App.Title, "NAG", "SHOWN", 0) = 0 Then
        frmNag.Show 1
    Else
        frmAbout.Show 1
    End If
End Sub

Private Sub Form_Load()
    btnType.ListIndex = 8
    Me.Caption = Me.Caption & "  (version " & cbWinXP.Version & ")"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmTestChameleon = Nothing
End Sub

Private Sub toolB_MouseOut(Index As Integer)
If lblInfo.Caption = toolB(Index).ToolTipText Then lblInfo.Caption = ""
End Sub

Private Sub toolB_MouseOver(Index As Integer)
lblInfo.Caption = toolB(Index).ToolTipText
End Sub
