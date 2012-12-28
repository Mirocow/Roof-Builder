VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form SlopeSampleModule 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6960
   Icon            =   "SlopeSampleModule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6960
   StartUpPosition =   1  'CenterOwner
   Begin СhameleonButton.chameleonButton Комманда2 
      Height          =   615
      Left            =   1800
      TabIndex        =   68
      Top             =   120
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "SlopeSampleModule.frx":030A
      PICN            =   "SlopeSampleModule.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      ICONS           =   32
   End
   Begin СhameleonButton.chameleonButton Комманда3 
      Height          =   615
      Left            =   1020
      TabIndex        =   67
      Top             =   120
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "SlopeSampleModule.frx":0778
      PICN            =   "SlopeSampleModule.frx":0794
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      ICONS           =   32
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   63
      Text            =   "0"
      Top             =   120
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   1
            Left            =   2760
            TabIndex        =   6
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   5
            Text            =   "0"
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   4
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   3
            Text            =   "0"
            Top             =   3600
            Width           =   615
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   480
            X2              =   1320
            Y1              =   3360
            Y2              =   240
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   480
            X2              =   3480
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line4 
            X1              =   1320
            X2              =   1800
            Y1              =   240
            Y2              =   3360
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   1320
            X2              =   3480
            Y1              =   240
            Y2              =   3360
         End
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   45
            Left            =   600
            TabIndex        =   60
            Text            =   "0"
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   44
            Left            =   2760
            TabIndex        =   59
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   43
            Left            =   1800
            TabIndex        =   58
            Text            =   "0"
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   42
            Left            =   120
            TabIndex        =   57
            Text            =   "0"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   41
            Left            =   2280
            TabIndex        =   56
            Text            =   "0"
            Top             =   285
            Width           =   615
         End
         Begin VB.Line Line46 
            X1              =   1440
            X2              =   1440
            Y1              =   720
            Y2              =   3360
         End
         Begin VB.Line Line45 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   1440
            X2              =   3600
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line44 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   120
            X2              =   1440
            Y1              =   3360
            Y2              =   720
         End
         Begin VB.Line Line43 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3600
            X2              =   3600
            Y1              =   3360
            Y2              =   720
         End
         Begin VB.Line Line42 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   120
            X2              =   3600
            Y1              =   3360
            Y2              =   3360
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   21
            Left            =   840
            TabIndex        =   17
            Text            =   "0"
            Top             =   285
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   20
            Left            =   2880
            TabIndex        =   16
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   19
            Left            =   1440
            TabIndex        =   15
            Text            =   "0"
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   17
            Left            =   240
            TabIndex        =   14
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   16
            Left            =   2520
            TabIndex        =   13
            Text            =   "0"
            Top             =   2880
            Width           =   615
         End
         Begin VB.Line Line22 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   120
            X2              =   3600
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line21 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   120
            X2              =   120
            Y1              =   3360
            Y2              =   720
         End
         Begin VB.Line Line20 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3600
            X2              =   2280
            Y1              =   3360
            Y2              =   720
         End
         Begin VB.Line Line19 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   120
            X2              =   2280
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line18 
            X1              =   2280
            X2              =   2280
            Y1              =   720
            Y2              =   3360
         End
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   40
            Left            =   1680
            TabIndex        =   53
            Text            =   "0"
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   39
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   38
            Left            =   1680
            TabIndex        =   51
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   37
            Left            =   480
            TabIndex        =   50
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   36
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "0"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   35
            Left            =   1680
            TabIndex        =   48
            Text            =   "0"
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "0"
            Height          =   255
            Left            =   2760
            TabIndex        =   54
            Top             =   840
            Width           =   615
         End
         Begin VB.Line Line41 
            X1              =   2640
            X2              =   2640
            Y1              =   720
            Y2              =   3360
         End
         Begin VB.Line Line40 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   120
            X2              =   3600
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line39 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   1320
            X2              =   120
            Y1              =   3360
            Y2              =   720
         End
         Begin VB.Line Line38 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   2640
            X2              =   3600
            Y1              =   3360
            Y2              =   720
         End
         Begin VB.Line Line37 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   1320
            X2              =   2640
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line36 
            X1              =   1320
            X2              =   1320
            Y1              =   720
            Y2              =   3360
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   10
            Left            =   1680
            TabIndex        =   24
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   12
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   9
            Left            =   480
            TabIndex        =   22
            Text            =   "0"
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   8
            Left            =   1680
            TabIndex        =   21
            Text            =   "0"
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   13
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "0"
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   11
            Left            =   1680
            TabIndex        =   19
            Text            =   "0"
            Top             =   285
            Width           =   615
         End
         Begin VB.Line Line9 
            X1              =   1320
            X2              =   1320
            Y1              =   720
            Y2              =   3360
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   1320
            X2              =   2640
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3600
            X2              =   2640
            Y1              =   3360
            Y2              =   720
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   120
            X2              =   1320
            Y1              =   3360
            Y2              =   720
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   120
            X2              =   3600
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line12 
            X1              =   2640
            X2              =   2640
            Y1              =   720
            Y2              =   3360
         End
         Begin VB.Label Label3 
            Caption         =   "0"
            Height          =   255
            Left            =   2760
            TabIndex        =   25
            Top             =   3000
            Width           =   615
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   26
            Left            =   1560
            TabIndex        =   36
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   25
            Left            =   2160
            TabIndex        =   35
            Text            =   "0"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   24
            Left            =   600
            TabIndex        =   34
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.Line Line28 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3480
            X2              =   480
            Y1              =   360
            Y2              =   3360
         End
         Begin VB.Line Line27 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   480
            X2              =   3480
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line26 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   480
            X2              =   480
            Y1              =   3360
            Y2              =   360
         End
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   29
            Left            =   2760
            TabIndex        =   40
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   28
            Left            =   1320
            TabIndex        =   39
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   27
            Left            =   1560
            TabIndex        =   38
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.Line Line31 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3480
            X2              =   3480
            Y1              =   3360
            Y2              =   360
         End
         Begin VB.Line Line30 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   480
            X2              =   3480
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line29 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   480
            X2              =   3480
            Y1              =   360
            Y2              =   3360
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   23
            Left            =   1440
            TabIndex        =   32
            Text            =   "0"
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   22
            Left            =   1320
            TabIndex        =   31
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   18
            Left            =   2640
            TabIndex        =   30
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.Line Line25 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3480
            X2              =   480
            Y1              =   240
            Y2              =   3360
         End
         Begin VB.Line Line24 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   480
            X2              =   3480
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line23 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3480
            X2              =   3480
            Y1              =   3360
            Y2              =   240
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   5
            Left            =   600
            TabIndex        =   11
            Text            =   "0"
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   7
            Left            =   2160
            TabIndex        =   10
            Text            =   "0"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   6
            Left            =   1440
            TabIndex        =   9
            Text            =   "0"
            Top             =   3480
            Width           =   615
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   480
            X2              =   480
            Y1              =   3360
            Y2              =   240
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   480
            X2              =   3480
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line13 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   480
            X2              =   3480
            Y1              =   240
            Y2              =   3360
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   34
            Left            =   1440
            TabIndex        =   46
            Text            =   "0"
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   33
            Left            =   1320
            TabIndex        =   45
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   32
            Left            =   720
            TabIndex        =   44
            Text            =   "0"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   31
            Left            =   2760
            TabIndex        =   43
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   30
            Left            =   240
            TabIndex        =   42
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.Line Line35 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3360
            X2              =   1800
            Y1              =   480
            Y2              =   3600
         End
         Begin VB.Line Line34 
            X1              =   1440
            X2              =   1800
            Y1              =   480
            Y2              =   3600
         End
         Begin VB.Line Line33 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   360
            X2              =   3360
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line32 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   1800
            X2              =   360
            Y1              =   3600
            Y2              =   480
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   3975
         Begin VB.TextBox value 
            Height          =   285
            Index           =   14
            Left            =   1800
            TabIndex        =   28
            Text            =   "0"
            Top             =   3120
            Width           =   615
         End
         Begin VB.TextBox value 
            Height          =   285
            Index           =   15
            Left            =   120
            TabIndex        =   27
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.Line Line14 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   840
            X2              =   840
            Y1              =   1200
            Y2              =   3000
         End
         Begin VB.Line Line15 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   840
            X2              =   3360
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Line Line16 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            X1              =   3360
            X2              =   3360
            Y1              =   3000
            Y2              =   1200
         End
         Begin VB.Line Line17 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            X1              =   840
            X2              =   3360
            Y1              =   1200
            Y2              =   1200
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   62
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "S: ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   61
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
      Begin СhameleonButton.chameleonButton Command3 
         Height          =   495
         Left            =   120
         TabIndex        =   66
         Top             =   3360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BTYPE           =   7
         TX              =   "&Ok"
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
         MICON           =   "SlopeSampleModule.frx":0BE6
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
      Begin СhameleonButton.chameleonButton Command2 
         Height          =   495
         Left            =   120
         TabIndex        =   65
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BTYPE           =   7
         TX              =   "&Calc"
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
         MICON           =   "SlopeSampleModule.frx":0C02
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
      Begin СhameleonButton.chameleonButton Command4 
         Height          =   495
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BTYPE           =   7
         TX              =   "&Clear"
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
         MICON           =   "SlopeSampleModule.frx":0C1E
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
Attribute VB_Name = "SlopeSampleModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub calc()
    Dim i As Integer
    Dim a2 As Single
    Dim sbox As Single

    value(8).ForeColor = vbBlack: value(11).ForeColor = vbBlack

    For i = 0 To value.Count - 1
        value(i).Text = ConvertData(value(i).Text, True)
        value(i).ForeColor = vbBlack
    Next

    On Error GoTo ERR

Select Case SlopeSampleModule.Text2
Case 0
        ' a = value(2), b = value(0), c = value(1)
        ' a1 = value(4), h = value(3)

        '      ___________________________
        'S = \\ p * ( p - a) * ( p - b ) * ( p - c )
        '
        '    Здесь p - это полупериметр, который вычисляется по формуле: p = ( a + b + c ) / 2.
        '
        If value(2) > 0 And value(0) > 0 And value(1) > 0 Then
            If ((CSng(value(2)) + CSng(value(0))) >= value(1) And (CSng(value(2)) + CSng(value(1))) >= value(0) And (CSng(value(0)) + CSng(value(1))) >= value(2)) Then
                'a, b, c - длины сторон заданного треугольника.
            
        Dim P As Single
            P = (CSng(value(2)) + CSng(value(0)) + CSng(value(1))) / 2
            Label2 = Round(Sqr(P * (P - value(2)) * (P - value(0)) * (P - value(1))), 2)
                
            ' S = 1/2*ah
            ' h = S/(a/2)
            '            value(3).Enabled = False
            value(3) = Label2 / (value(2) / 2)
            '            value(4).Enabled = False
            value(4) = Sqr(value(0) * value(0) - value(3) * value(3))
            '            Command2.Enabled = False
                
            Else: Label2 = 0: MsgBox "Ошибка: Одна из сторон превышает сумму противоположных сторон", vbCritical
        End If
            
    ElseIf value(0) > 0 And value(1) > 0 And value(3) > 0 Then
        ' b,c,h
        
        Label2 = Round((value(1) / 2) * value(3), 2)
        value(4) = Round(Sqr(value(0) * value(0) - value(3) * value(3)), 2)
        value(4).Enabled = False
        value(2) = (Label2 / value(3)) * 2
        '    value(2).Enabled = False
        '    Command2.Enabled = False
        
    ElseIf value(0) > 0 And value(2) > 0 And value(3) Then
        ' b,a,h
        
        Label2 = Round((value(2) / 2) * value(3), 2)
        value(4) = Round(Sqr(value(0) * value(0) - value(3) * value(3)), 2)
        '    value(4).Enabled = False
        value(1) = Sqr(((value(0) - value(4)) * (value(0) - value(4))) + CSng(value(3)) * value(3))
        '    value(1).Enabled = False
        '    Command2.Enabled = False
        
    ElseIf value(2) > 0 And value(1) > 0 And value(3) Then
        ' a,c,h
        
        Label2 = Round((value(1) / 2) * value(3), 2)
        a2 = Round(Sqr(value(1) * value(1) - value(3) * value(3)), 2)
        value(4) = value(2) - a2
        '    value(4).Enabled = False
        value(0) = (Label2 / value(3)) * 2
        '    value(0).Enabled = False
        '    Command2.Enabled = False
        
    ElseIf value(4) > 0 And value(1) > 0 And value(2) Then
        ' a1, c, a
        
        a2 = value(2) - value(4)
        value(3) = Sqr(value(1) * value(1) - a2 * a2)
        '    value(3).Enabled = False
        value(0) = Sqr(value(4) * value(4) + CSng(value(3)) * value(3))
        '    value(0).Enabled = False
        Label2 = Round((value(1) / 2) * value(3), 2)
        
        '    Command2.Enabled = False
        
    ElseIf value(4) > 0 And value(1) > 0 And value(0) Then
        ' a1, c, b
        
        value(3) = Sqr(value(0) * value(0) - value(4) * value(4))
        '    value(3).Enabled = False
        value(2) = Sqr(value(1) * value(1) - value(3) * value(3)) + CSng(value(4))
        '    value(2).Enabled = False
        Label2 = Round((value(1) / 2) * value(3), 2)
        
        '    Command2.Enabled = False
        
    ElseIf value(2) > 0 And value(4) > 0 And value(3) Then
        'a, a1, h
        
        value(0) = Sqr(value(4) * value(4) + value(3) * value(3))
        '    value(0).Enabled = False
        a2 = value(2) - value(4)
        value(1) = Sqr(a2 * a2 + value(3) * value(3))
        '    value(1).Enabled = False
        Label2 = Round((value(1) / 2) * value(3), 2)
        
    End If
    
Case 1
    If value(6) > 0 And value(7) > 0 Then
        value(5) = Sqr(value(7) * value(7) - value(6) * value(6))
        If value(5) > 0 Then Label2 = Round((value(5) * value(6)) / 2, 2)
    ElseIf value(5) > 0 And value(6) > 0 Then
        value(7) = Sqr(value(5) * value(5) + value(6) * value(6))
        If value(7) > 0 Then Label2 = Round((value(5) * value(6)) / 2, 2)
    ElseIf value(7) > 0 And value(5) > 0 Then
        value(6) = Sqr(value(7) * value(7) - value(5) * value(5))
        If value(6) > 0 Then Label2 = Round((value(5) * value(6)) / 2, 2)
    End If
    
Case 2
    
    If mCint(value(8)) <= mCint(value(11)) Then value(8).ForeColor = vbRed: value(11).ForeColor = vbRed: Exit Sub
    
    If value(12) > 0 And value(9) > 0 Then
        value(10) = Sqr(value(12) * value(12) - value(9) * value(9))
        '        value(10).Enabled = False
    End If
    
    If value(10) > 0 And value(9) > 0 Then
        value(12) = Sqr(value(10) * CSng(value(10)) + CSng(value(9)) * value(9))
        '        value(12).Enabled = False
    End If
    
    If value(11) > 0 And value(9) > 0 And value(8) > 0 Then
        a2 = value(8) - (CSng(value(9)) + CSng(value(11)))
        Label3 = a2
    End If
    
    If a2 > 0 And (value(10) > 0 Or value(13) > 0) Then
        If value(13) > 0 Then value(10) = Sqr(value(13) * value(13) - a2 * a2) 'value(13).Enabled = False
        If value(10) > 0 Then value(13) = Sqr(value(10) * value(10) + a2 * a2) ': value(10).Enabled = False
        Label2 = Round(0.5 * (CSng(value(8)) + CSng(value(11))) * value(10), 2)
        '        Exit Sub
    End If
    
Case 3
    If value(14) > 0 And value(15) > 0 Then
        Label2 = Round(value(14) * value(15), 2)
    End If

Case 4
    If value(17) > 0 And value(19) > 0 And value(16) > 0 Then
        value(21) = value(19) - value(16)
        sbox = value(21) * value(17)
        value(20) = Sqr(value(16) * value(16) + value(17) * value(17))
    ElseIf value(17) > 0 And value(19) > 0 And value(21) > 0 Then
        value(16) = value(19) - value(21)
        sbox = value(21) * value(17)
        value(20) = Sqr(value(16) * value(16) + value(17) * value(17))
    ElseIf value(20) > 0 And value(19) > 0 And value(21) > 0 Then
        value(16) = value(19) - value(21)
        value(17) = Sqr(value(20) * value(20) - value(16) * value(16))
        sbox = value(21) * value(17)
    ElseIf value(19) > 0 And value(16) > 0 And value(20) > 0 Then
        value(17) = Sqr(value(20) * value(20) - value(16) * value(16))
        value(21) = value(19) - value(16)
        sbox = value(21) * value(17)
    End If
    
    Label2 = Round(sbox + (value(16) * value(17)) / 2, 2)

Case 6
    If value(22) > 0 And value(23) > 0 Then
        value(18) = Sqr(value(22) * value(22) - value(23) * value(23))
        If value(18) > 0 Then Label2 = Round((value(18) * value(23)) / 2, 2)
    ElseIf value(18) > 0 And value(23) > 0 Then
        value(22) = Sqr(value(18) * value(18) + value(23) * value(23))
        If value(22) > 0 Then Label2 = Round((value(18) * value(23)) / 2, 2)
    ElseIf value(18) > 0 And value(22) > 0 Then
        value(23) = Sqr(value(22) * value(22) - value(18) * value(18))
        If value(23) > 0 Then Label2 = Round((value(18) * value(23)) / 2, 2)
    End If
    
Case 7
    ' a = value(2), b = value(0), c = value(1)
    ' a1 = value(4), h = value(3)
    
    '      ___________________________
    'S = \\ p * ( p - a) * ( p - b ) * ( p - c )
    '
    '    Здесь p - это полупериметр, который вычисляется по формуле: p = ( a + b + c ) / 2.
    '
    If value(34) > 0 And value(30) > 0 And value(31) > 0 Then
        If ((value(34) + value(30)) >= value(31) And (value(34) + value(31)) >= value(30) And (value(30) + value(31)) >= value(34)) Then
            'a, b, c - длины сторон заданного треугольника.
        
            '            Dim P As Single
            P = (CSng(value(34)) + CSng(value(30)) + CSng(value(31))) / 2
            Label2 = Round(Sqr(P * (P - value(34)) * (P - value(30)) * (P - value(31))), 2)
            
            ' S = 1/2*ah
            ' h = S/(a/2)
            '            value(33).Enabled = False
            value(33) = Label2 / (value(34) / 2)
            '            value(32).Enabled = False
            value(32) = Sqr(value(30) * value(30) - value(33) * value(33))
            '            Command2.Enabled = False
            
            Else: Label2 = 0: MsgBox "Ошибка: Одна из сторон превышает сумму противоположных сторон", vbCritical
        End If
        
    ElseIf value(30) > 0 And value(31) > 0 And value(33) > 0 Then
        ' b,c,h
    
        Label2 = Round((value(31) / 2) * value(33), 2)
        value(32) = Round(Sqr(value(30) * value(30) - value(33) * value(33)), 2)
        '    value(32).Enabled = False
        value(34) = (Label2 / value(33)) * 2
        '    value(34).Enabled = False
        '    Command2.Enabled = False
    
    ElseIf value(30) > 0 And value(34) > 0 And value(33) Then
        ' b,a,h
    
        Label2 = Round((value(34) / 2) * value(33), 2)
        value(32) = Round(Sqr(value(30) * value(30) - value(33) * value(33)), 2)
        '    value(32).Enabled = False
        value(31) = Sqr(((value(30) - value(32)) * (value(30) - value(32))) + value(33) * value(33))
        '    value(31).Enabled = False
        '    Command2.Enabled = False
    
    ElseIf value(34) > 0 And value(31) > 0 And value(33) Then
        ' a,c,h
    
        Label2 = Round((value(31) / 2) * value(33), 2)
        a2 = Round(Sqr(value(31) * value(31) - value(33) * value(33)), 2)
        value(32) = value(34) - a2
        '    value(32).Enabled = False
        value(30) = (Label2 / value(33)) * 2
        '    value(30).Enabled = False
        '    Command2.Enabled = False
    
    ElseIf value(32) > 0 And value(31) > 0 And value(34) Then
        ' a1, c, a
    
        a2 = value(34) - value(32)
        value(33) = Sqr(value(31) * value(31) - a2 * a2)
        '    value(33).Enabled = False
        value(30) = Sqr(value(32) * value(32) + value(33) * value(33))
        '    value(30).Enabled = False
        Label2 = Round((value(31) / 2) * value(33), 2)
    
        '    Command2.Enabled = False
    
    ElseIf value(32) > 0 And value(31) > 0 And value(30) Then
        ' a1, c, b
    
        value(33) = Sqr(value(30) * value(30) - value(32) * value(32))
        '    value(33).Enabled = False
        value(34) = Sqr(value(31) * value(31) - value(33) * value(33)) + value(32)
        '    value(34).Enabled = False
        Label2 = Round((value(31) / 2) * value(33), 2)
    
        '    Command2.Enabled = False
    
    ElseIf value(34) > 0 And value(32) > 0 And value(33) Then
        'a, a1, h
    
        value(30) = Sqr(value(32) * value(32) + value(33) * value(33))
        '    value(30).Enabled = False
        a2 = value(34) - value(32)
        value(31) = Sqr(a2 * a2 + value(33) * value(33))
        '    value(31).Enabled = False
        Label2 = Round((value(31) / 2) * value(33), 2)
    
        '    Command2.Enabled = False
    End If

Case 5
    If value(27) > 0 And value(28) > 0 Then
        value(29) = Sqr(value(28) * value(28) - value(27) * value(27))
        If value(29) > 0 Then Label2 = Round((value(29) * value(27)) / 2, 2)
    ElseIf value(28) > 0 And value(29) > 0 Then
        value(27) = Sqr(value(28) * value(28) - value(29) * value(29))
        If value(27) > 0 Then Label2 = Round((value(29) * value(27)) / 2, 2)
    ElseIf value(29) > 0 And value(27) > 0 Then
        value(28) = Sqr(value(29) * value(29) + value(27) * value(27))
        If value(28) > 0 Then Label2 = Round((value(29) * value(27)) / 2, 2)
    End If

Case 8
    If value(24) > 0 And value(25) > 0 Then
        value(26) = Sqr(value(25) * value(25) - value(24) * value(24))
        If value(26) > 0 Then Label2 = Round((value(24) * value(26)) / 2, 2)
    ElseIf value(25) > 0 And value(26) > 0 Then
        value(24) = Sqr(value(25) * value(25) - value(26) * value(26))
        If value(24) > 0 Then Label2 = Round((value(24) * value(26)) / 2, 2)
    ElseIf value(24) > 0 And value(26) > 0 Then
        value(25) = Sqr(value(24) * value(24) + value(26) * value(26))
        Label2 = Round((value(24) * value(26)) / 2, 2)
    End If

Case 9
    If mCint(value(38)) <= mCint(value(40)) Then value(38).ForeColor = vbRed: value(40).ForeColor = vbRed: Exit Sub
    
    If value(36) > 0 And value(37) > 0 Then
        value(35) = Sqr(value(36) * value(36) - value(37) * value(37))
        '        value(35).Enabled = False
    End If
    
    If value(35) > 0 And value(37) > 0 Then
        value(36) = Sqr(value(35) * value(35) + value(37) * value(37))
        '        value(36).Enabled = False
    End If
    
    If value(40) > 0 And value(37) > 0 And value(38) > 0 Then
        a2 = value(38) - (CSng(value(37)) + CSng(value(40)))
        Label3 = a2
    End If
    
    If a2 > 0 And (value(35) > 0 Or value(39) > 0) Then
        If value(39) > 0 Then value(35) = Sqr(value(39) * value(39) - a2 * a2) ': value(39).Enabled = False
        If value(35) > 0 Then value(39) = Sqr(value(35) * value(35) + a2 * a2) ': value(35).Enabled = False
        Label2 = Round(0.5 * (CSng(value(38)) + CSng(value(40))) * value(35), 2)
        '        Exit Sub
    End If
    
Case 10
    If value(43) > 0 And value(44) > 0 And value(45) > 0 Then
        value(41) = value(43) - value(45)
        sbox = value(41) * value(44)
        value(42) = Sqr(value(45) * value(45) + value(44) * value(44))
    ElseIf value(43) > 0 And value(44) > 0 And value(41) > 0 Then
        value(45) = value(43) - value(44)
        sbox = value(41) * value(44)
        value(42) = Sqr(value(45) * value(45) + value(44) * value(44))
    ElseIf value(42) > 0 And value(43) > 0 And value(45) > 0 Then
        value(41) = value(43) - value(45)
        value(44) = Sqr(value(42) * value(42) - value(45) * value(45))
        sbox = value(44) * value(41)
        '    ElseIf value(44) > 0 And value(45) > 0 And value(43) > 0 Then
        '        value(43) = Sqr(value(42) * value(42) - value(45) * value(45))
        '        value(41) = value(44) - value(45)
        '        sbox = value(41) * value(43)
    End If
    
    Label2 = Round(sbox + (value(45) * value(43)) / 2, 2)
    
End Select

Label2 = Format(Label2 / 10000, "# ##0.00") & " " & m2

Exit Sub
ERR:
'MsgBox "ERR: " & ERR.Description, vbCritical, "Err"
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.15." & ERR.Source & "]", ERR.Number, ERR.Description
Command4.value = True
End Sub


Private Sub Command2_Click()
    calc
    SetChange True
End Sub


Private Sub Command3_Click()
If IsLoadForm("Lapepic") Then

    If Label2 = 0 Or SlP(N_Slope).CountOfLines > 0 Or SlP(N_Slope).CountOfPoints > 0 Then
        
        Command2.value = True
        If Label2 <= 0 Then Unload Me: Exit Sub
    
    End If

    Lape_Lines(N_Slope, 1, 0) = 1
    Lape_Lines(N_Slope, 1, 1) = 2
    Lape_Points_X(N_Slope, 1) = 0
    Lape_Points_Y(N_Slope, 1) = 0

    Select Case SlopeSampleModule.Text2
        Case 0
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(2)
            Lape_Points_Y(N_Slope, 2) = 0
    
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 1
            Lape_Points_X(N_Slope, 3) = value(4)
            Lape_Points_Y(N_Slope, 3) = value(3)
    
            SlP(N_Slope).CountOfLines = 3
            SlP(N_Slope).CountOfPoints = 3
    
        Case 1
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(6)
            Lape_Points_Y(N_Slope, 2) = 0
    
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 1
            Lape_Points_X(N_Slope, 3) = 0
            Lape_Points_Y(N_Slope, 3) = value(5)
  
            SlP(N_Slope).CountOfLines = 3
            SlP(N_Slope).CountOfPoints = 3
  
        Case 2
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(8)
            Lape_Points_Y(N_Slope, 2) = 0
    
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 4
            Lape_Points_X(N_Slope, 3) = CSng(value(9)) + CSng(value(11))
            Lape_Points_Y(N_Slope, 3) = value(10)
    
            Lape_Lines(N_Slope, 4, 0) = 4
            Lape_Lines(N_Slope, 4, 1) = 1
            Lape_Points_X(N_Slope, 4) = value(9)
            Lape_Points_Y(N_Slope, 4) = value(10)
  
            SlP(N_Slope).CountOfLines = 4
            SlP(N_Slope).CountOfPoints = 4

        Case 3
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(14)
            Lape_Points_Y(N_Slope, 2) = 0
  
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 4
            Lape_Points_X(N_Slope, 3) = value(14)
            Lape_Points_Y(N_Slope, 3) = value(15)
  
            Lape_Lines(N_Slope, 4, 0) = 4
            Lape_Lines(N_Slope, 4, 1) = 1
            Lape_Points_X(N_Slope, 4) = 0
            Lape_Points_Y(N_Slope, 4) = value(15)
  
            SlP(N_Slope).CountOfLines = 4
            SlP(N_Slope).CountOfPoints = 4
  
        Case 4
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(19)
            Lape_Points_Y(N_Slope, 2) = 0
  
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 4
            Lape_Points_X(N_Slope, 3) = value(21)
            Lape_Points_Y(N_Slope, 3) = value(17)
  
            Lape_Lines(N_Slope, 4, 0) = 4
            Lape_Lines(N_Slope, 4, 1) = 1
            Lape_Points_X(N_Slope, 4) = 0
            Lape_Points_Y(N_Slope, 4) = value(17)
  
            SlP(N_Slope).CountOfLines = 4
            SlP(N_Slope).CountOfPoints = 4

        Case 5
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(27)
            Lape_Points_Y(N_Slope, 2) = 0
    
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 1
            Lape_Points_X(N_Slope, 3) = value(27)
            Lape_Points_Y(N_Slope, 3) = -value(29)
  
            SlP(N_Slope).CountOfLines = 3
            SlP(N_Slope).CountOfPoints = 3

        Case 6
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(23)
            Lape_Points_Y(N_Slope, 2) = 0
    
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 1
            Lape_Points_X(N_Slope, 3) = value(23)
            Lape_Points_Y(N_Slope, 3) = value(18)
  
            SlP(N_Slope).CountOfLines = 3
            SlP(N_Slope).CountOfPoints = 3

        Case 7
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(34)
            Lape_Points_Y(N_Slope, 2) = 0
    
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 1
            Lape_Points_X(N_Slope, 3) = value(32)
            Lape_Points_Y(N_Slope, 3) = -value(33)
    
            SlP(N_Slope).CountOfLines = 3
            SlP(N_Slope).CountOfPoints = 3
    
        Case 8
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(26)
            Lape_Points_Y(N_Slope, 2) = 0
    
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 1
            Lape_Points_X(N_Slope, 3) = 0
            Lape_Points_Y(N_Slope, 3) = -value(24)
  
            SlP(N_Slope).CountOfLines = 3
            SlP(N_Slope).CountOfPoints = 3
    
        Case 9
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(38)
            Lape_Points_Y(N_Slope, 2) = 0
    
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 4
            Lape_Points_X(N_Slope, 3) = CSng(value(37)) + CSng(value(40))
            Lape_Points_Y(N_Slope, 3) = -value(35)
    
            Lape_Lines(N_Slope, 4, 0) = 4
            Lape_Lines(N_Slope, 4, 1) = 1
            Lape_Points_X(N_Slope, 4) = value(37)
            Lape_Points_Y(N_Slope, 4) = -value(35)
  
            SlP(N_Slope).CountOfLines = 4
            SlP(N_Slope).CountOfPoints = 4

        Case 10
            Lape_Lines(N_Slope, 2, 0) = 2
            Lape_Lines(N_Slope, 2, 1) = 3
            Lape_Points_X(N_Slope, 2) = value(43)
            Lape_Points_Y(N_Slope, 2) = 0
  
            Lape_Lines(N_Slope, 3, 0) = 3
            Lape_Lines(N_Slope, 3, 1) = 4
            Lape_Points_X(N_Slope, 3) = value(43)
            Lape_Points_Y(N_Slope, 3) = value(44)
  
            Lape_Lines(N_Slope, 4, 0) = 4
            Lape_Lines(N_Slope, 4, 1) = 1
            Lape_Points_X(N_Slope, 4) = value(45)
            Lape_Points_Y(N_Slope, 4) = value(44)
  
            SlP(N_Slope).CountOfLines = 4
            SlP(N_Slope).CountOfPoints = 4
    End Select

    Lapepic.Command5.value = True
    
End If
Unload Me
End Sub


Private Sub Command4_Click()
    Dim i As Integer

        For i = 0 To value.Count - 1
            value(i).Enabled = True
            value(i).Text = 0
            '    value(i).ForeColor = vblack
        Next

        Label2 = 0
        Label3 = 0

        Command2.Enabled = True
End Sub


Private Sub Form_Load()
    Me.Caption = lng.GetResIDstring(9564)
    Command2.Caption = lng.GetResIDstring(9559)
    If IsLoadForm("Lapepic") Then
        If SlP(N_Slope).CountOfLines > 0 And SlP(N_Slope).CountOfPoints > 0 Then Command3.Enabled = False
    End If
End Sub


Private Sub Text2_Change()
    Select Case Text2
        Case 0
            SlopeSampleModule.Frame4.ZOrder 0
        Case 1 ', 6, 5, 8
            SlopeSampleModule.Frame5.ZOrder 0
        Case 2
            SlopeSampleModule.Frame6.ZOrder 0
        Case 3
            SlopeSampleModule.Frame7.ZOrder 0
        Case 4
            SlopeSampleModule.Frame8.ZOrder 0
        Case 5
            SlopeSampleModule.Frame10.ZOrder 0
        Case 6
            SlopeSampleModule.Frame3.ZOrder 0
        Case 7
            SlopeSampleModule.Frame11.ZOrder 0
        Case 8
            SlopeSampleModule.Frame9.ZOrder 0
        Case 9
            SlopeSampleModule.Frame12.ZOrder 0
        Case 10
            SlopeSampleModule.Frame13.ZOrder 0
    End Select
End Sub


Private Sub Комманда2_Click()
If Text2 < 10 Then
Text2.Text = Text2.Text + 1
Else
Text2.Text = 0
End If
End Sub


Private Sub Комманда3_Click()
If Text2 > 0 Then
Text2.Text = Text2.Text - 1
Else
Text2.Text = 10
End If
End Sub
