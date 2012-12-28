VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form matedit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material`s editor"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9615
   ControlBox      =   0   'False
   Icon            =   "matedit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6735
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   9380
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   8760
         TabIndex        =   66
         Top             =   1560
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   65
         Text            =   "Combo1"
         Top             =   1080
         Width           =   9135
      End
      Begin VB.Frame Frame8 
         Height          =   1455
         Left            =   90
         TabIndex        =   58
         Top             =   1800
         Width           =   9210
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   61
            Text            =   "0"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   60
            Text            =   "0"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   59
            Text            =   "0"
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Working width:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   7815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "General width:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   7815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "General height"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   960
            Width           =   7815
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1815
         Left            =   90
         TabIndex        =   49
         Top             =   3240
         Width           =   9210
         Begin VB.CommandButton Command16 
            Caption         =   "..."
            Height          =   255
            Left            =   8760
            TabIndex        =   71
            Top             =   600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   53
            Text            =   "0"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   52
            Text            =   "0"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   51
            Text            =   "0"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   50
            Text            =   "0"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Max lenght"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1320
            Width           =   7815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Min lenght"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Width           =   7815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Overlapping"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   600
            Width           =   7095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Step of wave:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   7815
         End
      End
      Begin VB.Label Label25 
         Caption         =   "Label25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1455
         Left            =   120
         TabIndex        =   72
         Top             =   5280
         Width           =   9135
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Label22"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   69
         Top             =   120
         Width           =   9255
      End
      Begin VB.Label Label13 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1560
         Width           =   8535
      End
      Begin VB.Label Label9 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   720
         Width           =   9135
      End
   End
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   180
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   9345
      Begin VB.Frame Frame12 
         Height          =   6735
         Left            =   6200
         TabIndex        =   38
         Top             =   0
         Width           =   3110
         Begin VB.ListBox List4 
            Height          =   4980
            IntegralHeight  =   0   'False
            ItemData        =   "matedit.frx":030A
            Left            =   120
            List            =   "matedit.frx":030C
            TabIndex        =   41
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Add/Edit"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Del"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame11 
         Height          =   6735
         Left            =   3240
         TabIndex        =   34
         Top             =   0
         Width           =   2880
         Begin VB.ListBox List3 
            Height          =   4995
            IntegralHeight  =   0   'False
            ItemData        =   "matedit.frx":030E
            Left            =   120
            List            =   "matedit.frx":0310
            TabIndex        =   37
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Add/Edit"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   2655
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Del"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame10 
         Height          =   6735
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   3135
         Begin VB.ListBox List2 
            Height          =   4980
            IntegralHeight  =   0   'False
            ItemData        =   "matedit.frx":0312
            Left            =   120
            List            =   "matedit.frx":0314
            TabIndex        =   33
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Add/Edit"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Del"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   2895
         End
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6735
      Left            =   180
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   9300
      Begin VB.Frame Frame13 
         Caption         =   "**"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         TabIndex        =   45
         Top             =   5640
         Width           =   9280
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00DCFBFC&
            Height          =   285
            Left            =   2235
            TabIndex        =   47
            Text            =   "10"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00DCFBFC&
            Height          =   285
            Left            =   8400
            TabIndex        =   46
            Text            =   "4"
            Top             =   270
            Width           =   735
         End
         Begin VB.Line Line22 
            BorderWidth     =   2
            X1              =   7320
            X2              =   7320
            Y1              =   720
            Y2              =   960
         End
         Begin VB.Line Line32 
            X1              =   6600
            X2              =   6480
            Y1              =   510
            Y2              =   390
         End
         Begin VB.Line Line31 
            X1              =   6600
            X2              =   6480
            Y1              =   270
            Y2              =   390
         End
         Begin VB.Line Line30 
            X1              =   6000
            X2              =   6120
            Y1              =   270
            Y2              =   390
         End
         Begin VB.Line Line29 
            X1              =   6000
            X2              =   6120
            Y1              =   510
            Y2              =   390
         End
         Begin VB.Line Line28 
            X1              =   5880
            X2              =   8400
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Line Line27 
            X1              =   6120
            X2              =   6120
            Y1              =   870
            Y2              =   270
         End
         Begin VB.Line Line26 
            X1              =   6480
            X2              =   6480
            Y1              =   750
            Y2              =   270
         End
         Begin VB.Line Line25 
            BorderWidth     =   2
            X1              =   6480
            X2              =   6120
            Y1              =   990
            Y2              =   870
         End
         Begin VB.Line Line24 
            BorderWidth     =   2
            X1              =   6480
            X2              =   6480
            Y1              =   750
            Y2              =   990
         End
         Begin VB.Line Line23 
            BorderWidth     =   2
            X1              =   8160
            X2              =   7320
            Y1              =   960
            Y2              =   720
         End
         Begin VB.Line Line21 
            BorderWidth     =   2
            X1              =   6480
            X2              =   7320
            Y1              =   750
            Y2              =   990
         End
         Begin VB.Line Line20 
            X1              =   1920
            X2              =   1800
            Y1              =   510
            Y2              =   390
         End
         Begin VB.Line Line19 
            X1              =   1800
            X2              =   1920
            Y1              =   390
            Y2              =   270
         End
         Begin VB.Line Line18 
            X1              =   1320
            X2              =   1200
            Y1              =   390
            Y2              =   270
         End
         Begin VB.Line Line17 
            X1              =   1320
            X2              =   1200
            Y1              =   390
            Y2              =   510
         End
         Begin VB.Line Line16 
            X1              =   1080
            X2              =   2280
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Line Line15 
            X1              =   1320
            X2              =   1320
            Y1              =   750
            Y2              =   270
         End
         Begin VB.Line Line14 
            X1              =   1800
            X2              =   1800
            Y1              =   870
            Y2              =   270
         End
         Begin VB.Line Line13 
            BorderWidth     =   2
            X1              =   480
            X2              =   120
            Y1              =   990
            Y2              =   870
         End
         Begin VB.Line Line12 
            BorderWidth     =   2
            X1              =   480
            X2              =   480
            Y1              =   750
            Y2              =   990
         End
         Begin VB.Line Line11 
            BorderWidth     =   2
            X1              =   1800
            X2              =   1320
            Y1              =   870
            Y2              =   750
         End
         Begin VB.Line Line10 
            BorderWidth     =   2
            X1              =   1320
            X2              =   1320
            Y1              =   750
            Y2              =   990
         End
         Begin VB.Line Line9 
            BorderWidth     =   2
            X1              =   480
            X2              =   1320
            Y1              =   750
            Y2              =   990
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4815
         Left            =   0
         TabIndex        =   15
         Top             =   490
         Width           =   3135
         Begin VB.CommandButton Command7 
            Caption         =   "Del"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   2895
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Add/Edit"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   2895
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3015
            Left            =   120
            TabIndex        =   18
            Top             =   1560
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   5318
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Length"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Count"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   19
            Top             =   4560
            Width           =   135
         End
      End
      Begin VB.Frame Frame7 
         Height          =   4815
         Left            =   6190
         TabIndex        =   9
         Top             =   490
         Width           =   3110
         Begin VB.CommandButton Command6 
            Caption         =   "Add/Edit"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Del"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   2895
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   3015
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   5318
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   13
            Top             =   4560
            Width           =   135
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4815
         Left            =   3230
         TabIndex        =   21
         Top             =   490
         Width           =   2880
         Begin VB.CommandButton Command5 
            Caption         =   "Del"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add/Edit"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   2655
         End
         Begin VB.ListBox List1 
            Height          =   2985
            ItemData        =   "matedit.frx":0316
            Left            =   120
            List            =   "matedit.frx":0318
            TabIndex        =   22
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "**"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   48
            Top             =   4560
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Label23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   70
         Top             =   120
         Width           =   6135
      End
      Begin VB.Label Label17 
         Caption         =   "* "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   5400
         Width           =   9135
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7215
      Left            =   50
      TabIndex        =   28
      Top             =   600
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   12726
      MultiRow        =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ghrthrtjrtjmrtmj"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "tjmjrtjkmkjytkytk"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "frgfgfgfrerwgfreg"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   8280
      Width           =   1695
   End
   Begin VB.OptionButton Check2 
      Caption         =   "DEL"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   2295
   End
   Begin VB.OptionButton Check1 
      Caption         =   "NEW/EDIT"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   9675
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   3120
         Width           =   8895
      End
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   7875
      Value           =   2  'Grayed
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Editor: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "matedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Check4.Enabled = True
    Frame6.Visible = False
    Frame6.ZOrder 1
    TabStrip1.Tabs(1).Selected = True
End Sub


Private Sub Check2_Click()
    Check4.value = 0
    Check4.Enabled = False
    Frame6.Visible = True
    Frame6.ZOrder 0
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Command10_Click()
    On Error Resume Next
    List2.RemoveItem IIf(List2.ListIndex = -1, 0, List2.ListIndex)
End Sub


Private Sub Command11_Click()
    On Error Resume Next
    If List2.ListIndex >= 0 Then
        List2.List(List2.ListIndex) = Val(InputBox(lng.GetResIDstring(9635), , List2.List(List2.ListIndex)))
    Else
    Dim L As String
        L = CStr(InputBox(lng.GetResIDstring(9635), , 0))
        If L > 0 And L <> "" Then List2.AddItem L
    End If

    List2.ListIndex = -1
End Sub


Private Sub Command12_Click()
    On Error Resume Next
    List3.RemoveItem IIf(List3.ListIndex = -1, 0, List3.ListIndex)
End Sub


Private Sub Command13_Click()
    On Error Resume Next
    If List3.ListIndex >= 0 Then
        List3.List(List3.ListIndex) = Val(InputBox(lng.GetResIDstring(9635), , List3.List(List3.ListIndex)))
        'List3.ListIndex = -1
    Else
    Dim L As String
        L = CStr(InputBox(lng.GetResIDstring(9635), , 0))
        If L > 0 And L <> "" Then List3.AddItem L
        'List3.ListIndex = List3.ListCount - 1
    End If

    List3.ListIndex = -1
End Sub


Private Sub Command14_Click()
    On Error Resume Next
    List4.RemoveItem IIf(List4.ListIndex = -1, 0, List4.ListIndex)
End Sub


Private Sub Command15_Click()
    On Error Resume Next
    If List4.ListIndex >= 0 Then
        List4.List(List4.ListIndex) = Val(InputBox(lng.GetResIDstring(9635), , List4.List(List4.ListIndex)))
    Else
    Dim L As String
        L = CStr(InputBox(lng.GetResIDstring(9635), , 0))
        If L > 0 And L <> "" Then List4.AddItem L
    End If

    List4.ListIndex = -1
End Sub


Private Sub Command2_Click()
On Error Resume Next
Dim Pid As Integer
If Me.Check1 And Label13.Caption <> "" And Val(Text2.Text) > 0 And Val(Text3.Text) > 0 Then


        Pid = SaveProfilData( _
        Label13, _
        ConvertData(Text2, True), _
        ConvertData(Text3, True), _
        ConvertData(Text4, True), _
        ConvertData(Text5, True), _
        ConvertData(Text6, True), _
        ConvertData(Text7, True), _
        ConvertData(text8, True), _
        ConvertData(Text9, True), _
        ConvertData(Text10, True), _
        List1.ListCount, _
        Combo1.ListIndex, _
        ChangeProfil.Label24.Tag _
        )
        
        Profil_Name = Label13

        ' Запись стандартных длин
        Dim i As Integer
            For i = 0 To List1.ListCount - 1
                SetProfilStandartLength Pid, ConvertData(List1.List(i), True)
            Next

        ' Запись невыполнимых длин
        Dim itmX As ListItem
            For i = 1 To ListView2.ListItems.Count Step 1
                Set itmX = ListView2.ListItems(i)
                SetProfilWrongLength Pid, ConvertData(itmX.Text, True), ConvertData(itmX.SubItems(1), True), IIf(itmX.Checked, 1, 0)
            Next
            
        

        ' Запись складских длин
        For i = 1 To ListView1.ListItems.Count Step 1
            Set itmX = ListView1.ListItems(i)
            SetWarehouseLength Pid, ConvertData(itmX.Text, True), itmX.SubItems(1), IIf(itmX.Checked, 1, 0)
        Next
    
        Set itmX = Nothing
        
        Execute "DELETE * FROM " & "Thickness where IDNAME=" & Pid
        Execute "DELETE * FROM " & "Coating where IDNAME=" & Pid
        Execute "DELETE * FROM " & "Color where IDNAME=" & Pid
        
        For i = 0 To List3.ListCount - 1
            If List3.List(i) <> "" Then SetProfilAdditionalData "Coating", List3.List(i), Pid, CurrentLocale
        Next
        For i = 0 To List2.ListCount - 1
            If List2.List(i) <> "" Then SetProfilAdditionalData "Thickness", List2.List(i), Pid, CurrentLocale
        Next
        For i = 0 To List4.ListCount - 1
            If List4.List(i) <> "" Then SetProfilAdditionalData "Color", List4.List(i), Pid, CurrentLocale
        Next
        
End If
    
' Удаления профиля
If Label13 <> "" And Check2.value = True Then
    Pid = GetProfilID(Label13)
    DelBaseData "select id from ProfiName where id=" & Pid
    Profil_Name = ""
End If


If Label13 <> "" Then
    ChangeProfil.ComboBox1.Text = Profil_Name
End If

ChangeProfil.lstprof_Click

If Check4.Enabled And Check4.value And Me.Check1 Then
    
    Dim n As Integer
    For n = 0 To ChangeProfil.ComboBox1.ListCount
        If ChangeProfil.ComboBox1.List(n) = Label13 Then Exit For
    Next

    ChangeProfil.ComboBox1.ListIndex = n
    
End If

Unload Me
End Sub



Private Sub Command3_Click()
    Dim n As String
        On Error Resume Next
        If Label13.Caption <> "" Then
            n = Trim(InputBox(lng.GetResIDstring(1491), lng.GetResIDstring(1100), Label13.Caption & "_copy"))
        Else
            n = Trim(InputBox(lng.GetResIDstring(1491), lng.GetResIDstring(1100), "new"))
        End If

        If n <> "" Then Label13.Caption = n
End Sub


Private Sub Command4_Click()
    On Error Resume Next
    If List1.ListIndex >= 0 Then
        List1.List(List1.ListIndex) = Val(InputBox(lng.GetResIDstring(9635), , List1.List(List1.ListIndex)))
    Else
    Dim value As Single
    Dim str As String
        str = InputBox(lng.GetResIDstring(9635), , 0)
        If IsNumeric(str) = False Then
            value = Val(str)
        Else
            value = CSng(str)
        End If
        If value = 0 Then Exit Sub
        List1.AddItem value
        List1.ListIndex = -1
    End If
End Sub


Private Sub Command5_Click()
    On Error Resume Next
    Dim Pid As Integer
    Pid = GetProfilID(Label13)
    DelBaseData "select id from ProfilsWrongLength where idname=" & Pid & " and length=" & ConvertData(List1.List(0), True)
    List1.RemoveItem IIf(List1.ListIndex = -1, 0, List1.ListIndex)
End Sub


Private Sub Command6_Click()
    Dim itmX As ListItem
    Dim value As Single
    Dim str As String
    On Error Resume Next
    
    str = InputBox(lng.GetResIDstring(9663), lng.GetResIDstring(9662), 0)
    If IsNumeric(str) = False Then
        value = Val(str)
    Else
        value = CSng(str)
    End If
    If value = 0 Then Exit Sub
    
    If ListView2.ListItems.Count > 0 Then Set itmX = ListView2.FindItem(value, 0, 1, 0)
    If itmX Is Nothing Then
        Set itmX = ListView2.ListItems.Add(, , value)
    End If

    value = 0
    str = InputBox(lng.GetResIDstring(9664), lng.GetResIDstring(9662), 0)
    If IsNumeric(str) = False Then
        value = Val(str)
    Else
        value = CSng(str)
    End If
    
    itmX.SubItems(1) = value
    itmX.Checked = True
End Sub


Private Sub Command7_Click()
    On Error Resume Next
    Dim Pid As Integer
        Pid = GetProfilID(Label13)
        DelBaseData "select id from Warehouse_profils where idname=" & Pid & " and length=" & ConvertData(ListView1.SelectedItem, True)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End Sub


Private Sub Command8_Click()
    Dim itmX As ListItem
    Dim value As Single
    Dim str As String
    On Error Resume Next
    
    str = InputBox(lng.GetResIDstring(9635), , 0)
    If IsNumeric(str) = False Then
        value = Val(str)
    Else
        value = CSng(str)
    End If
    If value = 0 Then Exit Sub
    
    If ListView1.ListItems.Count > 0 Then Set itmX = ListView1.FindItem(value, 0, 1, 0)
    If itmX Is Nothing Then
        Set itmX = ListView1.ListItems.Add(, , value)
    End If
    
    value = 0
    str = InputBox(lng.GetResIDstring(9636), , 0)
    If IsNumeric(str) = False Then
        value = Val(str)
    Else
        value = CSng(str)
    End If
    
    itmX.SubItems(1) = value
    itmX.Checked = True
End Sub


Private Sub Command9_Click()
    On Error Resume Next
    Dim Pid As Integer
        Pid = GetProfilID(Label13)
        DelBaseData "select id from ProfilsWLength where idname=" & Pid & " and length1=" & ConvertData(ListView2.SelectedItem, True)
        ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
End Sub



Private Sub Form_Load()
On Error GoTo ERR
    SetFont Me
    Dim i As Integer

        TabStrip1.Tabs(1).Caption = lng.GetResIDstring(9666)
        TabStrip1.Tabs(2).Caption = lng.GetResIDstring(9667)
        TabStrip1.Tabs(3).Caption = lng.GetResIDstring(9669)

        Me.Label3 = lng.GetResIDstring(1099)
        Me.Caption = lng.GetResIDstring(1099)
        Label1 = lng.GetResIDstring(1089)
        Label11 = lng.GetResIDstring(1093)

        Label2 = lng.GetResIDstring(1090)
        Label4 = lng.GetResIDstring(1091)
        Label10 = lng.GetResIDstring(1036)
        Label5 = lng.GetResIDstring(1055)
        Label6 = lng.GetResIDstring(1056)
        Label7 = lng.GetResIDstring(1057)
        Label8 = lng.GetResIDstring(1079)
        Label17 = lng.GetResIDstring(9665)
        Label22 = lng.GetResIDstring(9204)
        Label23 = lng.GetResIDstring(9206)

        'Check3.Caption = lng.GetResIDstring(9622)
        Label12.Caption = lng.GetResIDstring(9623)

        Label18.Caption = lng.GetResIDstring(9390)
        Label19.Caption = lng.GetResIDstring(9391)
        Label20.Caption = lng.GetResIDstring(9392)

        For i = 0 To ChangeProfil.lstprof.ListCount - 1
            Combo1.AddItem ChangeProfil.lstprof.List(i)
        Next
        
        Combo1.ListIndex = ChangeProfil.lstprof.ListIndex

        'Combo1.Text = Project.lstprof.Text

        Check1.Caption = lng.GetResIDstring(1100)
        Command8.Caption = Check1.Caption
        Command13.Caption = Check1.Caption
        Command11.Caption = Check1.Caption
        Command15.Caption = Check1.Caption
        Command4.Caption = Check1.Caption
        Command6.Caption = Check1.Caption

        Check2.Caption = lng.GetResIDstring(1101)
        Label14.Caption = Check2.Caption
        Command5.Caption = Check2.Caption
        Command9.Caption = Check2.Caption
        Command7.Caption = Check2.Caption
        Command10.Caption = Check2.Caption
        Command12.Caption = Check2.Caption
        Command14.Caption = Check2.Caption

        Check4.Caption = lng.GetResIDstring(9637)
        
        If Setup.Combo4.ListIndex = 1 Then
            Label25.Caption = lng.GetResIDstring(9702)
        Else
            Label25.Caption = lng.GetResIDstring(9703)
        End If
        
'        Label24.Caption = lng.GetResIDstring(9701) & ": " & setup.Combo4.list(setup.Combo4.ListIndex) ' Единица измерения

        If ChangeProfil.ListView1.ListItems.Count > 0 Then
    
        Dim PDataRS As Recordset
        Dim itmX As ListItem
        Dim mID As Integer
        mID = ChangeProfil.ListView1.ListItems(Setup.GetIDData(10)).ListSubItems(1).Text
        
        ' Стандартные длины
        Set PDataRS = RequestSQL("select * from ProfilsWrongLength where idname=" & mID & " order by length")
        If Not PDataRS Is Nothing Then
            Do While Not PDataRS.EOF
                matedit.List1.AddItem ConvertData(PDataRS.Fields(2))  ', PDataRS.Fields(0)
                PDataRS.MoveNext
            Loop

            PDataRS.Close
        End If
        
        ' Невыполнимые длины
        Set PDataRS = RequestSQL("select * from ProfilsWLength where idname=" & mID & " order by length1")
        If Not PDataRS Is Nothing Then
            Do While Not PDataRS.EOF
                Set itmX = ListView2.ListItems.Add(, , ConvertData(PDataRS.Fields(2)))
                itmX.SubItems(1) = ConvertData(PDataRS.Fields(3))
                itmX.Checked = CheckNullNomber(PDataRS.Fields(Setup.GetIDData(11)))
                PDataRS.MoveNext
            Loop

            PDataRS.Close
        End If
        
        ' Складские длины
        Set PDataRS = RequestSQL("select * from Warehouse_profils where idname=" & mID & " order by length")
        If Not PDataRS Is Nothing Then
            Do While Not PDataRS.EOF
                Set itmX = ListView1.ListItems.Add(, , ConvertData(PDataRS.Fields(2)))
                itmX.SubItems(1) = PDataRS.Fields(3)
                itmX.Checked = CheckNullNomber(PDataRS.Fields(4))
                PDataRS.MoveNext
            Loop

            PDataRS.Close
        End If
    
        Set PDataRS = Nothing
        Set itmX = Nothing
        
    End If

    ListView1.ColumnHeaders(1).Text = ""
    ListView1.ColumnHeaders(2).Text = ""

    For i = 0 To Project.Combo1(1).ListCount - 1
        If Project.Combo1(1).List(i) <> "" Then List2.AddItem Project.Combo1(1).List(i)
    Next

    For i = 0 To Project.Combo1(2).ListCount - 1
        If Project.Combo1(2).List(i) <> "" Then List3.AddItem Project.Combo1(2).List(i)
    Next

    For i = 0 To Project.Combo1(3).ListCount - 1
        If Project.Combo1(3).List(i) <> "" Then List4.AddItem Project.Combo1(3).List(i)
    Next
    
ERR:
End Sub


Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Index
        Case 1
            Frame3.Visible = False
            Frame9.Visible = False
            Frame1.Visible = True
        Case 2
            Frame3.Visible = True
            Frame1.Visible = False
            Frame9.Visible = False
        Case 3
            Frame9.Visible = True
            Frame3.Visible = False
            Frame1.Visible = False
    End Select

End Sub

