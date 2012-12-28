VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm OfficeStart 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00C0C0C0&
   Caption         =   "Roof Builder"
   ClientHeight    =   9840
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   9000
   Icon            =   "OfficeStart_.frx":0000
   LinkTopic       =   "MDIForm1"
   MousePointer    =   1  'Arrow
   OLEDropMode     =   1  'Manual
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9000
      TabIndex        =   5
      Top             =   9240
      Width           =   9000
      Begin VB.ListBox txtLog 
         CausesValidation=   0   'False
         Height          =   255
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   9015
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3720
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":0474
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":06DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":0871
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":0C83
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":0D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":0E8C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9000
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   9000
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   661
         ShowTips        =   0   'False
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Свойства проекта"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "План проекта"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Менеджер проекта"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Окно расчета"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Предпросмотр"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8940
      TabIndex        =   1
      ToolTipText     =   "55bfbf92e85cfce8500b9dbb62842c1c"
      Top             =   1200
      Width           =   9000
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DEMO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -360
         MouseIcon       =   "OfficeStart_.frx":110D
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   0
         Width           =   10935
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1680
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":125F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":12FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":1383
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   600
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2520
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":147F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":17B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":1D55
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":22F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":2738
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":2B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":313C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":372A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OfficeStart_.frx":39C5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   9495
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2647
            MinWidth        =   2647
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6376
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3705
            MinWidth        =   3705
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   407
            MinWidth        =   407
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   825
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   1455
      ButtonWidth     =   1323
      ButtonHeight    =   1402
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "rfd"
                  Text            =   "New project rfd"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "rbp"
                  Text            =   "New project rbp"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Renew"
            Object.ToolTipText     =   "Renew"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Undo"
            Description     =   "Ctrl + Z"
            Object.ToolTipText     =   "Ctrl + Z"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Redo"
            Description     =   "Ctrl + Y"
            Object.ToolTipText     =   "Ctrl + Y"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Man"
            Key             =   "man"
            Object.ToolTipText     =   "Manager of project"
            ImageIndex      =   6
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Setup"
            Object.ToolTipText     =   "Setup"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Preview"
            Object.ToolTipText     =   "Preview"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quit"
            Object.ToolTipText     =   "Quit"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu_fil 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu menu_uusi 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu menu_open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu menu_close 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu recfiles 
         Caption         =   "Recently Files"
         Begin VB.Menu menrfl_file 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_upd 
         Caption         =   "Update"
         Enabled         =   0   'False
         Shortcut        =   ^U
         Visible         =   0   'False
      End
      Begin VB.Menu menu_save 
         Caption         =   "Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu menu_save_as 
         Caption         =   "Save as..."
         Enabled         =   0   'False
      End
      Begin VB.Menu men3 
         Caption         =   "-"
      End
      Begin VB.Menu da 
         Caption         =   "View"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_print_valinta 
         Caption         =   "Print"
         Enabled         =   0   'False
         Begin VB.Menu menu_print 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu menu_print 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu menu_print 
            Caption         =   ""
            Index           =   2
            Shortcut        =   ^L
         End
         Begin VB.Menu menu_print 
            Caption         =   ""
            Index           =   3
         End
      End
      Begin VB.Menu menu_xls 
         Caption         =   "Send to XLS"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu send_sum_xls 
            Caption         =   "Send Summary"
         End
      End
      Begin VB.Menu men2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_setup 
         Caption         =   "Setup"
         Shortcut        =   ^P
      End
      Begin VB.Menu menu_end 
         Caption         =   "End"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menuedit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu menuUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu menuRedo 
         Caption         =   "Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu menuAddins 
      Caption         =   "Add-Ins"
      Begin VB.Menu menuprofmanager 
         Caption         =   "Profile manager"
      End
      Begin VB.Menu menupman 
         Caption         =   "Projects manager"
         Shortcut        =   ^F
      End
      Begin VB.Menu menudump 
         Caption         =   "Dump"
      End
      Begin VB.Menu menuslope 
         Caption         =   "FlatModels"
      End
      Begin VB.Menu mOpWp 
         Caption         =   "Operation with project"
         Visible         =   0   'False
         Begin VB.Menu mExportShow 
            Caption         =   "Export"
            Shortcut        =   ^E
         End
      End
   End
   Begin VB.Menu menu_view 
      Caption         =   "Option"
      Visible         =   0   'False
      Begin VB.Menu menu_view_m 
         Caption         =   "Рисование прямых"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Показывать размеры длин сторон"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Показывать разметку листов"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Закрашивать область листов"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Показывать обозначение листов"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Показывать модули листа"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Показать высоту листа"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Меню корректировки листа"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Меню перемещения"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu menu_view_m 
         Caption         =   "Менеджер проекта"
         Index           =   11
      End
   End
   Begin VB.Menu ab 
      Caption         =   "Help"
      Begin VB.Menu hot_keys 
         Caption         =   "Hot Kes"
      End
      Begin VB.Menu menu_help 
         Caption         =   "Video lessons"
      End
      Begin VB.Menu men4 
         Caption         =   "-"
      End
      Begin VB.Menu menuactrb 
         Caption         =   "Active Roof Builder"
      End
      Begin VB.Menu menu_lic 
         Caption         =   "License agreement..."
      End
      Begin VB.Menu men10 
         Caption         =   "-"
      End
      Begin VB.Menu m_chlog 
         Caption         =   "Change log"
      End
      Begin VB.Menu chknewversion 
         Caption         =   "Check new version"
      End
      Begin VB.Menu feedback 
         Caption         =   "Feedback"
         Begin VB.Menu bugreport 
            Caption         =   "Bug Report"
         End
      End
      Begin VB.Menu rbhomepage 
         Caption         =   "Roof Builder Home Page"
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu menuabout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu ToolBarSettings 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu ToolbarCustomize 
         Caption         =   "&Customize"
      End
   End
End
Attribute VB_Name = "OfficeStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Механизм хранения шага назад
'
Private fseek As Long ' Для отладки

Private INIT As Boolean
Private startp As Date

'Private Declare Function GetFreeResources Lib "RSRC32.dll" Alias "_MyGetFreeSystemResources32@4" (ByVal lWhat As Long) As Long
     
'Вот дабавил функцию для сокращённого представленя имени файла
Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hdc As Long, ByVal lpszPath As String, ByVal dx As Long) As Long
Private lhDC As Long, lCtlWidth As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private TimeStart As Date

Public HistoryWorking As Boolean

Private Sub bugreport_Click()
    SendTo Me, "support@roof-builder.ru", "Bugs report", "The following errors were found:"
End Sub


Private Sub chknewversion_Click()
    Navigate Me, "http://roof-builder.ru/rbcurrent.php?ver=" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub da_Click()
    Me.TabStrip1.Tabs(5).Selected = True
End Sub

Private Sub hot_keys_Click()
    Navigate Me, "http://roof-builder.ru/engine/zaregistrirovannye-goryachie-klavishi"
End Sub

Private Sub Label1_Click()
    Navigate Me, "http://roof-builder.ru/register.html"
End Sub

Private Sub m_chlog_Click()
    Navigate Me, "http://188.134.2.31:2121/redmine/projects/rb02/wiki/%D0%98%D1%81%D1%82%D0%BE%D1%80%D0%B8%D1%8F_%D0%B8%D0%B7%D0%BC%D0%B5%D0%BD%D0%B5%D0%BD%D0%B8%D0%B9"
End Sub


Private Sub MDIForm_Initialize()
On Error Resume Next

    If App.PrevInstance = False Then
        splash.Show vbModeless, Me
    End If

    Me.Caption = Me.Caption '& " BETA"
    
    Gl.TimerStart = Timer
    
    SetDebugMode
    
    LoadRB ' Загрузка программы
    
    Timer2.Interval = 100
    Timer2.Enabled = True
    
'
'     Открытие файла двойным щелчком
'
'    strCommand = LCase(Command$())
'    If strCommand <> "" Then
'        If Not FindTrafficWindow Then ' ищем окно, если находим - шлем масагу, пишем строку и выходим
'            If Not strCommand = "" Then OpenFilePreload strCommand
'        Else
'            Unload Me
'            End
'        End If
'    End If
'    If Gl.bDebug = False Then If Not FindTrafficWindow Then Call CreateTrafficWindow  ' если окна нет, значит это первая инстансь, создаем окно

    If App.PrevInstance = False Then
        Unload splash
    End If

End Sub

Sub SetDebugMode()
    Debug.Assert zSetTrue(Gl.bDebug)
'    Set PM = New SysPM
End Sub


Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
zSetTrue = True
bValue = True
End Function


Sub LoadRB()
Call VarPtr("VMProtect begin")
On Error GoTo ERR


splash.SetProgress 1

    '
    ' LNG
    '
    Set lng = New Lng_class
    If Not lng Is Nothing Then
        If lng.CountLanguages > 0 Then
            lng.LngCharset = Val(GetSetting(App.ProductName, "Main", "Charset", 204))
            lng.CurrentCodelanguage = GetSetting(App.ProductName, "Main", "CurrentLanguage", 25)
            If lng.CurrentCodelanguage = 0 Then
                lng.SwitchLanguage 25
            Else
                lng.SwitchLanguage lng.CurrentCodelanguage
            End If
            CurrentLocale = lng.GetResIDstring(100)
            If IsNumeric(CurrentLocale) Then MsgBox "The language table is not established!", vbCritical
        Else
            MsgBox "There are no data languages!"
            End
        End If
    Else
    End
    End If
    
splash.SetProgress 2
    
    '
    ' Регистрация расширений
    '
    Dim FILE As String
    FILE = LCase(Command$())
    If FILE <> "" Then
    If InStr(1, FILE, "/a") Then
        AssociateFile "rfdfile", ".rfd", "Roof Builder.rfd", lng.GetResIDstring(1440) & App.ProductName, "Roof Builder", App.Path & "\roof.exe", 0
        AssociateFile "rbpfile", ".rbp", "Roof Builder.rbp", lng.GetResIDstring(1440) & App.ProductName, "Roof Builder", App.Path & "\roof.exe", 0
        End
    ElseIf InStr(1, FILE, "/una") Then
        DellAssociateFile "Roof Builder.rfd", ".rfd"
        DellAssociateFile "Roof Builder.rbp", ".rbp"
        End
    End If
    End If
    
    Gl.UserName = mGetComputerName
    Gl.Uname = mGetUserName
    Ver = App.Major & "." & App.Minor & " build (" & App.Revision & ")"

    Gl.WindowsFont = GetSetting(App.ProductName, "Main", "w_font_name", "MS Sans Serif")
    If Gl.WindowsFont = "" Then Gl.WindowsFont = "MS Sans Serif"
    Gl.WindowsFontSize = Val(GetSetting(App.ProductName, "Main", "w_font_size", 11))
    If Gl.WindowsFontSize = 0 Then Gl.WindowsFontSize = 11
    
    ' подгрузка форм
    officestart_lng
    OfficeStart.initializa
    OfficeStart.Show
    
    ' запуск активации
    Setup.Timer1.Enabled = True
    Setup.Timer1.Interval = 1000 ' 1 секунда
    
'    If App.PrevInstance = False Then
'        splash.Show vbModeless, Me
'        OfficeStart.Show
'        splash.Refresh
'        OfficeStart.initializa
'    Else
'        Load splash
'    End If
    
    'If GetSetting(App.ProductName, "Main", "morecall", 0) = False And App.PrevInstance = True Then MsgBox Lng.GetResIDstring(1418): End

'    Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
'    Me.Top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.Top)
'    Me.Width = GetSetting(App.ProductName, "Position", Me.name & "width", Me.Width)
'    Me.Height = GetSetting(App.ProductName, "Position", Me.name & "height", Me.Height)

    Me.WindowState = GetSetting(App.ProductName, "Position", Me.name & "WindowState", Me.WindowState)

    startp = Time
    
splash.SetProgress 10
        
Exit Sub
ERR:
'MsgBox ERR.Description
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.29." & ERR.Source & "]", ERR.Number, ERR.Description
Call VarPtr("VMProtect end")
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OpenFilePreload Data.Files.Item(1)
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

If UnloadMode = 0 And Print_ALL.is_print = False Then  ' Not ((UnloadMode = vbAppWindows) Or (UnloadMode = vbAppTaskManager)) And
   If Module10.Close_Project(False, True) = True Then
   
    OfficeStart.Clear_project True
    
     '
     Me.Hide
   
     Timer2.Enabled = False ' Отключение системных проверок
     Timer1.Enabled = False ' Отключение автосохранения
     
'     If IsLic Then Setup.SendData 2 ' посылаем сигнал выхода

     Setup.Timer1.Enabled = False ' Отключение подключения к серверу лицензий
   
     CloseDB ' Закрываем базу
      
     If WindowState <> 2 Then
        SaveSetting App.ProductName, "Position", Me.name & "left", OfficeStart.Left
        SaveSetting App.ProductName, "Position", Me.name & "top", OfficeStart.Top
        SaveSetting App.ProductName, "Position", Me.name & "width", OfficeStart.Width
        SaveSetting App.ProductName, "Position", Me.name & "height", OfficeStart.Height
        SaveSetting App.ProductName, "Position", Me.name & "WindowState", OfficeStart.WindowState
     End If
     
     Dim i As Integer
     
     For i = 0 To menu_view_m.Count - 1
         SaveSetting App.ProductName, "menu_view_m", OfficeStart.menu_view_m(i).name & i & ".Checked", OfficeStart.menu_view_m(i).Checked
     Next
     
     ' Save Recently files
    If menrfl_file(0).Caption <> Empty Then
    For i = 0 To ArraySize(RecentlyFiles) - 1
        If RecentlyFiles(i) <> "" Then SaveSetting App.ProductName, "RecentlyFiles", (i), RecentlyFiles(i) Else SaveSetting App.ProductName, "RecentlyFiles", (i), ""
    Next
    End If
      
     ' Выгрузка форм
     Dim frm As Form
     FEXIT = True
     For Each frm In Forms   'Loop thru the forms collection
        Unload frm           'Unload the first form
        Set frm = Nothing    'Set it to nothing
     Next frm
      
     ' Отключение и очистка плагинов
     Erase Plgs
     
     Set lng = Nothing

     Cancel = False
     Unload Me
      
   Else
      Cancel = True
   End If
End If
End Sub


Private Sub MDIForm_Resize()
    If FEXIT Then Exit Sub
    On Error Resume Next
    TabStrip1.Width = Me.Width
    Label1.Width = Me.Width
'    Picture1.Width = Me.Width
    txtLog.Width = Me.Width
'    If Me.Width < 2000 Then Me.Width = 2000
'    If Me.Height < 3000 Then Me.Height = 3000
    
'    If Me.WindowState = 2 Then
'        SaveSetting App.ProductName, "Position", Me.name & "left", OfficeStart.Left
'        SaveSetting App.ProductName, "Position", Me.name & "top", OfficeStart.Top
'        SaveSetting App.ProductName, "Position", Me.name & "width", OfficeStart.Width
'        SaveSetting App.ProductName, "Position", Me.name & "height", OfficeStart.Height
'        SaveSetting App.ProductName, "Position", Me.name & "WindowState", OfficeStart.WindowState
'    Else
'        Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
'        Me.Top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.Top)
'        Me.Width = GetSetting(App.ProductName, "Position", Me.name & "width", Me.Width)
'        Me.Height = GetSetting(App.ProductName, "Position", Me.name & "height", Me.Height)
'    End If
End Sub

Sub Load_views_menu()
    Dim i As Integer
    On Error Resume Next
    'Просматриваем имеем ли мы вообще свойства
    Dim View_m
    View_m = GetSetting(App.ProductName, "menu_view_m", "menu_view_m0.Checked", "*")
    If View_m = "*" Then Exit Sub    'Записи в реестре нет
    View_m = GetAllSettings(App.ProductName, "menu_view_m")
    For i = LBound(View_m, 1) To UBound(View_m, 1)
        OfficeStart.menu_view_m(i).Checked = View_m(i, 1)
    Next
End Sub


Private Sub MDIForm_Load()

'Processing = lng.GetResIDstring(9339)
'Dim hSysMenu As Long
'hSysMenu = GetSystemMenu(Me.hWnd, 0)
'RemoveMenu hSysMenu, 6, MF_BYPOSITION
'RemoveMenu hSysMenu, 0, MF_BYPOSITION
'RemoveMenu hSysMenu, 2, MF_BYPOSITION
'RemoveMenu hSysMenu, 0, MF_BYPOSITION

Load_views_menu

On Error GoTo ERR

HistoryWorking = False

'Просматриваем имеем ли мы вообще меню
Dim i As Integer
Dim n As Integer
Dim RFname As String
Dim vMenu
vMenu = GetSetting(App.ProductName, "RecentlyFiles", "0", "*")
vMenu = GetAllSettings(App.ProductName, "RecentlyFiles")
If IsArray(vMenu) = False Then Exit Sub  'Записи в реестре нет
s2.Visible = True
n = -1
For i = LBound(vMenu, 1) To UBound(vMenu, 1)
    If Left(vMenu(i, 1), 2) <> "\\" And Left(vMenu(i, 1), 2) <> "" Then

        n = n + 1
        If dir(vMenu(i, 1)) <> "" Then
            '            If vMenu(i%, 0) = 0 Then
            '                menrfl_file(0).Visible = True
            '            Else
            menrfl_file(0).Visible = True
            If vMenu(n, 0) <> 0 Then Load menrfl_file(vMenu(n, 0))
            '            End If

            RFname = vMenu(i, 1)
            ReDim Preserve RecentlyFiles(n)
            RecentlyFiles(n) = RFname
            If Len(RFname) > 350 Then PathCompactPath lhDC, RFname, 350
            menrfl_file(vMenu(n, 0)).Caption = LCase(RFname)
        Else
            n = i - 1
        End If

    End If

Next

    'Просматриваем имеем ли мы вообще свойства
    Dim View_m
    View_m = GetSetting(App.ProductName, "menu_view_m", "menu_view_m.Checked", "*")
    If View_m = "*" Then Exit Sub    'Записи в реестре нет
    View_m = GetAllSettings(App.ProductName, "menu_view_m")

    For i = LBound(View_m, 1) To UBound(View_m, 1)
        OfficeStart.menu_view_m(i).Checked = View_m(i, 1)
    Next
    
    OfficeStart.menu_view_m(11).Checked = False

    Exit Sub
ERR:
    n = i - 1
Resume Next
End Sub



Private Sub officestart_lng()
On Error Resume Next

    SetFont Me

    menu_fil.Caption = lng.GetResIDstring(1300)
    
    menu_uusi.Caption = lng.GetResIDstring(1301)
    menu_open.Caption = lng.GetResIDstring(1302)
    menu_save.Caption = lng.GetResIDstring(1303)
    menupman.Caption = lng.GetResIDstring(1323)
    menu_setup.Caption = lng.GetResIDstring(1314)
    menu_end.Caption = lng.GetResIDstring(1313)
    da.Caption = lng.GetResIDstring(1317)
    
    hot_keys.Caption = lng.GetResIDstring(1412)
    
    Toolbar1.Buttons(1).Description = menu_uusi.Caption
    Toolbar1.Buttons(2).Description = menu_open.Caption
    Toolbar1.Buttons(3).Description = menu_save.Caption
'
    Toolbar1.Buttons(1).ToolTipText = menu_uusi.Caption
    Toolbar1.Buttons(2).ToolTipText = menu_open.Caption
    Toolbar1.Buttons(3).ToolTipText = menu_save.Caption
    
    menu_save_as.Caption = lng.GetResIDstring(1304)
    menu_print_valinta.Caption = lng.GetResIDstring(1305)
    menu_print(0).Caption = lng.GetResIDstring(1306)
    menu_print(1).Caption = lng.GetResIDstring(1307)
    menu_print(2).Caption = lng.GetResIDstring(1308)
    menu_print(3).Caption = lng.GetResIDstring(1309)
    menuactrb.Caption = lng.GetResIDstring(9650)
    menu_help.Caption = lng.GetResIDstring(1315)
    menuAddins.Caption = lng.GetResIDstring(1320)
    menu_view.Caption = lng.GetResIDstring(9314)
    
    menu_close.Caption = lng.GetResIDstring(9299)
    
    menu_view_m(0).Caption = lng.GetResIDstring(9315)
    menu_view_m(1).Caption = lng.GetResIDstring(9316)
    menu_view_m(2).Caption = lng.GetResIDstring(9317)
    menu_view_m(3).Caption = lng.GetResIDstring(9318)
    menu_view_m(4).Caption = lng.GetResIDstring(9319)
    menu_view_m(5).Caption = lng.GetResIDstring(9320)
    menu_view_m(6).Caption = lng.GetResIDstring(9312)
    
    menu_view_m(8).Caption = lng.GetResIDstring(9321)
    
    menu_view_m(11).Caption = lng.GetResIDstring(9323)
    
    ab.Caption = lng.GetResIDstring(9325)
    menu_help.Caption = lng.GetResIDstring(9326)
    
    chknewversion.Caption = lng.GetResIDstring(9327)
    menu_help.Caption = lng.GetResIDstring(9340)
    menu_lic.Caption = lng.GetResIDstring(9329)
    m_chlog.Caption = lng.GetResIDstring(9330)
    menuabout.Caption = lng.GetResIDstring(9331)
    
    menuslope.Caption = lng.GetResIDstring(9564)
    recfiles.Caption = lng.GetResIDstring(9615)
    rbhomepage.Caption = lng.GetResIDstring(9616)
    feedback.Caption = lng.GetResIDstring(9617)
    bugreport.Caption = lng.GetResIDstring(9618)
    
    Toolbar1.Buttons(1).ButtonMenus(1).Text = lng.GetResIDstring(1326)
    Toolbar1.Buttons(1).ButtonMenus(2).Text = lng.GetResIDstring(1327)
    
    Me.TabStrip1.Tabs(1).Caption = lng.GetResIDstring(9011)
    Me.TabStrip1.Tabs(2).Caption = lng.GetResIDstring(9012)
    Me.TabStrip1.Tabs(3).Caption = lng.GetResIDstring(9013)
    Me.TabStrip1.Tabs(4).Caption = lng.GetResIDstring(9014)
    Me.TabStrip1.Tabs(5).Caption = lng.GetResIDstring(1317)

    da.Caption = lng.GetResIDstring(1317)
    ab.Caption = lng.GetResIDstring(1318)
    ToolbarCustomize.Caption = lng.GetResIDstring(9658)
    
    menuprofmanager.Caption = lng.GetResIDstring(1025)

End Sub


Private Sub menrfl_file_Click(Index As Integer)
    Dim F As String
    Dim ans As Integer
    On Error Resume Next
    
        If menrfl_file.Item(Index).Caption = "" Then Exit Sub
        If menrfl_file(Index).Caption = "File corrupted" Then Exit Sub
        If OpenFilePreload(RecentlyFiles(Index)) = "" Then
            'RecentlyFiles(Index) = ""
            'Unload OfficeStart.menrfl_file.Item(Index)
            'menrfl_file(Index).Caption = "File corrupted"
            'menrfl_file(Index).Enabled = False
        End If
End Sub


Private Sub menu_close_Click()
    Close_Project True, False
    OfficeStart.Clear_project True
End Sub


Sub menu_end_Click()
Call MDIForm_QueryUnload(0, 0)
End Sub


Private Sub menu_lic_Click()
    lic.Show vbModal, OfficeStart
End Sub


Sub menu_open_Click()
    OfficeStart.MousePointer = 11
    OpenFilePreload
    OfficeStart.MousePointer = 0
End Sub


Sub menu_help_Click()
    Navigate Me, "http://www.firma-ms.ru/forum/index.php?showforum=3"
End Sub


Sub menu_print_Click(Index%)
Dim Slope As Integer

On Error GoTo ERR

CommonDialog1.CancelError = True
CommonDialog1.ShowPrinter
NumCopies = CommonDialog1.Copies

FlagPrinter = 1

Select Case Index
    Case 0 ' ALL

        Summary Printer

        MainPicture Printer

        For Slope = 1 To KolvoScatov Step 1
            If Slope > 26 Then
                If SlP(Slope).CountSheets > 0 Then Print_Scat Slope, Printer
            Else
                If SlP(Slope).CountSheets > 0 Then Print_Scat Slope, Printer
            End If

        Next Slope

    Case 2
        Print_lape.lst_print.Clear

        For Slope = 1 To KolvoScatov Step 1
            If Slope > 26 Then
                If SlP(Slope).CountSheets > 0 Then Print_lape.lst_print.AddItem Chr$(Slope + 70) & " " & SlP(Slope).ProfilName
            Else
                If SlP(Slope).CountSheets > 0 Then Print_lape.lst_print.AddItem Chr$(Slope + 64) & " " & SlP(Slope).ProfilName
            End If

        Next Slope

        Print_lape.Show vbModal, OfficeStart

    Case 1
        ' Settings Printer
        MainPicture Printer

    Case 3
        ' Settings Printer
        Summary Printer
End Select

NumCopies = 1
Exit Sub

ERR:
If ERR.Number <> 32755 Then
    'STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" & vbNewLine
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.30." & ERR.Source & "]", ERR.Number, ERR.Description
    If CommonDialog1.CancelError Then
        MsgBox lng.GetResIDstring(9332), vbExclamation, lng.GetResIDstring(1413)
    Else
        MsgBox ERR.Description, vbCritical, lng.GetResIDstring(1413)
    End If
End If
End Sub


Private Sub menu_print_lape_view_Click(Index As Integer)
    Print_Scat Index, frmView.pic_view(0)
End Sub


Private Sub menu_print_lape_Click(Index As Integer)
    Print_Scat Index, Printer
End Sub


Private Sub view_Click()
    Print_Scat N_Slope, frmView.pic_view(0)
End Sub

Sub menu_save_as_Click()
    If IsLic = False Then GoTo NOLIC
    SaveAS
    Exit Sub
NOLIC:
    Module10.withoutl
'    OfficeStart.Enabled = True
End Sub


Sub menu_save_Click()
If IsLic = False Then GoTo NOLIC
    Me.MousePointer = 11
    SaveFile IIf(CurrentProjectDir <> "", CurrentProjectDir, ProjectsDir), CurrentFile
    Me.MousePointer = 0
    Exit Sub
NOLIC:
    Module10.withoutl
'    OfficeStart.Enabled = True
End Sub

Sub menu_setup_Click()
'If Module10.Close_Project(True, True, True) Then
    Setup.Show vbModal, Me
'End If
End Sub


Private Sub menu_upd_Click()
    Dim lape As Integer
    On Error Resume Next
    lape = N_Slope
    If CurrentFile = "" Then Exit Sub
    If Not ROOFPIC.Visible = True And Not Project.Visible = True Then
        N_Slope = lape
        Lapemenu.Command4_Click
    Else
        Project.Command4.value = True
    End If
End Sub


Sub menu_uusi_Click()
On Error Resume Next
    If Module10.Close_Project(True, True) = True Then
        OfficeStart.Clear_project
        newproject.Show vbModal, OfficeStart
        If CurrentFile <> "" Then OfficeStart.TabStrip1.Tabs(1).Selected = True
    End If
End Sub


Public Function SaveFile(ByRef dir As String, ByRef FILE As String) As Boolean
If CurrentFile = "" Then Exit Function
If IsLic = False Then GoTo NOLIC
OfficeStart.Enabled = False
fpt_ dir, FILE
OfficeStart.Enabled = True
SaveFile = True
Exit Function
NOLIC:
SaveFile = False
End Function


Public Function SaveAS() As Boolean
  Dim tempfilename As String
  Dim Path As String
  
  If IsLic = False Then GoTo NOLIC

  OfficeStart.Enabled = False
  
    tempfilename = ""
    Do While tempfilename = ""

        If Gl.FileNameExtension = ".rfd" Then
            tempfilename = Dialog.GetFileName(tempfilename$, _
            "RFD v" & FILEVER & " (*.rfd)|*.rfd", ProjectsDir, False, Me.hwnd)
        Else
            tempfilename = Dialog.GetFileName(tempfilename$, _
            "Roof Builder Project v" & FILEVER & " (*.rbp)|*.rbp", ProjectsDir, False, Me.hwnd)
        End If

        If tempfilename = "" Then GoTo NOFILE

        Path = ProjectsDir
        
        ProjectsDir = LCase(tempfilename)
        ProjectsDir = Left$(tempfilename, InStrRev(tempfilename, "\", -1))
        CurrentFile = Right(tempfilename, Len(tempfilename) - InStrRev(tempfilename, "\", -1))
        
        CurrentProjectDir = ProjectsDir
           
        tempfilename = CheckFileName(CurrentFile)

    Loop
    
    CurrentFile = tempfilename

    fpt_ ProjectsDir, CurrentFile
    
NOFILE:

    OfficeStart.Enabled = True
    
SaveAS = True
Exit Function
NOLIC:
SaveAS = False
End Function


Public Function CheckFileName(F As String) As String
    Dim l00A2, l00A4 As Integer
    On Error Resume Next
    
        l00A2 = 0
        If Len(F) > Gl.file_name_size Then l00A2 = 0
        For l00A4 = 1 To Len(F) Step 1
            Select Case mID$(F, l00A4, 1)
                Case "a" To "z", "A" To "Z", "0" To "9", "а" To "я", "А" To "Я", " ", "-", ".", "_"
                Case Else
                    l00A2 = -1
                    Exit For
            End Select
        Next l00A4

        If l00A2 = -1 Then
            CheckFileName = ""
            MsgBox lng.GetResIDstring(1425), vbCritical, lng.GetResIDstring(1413) '"Вы употребили неверные символы в названии проекта.")
        Else
            CheckFileName = F
        End If
End Function


Public Function initializa() As Boolean
On Error GoTo ERR
   
Dim ln As String
Dim i As Integer, n As Integer

Screen.MousePointer = 11

    '
    ' Основные переменные
    '
    CountRicentlyFiles = 19
    Gl.file_name_size = 22

    
    NumCopies = 1
    isSave = False
    

    ProjectsDir = GetSetting(App.ProductName, "Main", "url_work", App.Path & "\data")

    Gl.TempDir = GetTempDir()
    If Gl.TempDir <> "" And CheckWriteDir(Gl.TempDir) Then
        Gl.TemporaryFileName = Gl.TempDir & "~tmp.rfd"
    Else
        MsgBox lng.GetResIDstring(9332, "%DIR%", Gl.TempDir), vbInformation, "Roof Builder"
        Gl.TemporaryFileName = ""
        Gl.TempDir = ""
    End If

    ' GL
    Gl.PrintFont = CStr(GetSetting(App.ProductName, "Main", "font_name", "MS Sans Serif"))
    Gl.PrintFontSize = Val(GetSetting(App.ProductName, "Main", "font_size", 11))

    Gl.UserCreatProject = Gl.Uname
    
    '
    ' Проверка и загрузка файла конфигурации
    '
    If Gl.ConfigDir = "" Then Gl.ConfigDir = App.Path & "\cfg"
    
    Gl.FileName = GetSetting(App.ProductName, "Main", "mdbcofigfile", Gl.ConfigDir & "\materials.mdb")
    Load Setup
    
    Setup.Text10.Text = GetSetting(App.ProductName, "Main", "fimname", "Moscow Builders ltd")
    Setup.Text11.Text = GetSetting(App.ProductName, "Main", "firminfo", " www.firma-ms.ru, www.roof-builder.ru (Rus), www.roofbuilder.net (Eng)")
    Setup.Text17.Text = GetSetting(App.ProductName, "Main", "firmcustomer", "")

splash.SetProgress 3
    
    Load Setup

splash.SetProgress 4
    
splash.SetProgress 4.1
    
    Clear_project True ' Подготовка массивов к работе (Инициализация переменных)

splash.SetProgress 5

    '
    ' Операция с плугинами
    '
    Plugins.GetPlugins

    If Setup.Combo3.ListCount <= 0 Then
        On Error GoTo ERRplg
        Dim txt_err As String
        For i = 0 To UBound(Plgs)
            txt_err = txt_err & Plgs(i).ERR & vbNewLine
        Next
ERRplg:
        MsgBox lng.GetResIDstring(9337) & vbNewLine & txt_err, vbCritical
        End
        
    Else
        LNC = GetSetting(App.ProductName, "CalcOption", "CCMn", "1")
        If Setup.Combo3.ListCount - 1 < LNC Then LNC = Setup.Combo3.ListCount - 1
        Setup.Combo3.ListIndex = LNC
    End If

splash.SetProgress 6

    '
    ' Вывод окна конфигурации при первом старте
    '
'    If GetSetting(App.ProductName, "RunUp", "FirstTime", False) = False Then
'        SaveSetting App.ProductName, "RunUp", "FirstTime", True
'    Dim im As Integer
'        Screen.MousePointer = 0
'        'im = MsgBox("This is a first time you are running " & App.ProductName & ". It is strongly recommend you to setup editor in your own taste. Do you want to do it now?", vbYesNo + vbQuestion, App.ProductName)
'        MsgBox lng.GetResIDstring(1437, "%PRODUCT%", App.ProductName), vbCritical, App.ProductName
'        If im = vbYes Then
'            Me.MousePointer = 0
'            Setup.Show vbModal, Me
'        End If
'        Screen.MousePointer = 11
'    End If
    
splash.SetProgress 7

    ' Загрузка главного окна управления
    ' Забивка данных профилей
    Load Project


Screen.MousePointer = 0
Exit Function
ERR:
Screen.MousePointer = 0
MsgBox ERR.Description, vbCritical
End Function


Public Sub menu_view_m_Click(Index As Integer)
On Error Resume Next

Select Case Index
    Case 0
        If OfficeStart.menu_view_m(0).Checked = False Then
            OfficeStart.menu_view_m(0).Checked = True
        Else
            OfficeStart.menu_view_m(0).Checked = False
        End If
    Case 1
        If OfficeStart.menu_view_m(1).Checked = False Then
            OfficeStart.menu_view_m(1).Checked = True
        Else
            OfficeStart.menu_view_m(1).Checked = False
        End If

        Lapepic.Draw_Systems Lapepic.Picture1
    Case 2
        If OfficeStart.menu_view_m(2).Checked = False Then
            OfficeStart.menu_view_m(2).Checked = True
        Else
            OfficeStart.menu_view_m(2).Checked = False
        End If

        Lapepic.Draw_Systems Lapepic.Picture1
    Case 3
        If OfficeStart.menu_view_m(3).Checked = False Then
            OfficeStart.menu_view_m(3).Checked = True
        Else
            OfficeStart.menu_view_m(3).Checked = False
        End If

        Lapepic.Draw_Systems Lapepic.Picture1
    Case 4
        If OfficeStart.menu_view_m(4).Checked = False Then
            OfficeStart.menu_view_m(4).Checked = True
        Else
            OfficeStart.menu_view_m(4).Checked = False
        End If

        Lapepic.Draw_Systems Lapepic.Picture1
    Case 5
        If OfficeStart.menu_view_m(5).Checked = False Then
            OfficeStart.menu_view_m(5).Checked = True
        Else
            OfficeStart.menu_view_m(5).Checked = False
        End If

        Lapepic.Draw_Systems Lapepic.Picture1
    Case 6
        If OfficeStart.menu_view_m(6).Checked = False Then
            OfficeStart.menu_view_m(6).Checked = True
        Else
            OfficeStart.menu_view_m(6).Checked = False
        End If

        Lapepic.Draw_Systems Lapepic.Picture1
    Case 8
        If OfficeStart.menu_view_m(8).Checked = False Then
            Lapepic.ListEdit True
        Else
            Lapepic.ListEdit False
        End If
    Case 11
        If IsLoadForm("Lapepic") Then
        If Lapepic.Command10.value = False Then
            Lapepic.Command10.value = True
        Else
            Lapepic.Command10.value = False
        End If
        End If
End Select
End Sub


Private Sub menuabout_Click()
    About.Show vbModal, Me
End Sub


Private Sub menuactrb_Click()
    Navigate Me, "http://roof-builder.ru/register.html"
End Sub


Private Sub menudump_Click()
    dump.Show vbModeless, OfficeStart
End Sub


Private Sub menupman_Click()
    OpenFilePreload , True
End Sub


Private Sub menuprofmanager_Click()
Load ChangeProfil
If ChangeProfil.lstprof.ListCount > 0 Then
ChangeProfil.lstprof.ListIndex = 0
ChangeProfil.lstprof_Click
End If
ChangeProfil.Show vbModal, OfficeStart
Unload ChangeProfil
End Sub

Private Sub menuRedo_Click()
    Dim Button As MSComctlLib.Button
    Set Button = OfficeStart.Toolbar1.Buttons(7)
    OfficeStart.Toolbar1_ButtonClick Button
End Sub

Public Sub menuslope_Click()
    SlopeSampleModule.Show vbModal, Me
End Sub


Private Sub menuUndo_Click()
    Dim Button As MSComctlLib.Button
    Set Button = OfficeStart.Toolbar1.Buttons(6)
    OfficeStart.Toolbar1_ButtonClick Button
End Sub

Private Sub rbhomepage_Click()
    Navigate Me, "http://roof-builder.ru"
End Sub


Private Sub TabStrip1_Click()
On Error Resume Next

Select Case TabStrip1.SelectedItem.Index
    Case 1
        
'        If IsLoadForm("Project") Then Unload Project
        If IsLoadForm("ROOFPIC") Then Unload ROOFPIC
        If IsLoadForm("Lapemenu") Then Unload Lapemenu
        If IsLoadForm("Lapepic") Then Unload Lapepic
        If IsLoadForm("frmView") Then Unload frmView
        
        
'        Load Project
        OfficeStart.menu_view.Visible = False
        Project.WindowState = 2
        Project.Show

    Case 2
    
        If IsLoadForm("ROOFPIC") Then Exit Sub

        If CurrentFile = "" Then
            Beep
            OfficeStart.menu_uusi_Click
            Exit Sub
        End If

        If IsLoadForm("ROOFPIC") Then Unload ROOFPIC
        If IsLoadForm("Lapemenu") Then Unload Lapemenu
        If IsLoadForm("Lapepic") Then Unload Lapepic
        If IsLoadForm("frmView") Then Unload frmView
        
        Load ROOFPIC
        ROOFPIC.MousePointer = 1
        ROOFPIC.sTabFx1.SelectTab 2
        ROOFPIC.WindowState = 2
        ROOFPIC.Show

        'OptionDMM = "Mdraw"
        OfficeStart.menu_view.Visible = False
        Project.Hide

    Case 3
    
        If IsLoadForm("Lapemenu") Then Exit Sub
    
        If Project.Label3.Caption = "" And Gl.FileNameExtension = ".rfd" Then
            If Project.Label3.Caption = "" Then
                MsgBox lng.GetResIDstring(1487), vbCritical, App.ProductName
            Else
                MsgBox lng.GetResIDstring(1437, "%PROFIL%", Project.Label3.Caption), vbCritical, App.ProductName
            End If

            Me.TabStrip1.Tabs(1).Selected = True
            Exit Sub
        End If
        
        If IsLoadForm("ROOFPIC") Then Unload ROOFPIC
        If IsLoadForm("Lapepic") Then Unload Lapepic
        If IsLoadForm("frmView") Then Unload frmView

        Project.Hide
        
        If IsLoadForm("Lapemenu") = False Then
            OfficeStart.menu_view.Visible = False
            Load Lapemenu
            Lapemenu.WindowState = 2
            Lapemenu.Show
        End If
    Case 4
    
        If Project.Label3.Caption = "" And FileNameExtension = ".rfd" Then
            OfficeStart.TabStrip1.Tabs(3).Selected = True
            Exit Sub
        End If
        If KolvoScatov = 0 Then Me.TabStrip1.Tabs(3).Selected = True: Exit Sub
        
        If IsLoadForm("ROOFPIC") Then Unload ROOFPIC
        If IsLoadForm("frmView") Then Unload frmView
        If IsLoadForm("Lapemenu") Then Unload Lapemenu
        
        If IsLoadForm("Lapepic") = False Then
            If IsLoadForm("Lapemenu") = False Then Load Lapemenu
            Lapemenu.List1.ListItems(N_Slope).Selected = True
            Lapemenu.Command4_Click
        End If

    Case 5
    
        If IsLoadForm("frmView") Then Exit Sub

        If CurrentFile = "" Then
'            Beep
            OfficeStart.menu_uusi_Click
            Exit Sub
        End If
        
        If IsLoadForm("ROOFPIC") Then Unload ROOFPIC
        If IsLoadForm("Lapemenu") Then Unload Lapemenu
        If IsLoadForm("Lapepic") Then Unload Lapepic

        OfficeStart.menu_view.Visible = False

        Project.Hide
        Load frmView
        frmView.List1.ListIndex = 0
        frmView.WindowState = 2
        frmView.Show
End Select
End Sub


Private Sub Timer1_Timer()
On Error Resume Next

    If FEXIT Then Exit Sub
    If IsLic = False Then Exit Sub
    
    If Setup.Check3.value = 1 And (Project.Label3.Caption <> "" Or Gl.FileNameExtension = ".rbp") Then
        If isSave And Format(DateAdd("n", Setup.SetDataSaveTimer.Text, TimeStart), "hh:mm:ss") < Format(Now, "hh:mm:ss") Then
            If CurrentFile <> "" Then
            
                SaveFile IIf(CurrentProjectDir <> "", CurrentProjectDir, ProjectsDir), CurrentFile
                TimeStart = Now
                
            End If
        End If
    End If
    
End Sub


Private Sub Timer2_Timer()
    If FEXIT Then Exit Sub
    On Error GoTo ERR
    If CurrentFile <> "" Then
    
        'If IsLoadForm("Project") Then

            If CurrentFile <> "" Then Project.Text1 = CurrentFile
            If ProjectsDir <> "" Then Project.Label11 = ProjectsDir
            'If Gl.PrjDescrib <> "" Then Project.Text3.Text = Gl.PrjDescrib

        'End If
        
        If TemporaryFileName = "" Then
            Toolbar1.Buttons(6).Enabled = False
            Toolbar1.Buttons(7).Enabled = False
        End If
        
        Timer1.Enabled = True
        ' Отключение свойств при открытом проекте
        menu_setup.Enabled = False
'        Toolbar1.Buttons(8).Enabled = False
        da.Enabled = True
        menu_print_valinta.Enabled = True
        menu_upd.Enabled = True
        menu_save_as.Enabled = True
        menu_save.Enabled = True
        menu_close.Enabled = True
'        menuprofmanager.Enabled = False
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(12).Enabled = True
        Picture3.Visible = True ' Управляющий таб
        mOpWp.Enabled = True
        Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(1).Picture
        OfficeStart.StatusBar.Panels(1).Text = OptionDMM & "(" & SlP(N_Slope).CountOfLines & "," & SlP(N_Slope).CountOfPoints & ")"
        
        Setup.Combo4.Enabled = False
        Setup.Label36.Enabled = False
        
        If IsLoadForm("Lapepic") Or IsLoadForm("ROOFPIC") Then
            
            If IsLoadForm("Lapepic") Then
                SlP(N_Slope).ScaleLeftS = Lapepic.Picture1.ScaleLeft
                SlP(N_Slope).ScaleWidthS = Lapepic.Picture1.ScaleWidth
                SlP(N_Slope).ScaleTopS = Lapepic.Picture1.ScaleTop
                SlP(N_Slope).ScaleHeightS = Lapepic.Picture1.ScaleHeight
            End If

            If isChange = True And HistoryWorking = False And Gl.TemporaryFileName <> "" Then

                Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(3).Picture
                If Positions = 500 Then
                    
                    ' Чистим так как превышено кол-во эл истории
                    CurentPosition = 0
                    Positions = 0
                    Toolbar1.Buttons(6).Enabled = False
                    Toolbar1.Buttons(7).Enabled = False
                    OfficeStart.menuUndo.Enabled = False
                    OfficeStart.menuRedo.Enabled = False
                    Kill TemporaryFileName
                    
                End If

                ' Производим запись точек в истории
                Positions = SaveF(TemporaryFileName, N_Slope, CurentPosition)
                CurentPosition = Positions
                SetChange False
                
                Toolbar1.Buttons(7).Enabled = False
                OfficeStart.menuRedo.Enabled = False
                    
                ' Сделать доступным механизм возврата действий
                If Positions > 1 Then
                    Toolbar1.Buttons(6).Enabled = True
                    OfficeStart.menuUndo.Enabled = True
                End If

                Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(1).Picture
                
            End If
            
            OfficeStart.StatusBar.Panels(5).Text = CurentPosition & "/" & Positions
            
        End If

    Else
        
        ' Включение свойств при закрытом проекте
        menu_setup.Enabled = True
'        Toolbar1.Buttons(8).Enabled = True
        Timer1.Enabled = False
        N_Slope = 1: az = "a"
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(12).Enabled = False
        menu_upd.Enabled = False
        StatusBar.Panels(3).Text = ""
        menu_save.Enabled = False
        menu_save_as.Enabled = False
        menu_close.Enabled = False
'        menuprofmanager.Enabled = True
        da.Enabled = False
        menu_print_valinta.Enabled = False
        mOpWp.Enabled = False
        Picture3.Visible = False
        Set OfficeStart.StatusBar.Panels(4).Picture = OfficeStart.ImageList2.ListImages(3).Picture
        OfficeStart.StatusBar.Panels(1).Text = ""
        
        Setup.Combo4.Enabled = True
        Setup.Label36.Enabled = True
        
        OfficeStart.Toolbar1.Buttons(7).Enabled = False
        OfficeStart.Toolbar1.Buttons(6).Enabled = False
        OfficeStart.menuUndo.Enabled = False
        OfficeStart.menuRedo.Enabled = False

    End If
    Exit Sub
ERR:
    OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.31." & ERR.Source & "]", ERR.Number, ERR.Description
    Timer1.Enabled = False
    Timer2.Enabled = False
    Resume Next
End Sub


Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        menu_uusi_Click
    Case 2
        menu_open_Click
    Case 3
        menu_save_Click
'    Case 5
'        menu_upd_Click ' Переоткрытие файла (все исправления сбрасываются)
    Case 6
        If TemporaryFileName = "" Then Exit Sub
        
        HistoryWorking = True
        
        If CurentPosition <> 1 Then
            CurentPosition = CurentPosition - 1
        End If
        Toolbar1.Buttons(7).Enabled = True
        OfficeStart.menuRedo.Enabled = True
        CurentPosition = ReadF(TemporaryFileName, N_Slope, CurentPosition)

        If CurentPosition = 1 Then
            Toolbar1.Buttons(6).Enabled = False
            OfficeStart.menuUndo.Enabled = False
        End If
        
        If IsLoadForm("Lapepic") Then
           Lapepic.Command5.value = True ' Прорисовка фигуры с автоцентрированием
'            Lapepic.Draw_Systems Lapepic.Picture1 ' Прорисовка фигуры
        ElseIf IsLoadForm("ROOFPIC") Then
'            Draw_Systems ROOFPIC.Picture1
            ROOFPIC.Command5.value = True
        End If
        
        SetChange False
        
        HistoryWorking = False

'        isSave = True
    Case 7
        If TemporaryFileName = "" Then Exit Sub
        
        HistoryWorking = True
        
        Toolbar1.Buttons(6).Enabled = True
        OfficeStart.menuUndo.Enabled = True
        CurentPosition = CurentPosition + 1
        CurentPosition = ReadF(TemporaryFileName, N_Slope, CurentPosition)
        If CurentPosition = Positions Then
            Toolbar1.Buttons(7).Enabled = False
            OfficeStart.menuRedo.Enabled = False
        End If
        
        If IsLoadForm("Lapepic") Then
           Lapepic.Command5.value = True ' Прорисовка фигуры с автоцентрированием
'            Lapepic.Draw_Systems Lapepic.Picture1 ' Прорисовка фигуры
        ElseIf IsLoadForm("ROOFPIC") Then
'            Draw_Systems ROOFPIC.Picture1
            ROOFPIC.Command5.value = True
        End If
        
        SetChange False
        
        HistoryWorking = False

'        isSave = True
    Case 9
        OpenFilePreload , True
    Case 11
        menu_setup_Click
    Case 12
       da_Click
    Case 13
    Call MDIForm_QueryUnload(0, 0)
End Select
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next

    Dim indexselected As String

        If ButtonMenu.Parent.Index = 1 Then
            If ButtonMenu.Key = "rfd" Then
                indexselected = ".rfd"
            Else
                indexselected = ".rbp"
            End If

        End If

        Call OfficeStart.Clear_project

        Gl.FileNameExtension = indexselected
        If CurrentFile = "" Then CurrentFile = Project.Label14 & Gl.FileNameExtension

        isSave = True

        Project.Text3 = ""

        If Project.Combo1(1).ListCount > 0 Then Project.Combo1(1).ListIndex = 0
        If Project.Combo1(2).ListCount > 0 Then Project.Combo1(2).ListIndex = 0
        If Project.Combo1(3).ListCount > 0 Then Project.Combo1(3).ListIndex = 0

        UserCreatProject = Gl.Uname

        Unload Me

        Project.Text6 = UserCreatProject
        Project.Show
        OfficeStart.StatusBar.Panels(2) = ProjectsDir & CurrentFile & " NEW [OK]"

End Sub


Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' To use the PopMenu, click on the Toolbar with the right button
    If Button = 2 Then PopupMenu ToolBarSettings, 2, , , ToolbarCustomize
    '
End Sub


Private Sub ToolbarCustomize_Click()
    Toolbar1.Customize
End Sub


Sub Clear_project(Optional pclose As Boolean)
On Error Resume Next

MAINMAXSLOPELINE = 100 ' количество допустимых линий на главном рисунке 200
MAXSLOPELINE = 40 ' количество допустимых точек 20
'
' SLOPE
'
' Ограничение лицензии
If IsLic Xor True Then
    MAXSLOPES = 1 Xor 4 ' = 5
Else
    MAXSLOPES = 52 ' количество допустимым расчитываемых поверхностей 52
End If
MAXSLOPELISTS = 10000 ' количество допустимым листов на скате (макс 32 000)

If pclose Then

    Erase Points_m_A, Points_m_B, Main_Points_X, Main_Points_Y, Label_X, Label_Y, Lape_Points_X, Lape_Points_Y, _
    Lape_Lines, SlP, List_Properties_Length, List_Properties_PX, List_Properties_PY
    
Else
  
    '*****************************
    ' MAIN
    '
    ReDim Points_m_A(1 To MAINMAXSLOPELINE) 'CountOfLines - ?
    ReDim Points_m_B(1 To MAINMAXSLOPELINE) 'CountOfPoints - ?
    
    ReDim Main_Points_X(MAINMAXSLOPELINE)
    ReDim Main_Points_Y(MAINMAXSLOPELINE)

    ReDim Label_X(1 To MAXSLOPES)
    ReDim Label_Y(1 To MAXSLOPES)

    ReDim Lape_Points_X(1 To MAXSLOPES, 1 To MAXSLOPELINE + 2)
    ReDim Lape_Points_Y(1 To MAXSLOPES, 1 To MAXSLOPELINE + 2)
    ReDim Lape_Lines(1 To MAXSLOPES, 1 To MAXSLOPELINE + 2, 1)
    
    ReDim SlP(1 To MAXSLOPES)
    
'    ReDim Lape_Guidings(1 To MAXSLOPES, 1 - GUIDINGS) ' Массив с координатами вспомогательных линий

    ReDim List_Properties_Length(1 To MAXSLOPES, MAXSLOPELISTS)   ' длина полосы
    ReDim List_Properties_PX(1 To MAXSLOPES, MAXSLOPELISTS)  ' Координаты по X (НАЧАЛО ПРОРИСОВКИ)
    ReDim List_Properties_PY(1 To MAXSLOPES, MAXSLOPELISTS)   ' Координаты по Y (НАЧАЛО ПРОРИСОВКИ)
    '*****************************
    
End If
    
    SelectLists.Clear

    N_Slope = 1
    
    ' Инициализация SCALE для MAIN
    ScaleLeft_Main = 0
    ScaleWidth_Main = 1200
    ScaleTop_Main = 625
    ScaleHeight_Main = -625
    
    MainCountOfPoints = 0
    MainCountOfLines = 0
    KolvoScatov = 0
    MainDescrib = ""
    
End Sub


Sub AddAtmel(ByVal FILE As String)
Dim i As Integer
Dim bFlag As Boolean
Dim RFile As String

On Error Resume Next

    RFile = FILE
    PathCompactPath lhDC, FILE, 350

    'Записываем вызванный файл в меню
    'но предварительно проверяем есть ли он у нас
    
    bFlag = False
        
    For i = 0 To menrfl_file.Count - 1
        If RFile = menrfl_file(i).Caption Then
            menrfl_file(i).Checked = True
            bFlag = True
        Else
            menrfl_file(i).Checked = False
        End If
    Next
        
    If bFlag = False Then
        'Смотрим не превысили мы планку в меню?
        If menrfl_file.UBound >= Setup.Text2 - 1 Then
            'Если да, то сдвигаем все имеющиеся записи на один в перед
            'и освобождаем первое меню, при этом файл записанный в
            'последнем меню исчезает
            For i = menrfl_file.UBound To menrfl_file.LBound Step -1
                If i = menrfl_file.LBound Then
                    menrfl_file(0).Caption = FILE
                    'menrfl_file(0).Checked = True
                    Dim RFC As Integer
                    RFC = ArraySize(RecentlyFiles)
                    ReDim Preserve RecentlyFiles(RFC)
                    RecentlyFiles(RFC) = RecentlyFiles(0)
                    RecentlyFiles(0) = RFile
                    Exit For
                End If

                menrfl_file(i).Caption = menrfl_file(i - 1).Caption
                RecentlyFiles(i) = RecentlyFiles(i - 1)
            Next
     
        Else
            ' Загружаем меню
            Dim mc As Integer
            mc = ArraySize(RecentlyFiles)
            If mc = -1 Then mc = 0
            If mc > 0 Then Load menrfl_file(mc)
            menrfl_file(mc).Visible = True
            menrfl_file(mc).Caption = FILE
            'menrfl_file(mc).Checked = True
            ReDim Preserve RecentlyFiles(mc)
            RecentlyFiles(mc) = RFile
        End If
    End If
    
End Sub



Function OpenFilePreload(Optional FILE As String, Optional isprojectsmanager As Boolean, Optional dontshowQ As Boolean) As String
On Error GoTo ERR

If IsLic = False Then GoTo NOLIC

If Module10.Close_Project(True, True) Then

OfficeStart.Clear_project

If isprojectsmanager = True Then
    If IsLic = False Then GoTo NOLIC
    SO.Show vbModal, Me ' Запуск менедера проектов
    If dontshowQ = False Then Exit Function
End If
    
Dim Catalog As String, name As String

If FILE = "" Then  ' если файл не передан открытие диалога открытия файла
    Dim temppath As String
  
'    Me.Enabled = False
      
    temppath = Dialog.GetFileName("", "RFD v0-" & FILEVER & " (*.rfd)|*.rfd" & _
    "|Roof Builder Project v0-" & FILEVER & " (*.rbp)|*.rbp", ProjectsDir, True, Me.hwnd)
    
'    Me.Enabled = True

    If temppath = "" Then
'    Me.MousePointer = 0
    Exit Function
    End If
    
    name = Right(temppath, Len(temppath) - InStrRev(temppath, "\", -1))
    Catalog = Left$(temppath, InStrRev(temppath, "\", -1))
Else
    name = Right(FILE, Len(FILE) - InStrRev(FILE, "\", -1))
    Catalog = Left$(FILE, InStrRev(FILE, "\", -1))
End If

Factory_Name = ""
Profil_Name = ""
    
OPENFILE:
OpenFilePreload = Ld_(Catalog, name) ' Физическое открытие файла

If InStr(1, OpenFilePreload, "~$", vbTextCompare) Then OpenFilePreload = Replace(OpenFilePreload, "~$", "")

If OpenFilePreload <> "" Then ' Если файл удалось открыть
        
    If Gl.FileNameExtension <> ".rbp" Then
    
        ' Необязательно
        Project.Label2.Tag = 0
        Project.Label2.Caption = ""
        Project.Label3.Caption = ""
        
        Dim FactoryID As Integer
        If Factory_Name <> "" Then
            FactoryID = GetFactoryID(Factory_Name)
        Else
            FactoryID = 0
        End If
    
        If Profil_Name <> "" Then
            
            ' Запрос профиля в базе данных
            Dim RS As Recordset
            Set RS = GetProfilData(Profil_Name, FactoryID)
            If Not RS Is Nothing Then
                
                ' Factory
                Project.Label2.Tag = FactoryID
                Project.Label2.Caption = Factory_Name
                ' Profil
                Project.Label3.Tag = RS!id
                Project.Label3.Caption = Profil_Name
                
                If KolvoScatov <> 0 Then
                    ' Заполнение массива имен производителей
                    loadFactory Factory_Name
                    ' Заполнение массива имен профилей
                    loadProfil Profil_Name
                End If
                
            End If
            Set RS = Nothing
            
            CurrentProjectDir = Catalog
            
        End If
        
    End If
    
    Project.SwitchProfil
            
    Project.Combo1(1) = Gl.width1
    Project.Combo1(2) = Gl.cover
    Project.Combo1(3) = Gl.ColorRoof
    Project.Text6 = UserCreatProject

    OfficeStart.TabStrip1.Tabs(1).Selected = True
Else
    OpenFilePreload = ""
    CurrentProjectDir = ""
    Exit Function
End If

' возврат имени файла
CurrentFile = OpenFilePreload
End If
Exit Function

ERR:
'STRERR = STRERR & time & ". (" & Me.name & ") ... [ERROR] N " & ERR.Number & " (" & ERR.Description & ")" &  vbNewLine
OfficeStart.AddToList OfficeStart.txtLog, "[ERROR.32." & ERR.Source & "]", ERR.Number, ERR.Description
Resume Next

NOLIC:
    Module10.withoutl
    OfficeStart.Enabled = True
End Function


Public Sub AddToList(List As ListBox, dtype As String, code As Integer, Item As String)
On Error Resume Next
If List.ListCount > 5000 Then List.RemoveItem (0)
List.AddItem Time & ": " & dtype & " (" & code & ") " & Item
'FILE.SaveToFile "C:\root\htdocs\bot", "C:\root\htdocs\bot\" & Date & ".debug.log", sKey & "> " & Item
List.Selected(List.ListCount - 1) = True
End Sub


Private Sub txtLog_DblClick()
On Error GoTo ERR
Dim i As Integer
Load Teksti
For i = 0 To Me.txtLog.ListCount
Teksti.Text1 = Teksti.Text1 & Me.txtLog.List(i) & vbNewLine
Next
ERR:
Teksti.Show vbModal, OfficeStart
Unload Teksti
End Sub
