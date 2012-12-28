VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "СhameleonButton.ocx"

Begin VB.Form Pitmani 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Форма печати"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6495
   Icon            =   "Pitmani.frx":0000
   LinkTopic       =   "Форма1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6495
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Height          =   4215
      IntegralHeight  =   0   'False
      ItemData        =   "Pitmani.frx":030A
      Left            =   0
      List            =   "Pitmani.frx":030C
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   840
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   6495
      Begin roof.isButton Command1 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   873
         Style           =   1
         Caption         =   "..."
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin roof.isButton Command2 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   873
         Style           =   1
         Caption         =   "..."
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label Метка1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Для редактирования"
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
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "Pitmani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim option_p As Single
Dim m001E As Single
Dim m0022 As Single


Private Sub Combo1_Click()
    Dim i As Integer
    Dim n As Integer

    List1.clear

    If Combo1.ListIndex = 0 Then
        n = Print_ALL.prepair(0) ' Сортировка по убыванию
    ElseIf Combo1.ListIndex = 1 Then
        n = Print_ALL.prepair(1) ' Сортировка по возростанию
    ElseIf Combo1.ListIndex = 2 Then
        n = Print_ALL.prepair(2) ' Сортировка по наименованию
    End If

    Dim ln As String
    For i = 1 To n Step 1
        Pitmani.List1.AddItem AddSpaceLeft$(5, Print_ALL.For_print_lapesn(i).len) & "        " & Print_ALL.For_print_lapesn(i).name & _
        Space(7) & Print_ALL.For_print_lapesn(i).prof & Space(7) & Print_ALL.For_print_lapesn(i).factory_name
    Next i
      
End Sub



Sub Command1_Click()
    Pitmani.Hide
    pitmanprintunload = False
    frmView.pFunc = Combo1.ListIndex
End Sub


Private Sub Command2_Click()
    Dim lape_c As Single
    Dim lape_m As Integer
    Dim temp As String
    Dim i As Integer

        For i = 0 To List1.ListCount - 1
    
            If List1.Selected(i) = True Then
  
                option_p = i + 1
  
                Print_ALL.For_print_lapesn(option_p).len = Val(Text1)

                lape_m = Asc(Left$(For_print_lapesn(option_p).name, 1))
        
                If Asc(lape_m) > 96 Then
                    lape_m = lape_m - 70
                Else
                    lape_m = lape_m - 64
                End If
        
                lape_c = InStr(2, For_print_lapesn(option_p).name, ".")
      
                If lape_c = 0 Then
                    m001E = Val(mID$(For_print_lapesn(option_p).name, 2, Len(For_print_lapesn(option_p).name) - 1))
                    m0022 = 0
                Else
                    m001E = Val(mID$(For_print_lapesn(option_p).name, 2, lape_c - 2))
                    m0022 = Val(mID$(For_print_lapesn(option_p).name, lape_c + 1, 1))
                End If
      
                Pitmani.List1.list(option_p - 1) = AddSpaceLeft$(5, Print_ALL.For_print_lapesn(option_p).len) & "        " & For_print_lapesn(option_p).name & Space(7) & For_print_lapesn(option_p).prof
                List_Properties_PY(lape_m, m0022) = List_Properties_PY(lape_m, m0022) - List_Properties_Length(lape_m, m0022) + Print_ALL.For_print_lapesn(option_p).len
                List_Properties_Length(lape_m, m0022) = Print_ALL.For_print_lapesn(option_p).len
  
            End If
  
        Next

End Sub


Private Sub Form_Load()
SetFont Me

'Me.Left = GetSetting(App.ProductName, "Position", Me.name & "left", Me.Left)
'Me.Top = GetSetting(App.ProductName, "Position", Me.name & "top", Me.Top)
'Me.Width = GetSetting(App.ProductName, "Position", Me.name & "width", Me.Width)
'Me.Height = GetSetting(App.ProductName, "Position", Me.name & "height", Me.Height)

Me.Caption = lng.GetResIDstring(9354)
Метка1.Caption = lng.GetResIDstring(9356)
Command1.Caption = lng.GetResIDstring(9355)
Command2.Caption = lng.GetResIDstring(9378)

Me.Combo1.AddItem lng.GetResIDstring(1129)
Me.Combo1.AddItem lng.GetResIDstring(1128)
Me.Combo1.AddItem lng.GetResIDstring(1130)
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    'Me.Width = Frame1.Width
    Метка1.Width = Me.ScaleWidth
    Combo1.Width = Me.ScaleWidth
    List1.Width = Me.ScaleWidth
    Frame1.Width = Me.ScaleWidth
    List1.Height = Me.ScaleHeight - Frame1.Height - Combo1.Height - Метка1.Height + 150
    Frame1.Top = Me.ScaleHeight - Frame1.Height '- 200
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    SaveSetting App.ProductName, "Position", Me.name & "left", Me.Left
'    SaveSetting App.ProductName, "Position", Me.name & "top", Me.Top
'    SaveSetting App.ProductName, "Position", Me.name & "width", Me.Width
'    SaveSetting App.ProductName, "Position", Me.name & "height", Me.Height
    pitmanprintunload = True
End Sub


Private Sub List1_DblClick()
    Dim lape_c As Single
    Dim lape_m As Integer
    Dim temp As String

        option_p = List1.ListIndex + 1
        temp = InputBox("Please write new length. " & vbNewLine & "For list " & For_print_lapesn(option_p).name & " and profil " & Print_ALL.For_print_lapesn(option_p).prof, , Format$(Print_ALL.For_print_lapesn(option_p).len))
        If temp = "" Then
            Exit Sub
        Else
            Print_ALL.For_print_lapesn(option_p).len = temp
        End If

        lape_m = Asc(Left$(For_print_lapesn(option_p).name, 1))
        If Asc(lape_m) > 96 Then
            lape_m = lape_m - 70
        Else
            lape_m = lape_m - 64
        End If

        lape_c = InStr(2, For_print_lapesn(option_p).name, ".")
        If lape_c = 0 Then
            m001E = Val(mID$(For_print_lapesn(option_p).name, 2, Len(For_print_lapesn(option_p).name) - 1))
            m0022 = 0
        Else
            m001E = Val(mID$(For_print_lapesn(option_p).name, 2, lape_c - 2))
            m0022 = Val(mID$(For_print_lapesn(option_p).name, lape_c + 1, 1))
        End If
  
        Pitmani.List1.list(option_p - 1) = AddSpaceLeft$(5, Print_ALL.For_print_lapesn(option_p).len) & "        " & For_print_lapesn(option_p).name & Space(7) & For_print_lapesn(option_p).prof

        List_Properties_PY(lape_m, m0022) = List_Properties_PY(lape_m, m0022) - List_Properties_Length(lape_m, m0022) + Print_ALL.For_print_lapesn(option_p).len
        List_Properties_Length(lape_m, m0022) = Print_ALL.For_print_lapesn(option_p).len

End Sub
