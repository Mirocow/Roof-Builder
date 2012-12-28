VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E9335107-9E4B-4A70-BDE4-6A18106C27BA}#1.0#0"; "SplitterModern.ocx"
Begin VB.Form ProgramsData 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programs data"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin SplitterModern.SplitHV SplitHV1 
      Height          =   6255
      Left            =   6840
      TabIndex        =   8
      Top             =   360
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   11033
      SizeArea        =   1
      BackColor       =   255
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   6960
      TabIndex        =   1
      Top             =   360
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mask"
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   4095
         Begin VB.CommandButton Command4 
            Caption         =   "Searh"
            Height          =   735
            Left            =   2760
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   240
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Text            =   "Combo1"
            Top             =   600
            Width           =   2535
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11668
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Customers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Employees"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Product info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Managers"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ProgramsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Sink As New RsGridCls

Dim WK As DAO.Workspace
Dim DB As DAO.Database
Dim RS As DAO.Recordset
Dim Flds As DAO.Fields


Private Sub Form_Load()
    Dim Col As TrueDBGrid60.Column
    Dim Cols As TrueDBGrid60.Columns
    
    Dim C As Integer
    
    ' Open a database and working recordset
    On Error GoTo OpenRecSetError
    Set WK = DBEngine.Workspaces(0)
    Set DB = WK.OpenDatabase(Gl.FileName)
    Set RS = DB.OpenRecordset("select * from data", dbOpenDynaset)
    
'    Sink.Recordset = RS
    
    Set Cols = TDBGrid1.Columns
    Set Flds = RS.Fields
    
    While Cols.count
        Cols.Remove 0
    Wend

    
    Exit Sub
OpenRecSetError:
    MsgBox "Error openning Recordset!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RS.Close
    DB.Close
    WK.Close
End Sub
