VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample manager"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame mans 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   9135
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   6000
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   7440
         TabIndex        =   20
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   1680
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   18
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   16
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   14
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.VScrollBar VScroll1 
         Height          =   4750
         Left            =   8790
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   5535
         Left            =   0
         ScaleHeight     =   5475
         ScaleWidth      =   8715
         TabIndex        =   2
         Top             =   0
         Width           =   8775
         Begin VB.Image sample 
            Height          =   1695
            Index           =   15
            Left            =   5280
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   14
            Left            =   6960
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   12
            Left            =   1920
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   11
            Left            =   3600
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   10
            Left            =   5280
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   9
            Left            =   6960
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   8
            Left            =   240
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   7
            Left            =   1920
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   6
            Left            =   3600
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   5
            Left            =   1920
            Top             =   120
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   4
            Left            =   3600
            Top             =   120
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   3
            Left            =   5280
            Top             =   120
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   2
            Left            =   6960
            Top             =   120
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   1
            Left            =   240
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Image sample 
            Height          =   1695
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

