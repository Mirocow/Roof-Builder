VERSION 5.00
Object = "{2635FD45-668B-432A-8A79-3D3CF73A0077}#1.0#0"; "ÑhameleonButton.ocx"

Begin VB.Form frmEvent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exception Handler"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   7335
      TabIndex        =   3
      Top             =   900
      Width           =   7335
      Begin VB.CheckBox chkReport 
         Caption         =   "Report this error message"
         Height          =   225
         Left            =   150
         TabIndex        =   11
         Top             =   3180
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CommandButton cmdStatus 
         Caption         =   "Exit Now"
         Height          =   330
         Index           =   1
         Left            =   6030
         TabIndex        =   10
         Top             =   3390
         Width           =   1095
      End
      Begin VB.CommandButton cmdStatus 
         Caption         =   "Continue"
         Height          =   330
         Index           =   0
         Left            =   4800
         TabIndex        =   6
         Top             =   3390
         Width           =   1095
      End
      Begin VB.CheckBox chkReStart 
         Caption         =   "On Exit, Auto Restart the application"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox txtData 
         BackColor       =   &H8000000F&
         Height          =   1425
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   960
         Width           =   7005
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Height          =   645
         Left            =   120
         TabIndex        =   9
         Top             =   60
         Width           =   7095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   135
         TabIndex        =   8
         Top             =   2520
         Width           =   6930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Exception Information:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   750
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   35
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "An Exception Error has Occured"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exception Handler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   210
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStatus_Click(Index As Integer)

Dim frm     As Form

On Error GoTo Handler

    Select Case Index
    Case 0
        '//send an error report
        If chkReport.value = 1 Then
            With cLog
                .Report_Mail Me, txtData
            End With
        End If
        Unload Me
        
    Case 1
        If chkReport.value = 1 Then
            With cLog
                .Report_Mail Me, txtData
            End With
        End If
        
        '//restart application
        If chkReStart.value = 1 Then
            Application_Restart
        Else
            For Each frm In Forms
                Unload frm
            Next frm
            GoTo Handler
        End If
    End Select

Exit Sub

Handler:
    End
    
End Sub

Public Sub User_Message(ByVal sMessage As String)

    lblMessage.Caption = sMessage
    
End Sub

Private Sub Form_Load()

Dim sMessage As String

    sMessage = "A serious Error has occured in " & App.Title & _
    " You can choose to try and continue, or exit and optionally restart the application."
    User_Message sMessage
    
End Sub

