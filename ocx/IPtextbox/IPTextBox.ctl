VERSION 5.00
Begin VB.UserControl IPTextBox 
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   720
   ScaleWidth      =   2580
   Begin VB.TextBox txtDot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Index           =   2
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "."
      Top             =   60
      Width           =   75
   End
   Begin VB.TextBox txtDot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Index           =   1
      Left            =   950
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "."
      Top             =   60
      Width           =   75
   End
   Begin VB.TextBox txtip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Index           =   3
      Left            =   1500
      MaxLength       =   3
      TabIndex        =   5
      Top             =   60
      Width           =   400
   End
   Begin VB.TextBox txtip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Index           =   2
      Left            =   1025
      MaxLength       =   3
      TabIndex        =   4
      Top             =   60
      Width           =   400
   End
   Begin VB.TextBox txtip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Index           =   1
      Left            =   525
      MaxLength       =   3
      TabIndex        =   3
      Top             =   60
      Width           =   400
   End
   Begin VB.TextBox txtDot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Index           =   0
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "."
      Top             =   60
      Width           =   75
   End
   Begin VB.TextBox txtip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Index           =   0
      Left            =   30
      MaxLength       =   3
      TabIndex        =   1
      Top             =   60
      Width           =   400
   End
   Begin VB.TextBox txtFra 
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2000
   End
End
Attribute VB_Name = "IPTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_IP = "127.0.0.1"
'Property Variables:
Dim m_IP As String
'Event Declarations:
Event Change() 'MappingInfo=txtip(0),txtip,0,Change
Attribute Change.VB_Description = "Êﬁ Ì « ›«ﬁ „Ì «› œ òÂ œ— „ﬁœ«—  €ÌÌ—Ì œÂÌœ"


Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_Description = "œ—»«—Â «Ì‰ ò‰ —·"
Attribute ShowAboutBox.VB_UserMemId = -552
'    dlgAbout.Show vbModal
'    Unload dlgAbout
'    Set dlgAbout = Nothing
End Sub




Private Function CalculateIP()

    If Val(txtip(0).Text) = 0 Or Val(txtip(3).Text) = 0 Then
        m_IP = ""
        Exit Function
    End If

    Dim IPStr As String
    
    IPStr = Val(txtip(0).Text) & "." & Val(txtip(1).Text) & "." & Val(txtip(2).Text) & "." & Val(txtip(3).Text)
    
    m_IP = IPStr
    
End Function

Private Function ShowIP(IPStr As String)

    If InStr(1, IPStr, ".", vbTextCompare) = 0 Then
        IP = "127.0.0.1"
        Exit Function
    End If

    Dim st As String
    Dim IPParts(3) As String
    
    st = IPStr
    
    For i = 0 To 3
        IPParts(i) = Split(st, ".")(0)
        If i <> 3 Then st = Right(st, Len(st) - Len(IPParts(i)) - 1)
        For j = 1 To Len(IPParts(i))
            If IsNumeric(mID(IPParts(i), j, 1)) = False Then
                IP = ""
                Exit Function
            End If
        Next j
    Next i
    
    If Val(IPParts(0)) = 0 Or Val(IPParts(3)) = 0 Then
        IP = ""
        Exit Function
    End If
    
    For i = 0 To 3
        txtip(i).Text = Val(IPParts(i))
    Next i

End Function



Private Sub txtDot_GotFocus(Index As Integer)
    txtip(Index + 1).SetFocus
End Sub


Private Sub txtip_Change(Index As Integer)

On Error Resume Next
   
If txtip(Index).Text <> "" Then
    For i = 1 To Len(txtip(Index).Text)
       If IsNumeric(mID(txtip(Index).Text, i, 1)) = False Then
            txtip(Index).Text = ""
            Exit Sub
       End If
    Next i
End If
If Val(txtip(Index).Text) > 255 Then txtip(Index).Text = 255
If Index < 3 And Len(txtip(Index).Text) = 3 Then txtip(Index + 1).SetFocus

CalculateIP

 RaiseEvent Change

End Sub


Private Sub txtip_GotFocus(Index As Integer)

    txtip(Index).SelStart = 0
    txtip(Index).SelLength = Len(txtip(Index).Text)

End Sub

Private Sub txtip_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If (KeyCode = 8 Or KeyCode = 37 Or KeyCode = 38) And Index <> 0 Then
            If txtip(Index).SelStart = 0 Then txtip(Index - 1).SetFocus
    End If
    
    If (KeyCode = 39 Or KeyCode = 40 Or KeyCode = 46) And Index <> 3 Then
            If txtip(Index).SelStart = Len(txtip(Index).Text) Then txtip(Index + 1).SetFocus
    End If

End Sub



Private Sub UserControl_Resize()
Height = 375
Width = 2010
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtip(0),txtip,0,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "„‘Œ’ „Ì ò‰œ òÂ «Ì‰ ‘Ì ›⁄«· »«‘œ Ì« ŒÌ—"
    Enabled = txtip(0).Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    For i = 0 To 3
        txtip(i).Enabled() = New_Enabled
    Next i
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtip(0),txtip,0,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "œ— ’Ê—  «‰ Œ«» «Ì‰ ê“Ì‰Â «Ì‰ ‘Ì —ÊÌ ›—„ ﬁ›· „Ì ‘Êœ"
    Locked = txtip(0).Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    For i = 0 To 3
        txtip(i).Locked() = New_Locked
    Next i
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtip(0),txtip,0,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "—«”  »Â çÅ «Ì‰ ‘Ì"
    RightToLeft = txtip(0).RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    For i = 0 To 3
        txtip(i).RightToLeft() = New_RightToLeft
    Next i
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFra,txtFra,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "‘—Õ òÊçò ‰„«Ì‘ œ«œÂ ‘œÂ Êﬁ Ì òÂ „«Ê” —ÊÌ «Ì‰ ‘Ì «” "
    ToolTipText = txtFra.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtFra.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get IP() As String
Attribute IP.VB_Description = "„ﬁœ«— ⁄œœ ¬Ì ÅÌ ¬œ—”"
    IP = m_IP
End Property

Public Property Let IP(ByVal New_IP As String)
    m_IP = New_IP
    PropertyChanged "IP"
    ShowIP New_IP
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_IP = m_def_IP
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    For i = 0 To 3
        txtip(i).Enabled = PropBag.ReadProperty("Enabled", True)
        txtip(i).Locked = PropBag.ReadProperty("Locked", False)
        txtip(i).RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    Next i
    txtFra.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_IP = PropBag.ReadProperty("IP", m_def_IP)
    
End Sub

Private Sub UserControl_Show()
ShowIP IP
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", txtip(0).Enabled, True)
    Call PropBag.WriteProperty("Locked", txtip(0).Locked, False)
    Call PropBag.WriteProperty("RightToLeft", txtip(0).RightToLeft, False)
    Call PropBag.WriteProperty("ToolTipText", txtFra.ToolTipText, "")
    Call PropBag.WriteProperty("IP", m_IP, m_def_IP)
End Sub

