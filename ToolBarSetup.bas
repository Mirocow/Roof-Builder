Attribute VB_Name = "ToolBarSetup"
Option Explicit

Public Sub SetToolbarSettings()

    If Setup.Check5.value Then
    OfficeStart.Toolbar1.Visible = True
    If Setup.Check12.value Then
        SetCaprtion True
    Else
        SetCaprtion False
    End If
    'If Setup.Option4.value Then
        ToolBarSetup.Large32
    'Else
    '    ToolBarSetup.Small16
    'End If
    Else
    OfficeStart.Toolbar1.Visible = False
    End If
    
End Sub


Public Function SetCaprtion(CaptionShow As Boolean)
    Dim i As Integer
        For i = 1 To OfficeStart.Toolbar1.Buttons.count
            If CaptionShow Then
                OfficeStart.Toolbar1.Buttons(i).Caption = OfficeStart.Toolbar1.Buttons(i).Description
            Else
                OfficeStart.Toolbar1.Buttons(i).Caption = ""
            End If

        Next i

        RestoreTB
End Function



'
'LOADING LARGE ICONS FOR THE TOOLBAR:

Public Sub Large32()
    On Error Resume Next

    Dim i As Integer
    Dim n As Integer
        OfficeStart.Toolbar1.ImageList = OfficeStart.ImageList1
        For i = 1 To OfficeStart.Toolbar1.Buttons.count
            If OfficeStart.Toolbar1.Buttons(i).Description <> "" Then
                OfficeStart.Toolbar1.Buttons(i).Image = i - n
            Else
                n = n + 1
            End If

        Next i

        '
End Sub

'
'LOADING SMALL ICONS FOR THE TOOLBAR:

Public Sub Small16()
    'On Error Resume Next

    Dim i As Integer
    Dim n As Integer
        OfficeStart.Toolbar1.ImageList = OfficeStart.ImageList3
        For i = 1 To OfficeStart.Toolbar1.Buttons.count
            If OfficeStart.Toolbar1.Buttons(i).Description <> "" Then
                OfficeStart.Toolbar1.Buttons(i).Image = i - n
            Else
                n = n + 1
            End If

        Next i

        '
End Sub


'15- MENU TO SELECT A BORDER-STYLE FOR THE TOOLBAR:
'Private Function MnuTBorder(border As Boolean)
'  If border Then
'     Toolbar1.BorderStyle = ccNone
'  Else
'     Toolbar1.BorderStyle = ccFixedSingle
'  End If
'
'  Call RestoreTB
'  '
'End Function

'RESTORING THE TOOLBAR SETTINGS:

Public Sub RestoreTB()
    '
    OfficeStart.Toolbar1.RestoreToolbar "AppName", "User1", "Toolbar1"
    '
End Sub

