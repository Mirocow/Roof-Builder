VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
'   cf.FOpen App.Path & "\test.bin", 0
'   cf.FOpen App.Path & "\test.bin", 1
'   cf.FOpen App.Path & "\test.bin", 2
'   cf.FWrite clng(1)
'   cf.FRead lng
'   cf.FSeek
'   cf.FSeek 5
'   cf.FWriteList clng(2),cstr("ferhgrtjh fdgdfh")
'   array = cf.FReadList(2)
'   cf.FWriteString str, 0-long
'   cf.FWriteString str, 1-single
'   cf.FReadString f8, 1
'   cf.FClose

'    Dim iData() As Byte
'    ReDim iData(5, 2)
'
'    iData(0, 0) = 45
'    iData(1, 0) = 34
'    iData(2, 0) = 33
'    iData(3, 0) = 33
'    iData(4, 0) = 37
'
'    iData(1, 1, 0) = 145
'    iData(1, 2, 1) = 134
'    iData(1, 1, 2) = 133
'
'    iData(2, 1, 0) = 245
'    iData(2, 2, 1) = 234
'    iData(2, 1, 2) = 233
'
'    iData(1, 2, 1) = 245
'    iData(1, 2, 1) = 234
'    iData(1, 2, 1) = 233
    
    
'    MsgBox UBound(iData, 1)

'    ScaleLeft_Main = 0
'    ScaleWidth_Main = -1005
'    ScaleTop_Main = -1.256
'    ScaleHeight_Main = -156.26
    
'    Dim cf As FileMan.clsFile
'    Set cf = New clsFile
'
'    cf.FOpen App.Path & "\test.bin", aWrite, True
'
'    cf.FWriteArray iData, 1
    
'    cf.FWriteString "Test of data string save for binary file", vSingle
'    cf.FWrite CSng(13)
'    cf.FWriteString "", 1
'    cf.FWrite ScaleLeft_Main
'    cf.FWrite ScaleWidth_Main
'    cf.FWrite ScaleTop_Main
'    cf.FWrite ScaleHeight_Main
'    cf.FWrite CLng(1564455)
'    cf.FWrite CInt(1564)
'    cf.FWrite CDbl(451864641867#)
'    cf.FClose
'
'    Erase iData
    
'    Dim f0 As Byte
'    Dim f1 As Long
'    Dim f2 As Integer
'    Dim f3 As Single, f7 As Single
'    Dim f5 As Double
'    Dim f6 As Boolean
'    Dim f8 As String ', f8 As String
    
'    Dim flist(), flist1() As Variant
    
'    cf.FOpen App.Path & "\test.bin", aRead
'
'    cf.FReadArray iData, 1
    
'    cf.FReadString f8, vSingle
'    cf.FRead f3
'    cf.FReadString f8, 1
'    cf.FRead ScaleLeft_Main
'    cf.FRead ScaleWidth_Main
'    cf.FRead ScaleTop_Main
'    cf.FRead ScaleHeight_Main
'    cf.FRead f1
'    cf.FRead f2
'    cf.FRead f5
'    cf.FClose
'
'    Set cf = Nothing

    Dim cf As New clsFile
    If cf.FOpen("o:\tmp\~tmp.rfd", 1) Then
        If cf.FN = 0 Then GoTo ERR
        If cf.FLOF() = 0 Then GoTo ERR
    End If
    Set cf = Nothing
    
    Dim cf1 As New clsFile
    If cf1.FOpen("o:\tmp\~tmp1.rfd", 1) Then
        If cf1.FN = 0 Then GoTo ERR
        If cf1.FLOF() = 0 Then GoTo ERR
    End If
    Set cf1 = Nothing
    

'    Dim iData() As Variant
'    ReDim iData(5, 1)
'
'    iData(0, 0) = CLng(45)
'    iData(1, 0) = CDbl(34)
'    iData(2, 0) = CInt(33)
'    iData(3, 0) = CByte(33)
'    iData(4, 0) = 37
'    iData(4, 1) = True
'
'    Dim cm As New iStorage
'    cm.Driver = clsHash
'    cm.Add CByte(0), 1
'    cm.Add CLng(123), 2
'    cm.Add CByte(23), 3
'    cm.Add iData, 4
'
'    Dim cData() As Variant
'    cData = Array(1, 2.3, 45, 6.7, 77, 0.0001)
'
'    cm.Add cData, 5
'    cm.Add iData, 6
'    cm.Add iData, 7
'    cm.Add iData, 8
'    cm.Add iData, 9
'
'    cm.FileName = App.Path & "\test.txt"
'    Debug.Print "Count: " & cm.Count
'    Debug.Print "Save: " & cm.Save
'    cm.Clear
'
'    Debug.Print "Read: " & cm.Read
'    Debug.Print "Read items: " & cm.Count
'
'    Dim arr() As Variant
'    arr = cm.items
'    'Debug.Print "Get: " & cm.Lookup(1, arr)
'    Debug.Print "Size item 4: " & UBound(arr(4))
'    Debug.Print "Type item 4: " & cm.GetType(4)
'
'    Set cm = Nothing
    Exit Sub
ERR:
    Debug.Print ERR.Description
End Sub

