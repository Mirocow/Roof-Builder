VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Columns() As String

Public Function AddColumns(ByRef Data) As String
Dim i As Integer
For i = 0 To AmountColums(Data) - 1
    ReDim Preserve Columns(i)
    Columns(i) = Data(i)
Next
End Function

Public Function Column(i As Integer) As String
On Error GoTo ERR
    Column = Columns(i)
    Exit Function
ERR:
    Column = -1
End Function

Private Function AmountColums(ByRef a)
On Error GoTo ERR
    AmountColums = UBound(a) + 1
    Exit Function
ERR:
    AmountColums = -1
End Function
