Attribute VB_Name = "ItemComboBox"
Private Columns() As String

Public Function AddColumns(ByRef Data) As String
Dim i As Integer
For i = 0 To AmountColums(Data)
    ReDim Preserve Columns(i)
    Columns(i) = Data(i)
Next
End Function

Public Function Column(i As Integer) As String
On Error GoTo ERR
    Column = Columns(i)
    Exit Function
ERR:
    AmountColums = -1
End Function

Private Function AmountColums(ByRef a)
On Error GoTo ERR
    AmountColums = UBound(a) + 1
    Exit Function
ERR:
    AmountColums = -1
End Function
