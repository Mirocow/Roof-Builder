Attribute VB_Name = "GL"
Public Function ShrinkArray_long(ByRef nArr() As Integer, ByVal nIndex As Long)
        If nIndex < LBound(nArr) Or nIndex > UBound(nArr) Then
            ERR.Raise 10, , "������ ����� ������?"
        Else
            If UBound(nArr) >= nIndex + 1 Then
            '������� ��� ��������
            CopyMemory VarPtr(nArr(nIndex)), VarPtr(nArr(nIndex + 1)), (UBound(nArr) - nIndex) * 4  '4 � ����� Long-����������
            End If
            '��������� ���������
            ReDim Preserve nArr(UBound(nArr) - 1)
        End If
End Function
