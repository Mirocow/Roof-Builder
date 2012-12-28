VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private zCaseSensitiveKeys As Boolean
Private zCount As Integer

Private Type ARRAY_TYPE
    zKey        As String   'The "key"
    zValue      As cItem  'The "value"
    zType       As Integer
End Type

Private Data() As ARRAY_TYPE

Public Property Let CaseSensitiveKeys(ByVal vData As Boolean)
    zCaseSensitiveKeys = vData
End Property

Public Property Get CaseSensitiveKeys() As Boolean
    CaseSensitiveKeys = zCaseSensitiveKeys
End Property

Private Sub Class_Initialize()
    Me.ClearArray
End Sub

Public Sub ClearArray()
    Erase Data
End Sub

Private Function IsBadVarType(ByVal vData As Variant) As Boolean

    Select Case VarType(vData)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, _
             vbString, vbBoolean, vbByte, vbDate, vbEmpty
                
        Case Is >= vbArray
            IsBadVarType = True
                
        Case Else
            IsBadVarType = True
    End Select
    
End Function

Public Sub SetValue(ByVal vValue As Variant, Optional ByVal vKey As String, Optional Replace As Boolean)
    Dim Key As Long, SearchResult As Long
    
    If IsBadVarType(vValue) Then Exit Sub
    
    SearchResult = SearchForKeyIndex(vKey)
    
    If SearchResult = 0 Then
        zCount = zCount + 1
        SearchResult = zCount
        ReDim Preserve Data(zCount)
    End If
    
    Data(SearchResult).zKey = Key
    Data(SearchResult).zValue = vValue
    Data(SearchResult).zType = VarType(vValue)

End Sub

Public Function GetValue(ByVal vKey As String) As Variant
    SearchResult = SearchForKeyIndex(vKey)
    
    If SearchResult > 0 Then
        GetValue = Data(SearchResult).zValue
    Else
        GetValue = Empty
    End If
End Function

Private Function SearchForKeyIndex(ByVal vKey As String) As Integer
    If zCount = 0 Then Exit Function
    
    Dim iLoop As Long
    
    For iLoop = 1 To UBound(Data)
        If zCaseSensitiveKeys Then
            If vKey = Data(iLoop).zKey Then Exit For
        Else
            If LCase(vKey) = LCase(Data(iLoop).zKey) Then Exit For
        End If
    Next iLoop
    
    If iLoop <= zCount Then
        SearchForKeyIndex = iLoop
    End If
        
End Function

Public Sub DeleteValue(ByVal vKey As String)
    
    SearchResult = SearchForKeyIndex(vKey)
    
    If SearchResult > 0 Then
        
        If SearchResult = zCount Then
            ReDim Preserve Data(zCount - 1)
        Else
            For iLoop = SearchResult To zCount - 1
                Data(iLoop).zKey = Data(iLoop + 1).zKey
                Data(iLoop).zValue = Data(iLoop + 1).zValue
                Data(iLoop).zType = Data(iLoop + 1).zType
            Next iLoop
            ReDim Preserve Data(zCount - 1)
        End If
        
        zCount = zCount - 1
    End If
    
End Sub

Public Function KeyExists(ByVal vKey As String) As Boolean
    If SearchForKeyIndex(vKey) > 0 Then
        KeyExists = True
    End If
End Function

Private Sub Class_Terminate()
    Me.ClearArray
End Sub
