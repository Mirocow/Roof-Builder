VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
' clsHash - An array class that hashes the keys for quick access.
'--------------------------------------------------------------------------------
Option Explicit

Implements IContainer

Private Enum enumState
    stEmpty
    stDeleted
    stInUse
End Enum


Private Type NameValue
    State As enumState
    Key As Long
    Value As Variant
End Type


Private mData() As NameValue
Private mMaxLoad As Long
Private mCnt As Long
Private Const module As String = "clsHash"
Private mCtr As Long


'$ PROBHIDE NO_ERROR_HANDLER BEGIN

Private Sub Class_Initialize()
    ' Set the initial table size to 8
    mCtr = 1
    InitializeTable 8
End Sub


Private Sub InitializeTable(ByVal InitialSize As Long)
    'Initialize the table to a certain size
    
    If InitialSize < 8 Then
        InitialSize = 8
    End If
    InitialSize = InitialSize * 1.25
    InitialSize = NextPrime(InitialSize)
    mMaxLoad = InitialSize * 0.8
    mCnt = 0
    mCtr = 1
    Erase mData
    ReDim mData(-1 To InitialSize)
End Sub


Private Function NextPrime(ByVal NumberDesired As Long) As Long
    'Return the next prime number past a certain number.
    
    Dim i As Long
    Dim HalfwayPoint As Long
    Dim IsDivisible As Boolean
    
    Do
        NumberDesired = NumberDesired + 1
        IsDivisible = False
        HalfwayPoint = NumberDesired \ 2
        For i = 2 To HalfwayPoint
            If NumberDesired Mod i = 0 Then
                IsDivisible = True
                Exit For
            End If
        Next i
        If IsDivisible = False Then
            Exit Do
        End If
    Loop
    NextPrime = NumberDesired

End Function


Private Sub Expand()
    'Expand the hash table to accommodate more entries.

    Dim i As Long
    Dim OldRows() As NameValue
    Dim OldMaxRows As Long
    
    'Store the old rows
    OldMaxRows = UBound(mData, 1)
    ReDim OldRows(0 To OldMaxRows)
    For i = 0 To OldMaxRows
        OldRows(i) = mData(i)
    Next i
    
    'Reinitialize the table
    InitializeTable OldMaxRows * 2
    
    ' Add the data from the old rows
    For i = 0 To OldMaxRows
        If OldRows(i).State = stInUse Then
            IContainer_Add OldRows(i).Value, OldRows(i).Key
        End If
    Next i

    ' Free the old table rows
    Erase OldRows

End Sub


Private Function FindSlot(ByVal Key As Long, ByVal FindExisting As Boolean) As Long
    ' Find a slot (filled or blank) in the array
    Dim i As Long
    Dim max As Long
    Dim Hash As Long
    
    If FindExisting = True Then
        ' Find existing key and leave early
        max = UBound(mData, 1)
        Hash = Key Mod max
        For i = 0 To max
            If mData(Hash).State = stInUse Then
                If mData(Hash).Key = Key Then
                    FindSlot = Hash
                    Exit Function
                End If
            ElseIf mData(Hash).State = stEmpty Then
                FindSlot = -2
                Exit Function
            End If
            Hash = (Hash + 1) Mod max
        Next i
        MsgBox "You should never reach this spot!", vbCritical
        Exit Function
    End If
        
    ' Find blank slot
    max = UBound(mData, 1)
    Hash = Key Mod max
    For i = 0 To max
        If mData(Hash).State <> stInUse Then
            FindSlot = Hash
            Exit Function
        End If
        Hash = (Hash + 1) Mod max
    Next i
    MsgBox "You should never get here!", vbCritical
    
End Function


Private Function IContainer_Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an item to the table
    Dim i As Long
    
    ' Assign a key
    If Key = 0 Then
        Key = mCtr
        mCtr = mCtr + 1
    End If
    
    ' Make sure key doesn't already exist
    i = FindSlot(Key, True)
    If i >= 0 Then
        IContainer_Add = 0
        Exit Function
    End If
    
    ' Enlarge table if necessary
    If mCnt > mMaxLoad Then
        Expand
    End If
    
    ' Set the value
    i = FindSlot(Key, False)
    With mData(i)
        .State = stInUse
        .Key = Key
        If IsObject(Data) Then
            Set .Value = Data
        Else
            .Value = Data
        End If
    End With
    
    mCnt = mCnt + 1
    
    IContainer_Add = Key
    

End Function


Private Sub IContainer_Clear()
    ' Remove all items from the table
    InitializeTable 8
End Sub


Private Function IContainer_Count() As Long
    ' Return the number of items in the table
    IContainer_Count = mCnt
End Function


Private Function IContainer_Exists(ByVal Key As Long) As Boolean
    ' Return whether or not an item exists in the container
    Dim i As Long
    i = FindSlot(Key, True)
    If i < 0 Then
        IContainer_Exists = False
    Else
        IContainer_Exists = True
    End If
End Function


Private Property Let IContainer_Item(ByVal Key As Long, ByVal Value As Variant)
    ' Set an item in the container
    Dim i As Long
    
    i = FindSlot(Key, True)
    
    If i < 0 Then
        IContainer_Add Value, Key
    Else
        With mData(i)
            If IsObject(Value) Then
                Set .Value = Value
            Else
                .Value = Value
            End If
        End With
    End If
End Property


Private Property Get IContainer_Item(ByVal Key As Long) As Variant
    ' Get an item from the container
    Dim i As Long
    i = FindSlot(Key, True)
    If i < 0 Then
        'Skip it
    ElseIf IsObject(mData(i).Value) Then
        Set IContainer_Item = mData(i).Value
    Else
        IContainer_Item = mData(i).Value
    End If
End Property


Private Function IContainer_Keys() As Long()
    ' Return an array of all the keys in the container
    Dim arr() As Long
    Dim i As Long
    Dim max As Long
    Dim ctr As Long
    
    ReDim arr(-1 To mCnt - 1)
    max = UBound(mData, 1)
    For i = 0 To max
        If mData(i).State = stInUse Then
            arr(ctr) = mData(i).Key
            ctr = ctr + 1
        End If
    Next i
    IContainer_Keys = arr
End Function

Private Function IContainer_items() As Variant()
    ' Return an array of all the keys in the container
    Dim arr() As Variant
    Dim i As Long
    Dim max As Long
    Dim ctr As Long
    
    ReDim arr(1 To mCnt)
    max = UBound(mData, 1) - 1
    For i = 1 To max
        If mData(i).State = stInUse Then
            ctr = ctr + 1
            arr(ctr) = mData(i).Value
        End If
    Next i
    IContainer_items = arr
End Function

Private Function IContainer_Lookup(ByVal Key As Long, ReturnData As Variant) As Boolean
    ' Get an item from the container
    Dim i As Long
        
    ' Find key
    i = FindSlot(Key, True)
    If i < 0 Then
        ' Key not found
        IContainer_Lookup = False
    Else
        ' Return entry
        IContainer_Lookup = True
        If IsObject(mData(i).Value) Then
            If IsObject(ReturnData) Then
                Set ReturnData = mData(i).Value
            Else
                MsgBox "The data is an object, and you passed in a non-object variable! The data was not set."
            End If
        Else
            If IsObject(ReturnData) Then
                MsgBox "The data is not an object, and you passed in an object variable! The data was not set."
            Else
                ReturnData = mData(i).Value
            End If
        End If
    End If
End Function


Private Property Get IContainer_Name() As String
    ' Return the name of the container
    IContainer_Name = module
End Property


Private Function IContainer_Remove(ByVal Key As Long) As Boolean
    ' Remove an item from the container
    Dim i As Long
    i = FindSlot(Key, True)
    If i < 0 Then
        IContainer_Remove = False
    Else
        IContainer_Remove = True
        With mData(i)
            .State = stDeleted
            .Key = 0
            .Value = 0
        End With
        mCnt = mCnt - 1
    End If
End Function


'$ END
