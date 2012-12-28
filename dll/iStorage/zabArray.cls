VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zabArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================
'  What is a zabArray
' =============================================================
'
' Some of you folks who are multi-lingual in the development
' world may have come accross a CPP/PHP concept call Serialize,
' and the sister concept of UnSerialize.  For those who don't,
' allow me to help you out a little:
'
' In PHP, arrays can contain any type of data.  Automatically.
' You almost never have to worry about "Type Mismatch" because
' of something dumb like adding 5 to "Hello".  PHP also handles
' Collections and Arrays the same way.  Even interchangeably.
' You can have a legal array in PHP that looks like this:
'
'           Action[1] = "Jump"
'           Action[2] = "Hide"
'           Action[3] = "Scream"
'           Action['Last'] = 3
'           Action['Next'] = 2
'
' And PHP doesn't care.  This would be a glorious addition to
' VB.  Now now, you in the back with their hand up.  What's that
' you say?  You can do almost all that in VB with a Collection?
' Well that would work fine in dandy, but what if you could
' convert your entire collection to a text string and store it
' in, let's say, an INI file.  Or a text value in the registry.
'
' I cannot programmatically mimic a collection here.  So I did
' what I could to emulate it and add a couple things that are
' missing.
'
' You can add key/value pairs to the zabArray.  You can iterate
' through them with For..Next.  The values can be any of the
' numeric or string data types (except vbDecimal).  You can swap
' values between two keys.  Most importantly, you can dump the
' entire array into a single line string perfect for text files
' or GetSetting/SaveSetting.
'
' Quick Tip:  Make a zabArray to store your two-dimensional
' data.  Then save the Serial string in another zabArray for
' a third dimension.  etc.
'
' With zabArrays, you can communicate with your app's object
' very easily.  Want to make a child form with options that
' affect the parent form?  Give it a "Serial" property.  The
' parent can Unserialize the data when the child unloads.  Have
' a DDE project and you want to send a command?  Great.  Save
' time- and error-prone transmissions with a quick zabArray.
'
' If you sat down to ponder it, there are many many many uses
' for a zabArray.  Comments below.
'
' =============================================================


' =============================================================
' WARNING: The moment you change these three values, whatever
' arrays you have made previously become invalid.
Private DATA_SEPARATOR    'This separates individual
                                    'array tokens.  Must be <64
Private FIELD_SEPARATOR  'This separates ID fields
                                    'must be <64
Private ARRAY_TAG    'Version ID
        ' The version ID allows forward compatibility.  If you
        ' make a new zabArray class, you should set it's tag
        ' to something else.  That makes whatever serial string
        ' it produces unique.  If you attempt to UnSerialize a
        ' string that was Serialized with a different tag, then
        ' the script should either import or fail.
' =============================================================


Private zCaseSensitiveKeys As Boolean   'Nice option that will force
                                        'the search function to pay
                                        'attention to case.
                                        
Private zCount As Integer   'The current number of items present in
                            'the zabArray.  Could be redundant, but
                            'it is slightly faster than Ubound().
                                        

' The key/value pair type array.  Many places in the class utilize this
' block, so if you add something here, make sure you hit all occurrences.
Private Type ARRAY_TYPE
    zKey        As String   'The "key"
    zValue      As Variant  'The "value"
End Type

' We don't predeclare any elements here.  Each time an element is
' added, we will redim the array.  Reverse effect when an element is
' deleted.  The array goes from 0 to Count.  Item 0 is not used for
' user data storage.  So logically, 1 to Count is the useful data.
Private data() As ARRAY_TYPE


Private Sub Class_Initialize()
    DATA_SEPARATOR = Chr(1)  'This separates individual
                             'array tokens.  Must be <64
    FIELD_SEPARATOR = Chr(0) 'This separates ID fields
                             'must be <64
    ARRAY_TAG = "CF01"    'Version ID
    'ClearArray simply ReDim's the Array at 0 elements without a
    'preserve clause.  Seems like the fastest way.
    Me.ClearArray
End Sub


Public Sub ClearArray()
    ' Initialize the Data to 0 elements.  The 0 element is used
    ' as a buffer for swapping data so I don't have to declare
    ' extra temporary variables on the side.  Since the 0 element
    ' doesn't get searched when a value is requested, we don't
    ' even have to delete it's contents.  For what it's worth it
    ' adds a little bit of speed.
    ReDim data(0)
End Sub





Public Sub SetValue(ByVal vKey As String, ByVal vValue As Variant)
' Sets a key/value pair in the array.  Simple function.  I should
' note at this moment that this class will do everything it can to
' avoid interrupting the application flow.  eg no errors will be
' purposely raised.  The default response to error is to just exit.
'
    
    ' Acceptable VarTypes are checked in the IsBadVarType function.
    If IsBadVarType(vValue) Then Exit Sub

    ' See if the key is in the database.  SearchForKeyIndex drops the
    ' physical index of the Key/Value.  Since we don't store data at
    ' element 0, then a 0 return would mean "not found."
    SearchResult = SearchForKeyIndex(vKey)
        
    If SearchResult = 0 Then
            'Didn't find anything.  We need to redim the array
            'and add it to the end.  No sorting is done.
            
            zCount = zCount + 1     'Such a loyal counter, eh?
            SearchResult = zCount   'We need this to happen
            ReDim Preserve data(zCount) 'Redim it.
    End If
    
    ' Well, we have a place to save the data.  One way or another
    ' "SearchResult" is now ready to occupy.  Call the movers.
    data(SearchResult).zKey = vKey
    data(SearchResult).zValue = vValue
    
    
End Sub




Public Function GetValue(ByVal vKey As String) As Variant
    ' Returns the value filed under "key" or empty.
    
    SearchResult = SearchForKeyIndex(vKey)
    
    If SearchResult > 0 Then
        GetValue = data(SearchResult).zValue
    Else
        GetValue = Empty 'NULL Sucks
    End If
        
End Function






Public Sub SetItemKey(ByVal vItem As Integer, ByVal vKey As String)
    ' This allows you to set a Key by index.  Use this if you want
    ' to iterate through the array.  Will not allow new items to be
    ' created, so you will have to do something like:
    '
    '       For i = 1 to 10: Data.SetValue trim(str(i)), "": next i
    '
    ' And then do what you need to do.
    
    If vItem < 1 Or vItem > zCount Then Exit Sub 'Testing, 1, 2, 3
    
    data(vItem).zKey = vKey 'Duh
End Sub


Public Sub SetItemValue(ByVal vItem As Integer, ByVal vValue As Variant)
    ' Basically the same concept as SetItemKey above.  The primary
    ' difference being in the VarType checking...
    
    If vItem < 1 Or vItem > zCount Then Exit Sub 'Pop-quiz
    
    If IsBadVarType(vValue) Then Exit Sub   'Well, is it?
    
    data(vItem).zValue = vValue 'Duh
End Sub


Public Function GetItemKey(ByVal vItem As Integer) As String
    ' Reverse of the SetItemKey.  This returns the Key based on
    ' an index.  Iteration Rules!
    If vItem < 1 Or vItem > zCount Then Exit Function
    GetItemKey = data(vItem).zKey
End Function

Public Function GetItemValue(ByVal vItem As Integer) As Variant
    ' Also a similar concept to the above function, except where noted.
    
    If vItem < 1 Or vItem > zCount Then
            ' I do not like returning NULL because that basically says
            ' Invalid when that may not be the case.  Empty at least
            ' converts in some instances to another type.  Better be
            ' safe than wonder later on...
            GetItemValue = Empty
            Exit Function    'Get Out!
    Else
            GetItemValue = data(vItem).zValue   'Looks good.
    End If
    
End Function







Public Sub DeleteValue(ByVal vKey As String)
    ' This deletes a key.  By saying that, I am generally implying
    ' that if you call this procedure, you could in effect delete
    ' a key.  This is similar to erasing, but heavier in syntax.
    
    SearchResult = SearchForKeyIndex(vKey)  'Why delete it if it
                                            'doesn't exist already?
    
    If SearchResult > 0 Then
        ' What we want to do is find out where the key is in the
        ' array.  Then we want to collapse the stack on to the key
        ' so that whatever order is maintained.  An alternative
        ' would be to take the last item and move it to the deleted
        ' slot, but this will jumble things.  If we add a sorting
        ' routine later, we wouldn't want some lowly deleter messing
        ' up our lovely collating.
        
        If SearchResult = zCount Then
                ' It's at the end.  Might as well shave off the last
                ' element and get it over with.
                ReDim Preserve data(zCount - 1)
        Else
            ' It's not at the end.  Start at the item and move the next
            ' one into the current spot.  Then move to the next spot.
            ' Eventually you hit the end and you shave off the excess.
            For iLoop = SearchResult To zCount - 1
                data(iLoop).zKey = data(iLoop + 1).zKey
                data(iLoop).zValue = data(iLoop + 1).zValue
            Next iLoop
            ReDim Preserve data(zCount - 1) 'Gillete
        End If
        
        zCount = zCount - 1 'Since we did some shaving, might as well
                            'update the array counter.
    End If
    
End Sub



Public Function SwapValues(ByVal vKey1 As String, ByVal vKey2 As String)
    ' Back in my QB/VBDOS days, a sometime handy function existed called
    ' SWAP.  I liked this function but can't seem to find a built-in
    ' version for todays VB.  Just in case somebody could use it I took
    ' the time to code one.
    '
    ' I honestly cannot tell you why VB doesn't have a Swap
    ' function anymore.  It is way beyond me...
    
    
    ' We should at least make sure both keys even exist.  If we don't
    ' it's not worth swapping anything.
    If SearchForKeyIndex(vKey1) <= 0 Then Exit Function
    If SearchForKeyIndex(vKey2) <= 0 Then Exit Function
    
    ' These next few lines are pretty self explanatory.  Remember that
    ' element 0 buffer I mentioned.  Here it is.
    data(0).zKey = data(vKey1).zKey
    data(0).zValue = data(vKey1).zValue
    
    data(vKey1).zKey = data(vKey2).zKey
    data(vKey1).zValue = data(vKey2).zValue
    
    data(vKey2).zKey = data(0).zKey
    data(vKey2).zValue = data(0).zValue
    
End Function



Public Function KeyExists(ByVal vKey As String) As Boolean
    'Wrapper function for the data search function.  The physical index
    'is redundant if you know the key, so why bother.  This seemed a bit
    'more logical than any other method.
    
    If SearchForKeyIndex(vKey) > 0 Then
        KeyExists = True
    End If
End Function



Private Function IsBadVarType(ByVal vData As Variant) As Boolean
    ' This takes a variant and gives a yay or nay on whether we
    ' want to store it in the zabArray.  Some vartypes do not
    ' convert to string and back very well.  I want to make this
    ' code as friendly as possible, I here is where I have blocked
    ' these bad boys.
    
    Select Case VarType(vData)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, _
             vbString, vbBoolean, vbByte, vbDate, vbEmpty
                ' OK Data.  This is the 'approved list' or, in
                ' other words, the list of data that converts easily
                ' into a string and back.  Different OS's tend to
                ' act differently with DataObject, Error, et al so
                ' I didn't even want to bother with them.  This list
                ' should take care of 99 percent of the problems.
                
        Case Is >= vbArray
            ' We are not implementing arrays at this time.  It requires
            ' us to nest the class in ever-deeper layers.  This is far
            ' beyond the scope of the zabArray.
            '
            ' I include this case separately with the other bad apples
            ' because you may want to enable support later.
            IsBadVarType = True
                
        Case Else
            ' Bad stuff.  We don't wanna play with the other stuff
            ' because we don't know how it will act with different
            ' OS's, VB's, and Service Packs.
            IsBadVarType = True
    End Select
    
End Function


Private Function SearchForKeyIndex(ByVal vKey As String) As Integer
    ' Our wonderful built-in search function.  Returns a physical
    ' index if the key is found or 0 if not.
        
    If zCount = 0 Then Exit Function 'Wanna search nothing?
    
    For iLoop = 1 To UBound(data)
        'Iterate through the keys and values.  We are looking for
        'a key match.  If you have set the CaseSensitiveKeys to
        'TRUE, then here is where it is honored.
        If zCaseSensitiveKeys Then
            If vKey = data(iLoop).zKey Then Exit For
        Else
            If LCase(vKey) = LCase(data(iLoop).zKey) Then Exit For
        End If
    Next iLoop
    
    ' When the final Next is processed, the loop variable is equal to
    ' one step above the target value.  In this case, iLoop is now
    ' equal to Ubound(Data) + 1.  So instead of dragging an extra
    ' "success" variable around, we just do a simple test.
    If iLoop <= zCount Then
        SearchForKeyIndex = iLoop
    End If
        
End Function


Public Function GetSerial() As String
    ' It is so much easier to make a serial than it is to
    ' break it apart.  Sheesh!
    Dim KeyTokens As String 'The keys are stored separately from
                            'the values.  Makes it easier to index
                            'later, especially if you read the data
                            'in binary format.
    
    Dim ValSizes As String  'More binary pre-planning. This stores the
                            'length of each encoded value.  Unused in
                            'in this class, but maybe later...
    
    Dim ValTokens As String 'The actual encoded values.
    
    Dim CurToken As String  'The current value token.  Used by ValSizes
                            'to make a size encoding before aded to the
                            'ValTokens string.
    
    Dim DataHash As String  'A checksum for the data.  This makes sure
                            'that the data keeps it's integrity.
    
    Dim vOutput As String 'The total output item.
    
    
    ' There are 6 parts to this version of Serial/UnSerial.  The first
    ' is a version tag.  The second is the count of the keys and values.
    ' the third is the Keys, tokenized and encoded.  The fourth is a
    ' list of length sizes for each tokenized value.  Fifth we have the
    ' tokenized values, also encoded, and lastly a simple hash to make
    ' sure nothing gets corrupted.
    
    ' Each field in the serial array is separated with a field separator
    ' that tells the UnSerializer where to break things apart.  The
    ' tokenizer has a similar data separator for the tokens.
    
    
    'Array tag and count
    vOutput = ARRAY_TAG & FIELD_SEPARATOR
    vOutput = vOutput & Trim(Str(zCount)) & FIELD_SEPARATOR
       
    
    For iLoop = 1 To zCount
            'Now we go through each key/value pair and tokenize the
            'data.  Since nothing is being written to the output yet
            'we don't care if there is no data.
            
            'These are the keys
            KeyTokens = KeyTokens & MakeToken(data(iLoop).zKey)
        
            'This is the current value
            CurToken = MakeToken(data(iLoop).zValue)
            
            'This is the len() of the current value
            ValSizes = ValSizes & Trim(Str(Len(CurToken)))
            
            'These are all the values
            ValTokens = ValTokens & CurToken
    
            If iLoop < zCount Then
                'If we still have more to go then add a data separator
                'to tell the Unserializer where to start breaking.
                KeyTokens = KeyTokens & DATA_SEPARATOR
                ValSizes = ValSizes & DATA_SEPARATOR
                ValTokens = ValTokens & DATA_SEPARATOR
            End If
    Next iLoop
    
    ' --------------------------------------------------
    ' Here you would produce a DataHash.  Its very simple
    ' to create a checksum program, but lot's of code,
    ' and I want the zabArray to be as small as possible.
    '
    ' Consider making a hash class.  You have 3 strings
    ' to hash (KeyTokens,ValSizes,ValTokens) and you will
    ' need to check the hashes at UnSerialize.
    ' --------------------------------------------------
    
    'Finish the output variable.
    vOutput = vOutput & KeyTokens & FIELD_SEPARATOR
    vOutput = vOutput & ValSizes & FIELD_SEPARATOR
    vOutput = vOutput & ValTokens & FIELD_SEPARATOR
    vOutput = vOutput & DataHash
    
    GetSerial = vOutput
End Function





Public Function SetSerial(ByVal vData As String) As Boolean
    ' This is way more difficult.  The Split() function is used
    ' a bunch here.  I do not trust this function.  Sometimes
    ' it will fail when there is no reason for it to do so.
    '
    ' Nevertheless, this function takes your premade serial
    ' string and populates the zabArray with the data.
    '
    
    Dim Fields As Variant
    Fields = Split(vData, FIELD_SEPARATOR, 6) 'Yuck *cringe*
        
    If UBound(Fields) < 5 Then
        ' Probably a bad zabArray.  Should be 6 elements
        ' pulled from the Serial string...
        Exit Function
    End If
    
    
    ' Check Array Tag
    Select Case Fields(0)
            Case ARRAY_TAG
                    ' The serial matches this class.  We can process it
                    ' as normal, provided everything else checks out...
            
            Case "zab2"
                    ' This is just an example of compatibility.  If we
                    ' can read 'zab2' serials then put it here, along
                    ' with any considerations we should be having...
            
            Case Else
                    ' Can't read it or the tag is wrong.
                    Exit Function
    End Select

    
    Dim KeyTokens As Variant    'Duh
    Dim ValSizes As Variant     'Duh
    Dim ValTokens As Variant    'Duh
    
    KeyTokens = Split(Fields(2), DATA_SEPARATOR)    'ouch
    ValSizes = Split(Fields(3), DATA_SEPARATOR)     'oooh
    ValTokens = Split(Fields(4), DATA_SEPARATOR)    'eek
    
    ' Check to make sure we have equal keys and values
    If UBound(KeyTokens) <> UBound(ValTokens) Then
            ' We have a bias between keys or values.  Since we can't
            ' tell where we are missing data from, we should just
            ' give up on this.
            Exit Function
    End If
       
    
    Me.ClearArray   'Clear us.  New data is here.
    For iLoop = 0 To UBound(KeyTokens)
        'Uses our public set value.  If you can't use your own code
        'then something is wrong with it.
        Me.SetValue UnMakeToken(KeyTokens(iLoop)), _
                    UnMakeToken(ValTokens(iLoop))
    Next iLoop
        
    '
    ' Check Fields(5) for a valid hash
    '
    
    SetSerial = True 'If we made it this far than it's good.
        
End Function







Private Function MakeToken(ByVal vData As Variant) As String
    ' This function creates a token from the data.  The data is
    ' converted to a string.  The string is encoded and the
    ' VarType is prepended to it.

    Dim VariantType As String
    Dim VariantOutput As String
    
    VariantType = Trim(Str(VarType(vData))) 'I don't like CStr
    VariantOutput = CStr(vData)             'But I'm stuck with it.
    
    ' If someone decides to allow vbArray to slide, then the VarType
    ' will be 4 digits long.  Might as well plan ahead...
    While Len(VariantType) < 4
        VariantType = "0" & VariantType
    Wend
    
    ' Encode the string as plain old text (ascii 64 to 127, actually)
    VariantOutput = EncodeData(VariantOutput)
            
    'Done
    MakeToken = VariantType & VariantOutput
            
End Function



Private Function UnMakeToken(ByVal vData As String) As Variant
    ' Reverses the damage that MakeToken does.  It's more difficult
    ' here because you can't go VarType(Data) = vbLong.  It only
    ' works one way.  So we have to hardcode our vartype
    ' conversions.  It hurts a bit to do it this way...
    
    Dim VariantType As Integer
    Dim VariantOutput As Variant
    
    'Get our 4-digit VarType
    VariantType = Val(Left(vData, 4))
    
    'The rest must be the data!  See what planning ahead does?
    vData = DecodeData(Right(vData, Len(vData) - 4))
        
    'Just guess on this next one, you'll be correct.
    Select Case VariantType
        Case vbInteger:     VariantOutput = CInt(vData)
        Case vbLong:        VariantOutput = CLng(vData)
        Case vbSingle:      VariantOutput = CSng(vData)
        Case vbDouble:      VariantOutput = CDbl(vData)
        Case vbCurrency:    VariantOutput = CCur(vData)
        Case vbString:      VariantOutput = CStr(vData)
        Case vbBoolean:     VariantOutput = CBool(vData)
        Case vbByte:        VariantOutput = CByte(vData)
        Case vbDate:        VariantOutput = CDate(vData)
        Case vbEmpty:       VariantOutput = Empty
        Case vbArray
            ' Unsupported, but included here for future
            ' consideration.  This one is on you.
            VariantOutput = CStr(vData)
        Case Else
            ' Unknown type.  Make it a string...
            VariantOutput = CStr(vData)
    End Select
       
    'Done!
    UnMakeToken = VariantOutput
        
End Function






Private Function EncodeData(ByVal vData As String) As String
    ' Simple Encoder.  I have been using this encoder for years.  It
    ' basically justs does some math on the character to make it fit
    ' within a certain range.  In this case, 64 to 127.
    '
    ' It works on a mode principal.  Each "mode" covers 64 characters
    ' from 0 to 63, 64 to 127, 128 to 191, 192 to 255.  When a char is
    ' read, it's mode is determined.  Whenever there is a change of
    ' mode, then a number is inserted in front of the mode.
    '
    ' In ASCII, numbers fall between 48 and 57.  If our final chars
    ' are between 64 and 127, then we don't have to worry about using
    ' numbers.  Simple but effective...
    '
    
    Dim CurMode As Integer  'Current mode, starts as 0
    Dim NewMode As Integer  'New mode.  Says we have a different char.
    
    Dim ModeBase As Integer 'This is the cusp of the modes.  4 modes
                            'means 4 cusps.  0, 64, 128, 192.
    
    Dim CurChar As Integer  'The current character code.
    
    Dim NewChar As Integer  'The new character.  Take the current char
                            'and subtract the next lowest cusp.  If the
                            'letter is A then the cusp is 64.  Then add
                            '64 to make a new character.  I know it is
                            'still 65, but so is NUL, ü, and chr(193).
    
    Dim vOutput As String   'The final output string.
    
    
    ' Now we will loop through each character an convert it to it's
    ' new character.  We have to be careful and make sure each Mode
    ' change is numbered, or the decoder won't work.
    For iLoop = 1 To Len(vData)
        CurChar = Asc(Mid(vData, iLoop, 1))
        
        Select Case CurChar
            Case 0 To 63:       NewMode = 1
            Case 64 To 127:     NewMode = 2
            Case 128 To 191:    NewMode = 3
            Case 192 To 255:    NewMode = 4
        End Select
        
        ' Check to see if the mode had changed.  The first time you
        ' run this, CurMode = 0.  This effectively puts the first
        ' Mode number in the text stream.
        If CurMode <> NewMode Then
            
                'Calculate the new modebase
                ModeBase = (NewMode * 64) - 64
                
                'Push the new mode number into the stream
                vOutput = vOutput & Trim(Str(NewMode))
                
                'Equalize the mode indicators
                CurMode = NewMode
                
        End If
        
        'Create a new charater
        NewChar = 64 + (CurChar - ModeBase)
        
        'Push the new character into the stream
        vOutput = vOutput & Chr(NewChar)
    Next iLoop
    
    'Done - Our string is now encoded.
    EncodeData = vOutput
End Function



Private Function DecodeData(ByVal vData As String) As String
    ' Reverse decodes the data that has been encoded.  Don't try
    ' to decode a string that wasn't encoded.  It will fail.
                
    Dim CurMode As Integer  'same
    Dim NewMode As Integer  'same
    Dim ModeBase As Integer 'same
    
    Dim CurChar As Integer  'same
    Dim NewChar As Integer  'same
    
    Dim vOutput As String   'same
    
    
    ' This process is far simpler than encoding.  We just need to
    ' know what the mode is.  We don't have to push extra characters
    ' or anything funky.  In fact, regular characters are fine.
    For iLoop = 1 To Len(vData)
        CurChar = Asc(Mid(vData, iLoop, 1))
        
        Select Case CurChar
            Case 48 To 51
                'Mode change is much simpler.
                NewMode = CurChar - 48
            
            Case Else
                If CurMode <> NewMode Then
                    'Set the mode base
                    ModeBase = (NewMode * 64) - 64
                    CurMode = NewMode   'equalize
                End If
                
                'Back to me old self again.
                NewChar = ModeBase + (CurChar - 64)
                vOutput = vOutput + Chr(NewChar)
        End Select
                
    Next iLoop
    
    'Done
    DecodeData = vOutput
       
End Function

Public Property Get ArrayCount() As Integer
    ArrayCount = zCount 'I would say this explains itself
End Property

Public Property Get ArrayUpperBound() As Integer
    ' There may come a time when this class is part of a DLL or
    ' some other library and the programmer doesn't trust instinct
    ' to 0- or 1-based arrays.  So I included this as a way for
    ' that programmer to know for sure.  Same as ArrayCount.
    ArrayUpperBound = zCount
End Property

Public Property Get ArrayLowerBound() As Integer
    ' Here is the qualifier.  Now a portable programmer can write:
    '       For iLoop = Zab.ArrayLowerBound to Zab.ArrayUpperBound
    ArrayLowerBound = 1
End Property




Public Property Let CaseSensitiveKeys(ByVal vData As Boolean)
    ' Makes the SearchForKeyIndex (private) validate or ignore
    ' the case of the keys.  Provided as an option, of course.
    zCaseSensitiveKeys = vData
End Property

Public Property Get CaseSensitiveKeys() As Boolean
    CaseSensitiveKeys = zCaseSensitiveKeys
End Property


'    DATA_SEPARATOR = Chr(1)  'This separates individual
'                             'array tokens.  Must be <64
'    FIELD_SEPARATOR = Chr(0) 'This separates ID fields
'                             'must be <64
'    ARRAY_TAG = "CF01"    'Version ID


Public Property Let DS(ByVal vData As Integer)
    DATA_SEPARATOR = Chr(vData)
End Property

Public Property Get DS() As Integer
    DS = Asc(DATA_SEPARATOR)
End Property

Public Property Let FS(ByVal vData As Integer)
    FIELD_SEPARATOR = Chr(vData)
End Property

Public Property Get FS() As Integer
    FS = Asc(FIELD_SEPARATOR)
End Property

Public Property Let TAG(ByVal vData As String)
    ARRAY_TAG = vData
End Property

Public Property Get TAG() As String
    TAG = ARRAY_TAG
End Property
