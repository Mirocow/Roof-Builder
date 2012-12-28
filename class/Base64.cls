VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Base64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' The original source code was from someone (AndrComm or Sebastien?) else.
' It's kinda unreadable (at least i'm the one who feel so, no offense).
' So i decided to recode the whole thing, but having reference to the original.
' Anyway thanks, whoever who are who wrote the original code (yeahh.. all creadits to him)
' since i learnt some performance optimization (my original codes are damn slow, sucks).
' I was so excited when i found that piece of code a few days ago and i work
' nonstop for 2 days and come out with my own.. kinda crazy, am i?
'
' I didn't make any major changes, just some code clean up and comments.
' The encode speed increased by 20%, decoded speed increased by 30%.
' In order to gain speed.. much of the safety checkups have been removed
' thus crashes more easily. Compile in NATIVE CODE to gain more speed.
' Check Project -> Base64 Properties -> Compile -> Advanced Optimizations.
'
' After running 15 times of test on my AMD K6-2 300Mhz 64mb Quantum FireBall HDD
' (kinda outdated :p) here's the average results that i come out with:
'
'   file : command.com (93,040 bytes)
'   load : 2.56 mb/sec (n/a in original code)
'   save : 3.39 mb/sec (n/a in original code)
' encode : 7.17 mb/sec (20% faster, original code runs at 6.05 mb/sec)
' decode : 7.31 mb/sec (30% faster, original code runs at 5.61 mb/sec)
'
' Here's the changes I've made :
'
' 1. *REWROTE* code clean up (i think it should be more readable)
' 2. *REWROTE* encode and decode table is created together.
' 3. *REWROTE* sub DECODE rewrote (should be faster).
' 4. *REWROTE* sub SPAN and UNSPAN now supports SpanSeparator.
' 5. *ADDED*   More detailed co1mments for novice, not experts
' 5. *ADDED*   some form controls, for input and output.
' 6. *ADDED*   file loading and saving feature
' 7. *FIXED*   encode / decode rate calculation fixed... more accurate
'              should be 1 sec = 1024 ticks (kernel 1044 ticks?)
'              and 1 mb = 1048576 bytes. (1024b x 1024b)
'
' and some other minor corrections as well...
' below are some useful details that i would like to share with others.
'
' P/s : someone please look at the load file sub? Too jerky and
'       i don't know how to optimize it...
'       Sorry if the lines are too long, i'm working in 1600x1200 environment
'
' Johnny (chinhuat@i.am, cknvally@tm.net.my)
' Friday, September 08, 2000
' 1418 GMT +0800
'
' (you may delete the shits above and redistribute, all credits to those guys not me)
'
' -------------------------------- snip snip snip --------------------------------
'
' ******************************************
' * Base64 Encode / Decode Algorythm Class *
' ******************************************
'
' +---------------------+
' | Optimizations Found |
' +---------------------+
'
' 1. predeclare all variable (avoid using variant)
' 2. precalculate all values, do not recalculate them (especially modulus)
' 3. use byte array instead of string (strings are array too but are slower)
' 4. use \ instead of / ( / returns rounded decimal)
' 5. use table for lookups instead of doing silly calculations
' 6. use logic operands (AND, OR) to mask out bits
'    (don't to silly DEC-BIN BIN-DEC convertion like i do!)
'
' +-------------------------------+
' | Base64 Encode / Decode method |
' +-------------------------------+
'
'   Source : RFC 1341 - MIME (Multipurpose Internet Mail Extensions)
'                             Mechanisms for Specifying and Describing
'                             the Format of Internet Message Bodies
'      URL : http://www.faqs.org/ (highly recommended)
'
' -------------------------------- snip snip snip --------------------------------
'
'            Borenstein & Freed                                 [Page 16]
'
'            RFC 1341MIME: Multipurpose Internet Mail ExtensionsJune 1992
'
'            5.2  Base64 Content-Transfer-Encoding
'
'            The  Base64   Content-Transfer-Encoding   is   designed   to
'            represent  arbitrary  sequences  of octets in a form that is
'            not humanly readable.  The encoding and decoding  algorithms
'            are simple, but the encoded data are consistently only about
'            33 percent larger than the unencoded data.  This encoding is
'            based on the one used in Privacy Enhanced Mail applications,
'            as defined in RFC 1113.   The  base64  encoding  is  adapted
'            from  RFC  1113, with one change:  base64 eliminates the "*"
'            mechanism for embedded clear text.
'
'            A 65-character subset of US-ASCII is used, enabling  6  bits
'            to  be  represented per printable character. (The extra 65th
'            character, "=", is used  to  signify  a  special  processing
'            function.)
'
'            NOTE:  This subset has the important  property  that  it  is
'            represented   identically   in  all  versions  of  ISO  646,
'            including US ASCII, and all characters  in  the  subset  are
'            also  represented  identically  in  all  versions of EBCDIC.
'            Other popular encodings, such as the encoding  used  by  the
'            UUENCODE  utility  and the base85 encoding specified as part
'            of Level 2 PostScript, do not share  these  properties,  and
'            thus  do  not  fulfill the portability requirements a binary
'            transport encoding for mail must meet.
'
'            The encoding process represents 24-bit groups of input  bits
'            as  output  strings of 4 encoded characters. Proceeding from
'            left  to  right,  a  24-bit  input  group   is   formed   by
'            concatenating  3  8-bit input groups. These 24 bits are then
'            treated as 4 concatenated 6-bit groups,  each  of  which  is
'            translated  into a single digit in the base64 alphabet. When
'            encoding a bit stream  via  the  base64  encoding,  the  bit
'            stream  must  be  presumed  to  be  ordered  with  the most-
'            significant-bit first.  That is, the first bit in the stream
'            will be the high-order bit in the first byte, and the eighth
'            bit will be the low-order bit in the first byte, and so on.
'
'            Each 6-bit group is used as an index into  an  array  of  64
'            printable  characters. The character referenced by the index
'            is placed in the output string. These characters, identified
'            in  Table  1,  below,  are  selected so as to be universally
'            representable,  and  the  set   excludes   characters   with
'            particular  significance to SMTP (e.g., ".", "CR", "LF") and
'            to the encapsulation boundaries  defined  in  this  document
'            (e.g., "-").
'
'            Borenstein & Freed                                 [Page 17]
'
'            RFC 1341MIME: Multipurpose Internet Mail ExtensionsJune 1992
'
'                            Table 1: The Base64 Alphabet
'
'               Value Encoding  Value  Encoding   Value  Encoding   Value
'            encoding
'                   0 A            17 R            34 i            51 z
'                   1 B            18 S            35 j            52 0
'                   2 C            19 T            36 k            53 1
'                   3 D            20 U            37 l            54 2
'                   4 E            21 V            38 m            55 3
'                   5 F            22 W            39 n            56 4
'                   6 G            23 X            40 o            57 5
'                   7 H            24 Y            41 p            58 6
'                   8 I            25 Z            42 q            59 7
'                   9 J            26 a            43 r            60 8
'                  10 K            27 b            44 s            61 9
'                  11 L            28 c            45 t            62 +
'                  12 M            29 d            46 u            63 /
'                  13 N            30 e            47 v
'                  14 O            31 f            48 w         (pad) =
'                  15 P            32 g            49 x
'                  16 Q            33 h            50 y
'
'            The output stream (encoded bytes)  must  be  represented  in
'            lines  of  no more than 76 characters each.  All line breaks
'            or other characters not found in Table 1 must be ignored  by
'            decoding  software.   In  base64 data, characters other than
'            those in  Table  1,  line  breaks,  and  other  white  space
'            probably  indicate  a  transmission  error,  about  which  a
'            warning  message  or  even  a  message  rejection  might  be
'            appropriate under some circumstances.''
'
'            Special processing is performed if fewer than  24  bits  are
'            available  at  the  end  of  the data being encoded.  A full
'            encoding quantum is always completed at the end of  a  body.
'            When  fewer  than  24  input  bits are available in an input
'            group, zero bits  are  added  (on  the  right)  to  form  an
'            integral number of 6-bit groups.  Output character positions
'            which are not required to represent actual  input  data  are
'            set  to  the  character  "=".   Since all base64 input is an
'            integral number of octets,  only  the  following  cases  can
'            arise:  (1)  the  final  quantum  of  encoding  input  is an
'            integral multiple of  24  bits;  here,  the  final  unit  of
'            encoded  output will be an integral multiple of 4 characters
'            with no "=" padding, (2) the final quantum of encoding input
'            is  exactly  8  bits; here, the final unit of encoded output
'            will  be  two  characters  followed  by  two   "="   padding
'            characters,  or  (3)  the final quantum of encoding input is
'            exactly 16 bits; here, the final unit of encoded output will
'            be three characters followed by one "=" padding character.
'
'            Care must be taken to use the proper octets for line  breaks
'            if base64 encoding is applied directly to text material that
'            has not been converted to  canonical  form.  In  particular,
'            text  line  breaks  should  be converted into CRLF sequences
'
'            Borenstein & Freed                                 [Page 18]
'
'            RFC 1341MIME: Multipurpose Internet Mail ExtensionsJune 1992
'
'            prior to base64 encoding. The important  thing  to  note  is
'            that this may be done directly by the encoder rather than in
'            a prior canonicalization step in some implementations.
'
'            NOTE: There is no  need  to  worry  about  quoting  apparent
'            encapsulation  boundaries  within  base64-encoded  parts  of
'            multipart entities because no hyphen characters are used  in
'            the base64 encoding.
'
'
' -------------------------------- snip snip snip --------------------------------

Option Explicit                     ' just to make sure everything is declared

Private B64EncTable(63) As Byte     ' encode table
Private B64DecTable(255) As Byte    ' decode table


' +------------------------------------------------------+
' | String to byte array conversion                      |
' +------------------------------------------------------+
' | Parameter   :                                        |
' | strIn  - input string                                |
' | aryOut - output array                                |
' |                                                      |
' | Description :                                        |
' | Note that output array does not have be to dimension |
' | instead is automatically done by StrConv             |
' +------------------------------------------------------+

Public Function Str2ByteArray(ByVal strIn As String) As Byte()
Str2ByteArray = StrConv(strIn, vbFromUnicode)  ' Convert string to byte array
End Function

' +----------------------------+
' | Array to string conversion |
' +----------------------------+
' | Parameter   :              |
' | aryIn  - input array       |
' | strOut - output string     |
' |                            |
' | Description :              |
' | None                       |
' +----------------------------+

Public Function ByteArray2Str(ByRef aryIn() As Byte) As String
ByteArray2Str = StrConv(aryIn, vbUnicode)      ' Convert byte array to string
End Function

' +-------------------------------------------------------------------------------+
' | Encode an byte array                                                          |
' +-------------------------------------------------------------------------------+
' | Parameter                                                                     |
' | aryIn  - Input array                                                          |
' | aryOut - Output array (will be redimensioned)                                 |
' |                                                                               |
' | Description :                                                                 |
' | Size of output is always 1/3 larger (3 bytes becomes 4) due to the algorythm. |
' | Take note that the output size should ALWAYS be dividable by 4 perfectly.     |
' | If the input size is not dividable by 3 perfectly, the next empty byte is     |
' | filled with 0 (null) to round the last group of bits into 6. If input byte is |
' | 1 then output will be 2 encoded bytes followed with 2 "=", if 2 bytes is      |
' | input then output should be 3 encoded bytes with 1 "="                        |
' +-------------------------------------------------------------------------------+

Public Function Encode(ByRef aryIn() As Byte) As Byte()

Dim SizeIn As Long      ' Size of input array
Dim SizeOut As Long     ' Size of output array (to be calculate)
Dim SizeMod As Integer  ' Remainder of input (due to base64 encode algorythm)
Dim Buffer(2) As Byte   ' Temporary buffer to store 3 bytes of data to be encode
Dim Index As Long       ' Encode index
Dim oIndex As Long      ' Output Index
Dim aryOut() As Byte

On Error Resume Next

SizeIn = UBound(aryIn) + 1                  ' size of input
SizeMod = SizeIn Mod 3                      ' remaining input bytes
SizeOut = (SizeIn - SizeMod) / 3 * 4        ' size of output (dividable by 3)
If SizeMod > 0 Then SizeOut = SizeOut + 4   ' if there is remaining bytes, add another group of 4 byte

ReDim aryOut(SizeOut - 1)                   ' redimension output array

If SizeIn >= 3 Then                         ' Input has more than 1 group of 3 bytes

    For Index = 0 To SizeIn - SizeMod - 1 Step 3    ' Start from 1st element in input byte array
    
        Buffer(0) = aryIn(Index)            ' Fetch next 3 bytes
        Buffer(1) = aryIn(Index + 1)        ' into temporary array
        Buffer(2) = aryIn(Index + 2)        ' (arrays are faster)

        ' Encode 3 bytes to 4
        ' by concatenating 3 x 8 bits into 24 bits
        ' then divide into 4 x 6 bit group
        ' then encode using the encode the table
        
        aryOut(oIndex) = B64EncTable((Buffer(0) And &HFC) \ 4)
        aryOut(oIndex + 1) = B64EncTable((Buffer(0) And &H3) * 16 Or (Buffer(1) And &HF0) \ 16)
        aryOut(oIndex + 2) = B64EncTable((Buffer(1) And &HF) * 4 Or (Buffer(2) And &HC0) \ 64)
        aryOut(oIndex + 3) = B64EncTable((Buffer(2) And &H3F))
        
        oIndex = oIndex + 4                 ' Increment output array index

    Next

End If

If SizeMod = 1 Then                         ' 1 byte remaining in input buffer
    
    Buffer(0) = aryIn(SizeIn - 1)           ' Get the remaining 1 byte
    Buffer(1) = 0                           ' Fill the 2nd buffer with null (to round the 2nd byte into 6 bits)
    
    ' Output 2 bytes with 2 "="
    aryOut(oIndex) = B64EncTable((Buffer(0) And &HFC) \ 4)
    aryOut(oIndex + 1) = B64EncTable((Buffer(0) And &H3) * 16 Or (Buffer(1) And &HF0) \ 16)
    aryOut(oIndex + 2) = 61                 ' Neccessary to assign "==" to round output into 1 group of 4 bytes
    aryOut(oIndex + 3) = 61                 ' according to Base64 Encode Standard

ElseIf SizeMod = 2 Then                     ' 2 bytes remaining in input buffer

    Buffer(0) = aryIn(SizeIn - 2)           ' Get the remaining 2 bytes
    Buffer(1) = aryIn(SizeIn - 1)
    Buffer(2) = 0                           ' Fill the 3rd buffer with null (to round the 3rd byte into 6 bits)
    
    ' Output 3 bytes with 1 "="
    aryOut(oIndex) = B64EncTable((Buffer(0) And &HFC) \ 4)
    aryOut(oIndex + 1) = B64EncTable((Buffer(0) And &H3) * 16 Or (Buffer(1) And &HF0) \ 16)
    aryOut(oIndex + 2) = B64EncTable((Buffer(1) And &HF) * 4 Or (Buffer(2) And &HC0) \ 64)
    aryOut(oIndex + 3) = 61                 ' Necessary to assign "=" to round output into 1 group of 4 bytes
    
End If

Encode = aryOut
End Function

' +-----------------------------------------------------------------------------+
' | Decode a byte array                                                         |
' +-----------------------------------------------------------------------------+
' | Parameter   :                                                               |
' | aryIn  - input array                                                        |
' | aryOut - output array (will be redimensioned)                               |
' |                                                                             |
' | Description :                                                               |
' | Size of output is always 1/4 smaller than the original (4 bytes becomes 3). |
' | The reversed mechanisme was used base on the encode mechanisme. Output size |
' | maybe not dividable perfectly by 3, thus a checkup on last 2 bytes of input |
' | is done. If 1 "=" encounted then the last group of 4 bytes is decoded into  |
' | 2 bytes, and if 2 "=" is encounted 1 bytes will be output instead.          |
' +-----------------------------------------------------------------------------+

Public Function Decode(ByRef aryIn() As Byte) As Byte()

Dim SizeIn As Long      ' Size of input array
Dim SizeOut As Long     ' Size of output array (to be calculate)
Dim SizeMod As Integer  ' Remainder of output
Dim ActSizeIn As Long   ' Actual input (dividable by 4)
Dim Index As Long       ' Decode Input
Dim oIndex As Long      ' Output Index
Dim Buffer(3) As Byte   ' Temporary buffer to store 4 bytes of data to be decoded
Dim aryOut() As Byte

On Error Resume Next

SizeIn = UBound(aryIn) + 1      ' Get the size of input
ActSizeIn = SizeIn

If aryIn(SizeIn - 1) = 61 Then          ' Last byte is a "="
    If aryIn(SizeIn - 2) = 61 Then      ' Second Last byte is a"="
        SizeMod = 1                     ' Output remainder should be 1 byte
    Else
        SizeMod = 2                     ' Output remainder should be 2 bytes
    End If
    ActSizeIn = SizeIn - 4              ' Substitude the last group of 4 bytes since they are incomplete (contains "=")
    If ActSizeIn < 0 Then ActSizeIn = 0 ' Since input maybe just "dA==" or "dHA=" less than 4 complete bytes instead of "GfgrDfF="
End If

SizeOut = (ActSizeIn / 4 * 3) + SizeMod ' Output size is 3/4 of input plus the output remainder

ReDim aryOut(SizeOut - 1)               ' Redimension output array

If ActSizeIn >= 4 Then                  ' Have more than 1 group of complete 4 bytes

    For Index = 0 To ActSizeIn - 1 Step 4
    
        Buffer(0) = B64DecTable(aryIn(Index))       ' Input 4 bytes
        Buffer(1) = B64DecTable(aryIn(Index + 1))   ' and decode them
        Buffer(2) = B64DecTable(aryIn(Index + 2))
        Buffer(3) = B64DecTable(aryIn(Index + 3))
        
        aryOut(oIndex + 0) = ((Buffer(0) * 4) Or ((Buffer(1) And &H30) \ 16))
        aryOut(oIndex + 1) = (((Buffer(1) And &HF) * 16) Or ((Buffer(2) And &H3C)) \ 4)
        aryOut(oIndex + 2) = (((Buffer(2) And &H3) * 64) Or Buffer(3))
    
        oIndex = oIndex + 3             ' Increment output index
        
    Next

End If

If SizeMod = 1 Then         ' Output has 1 remainder byte

        Buffer(0) = B64DecTable(aryIn(SizeIn - 4))  ' input 2 remaining bytes from input
        Buffer(1) = B64DecTable(aryIn(SizeIn - 3))  ' this byte has portion of 1st output byte ONLY
        Buffer(2) = 0                               ' ignore the 2 "="
        
        ' output 1 byte
        aryOut(oIndex + 0) = ((Buffer(0) * 4) Or ((Buffer(1) And &H30) \ 16))
        
ElseIf SizeMod = 2 Then     ' Output has 2 remainder byte

        Buffer(0) = B64DecTable(aryIn(SizeIn - 4))  ' input the 3 remaining bytes
        Buffer(1) = B64DecTable(aryIn(SizeIn - 3))
        Buffer(2) = B64DecTable(aryIn(SizeIn - 2))  ' this byte has portion of 2nd output byte ONLY
        Buffer(3) = 0                               ' ignore the last "="
        
        ' output 2 bytes
        aryOut(oIndex + 0) = ((Buffer(0) * 4) Or ((Buffer(1) And &H30) \ 16))
        aryOut(oIndex + 1) = (((Buffer(1) And &HF) * 16) Or ((Buffer(2) And &H3C)) \ 4)

End If
Decode = aryOut
End Function


' +---------------------------------------------------+
' | Build Codec Table                                 |
' +---------------------------------------------------+
' | Parameter   :                                     |
' | None                                              |
' |                                                   |
' | Description :                                     |
' | Eg 0 is encoded into A so we store A at element 0 |
' | of encode table and vice versa for decode table.  |
' | Note that byte array is used instead of string.   |
' | Refer to the RFC on the codec table.              |
' +---------------------------------------------------+


Private Sub Class_Initialize()
Dim Index As Integer

' Initialize Codec Table

' A-Z
For Index = 65 To 90
    B64EncTable(Index - 65) = Index
    B64DecTable(Index) = Index - 65
Next

' a-z
For Index = 97 To 122
    B64EncTable(Index - 71) = Index
    B64DecTable(Index) = Index - 71
Next

' 0-9
For Index = 48 To 57
    B64EncTable(Index + 4) = Index
    B64DecTable(Index) = Index + 4
Next

' + and /
B64EncTable(62) = 43: B64DecTable(43) = 62
B64EncTable(63) = 47: B64DecTable(47) = 63
End Sub

