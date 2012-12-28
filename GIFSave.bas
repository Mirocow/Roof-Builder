Attribute VB_Name = "GIFSave"
' GIFSave.bas

' Modified from:-
' mGIFSave.bas  -  master file for writing GIF files
'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

' from the C copyright ©1997 Ron van Tilburg 25.12.1997
' VB copyright ©2000 Ron van Tilburg 24.12.2000
' and copyrights of the original C code from which this is derived are given in the body
' Documentation of GIF structures is from the GIF standard as attached as html documents
' All copyrights applying there continue to apply

' Modified from mGIFSave.bas by Ron van Tilburg & Carles P V
' (See PSC CodeId=49875 for full references) to just save
' 8bpp GIF 87a - no transparency, no comments, non-interlaced.

Option Explicit
Option Base 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

'-- GIF format. Structures and Constants.

Private Const GIF_SIGNATURE                 As String = "GIF"
Private Const GIF_VERSION_87a               As String = "87a"
Private Const GIF_TRAILER                   As Byte = &H3B

Private Const GIF_IMAGE_DESCRIPTOR          As Byte = &H2C
Private Const GIF_BLOCK_TERMINATOR          As Byte = &H0

Private Type tGIFScreenDescriptor
    sdLogicalScreenWidth     As Integer
    sdLogicalScreenHeight    As Integer
    sdFlags                  As Byte
    sdBackgroundColorIndex   As Byte
    sdPixelAspectRatio       As Byte
End Type

Private Type tGIFImageDescriptor
    idImageLeftPosition      As Integer
    idImageTopPosition       As Integer
    idImageWidth             As Integer
    idImageHeight            As Integer
    idFlags                  As Byte
End Type

'-- GLOBAL VARIABLES for the Encoding Routines

Private Const MAX_BITS                    As Long = 12           ' User settable max - bits/code
Private Const MAX_BITSHIFT                As Long = 2 ^ MAX_BITS
Private Const MAX_CODE                    As Long = 2 ^ MAX_BITS ' Should NEVER generate this code
Private Const EOF_CODE                    As Long = -1           ' END of input
Private Const TABLE_SIZE                  As Long = 5003         ' 80% occupancy
Private m_lBits                           As Long                ' Number of bits/code
Private m_lMaxCode                        As Long                ' Maximum code, given m_lBits
Private m_lHashTable(0 To TABLE_SIZE - 1) As Long
Private m_lCodeTable(0 To TABLE_SIZE - 1) As Long
Private m_lFreeEntry                      As Long                ' First unused entry

'-- Block compression parameters.
'   After all codes are used up, and compression rate changes, start over.
Private m_lClearFlag     As Long
Private m_lInitBits      As Long
Private m_lClearCode     As Long
Private m_lEOFCode       As Long

'-- Variables for positioning and control
Private m_lx             As Long    ' Image current x pos.
Private m_ly             As Long    ' Image current y pos.
Private m_lImageWidth    As Long    ' Image Width
Private m_lImageHeight   As Long    ' Image Height
Private m_lPixelCount    As Long    ' Pixels left to do
Private m_lPass          As Long    ' Which m_lPass in interlaced mode
Private m_lOutputBytes   As Long    ' Bytes output so far

'-- Variables for the code accumulator (pvOutputCode)
Private m_lOutputBucket  As Long
Private m_lOutputBits    As Long
Private m_lMask(0 To 16) As Long    ' Powers of 2 -1

'-- Variables for the output byte accumulator
Private m_lCharCount     As Long    ' Number of characters so far in this 'packet'
Private m_aChar()        As Byte    ' Will be max 256 bytes long, first byte is length

'-- Global file handler
Private m_hFile   As Long

'-- 8bpp-DIB mapped bytes
Private m_aBits() As Byte

Private bred As Byte, bgreen As Byte, bblue As Byte

Public Function MSaveGIF(ByVal FileName As String, _
                         bA() As Byte, _
                         gWidth As Integer, _
                         gHeight As Integer, _
                         gPAL() As Long, _
                         aConv As Boolean)
    'In: bA() Byte array ( 1 to gWidth, 1 to gHeight)
    '    NB gWidth & gHeight are Integers
    '    gPAL(0 to 255) Long RGBA

    Dim tScreenDescriptor        As tGIFScreenDescriptor
    Dim tImageDescriptor         As tGIFImageDescriptor
    Dim aBPP                     As Byte
    Dim LIdx As Long

    Dim k As Long
    Dim ix As Long
    Dim iy As Long

        '-- Init LUT for fast 2 ^ x - 1
        m_lMask(0) = 0
        For LIdx = 1 To 16
            m_lMask(LIdx) = 2 * (m_lMask(LIdx - 1) + 1) - 1
        Next LIdx

        ' Copy bA() (height already reversed for BMP saving) to m_aBits
        ReDim m_aBits(1 To gWidth, 1 To gHeight)
        ' If aConvert = True need to swap Y scan lines
        If aConv Then
            For iy = 1 To gHeight
                CopyMemory m_aBits(1, gHeight - iy + 1), bA(1, iy), gWidth
            Next iy

        Else
            CopyMemory m_aBits(1, 1), bA(1, 1), 1& * gWidth * gHeight
        End If
   
        '-- Kill previous
        On Error Resume Next
        Kill FileName
        On Error GoTo 0
   
        '-- Get a free file handle and open a new one
        m_hFile = FreeFile()
        Open FileName For Binary Access Write As #m_hFile
        On Error GoTo ErrSave
   
        '-- Write GIF header
        Put #m_hFile, , GIF_SIGNATURE
        Put #m_hFile, , GIF_VERSION_87a
        aBPP = 8
        '-- Prepare screen descriptor
        With tScreenDescriptor
            .sdLogicalScreenWidth = gWidth
            .sdLogicalScreenHeight = gHeight
            .sdFlags = &HF0 Or (aBPP - 1)
            .sdBackgroundColorIndex = 0
            .sdPixelAspectRatio = 0
        End With
   
        '-- Prepare image descriptor
        With tImageDescriptor
            .idImageLeftPosition = 0
            .idImageTopPosition = 0
            .idImageWidth = gWidth
            .idImageHeight = gHeight
            .idFlags = &H7
        End With
   
        '-- Write screen descriptor and global palette
        Put #m_hFile, , tScreenDescriptor
   
        'Write RGB palette
        For k = 0 To 255
            Put #m_hFile, , CByte(gPAL(k) And &HFF&)                 ' Red
            Put #m_hFile, , CByte((gPAL(k) And &HFF00&) / &H100&)    ' Green
            Put #m_hFile, , CByte((gPAL(k) And &HFF0000) / &H10000)  ' Blue
        Next k
   
        '-- Write GIF image descriptor
        Put #m_hFile, , GIF_IMAGE_DESCRIPTOR
        Put #m_hFile, , tImageDescriptor
        '-- Write GIF-LZW code size
        Put #m_hFile, , aBPP
   
        '-- Prepare some vars. for compress and write image data
        m_lImageWidth = CLng(gWidth)
        m_lImageHeight = CLng(gHeight)
        m_lPixelCount = 1& * gWidth * gHeight
   
        '-- Compress/Write image data
        Call pvCompressAndWriteBits(aBPP + 1)
   
        Put #m_hFile, , GIF_BLOCK_TERMINATOR
   
        '-- Finaly, write trailer label
        Put #m_hFile, , GIF_TRAILER
   
        '-- Close file: success
        Close #m_hFile
        Erase m_aBits()
ErrSave:
        On Error GoTo 0
End Function


Private Sub pvCompressAndWriteBits(nInitBits As Integer)

    Dim LIdx     As Long
    Dim lFCode   As Long
    Dim lC       As Long
    Dim lEnt     As Long
    Dim lDisp    As Long
    Dim m_lShift As Long

        '-- Set up where we are starting
        LIdx = 0
        m_lOutputBytes = 0
        m_lPass = 0
        m_lx = 1   '0
        m_ly = 1   '0
   
        '-- Set up the code accumulator
        m_lOutputBucket = 0
        m_lOutputBits = 0
   
        '-- Set up initial number of bits
        m_lInitBits = nInitBits
   
        '-- Set up the necessary values
        m_lClearFlag = 0
        m_lBits = m_lInitBits
        m_lMaxCode = m_lMask(m_lBits)
        m_lClearCode = 2 ^ (nInitBits - 1)
        m_lEOFCode = m_lClearCode + 1
        m_lFreeEntry = m_lClearCode + 2
   
        '-- Set up output buffers
        Call pvCharInit
   
        m_lShift = 0
        lFCode = TABLE_SIZE
        Do While lFCode < 65536
            m_lShift = m_lShift + 1
            lFCode = lFCode + lFCode
        Loop
   
        '-- Set hash code range bound for shifting
        m_lShift = 1 + m_lMask(8 - m_lShift)
   
        Call pvClearTable
        Call pvOutputCode(m_lClearCode)
   
        '-- Start...
        lEnt = pvGetPixel: lC = pvGetPixel
   
        Do While lC <> EOF_CODE
   
            lFCode = lC * MAX_BITSHIFT + lEnt
            LIdx = (lC * m_lShift) Xor lEnt      ' XOR hashing
   
            If (m_lHashTable(LIdx) = lFCode) Then
                lEnt = m_lCodeTable(LIdx)
                GoTo NextPixel
            ElseIf (m_lHashTable(LIdx) < 0) Then ' Empty slot
                GoTo NoMatch
            End If
   
            lDisp = TABLE_SIZE - LIdx            ' Secondary hash (after G. Knott)
            If (LIdx = 0) Then lDisp = 1

Probe:
            LIdx = LIdx - lDisp
            If (LIdx < 0) Then LIdx = LIdx + TABLE_SIZE
      
            If (m_lHashTable(LIdx) = lFCode) Then
                lEnt = m_lCodeTable(LIdx)
                GoTo NextPixel
            End If
      
            If (m_lHashTable(LIdx) > 0) Then GoTo Probe

NoMatch:
            Call pvOutputCode(lEnt)
            lEnt = lC

            If (m_lFreeEntry < MAX_CODE) Then
                m_lCodeTable(LIdx) = m_lFreeEntry
                m_lFreeEntry = m_lFreeEntry + 1  ' Code -> Hash table
                m_lHashTable(LIdx) = lFCode
            Else
                Call pvClearBlock
            End If

NextPixel:
            lC = pvGetPixel
    
        Loop

        '--  Put out the final code
        Call pvOutputCode(lEnt)
        Call pvOutputCode(m_lEOFCode)
End Sub


Private Function pvGetPixel() As Integer

    If (m_lPixelCount = 0) Then
        '-- End of data
        pvGetPixel = EOF_CODE
   
    Else
        '-- Return the next pixel from the image and increment positions
        pvGetPixel = m_aBits(m_lx, m_ly)
   
        m_lx = m_lx + 1
        If (m_lx > m_lImageWidth) Then    ' =  >
            m_lx = 1
            m_ly = m_ly + 1
        End If

        m_lPixelCount = m_lPixelCount - 1
    End If

End Function


Private Sub pvOutputCode(ByVal lCode As Long)
    '-- Output the given code.
    '   Assumptions:
    '     - Chars are 8 bits long.
    '   Algorithm:
    '     - Maintain a MAX_BITS character long buffer (so that 8 codes will fit in it exactly).
    '     - When the buffer fills up empty it and start over.

    m_lOutputBucket = m_lOutputBucket And m_lMask(m_lOutputBits)
   
    If (m_lOutputBits > 0) Then
        m_lOutputBucket = m_lOutputBucket Or (lCode * (1 + m_lMask(m_lOutputBits)))
    Else
        m_lOutputBucket = lCode
    End If

    m_lOutputBits = m_lOutputBits + m_lBits
   
    Do While (m_lOutputBits >= 8)
        Call pvCharOut(m_lOutputBucket And &HFF&)
        m_lOutputBucket = m_lOutputBucket / 256&
        m_lOutputBits = m_lOutputBits - 8
    Loop
   
    '-- If the next entry is going to be too big for the code size, then increase it, if possible.
    If (m_lFreeEntry > m_lMaxCode Or m_lClearFlag = -1) Then
        If (m_lClearFlag = -1) Then
            m_lBits = m_lInitBits
            m_lMaxCode = m_lMask(m_lBits)
            m_lClearFlag = 0
        Else
            m_lBits = m_lBits + 1
            If (m_lBits = MAX_BITS) Then
                m_lMaxCode = MAX_CODE
            Else
                m_lMaxCode = m_lMask(m_lBits)
            End If

        End If

    End If
   
    '-- At EOF, write the rest of the buffer.
    If (lCode = m_lEOFCode) Then
        Do While (m_lOutputBits > 0)
            Call pvCharOut(m_lOutputBucket And &HFF&)
            m_lOutputBucket = m_lOutputBucket / 256&
            m_lOutputBits = m_lOutputBits - 8
        Loop

        Call pvFlushChar
    End If

End Sub


Private Sub pvClearBlock()
    '-- Clear out the hash table for block compress
    Call pvClearTable
    m_lFreeEntry = m_lClearCode + 2
    m_lClearFlag = -1
    Call pvOutputCode(m_lClearCode)
End Sub


Private Sub pvClearTable()
    '-- Reset code table
    Dim LIdx As Long

        For LIdx = 0 To TABLE_SIZE - 1
            m_lHashTable(LIdx) = -1
        Next LIdx

End Sub


Private Sub pvCharInit()
    '-- Set up the 'byte output' routine and define the storage for the packet accumulator
    m_lCharCount = 0
    ReDim m_aChar(0 To 255) As Byte
End Sub


Private Sub pvCharOut(ByVal lChar As Long)
    '-- Add a character to the end of the current packet, and if it is 254 characters,
    '   flush the packet to disk
    m_aChar(m_lCharCount + 1) = lChar              ' 0,...,n mapped to 1,...,n+1
    m_lCharCount = m_lCharCount + 1
    If (m_lCharCount >= 254) Then Call pvFlushChar
End Sub


Private Sub pvFlushChar()
    '-- Flush the current packet to disk, and reset the accumulator
    If (m_lCharCount > 0) Then
        m_aChar(0) = m_lCharCount                          ' Set block length,
        ReDim Preserve m_aChar(0 To m_lCharCount) As Byte  ' and redimension to this length
        Put #m_hFile, , m_aChar()                          ' Write it to disk
        m_lOutputBytes = m_lOutputBytes + m_lCharCount + 1 ' Track bytes written
        Call pvCharInit
    End If

End Sub


'Private Sub pvCheckAndWriteComment(sComment As String)
'Dim aBuff()   As Byte
'Dim aBuffSize As Byte
'
'   '-- 255 chars max.
'   aBuffSize = Len(sComment)
'   If (aBuffSize > 255) Then
'       aBuffSize = 255
'   End If
'   '-- Fill byte array buffer
'   ReDim aBuff(1 To aBuffSize)
'   CopyMemory aBuff(1), ByVal sComment, aBuffSize
'
'   '-- Write
'   Put #m_hFile, , aBuffSize ' Block size
'   Put #m_hFile, , aBuff()   ' Block itself
'End Sub



