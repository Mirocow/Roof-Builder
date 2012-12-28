Attribute VB_Name = "modAniGif"
Option Explicit
Public RepeatTimes As Long 'This one calculates,
' but don't use in this sample. If You need, You
' can add simple checking at Timer1_Timer Procedure
Public TotalFrames As Long

Public Function LoadGif(sFile As String, aImg As Variant) As Boolean
    LoadGif = False
    If Dir$(sFile) = "" Or sFile = "" Then
       MsgBox "File " & sFile & " not found", vbCritical
       Exit Function
    End If
    On Error GoTo ErrHandler
    Dim fNum As Integer
    Dim imgHeader As String, fileHeader As String
    Dim buf$, picbuf$
    Dim imgCount As Integer
    Dim i&, j&, xOff&, yOff&, TimeWait&
    Dim GifEnd As String
    GifEnd = chr$(0) & chr$(33) & chr$(249)
    For i = 1 To aImg.Count - 1
        Unload aImg(i)
    Next i
    fNum = FreeFile
    Open sFile For Binary Access Read As fNum
        buf = string$(LOF(fNum), chr$(0))
        Get #fNum, , buf 'Get GIF File into buffer
    Close fNum
    
    i = 1
    imgCount = 0
    j = InStr(1, buf, GifEnd) + 1
    fileHeader = Left(buf, j)
    If Left$(fileHeader, 3) <> "GIF" Then
       MsgBox "This file is not a *.gif file", vbCritical
       Exit Function
    End If
    LoadGif = True
    i = j + 2
    If Len(fileHeader) >= 127 Then
        RepeatTimes& = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * 256&)
    Else
        RepeatTimes = 0
    End If

    Do ' Split GIF Files at separate pictures
       ' and load them into Image Array
        imgCount = imgCount + 1
        j = InStr(i, buf, GifEnd) + 3
        If j > Len(GifEnd) Then
            fNum = FreeFile
            Open "temp.gif" For Binary As fNum
                picbuf = string$(Len(fileHeader) + j - i, chr$(0))
                picbuf = fileHeader & Mid(buf, i - 1, j - i)
                Put #fNum, 1, picbuf
                imgHeader = Left(Mid(buf, i - 1, j - i), 16)
            Close fNum
            TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256&)) * 10&
            If imgCount > 1 Then
                xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&)
                yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256&)
                Load aImg(imgCount - 1)
                aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
                aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
            End If
            ' Use .Tag Property to save TimeWait interval for separate Image
            aImg(imgCount - 1).Tag = TimeWait
            aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
            Kill ("temp.gif")
            i = j
        End If
        DoEvents
    Loop Until j = 3
' If there are one more Image - Load it
    If i < Len(buf) Then
        fNum = FreeFile
        Open "temp.gif" For Binary As fNum
            picbuf = string$(Len(fileHeader) + Len(buf) - i, chr$(0))
            picbuf = fileHeader & Mid(buf, i - 1, Len(buf) - i)
            Put #fNum, 1, picbuf
            imgHeader = Left(Mid(buf, i - 1, Len(buf) - i), 16)
        Close fNum
        TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256)) * 10
        If imgCount > 1 Then
            xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256)
            yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256)
            Load aImg(imgCount - 1)
            aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
            aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
        End If
        aImg(imgCount - 1).Tag = TimeWait
        aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
        Kill ("temp.gif")
    End If
    TotalFrames = aImg.Count - 1
    Exit Function
ErrHandler:
    MsgBox "Error No. " & Err.Number & " when reading file", vbCritical
    LoadGif = False
    On Error GoTo 0
End Function
