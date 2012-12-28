VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' clsTimer
'--------------------------------------------------------------------------------
Option Explicit


Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private StartNormal As Long

Private mState As Long
Private mElapsedMilliseconds As Long
Private mLoopsPerSecond As Long
Private mMillisecondsPerLoop As Double
Private memStart As MEMORYSTATUS
Private memStop As MEMORYSTATUS
Private mBytesPerLoop As Long


'$ PROBHIDE NO_ERROR_HANDLER BEGIN

Private Sub Class_Initialize()
    ' Set the state to uninitialized
    State = -1
End Sub


Private Property Let State(ByVal Value As Long)
    ' Set the state
    If Value = 1 And mState <> 0 Then
        MsgBox "You must start the timer before stopping it!", vbCritical
    Else
        mState = Value
    End If
End Property


Public Sub StartTimer()
    ' Start the timer
    StartNormal = timeGetTime
    State = 0
    GlobalMemoryStatus memStart
End Sub


Public Sub StopTimer(ByVal Loops As Long)
    ' Stop the timer
    Dim startm As Long
    Dim stopm As Long
    Dim mem As Long
    Dim divisor As Long
    
    State = 1
    GlobalMemoryStatus memStop
    startm = (memStart.dwTotalPageFile - memStart.dwAvailPageFile)
    stopm = (memStop.dwTotalPageFile - memStop.dwAvailPageFile)
    mem = stopm - startm
    mBytesPerLoop = mem \ Loops
    
    ' Calculate elapsed milliseconds
    mElapsedMilliseconds = timeGetTime - StartNormal
    divisor = mElapsedMilliseconds
    If divisor = 0 Then
        mLoopsPerSecond = -1
        mMillisecondsPerLoop = -1
    Else
        mLoopsPerSecond = Loops / divisor * 1000
        mMillisecondsPerLoop = Loops / divisor
    End If
    
    Exit Sub
End Sub


Public Property Get ElapsedMilliseconds() As Long   '$ PROBHIDE DEAD
    ' Display elapsed milliseconds
    If mState = 0 Then
        ElapsedMilliseconds = timeGetTime - StartNormal
    Else
        ElapsedMilliseconds = mElapsedMilliseconds
    End If
End Property


Public Property Get LoopsPerSecond() As Long
    ' Return the loops per second
    LoopsPerSecond = mLoopsPerSecond
End Property


Public Property Get MillisecondsPerLoop() As Double '$ PROBHIDE DEAD
    ' Return the milliseconds per loop
    MillisecondsPerLoop = mMillisecondsPerLoop
End Property


Public Property Get BytesPerLoop() As Long
    ' Return the bytes per loop
    BytesPerLoop = mBytesPerLoop
End Property

'$ END
