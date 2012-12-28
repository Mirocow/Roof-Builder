VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' clsTimings
'--------------------------------------------------------------------------------
Option Explicit

Public AddsPerSecond As Long
Public ExistsPerSecond As Long
Public GetsPerSecond As Long
Public SetsPerSecond As Long
Public DeletesPerSecond As Long
Public KeysPerSecond As Long
Public LookupsPerSecond As Long
Public BytesPerLoop As Long

