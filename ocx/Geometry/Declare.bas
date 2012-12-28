Attribute VB_Name = "Declare"
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

