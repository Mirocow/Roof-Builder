Attribute VB_Name = "HardwareFingerPrint"
 Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
  Private Const MAX_FILENAME_LEN = 256

  Dim ReadRegistry_Entry As String
  Dim tempBuildKey As String
  Dim ReadDays_In_Use As String
  Dim ReadDays_One As String
Private Function DriveSerial(ByVal sDrv As String) As Long
     Dim Str As String * MAX_FILENAME_LEN
     Dim str2 As String * MAX_FILENAME_LEN
     Dim a As Long
     Dim b As Long
     Call GetVolumeInformation(sDrv & ":\", Str, MAX_FILENAME_LEN, RetVal, a, b, str2, MAX_FILENAME_LEN)
     DriveSerial = RetVal
End Function
Public Property Get GetSystemSerial() As String
GetSystemSerial = DriveSerial("c")
End Property

