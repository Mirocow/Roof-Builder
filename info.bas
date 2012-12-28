Attribute VB_Name = "info"
Option Explicit

Private Const VER_PLATFORM_WIN32S = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Public Type OSVERSIONINFO
    OSVSize         As Long         'size, in bytes, of this data structure
    dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
    dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
    dwBuildNumber   As Long         'NT: build number of the OS
    'Win9x: build number of the OS in low-order word.
    '       High-order word contains major & minor ver nos.
    PlatformID      As Long         'Identifies the operating system platform.
    szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"                                'Win9x: 'arbitrary additional information'
End Type

' Per ottenere la versione del SO
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                                                                             ByVal pVolumeNameBuffer As String, _
                                                                                             ByVal nVolumeNameSize As Long, _
                                                                                             lpVolumeSerialNumber As Long, _
                                                                                             lpMaximumComponentLength As Long, _
                                                                                             lpFileSystemFlags As Long, _
                                                                                             ByVal lpFileSystemNameBuffer As String, _
                                                                                             ByVal nFileSystemNameSize As Long)
'    Const MAX_FILENAME_LEN = 256

Public Function Diskinfo() As String
    Dim SerialNum As Long
    Dim VolNameBuf As String
    Dim FileSysNameBuf As String
        On Error GoTo ERR
        VolNameBuf = String(255, Chr(0))
        FileSysNameBuf = String(255, Chr$(0))
        GetVolumeInformation "c:\", VolNameBuf, _
Len(VolNameBuf), SerialNum, 0, 0, _
FileSysNameBuf, Len(FileSysNameBuf)
        Diskinfo = Right("00000000" & Hex(SerialNum), 8) & Chr(1) & Abs(SerialNum)
        Exit Function
ERR:
        Diskinfo = ""
End Function


'Public Function osinfo()
'  Dim SWbemSet(6) As SWbemObjectSet
'  Dim SWbemObj As SWbemObject
'  Dim varObjectToId(6) As String
'  Dim varSerial(6) As String
'  Dim varSerials As String
'  Dim i, j As Integer
'  On Error Resume Next
'
'  varObjectToId(1) = "Win32_Processor,Name"
'  varObjectToId(2) = "Win32_Processor,Manufacturer"
'  varObjectToId(3) = "Win32_Processor,ProcessorId"
''  varObjectToId(4) = "Win32_BaseBoard,SerialNumber"
''  varObjectToId(5) = "Win32_BaseBoard,manufacturer"
''  varObjectToId(6) = "Win32_Baseboard,product"
'  varObjectToId(4) = "Win32_BIOS,Manufacturer"
'  varObjectToId(5) = "Win32_OperatingSystem,SerialNumber"
'  varObjectToId(6) = "Win32_OperatingSystem,Caption"
''  varObjectToId(7) = "Win32_DiskDrive,Model"
'
'  For i = 1 To 6
'    Set SWbemSet(i) = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf(Split(varObjectToId(i), ",")(0))
'    varSerial(i) = ""
'    For Each SWbemObj In SWbemSet(i)
'      varSerial(i) = SWbemObj.Properties_(Split(varObjectToId(i), ",")(1)) 'Property value
'      varSerial(i) = Trim(varSerial(i))
'      If Len(varSerial(i)) < 1 Then varSerial(i) = "Unknown value"
'    Next
''    Text1(i) = varSerial(i)
'     varSerials = varSerials & varSerial(i) & vbNewLine
'  Next
'
'  Text1.Text = varSerials
'End Function


Public Function CPinfo() As String
    Dim strObject

        On Error GoTo ERR

    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer As String, strUser, strPassword, name

        Set objLocator = CreateObject("WbemScripting.SWbemLocator")

        strComputer = "."

        Set objWMIService = objLocator.ConnectServer(strComputer, "root\cimv2")

        'Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
        'Set objWMIService = objLocator.ConnectServer(strComputer, "rootcimv2")

        objWMIService.Security_.impersonationlevel = 3

        Set colItems = objWMIService.ExecQuery("Select * From Win32_Processor")

        CPinfo = colItems.count & Chr(1)

        For Each objItem In colItems
            CPinfo = CPinfo & CStr(objItem.ProcessorId) & Chr(1)
        Next

        Set colItems = Nothing
        Set objWMIService = Nothing

        Exit Function
ERR:
        CPinfo = ""
End Function


Public Function MBinfo() As String

    'RETRIEVES SERIAL NUMBER OF MOTHERBOARD
    'IF THERE IS MORE THAN ONE MOTHERBOARD, THE SERIAL
    'NUMBERS WILL BE DELIMITED BY COMMAS

    'YOU MUST HAVE WMI INSTALLED AND A REFERENCE TO
    'Microsoft WMI Scripting Library IS REQUIRED

    Dim objs As Object

    Dim Obj As Object
    Dim WMI As Object
    Dim sAns As String

        On Error GoTo ERR

    Dim strComputer As String

        strComputer = "."

        Set WMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        'GetObject ("WinMgmts:")
        Set objs = WMI.InstancesOf("Win32_BaseBoard")
        For Each Obj In objs
            sAns = sAns & Obj.SerialNumber
            If sAns < objs.count Then sAns = sAns & ","
        Next

        Set WMI = Nothing
        Set Obj = Nothing

        MBinfo = sAns

        Exit Function
ERR:
        MBinfo = ""
End Function


Public Function OSinfo() As String
    #If Win32 Then
    Dim OSV As OSVERSIONINFO
        OSV.OSVSize = Len(OSV)
        If GetVersionEx(OSV) = 1 Then
            Select Case OSV.PlatformID
                Case VER_PLATFORM_WIN32S: OSinfo = "32s"
                Case VER_PLATFORM_WIN32_NT:
                    Select Case OSV.dwVerMajor
                        Case 3:
                            Select Case OSV.dwVerMinor
                                Case 0:  OSinfo = "NT3"
                                Case 1:  OSinfo = "NT3.1"
                                Case 5:  OSinfo = "NT3.5"
                                Case 51: OSinfo = "NT3.51"
                            End Select

                        Case 4: OSinfo = "NT 4"
                        Case 5:
                            Select Case OSV.dwVerMinor
                                Case 0:  OSinfo = "Win2000"
                                Case 1:  OSinfo = "WinXP"
                                Case 2:  OSinfo = "Win2k3" ' Windows Server 2003
                            End Select

                        Case 6: OSinfo = "Vista" 'Vista
                        Case Else
                            OSinfo = OSV.dwVerMajor & "." & OSV.dwVerMinor
                    End Select

                Case VER_PLATFORM_WIN32_WINDOWS:
                    Select Case OSV.dwVerMinor
                        Case 0:
                            If OSV.dwBuildNumber = 950 Then
                                OSinfo = "Windows 95"
                            Else
                                ' 1111 for OSR 2. For OSR 2.5 = ???
                                OSinfo = "Windows 95 OSR2"
                            End If

                        Case 90:   OSinfo = "ME"
                        Case Else:
                            If OSV.dwBuildNumber = 1998 Then
                                OSinfo = "Windows 98"
                            Else
                                OSinfo = "Windows 98 Release 2"
                            End If

                    End Select

            End Select

        End If

        #Else
            OSinfo = "3x"
        #End If
    
        Exit Function
ERR:
        OSinfo = ""
End Function

