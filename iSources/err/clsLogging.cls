VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'***************************************************************************************
'*  EXTREME! Logging and Event Management Class ver. 1.3.0                             *
'*                                                                                     *
'*  Created:     December 05, 2005                                                     *
'*  Revised:     January 09, 2006                                                      *
'*  Purpose:     Comprehensive Application Logging and Managment Interface             *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.1.0                                                                 *
'*  Compile:     Native                                                                *
'*  Referenced:  mHandler, frmEvent                                                    *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************


'** List of Exposed Functions **

'** Handler_Start   - Start Error Handling Engine
'** Handler_Stop    - Stop Error Handling Engine
'** Log_Commit      - Write contents of application Log (automatic on class terminate)
'** Log_Data        - Write a Single Line to Application Log
'** Log_Event       - Log an Application Event
'** Log_Start       - Start Application Logging Engine
'** Log_SubHeader   - Create a List Sub Header
'** Log_Text        - Send Log to TextBox
'** Report_Error    - Write an Error Log
'** Report_Mail     - Write an Email Report
'** Report_Web      - Write a Web Report
'** Report_Note     - Display log in Text Editor


'** List of Exposed Properties **

'** Application Log **
'** OnStart         - Application Log Activation (boolean)
'** AppCpr          - Application Log Copyright String (string)
'** Append          - New or Appended Log File (boolean)
'** AppFtr          - Application Log Footer (string)
'** AppMsg          - Application Log Message (string)
'** AppTitle        - Application Log Title (string)
'** CssBackColor    - Style backcolor property (string) - optional
'** CssForeColor    - Style forecolor property (string) - optional
'** CssFont         - Style font property (string) - optional
'** CssFontSize     - Style body font size (integer)
'** CssHeader       - Style use header property (boolean) - optional
'** CssHeadColor    - Style header color property (string) - optional
'** CssHeadSize     - Style header size property (integer) - optional
'** CssStyle        - Style Turns feature on (boolean)
'** LogPath         - Application Log Path\Name (string) -optional: ex. logs\myapp.log
'** WebPath         - Web Log Path\Name (string) -optional: ex. logs\log.htm
'** WebImage        - Web Log Title Image Path\Name (string) ex. logs\title.gif
'**
'** Data Collection (WMI) switches **
'** DataCmp         - Computer Data (boolean)
'** DataOS          - OS Data (boolean)
'** DataPrc         - Process Data (boolean)
'** DataSft         - Software Data (boolean)
'** DataSrv         - Service Data (boolean)
'**
'** Error Logging **
'** EHandler        - Error Handler Switch (boolean)
'** ELocale         - Relative Error Position (string)
'** EData           - Module/Routine Information (string)
'** ErrPath         - Error Log Path\Name (string) -optional: ex. logs\error.log
'** EResume         - Resume Without User Notification (boolean)
'** EDump           - Dump Context Memory to Error Log (boolean)
'**
'** Mail Settings **
'** MailCc          - Mail CC (string)
'** MailSubject     - Mail Subject (string)
'** MailTo          - Mail To (string)


'** Original Logging Class credit goes out to: Richard Allsebrook  richardallsebrook@earlybirdmarketing.com
'** WMI routines based on work by: L.J. Johnson, http://mvps.org/st-software/Ask_NT_Pro.htm - dude loves his wmi!
'** Some ideas borrowed from Ulli's GPF Interveptor project - whatever happened to Ulli?
'**
'** The great thing about WMI is it works across all platforms, and is installed by default on ME on up.
'** You may want to test for 98, and shell a build with your install script if it is absent..
'** M$ update can be had at: http://www.microsoft.com/downloads/details.aspx?FamilyId=98A4C5BA-337B-4E92-8C18-A63847760EA5&displaylang=en
'**
'** There are several strong indicators of well written software to the user: the interface, speed and functionality
'** but perhaps some of the most important - but often overlooked queues to a good design, are how it
'** handles errors, and logs event processing.
'** I looked at a number of software packages that are in the same vein as the software I am currently developing
'** and looked at their logging features. At first I was shocked to see how much information they were commiting to
'** usage logs. Have you ever seen an Ad-aware or Spybot log? There is a lot of system information being logged,
'** along with a thorough accounting of event processes.
'** I hooked them and forced them to crash, but to my suprise, it appears neither of these had a good GPF logging
'** system -that- is an oversight. There is nothing worse then getting vague complaints from EUs, and trying to fish
'** the information out of them with drawn out, (and pointless/frustrating), email dialogs.
'** So I built a comprehensive logging engine - including GPF protection, error origination and description
'** data, application event tracking, and options to display as text or a web page.
'**
'** Update - December 8, 2005
'** Spotted a couple small bugs and fixed. Added wmi bypass on class init, (if wmi object load fails)
'** Set optional fixed paths for logs if path properties are empty, will use application local paths.
'** Added embedded css style option for web log, path property for title image.
'**
'** For a comment, (or a job..) email: steppenwolfe_2000@yahoo.com
'** Check it out.. if you spot problems or think of improvements, post a comment..

'** Cheers
'** John

Option Explicit

'//shellexecute
Private Const SW_NORMAL     As Long = 1
Private Const SW_MAXIMIZED  As Long = 3
Private Const SW_DEFAULT    As Long = 10

'//property vars
Private bOnStart            As Boolean
Private bAppend             As Boolean
Private bDataSft            As Boolean
Private bDataCmp            As Boolean
Private bDataOS             As Boolean
Private bDataSrv            As Boolean
Private bDataPrc            As Boolean
Private bEHandler           As Boolean
Private dTime               As Date
Private lPaddress           As Long
Private objWmiRoot          As Object
Private sLogPath            As String
Private sErrPath            As String
Private sWebPath            As String
Private sWebImage           As String
Private sMailPath           As String
Private sAppTitle           As String
Private sAppMsg             As String
Private sAppCpr             As String
Private sAppFtr             As String
Private sMailTo             As String
Private sMailCc             As String
Private sMailSubject        As String
Private bCssStyle           As Boolean
Private sCssBackColor       As String
Private sCssForeColor       As String
Private sCssFont            As String
Private iCssFontSize        As Integer
Private bCssHeader          As String
Private sCssHeadColor       As String
Private iCssHeadSize        As Integer
'//internal error record
'//use this to track module/routine
'//source of error
Private sELocale            As String
Private sEData              As String
'//handler switches
Private bEResume            As Boolean
Private bEDump              As Boolean
'//data collections
Private cLogData            As New Collection
Private cLogFooter          As New Collection
Private cLogDetail          As New Collection
Private cLogHeader          As New Collection

'//registry constants
'Public Enum HKEY_Type
'    HKEY_CLASSES_ROOT = &H80000000
'    HKEY_CURRENT_USER = &H80000001
'    HKEY_LOCAL_MACHINE = &H80000002
'    HKEY_USERS = &H80000003
'    HKEY_PERFORMANCE_DATA = &H80000004
'    HKEY_CURRENT_CONFIG = &H80000005
'    HKEY_DYN_DATA = &H80000006
'End Enum

Private Const KEY_ALL_ACCESS As Long = &HF003F

'//desktop hwnd api
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'//exception handler api
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
'//shell execute api
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long
'//registry api
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'** logging options **
'//log title
Public Property Get AppTitle() As String
    AppTitle = sAppTitle
End Property

Public Property Let AppTitle(ByVal PropVal As String)
    sAppTitle = PropVal
End Property
'//subheader message
Public Property Get AppMsg() As String
    AppMsg = sAppMsg
End Property

Public Property Let AppMsg(ByVal PropVal As String)
    sAppMsg = PropVal
End Property
'//subheader copyright
Public Property Get AppCpr() As String
    AppCpr = sAppCpr
End Property

Public Property Let AppCpr(ByVal PropVal As String)
    sAppCpr = PropVal
End Property
'//footer message
Public Property Get AppFtr() As String
    AppFtr = sAppFtr
End Property

Public Property Let AppFtr(ByVal PropVal As String)
    sAppFtr = PropVal
End Property

'//log path
Public Property Get LogPath() As String
    LogPath = sLogPath
End Property

Public Property Let LogPath(NewValue As String)
    sLogPath = NewValue
End Property

'//error log path
Public Property Get ErrPath() As String
    ErrPath = sErrPath
End Property

Public Property Let ErrPath(NewValue As String)
    sErrPath = NewValue
End Property

'//web log path
Public Property Get WebPath() As String
    WebPath = sWebPath
End Property

Public Property Let WebPath(NewValue As String)
    sWebPath = NewValue
End Property

'//web log image path
Public Property Get WebImage() As String
    WebImage = sWebImage
End Property

Public Property Let WebImage(NewValue As String)
    sWebImage = NewValue
End Property

'//email path
Public Property Get MailPath() As String
    MailPath = sMailPath
End Property

Public Property Let MailPath(NewValue As String)
    sMailPath = NewValue
End Property

'** wmi options **
'//software
Public Property Get DataSft() As Boolean
    DataSft = bDataSft
End Property

Public Property Let DataSft(ByVal PropVal As Boolean)
    bDataSft = PropVal
End Property
'//computer data
Public Property Get DataCmp() As Boolean
    DataCmp = bDataCmp
End Property

Public Property Let DataCmp(ByVal PropVal As Boolean)
    bDataCmp = PropVal
End Property
'//os data
Public Property Get DataOS() As Boolean
    DataOS = bDataOS
End Property

Public Property Let DataOS(ByVal PropVal As Boolean)
    bDataOS = PropVal
End Property
'//service data
Public Property Get DataSrv() As Boolean
    DataSrv = bDataSrv
End Property

Public Property Let DataSrv(ByVal PropVal As Boolean)
    bDataSrv = PropVal
End Property
'//process data
Public Property Get DataPrc() As Boolean
    DataPrc = bDataPrc
End Property

Public Property Let DataPrc(ByVal PropVal As Boolean)
    bDataPrc = PropVal
End Property

'//load on class start
Public Property Get OnStart() As Boolean
    OnStart = bOnStart
End Property

Public Property Let OnStart(ByVal PropVal As Boolean)
    bOnStart = PropVal
End Property

'//append or create new log
Public Property Get Append() As Boolean
    Append = bAppend
End Property

Public Property Let Append(ByVal PropVal As Boolean)
    bAppend = PropVal
End Property

'** exception handler properties **
'//engage handler
Public Property Get EHandler() As Boolean
    EHandler = bEHandler
End Property

Public Property Let EHandler(ByVal PropVal As Boolean)
    bEHandler = PropVal
End Property

'** mail properties **
'//send to
Public Property Get MailTo() As String
    MailTo = sMailTo
End Property

Public Property Let MailTo(ByVal PropVal As String)
    sMailTo = PropVal
End Property
'//mail cc
Public Property Get MailCc() As String
    MailCc = sMailCc
End Property

Public Property Let MailCc(ByVal PropVal As String)
    sMailCc = PropVal
End Property
'//mail subject
Public Property Get MailSubject() As String
    MailSubject = sMailSubject
End Property

Public Property Let MailSubject(ByVal PropVal As String)
    sMailSubject = PropVal
End Property

'** style sheet properties
'//style switch
Public Property Get CssStyle() As Boolean
    CssStyle = bCssStyle
End Property

Public Property Let CssStyle(ByVal PropVal As Boolean)
    bCssStyle = PropVal
End Property

'//style backcolor
Public Property Get CssBackColor() As String
    CssBackColor = sCssBackColor
End Property

Public Property Let CssBackColor(ByVal PropVal As String)
    sCssBackColor = PropVal
End Property

'//style forecolor
Public Property Get CssForeColor() As String
    CssForeColor = sCssForeColor
End Property

Public Property Let CssForeColor(ByVal PropVal As String)
    sCssForeColor = PropVal
End Property

'//style font size
Public Property Get CssFontSize() As Integer
    CssFontSize = iCssFontSize
End Property

Public Property Let CssFontSize(ByVal PropVal As Integer)
    iCssFontSize = PropVal
End Property

'//style forecolor
Public Property Get CssFont() As String
    CssFont = sCssFont
End Property

Public Property Let CssFont(ByVal PropVal As String)
    sCssFont = PropVal
End Property

'//use header style
Public Property Get CssHeader() As Boolean
    CssHeader = bCssHeader
End Property

Public Property Let CssHeader(ByVal PropVal As Boolean)
    bCssHeader = PropVal
End Property

'//style header color
Public Property Get CssHeadColor() As String
    CssHeadColor = sCssHeadColor
End Property

Public Property Let CssHeadColor(ByVal PropVal As String)
    sCssHeadColor = PropVal
End Property

'//style header size
Public Property Get CssHeadSize() As Integer
    CssHeadSize = iCssHeadSize
End Property

Public Property Let CssHeadSize(ByVal PropVal As Integer)
    iCssHeadSize = PropVal
End Property

'** error tracking
'//error location ex module/routine/call - M1/R3/C3
Public Property Get ELocale() As String
    ELocale = sELocale
End Property

Public Property Let ELocale(ByVal PropVal As String)
    sELocale = PropVal
End Property

'//description of routine
Public Property Get EData() As String
    EData = sEData
End Property

Public Property Let EData(ByVal PropVal As String)
    sEData = PropVal
End Property

'** error handler switches
'//resume/ignore GPF
Public Property Get EResume() As Boolean
    EResume = bEResume
End Property

Public Property Let EResume(ByVal PropVal As Boolean)
    bEResume = PropVal
End Property

'//dump context memory
Public Property Get EDump() As Boolean
    EDump = bEDump
End Property

Public Property Let EDump(ByVal PropVal As Boolean)
    bEDump = PropVal
End Property

Private Sub Class_Initialize()
'//class load

Dim bTest As Boolean

On Error Resume Next
    
    '//data collections
    Set cLogData = New Collection
    Set cLogFooter = New Collection
    Set cLogDetail = New Collection
    Set cLogHeader = New Collection

On Error GoTo Handler

    '//initialize wmi object
    Set objWmiRoot = GetObject("winmgmts:{impersonationLevel=impersonate}")
    bTest = True
    
On Error GoTo 0
Exit Sub

'//test for wmi compatability
Handler:
    If bTest = False Then
        DataSft = False
        DataCmp = False
        DataOS = False
        DataSrv = False
        DataPrc = False
    End If
On Error GoTo 0

End Sub

Public Sub Log_Start()
'//set up the event log

    dTime = Now
    '//test/set paths properties
    '//app log
    If Len(LogPath) = 0 Then
        LogPath = App.Path & Chr$(92) & App.Title & Format$(dTime, "MM-DD-YYYY") & ".log"
    Else
        LogPath = App.Path & Chr$(92) & LogPath
    End If
    
    '//error log
    If Len(ErrPath) = 0 Then
        ErrPath = App.Path & Chr$(92) & App.Title & Chr$(45) & "-Error-" & Format$(dTime, "MM-DD-YYYY") & ".log"
    Else
        ErrPath = App.Path & Chr$(92) & ErrPath
    End If
    
    '//html log
    If Len(WebPath) = 0 Then
        WebPath = App.Path & Chr$(92) & App.Title & Chr$(45) & Format$(dTime, "MM-DD-YYYY") & ".htm"
    Else
        WebPath = App.Path & Chr$(92) & WebPath
    End If
    
    '//email log
    If Len(MailPath) = 0 Then
        MailPath = App.Path & Chr$(92) & App.Title & Chr$(45) & Format$(dTime, "MM-DD-YYYY") & ".eml"
    Else
        MailPath = App.Path & Chr$(92) & MailPath
    End If
    
    '//start logging
    If OnStart Then
        Log_Header
    End If

    '//start handler
    If EHandler Then
        Handler_Start
    End If

End Sub

Private Sub Log_Header()
'//build header

Dim i       As Integer
Dim sLsep   As String
Dim sHsep   As String
Dim x       As Integer
Dim y       As Integer
Dim sBuild  As String
Dim sCreate As String
Dim sLmrk   As String
Dim sRmrk   As String
Dim sPath   As String

    '//build strings
    sLmrk = Chr$(187) & Chr$(32) & Chr$(32)
    sBuild = "Build: " & App.Major & "." & App.Minor & "." & App.Revision
    sCreate = "Log File Created:  " & dTime
    sPath = "Path: " & App.Path & "\" & App.EXEName & ".exe"

    '//get longest string
    x = Len(sBuild)
    y = Len(sCreate)
    If y > x Then x = y
    y = Len(sPath)
    If y > x Then x = y
    y = Len(AppMsg)
    If y > x Then x = y
    y = Len(AppCpr)
    If y > x Then x = y
    y = Len(AppTitle)
    If y > x Then x = y

    '//spacer
    Log_Data ""

    '//build title frame
    For i = 1 To x + 6
        sHsep = sHsep + Chr$(187)
        sLsep = sLsep + Chr$(171)
    Next i

    '//add and frame data
    cLogHeader.Add sHsep
    For y = Len(AppTitle) To Len(sHsep) - 5
        sRmrk = sRmrk & Chr$(32)
    Next y
    cLogHeader.Add sLmrk & AppTitle & sRmrk & Chr$(171)
    sRmrk = vbNullString

    For y = Len(sBuild) To Len(sHsep) - 5
        sRmrk = sRmrk & Chr$(32)
    Next y
    cLogHeader.Add sLmrk & sBuild & sRmrk & Chr$(171)
    sRmrk = vbNullString

    For y = Len(sCreate) To Len(sHsep) - 5
        sRmrk = sRmrk & Chr$(32)
    Next y
    cLogHeader.Add sLmrk & sCreate & sRmrk & Chr$(171)
    sRmrk = vbNullString

    For y = Len(sPath) To Len(sHsep) - 5
        sRmrk = sRmrk & Chr$(32)
    Next y
    cLogHeader.Add sLmrk & sPath & sRmrk & Chr$(171)
    sRmrk = vbNullString

    If Len(AppMsg) > 0 Then
        For y = Len(AppMsg) To Len(sHsep) - 5
            sRmrk = sRmrk & Chr$(32)
        Next y
        cLogHeader.Add sLmrk & AppMsg & sRmrk & Chr$(171)
        sRmrk = vbNullString
    End If

    If Len(AppCpr) > 0 Then
        For y = Len(AppCpr) To Len(sHsep) - 5
            sRmrk = sRmrk & Chr$(32)
        Next y
        cLogHeader.Add sLmrk & AppCpr & sRmrk & Chr$(171)
        sRmrk = vbNullString
    End If

    '//end frame
    cLogHeader.Add sLsep

    '//add optional computer data
'    WMI_Control

End Sub

Public Sub Log_Data(ByVal sMessage As String)
'//add a single entry

    cLogData.Add sMessage

End Sub

Private Sub Log_Footer()
'//build footer

    With cLogFooter
        .Add ""
        .Add "»»»»»»»»»»»»»»»»»»»"
        .Add "Application Stopped"
        .Add "Run Time: " & DateDiff("s", dTime, Now) & " seconds"
        .Add ""
    End With

End Sub

Public Sub Log_Commit(Optional ByVal bDisplay As Boolean)
'//write the log file

Dim FF    As Integer
Dim vItem As Variant
Dim lHwnd   As Long

On Error GoTo Handler

    FF = FreeFile
    '//choice: new or append
    If Append Then
        Open LogPath For Append As #FF
    Else
        Open LogPath For Output As #FF
    End If
    '//add header
    For Each vItem In cLogHeader
        Print #FF, vItem
    Next vItem
    '//add data
    For Each vItem In cLogData
        Print #FF, vItem
    Next vItem
    '//add system details
    For Each vItem In cLogDetail
        Print #FF, vItem
    Next vItem
    '//write footer
    Log_Footer
    '//add footer
    For Each vItem In cLogFooter
        Print #FF, vItem
    Next vItem
    Close #FF
    
    If bDisplay Then
        lHwnd = GetDesktopWindow()
        ShellExecute lHwnd, "Open", LogPath, vbNullString, vbNullString, SW_NORMAL
    End If
    
    '//reset footer
    Set cLogFooter = Nothing
    Set cLogFooter = New Collection
    
Handler:

End Sub

Public Sub Log_SubHeader(ByVal sName As String)
'//underline and subheader

Dim sRmrk As String
Dim x     As Integer

    Log_Data ""
    
    For x = 1 To Len(sName)
        sRmrk = sRmrk & Chr$(187)
    Next x

    Log_Data sName
    Log_Data sRmrk
    Log_Data ""

End Sub

Public Sub Log_Event(ByVal sName As String, _
                     Optional ByVal sDesc1 As String, _
                     Optional ByVal sDesc2 As String, _
                     Optional ByVal sDesc3 As String, _
                     Optional ByVal sDesc4 As String, _
                     Optional ByVal sErr As String)
'//log an application event

    '//entry name
    cLogData.Add sName
    '//add data
    If Len(sDesc1) > 0 Then
        cLogData.Add sDesc1
    End If
    If Len(sDesc2) > 0 Then
        cLogData.Add sDesc2
    End If
    If Len(sDesc3) > 0 Then
        cLogData.Add sDesc3
    End If
    If Len(sDesc4) > 0 Then
        cLogData.Add sDesc4
    End If
    If Len(sErr) > 0 Then
        cLogData.Add sErr
    End If
    '//spacer
    cLogData.Add ""

End Sub


Private Sub Class_Terminate()
'//clean up

    '//write the log
    Log_Commit

    '//clean up
    Set cLogData = Nothing
    Set cLogFooter = Nothing
    Set cLogDetail = Nothing
    Set cLogHeader = Nothing
    Set objWmiRoot = Nothing

    '//stop handler
    If EHandler Then
        Handler_Stop
    End If

End Sub

Public Sub Log_Text(ByRef lText As Object)
'//send log to textbox

Dim vItem As Variant

On Error Resume Next

    '//add header
    lText.Text = lText.Text & AppTitle & vbCrLf
    lText.Text = lText.Text & AppCpr & vbCrLf
    lText.Text = lText.Text & AppMsg & vbCrLf
    lText.Text = lText.Text & "Log Created: " & dTime & vbCrLf
    
    '//add data
    For Each vItem In cLogData
        lText.Text = lText.Text & vItem & vbCrLf
    Next vItem
    
    '//add data
    For Each vItem In cLogDetail
        lText.Text = lText.Text & vItem & vbCrLf
    Next vItem
    
    '//write footer
    Log_Footer
    '//add footer
    For Each vItem In cLogFooter
        lText.Text = lText.Text & vItem & vbCrLf
    Next vItem
    
    '//reset footer
    Set cLogFooter = Nothing
    Set cLogFooter = New Collection
    
On Error GoTo 0

End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           EXCEPTION HANDLER
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Function Pointer(ByVal lHandle As Long) As Long
'//helper function

    Pointer = lHandle

End Function

Public Sub Handler_Start()
'//start the exception handler

    lPaddress = SetUnhandledExceptionFilter(Pointer(AddressOf Exception_Handler))

End Sub

Public Sub Handler_Stop()
'//stop the exception handler

    SetUnhandledExceptionFilter &H0

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           WMI INSTRUMENTATION
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Private Sub WMI_Control()
''//wmi processing hub
'
'Dim vntRtn As Variant
'Dim lLoop  As Long
'Dim lData  As Long
'
'On Error Resume Next
'
'    '//software data
'    If DataSft Then
'        '//set header
'        vntRtn = Get_SoftData(objWmiRoot)
'        If Not IsEmpty(vntRtn) Then
'            With cLogDetail
'                .Add ""
'                .Add "Registered Software"
'                .Add "»»»»»»»»»»»»»»»»»»»"
'            End With
'            '//add data strings
'            For lLoop = 1 To UBound(vntRtn)
'                With cLogDetail
'                    .Add ""
'                    .Add vntRtn(lLoop)(1)
'                End With
'                For lData = 2 To UBound(vntRtn(1))
'                    cLogDetail.Add vntRtn(lLoop)(lData)
'                Next lData
'
'            Next lLoop
'        End If
'    End If
'
'    '//computer data
'    If DataCmp Then
'        vntRtn = Get_CompData(objWmiRoot)
'        If Not IsEmpty(vntRtn) Then
'            With cLogDetail
'                .Add ""
'                .Add "Computer Configuration"
'                .Add "»»»»»»»»»»»»»»»»»»»»»»"
'                .Add ""
'            End With
'            For lLoop = 1 To UBound(vntRtn, 2)
'                cLogDetail.Add vntRtn(0, lLoop) & ": " & vntRtn(1, lLoop)
'            Next lLoop
'        End If
'    End If
'
'    '//os data
'    If DataOS Then
'        vntRtn = Get_OsData(objWmiRoot)
'        If Not IsEmpty(vntRtn) Then
'            With cLogDetail
'                .Add ""
'                .Add "OS Configuration"
'                .Add "»»»»»»»»»»»»»»»»"
'                .Add ""
'            End With
'            For lLoop = 1 To UBound(vntRtn, 2)
'                cLogDetail.Add vntRtn(0, lLoop) & ": " & vntRtn(1, lLoop)
'            Next lLoop
'        End If
'    End If
'
'    '//servive info
'    If DataSrv Then
'        vntRtn = Get_SrvData(objWmiRoot)
'        If Not IsEmpty(vntRtn) Then
'            With cLogDetail
'                .Add ""
'                .Add "Service Information"
'                .Add "»»»»»»»»»»»»»»»»»»»"
'                .Add ""
'            End With
'            For lLoop = 1 To UBound(vntRtn)
'                With cLogDetail
'                    .Add ""
'                    .Add vntRtn(lLoop)(1)
'                End With
'                For lData = 2 To UBound(vntRtn(1))
'                    cLogDetail.Add vntRtn(lLoop)(lData)
'                Next lData
'            Next lLoop
'        End If
'    End If
'
'    '//process info
'    If DataPrc Then
'        vntRtn = Get_PrcData(objWmiRoot)
'        If Not IsEmpty(vntRtn) Then
'            With cLogDetail
'                .Add ""
'                .Add "Process Information"
'                .Add "»»»»»»»»»»»»»»»»»»»"
'            End With
'            For lLoop = 1 To UBound(vntRtn)
'                With cLogDetail
'                    .Add ""
'                    .Add vntRtn(lLoop)(1)
'                End With
'                For lData = 2 To UBound(vntRtn(1))
'                    cLogDetail.Add vntRtn(lLoop)(lData)
'                Next lData
'            Next lLoop
'        End If
'    End If
'
'On Error GoTo 0
'
'End Sub

'Private Function Get_SoftData(ByVal objWbm As SWbemServices) As Variant
''//interrogate software data
'
'Dim vTemp    As Variant
'Dim objISet  As SWbemObjectSet
'Dim objISoft As SWbemObject
'Dim objProp  As SWbemProperty
'Dim lIndex   As Long
'Dim vData    As Variant
'
'On Error Resume Next
'
'    Set objISet = objWbm.InstancesOf("Win32_Product")
'    ReDim vData(1 To objISet.count) As Variant
'
'    For Each objISoft In objISet
'        ReDim vTemp(1 To 7) As String
'
'        For Each objProp In objISoft.Properties_
'            With objProp
'                Select Case .name
'                Case "Description"
'                    vTemp(1) = "Description: " & .value
'                Case "Version"
'                    vTemp(2) = "Version: " & .value
'                Case "InstallDate"
'                    vTemp(3) = "Installed: " & Trans_Soft(.value)
'                Case "PackageCache"
'                    vTemp(4) = "Cache: " & .value
'                Case "InstallState"
'                    Select Case .value
'                    Case -6
'                        vTemp(5) = "State: Bad Configuration"
'                    Case -2
'                        vTemp(5) = "State: Invalid Argument"
'                    Case -1
'                        vTemp(5) = "State: Unknown Package"
'                    Case 1
'                        vTemp(5) = "State: Advertised"
'                    Case 2
'                        vTemp(5) = "State: Absent"
'                    Case 5
'                        vTemp(5) = "State: Installed"
'                    End Select
'                Case "IdentifyingNumber"
'                    vTemp(6) = "Clsid: " & .value
'                Case "Vendor"
'                    vTemp(7) = "Vendor: " & .value
'                Case Else
'                End Select
'            End With
'        Next objProp
'        If Len(Trim$(vTemp(1))) > 0 Then
'            lIndex = lIndex + 1
'            vData(lIndex) = vTemp
'        End If
'        DoEvents
'    Next objISoft
'    Get_SoftData = vData
'
'On Error GoTo 0
'
'End Function
'
'Private Function Get_CompData(ByVal objWbm As SWbemServices) As Variant
''//interrogate computer data
'
'Dim objSysSet As SWbemObjectSet
'Dim objSystem As SWbemObject
'Dim objProp   As SWbemProperty
'Dim vData     As Variant
'
'On Error Resume Next
'
'    Set objSysSet = objWbm.InstancesOf("Win32_ComputerSystem")
'    ReDim vData(0 To 1, 1 To 24) As Variant
'
'    For Each objSystem In objSysSet
'        For Each objProp In objSystem.Properties_
'            With objProp
'                Select Case .name
'                Case "TotalPhysicalMemory"
'                    vData(0, 1) = "Total Physical Memory"
'                    vData(1, 1) = Trans_Calc(.value) & " K"
'                Case "SystemType"
'                    vData(0, 2) = "System Type"
'                    vData(1, 2) = .value
'                Case "Domain"
'                    vData(0, 3) = "Domain"
'                    vData(1, 3) = .value
'                Case "NumberOfProcessors"
'                    vData(0, 4) = "Number of Processors"
'                    vData(1, 4) = .value
'                Case "PrimaryOwnerName"
'                    vData(0, 5) = "Primary Owner Name"
'                    vData(1, 5) = .value
'                Case "Description"
'                    vData(0, 6) = "Description"
'                    vData(1, 6) = .value
'                Case "EnableDaylightSavingsTime"
'                    vData(0, 7) = "Daylight Savings Enabled"
'                    vData(1, 7) = CStr(.value)
'                Case "UserName"
'                    vData(0, 8) = "Current User Name"
'                    vData(1, 8) = .value
'                Case "AdminPasswordStatus"
'                    vData(0, 9) = "Admin Password Status"
'                    Select Case .value
'                    Case 1
'                        vData(1, 9) = "Disabled"
'                    Case 2
'                        vData(1, 9) = "Enabled"
'                    Case 3
'                        vData(1, 9) = "Not Implemented"
'                    Case 4
'                        vData(1, 9) = "Unknown"
'                    End Select
'                Case "AutomaticResetBootOption"
'                    vData(0, 10) = "Auto Reset Boot Option"
'                    vData(1, 10) = CStr(.value)
'                Case "AutomaticResetCapability"
'                    vData(0, 11) = "Auto Reset Capability"
'                    vData(1, 11) = CStr(.value)
'                Case "BootROMSupported"
'                    vData(0, 12) = "Boot ROM Supported"
'                    vData(1, 12) = CStr(.value)
'                Case "InfraredSupported"
'                    vData(0, 13) = "Infrared Supported"
'                    vData(1, 13) = CStr(.value)
'                Case "NetworkServerModeEnabled"
'                    vData(0, 14) = "Network Server Mode"
'                    vData(1, 14) = CStr(.value)
'                Case "DomainRole"
'                    vData(0, 15) = "Domain Role"
'                    Select Case .value
'                    Case 0
'                        vData(1, 15) = "Standalone Workstation"
'                    Case 1
'                        vData(1, 15) = "Member Workstation"
'                    Case 2
'                        vData(1, 15) = "Standalone Server"
'                    Case 3
'                        vData(1, 15) = "Member Server"
'                    Case 4
'                        vData(1, 15) = "Backup DC"
'                    Case 5
'                        vData(1, 15) = "Primary DC"
'                    Case Else
'                        vData(1, 15) = "Unknown"
'                    End Select
'                Case "PartOfDomain"
'                    vData(0, 16) = "Part of Domain"
'                    vData(1, 16) = CStr(.value)
'                Case "BootupState"
'                    vData(0, 17) = "Bootup State"
'                    vData(1, 17) = .value
'                Case "FrontPanelResetStatus"
'                    vData(0, 18) = "Front Panel Reset Status"
'                    Select Case .value
'                    Case 0
'                        vData(1, 18) = "Disabled"
'                    Case 1
'                        vData(1, 18) = "Enabled"
'                    Case 2
'                        vData(1, 18) = "Not Implemented"
'                    Case 3
'                        vData(1, 18) = "Unknown"
'                    End Select
'                Case "KeyboardPasswordStatus"
'                    vData(0, 19) = "Keyboard Password Status"
'                    Select Case .value
'                    Case 0
'                        vData(1, 19) = "Disabled"
'                    Case 1
'                        vData(1, 19) = "Enabled"
'                    Case 2
'                        vData(1, 19) = "Not Implemented"
'                    Case 3
'                        vData(1, 19) = "Unknown"
'                    End Select
'                Case "PowerOnPasswordStatus"
'                    vData(0, 20) = "Power-on Password Status"
'                    Select Case .value
'                    Case 0
'                        vData(1, 20) = "Disabled"
'                    Case 1
'                        vData(1, 20) = "Enabled"
'                    Case 2
'                        vData(1, 20) = "Not Implemented"
'                    Case 3
'                        vData(1, 20) = "Unknown"
'                    End Select
'                Case "PowerSupplyState"
'                    vData(0, 21) = "Power Supply State"
'                    Select Case .value
'                    Case 1
'                        vData(1, 21) = "Other"
'                    Case 2
'                        vData(1, 21) = "Unknown"
'                    Case 3
'                        vData(1, 21) = "Save"
'                    Case 4
'                        vData(1, 21) = "Warning"
'                    Case 5
'                        vData(1, 21) = "Critical"
'                    Case 6
'                        vData(1, 21) = "Non-recoverable"
'                    End Select
'                Case "ThermalState"
'                    vData(0, 22) = "Thermal State"
'                    Select Case .value
'                    Case 1
'                        vData(1, 22) = "Other"
'                    Case 2
'                        vData(1, 22) = "Unknown"
'                    Case 3
'                        vData(1, 22) = "Safe"
'                    Case 4
'                        vData(1, 22) = "Warning"
'                    Case 5
'                        vData(1, 22) = "Critical"
'                    Case 6
'                        vData(1, 22) = "Non-recoverable"
'                    End Select
'                Case "ResetCapability"
'                    vData(0, 23) = "Reset Capability"
'                    Select Case .value
'                    Case 1
'                        vData(1, 23) = "Other"
'                    Case 2
'                        vData(1, 23) = "Unknown"
'                    Case 3
'                        vData(1, 23) = "Disabled"
'                    Case 4
'                        vData(1, 23) = "Enabled"
'                    Case 5
'                        vData(1, 23) = "Non-recoverable"
'                    End Select
'                Case "SystemStartupDelay"
'                    vData(0, 24) = "System Startup Delay (secs)"
'                    vData(1, 24) = .value
'                End Select
'            End With
'        Next objProp
'        DoEvents
'    Next objSystem
'    Get_CompData = vData
'
'On Error GoTo 0
'
'End Function
'
'Private Function Get_OsData(ByVal objWbm As SWbemServices) As Variant
''//interrogate operating system data
'
'Dim OSSet   As SWbemObjectSet
'Dim OS      As SWbemObject
'Dim objProp As SWbemProperty
'Dim vData   As Variant
'
'On Error Resume Next
'
'    Set OSSet = objWbm.InstancesOf("Win32_OperatingSystem")
'    ReDim vData(0 To 1, 1 To 32) As Variant
'
'    For Each OS In OSSet
'        For Each objProp In OS.Properties_
'            With objProp
'                Select Case .name
'                Case "FreePhysicalMemory"
'                    vData(0, 1) = "Free Physical Memory"
'                    vData(1, 1) = Trans_Calc(.value) & " K"
'                Case "FreeVirtualMemory"
'                    vData(0, 2) = "Free Virtual Memory"
'                    vData(1, 2) = Trans_Calc(.value) & " K"
'                Case "TotalVirtualMemorySize"
'                    vData(0, 3) = "Total Virtual Memmory Size"
'                    vData(1, 3) = Trans_Calc(.value) & " K"
'                Case "TotalVisibleMemorySize"
'                    vData(0, 4) = "Total Visible Memory Size"
'                    vData(1, 4) = Trans_Calc(.value) & " K"
'                Case "SizeStoredInPagingFiles"
'                    vData(0, 5) = "Size Stored in Paging File(s)"
'                    vData(1, 5) = Trans_Calc(.value) & " K"
'                Case "FreeSpaceInPagingFiles"
'                    vData(0, 6) = "Free Space in Paging File(s)"
'                    vData(1, 6) = Trans_Calc(.value) & " K"
'                Case "CodeSet"
'                    vData(0, 7) = "Code Set"
'                    vData(1, 7) = .value
'                Case "CountryCode"
'                    vData(0, 8) = "Country Code"
'                    vData(1, 8) = .value
'                Case "Locale"
'                    vData(0, 9) = "Locale"
'                    vData(1, 9) = .value
'                Case "InstallDate"
'                    vData(0, 10) = "OS Installation Date"
'                    vData(1, 10) = Trans_Date(.value)
'                Case "LastBootUpTime"
'                    vData(0, 11) = "Last Boot-up Time"
'                    vData(1, 11) = Trans_Date(.value)
'                Case "Organization"
'                    vData(0, 12) = "Organization"
'                    vData(1, 12) = .value
'                Case "Version"
'                    vData(0, 13) = "Version"
'                    vData(1, 13) = .value
'                Case "WindowsDirectory"
'                    vData(0, 14) = "Windows Directory"
'                    vData(1, 14) = .value
'                Case "SystemDirectory"
'                    vData(0, 15) = "System Directory"
'                    vData(1, 15) = .value
'                Case "RegisteredUser"
'                    vData(0, 16) = "Registered User"
'                    vData(1, 16) = .value
'                Case "NumberOfUsers"
'                    vData(0, 17) = "Number of Users"
'                    vData(1, 17) = .value
'                Case "NumberOfProcesses"
'                    vData(0, 18) = "Number of Processes"
'                    vData(1, 18) = .value
'                Case "BuildType"
'                    vData(0, 19) = "Build Type"
'                    vData(1, 19) = .value
'                Case "BuildNumber"
'                    vData(0, 20) = "Build Number"
'                    vData(1, 20) = .value
'                Case "Caption"
'                    vData(0, 21) = "Caption"
'                    vData(1, 21) = .value
'                Case "Primary"
'                    vData(0, 22) = "Primary OS?"
'                    vData(1, 22) = CStr(.value)
'                Case "OSType"
'                    vData(0, 23) = "OS Type"
'                    Select Case .value
'                    Case 15
'                        vData(1, 23) = "Win 3x"
'                    Case 16
'                        vData(1, 23) = "Windows 95"
'                    Case 17
'                        vData(1, 23) = "Windows 98"
'                    Case 18
'                        vData(1, 23) = "NT Family"
'                    Case 19
'                        vData(1, 23) = "Win CE"
'                    Case Else
'                        vData(1, 23) = "Unknown"
'                    End Select
'                Case "CSDVersion"
'                    vData(0, 24) = "SP Text"
'                    vData(1, 24) = .value
'                Case "ServicePackMajorVersion"
'                    vData(0, 25) = "SP, Major Ver"
'                    vData(1, 25) = .value
'                Case "ServicePackMinorVersion"
'                    vData(0, 26) = "SP, Minor Ver"
'                    vData(1, 26) = .value
'                Case "SystemDrive"
'                    vData(0, 27) = "System Drive"
'                    vData(1, 27) = .value
'                Case "BootDevice"
'                    vData(0, 28) = "Boot Device"
'                    vData(1, 28) = .value
'                Case "SerialNumber"
'                    vData(0, 29) = "Serial Number"
'                    vData(1, 29) = .value
'                Case "CurrentTimeZone"
'                    vData(0, 30) = "Minutes Offset GMT"
'                    vData(1, 30) = .value
'                Case "EncryptionLevel"
'                    vData(0, 31) = "Encryption Level"
'                    vData(1, 31) = .value
'                Case "ForegroundApplicationBoost"
'                    vData(0, 32) = "Foreground App Boost"
'                    Select Case .value
'                    Case 0
'                        vData(1, 32) = "None"
'                    Case 1
'                        vData(1, 32) = "Minimum"
'                    Case 2
'                        vData(1, 32) = "(Default) Maximum"
'                    End Select
'                End Select
'            End With
'        Next objProp
'        DoEvents
'    Next OS
'    Get_OsData = vData
'
'On Error GoTo 0
'
'End Function
'
'Private Function Get_SrvData(ByVal objWbm As SWbemServices) As Variant
''//interrogate service data
'
''** Note, interrogating system services will likely
''** raise a firewall or av alert on some systems
'
'Dim vTemp          As Variant
'Dim objServicesSet As SWbemObjectSet
'Dim objService     As SWbemObject
'Dim objProp        As SWbemProperty
'Dim lIndex         As Long
'Dim vData          As Variant
'
'On Error Resume Next
'
'    Set objServicesSet = objWbm.InstancesOf("Win32_Service")
'    ReDim vData(1 To objServicesSet.count) As Variant
'
'    For Each objService In objServicesSet
'        ReDim vTemp(1 To 13) As String
'        For Each objProp In objService.Properties_
'            With objProp
'                Select Case .name
'                Case "DisplayName"
'                    vTemp(1) = "DisplayName: " & .value
'                Case "Description"
'                    vTemp(2) = "Description: " & .value
'                Case "Started"
'                    vTemp(3) = "Started: " & CStr(.value)
'                Case "StartMode"
'                    vTemp(4) = "StartMode: " & .value
'                Case "State"
'                    vTemp(5) = "State: " & .value
'                Case "StartName"
'                    vTemp(6) = "StartName: " & .value
'                Case "ServiceType"
'                    vTemp(7) = "ServiceType: " & .value
'                Case "AcceptPause"
'                    vTemp(8) = "AcceptPause: " & CStr(.value)
'                Case "AcceptStop"
'                    vTemp(9) = "AcceptStop: " & CStr(.value)
'                Case "DesktopInteract"
'                    vTemp(10) = "DesktopInteract: " & CStr(.value)
'                Case "ErrorControl"
'                    vTemp(11) = "ErrorControl: " & .value
'                Case "ProcessId"
'                    vTemp(12) = "ProcessId: " & .value
'                Case "PathName"
'                    vTemp(13) = "PathName: " & .value
'                End Select
'            End With
'        Next objProp
'
'        If Len(Trim$(vTemp(1))) > 0 Then
'            lIndex = lIndex + 1
'            vData(lIndex) = vTemp
'        End If
'        DoEvents
'    Next objService
'    Get_SrvData = vData
'
'On Error GoTo 0
'
'End Function
'
'Private Function Get_PrcData(ByVal objWbm As SWbemServices) As Variant
''//interrogate process data
'
'Dim objPrcSet   As SWbemObjectSet
'Dim objProcess  As SWbemObject
'Dim Property    As SWbemProperty
'Dim vData       As Variant
'Dim vTemp       As Variant
'Dim lIndex      As Long
'Dim sQuery      As String
'Dim objEnum     As SWbemObjectSet
'Dim objInstance As SWbemObject
'Dim cTotal      As Currency
'Dim cProc       As Currency
'
'On Error Resume Next
'
'    Set objPrcSet = objWbm.InstancesOf("Win32_Process")
'    ReDim vTemp(1 To 8) As Variant
'    ReDim vData(1 To objPrcSet.count) As Variant
'    sQuery = "Select * From Win32_PerfRawData_PerfProc_Process Where Name = '_Total'"
'    Set objEnum = objWbm.ExecQuery(sQuery, "WQL")
'
'    For Each objInstance In objEnum
'        If objInstance.name = "_Total" Then
'            cTotal = CCur(objInstance.PercentProcessorTime)
'            Exit For
'        End If
'    Next objInstance
'
'    For Each objProcess In objPrcSet
'        For Each Property In objProcess.Properties_
'            With Property
'                Select Case .name
'                Case "Name"
'                    vTemp(1) = "Name: " & .value
'                Case "ProcessId"
'                    vTemp(2) = "ProcessId: " & .value
'                Case "WorkingSetSize"
'                    vTemp(3) = "WorkingSetSize: " & Trans_Calc(.value) & " K"
'                Case "PeakWorkingSetSize"
'                    vTemp(4) = "PeakWorkingSetSize: " & Trans_Calc(.value) & " K"
'                Case "HandleCount"
'                    vTemp(5) = "HandleCount: " & .value
'                Case "ThreadCount"
'                    vTemp(6) = "ThreadCount: " & .value
'                Case "Priority"
'                    vTemp(7) = "Priority: " & .value
'                End Select
'            End With
'        Next Property
'
'        If Len(Trim$(vTemp(1))) > 0 Then
'            lIndex = lIndex + 1
'            vData(lIndex) = vTemp
'        End If
'        DoEvents
'    Next objProcess
'    Get_PrcData = vData
'
'On Error GoTo 0
'
'End Function

Private Function Trans_Calc(ByVal vNum As Variant) As String
'//format size (mb)

Dim lFactor As Double
Dim lMegs   As Double

On Error Resume Next

    lFactor = 1024#
    lMegs = CDbl(vNum) / lFactor
    Trans_Calc = Format$(CStr(lMegs), "Standard")

On Error GoTo 0

End Function

Private Function Trans_Soft(ByVal sDate As String) As String
'//format time

On Error Resume Next

    Trans_Soft = CStr(DateSerial(Left$(sDate, 4), Mid$(sDate, 5, 2), Mid$(sDate, 7, 2)))

On Error GoTo 0

End Function

Private Function Trans_Date(ByVal sDate As String) As String
'//format date

On Error Resume Next

    Trans_Date = CStr(DateSerial(Left$(sDate, 4), Mid$(sDate, 5, 2), Mid$(sDate, 7, 2))) & " -- " & CStr(TimeSerial(Mid$(sDate, 9, 2), Mid$(sDate, 11, 2), Mid$(sDate, 13, 2)))

On Error GoTo 0

End Function

Private Function Trans_Percent(ByVal vNum As Variant) As String
'//format percent

On Error Resume Next

    Trans_Percent = Format$(vNum, "Percent")

On Error GoTo 0

End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           WEB INTERFACE
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Sub Report_Web(ByRef lFrm As Form)
'//display log as a web page
'//could be expanded to special style envelops
'//with div or paragraph styles, skies the limit..

Dim sTemp   As String
Dim iPos    As Integer
Dim vItem   As Variant

Const NLINE As String = vbCr + vbLf

On Error Resume Next

    If Len(dir(WebPath)) > 0 Then
        Kill WebPath
    End If

    '//write html header
    Open WebPath For Output As #1
    Print #1, "<HTML>" & NLINE
    Print #1, "<HEAD>" & NLINE
    Print #1, "<TITLE>" & AppTitle & Chr$(32) & CStr(Now) & "</TITLE>" & NLINE
    
    '//test for style sheet
    If CssStyle Then
        If Len(CssBackColor) = 0 Then CssBackColor = "ffffff"
        If Len(CssForeColor) = 0 Then CssForeColor = "555555"
        If CssFontSize = 0 Then CssFontSize = 24
        If CssHeadSize = 0 Then CssHeadSize = 32
        If Len(CssHeadColor) = 0 Then CssHeadColor = "333333"
        If Len(CssFont) = 0 Then CssFont = "9 verdana"
        
        Print #1, "<style TYPE=" & Chr$(34) & "text/css" & Chr$(34) & ">" & NLINE
        Print #1, "<!-- " & NLINE
        Print #1, "body {" & NLINE
        Print #1, "background: " & "#" & CssBackColor & ";" & NLINE
        Print #1, "color: " & "#" & CssForeColor & ";" & NLINE
        Print #1, "font: " & CssFont & ";" & NLINE
        Print #1, "font-size: " & CssFontSize & "pt;" & "}" & NLINE
        Print #1, "}" & NLINE
        Print #1, "H5 {" & NLINE
        Print #1, "font-size: " & CssHeadSize & NLINE
        Print #1, "color: #" & CssHeadColor & " }" & NLINE
        Print #1, "--> " & NLINE
        Print #1, "</style> " & NLINE
    End If

    Print #1, "</HEAD>" & NLINE
    Print #1, "<BODY>" & NLINE

    '//test for image path
    If Len(WebImage) > 0 Then
        '//format image path for html
        sTemp = "file:///" & Replace(WebImage, Chr$(92), Chr$(47), 1)
        '//add path
        Print #1, "<p><img border=" & Chr$(34) & "0" & Chr$(34) & " src=" & Chr$(34) & sTemp & Chr$(34) & "></p>"
    End If
    
    '//header style
    If CssHeader Then
        Print #1, "<H5>" & AppTitle & "</H5><BR>"
    Else
        Print #1, AppTitle & "<BR>"
    End If
    
    '//rest of intro message
    Print #1, AppCpr & "<BR>"
    Print #1, AppMsg & "<BR>"
    Print #1, "Log Created: " & dTime & "<BR>"
    
    '//add data
    For Each vItem In cLogData
        Print #1, vItem & "<BR>"
    Next vItem
    
    '//add details
    For Each vItem In cLogDetail
        Print #1, vItem & "<BR>"
    Next vItem
    
    '//add footer
    Log_Footer
    For Each vItem In cLogFooter
        Print #1, vItem & "<BR>"
    Next vItem
    
    '//close tags
    Print #1, NLINE
    Print #1, "</BODY>" & NLINE
    Print #1, "</HTML>"
    Close #1

    '//display
    ShellExecute lFrm.hwnd, "open", WebPath, vbNullString, vbNullString, SW_NORMAL
    
    '//reset footer
    Set cLogFooter = Nothing
    Set cLogFooter = New Collection
    
On Error GoTo 0

End Sub

Public Sub Report_Mail(lFrm As Form, _
                       lTxt As TextBox)
'//create an email report

'//first collect the data for the message
'//then format and create an eml file
'//and launch with the default mail client

Dim sBody   As String
Dim vItem   As Variant
Dim FF      As Integer
Dim sPath   As String
Dim sFrom   As String

On Error Resume Next

    FF = FreeFile
    
    '//build message body
    With lTxt
        sBody = .Text
        sBody = sBody & vbNewLine
'        For Each vItem In Get_Stats
'            sBody = sBody & vItem & vbNewLine
'        Next vItem
    End With
    
    '//test for setting
    If Len(MailTo) = 0 Then Exit Sub
    
    '//get users email address
    sFrom = Read_String(HKEY_CURRENT_USER, _
    "Software\Microsoft\Internet Account Manager\Accounts\00000001", _
    "SMTP Email Address")
    
    '//construct email
    Open MailPath For Output As #FF
    Print #FF, "From: <"; sFrom; ">"
    Print #FF, "To: "; Chr$(34); MailTo; Chr$(34)
    
    '//optional CC
    If Len(MailCc) > 0 Then
        Print #FF, "CC: "; Chr$(34); MailCc; Chr$(34)
    End If
    
    '//write email
    Print #FF, "Subject: "; MailSubject
    Print #FF, "X-Unsent: 1"
    Print #FF, ""
    Print #FF, "Add a Note Here:"
    Print #FF, ""
    Print #FF, ""
    Print #FF, ""
    Print #FF, "<<<BEGIN ERROR REPORT - DO NOT WRITE BELOW THIS LINE>>>"
    Print #FF, sBody
    Print #FF, "<<END ERROR>>"
    Close #FF
    
    '//launch in default client
    ShellExecute lFrm.hwnd, "Open", MailPath, vbNullString, vbNullString, SW_NORMAL

On Error GoTo 0

End Sub

Public Sub Report_Error(ByRef cLog As Collection)
'//generate an error log
    
Dim vItem   As Variant
Dim FF      As Integer

On Error Resume Next

    FF = FreeFile
    '//write to log
    Open ErrPath For Output As #FF
        For Each vItem In cLog
            Print #FF, vItem
        Next vItem
    Close #FF

On Error GoTo 0

End Sub

Public Sub Report_Note()
'//launch text application log

    Log_Commit True
    
End Sub

Private Function Read_String(ByVal RootKey As REG_HKEY, _
                             ByVal SubKey As String, _
                             ByVal value As String) As String
'//read an SZ value

Dim hKey    As Long
Dim RetVal  As Long
Dim sBuffer As String
Dim slength As Long

On Error GoTo Handler

    '//open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = 0 Then
        Exit Function
    Else
        '//create a buffer
        sBuffer = Space$(255)
        slength = 255
        '//query key for string value
        RetVal = RegQueryValueEx(hKey, value, 0, 1, ByVal sBuffer, slength)
        '//read string into buffer
        If RetVal = 0 Then
            sBuffer = Left$(sBuffer, slength - 1)
            Read_String = sBuffer
        End If
    End If

Handler:
    '//close key and set result
    RetVal = RegCloseKey(hKey)
    On Error GoTo 0

End Function

'Private Function Get_Stats() As Collection
''//get system statistics for error log
'
'Dim vntRtn  As Variant
'Dim lLoop   As Long
'Dim lData   As Long
'Dim cTemp   As New Collection
'
'On Error Resume Next
'
'    Set cTemp = New Collection
'
'    '//software data
'    vntRtn = Get_SoftData(objWmiRoot)
'    If Not IsEmpty(vntRtn) Then
'        With cTemp
'            .Add ""
'            .Add "Registered Software"
'            .Add "»»»»»»»»»»»»»»»»»»»"
'        End With
'        '//add data strings
'        For lLoop = 1 To UBound(vntRtn)
'            With cTemp
'                .Add ""
'                .Add vntRtn(lLoop)(1)
'            End With
'            For lData = 2 To UBound(vntRtn(1))
'                cTemp.Add vntRtn(lLoop)(lData)
'            Next lData
'        Next lLoop
'    End If
'
'    '//computer data
'    vntRtn = Get_CompData(objWmiRoot)
'    If Not IsEmpty(vntRtn) Then
'        With cTemp
'            .Add ""
'            .Add "Computer Configuration"
'            .Add "»»»»»»»»»»»»»»»»»»»»»»"
'            .Add ""
'        End With
'        For lLoop = 1 To UBound(vntRtn, 2)
'            cTemp.Add vntRtn(0, lLoop) & ": " & vntRtn(1, lLoop)
'        Next lLoop
'    End If
'
'    '//os data
'    vntRtn = Get_OsData(objWmiRoot)
'    If Not IsEmpty(vntRtn) Then
'        With cTemp
'            .Add ""
'            .Add "OS Configuration"
'            .Add "»»»»»»»»»»»»»»»»"
'            .Add ""
'        End With
'        For lLoop = 1 To UBound(vntRtn, 2)
'            cTemp.Add vntRtn(0, lLoop) & ": " & vntRtn(1, lLoop)
'        Next lLoop
'    End If
'
'    Set Get_Stats = cTemp
'    Set cTemp = Nothing
'
'On Error GoTo 0
'
'End Function
'
'
