VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  clsScheduler! 1.0 Event Scheduler Tool                                             *
'*                                                                                     *
'*  Created:     Febuary 11, 2006                                                      *
'*  Updated:     March 27 2006                                                         *
'*  Purpose:     Task Event Scheduling Tool                                            *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.1                                                                   *
'*  Compile:     Native                                                                *
'*  Referenced:  Throughout Project                                                    *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*  Credits:     Thanks to Fred Cpp for the cool isButton                              *
'*                                                                                     *
'***************************************************************************************

'*~ List of exposed functions ~*

'/*  Scheduler_Start    - create a running instance of the scheduler
'/*  Scheduler_Add      - add a task to the event list
'/*  Scheduler_Remove   - remove a task from the event list
'/*  Scheduler_Reset    - rewrite events list
'/*  Scheduler_Pause    - pause scheduler timer
'/*  Scheduler_Clear    - remove all scheduled events
'/*  Scheduler_End      - unload scheduler
'/*  Get_Idle           - get system idle time
'/*  Get_Instance       - test for previous instance

'/~ I recently discovered that the scheduler tool in XP (SP2),
'/~ is effectively broken. It now requires a user logon to run
'/~ a task. What this means, is that if you use the scheduler service
'/~ to update your applications, (via netapi or wmi), they will not run.
'/~ No error is logged, and the task appears to run, but without an
'/~ authenticated user account, the task never actually commits.
'/~ I can only imagine that this is due to some heavy handed M$ security
'/~ fix, there is probably a registry hack out there somewhere to get
'/~ around this, (though I could not find it), or, you could create a user
'/~ account for your software, and use this to log in and run the task,
'/~ (security? what security??), and simply wait for them to break it again
'/~ in some other way.. /or/ you could create your own scheduler..
'/~ This tool is a viable replacement for the windows scheduler. It can
'/~ schedule multiple simultaneous events, (theoretically 1000+), has
'/~ almost no cpu impact, and is completely accessable from the command line.
'/~ You can use this tool in any way you like, (excepting malicious), but..
'/~ if you publish, you can not use this interface, and you must give
'/~ proper credit to the author, (that's me!), ex. 'scheduler based on the
'/~ Task Master scheduler by John Underhill, www.nspowertools.com',
'/~ placed in an appropriate place in your application, ie. About/Help/web page.
'/~ Anyways, hope this helps people out, if you have a comment or a job
'/~ drop me a line at steppenwolfe_2000@yahoo.com, Enjoy..

'/~ Updated Febuary 11, 2006
'/~ Consolidated modules into a class. Tightened up the code in a few spots
'/~ changed time readout to 12 hour clock for clarity.

'/~ Updated March 27, 2006
'/~ Added a skin, a la NeoClass, fixed a couple bugs, added a couple features, enjoy..


Private Type POINTAPI
    xPos As Long
    yPos As Long
End Type

'/~ registry handlers /~

'/* security structure
Private Type SECURITY_ATTRIBUTES
    nLength                                    As Long
    lpSecurityDescriptor                       As Long
    bInheritHandle                             As Boolean
End Type

'/* key constants
Public Enum HKEY_Type
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'/* schedule type
Public Enum eTime
    Hourly = 1
    daily = 2
    weekly = 3
    specific = 4
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)
                                                                     
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
                                                                                
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long
                                                                                      
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                                                                    ByVal lpValueName As String) As Long

Private Declare Function RegCreatekey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                                                                    ByVal lpSubKey As String, _
                                                                                    ByVal Reserved As Long, _
                                                                                    ByVal lpClass As String, _
                                                                                    ByVal dwOptions As Long, _
                                                                                    ByVal samDesired As Long, _
                                                                                    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                                    phkResult As Long, _
                                                                                    lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  lpData As Any, _
                                                                                  ByVal cbData As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private m_bIsActive     As Boolean
Private m_bIsLoaded     As Boolean
Private m_lInterval     As Long
Private m_lLastPos      As Long
Private c_Duration      As Collection
Private c_Tasks         As Collection
Private c_Recurring     As Collection
Private c_Interval      As Collection

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                              CORE PROCESSORS
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


'/~ Properties /~

'/* timer state
Public Property Get p_bIsActive() As Boolean
    p_bIsActive = m_bIsActive
End Property

Public Property Let p_bIsActive(ByVal PropVal As Boolean)
    m_bIsActive = PropVal
End Property

'/* running instance
Public Property Get p_bIsLoaded() As Boolean
    p_bIsLoaded = m_bIsLoaded
End Property

Public Property Let p_bIsLoaded(ByVal PropVal As Boolean)
    m_bIsLoaded = PropVal
End Property

Public Property Get p_Count() As Long
    p_Count = c_Duration.Count
End Property

Public Sub Scheduler_Start()
'/* start scheduler

    '/* test prev instance
    If Not m_bIsLoaded Then
        '/* test is loaded
        If Not m_bIsActive Then
            Timer_Enabled True
        End If
        Set_RunState True
    End If

End Sub

Public Sub Scheduler_Add(ByVal sDuration As String, _
                         ByVal sTask As String, _
                         ByVal bRecur As Boolean, _
                         ByVal enTime As eTime)

'/* add a task - format: duration, task, recurring, type

Dim vItem As Variant
Dim cTemp As Collection
Dim lItem As Long
Dim sTemp As String

On Error Resume Next

    '/* test time paramaters
    Select Case enTime
    '/* hourly
    Case 1
        If Len(sDuration) > 2 Then
            MsgBox "The time format is invalid." & vbCrLf & _
            "Use minute format when not using the hourly switch. Type -? for help." & vbCrLf & _
            "example usage: 53,notepad.exe,False,1", vbInformation, "Bad Time Format!"
            Exit Sub
        End If
        
    '/* daily
    Case 2
        If Not InStr(1, sDuration, Chr$(58)) > 0 Then
            MsgBox "The time format is invalid." & vbCrLf & _
            "Use daily time format when using the daily switch. Type -? for help." & vbCrLf & _
            "example usage: 14:26,notepad.exe,False,2", vbInformation, "Bad Time Format!"
            Exit Sub
        End If
        
    '/* weekly
    Case 3
        If (InStr(1, sDuration, Chr$(32)) < 1) Or (InStr(1, sDuration, Chr$(58)) < 1) Then
            MsgBox "The time format is invalid." & vbCrLf & _
            "Use standard date/time format when using the specific switch. Type -? for help." & vbCrLf & _
            "example usage: Monday 14:30,notepad.exe,False,4", vbInformation, "Bad Time Format!"
            Exit Sub
        End If
        
    '/* specific
    Case 4
        If (InStr(1, sDuration, Chr$(58)) < 1) Or (InStr(1, sDuration, Chr$(47)) < 2) Then
            MsgBox "The time format is invalid." & vbCrLf & _
            "Use standard date/time format when using the specific switch. Type -? for help." & vbCrLf & _
            "example usage: 02/06/2006 14:28:42,notepad.exe,False,3", vbInformation, "Bad Time Format!"
            Exit Sub
        End If
    End Select
    
    '/* test boolean and numeric vals
    If Test_Input(bRecur, 0) Then
        MsgBox "The Recurring switch is invalid." & vbCrLf & _
        "Use a boolean statement, ex 'True' or 'False'. Type -? for help." & vbCrLf & _
        "example usage: 02/06/2006 14:28:42,notepad.exe,False,3", vbInformation, "Bad Switch Format!"
        Exit Sub
    ElseIf Test_Input(enTime, 1) Then
        MsgBox "The Type switch is invalid." & vbCrLf & _
        "Use numeric format 1 to 4 for this switch. Type -? for help." & vbCrLf & _
        "example usage: 02/06/2006 14:28:42,notepad.exe,False,3", vbInformation, "Bad Switch Format!"
        Exit Sub
    End If
    
    '/* test command string
    If Not InStr(1, sTask, ".") > 0 Or Len(sTask) < 4 Then
        MsgBox "The task command is invalid." & vbCrLf & _
        "Use the path format: 'c:\windows\calc.exe' or for registered files 'notepad.exe'. Type -? for help." & vbCrLf & _
        "example usage: nspsched -a 00:03:00,notepad.exe,False,2", vbInformation, "Bad Type Format!"
        Exit Sub
    End If
    
    '/* build command string
    sTemp = sDuration + Chr$(44) + sTask + Chr$(44) + CStr(bRecur) + Chr$(44) + CStr(enTime)
    Set cTemp = New Collection
    
    '/* first entry
    If Read_MultiCN(HKEY_CURRENT_USER, m_sRegPath, "ctasks") Is Nothing Then
        cTemp.Add sTemp, "1"
        Write_MultiCN HKEY_CURRENT_USER, m_sRegPath, "ctasks", cTemp
        GoTo Handler
    End If

    '/* update multi_sz entry
    For Each vItem In Read_MultiCN(HKEY_CURRENT_USER, m_sRegPath, "ctasks")
        lItem = lItem + 1
        cTemp.Add vItem, CStr(lItem)
    Next vItem
    
    cTemp.Add sTemp, CStr(lItem + 1)
    Delete_Value HKEY_CURRENT_USER, m_sRegPath, "ctasks"
    Write_MultiCN HKEY_CURRENT_USER, m_sRegPath, "ctasks", cTemp

Handler:
    On Error GoTo 0
    '/* update tasks
    Scheduler_Reset

End Sub

Public Function Jobs_Pending() As Boolean
'/* pending job check

On Error Resume Next

    If Read_MultiCN(HKEY_CURRENT_USER, m_sRegPath, "ctasks").Count > 0 Then
        Jobs_Pending = True
    End If

On Error GoTo 0

End Function

Public Sub Scheduler_Remove(ByVal lmItem As Long)
'/* remove item

Dim vItem As Variant
Dim lItem As Long
Dim cTemp As Collection

On Error Resume Next

    Set cTemp = New Collection
    '/* check for values
    If Read_MultiCN(HKEY_CURRENT_USER, m_sRegPath, "ctasks").Count = 0 Then
        Scheduler_Pause
        Unload_Instance
    End If
    
    '/* iterate through entries
    For Each vItem In Read_MultiCN(HKEY_CURRENT_USER, m_sRegPath, "ctasks")
        lItem = lItem + 1
        '/* ignore entry
        If Not lItem = lmItem Then
            cTemp.Add vItem, CStr(lItem)
        End If
    Next vItem
    
    Delete_Value HKEY_CURRENT_USER, m_sRegPath, "ctasks"
    
    If cTemp.Count > 0 Then
        Write_MultiCN HKEY_CURRENT_USER, m_sRegPath, "ctasks", cTemp
    End If
    
On Error GoTo 0
    
    '/* update tasks
    Scheduler_Reset

End Sub

Private Sub Scheduler_Reset()
'/* duration/task/repeat

Dim vItem  As Variant
Dim cTemp  As Collection
Dim sCmd() As String

On Error Resume Next

    '/* hold for changes
    If m_bIsActive Then
        Scheduler_Pause
    End If
    
    '/* if no tasks - pause
    Set cTemp = New Collection
    If Read_MultiCN(HKEY_CURRENT_USER, m_sRegPath, "ctasks") Is Nothing Then
        Scheduler_Pause
        Exit Sub
    End If
    
    '/* reset storage
    Set c_Duration = New Collection
    Set c_Tasks = New Collection
    Set c_Recurring = New Collection
    Set c_Interval = New Collection
    
    '/* get each entry and assign to collection
    For Each vItem In Read_MultiCN(HKEY_CURRENT_USER, m_sRegPath, "ctasks")
        If Len(vItem) = 0 Then GoTo skip
        sCmd = Split(vItem, Chr$(44))
        c_Duration.Add Trim$(sCmd(0))
        c_Tasks.Add Trim$(sCmd(1))
        c_Recurring.Add Trim$(sCmd(2))
        c_Interval.Add Trim$(sCmd(3))
        ReDim sCmd(0 To 3) As String
skip:
    Next vItem
        '/* reset todo flag
    Write_DWord HKEY_CURRENT_USER, m_sRegPath, "status", 0
    
    '/* start timer
    If c_Duration.Count > 0 Then
        Scheduler_Start
    '/* no events
    Else
        Scheduler_Pause
    End If

On Error GoTo 0

End Sub

Public Sub Scheduler_Pause()
'/* pause scheduler

    If Not m_bIsLoaded Then
        Timer_Enabled False
        '/* flag a running instance
        Set_RunState False
    End If

End Sub

Public Sub Scheduler_Clear()
'/* clear tasks

    Scheduler_Pause
    Delete_Value HKEY_CURRENT_USER, m_sRegPath, "status"
    Delete_Value HKEY_CURRENT_USER, m_sRegPath, "ctasks"

End Sub

Public Sub Schedule_Check()
'/* test task status

Dim lItem       As Long
Dim bValid      As Boolean
Dim dCurrent    As Date

On Error Resume Next

    dCurrent = Now
    '/* test for run flag
    If Not Get_RunState Then
        Scheduler_Pause
        Exit Sub
    End If
    
    '/* test for changes
    If Get_Status Then
        Scheduler_Reset
    End If
    
    '/* sanity check
    If c_Duration.Count = 0 Then
        Scheduler_Pause
        Exit Sub
    End If
    
    For lItem = c_Duration.Count To 1 Step -1
        bValid = False
        Select Case c_Interval.Item(lItem)
        '/*  hourly
        Case 1
            If Format$(c_Duration.Item(lItem), "00") = Format$(dCurrent, "nn") Then
                Launch_App CStr(c_Tasks.Item(lItem))
                bValid = True
            End If
            
        '/* daily
        Case 2
            If Format$(c_Duration.Item(lItem), "HH:nn") = Format$(dCurrent, "HH:nn") Then
                Launch_App CStr(c_Tasks.Item(lItem))
                bValid = True
            End If
            
        '/* weekday
        Case 3
            If Format$(c_Duration.Item(lItem), "dddd HH:nn") = Format$(dCurrent, "dddd HH:nn") Then
                Launch_App CStr(c_Tasks.Item(lItem))
                bValid = True
            End If
            
        '/*  specific
        Case 4
            If Format$(c_Duration.Item(lItem), "mm/dd/yyyy HH:nn") = Format$(dCurrent, "mm/dd/yyyy HH:nn") Then
                Launch_App CStr(c_Tasks.Item(lItem))
                bValid = True
            End If
        End Select
        
        '/* remove item if not recurring
        If bValid Then
            If Not CBool(c_Recurring.Item(lItem)) Then
                Scheduler_Remove lItem
            End If
        End If
    Next lItem

On Error GoTo 0

End Sub

Public Sub Scheduler_End()
'/* end scheduler

On Error Resume Next

    Scheduler_Pause
    Unload_Instance

End Sub

Public Function Get_Idle() As Long
'/* track system idle time

Dim tPoint  As POINTAPI
Dim lPos    As Long

    GetCursorPos tPoint
    '/* current position
    lPos = tPoint.xPos + tPoint.yPos
    
    '/* cursor position compare
    If lPos = m_lLastPos Then
        m_lInterval = m_lInterval + 1
    Else
        m_lInterval = 0
    End If
    
    '/* current idle time
    Get_Idle = m_lInterval
    
    '/* last position
    m_lLastPos = lPos
    
    '/* safe reset
    If m_lInterval > 10080 Then
        m_lInterval = 0
    End If
    
End Function

Public Function Get_Instance() As Boolean
'/* test for previous instance

    p_bIsLoaded = App.PrevInstance
    Get_Instance = p_bIsLoaded

End Function

Private Sub Timer_Enabled(ByVal bEnabled As Boolean)
'/* boolean timer switch

    If bEnabled Then
        '/ * 1 minute interval
        Start_Timer 60000
        p_bIsActive = True
    Else
        Kill_Timer
        p_bIsActive = False
    End If

End Sub

Public Function Get_RunState() As Boolean
'/* test active status

    Get_RunState = Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "jobstatus") > 0

End Function

Private Sub Set_RunState(ByVal bState As Boolean)
'/* set running flag

    '/* running
    If bState Then
        Write_DWord HKEY_CURRENT_USER, m_sRegPath, "jobstatus", 1
    '/ stop
    Else
        Write_DWord HKEY_CURRENT_USER, m_sRegPath, "jobstatus", 0
    End If

End Sub

Private Function Get_Status() As Boolean
'/* test for change in status

    Get_Status = Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "status") > 0

End Function

Private Sub Launch_App(ByVal sCmd As String)
'/* launch application

Dim lStatus As Long

On Error Resume Next

    '/* launch
    lStatus = ShellExecute(0&, "open", sCmd, "", "", 1)

    '/* track error
    If lStatus = 0 Then
        Log_Error "The application: " & _
        Mid$(sCmd, InStrRev(sCmd, Chr$(92), 1) - 1) & "could not be launched."
    End If
    
On Error GoTo 0

End Sub

Private Sub Log_Error(ByVal sErrDesc As String)
'/* log errors

On Error Resume Next

    '/* write to log
    Open App.Path & "\application.log" For Append As #1
    Print #1, CStr(Now) & vbTab & sErrDesc
    Close #1
    
On Error GoTo 0

End Sub

Public Function Test_Input(ByVal sInput As String, _
                           ByVal iType As Integer) As Boolean

'/* invalid input will throw err. 13

Dim bTest As Boolean
Dim lTest As Long

On Error Resume Next

    sInput = Trim$(sInput)
    Err.Clear
    
    Select Case iType
    '/* test boolean
    Case 0
        bTest = CBool(sInput)
        If Err.Number <> 0 Then
            Test_Input = True
        End If
        
    '/* test long
    Case 1
        lTest = CLng(sInput)
        If Err.Number <> 0 Or lTest < 1 Or lTest > 4 Then
            Test_Input = True
        End If
    End Select

On Error GoTo 0

End Function

Private Sub Preload_Instance()
'/* initialize storage

    Set c_Duration = New Collection
    Set c_Tasks = New Collection
    Set c_Recurring = New Collection
    Set c_Interval = New Collection
    Get_Instance

End Sub

Private Sub Unload_Instance()
'/* unload storage

    Set c_Duration = Nothing
    Set c_Tasks = Nothing
    Set c_Recurring = Nothing
    Set c_Interval = Nothing

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                              REGISTRY PROCESSING
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Public Function Read_DWord(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String, _
                           ByVal Value As Variant, _
                           Optional ByVal lErrdesc As Long) As Long

'/* read a dword value
Dim hKey    As Long
Dim RetVal  As Long
Dim lBuffer As Long

On Error GoTo Handler

    '/* open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, &HF003F, hKey)
    If Not RetVal = &H0 Then
        lErrdesc = RetVal
    End If
    '/* query value
    RetVal = RegQueryValueEx(hKey, Value, 0, &H4, lBuffer, 4)
    '/* set value
    If RetVal = &H0 Then
        Read_DWord = lBuffer
    Else
        Read_DWord = 0
    End If

Handler:
    '/* close key
    lErrdesc = RetVal
    RetVal = RegCloseKey(hKey)
    On Error GoTo 0

End Function

Public Function Write_DWord(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal Value As String, _
                            ByVal lData As Long) As Long

'/* write a DWORD value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '/* required security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '/* open key and test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, &H20006, secattr, hKey, deposit)
    If Not RetVal = &H0 Then
        Write_DWord = RetVal
        Exit Function
    End If
    '/* write value, set error, and close key
    RetVal = RegSetValueEx(hKey, Value, 0&, &H4, lData, 4)

Handler:
    Write_DWord = RetVal
    RetVal = RegCloseKey(hKey)
    On Error GoTo 0

End Function

Public Function Read_Multi(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String, _
                           ByVal Value As String, _
                           Optional ByVal lErrdesc As Long) As String

'/* read a multi_sz value

Dim hKey        As Long
Dim RetVal      As Long
Dim sBuffer     As String
Dim Length      As Long
Dim resBinary() As Byte
Dim resString   As String

On Error GoTo Handler

    '/* open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, &HF003F, hKey)
    If Not RetVal = &H0 Then
        lErrdesc = RetVal
    Else
        '/* get buffer size
        Length = 1024
        ReDim resBinary(0 To Length - 1) As Byte
        RetVal = RegQueryValueEx(hKey, Value, 0, &H7, resBinary(0), Length)
        '/* if more data, resize our byte array
        If RetVal = &HEA Then
            ReDim resBinary(0 To Length - 1) As Byte
            RetVal = RegQueryValueEx(hKey, Value, 0, &H7, resBinary(0), Length)
        End If
        '/* read into buffer
        If RetVal = &H0 Then
            resString = Space$(Length - 2)
            CopyMemory ByVal resString, resBinary(0), Length - 2
            sBuffer = resString
            If Len(Trim_Null(sBuffer)) > 0 Then
                Read_Multi = resString
            End If
        Else
            lErrdesc = RetVal
        End If
    End If

Handler:
    '/* set error and close key
    lErrdesc = 0
    RetVal = RegCloseKey(hKey)
    On Error GoTo 0

End Function

Public Function Write_Multi(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal Value As String, _
                            ByVal sData As String) As Long

'/* write a multi-sz value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '/* security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '/* test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, &H20006, secattr, hKey, deposit)
    If Not RetVal = &H0 Then
        Write_Multi = RetVal
        Exit Function
    End If
    '/* write multi, set error, and close key
    RetVal = RegSetValueEx(hKey, Value, 0, &H7, ByVal sData, Len(sData))

Handler:
    Write_Multi = RetVal
    RetVal = RegCloseKey(hKey)
    On Error GoTo 0

End Function

Public Function Read_MultiCN(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String) As Collection

'//process multi_sz value
'//and return collection
Dim aTemp()     As String
Dim cTemp       As New Collection
Dim sTemp       As String
Dim l           As Long
Dim lReturn     As Long

On Error GoTo Handler

    Set cTemp = New Collection
    '//get multi string
    sTemp = Read_Multi(RootKey, SubKey, KeyVal, lReturn)
    If Len(sTemp) = 0 Then GoTo Handler
    
    '//get values and add to collection
    aTemp = Split(sTemp, vbNullChar)
    For l = 0 To UBound(aTemp)
        cTemp.Add aTemp(l)
    Next l
    
Handler:
    '//test for error
    If lReturn = 0 Then
        '//set collection and release library
        Set Read_MultiCN = cTemp
        Set cTemp = Nothing
    End If

End Function

Public Sub Write_MultiCN(ByVal RootKey As HKEY_Type, _
                         ByVal SubKey As String, _
                         ByVal KeyVal As String, _
                         ByVal DataSet As Collection)

'//process a collection and
'//write to registry as a
'//multi_sz entry
Dim sTemp       As String
Dim sValue      As String
Dim cTemp       As New Collection
Dim l           As Long
Dim lReturn     As Long
Dim Item        As Variant

On Error Resume Next

    For Each Item In DataSet
        sTemp = sTemp & Item & vbNullChar
    Next
    '//add terminating null char
    sTemp = sTemp & vbNullChar
    '//write to registry
    lReturn = Write_Multi(RootKey, SubKey, KeyVal, sTemp)
    
On Error GoTo 0

End Sub

Public Function Read_String(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal Value As String, _
                            Optional ByVal lErrdesc As Long) As String

'/* read an SZ value
Dim hKey    As Long
Dim RetVal  As Long
Dim sBuffer As String
Dim slength As Long

On Error GoTo Handler

    '/* open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, &HF003F, hKey)
    If Not RetVal = &H0 Then
        lErrdesc = RetVal
    Else
        '/* create a buffer
        sBuffer = Space$(255)
        slength = 255
        '/* query key for string value
        RetVal = RegQueryValueEx(hKey, Value, 0, &H1, ByVal sBuffer, slength)
        '/* read string into buffer
        If RetVal = &H0 Then
            sBuffer = Left$(sBuffer, slength - 1)
            Read_String = sBuffer
        End If
    End If

Handler:
    '/* close key and set result
    lErrdesc = RetVal
    RetVal = RegCloseKey(hKey)
    On Error GoTo 0

End Function

Public Function Write_String(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal Value As String, _
                            ByVal Data As String) As Long

'/* write an SZ value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '/* security attributes
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    
    '/* open key
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, &H20006, secattr, hKey, deposit)
    If Not RetVal = &H0 Then
        Write_String = RetVal
        Exit Function
    End If
    '/* write string, set error, and close
    RetVal = RegSetValueEx(hKey, Value, 0, &H1, ByVal Data, Len(Data))

Handler:
    Write_String = RetVal
    RetVal = RegCloseKey(hKey)
    On Error GoTo 0

End Function

Public Function Create_Key(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String) As Long

'/* create a new key
Dim hKey   As Long
Dim RetVal As Long

On Error GoTo Handler

    '/* create the key
    RetVal = RegCreatekey(RootKey, SubKey, hKey)

Handler:
    '/* set error and close
    Create_Key = RetVal
    RetVal = RegCloseKey(hKey)
    On Error GoTo 0

End Function

Public Function Delete_Key(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String) As Long

'/* delete a key
Dim RetVal  As Long

On Error GoTo Handler

    '/* delete key
    RetVal = RegDeleteKey(RootKey, SubKey)
    '/* set error
    Delete_Key = RetVal

Handler:
    On Error GoTo 0

End Function

Public Function Delete_Value(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal Value As Variant) As Long

'/* delete a value
Dim RetVal As Long
Dim handle As Long

On Error GoTo Handler

    '/* open key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, &HF003F, handle)
    '/* exit on error
    If Not RetVal = &H0 Then
        Delete_Value = RetVal
    Else
        '/* delete value
        RetVal = RegDeleteValue(handle, Value)
    End If

Handler:
    '/* set error and close key
    Delete_Value = RetVal
    RetVal = RegCloseKey(handle)
    On Error GoTo 0

End Function

Private Function Trim_Null(Item As String) As String

'/* trim nulls

Dim pos As Integer

On Error GoTo Handler

    '/* trim nulls for sz_multi
    pos = InStr(Item, vbNullChar)
    If pos Then
        Item = Left$(Item, pos - 1)
    End If
    Trim_Null = Item

Handler:
    On Error GoTo 0

End Function

Private Sub Class_Initialize()
'/* load storage

    Preload_Instance

End Sub

Private Sub Class_Terminate()
'/* unload storage

    Unload_Instance

End Sub
