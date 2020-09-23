Attribute VB_Name = "mMain"
Option Explicit

Public mScheduler       As clsScheduler
Public m_sRegPath       As String

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Sub Main()
'/* entry point

Dim sCommand As String

On Error GoTo Handler

    '/* init common controls
    Init_Controls

    '/* get app strings
    Get_Settings
    
    '/* instantiate objects
    Preload_Scheduler
    
    '/* test for command string
    If Not Len(Command) > 0 Then
        frmSchedule.Show
        If Not mScheduler.Get_RunState And mScheduler.Jobs_Pending Then
            mScheduler.Scheduler_Start
        End If
    Else
        sCommand = LCase$(CStr(Command))
        Command_Interpreter sCommand
    End If

Exit Sub

Handler:
    Launch_Help 0

End Sub

Private Sub Init_Controls()

On Error Resume Next

    InitCommonControls

On Error GoTo 0

End Sub

Public Sub Command_Interpreter(ByVal sCommand As String)

Dim sTemp  As String
Dim sCmd() As String
Dim lTemp  As Long

On Error GoTo Handler

    Select Case True
    '/* add task
    Case sCommand = "-s"
        If mScheduler.Get_RunState Then
            Load frmSysTray
            mScheduler.Scheduler_Start
        Else
            mScheduler.Delete_Value HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "nspsched"
            mScheduler.Write_DWord HKEY_CURRENT_USER, m_sRegPath, "stay", 0
            Unload_Scheduler
            Application_Terminate
        End If
        
    Case InStr(1, sCommand, "-a") > 0
        sTemp = Mid$(sCommand, 3)
        sCmd = Split(sTemp, Chr$(44))
        '/* test user input
        If mScheduler.Test_Input(sCmd(2), 0) Then
            Exit Sub
        ElseIf mScheduler.Test_Input(sCmd(3), 1) Then
            Exit Sub
        End If
        '/* send to scheduler
        mScheduler.Scheduler_Add Trim$(sCmd(0)), Trim$(sCmd(1)), _
        CBool(Trim$(sCmd(2))), CLng(Trim$(sCmd(3)))
        
    '/* pause
    Case sCommand = "-p"
        mScheduler.Scheduler_Pause
        
    '/* restart
    Case sCommand = "-r"
        mScheduler.Scheduler_Start
        
    '/* remove task
    Case InStr(1, sCommand, "-t") > 0
        lTemp = CLng(Mid$(sCommand, 3))
        mScheduler.Scheduler_Remove lTemp
        
    '/* clear tasks
    Case sCommand = "-c"
        mScheduler.Scheduler_Clear
        
    '/* end
    Case sCommand = "-e"
        mScheduler.Scheduler_End
        
    '/* help
    Case sCommand = "-?"
        GoTo Handler
        
    '/* unknown
    Case Else
        Launch_Help 0
    End Select

    Exit Sub
    
Handler:
    Launch_Help 0

End Sub

Public Sub Launch_Help(ByVal iType As Integer)

Dim sHelp As String

    Select Case iType
    '/* command line
    Case 0
        sHelp = "Schedule a task: -s (duration {str}, task {str}, recurring {bool}, type index {long}" _
        & vbCrLf & "Duration -standard time formats determined by type index" & vbCrLf & "Recurring -true|false" _
        & vbCrLf & "Task -path to application" & vbCrLf & "Type Index Formats" & vbCrLf & "1) hourly -minutes ex. 17" _
        & vbCrLf & "2) daily -hours:minutes ex. 18:52 " & vbCrLf & "3) weekly -day hours:minutes ex. Saturday 18:52:00" _
        & vbCrLf & "4) user specify -MM/DD/YYYY HH:mm:ss ex. 02/17/2006 21:17:52" & vbCrLf & "Add a task: -a {duration, task, recurring, index}" _
        & vbCrLf & "Remove a task: -t {task id {long}}" & vbCrLf & "Clear Tasks: -c {clear the task list}" _
        & vbCrLf & "Close Scheduler: -e {clear tasks and close}" & vbCrLf & "Hourly Ex: nspsched.exe -a 59, notepad.exe, true, 1" _
        & vbCrLf & "nspsched.exe -a Monday 15:30:00, c:\app.exe, false, 3" & vbCrLf & "Starts Monday at 3:30 pm, runs once, launches myapp.exe" _
        & vbCrLf & "User Ex: nspsched.exe -a 02/17/2006 21:17:00, myapp.exe, false, 4" & vbCrLf & "Starts Febuary 17, 2006 at 9:17 pm, runs once, launches myapp.exe"

    '/* gui
    Case 1
        sHelp = "Task Scheduler Help" & vbCrLf & _
        vbCrLf & "Interval: The time you want the event to begin. Example: Thursday 17:30" & _
        vbCrLf & "would start the application on Thursday at 5:30 pm." & _
        vbCrLf & "Recurring: When checked, the event will reoccur at the same time every day." & _
        vbCrLf & "Start Up: When selected, the scheduler will start automatically each" & _
        vbCrLf & "time you restart your computer." & _
        vbCrLf & "Application: Select the program you want to run with the scheduler." & _
        vbCrLf & "Add Job: Once you have chosen the time and application, this will" & _
        vbCrLf & "add it to the scheduler." & _
        vbCrLf & "Remove Job: This option removes the job from the scheduler." & _
        vbCrLf & "Clear Jobs: This option removes all jobs from the scheduler." & _
        vbCrLf & "Finish: This closes the scheduler window. To stop the scheduler" & _
        vbCrLf & "completely, just remove all tasks from the job list."
    End Select
    
    MsgBox sHelp, vbOKOnly, "NSP Scheduler Lite - Help"

End Sub

Public Sub Preload_Scheduler()
'/* initialize classes

    Set mScheduler = New clsScheduler
    Load frmSchedule

End Sub

Public Sub Unload_Scheduler()
'/* early exit

On Error Resume Next

    If Not mScheduler Is Nothing Then
        With mScheduler
            .Delete_Value HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "nspsched"
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "stay", 0
            .Scheduler_Clear
        End With
        Set mScheduler = Nothing
    Else
        Set mScheduler = New clsScheduler
        With mScheduler
            .Delete_Value HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "nspsched"
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "stay", 0
            .Scheduler_Clear
        End With
        Set mScheduler = Nothing
    End If

On Error GoTo 0

End Sub

Private Sub Get_Settings()
'/* get application settings

    '/* reg path
    m_sRegPath = "Software\" & App.ProductName

End Sub

Public Sub Application_Terminate()
'/* close it all

Dim frm As Form

On Error Resume Next

    For Each frm In Forms
        Unload frm
    Next frm
    
On Error GoTo 0

End Sub


