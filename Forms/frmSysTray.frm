VERSION 5.00
Begin VB.Form frmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   1575
   ClientLeft      =   6330
   ClientTop       =   5595
   ClientWidth     =   2055
   ControlBox      =   0   'False
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuSched 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSubSched 
         Caption         =   "Open Scheduler"
         Index           =   0
      End
      Begin VB.Menu mnuSubSched 
         Caption         =   "Help"
         Index           =   1
      End
      Begin VB.Menu mnuSubSched 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSubSched 
         Caption         =   "About"
         Index           =   3
      End
      Begin VB.Menu mnuSubSched 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSubSched 
         Caption         =   "Exit"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NIF_ICON                              As Long = &H2
Private Const NIF_MESSAGE                           As Long = &H1
Private Const NIF_TIP                               As Long = &H4
Private Const NIM_ADD                               As Long = &H0
Private Const NIM_MODIFY                            As Long = &H1
Private Const NIM_DELETE                            As Long = &H2
Private Const MAX_TOOLTIP                           As Integer = 64
Private Const WM_MOUSEMOVE                          As Long = &H200
Private Const WM_LBUTTONDBLCLK                      As Long = &H203
Private Const WM_LBUTTONDOWN                        As Long = &H201
Private Const WM_LBUTTONUP                          As Long = &H202
Private Const WM_RBUTTONDBLCLK                      As Long = &H206
Private Const WM_RBUTTONDOWN                        As Long = &H204
Private Const WM_RBUTTONUP                          As Long = &H205

Private Type POINTAPI
    x                                               As Long
    y                                               As Long
End Type

Private Type RECT
    Left                                            As Long
    Top                                             As Long
    Right                                           As Long
    Bottom                                          As Long
End Type

Private Type NOTIFYICONDATA
    cbSize                                          As Long
    hwnd                                            As Long
    uID                                             As Long
    uFlags                                          As Long
    uCallbackMessage                                As Long
    hIcon                                           As Long
    szTip                                           As String * MAX_TOOLTIP
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                       lpData As NOTIFYICONDATA) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private nfIconData  As NOTIFYICONDATA


Public Sub ShowMenu()

Dim lR As Long
Dim tP As POINTAPI

    SetForegroundWindow Me.hwnd
    GetCursorPos tP
    

End Sub

Public Property Get ToolTip() As String

Dim sTip As String
Dim iPos As Long

    sTip = nfIconData.szTip
    iPos = InStr(sTip, vbNullChar)
    If iPos <> 0 Then
        sTip = Left$(sTip, iPos - 1)
    End If
    ToolTip = sTip

End Property

Public Property Let ToolTip(ByVal sTip As String)

    With nfIconData
        If (sTip & vbNullChar <> .szTip) Then
            .szTip = sTip & vbNullChar
            .uFlags = NIF_TIP
            Shell_NotifyIcon NIM_MODIFY, nfIconData
        End If
    End With

End Property

Public Sub Initialise()

    With nfIconData
        .hwnd = Me.hwnd
        .uID = Me.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.handle
        .szTip = App.FileDescription & vbNullChar
        .cbSize = Len(nfIconData)
    End With
    Shell_NotifyIcon NIM_ADD, nfIconData

    ToolTip = "NSP Scheduler Lite - Click for Options"

End Sub

Private Sub Form_Load()

    Initialise

End Sub

Private Sub Form_LostFocus()
    Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

Dim lX As Long

    lX = Me.ScaleX(x, Me.ScaleMode, vbPixels)
    Select Case lX
    Case WM_LBUTTONDBLCLK
        frmSchedule.Show
    Case WM_RBUTTONUP
        Me.PopupMenu Me.mnuSched
    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Me.Hide
    Shell_NotifyIcon NIM_DELETE, nfIconData

End Sub

Private Sub mnuSubSched_Click(Index As Integer)
'/* popup menu items

On Error Resume Next

    Select Case Index
    Case 0
        frmSchedule.Show
    Case 1
        frmHelp.Show
    Case 3
        frmAbout.Show
    Case 5
        Dim iQuery As Integer
        iQuery = MsgBox("Do you want to Stop the Scheduler?" & vbNewLine & _
        "Scheduled tasks will Not be Run, and All Events will be Cleared.", vbYesNo, "Unload Scheduler?")
        If iQuery = 7 Then Exit Sub
        Unload_Scheduler
        Application_Terminate
    End Select

On Error GoTo 0

End Sub


