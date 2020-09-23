VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduler Lite - Help"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   330
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":57E2
      Top             =   240
      Width           =   4905
   End
   Begin prjCSchedule.isButton cmdAbout 
      Height          =   405
      Left            =   330
      TabIndex        =   1
      Top             =   3990
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   714
      Icon            =   "frmHelp.frx":57E8
      Style           =   8
      Caption         =   "About Us"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackDrop        =   0
   End
   Begin prjCSchedule.isButton cmdClose 
      Height          =   405
      Left            =   4080
      TabIndex        =   2
      Top             =   3990
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   714
      Icon            =   "frmHelp.frx":5804
      Style           =   8
      Caption         =   "Exit"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackDrop        =   0
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3780
      Picture         =   "frmHelp.frx":5820
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1860
      Picture         =   "frmHelp.frx":691E
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2820
      Picture         =   "frmHelp.frx":7A1C
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   900
      Picture         =   "frmHelp.frx":8B1A
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   900
      Picture         =   "frmHelp.frx":9C18
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2010
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   900
      Picture         =   "frmHelp.frx":A74A
      ScaleHeight     =   435
      ScaleWidth      =   3750
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   3750
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_NCDH      As cNeoClass

Private Sub cmdAbout_Click()
    frmAbout.Show vbModeless, Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set_Skin
    Set_Help
End Sub

Private Sub Set_Skin()
'/* load skin

On Error Resume Next

On Error Resume Next

    Set m_NCDH = New cNeoClass
    
    '/* mac
        With m_NCDH
            Set .p_ICaption = picBar.Picture
            Set .p_IBorders = picFrame.Picture
            Set .p_ICBoxMin = picMin.Picture
            Set .p_ICBoxMax = picMax.Picture
            Set .p_ICBoxRst = picRst.Picture
            Set .p_ICBoxCls = picCls.Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HEEEEEE
            Dim lLen As Long
            lLen = (LenB(Me.Caption) * Screen.TwipsPerPixelX)
            .p_CaptionOffset = ((Me.Width - lLen) / Screen.TwipsPerPixelX) / 3
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -11
            .p_ButtonOffsetY = 0
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_Offset = 0
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 0
            .Attach
        End With
        
End Sub

Private Sub Set_Help()

Dim sHelp As String

    sHelp = "Task Scheduler Help" & vbCrLf & _
    vbCrLf & "Interval: The time you want the event to begin. Example: Thursday " & _
    "at 17:30, would start the application on Thursday at 5:30 pm." & _
    vbCrLf & vbCrLf & "Recurring: When checked, the event will reoccur at the same time every day." & _
    vbCrLf & vbCrLf & "Start Up: When selected, the scheduler will start automatically each " & _
    "time you restart your computer." & _
    vbCrLf & vbCrLf & "Application: Select the program you want to run with the scheduler." & _
    "Add Job: Once you have chosen the time and application, this will " & _
    "add it to the scheduler." & _
    vbCrLf & vbCrLf & "Remove Job: This option removes the job from the scheduler." & _
    vbCrLf & vbCrLf & "Clear Jobs: This option removes all jobs from the scheduler." & _
    vbCrLf & vbCrLf & "Finish: This closes the scheduler window. To stop the scheduler " & _
    "completely, just remove all tasks from the job list."
        
    txtHelp.Text = ""
    txtHelp.Text = sHelp
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Set m_NCDH = Nothing
End Sub
