VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduler Lite - About Us"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbout 
      Height          =   2355
      Left            =   330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAbout.frx":57E2
      Top             =   330
      Width           =   4395
   End
   Begin prjCSchedule.isButton cmdClose 
      Height          =   405
      Left            =   3570
      TabIndex        =   1
      Top             =   2850
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   714
      Icon            =   "frmAbout.frx":57E8
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
      Left            =   3450
      Picture         =   "frmAbout.frx":5804
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1530
      Picture         =   "frmAbout.frx":6902
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2490
      Picture         =   "frmAbout.frx":7A00
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   570
      Picture         =   "frmAbout.frx":8AFE
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   570
      Picture         =   "frmAbout.frx":9BFC
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   570
      Picture         =   "frmAbout.frx":A72E
      ScaleHeight     =   435
      ScaleWidth      =   3750
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   3750
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_NCDA      As cNeoClass

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Load()
    Set_Skin
    Set_About
End Sub

Private Sub Set_About()

Dim sAbout As String

    sAbout = "~About NSP Task Scheduler~" & vbCrLf & _
    vbCrLf & "NSP Task Scheduler is provided as a freeware alternative " & _
    "to the native scheduler component in Windows." & _
    vbCrLf & vbCrLf & "NSPowertools.com offers several free software components to " & _
    "the public as a service." & _
    vbCrLf & vbCrLf & "If you would like to contact us about our products, email: support@nspowertools.com. " & _
    "For more information on our products or services, feel free to " & _
    "email us, or visit our website at www.nspowertools.com."

    txtAbout.Text = sAbout
        
End Sub

Private Sub Set_Skin()
'/* load skin

On Error Resume Next

On Error Resume Next

    Set m_NCDA = New cNeoClass
    
    '/* mac
        With m_NCDA
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

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Set m_NCDA = Nothing
End Sub

