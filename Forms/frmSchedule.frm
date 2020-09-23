VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSchedule 
   BackColor       =   &H00FFFFFF&
   Caption         =   "NSPowertools - Scheduler Lite 1.0"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
      Height          =   3315
      Left            =   270
      TabIndex        =   28
      Top             =   3690
      Width           =   8325
      Begin prjCSchedule.isButton cmdHelp 
         Height          =   405
         Left            =   150
         TabIndex        =   35
         Top             =   2790
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   714
         Icon            =   "frmSchedule.frx":57E2
         Style           =   8
         Caption         =   "Help"
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
      Begin MSComctlLib.ListView lstJobs 
         Height          =   2355
         Left            =   150
         TabIndex        =   29
         Top             =   300
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   4154
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin prjCSchedule.isButton cmdJob 
         Height          =   405
         Index           =   0
         Left            =   3180
         TabIndex        =   36
         Top             =   2790
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   714
         Icon            =   "frmSchedule.frx":57FE
         Style           =   8
         Caption         =   "Add Job"
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
      Begin prjCSchedule.isButton cmdJob 
         Height          =   405
         Index           =   1
         Left            =   4410
         TabIndex        =   37
         Top             =   2790
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   714
         Icon            =   "frmSchedule.frx":581A
         Style           =   8
         Caption         =   "Remove Job"
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
      Begin prjCSchedule.isButton cmdJob 
         Height          =   405
         Index           =   2
         Left            =   5790
         TabIndex        =   38
         Top             =   2790
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   714
         Icon            =   "frmSchedule.frx":5836
         Style           =   8
         Caption         =   "Clear Jobs"
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
      Begin prjCSchedule.isButton cmdJob 
         Height          =   405
         Index           =   3
         Left            =   7020
         TabIndex        =   40
         Top             =   2790
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   714
         Icon            =   "frmSchedule.frx":5852
         Style           =   8
         Caption         =   "Finish"
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Application"
      Height          =   735
      Left            =   270
      TabIndex        =   26
      Top             =   2850
      Width           =   8325
      Begin VB.TextBox txtCommand 
         Height          =   285
         Left            =   210
         TabIndex        =   27
         Top             =   270
         Width           =   6645
      End
      Begin prjCSchedule.isButton cmdBrowse 
         Height          =   345
         Left            =   7020
         TabIndex        =   39
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         Icon            =   "frmSchedule.frx":586E
         Style           =   8
         Caption         =   "Browse"
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      Height          =   2505
      Left            =   6000
      TabIndex        =   25
      Top             =   300
      Width           =   2595
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   120
         ScaleHeight     =   885
         ScaleWidth      =   2385
         TabIndex        =   30
         Top             =   330
         Width           =   2385
         Begin VB.OptionButton optMode 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Run on System Start Up"
            Height          =   225
            Index           =   1
            Left            =   30
            TabIndex        =   32
            Top             =   600
            Width           =   2115
         End
         Begin VB.OptionButton optMode 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Run Once"
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   31
            Top             =   330
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Up:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H004A4A4A&
            Height          =   210
            Index           =   0
            Left            =   30
            TabIndex        =   33
            Top             =   30
            Width           =   690
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Time"
      Height          =   2505
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   5655
      Begin VB.PictureBox picType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   0
         Left            =   1620
         ScaleHeight     =   1065
         ScaleWidth      =   1185
         TabIndex        =   1
         Top             =   330
         Width           =   1215
         Begin MSComCtl2.DTPicker dtpFreqTime 
            Height          =   315
            Left            =   30
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   270
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:mm tt"
            Format          =   22937603
            CurrentDate     =   36742
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H004A4A4A&
            Height          =   210
            Index           =   3
            Left            =   30
            TabIndex        =   3
            Top             =   30
            Width           =   465
         End
      End
      Begin VB.PictureBox picType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   3540
         ScaleHeight     =   705
         ScaleWidth      =   1365
         TabIndex        =   4
         Top             =   330
         Width           =   1395
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   315
            Left            =   30
            TabIndex        =   5
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   22937601
            CurrentDate     =   38752
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H004A4A4A&
            Height          =   210
            Index           =   5
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   405
         End
      End
      Begin VB.CheckBox chkDuration 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recurring"
         ForeColor       =   &H004A4A4A&
         Height          =   225
         Left            =   4380
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1080
      End
      Begin VB.PictureBox picOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   120
         ScaleHeight     =   1725
         ScaleWidth      =   945
         TabIndex        =   16
         Top             =   330
         Width           =   945
         Begin VB.OptionButton optOccurs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Daily"
            ForeColor       =   &H004A4A4A&
            Height          =   210
            Index           =   1
            Left            =   30
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   600
            Width           =   765
         End
         Begin VB.OptionButton optOccurs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Weekly"
            ForeColor       =   &H004A4A4A&
            Height          =   210
            Index           =   2
            Left            =   30
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   870
            Width           =   1005
         End
         Begin VB.OptionButton optOccurs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hourly"
            ForeColor       =   &H004A4A4A&
            Height          =   210
            Index           =   0
            Left            =   30
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   330
            Width           =   765
         End
         Begin VB.OptionButton optOccurs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Specify"
            ForeColor       =   &H004A4A4A&
            Height          =   210
            HelpContextID   =   30
            Index           =   3
            Left            =   30
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1065
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Interval:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H004A4A4A&
            Height          =   210
            Index           =   1
            Left            =   30
            TabIndex        =   21
            Top             =   30
            Width           =   660
         End
      End
      Begin VB.PictureBox picType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2055
         Index           =   1
         Left            =   1590
         ScaleHeight     =   2025
         ScaleWidth      =   1305
         TabIndex        =   7
         Top             =   330
         Width           =   1335
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tuesday"
            ForeColor       =   &H004A4A4A&
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   570
            Width           =   1110
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Monday"
            ForeColor       =   &H004A4A4A&
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   330
            Width           =   1050
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Saturday"
            ForeColor       =   &H004A4A4A&
            Height          =   195
            Index           =   5
            Left            =   30
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1530
            Width           =   1020
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Friday"
            ForeColor       =   &H004A4A4A&
            Height          =   195
            Index           =   4
            Left            =   30
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1290
            Width           =   840
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Thursday"
            ForeColor       =   &H004A4A4A&
            Height          =   195
            Index           =   3
            Left            =   30
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1050
            Width           =   1035
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Wednesday"
            ForeColor       =   &H004A4A4A&
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   810
            Width           =   1215
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sunday"
            ForeColor       =   &H004A4A4A&
            Height          =   195
            Index           =   6
            Left            =   30
            TabIndex        =   8
            Top             =   1770
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Weekly:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H004A4A4A&
            Height          =   210
            Index           =   2
            Left            =   30
            TabIndex        =   15
            Top             =   30
            Width           =   645
         End
      End
      Begin VB.Label lblNotice 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "At:"
         Height          =   210
         Left            =   1080
         TabIndex        =   24
         Top             =   660
         Width           =   210
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "On the:"
         Height          =   210
         Left            =   2940
         TabIndex        =   23
         Top             =   660
         Width           =   525
      End
   End
   Begin MSComDlg.CommonDialog cdFolder 
      Left            =   0
      Top             =   6420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Arial"
      Orientation     =   2
   End
   Begin VB.PictureBox pic16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1110
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   34
      Top             =   6690
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   540
      Top             =   6330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":588A
            Key             =   "dft"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3780
      Picture         =   "frmSchedule.frx":5C24
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1860
      Picture         =   "frmSchedule.frx":6D22
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2820
      Picture         =   "frmSchedule.frx":7E20
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   900
      Picture         =   "frmSchedule.frx":8F1E
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   930
      Picture         =   "frmSchedule.frx":A01C
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   900
      Picture         =   "frmSchedule.frx":AB4E
      ScaleHeight     =   435
      ScaleWidth      =   3750
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1710
      Visible         =   0   'False
      Width           =   3750
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/* icon flags
Private Const BASIC_SHGFI_FLAGS As Double = _
&H4 Or &H200 Or &H400 Or &H2000 Or &H4000

'/* icon struct
Private Type SHFILEINFO
    hIcon                              As Long
    iIcon                              As Long
    dwAttributes                       As Long
    szDisplayName                      As String * 260
    szTypeName                         As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwFileAttributes As Long, _
                                                                                 psfi As SHFILEINFO, _
                                                                                 ByVal cbSizeFileInfo As Long, _
                                                                                 ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, _
                                                            ByVal i As Long, _
                                                            ByVal hDCDest As Long, _
                                                            ByVal x As Long, _
                                                            ByVal y As Long, _
                                                            ByVal Flags As Long) As Long

Private m_cSIcon    As Collection
Private tShInfo     As SHFILEINFO
Private m_NCDG      As cNeoClass
Public m_iCurrSkin  As Integer

Private Sub cmdAbout_Click()
    frmAbout.Show vbModeless, Me
End Sub

Private Sub cmdBrowse_Click()
'/* browse for folder

On Error GoTo Handler

    With cdFolder
        .DialogTitle = "Select A File"
        .InitDir = Left$(App.Path, 3)
        .CancelError = True
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        txtCommand = .FileName
    End With

Handler:

End Sub

Private Sub cmdHelp_Click()
'/* help dialog

    frmHelp.Show vbModeless, Me

End Sub

Private Sub cmdJob_Click(Index As Integer)
'/* command list

Dim sTimeStr As String
Dim cDay     As CheckBox
Dim lInd     As Long

    Select Case Index
    '/* add job
    Case 0
        '/* test parameters
        If Len(txtCommand.Text) < 4 Then
            MsgBox "Please enter in an application to launch before proceeding!", _
            vbInformation, "No Application Specified"
            Exit Sub
        End If
        
        Select Case True
        '/* hourly
        Case optOccurs(0).Value
            sTimeStr = Format$(dtpFreqTime.Minute, "00")
            mScheduler.Scheduler_Add sTimeStr, txtCommand.Text, _
            CBool(chkDuration.Value), Hourly
            
        '/* daily
        Case optOccurs(1).Value
            sTimeStr = Format$(dtpFreqTime.Value, "hh:mm")
            mScheduler.Scheduler_Add sTimeStr, txtCommand.Text, _
            CBool(chkDuration.Value), daily
            
            
        '/* weekly
        Case optOccurs(2).Value
            For Each cDay In chkDay
                If cDay.Value = 1 Then
                    sTimeStr = cDay.Caption + Chr$(32) + Format$(dtpFreqTime.Value, "hh:mm")
                    mScheduler.Scheduler_Add sTimeStr, txtCommand.Text, _
                    CBool(chkDuration.Value), weekly
                End If
            Next cDay
            
        '/*specific
        Case optOccurs(3).Value
            sTimeStr = dtpDate.Value & Chr$(32) + Format$(dtpFreqTime.Value, "hh:mm")
            mScheduler.Scheduler_Add Format$(sTimeStr, "mm/dd/yyyy HH:nn"), txtCommand.Text, _
            CBool(chkDuration.Value), specific
        End Select
        '/* refresh list
        List_Init
        
    '/* remove a job
    Case 1
        lInd = lstJobs.SelectedItem.Index
        If MsgBox("Do you want to remove the job: " + lstJobs.SelectedItem.Text, _
        vbOKCancel, "Remove a Task!") = vbOK Then
            mScheduler.Scheduler_Remove lInd
        End If
        If Not mScheduler.Jobs_Pending Then
            mScheduler.Scheduler_Clear
        End If
        
        '/* refresh list
        List_Init
        
    '/* clear scheduler
    Case 2
        If MsgBox("Remove All Jobs from the Scheduler?", vbOKCancel, _
        "Clear Tasks!") = vbOK Then
            mScheduler.Scheduler_Clear
            List_Init
        End If
        
    '/* finish
    Case 3
        Unload Me

    End Select

End Sub

Private Sub Form_Load()
'/* set control options

Dim i As Integer

    Set m_cSIcon = New Collection
    Set_Skin
    optOccurs(1).Value = True
    
    With lstJobs
        .SmallIcons = iml16
        .AllowColumnReorder = True
        .Font.SIZE = 8
        .Appearance = ccFlat
        .FullRowSelect = True
        .GridLines = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Name", (.Width / 32) * 8
        .ColumnHeaders.Add 2, , "Path", (.Width / 32) * 12
        .ColumnHeaders.Add 3, , "Schedule", (.Width / 32) * 8
        .ColumnHeaders.Add 4, , "Recurring", (.Width / 32) * 4
    End With
    
    For i = 0 To chkDay.Count - 1
        chkDay(i).ToolTipText = "Run Task on these Days"
    Next i
    
    optMode(0).ToolTipText = "Run Task Manager One Time"
    optMode(1).ToolTipText = "Keep Task Manager Loaded"
    chkDuration.ToolTipText = "Repeat Scheduled Task"
    txtCommand.ToolTipText = "Choose the Application you want to Run"
    optMode(1).Value = mScheduler.Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "stay") > 0
    List_Init

End Sub

Private Sub Hide_Controls(ByVal iIndex As Integer)
'/* controls arrangement

Dim mPic As PictureBox

On Error Resume Next

    For Each mPic In picType
        mPic.Visible = False
        mPic.BorderStyle = 0
    Next mPic
    lblDate.Visible = False
    
    Select Case iIndex
    '/* hourly
    Case 0
        picType(0).Left = 1690
        picType(0).Visible = True
        lblNotice.Caption = "Start At:"
        dtpFreqTime.ToolTipText = "Start Task in Number of Minutes"
        
    '/* daily
    Case 1
        picType(0).Left = 1300
        picType(0).Visible = True
        lblNotice.Caption = "At:"
        lblDate.Visible = False
        dtpFreqTime.ToolTipText = "Start Task at this Time"
        
    '/* weekly
    Case 2
        picType(0).Left = 3250
        picType(0).Visible = True
        picType(1).Visible = True
        lblNotice.Caption = "Every:"
        lblDate.Left = 3000
        lblDate.Caption = "At:"
        lblDate.Visible = True
        dtpFreqTime.ToolTipText = "Start Task at this Time"
        
    '/* specify
    Case 3
        picType(0).Left = 3250
        picType(0).Visible = True
        picType(2).Left = 1350
        picType(2).Visible = True
        lblNotice.Caption = "On:"
        With lblDate
            .Left = 3020
            .Caption = "At:"
            .Visible = True
        End With
        dtpDate.ToolTipText = "Start Task on this Date"
        dtpFreqTime.ToolTipText = "Start Task at this Time"
    End Select

On Error GoTo 0

End Sub

Public Sub List_Init()
'/* load listview

Dim vItem  As Variant
Dim sCmd() As String
Dim cItem  As ListItem
Dim sName  As String
Dim sKey   As String

On Error Resume Next

    lstJobs.ListItems.Clear
    For Each vItem In mScheduler.Read_MultiCN(HKEY_CURRENT_USER, m_sRegPath, "ctasks")
        If Len(vItem) = 0 Then GoTo skip
        sCmd = Split(vItem, Chr$(44))
        
        '/* filter switches
        If InStr(1, sCmd(1), " -") > 0 Then
            sCmd(1) = Left$(sCmd(1), InStrRev(sCmd(1), " -") - 1)
        End If
        If InStr(1, sCmd(1), " /") > 0 Then
            sCmd(1) = Left$(sCmd(1), InStrRev(sCmd(1), " /") - 1)
        End If
        
        '/* get icon key
        sKey = Get_SmallIcon(sCmd(1))
        '/* add name
        sName = Mid$(sCmd(1), InStrRev(sCmd(1), Chr$(92)) + 1)
        Set cItem = lstJobs.ListItems.Add(Text:=Trim$(sName), SmallIcon:=sKey)
        cItem.SubItems(1) = Trim$(sCmd(1))
        
        '/* get schedule type
        Select Case CLng(Trim$(sCmd(3)))
        Case 1
            sCmd(0) = "At " + Format$(Now, "HH") + Chr$(58) + sCmd(0) + " Minutes"
        Case 2
            sCmd(0) = "At " + sCmd(0) + ":00"
        Case 3
            sCmd(0) = "On " + sCmd(0) + ":00"
        Case 4
            sCmd(0) = "On " + sCmd(0) + ":00"
        End Select
        
        cItem.SubItems(2) = sCmd(0)
        '/* repeating
        cItem.SubItems(3) = Trim$(sCmd(2))
        ReDim sCmd(0 To 3) As String
skip:
    Next vItem
    
On Error GoTo 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

'/* update engine status
'/* scheduler is empty - shut it down
    
    If mScheduler Is Nothing Then
        Unload_Scheduler
        Application_Terminate
        Exit Sub
    End If
    
    With mScheduler
        If Not .p_bIsActive Then
            .Scheduler_End
            Unload_Scheduler
            Application_Terminate
        Else
            If Me.Visible Then
                If Not .p_bIsLoaded And .Jobs_Pending Then
                    Me.Hide
                    Load frmSysTray
                    Cancel = 1
                Else
                    Unload frmSysTray
                    Unload Me
                End If
            End If
        End If
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Me.Hide
    Set mScheduler = Nothing
    Set m_NCDG = Nothing
    
End Sub

Private Sub optMode_Click(Index As Integer)
'/* run on startup

    Select Case Index
    Case 0
        With mScheduler
            .Delete_Value HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "nspsched"
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "stay", 0
        End With
    Case 1
        With mScheduler
            .Write_String HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
            "nspsched", App.Path & "\nspsched.exe -s"
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "stay", 1
        End With
    End Select

End Sub

Private Sub optOccurs_Click(Index As Integer)
'/* time format options

    Hide_Controls Index
    Select Case Index
    '/* hourly
    Case 0
        With dtpFreqTime
            .Width = 700
            .Format = dtpCustom
            .UpDown = True
            .CustomFormat = "mm"
        End With
        
    '/* daily
    Case 1
        With dtpFreqTime
            .Width = 1200
            .Format = dtpCustom
            .UpDown = True
            .CustomFormat = "hh:mm tt"
        End With
        
    '/* weekly
    Case 2
        With dtpFreqTime
            .Width = 1200
            .Format = dtpCustom
            .UpDown = True
            .CustomFormat = "hh:mm tt"
        End With
        
    '/* user specified
    Case 3
        With dtpFreqTime
            .Width = 1200
            .Format = dtpCustom
            .UpDown = True
            .CustomFormat = "hh:mm tt"
        End With
    End Select

End Sub

Private Function Get_SmallIcon(ByVal sFile As String) As String
'/* get listview icons
'/* names image keys with file extensions,
'/* for reusable image items in imagelist

Dim hSIcon As Long
Dim imgObj As ListImage
Dim sKey   As String
Dim sExt   As String

On Error Resume Next

    '/* get associated extension
    sExt = LCase$(Mid$(sFile, InStrRev(sFile, Chr$(92)) + 1))
    '/* get handle to icon
    hSIcon = SHGetFileInfo(sFile, 0&, tShInfo, Len(tShInfo), BASIC_SHGFI_FLAGS Or &H1)
    '/* load icon to picturebox
    If Not hSIcon = 0 Then
        With pic16
            Set .Picture = LoadPicture("")
            .AutoRedraw = True
            ImageList_Draw hSIcon, tShInfo.iIcon, .hdc, 0&, 0&, &H1
            .Refresh
        End With
        '/* test for icon presence in collection
        sKey = m_cSIcon.Item(sExt)
        '/* if not present, add to collection
        If LenB(sKey) = 0 Then
            m_cSIcon.Add 1, sExt
            '/* add icon to image list
            '/* use file extension as image key
            Set imgObj = iml16.ListImages.Add(Key:=sExt, Picture:=pic16.Image)
        End If
        Get_SmallIcon = sExt
    Else
        '/* no icon, use default image
        Get_SmallIcon = "dft"
    End If

On Error GoTo 0

End Function

Private Sub Set_Skin()
'/* load skin

On Error Resume Next

    Set m_NCDG = Nothing
    Set m_NCDG = New cNeoClass
    
    '/* mac
        With m_NCDG
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


