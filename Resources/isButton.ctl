VERSION 5.00
Begin VB.UserControl isButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   Begin VB.PictureBox m_About 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7440
      ScaleHeight     =   2175
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "isButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************
'
' Control Name: isButton
'
' Author:       Fred.cpp
'               fred_cpp@msn.com
'
' Page:         http://mx.geocities.com/fred_cpp/
'
' Current
' Version:      3.0
'
' Description:  Multy Style Command Button. Made from scratch,
'               API drawn.
'               I started this project almost 4 years ago, I've
'               learned a lot since then, I think this version
'               Is the first one that can be used In real world,
'               I've keep this control for my own use for a
'               long Time, But I've used It for personal
'               projects, The only one public release was here:
'               http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=51271&lngWId=1
'               I Used It here cuz there was not too much to
'               show then. Afther some months I feel like I
'               could make the final Version, So Here I am.
'               I hope you like It And If you find It useful
'               please vote, leave comments and suggestions,
'               Everything Is wellcome.
'               Best Regards.
'
' Started on:   6-6-2004
'
' 2004-12-6
' Added Soft, Flat, java and Office XP Style
'
' 2004-18-6
' Added Windows XP style
'
' 2004-20-6
' Enhanced Office XP Style
'
' 2004-2-7
' Code Clean Up. Removed Timer, Now Uses the
' Paul Caton's self subclaser
'
' 2004-3-7
' Added Windows Themed Style
'
' 2004-4-7
' Added Galaxy style
'
' 2004-7-7
' Added Enabled/ShowFocus and Custom Colors
'
' 2004-8-7
' Added Plastik Style
'
' 2004-14-7
' Added Keramik Style
'
' 2004-15-7
' Added Multyple styles for Tooltips!
'
' 2004-19-7
' Added Custom WinXP Colors, Disabled State
' for the WinXP style(Normal and custom colors)
'
' 2004-20-7
' Added Mac OSX Style
'
' 2004-24-7
' Keramik Rewritten from scratch
'
' 2004-3-9
' Afther some weeks of inactivity added disabled state for
' plastik and keramik styles
'
' 2004-6-9
' Added a nice About box without a extra form!
' Also some corrections to the MAc OSX Style, But still
' needs more work.
'
' 2004-7-9
' PSC downloads are down. The post was deleted when I tried
' to update. I'm trying to make some improvements to the
' MAc OSX Style, but Nothing too important. just a bugfix
' with the disabled state.
'
' 2004-8-9
' Added font property, also tooltips info is not reseted
' each run. Thanks to the comments of the PSC users I have
' lot of things to improve!
'
' 2004-10-9
' OK, Lot, I mean LOT of suggestions, I would love to be at
' school when I had enought time to code (And I was young
' and was able to sleep one or two hours per day=( )
'
' 2004-12-9
' added font and small offsets for each button style. Also
' Enum names are fixed.
'
' 2004-15-9
' Version 3.0 The last one. added CheckBox Type and value
' property. maybe some Other smal fixes will be made
' soon...maybe.
' Have fun
'
' the code starts Here!

Option Explicit

Private Const strCurrentVersion              As String = "3.0"
Private Const COLOR_ACTIVEBORDER             As Long = 10
Private Const COLOR_ACTIVECAPTION            As Long = 2
Private Const COLOR_ADJ_MAX                  As Long = 100
Private Const COLOR_ADJ_MIN                  As Long = -100
Private Const COLOR_APPWORKSPACE             As Long = 12
Private Const COLOR_BACKGROUND               As Long = 1
Private Const COLOR_BTNFACE                  As Long = 15
Private Const COLOR_BTNHIGHLIGHT             As Long = 20
Private Const COLOR_BTNSHADOW                As Long = 16
Private Const COLOR_BTNTEXT                  As Long = 18
Private Const COLOR_CAPTIONTEXT              As Long = 9
Private Const COLOR_GRAYTEXT                 As Long = 17
Private Const COLOR_HIGHLIGHT                As Long = 13
Private Const COLOR_HIGHLIGHTTEXT            As Long = 14
Private Const COLOR_INACTIVEBORDER           As Long = 11
Private Const COLOR_INACTIVECAPTION          As Long = 3
Private Const COLOR_INACTIVECAPTIONTEXT      As Long = 19
Private Const COLOR_MENU                     As Long = 4
Private Const COLOR_MENUTEXT                 As Long = 7
Private Const COLOR_SCROLLBAR                As Long = 0
Private Const COLOR_WINDOW                   As Long = 5
Private Const COLOR_WINDOWFRAME              As Long = 6
Private Const COLOR_WINDOWTEXT               As Long = 8
Private Const COLOR_INFOTEXT                 As Long = 23
Private Const COLOR_INFOBK                   As Long = 24

'Gradient Constants
Private Const GRADIENT_FILL_RECT_H           As Long = &H0
Private Const GRADIENT_FILL_RECT_V           As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE         As Long = &H2
Private Const GRADIENT_FILL_OP_FLAG          As Long = &HFF

'flags for DrawFrameControl
Private Const DFC_CAPTION                    As Long = 1      'Title bar
Private Const DFC_MENU                       As Long = 2      'Menu
Private Const DFC_SCROLL                     As Long = 3      'Scroll bar
Private Const DFC_BUTTON                     As Long = 4      'Standard button
Private Const DFCS_CAPTIONCLOSE              As Long = &H0    'Close button
Private Const DFCS_CAPTIONMIN                As Long = &H1    'Minimize button
Private Const DFCS_CAPTIONMAX                As Long = &H2    'Maximize button
Private Const DFCS_CAPTIONRESTORE            As Long = &H3    'Restore button
Private Const DFCS_CAPTIONHELP               As Long = &H4    'Windows 95 only: Help button
Private Const DFCS_MENUARROW                 As Long = &H0    'Submenu arrow
Private Const DFCS_MENUCHECK                 As Long = &H1    'Check mark
Private Const DFCS_MENUBULLET                As Long = &H2    'Bullet
Private Const DFCS_MENUARROWRIGHT            As Long = &H4
Private Const DFCS_SCROLLUP                  As Long = &H0    'Up arrow of scroll bar
Private Const DFCS_SCROLLDOWN                As Long = &H1    'Down arrow of scroll bar
Private Const DFCS_SCROLLLEFT                As Long = &H2    'Left arrow of scroll bar
Private Const DFCS_SCROLLRIGHT               As Long = &H3    'Right arrow of scroll bar
Private Const DFCS_SCROLLCOMBOBOX            As Long = &H5    'Combo box scroll bar
Private Const DFCS_SCROLLSIZEGRIP            As Long = &H8    'Size grip
Private Const DFCS_SCROLLSIZEGRIPRIGHT       As Long = &H10
Private Const DFCS_BUTTONCHECK               As Long = &H0    'Check box
Private Const DFCS_BUTTONRADIO               As Long = &H4    'Radio button
Private Const DFCS_BUTTON3STATE              As Long = &H8    'Three-state button
Private Const DFCS_BUTTONPUSH                As Long = &H10 'Push button
Private Const DFCS_INACTIVE                  As Long = &H100 'Button is inactive (grayed)
Private Const DFCS_PUSHED                    As Long = &H200 'Button is pushed
Private Const DFCS_CHECKED                   As Long = &H400 'Button is checked
Private Const DFCS_ADJUSTRECT                As Long = &H2000
Private Const DFCS_FLAT                      As Long = &H4000 'Button has a flat border
Private Const DFCS_MONO                      As Long = &H8000 'Button has a monochrome border
Private Const BDR_RAISEDOUTER                As Long = &H1
Private Const BDR_SUNKENOUTER                As Long = &H2
Private Const BDR_RAISEDINNER                As Long = &H4
Private Const BDR_SUNKENINNER                As Long = &H8
Private Const BDR_OUTER                      As Long = &H3
Private Const BDR_INNER                      As Long = &HC
Private Const BDR_RAISED                     As Long = &H5
Private Const BDR_SUNKEN                     As Long = &HA
Private Const EDGE_RAISED                    As Double = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN                    As Double = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED                    As Double = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP                      As Double = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT                        As Long = &H1
Private Const BF_TOP                         As Long = &H2
Private Const BF_RIGHT                       As Long = &H4
Private Const BF_BOTTOM                      As Long = &H8
Private Const BF_TOPLEFT                     As Long = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT                    As Long = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT                  As Long = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT                 As Long = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT                        As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_DIAGONAL                    As Long = &H10
Private Const BF_DIAGONAL_ENDTOPRIGHT        As Long = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT         As Long = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT      As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT     As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_MIDDLE                      As Long = &H800 ' Fill in the middle.
Private Const BF_SOFT                        As Long = &H1000 ' Use for softer buttons.
Private Const BF_ADJUST                      As Long = &H2000 ' Calculate the space left over.
Private Const BF_FLAT                        As Long = &H4000 ' For flat rather than 3-D borders.
Private Const BF_MONO                        As Long = &H8000 ' For monochrome borders.
Private Const WM_LBUTTONDBLCLK               As Long = &H203
Private Const WM_RBUTTONDBLCLK               As Long = &H206
Private Const WM_LBUTTONDOWN                 As Long = &H201
Private Const WM_RBUTTONDOWN                 As Long = &H204
Private Const WM_MOUSEMOVE                   As Long = &H200
Private Const WM_MOUSELEAVE                  As Long = &H2A3
Private Const WM_MOVING                      As Long = &H216
Private Const WM_SIZING                      As Long = &H214
Private Const WM_EXITSIZEMOVE                As Long = &H232
Private Const WM_THEMECHANGED                As Long = &H31A
Private Const WM_SYSCOLORCHANGE              As Long = &H15
Private Const WM_USER                        As Long = &H400
Private Const GWL_STYLE                      As Long = -16
Private Const WS_CAPTION                     As Long = &HC00000
Private Const WS_THICKFRAME                  As Long = &H40000
Private Const WS_SYSMENU                     As Long = &H80000
Private Const WS_MINIMIZEBOX                 As Long = &H20000
Private Const SWP_REFRESH                    As Long = (&H1 Or &H2 Or &H4 Or &H20)
Private Const WS_EX_TOOLWINDOW               As Long = &H80
Private Const GWL_EXSTYLE                    As Long = -20
Private Const SW_SHOWDEFAULT                 As Long = 10
Private Const SW_SHOWMAXIMIZED               As Long = 3
Private Const SW_SHOWMINIMIZED               As Long = 2
Private Const SW_SHOWMINNOACTIVE             As Long = 7
Private Const SW_SHOWNA                      As Long = 8
Private Const SW_SHOWNOACTIVATE              As Long = 4
Private Const SW_SHOWNORMAL                  As Long = 1
Private Const SWP_FRAMECHANGED               As Long = &H20
Private Const SWP_DRAWFRAME                  As Long = SWP_FRAMECHANGED
Private Const SWP_HIDEWINDOW                 As Long = &H80
Private Const SWP_NOACTIVATE                 As Long = &H10
Private Const SWP_NOCOPYBITS                 As Long = &H100
Private Const SWP_NOMOVE                     As Long = &H2
Private Const SWP_NOOWNERZORDER              As Long = &H200
Private Const SWP_NOREDRAW                   As Long = &H8
Private Const SWP_NOREPOSITION               As Long = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE                     As Long = &H1
Private Const SWP_NOZORDER                   As Long = &H4
Private Const SWP_SHOWWINDOW                 As Long = &H40
Private Const HWND_TOPMOST                   As Long = -&H1
Private Const CW_USEDEFAULT                  As Long = &H80000000
Private Const RGN_AND                        As Long = &H1
Private Const RGN_OR                         As Long = &H2
Private Const RGN_XOR                        As Long = &H3
Private Const RGN_DIFF                       As Long = &H4
Private Const RGN_COPY                       As Long = &H5
Private Const DST_BITMAP                     As Long = &H4
Private Const DST_COMPLEX                    As Long = &H0
Private Const DST_ICON                       As Long = &H3
Private Const DSS_MONO                       As Long = &H80
Private Const DSS_NORMAL                     As Long = &H0
Private Const NULLREGION                     As Long = &H1    'Empty region
Private Const SIMPLEREGION                   As Long = &H2    'Rectangle Region
Private Const COMPLEXREGION                  As Long = &H3    'The region is complex
Private Const ALTERNATE                      As Long = 1
Private Const WINDING                        As Long = 2
Private Const TTS_NOPREFIX                   As Long = &H2
Private Const TTF_TRANSPARENT                As Long = &H100
Private Const TTF_CENTERTIP                  As Long = &H2
Private Const TTM_ADDTOOLA                   As Long = (WM_USER + 4)
Private Const TTM_ACTIVATE                   As Long = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA             As Long = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH             As Long = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR              As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR            As Long = (WM_USER + 20)
Private Const TTM_SETTITLE                   As Long = (WM_USER + 32)
Private Const TTS_BALLOON                    As Long = &H40
Private Const TTS_ALWAYSTIP                  As Long = &H1
Private Const TTF_SUBCLASS                   As Long = &H10
Private Const TOOLTIPS_CLASSA                As String = "tooltips_class32"
Private Const ALL_MESSAGES                   As Long = -1     'All messages added or deleted
Private Const GMEM_FIXED                     As Long = 0      'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC                    As Long = -4
Private Const PATCH_04                       As Long = 88
Private Const PATCH_05                       As Long = 93
Private Const PATCH_08                       As Long = 132    'Table A (after) address patch offset
Private Const PATCH_09                       As Long = 137

Private Type POINT
    x                                            As Long
    y                                            As Long
End Type

Private Type SIZE
    cx                                           As Long
    cy                                           As Long
End Type

Private Type RECT
    Left                                         As Long
    Top                                          As Long
    Right                                        As Long
    Bottom                                       As Long
End Type

Private Type RGB
    Red                                          As Byte
    Green                                        As Byte
    Blue                                         As Byte
End Type

Private Type tSubData
    hwnd                                         As Long
    nAddrSub                                     As Long
    nAddrOrig                                    As Long
    nMsgCntA                                     As Long
    nMsgCntB                                     As Long
    aMsgTblA()                                   As Long
    aMsgTblB()                                   As Long
End Type

Private Type TOOLINFO
    lSize                                        As Long
    lFlags                                       As Long
    lHwnd                                        As Long
    lId                                          As Long
    lpRect                                       As RECT
    hInstance                                    As Long
    lpStr                                        As String
    lParam                                       As Long
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize                          As Long
    dwMajorVersion                               As Long
    dwMinorVersion                               As Long
    dwBuildNumber                                As Long
    dwPlatformId                                 As Long
    szCSDVersion                                 As String * 128
    wServicePackMajor                            As Integer
    wServicePackMinor                            As Integer
    wSuiteMask                                   As Integer
    wProductType                                 As Byte
    wReserved                                    As Byte
End Type

Public Enum isbStyle
    [isbNormal] = &H0
    [isbSoft] = &H1
    [isbFlat] = &H2
    [isbJava] = &H3
    [isbOfficeXP] = &H4
    [isbWindowsXP] = &H5
    [isbWindowsTheme] = &H6
    [isbPlastik] = &H7
    [isbGalaxy] = &H8
    [isbKeramik] = &H9
    [isbMacOSX] = &HA
End Enum

Public Enum isbButtonType
    isbButton = &H0
    isbCheckBox = &H1
End Enum

Public Enum isbAlign
    [isbCenter] = &H0
    [isbLeft] = &H1
    [isbRight] = &H2
    [isbTop] = &H3
    [isbBottom] = &H4
End Enum

Private Enum isEstate
    statenormal = &H1
    stateHot = &H2
    statePressed = &H3
    statedisabled = &H4
    stateDefaulted = &H5
End Enum

Private Type msg
    hwnd                                         As Long
    message                                      As Long
    wParam                                       As Long
    lParam                                       As Long
    Time                                         As Long
    pt                                           As POINT
End Type

Private Type tagTRACKMOUSEEVENT
    cbSize                                       As Long
    dwFlags                                      As Long
    hwndTrack                                    As Long
    dwHoverTime                                  As Long
End Type

Private Type TRIVERTEX
    x                                            As Long
    y                                            As Long
    Red                                          As Integer
    Green                                        As Integer
    Blue                                         As Integer
    Alpha                                        As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft                                    As Long
    LowerRight                                   As Long
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1                                      As Long
    Vertex2                                      As Long
    Vertex3                                      As Long
End Type

Private Type DRAWTEXTPARAMS
    cbSize                                       As Long
    iTabLength                                   As Long
    iLeftMargin                                  As Long
    iRightMargin                                 As Long
    uiLengthDrawn                                As Long
End Type

Private Type BLENDFUNCTION
    BlendOp                                      As Byte
    BlendFlags                                   As Byte
    SourceConstantAlpha                          As Byte
    AlphaFormat                                  As Byte
End Type

Private Type RGBQUAD
    rgbBlue                                      As Byte
    rgbgreen                                     As Byte
    rgbred                                       As Byte
    rgbReserved                                  As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize                                       As Long
    biWidth                                      As Long
    biHeight                                     As Long
    biPlanes                                     As Integer
    biBitCount                                   As Integer
    biCompression                                As Long
    biSizeImage                                  As Long
    biXPelsPerMeter                              As Long
    biYPelsPerMeter                              As Long
    biClrUsed                                    As Long
    biClrImportant                               As Long
End Type

Private Type BITMAPINFO
    bmiHeader                                    As BITMAPINFOHEADER
    bmiColors                                    As RGBQUAD
End Type

Private Type UxTheme
    sClass                                       As String
    Part                                         As Long
    State                                        As Long
    hdc                                          As Long
    hwnd                                         As Long
    Left                                         As Long
    Top                                          As Long
    Width                                        As Long
    Height                                       As Long
    Text                                         As String
    TextAlign                                    As DrawTextFlags
    IconIndex                                    As Long
    himl                                         As Long
    RaiseError                                   As Boolean
    UseThemeSize                                 As Boolean
    UseTheme                                     As Boolean
    TextOffset                                   As Long
    RightTextOffset                              As Long
End Type

Private Enum DrawTextAdditionalFlags
    DTT_GRAYED = &H1
End Enum

Private Enum THEMESIZE
    TS_MIN
    TS_TRUE
    TS_DRAW
End Enum

Private Enum eMsgWhen
    MSG_AFTER = 1
    MSG_BEFORE = 2
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                       As Long
    dwFlags                                      As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                    As Long
    dwHoverTime                                  As Long
End Type

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

Private Enum GRADIENT_FILL_RECT
    FillHor = GRADIENT_FILL_RECT_H
    FillVer = GRADIENT_FILL_RECT_V
End Enum

Private Enum GRADIENT_TO_CORNER
    All
    TopLeft
    TopRight
    BottomLeft
    BottomRight
End Enum

Private Enum CRADIENT_DIRECTION
    DirectionSlash
    DirectionBackSlash
End Enum

Private Enum DrawTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

Private m_bOver                              As Boolean
Private m_bFocused                           As Boolean
Private m_iState                             As isEstate
Private m_iStyle                             As isbStyle
Private m_iNonThemeStyle                     As isbStyle
Private m_btnRect                            As RECT
Private m_txtRect                            As RECT
Private m_lRegion                            As Long
Private m_sCaption                           As String
Private m_CaptionAlign                       As isbAlign
Private m_IconAlign                          As isbAlign
Private m_Icon                               As StdPicture
Private m_Font                               As StdFont
Private m_IconSize                           As Long
Private m_bEnabled                           As Boolean
Private m_bShowFocus                         As Boolean
Private m_bUseCustomColors                   As Boolean
Private m_lBackColor                         As Long
Private m_lHighlightColor                    As Long
Private m_lFontColor                         As Long
Private m_lFontHighlightColor                As Long
Private m_sToolTipText                       As String
Private m_sTooltiptitle                      As String
Private m_lToolTipIcon                       As ttIconType
Private m_lToolTipType                       As ttStyleEnum
Private m_lttBackColor                       As Long
Private m_lttForeColor                       As Long
Private m_lttCentered                        As Boolean
Private m_lttHwnd                            As Long
Private m_ButtonType                         As isbButtonType
Private m_Value                              As Boolean
Private lPrevStyle                           As Long
Private sc_aSubData()                        As tSubData
Private bTrack                               As Boolean
Private bTrackUser32                         As Boolean
Private bInCtrl                              As Boolean
Private m_BDColor                            As OLE_COLOR
'Auxiliary Variables
Private lwFontAlign                          As Long
Private lPrevButton                          As Long
Private OSVI                                 As OSVERSIONINFOEX
Private ttip                                 As TOOLINFO
Public Event Click()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)
                                                                     
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, _
                                                  Source As Any, _
                                                  ByVal Length As Long)
                                                  
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, _
                                                                                ByVal lpString As String) As Long
                                                                                
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long
                                                        
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                                     ByVal dwBytes As Long) As Long
                                                     
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long
                                                      
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long
                                                                            
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long


Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
                                                                            
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, _
                                                                      ByVal lpCursorName As Long) As Long
                                                                      

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal XSrc As Long, _
                                             ByVal YSrc As Long, _
                                             ByVal dwRop As Long) As Long
                                             
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal x As Long, _
                                                 ByVal y As Long, _
                                                 ByVal nWidth As Long, _
                                                 ByVal nHeight As Long, _
                                                 ByVal hSrcDC As Long, _
                                                 ByVal XSrc As Long, _
                                                 ByVal YSrc As Long, _
                                                 ByVal nSrcWidth As Long, _
                                                 ByVal nSrcHeight As Long, _
                                                 ByVal dwRop As Long) As Long
                                                 
Private Declare Function SetRect Lib "user32" (lpRect As RECT, _
                                               ByVal x1 As Long, _
                                               ByVal y1 As Long, _
                                               ByVal x2 As Long, _
                                               ByVal y2 As Long) As Long
                                               
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long
                                                                  
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, _
                                                    lprcUpdate As RECT, _
                                                    ByVal hrgnUpdate As Long, _
                                                    ByVal fuRedraw As Long) As Long
                                                    
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDCDest As Long, _
                                                   ByVal nXOriginDest As Long, _
                                                   ByVal nYOriginDest As Long, _
                                                   ByVal nWidthDest As Long, _
                                                   ByVal hHeightDest As Long, _
                                                   ByVal hDCSrc As Long, _
                                                   ByVal nXOriginSrc As Long, _
                                                   ByVal nYOriginSrc As Long, _
                                                   ByVal nWidthSrc As Long, _
                                                   ByVal nHeightSrc As Long, _
                                                   ByVal blendFunc As Long) As Boolean

Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDCDest As Long, _
                                                       ByVal nXOriginDest As Long, _
                                                       ByVal nYOriginDest As Long, _
                                                       ByVal nWidthDest As Long, _
                                                       ByVal hHeightDest As Long, _
                                                       ByVal hDCSrc As Long, _
                                                       ByVal nXOriginSrc As Long, _
                                                       ByVal nYOriginSrc As Long, _
                                                       ByVal nWidthSrc As Long, _
                                                       ByVal nHeightSrc As Long, _
                                                       ByVal crTransparent As Long) As Boolean

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                          ByVal pszClassList As Long) As Long
                                                          
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                ByVal lHDC As Long, _
                                                                ByVal iPartId As Long, _
                                                                ByVal iStateId As Long, _
                                                                pRect As RECT, _
                                                                pClipRect As RECT) As Long
                                                                
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                                      ByVal hdc As Long, _
                                                                      prc As RECT) As Long

Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                          ByVal hdc As Long, _
                                                                          ByVal iPartId As Long, _
                                                                          ByVal iStateId As Long, _
                                                                          pBoundingRect As RECT, _
                                                                          pContentRect As RECT) As Long

Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                          ByVal hdc As Long, _
                                                          ByVal iPartId As Long, _
                                                          ByVal iStateId As Long, _
                                                          ByVal pszText As Long, _
                                                          ByVal iCharCount As Long, _
                                                          ByVal dwTextFlag As Long, _
                                                          ByVal dwTextFlags2 As Long, _
                                                          pRect As RECT) As Long

Private Declare Function DrawThemeIcon Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                          ByVal hdc As Long, _
                                                          ByVal iPartId As Long, _
                                                          ByVal iStateId As Long, _
                                                          pRect As RECT, _
                                                          ByVal himl As Long, _
                                                          ByVal iImageIndex As Long) As Long

Private Declare Function GetThemePartSize Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                             ByVal hdc As Long, _
                                                             ByVal iPartId As Long, _
                                                             ByVal iStateId As Long, _
                                                             prc As RECT, _
                                                             ByVal eSize As THEMESIZE, _
                                                             psz As SIZE) As Long

Private Declare Function GetThemeTextExtent Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                               ByVal hdc As Long, _
                                                               ByVal iPartId As Long, _
                                                               ByVal iStateId As Long, _
                                                               ByVal pszText As Long, _
                                                               ByVal iCharCount As Long, _
                                                               ByVal dwTextFlags As DrawTextFlags, _
                                                               pBoundingRect As RECT, _
                                                               pExtentRect As RECT) As Long

Private Declare Function IsThemePartDefined Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                               ByVal iPartId As Long, _
                                                               ByVal iStateId As Long) As Long

Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" (ByVal himl As Long, _
                                                                    ByVal i As Long, _
                                                                    prcImage As RECT) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long
                                                                               
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, _
                                                                                      ByVal lpFileName As String, _
                                                                                      ByVal nSize As Long) As Long
                                                                                      
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
                                                   ByVal nHeight As Long, _
                                                   ByVal nPlanes As Long, _
                                                   ByVal nBitCount As Long, _
                                                   lpBits As Any) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
                                                   
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long
                                                   
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long, _
                                                   ByVal crColor As Long) As Long
                                                   
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                                                              ByVal lpClassName As String, _
                                                                              ByVal lpWindowName As String, _
                                                                              ByVal dwStyle As Long, _
                                                                              ByVal x As Long, _
                                                                              ByVal y As Long, _
                                                                              ByVal nWidth As Long, _
                                                                              ByVal nHeight As Long, _
                                                                              ByVal hwndParent As Long, _
                                                                              ByVal hMenu As Long, _
                                                                              ByVal hInstance As Long, _
                                                                              lpParam As Any) As Long
                                                                              
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
                                                                        
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
                                                    
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
                                                     
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
                                                qrc As RECT, _
                                                ByVal edge As Long, _
                                                ByVal grfFlags As Long) As Long
                                                
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINT) As Long
                                               
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long
                                             
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                    ByVal hBrush As Long, _
                                                                    ByVal lpDrawStateProc As Long, _
                                                                    ByVal lParam As Long, _
                                                                    ByVal wParam As Long, _
                                                                    ByVal x As Long, _
                                                                    ByVal y As Long, _
                                                                    ByVal cx As Long, _
                                                                    ByVal cy As Long, _
                                                                    ByVal fuFlags As Long) As Long
                                                                    
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
                                                                ByVal dwMaxNameChars As Long, _
                                                                ByVal pszColorBuff As Long, _
                                                                ByVal cchMaxColorChars As Long, _
                                                                ByVal pszSizeBuff As Long, _
                                                                ByVal cchMaxSizeChars As Long) As Long

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByRef lpvParam As Long, _
                                                                                          ByVal fuWinIni As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
                                                
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long
                                                
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
                                                       ByVal yPoint As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Boolean) As Long
                                                    
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, _
                                                        ByVal y1 As Long, _
                                                        ByVal x2 As Long, _
                                                        ByVal y2 As Long) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, _
                                                         ByVal y1 As Long, _
                                                         ByVal x2 As Long, _
                                                         ByVal y2 As Long, _
                                                         ByVal X3 As Long, _
                                                         ByVal Y3 As Long) As Long
                                                         
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINT, _
                                                       ByVal nCount As Long, _
                                                       ByVal nPolyFillMode As Long) As Long
                                                       
Private Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINT, _
                                                           lpPolyCounts As Long, _
                                                           ByVal nCount As Long, _
                                                           ByVal nPolyFillMode As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal x2 As Long, _
                                                    ByVal y2 As Long) As Long
                                                    
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                                                 ByVal hSrcRgn1 As Long, _
                                                 ByVal hSrcRgn2 As Long, _
                                                 ByVal nCombineMode As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, _
                                                     ByVal hWndNewParent As Long) As Long

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, _
                          ByRef bHandled As Boolean, _
                          ByRef lReturn As Long, _
                          ByRef lng_hWnd As Long, _
                          ByRef uMsg As Long, _
                          ByRef wParam As Long, _
                          ByRef lParam As Long)

    Select Case uMsg
    Case WM_MOUSEMOVE
        If Not bInCtrl Then
            bInCtrl = True
            TrackMouseLeave lng_hWnd
            m_iState = stateHot
            Refresh
            RaiseEvent MouseEnter
            CreateToolTip
        End If

    Case WM_MOUSELEAVE
        bInCtrl = False
        m_iState = statenormal
        Refresh
        RaiseEvent MouseLeave

    Case WM_SYSCOLORCHANGE
        Refresh

    Case WM_THEMECHANGED
        Refresh

    End Select

End Sub

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, _
                            ByVal uMsg As Long, _
                            Optional ByVal When As eMsgWhen = MSG_AFTER)

    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            zAddMsg uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub
        End If
        If When And eMsgWhen.MSG_AFTER Then
            zAddMsg uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub
        End If
    End With

End Sub

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, _
                            ByVal uMsg As Long, _
                            Optional ByVal When As eMsgWhen = MSG_AFTER)

    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            zDelMsg uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub
        End If
        If When And eMsgWhen.MSG_AFTER Then
            zDelMsg uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub
        End If
    End With

End Sub

Private Function Subclass_InIDE() As Boolean

    Debug.Assert zSetTrue(Subclass_InIDE)

End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long


Const CODE_LEN             As Long = 200
Const FUNC_CWP             As String = "CallWindowProcA"
Const FUNC_EBM             As String = "EbMode"
Const FUNC_SWL             As String = "SetWindowLongA"
Const MOD_USER             As String = "user32"
Const MOD_VBA5             As String = "vba5"
Const MOD_VBA6             As String = "vba6"
Const PATCH_01             As Long = 18
Const PATCH_02             As Long = 68
Const PATCH_03             As Long = 78
Const PATCH_06             As Long = 116
Const PATCH_07             As Long = 121
Const PATCH_0A             As Long = 186
Static aBuf(1 To CODE_LEN) As Byte
Static pCWP                As Long
Static pEbMode             As Long
Static pSWL                As Long
Dim i                      As Long
Dim j                      As Long
Dim nSubIdx                As Long
Dim sHex                   As String

    If aBuf(1) = 0 Then
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
          "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
          "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
          "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop

        If Subclass_InIDE Then
            aBuf(16) = &H90
            aBuf(17) = &H90
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If pEbMode = 0 Then
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If

        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lng_hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)
        RtlMoveMemory ByVal .nAddrSub, aBuf(1), CODE_LEN
        zPatchRel .nAddrSub, PATCH_01, pEbMode
        zPatchVal .nAddrSub, PATCH_02, .nAddrOrig
        zPatchRel .nAddrSub, PATCH_03, pSWL
        zPatchVal .nAddrSub, PATCH_06, .nAddrOrig
        zPatchRel .nAddrSub, PATCH_07, pCWP
        zPatchVal .nAddrSub, PATCH_0A, ObjPtr(Me)
    End With

End Function

Private Sub Subclass_StopAll()


Dim i As Long


    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If .hwnd <> 0 Then
                Subclass_Stop .hwnd
            End If
        End With
        i = i - 1
    Loop

End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)

    With sc_aSubData(zIdx(lng_hWnd))
        SetWindowLongA .hwnd, GWL_WNDPROC, .nAddrOrig
        zPatchVal .nAddrSub, PATCH_05, 0
        zPatchVal .nAddrSub, PATCH_09, 0
        GlobalFree .nAddrSub
        .hwnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB
        Erase .aMsgTblA
    End With

End Sub

Private Sub zAddMsg(ByVal uMsg As Long, _
                    ByRef aMsgTbl() As Long, _
                    ByRef nMsgCnt As Long, _
                    ByVal When As eMsgWhen, _
                    ByVal nAddr As Long)


Dim nEntry As Long
Dim nOff1  As Long
Dim nOff2  As Long


    If uMsg = ALL_MESSAGES Then
        nMsgCnt = ALL_MESSAGES
    Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = 0 Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf aMsgTbl(nEntry) = uMsg Then
                Exit Sub
            End If
        Loop

        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If

    If When = eMsgWhen.MSG_BEFORE Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
    Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If

    If uMsg <> ALL_MESSAGES Then
        zPatchVal nAddr, nOff1, VarPtr(aMsgTbl(1))
    End If
    zPatchVal nAddr, nOff2, nMsgCnt

End Sub

Private Function zAddrFunc(ByVal sDLL As String, _
                           ByVal sProc As String) As Long

    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc

End Function

Private Sub zDelMsg(ByVal uMsg As Long, _
                    ByRef aMsgTbl() As Long, _
                    ByRef nMsgCnt As Long, _
                    ByVal When As eMsgWhen, _
                    ByVal nAddr As Long)


Dim nEntry As Long


    If uMsg = ALL_MESSAGES Then
        nMsgCnt = 0
        If When = eMsgWhen.MSG_BEFORE Then
            nEntry = PATCH_05
        Else
            nEntry = PATCH_09
        End If
        zPatchVal nAddr, nEntry, 0
    Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then
                aMsgTbl(nEntry) = 0
                Exit Do
            End If
        Loop
    End If

End Sub

Private Function zIdx(ByVal lng_hWnd As Long, _
                      Optional ByVal bAdd As Boolean = False) As Long

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If .hwnd = lng_hWnd Then
                If Not bAdd Then
                    Exit Function
                End If
            ElseIf .hwnd = 0 Then
                If bAdd Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop

    If Not bAdd Then
        Debug.Assert False
    End If

End Function

Private Sub zPatchRel(ByVal nAddr As Long, _
                      ByVal nOffset As Long, _
                      ByVal nTargetAddr As Long)

    RtlMoveMemory ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4

End Sub

Private Sub zPatchVal(ByVal nAddr As Long, _
                      ByVal nOffset As Long, _
                      ByVal nValue As Long)

    RtlMoveMemory ByVal nAddr + nOffset, nValue, 4

End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean

    zSetTrue = True
    bValue = True

End Function

Private Sub APILine(x1 As Long, _
                    y1 As Long, _
                    x2 As Long, _
                    y2 As Long, _
                    lcolor As Long)


Dim pt      As POINT
Dim hPen    As Long
Dim hPenOld As Long


    hPen = CreatePen(0, 1, lcolor)
    With UserControl
        hPenOld = SelectObject(.hdc, hPen)
        MoveToEx .hdc, x1, y1, pt
        LineTo .hdc, x2, y2
        SelectObject .hdc, hPenOld
    End With
    DeleteObject hPen

End Sub

Private Sub APILineEx(lhdcEx As Long, _
                      x1 As Long, _
                      y1 As Long, _
                      x2 As Long, _
                      y2 As Long, _
                      lcolor As Long)


Dim pt      As POINT
Dim hPen    As Long
Dim hPenOld As Long


    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, x1, y1, pt
    LineTo lhdcEx, x2, y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen

End Sub

Private Sub APIFillRect(lngHdc As Long, _
                        rc As RECT, _
                        Color As Long)


Dim NewBrush As Long


    NewBrush = CreateSolidBrush(Color)
    FillRect lngHdc, rc, NewBrush
    DeleteObject NewBrush

End Sub

Private Sub APIFillRectByCoords(lngHdc As Long, _
                                ByVal x As Long, _
                                ByVal y As Long, _
                                ByVal w As Long, _
                                ByVal h As Long, _
                                Color As Long)


Dim NewBrush As Long
Dim tmprect  As RECT


    NewBrush = CreateSolidBrush(Color)
    SetRect tmprect, x, y, x + w, y + h
    FillRect lngHdc, tmprect, NewBrush
    DeleteObject NewBrush

End Sub

Private Function APIRectangle(ByVal lngHdc As Long, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal w As Long, _
                              ByVal h As Long, _
                              Optional lcolor As OLE_COLOR = -1) As Long


Dim hPen    As Long
Dim hPenOld As Long
Dim pt      As POINT


    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(lngHdc, hPen)
    MoveToEx lngHdc, x, y, pt
    LineTo lngHdc, x + w, y
    LineTo lngHdc, x + w, y + h
    LineTo lngHdc, x, y + h
    LineTo lngHdc, x, y
    SelectObject lngHdc, hPenOld
    DeleteObject hPen

End Function

Private Sub DrawCtlEdgeByRect(lngHdc As Long, _
                              rt As RECT, _
                              Optional Style As Long = EDGE_RAISED, _
                              Optional Flags As Long = BF_RECT)

    DrawEdge lngHdc, rt, Style, Flags

End Sub

Private Sub DrawCtlEdge(lngHdc As Long, _
                        ByVal x As Single, _
                        ByVal y As Single, _
                        ByVal w As Single, _
                        ByVal h As Single, _
                        Optional Style As Long = EDGE_RAISED, _
                        Optional ByVal Flags As Long = BF_RECT)


Dim R As RECT


    With R
        .Left = x
        .Top = y
        .Right = x + w
        .Bottom = y + h
    End With
    DrawEdge lngHdc, R, Style, Flags

End Sub

Private Function BlendColors(ByVal lcolor1 As Long, _
                             ByVal lcolor2 As Long) As Single

    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)

End Function

Private Function TranslateColor(ByVal lcolor As Long) As Long

    If OleTranslateColor(lcolor, 0, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Private Function SoftColor(lcolor As OLE_COLOR) As OLE_COLOR


Dim lRed   As OLE_COLOR
Dim lGreen As OLE_COLOR
Dim lBlue  As OLE_COLOR
Dim lR     As OLE_COLOR
Dim lg     As OLE_COLOR
Dim lb     As OLE_COLOR

    lR = (lcolor And &HFF)
    lg = ((lcolor And 65280) \ 256)
    lb = ((lcolor) And 16711680) \ 65536
    lRed = (76 - Int(((lcolor And &HFF) + 32) \ 64) * 19)
    lGreen = (76 - Int((((lcolor And 65280) \ 256) + 32) \ 64) * 19)
    lBlue = (76 - Int((((lcolor And &HFF0000) \ &H10000) + 32) / 64) * 19)
    SoftColor = RGB(lR + lRed, lg + lGreen, lb + lBlue)

End Function

Private Function MSOXPShiftColor(ByVal theColor As Long, _
                                 Optional ByVal Base As Long = &HB0) As Long


Dim Red   As Long
Dim Blue  As Long
Dim Green As Long
Dim Delta As Long


    Blue = ((theColor \ &H10000) Mod &H100)
    Green = ((theColor \ &H100) Mod &H100)
    Red = (theColor And &HFF)
    Delta = &HFF - Base

    Blue = Base + Blue * Delta \ &HFF
    Green = Base + Green * Delta \ &HFF
    Red = Base + Red * Delta \ &HFF

    If Red > 255 Then
        Red = 255
    End If
    If Green > 255 Then
        Green = 255
    End If
    If Blue > 255 Then
        Blue = 255
    End If

    MSOXPShiftColor = Red + 256& * Green + 65536 * Blue

End Function

Private Function msSoftColor(ByVal lcolor As Long) As Long

Dim lRed   As Long
Dim lGreen As Long
Dim lBlue  As Long
Dim lR     As Long
Dim lg     As Long
Dim lb     As Long

    lR = (lcolor And &HFF)
    lg = ((lcolor And 65280) \ 256)
    lb = ((lcolor) And 16711680) \ 65536
    lRed = (76 - Int(((lcolor And &HFF) + 32) \ 64) * 19)
    lGreen = (76 - Int((((lcolor And 65280) \ 256) + 32) \ 64) * 19)
    lBlue = (76 - Int((((lcolor And &HFF0000) \ &H10000) + 32) / 64) * 19)
    msSoftColor = RGB(lR + lRed, lg + lGreen, lb + lBlue)

End Function

Private Function OffsetColor(lcolor As OLE_COLOR, _
                             ByVal lOffset As Long) As OLE_COLOR


Dim lRed   As OLE_COLOR
Dim lGreen As OLE_COLOR
Dim lBlue  As OLE_COLOR
Dim lR     As OLE_COLOR
Dim lg     As OLE_COLOR
Dim lb     As OLE_COLOR

    lR = (lcolor And &HFF)
    lg = ((lcolor And 65280) \ 256)
    lb = ((lcolor) And 16711680) \ 65536
    lRed = (lOffset + lR)
    lGreen = (lOffset + lg)
    lBlue = (lOffset + lb)
    If lRed > 255 Then
        lRed = 255
    End If
    If lRed < 0 Then
        lRed = 0
    End If
    If lGreen > 255 Then
        lGreen = 255
    End If
    If lGreen < 0 Then
        lGreen = 0
    End If
    If lBlue > 255 Then
        lBlue = 255
    End If
    If lBlue < 0 Then
        lBlue = 0
    End If
    OffsetColor = RGB(lRed, lGreen, lBlue)

End Function

Private Sub DrawCaption()


Dim lcolor    As Long
Dim ltmpColor As Long


    If Not m_bUseCustomColors Then
        If m_iState <> statedisabled Then
            lcolor = GetSysColor(COLOR_BTNTEXT)
        Else
            lcolor = TranslateColor(vbGrayText)
        End If
    Else
        Select Case m_iState
        Case statenormal
            lcolor = m_lFontColor
        Case statedisabled
            lcolor = TranslateColor(vbGrayText)
        Case Else
            lcolor = m_lFontHighlightColor
        End Select
    End If

    With UserControl
        ltmpColor = .ForeColor
        .ForeColor = lcolor
        DrawText .hdc, m_sCaption, -1, m_txtRect, lwFontAlign
        .ForeColor = ltmpColor
    End With

End Sub

Private Sub fPaintPicture(ByRef m_Picture As StdPicture, _
                          ByVal x As Long, _
                          ByVal y As Long, _
                          ByVal w As Long, _
                          ByVal h As Long)



Dim memDC As Long
Dim oldW  As Long
Dim oldH  As Long

    oldW = m_Picture.Width
    oldH = m_Picture.Height
    memDC = CreateCompatibleDC(UserControl.hdc)
    SelectObject memDC, m_Picture.handle
    StretchBlt UserControl.hdc, x, y, w, h, memDC, 0, 0, oldW, oldH, vbSrcCopy
    BitBlt UserControl.hdc, x, y, w, h, memDC, 0, 0, vbSrcCopy

End Sub

Private Sub fDrawPicture(ByRef m_Picture As StdPicture, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal bShadow As Boolean)


Dim lFlags As Long
Dim hBrush As Long


    Select Case m_Picture.Type
    Case vbPicTypeBitmap
        lFlags = DST_BITMAP
    Case vbPicTypeIcon
        lFlags = DST_ICON
    Case Else
        lFlags = DST_COMPLEX
    End Select
    If bShadow Then
        hBrush = CreateSolidBrush(RGB(128, 128, 128))
    End If
    DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.handle, 0, x, y, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
    If bShadow Then
        DeleteObject hBrush
    End If

End Sub

Private Sub DrawVGradient(ByVal lEndColor As Long, _
                          ByVal lStartcolor As Long, _
                          ByVal x As Long, _
                          ByVal y As Long, _
                          ByVal x2 As Long, _
                          ByVal y2 As Long)


Dim dR As Single
Dim dG As Single
Dim dB As Single
Dim sR As Single
Dim sG As Single
Dim sB As Single
Dim eR As Single
Dim eG As Single
Dim eB As Single
Dim ni As Long

    'lh = UserControl.ScaleHeight
    'lw = UserControl.ScaleWidth
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / y2
    dG = (sG - eG) / y2
    dB = (sB - eB) / y2
    For ni = 0 To y2
        APILine x, y + ni, x2, y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni

End Sub

Private Sub DrawVGradientEx(lhdcEx As Long, _
                            ByVal lEndColor As Long, _
                            ByVal lStartcolor As Long, _
                            ByVal x As Long, _
                            ByVal y As Long, _
                            ByVal x2 As Long, _
                            ByVal y2 As Long)

Dim dR As Single
Dim dG As Single
Dim dB As Single
Dim sR As Single
Dim sG As Single
Dim sB As Single
Dim eR As Single
Dim eG As Single
Dim eB As Single
Dim ni As Long

    'lh = UserControl.ScaleHeight
    'lw = UserControl.ScaleWidth
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / y2
    dG = (sG - eG) / y2
    dB = (sB - eB) / y2
    For ni = 0 To y2
        APILineEx lhdcEx, x, y + ni, x2, y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni

End Sub

Private Sub DrawHGradient(ByVal lEndColor As Long, _
                          ByVal lStartcolor As Long, _
                          ByVal x As Long, _
                          ByVal y As Long, _
                          ByVal x2 As Long, _
                          ByVal y2 As Long)


Dim dR As Single
Dim dG As Single
Dim dB As Single
Dim sR As Single
Dim sG As Single
Dim sB As Single
Dim eR As Single
Dim eG As Single
Dim eB As Single
Dim lw As Long
Dim ni As Long

    lw = x2 - x
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / lw
    dG = (sG - eG) / lw
    dB = (sB - eB) / lw
    For ni = 0 To lw
        APILine x + ni, y, x + ni, y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni

End Sub

Private Sub DrawJavaBorder(ByVal x As Long, _
                           ByVal y As Long, _
                           ByVal w As Long, _
                           ByVal h As Long, _
                           ByVal lColorShadow As Long, _
                           ByVal lColorLight As Long, _
                           ByVal lColorBack As Long)

    With UserControl
        APIRectangle .hdc, x, y, w - 1, h - 1, lColorShadow
        APIRectangle .hdc, x + 1, y + 1, w - 1, h - 1, lColorLight
        SetPixel .hdc, x, y + h, lColorBack
        SetPixel .hdc, x + w, y, lColorBack
        SetPixel .hdc, x + 1, y + h - 1, BlendColors(lColorLight, lColorShadow)
        SetPixel .hdc, x + w - 1, y + 1, BlendColors(lColorLight, lColorShadow)

    End With

End Sub

Private Function DrawTheme(sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long) As Boolean


Dim hTheme  As Long
Dim lResult As Long


    On Error GoTo NoXP

    hTheme = OpenThemeData(UserControl.hwnd, StrPtr(sClass))
    If hTheme Then
        lResult = DrawThemeBackground(hTheme, UserControl.hdc, iPart, iState, m_btnRect, m_btnRect)
        DrawTheme = IIf(lResult, False, True)
    Else
        DrawTheme = False
    End If


Exit Function

NoXP:
    DrawTheme = False

End Function

Private Function CreateWinXPRegion() As Long


Dim pPoligon(8) As POINT
Dim lw          As Long
Dim lh          As Long


    lw = UserControl.ScaleWidth
    lh = UserControl.ScaleHeight
    pPoligon(0).x = 0
    pPoligon(0).y = 1
    pPoligon(1).x = 1
    pPoligon(1).y = 0
    pPoligon(2).x = lw - 1
    pPoligon(2).y = 0
    pPoligon(3).x = lw
    pPoligon(3).y = 1
    pPoligon(4).x = lw
    pPoligon(4).y = lh - 2
    pPoligon(5).x = lw - 2
    pPoligon(5).y = lh
    pPoligon(6).x = 2
    pPoligon(6).y = lh
    pPoligon(7).x = 0
    pPoligon(7).y = lh - 2
    CreateWinXPRegion = CreatePolygonRgn(pPoligon(0), 8, ALTERNATE)

End Function

Private Function CreateGalaxyRegion() As Long


Dim pPoligon(8) As POINT
Dim lw          As Long
Dim lh          As Long


    lw = UserControl.ScaleWidth
    lh = UserControl.ScaleHeight
    pPoligon(0).x = 0
    pPoligon(0).y = 2
    pPoligon(1).x = 2
    pPoligon(1).y = 0
    pPoligon(2).x = lw - 3
    pPoligon(2).y = 0
    pPoligon(3).x = lw
    pPoligon(3).y = 3
    pPoligon(4).x = lw
    pPoligon(4).y = lh - 3
    pPoligon(5).x = lw - 3
    pPoligon(5).y = lh
    pPoligon(6).x = 4
    pPoligon(6).y = lh
    pPoligon(7).x = 0
    pPoligon(7).y = lh - 4
    CreateGalaxyRegion = CreatePolygonRgn(pPoligon(0), 8, ALTERNATE)

End Function

Private Function CreateMacOSXButtonRegion() As Long

    CreateMacOSXButtonRegion = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1, 18, 18)

End Function

Private Sub DrawWinXPButton(Mode As isEstate)


Dim lHDC      As Long
Dim tempColor As Long
Dim lh        As Long
Dim lw        As Long


    lw = UserControl.ScaleWidth
    lh = UserControl.ScaleHeight
    lHDC = UserControl.hdc
    UserControl.BackColor = GetSysColor(COLOR_BTNFACE)

    Select Case Mode
    Case statenormal, stateHot

        DrawVGradient &HFBFCFC, &HF0F0F0, 1, 1, lw - 2, 4
        DrawVGradient &HF9FAFA, &HEAF0F0, 1, 4, lw - 2, lh - 8
        DrawVGradient &HE6EBEB, &HC5D0D6, 1, lh - 4, lw - 2, 3
        DrawVGradient &HFAFBFB, &HDAE2E4, lw - 3, 3, lw - 2, lh - 5
        DrawVGradient &HF2F4F5, &HCDD7DB, lw - 2, 3, lw - 1, lh - 5
        APILine 1, 0, lw - 1, 0, &H743C00
        APILine 0, 1, 0, lh - 1, &H743C00
        APILine lw - 1, 1, lw - 1, lh - 1, &H743C00
        APILine 1, lh - 1, lw - 1, lh - 1, &H743C00
        SetPixel lHDC, 1, 1, &H906E48
        SetPixel lHDC, 1, lh - 2, &H906E48
        SetPixel lHDC, lw - 2, 1, &H906E48
        SetPixel lHDC, lw - 2, lh - 2, &H906E48
        SetPixel lHDC, 0, 1, &HA28B6A
        SetPixel lHDC, 1, 0, &HA28B6A
        SetPixel lHDC, 1, lh - 1, &HA28B6A
        SetPixel lHDC, 0, lh - 2, &HA28B6A
        SetPixel lHDC, lw - 1, lh - 2, &HA28B6A
        SetPixel lHDC, lw - 2, lh - 1, &HA28B6A
        SetPixel lHDC, lw - 2, 0, &HA28B6A
        SetPixel lHDC, lw - 1, 1, &HA28B6A
        SetPixel lHDC, 1, 2, &HCAC7BF
        SetPixel lHDC, 2, 1, &HCAC7BF
        SetPixel lHDC, 2, lh - 2, &HCAC7BF
        SetPixel lHDC, 1, lh - 3, &HCAC7BF
        SetPixel lHDC, lw - 2, lh - 3, &HCAC7BF
        SetPixel lHDC, lw - 3, lh - 2, &HCAC7BF
        SetPixel lHDC, lw - 3, 1, &HCAC7BF
        SetPixel lHDC, lw - 2, 2, &HCAC7BF
        If Mode = stateHot Then
            APILine 2, 1, lw - 2, 1, &HCFF0FF
            APILine 2, 2, lw - 2, 2, &H89D8FD
            APILine 2, lh - 3, lw - 2, lh - 3, &H30B3F8
            APILine 2, lh - 2, lw - 2, lh - 2, &H1097E5
            DrawVGradient &H89D8FD, &H30B3F8, 1, 2, 3, lh - 5
            DrawVGradient &H89D8FD, &H30B3F8, lw - 3, 2, lw - 1, lh - 5
        ElseIf Mode = statenormal And m_bFocused Then
            APILine 2, lh - 2, lw - 2, lh - 2, &HEE8269
            APILine 2, 1, lw - 2, 1, &HFFE7CE
            APILine 2, 2, lw - 2, 2, &HF6D4BC
            APILine 2, lh - 3, lw - 2, lh - 3, &HE4AD89
            DrawVGradient &HF6D4BC, &HE4AD89, 1, 2, 3, lh - 5
            DrawVGradient &HF6D4BC, &HE4AD89, lw - 3, 2, lw - 1, lh - 5
        End If

    Case statePressed
        DrawVGradient &HC1CCD1, &HDCE3E4, 2, 1, lw - 1, 4
        DrawVGradient &HDCE3E4, &HDBE2E3, 2, 4, lw - 1, lh - 8
        DrawVGradient &HDBE2E3, &HEEF1F2, 3, lh - 4, lw - 1, 3
        DrawVGradient &HCED8DA, &HDBE2E3, 1, 3, 2, lh - 5
        DrawVGradient &HCED8DA, &HDBE2E3, 2, 4, 3, lh - 7
        APILine 1, 0, lw - 1, 0, &H743C00
        APILine 0, 1, 0, lh - 1, &H743C00
        APILine lw - 1, 1, lw - 1, lh - 1, &H743C00
        APILine 1, lh - 1, lw - 1, lh - 1, &H743C00
        SetPixel lHDC, 1, 1, &H906E48
        SetPixel lHDC, 1, lh - 2, &H906E48
        SetPixel lHDC, lw - 2, 1, &H906E48
        SetPixel lHDC, lw - 2, lh - 2, &H906E48
        SetPixel lHDC, 0, 1, &HA28B6A
        SetPixel lHDC, 1, 0, &HA28B6A
        SetPixel lHDC, 1, lh - 1, &HA28B6A
        SetPixel lHDC, 0, lh - 2, &HA28B6A
        SetPixel lHDC, lw - 1, lh - 2, &HA28B6A
        SetPixel lHDC, lw - 2, lh - 1, &HA28B6A
        SetPixel lHDC, lw - 2, 0, &HA28B6A
        SetPixel lHDC, lw - 1, 1, &HA28B6A
        SetPixel lHDC, 1, 2, &HCAC7BF
        SetPixel lHDC, 2, 1, &HCAC7BF
        SetPixel lHDC, 2, lh - 2, &HCAC7BF
        SetPixel lHDC, 1, lh - 3, &HCAC7BF
        SetPixel lHDC, lw - 2, lh - 3, &HCAC7BF
        SetPixel lHDC, lw - 3, lh - 2, &HCAC7BF
        SetPixel lHDC, lw - 3, 1, &HCAC7BF
        SetPixel lHDC, lw - 2, 2, &HCAC7BF

    Case statedisabled
        tempColor = &HEAF4F5
        UserControl.BackColor = tempColor
        lHDC = UserControl.hdc
        APIRectangle lHDC, 0, 0, lw - 1, lh - 1, &HBAC7C9
        tempColor = &HC7D5D8
        SetPixel lHDC, 0, 1, tempColor
        SetPixel lHDC, 1, 1, tempColor
        SetPixel lHDC, 1, 0, tempColor
        SetPixel lHDC, 0, lh - 2, tempColor
        SetPixel lHDC, 1, lh - 2, tempColor
        SetPixel lHDC, 1, lh - 1, tempColor
        SetPixel lHDC, lw - 1, 1, tempColor
        SetPixel lHDC, lw - 2, 1, tempColor
        SetPixel lHDC, lw - 2, 0, tempColor
        SetPixel lHDC, lw - 1, lh - 2, tempColor
        SetPixel lHDC, lw - 2, lh - 2, tempColor
        SetPixel lHDC, lw - 2, lh - 1, tempColor
    End Select

End Sub

Private Sub DrawCustomWinXPButton(Mode As isEstate)


Dim tmpcolor As Long
Dim lh       As Long
Dim lw       As Long
Dim lHDC     As Long


    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth

    Select Case Mode
    Case statenormal, stateDefaulted, stateHot
        tmpcolor = m_lBackColor
        With UserControl
            .BackColor = tmpcolor
            DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&HF), 1, 1, .ScaleWidth - 2, .ScaleHeight - 3
            DrawVGradient OffsetColor(tmpcolor, &H15), OffsetColor(tmpcolor, -&H15), 1, 2, 2, .ScaleHeight - 5
            DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H20), .ScaleWidth - 2, 2, .ScaleWidth - 3, .ScaleHeight - 5
            APILine 2, 1, .ScaleWidth - 2, 1, OffsetColor(tmpcolor, &H5)
            APILine 1, 2, .ScaleWidth - 1, 2, OffsetColor(tmpcolor, &H2)
            APILine 2, .ScaleHeight - 4, .ScaleWidth - 2, .ScaleHeight - 4, OffsetColor(tmpcolor, -&H10)
            APILine 2, .ScaleHeight - 3, .ScaleWidth - 2, .ScaleHeight - 3, OffsetColor(tmpcolor, -&H18)
            APILine 2, .ScaleHeight - 2, .ScaleWidth - 2, .ScaleHeight - 2, OffsetColor(tmpcolor, -&H25)
        End With

        tmpcolor = OffsetColor(tmpcolor, -&H80)

        With UserControl
            APILine 2, 0, .ScaleWidth - 2, 0, tmpcolor
            APILine 0, 2, 0, .ScaleHeight - 2, tmpcolor
            APILine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, tmpcolor
            APILine .ScaleWidth - 1, 0, .ScaleWidth - 1, .ScaleHeight, tmpcolor
            SetPixel .hdc, 1, 1, tmpcolor
            SetPixel .hdc, 1, .ScaleHeight - 2, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, .ScaleHeight - 2, tmpcolor
        End With

        tmpcolor = OffsetColor(m_lBackColor, -&H15)

        With UserControl
            SetPixel .hdc, 1, 0, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, 0, tmpcolor
            SetPixel .hdc, 0, 1, tmpcolor
            SetPixel .hdc, 0, .ScaleHeight - 2, tmpcolor
            SetPixel .hdc, 1, .ScaleHeight - 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, .ScaleHeight - 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 1, 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 1, .ScaleHeight - 2, tmpcolor
        End With

        If Mode = stateDefaulted Or Mode = stateHot Or (m_bFocused And m_bShowFocus) Then
            tmpcolor = IIf((Mode = stateHot), m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
            APILine 2, 1, lw - 2, 1, OffsetColor(tmpcolor, &H55)
            APILine 2, 2, lw - 2, 2, OffsetColor(tmpcolor, &H45)
            APILine 2, lh - 3, lw - 2, lh - 3, OffsetColor(tmpcolor, &H10)
            APILine 2, lh - 2, lw - 2, lh - 2, tmpcolor
            DrawVGradient OffsetColor(tmpcolor, &H45), OffsetColor(tmpcolor, &H10), 1, 2, 3, lh - 5
            DrawVGradient OffsetColor(tmpcolor, &H45), OffsetColor(tmpcolor, &H10), lw - 3, 2, lw - 1, lh - 5
        End If

    Case statePressed
        tmpcolor = m_lBackColor
        lHDC = UserControl.hdc
        DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H15), 2, 1, lw - 1, 4
        DrawVGradient OffsetColor(tmpcolor, -&H15), OffsetColor(tmpcolor, -&H5), 2, 4, lw - 1, lh - 8
        DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, &H5), 3, lh - 4, lw - 1, 3
        DrawVGradient OffsetColor(tmpcolor, -&H20), OffsetColor(tmpcolor, -&H16), 1, 3, 2, lh - 5
        DrawVGradient OffsetColor(tmpcolor, -&H18), OffsetColor(tmpcolor, -&HF), 2, 4, 3, lh - 7
        SetPixel lHDC, 1, 2, OffsetColor(tmpcolor, -&H30)
        SetPixel lHDC, 2, 1, OffsetColor(tmpcolor, -&H30)
        SetPixel lHDC, 2, lh - 2, OffsetColor(tmpcolor, -&H5)
        SetPixel lHDC, 1, lh - 3, OffsetColor(tmpcolor, -&H10)
        SetPixel lHDC, lw - 2, lh - 3, OffsetColor(tmpcolor, &H12)
        SetPixel lHDC, lw - 3, lh - 2, OffsetColor(tmpcolor, &H8)
        SetPixel lHDC, lw - 3, 1, OffsetColor(tmpcolor, -&H30)
        SetPixel lHDC, lw - 2, 2, OffsetColor(tmpcolor, -&H25)
        tmpcolor = OffsetColor(m_lBackColor, -&H80)
        APILine 1, 0, lw - 1, 0, tmpcolor
        APILine 0, 1, 0, lh - 1, tmpcolor
        APILine lw - 1, 1, lw - 1, lh - 1, tmpcolor
        APILine 1, lh - 1, lw - 1, lh - 1, tmpcolor
        tmpcolor = OffsetColor(m_lBackColor, -&H60)
        SetPixel lHDC, 1, 1, tmpcolor
        SetPixel lHDC, 1, lh - 2, tmpcolor
        SetPixel lHDC, lw - 2, 1, tmpcolor
        SetPixel lHDC, lw - 2, lh - 2, tmpcolor

    Case statedisabled
        tmpcolor = m_lBackColor
        UserControl.BackColor = tmpcolor
        lHDC = UserControl.hdc
        APIRectangle lHDC, 0, 0, lw - 1, lh - 1, OffsetColor(m_lBackColor, -&H40)
        tmpcolor = OffsetColor(m_lBackColor, -&H35)
        SetPixel lHDC, 0, 1, tmpcolor
        SetPixel lHDC, 1, 1, tmpcolor
        SetPixel lHDC, 1, 0, tmpcolor
        SetPixel lHDC, 0, lh - 2, tmpcolor
        SetPixel lHDC, 1, lh - 2, tmpcolor
        SetPixel lHDC, 1, lh - 1, tmpcolor
        SetPixel lHDC, lw - 1, 1, tmpcolor
        SetPixel lHDC, lw - 2, 1, tmpcolor
        SetPixel lHDC, lw - 2, 0, tmpcolor
        SetPixel lHDC, lw - 1, lh - 2, tmpcolor
        SetPixel lHDC, lw - 2, lh - 2, tmpcolor
        SetPixel lHDC, lw - 2, lh - 1, tmpcolor
    End Select

End Sub

Private Sub DrawMacOSXButton()

    If m_iState = stateHot Or m_iState = stateDefaulted Then
        DrawMacOSXButtonHot
    ElseIf m_iState = statenormal Or m_iState = statedisabled Then
        If m_bFocused Then
            DrawMacOSXButtonHot
        Else
            DrawMacOSXButtonNormal
        End If
    Else
        DrawMacOSXButtonPressed
    End If

End Sub

Private Sub DrawMacOSXButtonNormal()


Dim lHDC  As Long
Dim lh    As Long
Dim lw    As Long
Dim tmph  As Long
Dim tmpw  As Long
Dim tmph1 As Long


    With UserControl
        lHDC = .hdc
        lh = .ScaleHeight
        lw = .ScaleWidth
    End With
    APIFillRectByCoords hdc, 18, 11, lw - 34, lh - 19, &HEAE7E8
    SetPixel lHDC, 6, 0, &HFEFEFE
    SetPixel lHDC, 7, 0, &HE6E6E6
    SetPixel lHDC, 8, 0, &HACACAC
    SetPixel lHDC, 9, 0, &H7A7A7A
    SetPixel lHDC, 10, 0, &H6C6C6C
    SetPixel lHDC, 11, 0, &H6B6B6B
    SetPixel lHDC, 12, 0, &H6F6F6F
    SetPixel lHDC, 13, 0, &H716F6F
    SetPixel lHDC, 14, 0, &H727070
    SetPixel lHDC, 15, 0, &H676866
    SetPixel lHDC, 16, 0, &H6C6D6B
    SetPixel lHDC, 17, 0, &H67696A
    SetPixel lHDC, 5, 1, &HEFEFEF
    SetPixel lHDC, 6, 1, &H939393
    SetPixel lHDC, 7, 1, &H676767
    SetPixel lHDC, 8, 1, &H797979
    SetPixel lHDC, 9, 1, &HB3B3B3
    SetPixel lHDC, 10, 1, &HDBDBDB
    SetPixel lHDC, 11, 1, &HEBEDEE
    SetPixel lHDC, 12, 1, &HF5F4F6
    SetPixel lHDC, 13, 1, &HF5F4F6
    SetPixel lHDC, 14, 1, &HF5F4F6
    SetPixel lHDC, 15, 1, &HF5F4F6
    SetPixel lHDC, 16, 1, &HF5F4F6
    SetPixel lHDC, 17, 1, &HF5F4F6
    SetPixel lHDC, 3, 2, &HFEFEFE
    SetPixel lHDC, 4, 2, &HE5E5E5
    SetPixel lHDC, 5, 2, &H737373
    SetPixel lHDC, 6, 2, &H656565
    SetPixel lHDC, 7, 2, &H939393
    SetPixel lHDC, 8, 2, &HDCDCDC
    SetPixel lHDC, 9, 2, &HE9E9E9
    SetPixel lHDC, 10, 2, &HF2F1F3
    SetPixel lHDC, 11, 2, &HF3F2F4
    SetPixel lHDC, 12, 2, &HF2F1F3
    SetPixel lHDC, 13, 2, &HF3F2F4
    SetPixel lHDC, 14, 2, &HF2F1F3
    SetPixel lHDC, 15, 2, &HF3F2F4
    SetPixel lHDC, 16, 2, &HF2F1F3
    SetPixel lHDC, 17, 2, &HF3F2F4
    SetPixel lHDC, 3, 3, &HEEEEEE
    SetPixel lHDC, 4, 3, &H717171
    SetPixel lHDC, 5, 3, &H6C6C6C
    SetPixel lHDC, 6, 3, &H909090
    SetPixel lHDC, 7, 3, &HD2D2D2
    SetPixel lHDC, 8, 3, &HE3E3E3
    SetPixel lHDC, 9, 3, &HECECEC
    SetPixel lHDC, 10, 3, &HEDEDED
    SetPixel lHDC, 11, 3, &HEEEEEE
    SetPixel lHDC, 12, 3, &HEDEDED
    SetPixel lHDC, 13, 3, &HEEEEEE
    SetPixel lHDC, 14, 3, &HEDEDED
    SetPixel lHDC, 15, 3, &HEEEEEE
    SetPixel lHDC, 16, 3, &HEDEDED
    SetPixel lHDC, 17, 3, &HEEEEEE
    SetPixel lHDC, 2, 4, &HFBFBFB
    SetPixel lHDC, 3, 4, &H858585
    SetPixel lHDC, 4, 4, &H686868
    SetPixel lHDC, 5, 4, &H959595
    SetPixel lHDC, 6, 4, &HB1B1B1
    SetPixel lHDC, 7, 4, &HDCDCDC
    SetPixel lHDC, 8, 4, &HE3E3E3
    SetPixel lHDC, 9, 4, &HE3E3E3
    SetPixel lHDC, 10, 4, &HEAEAEA
    SetPixel lHDC, 11, 4, &HEBEBEB
    SetPixel lHDC, 12, 4, &HEBEBEB
    SetPixel lHDC, 13, 4, &HEBEBEB
    SetPixel lHDC, 14, 4, &HEBEBEB
    SetPixel lHDC, 15, 4, &HEBEBEB
    SetPixel lHDC, 16, 4, &HEBEBEB
    SetPixel lHDC, 17, 4, &HEBEBEB
    SetPixel lHDC, 1, 5, &HFEFEFE
    SetPixel lHDC, 2, 5, &HCACACA
    SetPixel lHDC, 3, 5, &H696969
    SetPixel lHDC, 4, 5, &H949494
    SetPixel lHDC, 5, 5, &HA6A6A6
    SetPixel lHDC, 6, 5, &HC5C5C5
    SetPixel lHDC, 7, 5, &HD8D8D8
    SetPixel lHDC, 8, 5, &HE0E0E0
    SetPixel lHDC, 9, 5, &HE1E1E1
    SetPixel lHDC, 10, 5, &HEAE9EA
    SetPixel lHDC, 11, 5, &HE7E7E7
    SetPixel lHDC, 12, 5, &HE9E7E8
    SetPixel lHDC, 13, 5, &HEBE8EA
    SetPixel lHDC, 14, 5, &HEAE7E9
    SetPixel lHDC, 15, 5, &HEBE8EA
    SetPixel lHDC, 16, 5, &HEAE7E9
    SetPixel lHDC, 17, 5, &HEBE8EA
    SetPixel lHDC, 1, 6, &HF9F9F9
    SetPixel lHDC, 2, 6, &H808080
    SetPixel lHDC, 3, 6, &H878787
    SetPixel lHDC, 4, 6, &HA8A8A8
    SetPixel lHDC, 5, 6, &HB3B3B3
    SetPixel lHDC, 6, 6, &HC6C6C6
    SetPixel lHDC, 7, 6, &HDEDEDE
    SetPixel lHDC, 8, 6, &HE0E0E0
    SetPixel lHDC, 9, 6, &HE2E2E2
    SetPixel lHDC, 10, 6, &HE3E2E2
    SetPixel lHDC, 11, 6, &HE9EAE9
    SetPixel lHDC, 12, 6, &HE9E8E9
    SetPixel lHDC, 13, 6, &HEBE8EA
    SetPixel lHDC, 14, 6, &HEBE8EA
    SetPixel lHDC, 15, 6, &HEBE8EA
    SetPixel lHDC, 16, 6, &HEBE8EA
    SetPixel lHDC, 17, 6, &HEBE8EA
    SetPixel lHDC, 1, 7, &HE8E8E8
    SetPixel lHDC, 2, 7, &H777777
    SetPixel lHDC, 3, 7, &H9B9B9B
    SetPixel lHDC, 4, 7, &HB1B1B1
    SetPixel lHDC, 5, 7, &HB9B9B9
    SetPixel lHDC, 6, 7, &HC5C5C5
    SetPixel lHDC, 7, 7, &HD6D6D6
    SetPixel lHDC, 8, 7, &HE0E0E0
    SetPixel lHDC, 9, 7, &HE0E0E0
    SetPixel lHDC, 10, 7, &HE7E7E7
    SetPixel lHDC, 11, 7, &HE7E7E7
    SetPixel lHDC, 12, 7, &HE9E9E9
    SetPixel lHDC, 13, 7, &HEAEAEA
    SetPixel lHDC, 14, 7, &HEAEAEA
    SetPixel lHDC, 15, 7, &HEAEAEA
    SetPixel lHDC, 16, 7, &HEAEAEA
    SetPixel lHDC, 17, 7, &HEAEAEA
    SetPixel lHDC, 0, 8, &HFDFDFD
    SetPixel lHDC, 1, 8, &HC6C6C6
    SetPixel lHDC, 2, 8, &H7E7E7E
    SetPixel lHDC, 3, 8, &HABABAB
    SetPixel lHDC, 4, 8, &HC1C1C1
    SetPixel lHDC, 5, 8, &HC1C1C1
    SetPixel lHDC, 6, 8, &HCBCBCB
    SetPixel lHDC, 7, 8, &HCECECE
    SetPixel lHDC, 8, 8, &HD5D5D5
    SetPixel lHDC, 9, 8, &HD8D8D8
    SetPixel lHDC, 10, 8, &HDADADA
    SetPixel lHDC, 11, 8, &HDDDDDD
    SetPixel lHDC, 12, 8, &HDEDEDE
    SetPixel lHDC, 13, 8, &HE1E1E1
    SetPixel lHDC, 14, 8, &HE0E0E0
    SetPixel lHDC, 15, 8, &HE1E1E1
    SetPixel lHDC, 16, 8, &HE0E0E0
    SetPixel lHDC, 17, 8, &HE1E1E1
    SetPixel lHDC, 0, 9, &HFAFAFA
    SetPixel lHDC, 1, 9, &HAEAEAE
    SetPixel lHDC, 2, 9, &H919191
    SetPixel lHDC, 3, 9, &HB9B9B9
    SetPixel lHDC, 4, 9, &HC4C4C4
    SetPixel lHDC, 5, 9, &HCECECE
    SetPixel lHDC, 6, 9, &HD1D1D1
    SetPixel lHDC, 7, 9, &HDADADA
    SetPixel lHDC, 8, 9, &HDCDCDC
    SetPixel lHDC, 9, 9, &HDBDBDB
    SetPixel lHDC, 10, 9, &HDFDFDF
    SetPixel lHDC, 11, 9, &HE1E3E1
    SetPixel lHDC, 12, 9, &HE2E3E2
    SetPixel lHDC, 13, 9, &HE5E2E3
    SetPixel lHDC, 14, 9, &HE5E2E3
    SetPixel lHDC, 15, 9, &HE5E2E3
    SetPixel lHDC, 16, 9, &HE5E2E3
    SetPixel lHDC, 17, 9, &HE5E2E3
    SetPixel lHDC, 0, 10, &HF7F7F7
    SetPixel lHDC, 1, 10, &HA0A0A0
    SetPixel lHDC, 2, 10, &H999999
    SetPixel lHDC, 3, 10, &HC3C3C3
    SetPixel lHDC, 4, 10, &HC9C9C9
    SetPixel lHDC, 5, 10, &HD5D5D5
    SetPixel lHDC, 6, 10, &HD7D7D7
    SetPixel lHDC, 7, 10, &HDFDFDF
    SetPixel lHDC, 8, 10, &HE0E0E0
    SetPixel lHDC, 9, 10, &HE0E0E0
    SetPixel lHDC, 10, 10, &HE4E4E4
    SetPixel lHDC, 11, 10, &HE6E8E6
    SetPixel lHDC, 12, 10, &HE8E7E7
    SetPixel lHDC, 13, 10, &HEAE7E8
    SetPixel lHDC, 14, 10, &HEAE7E8
    SetPixel lHDC, 15, 10, &HEAE7E8
    SetPixel lHDC, 16, 10, &HEAE7E8
    SetPixel lHDC, 17, 10, &HEAE7E8
    SetPixel lHDC, 0, 11, &HF5F5F5
    SetPixel lHDC, 1, 11, &HA3A3A3
    SetPixel lHDC, 2, 11, &H9B9B9B
    SetPixel lHDC, 3, 11, &HC6C6C6
    SetPixel lHDC, 4, 11, &HD3D3D3
    SetPixel lHDC, 5, 11, &HD6D6D6
    SetPixel lHDC, 6, 11, &HDDDDDD
    SetPixel lHDC, 7, 11, &HE1E1E1
    SetPixel lHDC, 8, 11, &HE3E3E3
    SetPixel lHDC, 9, 11, &HE6E6E6
    SetPixel lHDC, 10, 11, &HE7E8E7
    SetPixel lHDC, 11, 11, &HE9EAE9
    SetPixel lHDC, 12, 11, &HE8EAE9
    SetPixel lHDC, 13, 11, &HE8EBE9
    SetPixel lHDC, 14, 11, &HE8EBE9
    SetPixel lHDC, 15, 11, &HE8EBE9
    SetPixel lHDC, 16, 11, &HE8EBE9
    SetPixel lHDC, 17, 11, &HE8EBE9
    SetPixel lHDC, 0, 12, &HF5F5F5
    SetPixel lHDC, 1, 12, &HAAAAAA
    SetPixel lHDC, 2, 12, &H8E8E8E
    SetPixel lHDC, 3, 12, &HD0D0D0
    SetPixel lHDC, 4, 12, &HDADADA
    SetPixel lHDC, 5, 12, &HDFDFDF
    SetPixel lHDC, 6, 12, &HE4E4E4
    SetPixel lHDC, 7, 12, &HE6E6E6
    SetPixel lHDC, 8, 12, &HE8E8E8
    SetPixel lHDC, 9, 12, &HECECEC
    SetPixel lHDC, 10, 12, &HEEEFEE
    SetPixel lHDC, 11, 12, &HEEF0EF
    SetPixel lHDC, 12, 12, &HEEF0EF
    SetPixel lHDC, 13, 12, &HEEF1EF
    SetPixel lHDC, 14, 12, &HEEF1EF
    SetPixel lHDC, 15, 12, &HEEF1EF
    SetPixel lHDC, 16, 12, &HEEF1EF
    SetPixel lHDC, 17, 12, &HEEF1EF

    tmph = lh - 22
    SetPixel lHDC, 0, tmph + 12, &HF5F5F5
    SetPixel lHDC, 1, tmph + 12, &HAAAAAA
    SetPixel lHDC, 2, tmph + 12, &H8E8E8E
    SetPixel lHDC, 3, tmph + 12, &HD0D0D0
    SetPixel lHDC, 4, tmph + 12, &HDADADA
    SetPixel lHDC, 5, tmph + 12, &HDFDFDF
    SetPixel lHDC, 6, tmph + 12, &HE4E4E4
    SetPixel lHDC, 7, tmph + 12, &HE6E6E6
    SetPixel lHDC, 8, tmph + 12, &HE8E8E8
    SetPixel lHDC, 9, tmph + 12, &HECECEC
    SetPixel lHDC, 10, tmph + 12, &HEEEFEE
    SetPixel lHDC, 11, tmph + 12, &HEEF0EF
    SetPixel lHDC, 12, tmph + 12, &HEEF0EF
    SetPixel lHDC, 13, tmph + 12, &HEEF1EF
    SetPixel lHDC, 14, tmph + 12, &HEEF1EF
    SetPixel lHDC, 15, tmph + 12, &HEEF1EF
    SetPixel lHDC, 16, tmph + 12, &HEEF1EF
    SetPixel lHDC, 17, tmph + 12, &HEEF1EF
    SetPixel lHDC, 0, tmph + 13, &HF7F7F7
    SetPixel lHDC, 1, tmph + 13, &HC2C2C2
    SetPixel lHDC, 2, tmph + 13, &H838383
    SetPixel lHDC, 3, tmph + 13, &HCFCFCF
    SetPixel lHDC, 4, tmph + 13, &HDEDEDE
    SetPixel lHDC, 5, tmph + 13, &HE3E3E3
    SetPixel lHDC, 6, tmph + 13, &HE8E8E8
    SetPixel lHDC, 7, tmph + 13, &HEAEAEA
    SetPixel lHDC, 8, tmph + 13, &HEDEDED
    SetPixel lHDC, 9, tmph + 13, &HF1F1F1
    SetPixel lHDC, 10, tmph + 13, &HF2F2F2
    SetPixel lHDC, 11, tmph + 13, &HF2F2F2
    SetPixel lHDC, 12, tmph + 13, &HF2F2F2
    SetPixel lHDC, 13, tmph + 13, &HF2F2F2
    SetPixel lHDC, 14, tmph + 13, &HF2F2F2
    SetPixel lHDC, 15, tmph + 13, &HF2F2F2
    SetPixel lHDC, 16, tmph + 13, &HF2F2F2
    SetPixel lHDC, 17, tmph + 13, &HF2F2F2
    SetPixel lHDC, 0, tmph + 14, &HFBFBFB
    SetPixel lHDC, 1, tmph + 14, &HE1E1E1
    SetPixel lHDC, 2, tmph + 14, &H818181
    SetPixel lHDC, 3, tmph + 14, &HABABAB
    SetPixel lHDC, 4, tmph + 14, &HDCDCDC
    SetPixel lHDC, 5, tmph + 14, &HE5E5E5
    SetPixel lHDC, 6, tmph + 14, &HEDEDED
    SetPixel lHDC, 7, tmph + 14, &HEFEFEF
    SetPixel lHDC, 8, tmph + 14, &HF1F1F1
    SetPixel lHDC, 9, tmph + 14, &HF4F4F4
    SetPixel lHDC, 10, tmph + 14, &HF5F5F5
    SetPixel lHDC, 11, tmph + 14, &HF5F5F5
    SetPixel lHDC, 12, tmph + 14, &HF5F5F5
    SetPixel lHDC, 13, tmph + 14, &HF5F5F5
    SetPixel lHDC, 14, tmph + 14, &HF5F5F5
    SetPixel lHDC, 15, tmph + 14, &HF5F5F5
    SetPixel lHDC, 16, tmph + 14, &HF5F5F5
    SetPixel lHDC, 17, tmph + 14, &HF5F5F5
    SetPixel lHDC, 0, tmph + 15, &HFEFEFE
    SetPixel lHDC, 1, tmph + 15, &HEDEDED
    SetPixel lHDC, 2, tmph + 15, &HA0A0A0
    SetPixel lHDC, 3, tmph + 15, &H898989
    SetPixel lHDC, 4, tmph + 15, &HDEDEDE
    SetPixel lHDC, 5, tmph + 15, &HE9E9E9
    SetPixel lHDC, 6, tmph + 15, &HEEEEEE
    SetPixel lHDC, 7, tmph + 15, &HF4F4F4
    SetPixel lHDC, 8, tmph + 15, &HF5F5F5
    SetPixel lHDC, 9, tmph + 15, &HFAFAFA
    SetPixel lHDC, 10, tmph + 15, &HFFFDFD
    SetPixel lHDC, 11, tmph + 15, &HFFFEFE
    SetPixel lHDC, 12, tmph + 15, &HFFFDFD
    SetPixel lHDC, 13, tmph + 15, &HFFFEFE
    SetPixel lHDC, 14, tmph + 15, &HFFFDFD
    SetPixel lHDC, 15, tmph + 15, &HFFFEFE
    SetPixel lHDC, 16, tmph + 15, &HFFFDFD
    SetPixel lHDC, 17, tmph + 15, &HFFFEFE
    SetPixel lHDC, 1, tmph + 16, &HF6F6F6
    SetPixel lHDC, 2, tmph + 16, &HD6D6D6
    SetPixel lHDC, 3, tmph + 16, &H7B7B7B
    SetPixel lHDC, 4, tmph + 16, &H8D8D8D
    SetPixel lHDC, 5, tmph + 16, &HE4E4E4
    SetPixel lHDC, 6, tmph + 16, &HF0F0F0
    SetPixel lHDC, 7, tmph + 16, &HF6F6F6
    SetPixel lHDC, 8, tmph + 16, &HFEFEFE
    SetPixel lHDC, 9, tmph + 16, &HFEFEFE
    SetPixel lHDC, 10, tmph + 16, &HFFFEFE
    SetPixel lHDC, 12, tmph + 16, &HFFFEFE
    SetPixel lHDC, 14, tmph + 16, &HFFFEFE
    SetPixel lHDC, 16, tmph + 16, &HFFFEFE
    SetPixel lHDC, 1, tmph + 17, &HFDFDFD
    SetPixel lHDC, 2, tmph + 17, &HEDEDED
    SetPixel lHDC, 3, tmph + 17, &HBEBEBE
    SetPixel lHDC, 4, tmph + 17, &H727272
    SetPixel lHDC, 5, tmph + 17, &H898989
    SetPixel lHDC, 6, tmph + 17, &HEBEBEB
    SetPixel lHDC, 7, tmph + 17, &HF5F5F5
    SetPixel lHDC, 8, tmph + 17, &HFCFCFC
    SetPixel lHDC, 10, tmph + 17, &HFDFDFD
    SetPixel lHDC, 11, tmph + 17, &HFDFDFD
    SetPixel lHDC, 12, tmph + 17, &HFDFDFD
    SetPixel lHDC, 13, tmph + 17, &HFDFDFD
    SetPixel lHDC, 14, tmph + 17, &HFDFDFD
    SetPixel lHDC, 15, tmph + 17, &HFDFDFD
    SetPixel lHDC, 16, tmph + 17, &HFDFDFD
    SetPixel lHDC, 17, tmph + 17, &HFDFDFD
    SetPixel lHDC, 2, tmph + 18, &HF9F9F9
    SetPixel lHDC, 3, tmph + 18, &HE6E6E6
    SetPixel lHDC, 4, tmph + 18, &HB9B9B9
    SetPixel lHDC, 5, tmph + 18, &H717171
    SetPixel lHDC, 6, tmph + 18, &H787878
    SetPixel lHDC, 7, tmph + 18, &HB6B6B6
    SetPixel lHDC, 8, tmph + 18, &HF7F7F7
    SetPixel lHDC, 9, tmph + 18, &HFCFCFC
    SetPixel lHDC, 10, tmph + 18, &HFEFEFE
    SetPixel lHDC, 11, tmph + 18, &HFEFEFE
    SetPixel lHDC, 12, tmph + 18, &HFEFEFE
    SetPixel lHDC, 13, tmph + 18, &HFEFEFE
    SetPixel lHDC, 14, tmph + 18, &HFEFEFE
    SetPixel lHDC, 15, tmph + 18, &HFEFEFE
    SetPixel lHDC, 16, tmph + 18, &HFEFEFE
    SetPixel lHDC, 17, tmph + 18, &HFEFEFE
    SetPixel lHDC, 2, tmph + 19, &HFEFEFE
    SetPixel lHDC, 3, tmph + 19, &HF8F8F8
    SetPixel lHDC, 4, tmph + 19, &HE6E6E6
    SetPixel lHDC, 5, tmph + 19, &HC8C8C8
    SetPixel lHDC, 6, tmph + 19, &H8E8E8E
    SetPixel lHDC, 7, tmph + 19, &H6C6C6C
    SetPixel lHDC, 8, tmph + 19, &H757575
    SetPixel lHDC, 9, tmph + 19, &H9F9F9F
    SetPixel lHDC, 10, tmph + 19, &HC7C7C7
    SetPixel lHDC, 11, tmph + 19, &HE9E9E9
    SetPixel lHDC, 12, tmph + 19, &HFBFBFB
    SetPixel lHDC, 13, tmph + 19, &HFBFBFB
    SetPixel lHDC, 14, tmph + 19, &HFBFBFB
    SetPixel lHDC, 15, tmph + 19, &HFBFBFB
    SetPixel lHDC, 16, tmph + 19, &HFBFBFB
    SetPixel lHDC, 17, tmph + 19, &HFBFBFB
    SetPixel lHDC, 3, tmph + 20, &HFEFEFE
    SetPixel lHDC, 4, tmph + 20, &HF9F9F9
    SetPixel lHDC, 5, tmph + 20, &HECECEC
    SetPixel lHDC, 6, tmph + 20, &HDADADA
    SetPixel lHDC, 7, tmph + 20, &HC1C1C1
    SetPixel lHDC, 8, tmph + 20, &H9D9D9D
    SetPixel lHDC, 9, tmph + 20, &H7B7B7B
    SetPixel lHDC, 10, tmph + 20, &H5E5E5E
    SetPixel lHDC, 11, tmph + 20, &H535353
    SetPixel lHDC, 12, tmph + 20, &H4D4D4D
    SetPixel lHDC, 13, tmph + 20, &H4B4B4B
    SetPixel lHDC, 14, tmph + 20, &H505050
    SetPixel lHDC, 15, tmph + 20, &H525252
    SetPixel lHDC, 16, tmph + 20, &H555555
    SetPixel lHDC, 17, tmph + 20, &H545454
    SetPixel lHDC, 5, tmph + 21, &HFCFCFC
    SetPixel lHDC, 6, tmph + 21, &HF5F5F5
    SetPixel lHDC, 7, tmph + 21, &HEBEBEB
    SetPixel lHDC, 8, tmph + 21, &HE1E1E1
    SetPixel lHDC, 9, tmph + 21, &HD6D6D6
    SetPixel lHDC, 10, tmph + 21, &HCECECE
    SetPixel lHDC, 11, tmph + 21, &HC9C9C9
    SetPixel lHDC, 12, tmph + 21, &HC7C7C7
    SetPixel lHDC, 13, tmph + 21, &HC7C7C7
    SetPixel lHDC, 14, tmph + 21, &HC6C6C6
    SetPixel lHDC, 15, tmph + 21, &HC6C6C6
    SetPixel lHDC, 16, tmph + 21, &HC5C5C5
    SetPixel lHDC, 17, tmph + 21, &HC5C5C5
    SetPixel lHDC, 7, tmph + 22, &HFDFDFD
    SetPixel lHDC, 8, tmph + 22, &HF9F9F9
    SetPixel lHDC, 9, tmph + 22, &HF4F4F4
    SetPixel lHDC, 10, tmph + 22, &HF0F0F0
    SetPixel lHDC, 11, tmph + 22, &HEEEEEE
    SetPixel lHDC, 12, tmph + 22, &HEDEDED
    SetPixel lHDC, 13, tmph + 22, &HECECEC
    SetPixel lHDC, 14, tmph + 22, &HECECEC
    SetPixel lHDC, 15, tmph + 22, &HECECEC
    SetPixel lHDC, 16, tmph + 22, &HECECEC
    SetPixel lHDC, 17, tmph + 22, &HECECEC

    tmpw = lw - 34
    SetPixel lHDC, tmpw + 17, 0, &H67696A
    SetPixel lHDC, tmpw + 18, 0, &H666869
    SetPixel lHDC, tmpw + 19, 0, &H716F6F
    SetPixel lHDC, tmpw + 20, 0, &H6F6D6D
    SetPixel lHDC, tmpw + 21, 0, &H6F706E
    SetPixel lHDC, tmpw + 22, 0, &H727371
    SetPixel lHDC, tmpw + 23, 0, &H6E6E6E
    SetPixel lHDC, tmpw + 24, 0, &H707070
    SetPixel lHDC, tmpw + 25, 0, &HA6A6A6
    SetPixel lHDC, tmpw + 26, 0, &HEEEEEE
    SetPixel lHDC, tmpw + 34, 0, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 1, &HF5F4F6
    SetPixel lHDC, tmpw + 18, 1, &HF5F4F6
    SetPixel lHDC, tmpw + 19, 1, &HF5F4F6
    SetPixel lHDC, tmpw + 20, 1, &HF5F4F6
    SetPixel lHDC, tmpw + 21, 1, &HF4F3F5
    SetPixel lHDC, tmpw + 22, 1, &HF1F0F2
    SetPixel lHDC, tmpw + 23, 1, &HE0E0E0
    SetPixel lHDC, tmpw + 24, 1, &HC3C3C3
    SetPixel lHDC, tmpw + 25, 1, &H848484
    SetPixel lHDC, tmpw + 26, 1, &H6B6B6B
    SetPixel lHDC, tmpw + 27, 1, &HA0A0A0
    SetPixel lHDC, tmpw + 28, 1, &HF7F7F7
    SetPixel lHDC, tmpw + 34, 1, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 2, &HF3F2F4
    SetPixel lHDC, tmpw + 18, 2, &HF2F1F3
    SetPixel lHDC, tmpw + 19, 2, &HF3F2F4
    SetPixel lHDC, tmpw + 20, 2, &HF3F2F4
    SetPixel lHDC, tmpw + 21, 2, &HF0EFF1
    SetPixel lHDC, tmpw + 22, 2, &HF2F1F3
    SetPixel lHDC, tmpw + 23, 2, &HF6F6F6
    SetPixel lHDC, tmpw + 24, 2, &HE8E8E8
    SetPixel lHDC, tmpw + 25, 2, &HE0E0E0
    SetPixel lHDC, tmpw + 26, 2, &H999999
    SetPixel lHDC, tmpw + 27, 2, &H696969
    SetPixel lHDC, tmpw + 28, 2, &H717171
    SetPixel lHDC, tmpw + 29, 2, &HEBEBEB
    SetPixel lHDC, tmpw + 34, 2, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 3, &HEEEEEE
    SetPixel lHDC, tmpw + 18, 3, &HEDEDED
    SetPixel lHDC, tmpw + 19, 3, &HEEEEEE
    SetPixel lHDC, tmpw + 20, 3, &HEEEEEE
    SetPixel lHDC, tmpw + 21, 3, &HEEEEEE
    SetPixel lHDC, tmpw + 22, 3, &HEEEEEE
    SetPixel lHDC, tmpw + 23, 3, &HE9E9E9
    SetPixel lHDC, tmpw + 24, 3, &HEAEAEA
    SetPixel lHDC, tmpw + 25, 3, &HE7E7E7
    SetPixel lHDC, tmpw + 26, 3, &HD0D0D0
    SetPixel lHDC, tmpw + 27, 3, &H939393
    SetPixel lHDC, tmpw + 28, 3, &H727272
    SetPixel lHDC, tmpw + 29, 3, &H6F6F6F
    SetPixel lHDC, tmpw + 30, 3, &HEFEFEF
    SetPixel lHDC, tmpw + 34, 3, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 4, &HEBEBEB
    SetPixel lHDC, tmpw + 18, 4, &HEBEBEB
    SetPixel lHDC, tmpw + 19, 4, &HEBEBEB
    SetPixel lHDC, tmpw + 20, 4, &HEBEBEB
    SetPixel lHDC, tmpw + 21, 4, &HEDEDED
    SetPixel lHDC, tmpw + 22, 4, &HE6E6E6
    SetPixel lHDC, tmpw + 23, 4, &HE9E9E9
    SetPixel lHDC, tmpw + 24, 4, &HE6E6E6
    SetPixel lHDC, tmpw + 25, 4, &HDEDEDE
    SetPixel lHDC, tmpw + 26, 4, &HDCDCDC
    SetPixel lHDC, tmpw + 27, 4, &HB2B2B2
    SetPixel lHDC, tmpw + 28, 4, &H919191
    SetPixel lHDC, tmpw + 29, 4, &H6E6E6E
    SetPixel lHDC, tmpw + 30, 4, &H7F7F7F
    SetPixel lHDC, tmpw + 31, 4, &HFAFAFA
    SetPixel lHDC, tmpw + 34, 4, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 5, &HEBE8EA
    SetPixel lHDC, tmpw + 18, 5, &HEAE7E9
    SetPixel lHDC, tmpw + 19, 5, &HEBE8EA
    SetPixel lHDC, tmpw + 20, 5, &HEBE8EA
    SetPixel lHDC, tmpw + 21, 5, &HE5E8E6
    SetPixel lHDC, tmpw + 22, 5, &HE7EAE8
    SetPixel lHDC, tmpw + 23, 5, &HE5E5E5
    SetPixel lHDC, tmpw + 24, 5, &HE3E3E3
    SetPixel lHDC, tmpw + 25, 5, &HDFDFDF
    SetPixel lHDC, tmpw + 26, 5, &HDCDCDC
    SetPixel lHDC, tmpw + 27, 5, &HC3C3C3
    SetPixel lHDC, tmpw + 28, 5, &HA7A7A7
    SetPixel lHDC, tmpw + 29, 5, &H969696
    SetPixel lHDC, tmpw + 30, 5, &H717171
    SetPixel lHDC, tmpw + 31, 5, &HC5C5C5
    SetPixel lHDC, tmpw + 32, 5, &HFEFEFE
    SetPixel lHDC, tmpw + 34, 5, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 6, &HEBE8EA
    SetPixel lHDC, tmpw + 18, 6, &HEBE8EA
    SetPixel lHDC, tmpw + 19, 6, &HEBE8EA
    SetPixel lHDC, tmpw + 20, 6, &HEBE8EA
    SetPixel lHDC, tmpw + 21, 6, &HE8EBE9
    SetPixel lHDC, tmpw + 22, 6, &HE3E6E4
    SetPixel lHDC, tmpw + 23, 6, &HE5E5E5
    SetPixel lHDC, tmpw + 24, 6, &HE2E2E2
    SetPixel lHDC, tmpw + 25, 6, &HE0E0E0
    SetPixel lHDC, tmpw + 26, 6, &HDADADA
    SetPixel lHDC, tmpw + 27, 6, &HC7C7C7
    SetPixel lHDC, tmpw + 28, 6, &HB5B5B5
    SetPixel lHDC, tmpw + 29, 6, &HA6A6A6
    SetPixel lHDC, tmpw + 30, 6, &H8C8C8C
    SetPixel lHDC, tmpw + 31, 6, &H808080
    SetPixel lHDC, tmpw + 32, 6, &HF8F8F8
    SetPixel lHDC, tmpw + 34, 6, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 7, &HEAEAEA
    SetPixel lHDC, tmpw + 18, 7, &HEAEAEA
    SetPixel lHDC, tmpw + 19, 7, &HEAEAEA
    SetPixel lHDC, tmpw + 20, 7, &HEAEAEA
    SetPixel lHDC, tmpw + 21, 7, &HE9E6E8
    SetPixel lHDC, tmpw + 22, 7, &HE9E6E8
    SetPixel lHDC, tmpw + 23, 7, &HE4E4E4
    SetPixel lHDC, tmpw + 24, 7, &HE2E2E2
    SetPixel lHDC, tmpw + 25, 7, &HDFDFDF
    SetPixel lHDC, tmpw + 26, 7, &HD7D7D7
    SetPixel lHDC, tmpw + 27, 7, &HC4C4C4
    SetPixel lHDC, tmpw + 28, 7, &HB7B7B7
    SetPixel lHDC, tmpw + 29, 7, &HB4B5B3
    SetPixel lHDC, tmpw + 30, 7, &H9D9E9C
    SetPixel lHDC, tmpw + 31, 7, &H777777
    SetPixel lHDC, tmpw + 32, 7, &HE7E7E7
    SetPixel lHDC, tmpw + 34, 7, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 8, &HE1E1E1
    SetPixel lHDC, tmpw + 18, 8, &HE0E0E0
    SetPixel lHDC, tmpw + 19, 8, &HE1E1E1
    SetPixel lHDC, tmpw + 20, 8, &HE1E1E1
    SetPixel lHDC, tmpw + 21, 8, &HDFDCDE
    SetPixel lHDC, tmpw + 22, 8, &HDDDADC
    SetPixel lHDC, tmpw + 23, 8, &HDBDBDB
    SetPixel lHDC, tmpw + 24, 8, &HD6D6D6
    SetPixel lHDC, tmpw + 25, 8, &HD5D5D5
    SetPixel lHDC, tmpw + 26, 8, &HD1D1D1
    SetPixel lHDC, tmpw + 27, 8, &HC9C9C9
    SetPixel lHDC, tmpw + 28, 8, &HC4C4C4
    SetPixel lHDC, tmpw + 29, 8, &HC0C1BF
    SetPixel lHDC, tmpw + 30, 8, &HAFB0AE
    SetPixel lHDC, tmpw + 31, 8, &H818181
    SetPixel lHDC, tmpw + 32, 8, &HC3C3C3
    SetPixel lHDC, tmpw + 33, 8, &HFDFDFD
    SetPixel lHDC, tmpw + 34, 8, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 9, &HE5E2E3
    SetPixel lHDC, tmpw + 18, 9, &HE5E2E3
    SetPixel lHDC, tmpw + 19, 9, &HE5E2E3
    SetPixel lHDC, tmpw + 20, 9, &HE5E2E3
    SetPixel lHDC, tmpw + 21, 9, &HE1E1E1
    SetPixel lHDC, tmpw + 22, 9, &HE1E1E1
    SetPixel lHDC, tmpw + 23, 9, &HE1E1E1
    SetPixel lHDC, tmpw + 24, 9, &HDDDDDD
    SetPixel lHDC, tmpw + 25, 9, &HDBDBDB
    SetPixel lHDC, tmpw + 26, 9, &HD8D8D8
    SetPixel lHDC, tmpw + 27, 9, &HD2D2D2
    SetPixel lHDC, tmpw + 28, 9, &HCBCBCB
    SetPixel lHDC, tmpw + 29, 9, &HC4C4C4
    SetPixel lHDC, tmpw + 30, 9, &HBABABA
    SetPixel lHDC, tmpw + 31, 9, &H989898
    SetPixel lHDC, tmpw + 32, 9, &HA6A6A6
    SetPixel lHDC, tmpw + 33, 9, &HF9F9F9
    SetPixel lHDC, tmpw + 34, 9, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 10, &HEAE7E8
    SetPixel lHDC, tmpw + 18, 10, &HEAE7E8
    SetPixel lHDC, tmpw + 19, 10, &HEAE7E8
    SetPixel lHDC, tmpw + 20, 10, &HEAE7E8
    SetPixel lHDC, tmpw + 21, 10, &HE7E7E7
    SetPixel lHDC, tmpw + 22, 10, &HE6E6E6
    SetPixel lHDC, tmpw + 23, 10, &HE4E4E4
    SetPixel lHDC, tmpw + 24, 10, &HE0E0E0
    SetPixel lHDC, tmpw + 25, 10, &HE0E0E0
    SetPixel lHDC, tmpw + 26, 10, &HDEDEDE
    SetPixel lHDC, tmpw + 27, 10, &HD9D9D9
    SetPixel lHDC, tmpw + 28, 10, &HD3D3D3
    SetPixel lHDC, tmpw + 29, 10, &HCCCCCC
    SetPixel lHDC, tmpw + 30, 10, &HC3C3C3
    SetPixel lHDC, tmpw + 31, 10, &HA3A3A3
    SetPixel lHDC, tmpw + 32, 10, &H9C9C9C
    SetPixel lHDC, tmpw + 33, 10, &HF6F6F6
    SetPixel lHDC, tmpw + 34, 10, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 11, &HE8EBE9
    SetPixel lHDC, tmpw + 18, 11, &HE8EBE9
    SetPixel lHDC, tmpw + 19, 11, &HE8EBE9
    SetPixel lHDC, tmpw + 20, 11, &HE8EBE9
    SetPixel lHDC, tmpw + 21, 11, &HE9EAE8
    SetPixel lHDC, tmpw + 22, 11, &HE8E9E7
    SetPixel lHDC, tmpw + 23, 11, &HE9E9E9
    SetPixel lHDC, tmpw + 24, 11, &HE5E5E5
    SetPixel lHDC, tmpw + 25, 11, &HE4E4E4
    SetPixel lHDC, tmpw + 26, 11, &HE2E2E2
    SetPixel lHDC, tmpw + 27, 11, &HDBDBDB
    SetPixel lHDC, tmpw + 28, 11, &HD9D9D9
    SetPixel lHDC, tmpw + 29, 11, &HD1D1D1
    SetPixel lHDC, tmpw + 30, 11, &HC8C8C8
    SetPixel lHDC, tmpw + 31, 11, &HA4A4A4
    SetPixel lHDC, tmpw + 32, 11, &HA2A2A2
    SetPixel lHDC, tmpw + 33, 11, &HF4F4F4
    SetPixel lHDC, tmpw + 34, 11, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 12, &HEEF1EF
    SetPixel lHDC, tmpw + 18, 12, &HEEF1EF
    SetPixel lHDC, tmpw + 19, 12, &HEEF1EF
    SetPixel lHDC, tmpw + 20, 12, &HEEF1EF
    SetPixel lHDC, tmpw + 21, 12, &HEEEFED
    SetPixel lHDC, tmpw + 22, 12, &HEFF0EE
    SetPixel lHDC, tmpw + 23, 12, &HEEEEEE
    SetPixel lHDC, tmpw + 24, 12, &HECECEC
    SetPixel lHDC, tmpw + 25, 12, &HEAEAEA
    SetPixel lHDC, tmpw + 26, 12, &HE7E7E7
    SetPixel lHDC, tmpw + 27, 12, &HE2E2E2
    SetPixel lHDC, tmpw + 28, 12, &HDFDFDF
    SetPixel lHDC, tmpw + 29, 12, &HD8D8D8
    SetPixel lHDC, tmpw + 30, 12, &HD4D4D4
    SetPixel lHDC, tmpw + 31, 12, &H999999
    SetPixel lHDC, tmpw + 32, 12, &HAFAFAF
    SetPixel lHDC, tmpw + 33, 12, &HF5F5F5
    SetPixel lHDC, tmpw + 34, 12, &HFFFFFFFF

    tmph = lh - 22
    tmpw = lw - 34
    SetPixel lHDC, tmpw + 17, tmph + 12, &HEEF1EF
    SetPixel lHDC, tmpw + 18, tmph + 12, &HEEF1EF
    SetPixel lHDC, tmpw + 19, tmph + 12, &HEEF1EF
    SetPixel lHDC, tmpw + 20, tmph + 12, &HEEF1EF
    SetPixel lHDC, tmpw + 21, tmph + 12, &HEEEFED
    SetPixel lHDC, tmpw + 22, tmph + 12, &HEFF0EE
    SetPixel lHDC, tmpw + 23, tmph + 12, &HEEEEEE
    SetPixel lHDC, tmpw + 24, tmph + 12, &HECECEC
    SetPixel lHDC, tmpw + 25, tmph + 12, &HEAEAEA
    SetPixel lHDC, tmpw + 26, tmph + 12, &HE7E7E7
    SetPixel lHDC, tmpw + 27, tmph + 12, &HE2E2E2
    SetPixel lHDC, tmpw + 28, tmph + 12, &HDFDFDF
    SetPixel lHDC, tmpw + 29, tmph + 12, &HD8D8D8
    SetPixel lHDC, tmpw + 30, tmph + 12, &HD4D4D4
    SetPixel lHDC, tmpw + 31, tmph + 12, &H999999
    SetPixel lHDC, tmpw + 32, tmph + 12, &HAFAFAF
    SetPixel lHDC, tmpw + 33, tmph + 12, &HF5F5F5
    SetPixel lHDC, tmpw + 17, tmph + 13, &HF2F2F2
    SetPixel lHDC, tmpw + 18, tmph + 13, &HF2F2F2
    SetPixel lHDC, tmpw + 19, tmph + 13, &HF2F2F2
    SetPixel lHDC, tmpw + 20, tmph + 13, &HF2F2F2
    SetPixel lHDC, tmpw + 21, tmph + 13, &HF5F4F6
    SetPixel lHDC, tmpw + 22, tmph + 13, &HF0EFF1
    SetPixel lHDC, tmpw + 23, tmph + 13, &HF2F2F2
    SetPixel lHDC, tmpw + 24, tmph + 13, &HF2F2F2
    SetPixel lHDC, tmpw + 25, tmph + 13, &HECECEC
    SetPixel lHDC, tmpw + 26, tmph + 13, &HEAEAEA
    SetPixel lHDC, tmpw + 27, tmph + 13, &HEBEBEB
    SetPixel lHDC, tmpw + 28, tmph + 13, &HE3E3E3
    SetPixel lHDC, tmpw + 29, tmph + 13, &HDEDEDE
    SetPixel lHDC, tmpw + 30, tmph + 13, &HD1D1D1
    SetPixel lHDC, tmpw + 31, tmph + 13, &H8A8A8A
    SetPixel lHDC, tmpw + 32, tmph + 13, &HD5D5D5
    SetPixel lHDC, tmpw + 33, tmph + 13, &HF8F8F8
    SetPixel lHDC, tmpw + 17, tmph + 14, &HF5F5F5
    SetPixel lHDC, tmpw + 18, tmph + 14, &HF5F5F5
    SetPixel lHDC, tmpw + 19, tmph + 14, &HF5F5F5
    SetPixel lHDC, tmpw + 20, tmph + 14, &HF5F5F5
    SetPixel lHDC, tmpw + 21, tmph + 14, &HF8F7F9
    SetPixel lHDC, tmpw + 22, tmph + 14, &HF7F6F8
    SetPixel lHDC, tmpw + 23, tmph + 14, &HF7F7F7
    SetPixel lHDC, tmpw + 24, tmph + 14, &HF5F5F5
    SetPixel lHDC, tmpw + 25, tmph + 14, &HEFEFEF
    SetPixel lHDC, tmpw + 26, tmph + 14, &HEEEEEE
    SetPixel lHDC, tmpw + 27, tmph + 14, &HECECEC
    SetPixel lHDC, tmpw + 28, tmph + 14, &HE5E5E5
    SetPixel lHDC, tmpw + 29, tmph + 14, &HDEDEDE
    SetPixel lHDC, tmpw + 30, tmph + 14, &HB3B3B3
    SetPixel lHDC, tmpw + 31, tmph + 14, &H808080
    SetPixel lHDC, tmpw + 32, tmph + 14, &HE8E8E8
    SetPixel lHDC, tmpw + 33, tmph + 14, &HFDFDFD
    SetPixel lHDC, tmpw + 17, tmph + 15, &HFFFEFE
    SetPixel lHDC, tmpw + 18, tmph + 15, &HFFFDFD
    SetPixel lHDC, tmpw + 19, tmph + 15, &HFFFEFE
    SetPixel lHDC, tmpw + 20, tmph + 15, &HFFFEFE
    SetPixel lHDC, tmpw + 21, tmph + 15, &HFBFBFB
    SetPixel lHDC, tmpw + 22, tmph + 15, &HFCFCFC
    SetPixel lHDC, tmpw + 23, tmph + 15, &HFEFEFE
    SetPixel lHDC, tmpw + 24, tmph + 15, &HF8F8F8
    SetPixel lHDC, tmpw + 25, tmph + 15, &HF7F7F7
    SetPixel lHDC, tmpw + 26, tmph + 15, &HF5F5F5
    SetPixel lHDC, tmpw + 27, tmph + 15, &HEDEDED
    SetPixel lHDC, tmpw + 28, tmph + 15, &HEAEAEA
    SetPixel lHDC, tmpw + 29, tmph + 15, &HE0E0E0
    SetPixel lHDC, tmpw + 30, tmph + 15, &H8D8D8D
    SetPixel lHDC, tmpw + 31, tmph + 15, &HBABABA
    SetPixel lHDC, tmpw + 32, tmph + 15, &HF1F1F1
    SetPixel lHDC, tmpw + 18, tmph + 16, &HFFFEFE
    SetPixel lHDC, tmpw + 22, tmph + 16, &HFEFEFE
    SetPixel lHDC, tmpw + 23, tmph + 16, &HFEFEFE
    SetPixel lHDC, tmpw + 25, tmph + 16, &HFCFCFC
    SetPixel lHDC, tmpw + 26, tmph + 16, &HF6F6F6
    SetPixel lHDC, tmpw + 27, tmph + 16, &HF2F2F2
    SetPixel lHDC, tmpw + 28, tmph + 16, &HE7E7E7
    SetPixel lHDC, tmpw + 29, tmph + 16, &H989898
    SetPixel lHDC, tmpw + 30, tmph + 16, &H828282
    SetPixel lHDC, tmpw + 31, tmph + 16, &HE2E2E2
    SetPixel lHDC, tmpw + 32, tmph + 16, &HF9F9F9
    SetPixel lHDC, tmpw + 17, tmph + 17, &HFDFDFD
    SetPixel lHDC, tmpw + 18, tmph + 17, &HFDFDFD
    SetPixel lHDC, tmpw + 19, tmph + 17, &HFDFDFD
    SetPixel lHDC, tmpw + 20, tmph + 17, &HFDFDFD
    SetPixel lHDC, tmpw + 21, tmph + 17, &HFEFEFE
    SetPixel lHDC, tmpw + 23, tmph + 17, &HFEFEFE
    SetPixel lHDC, tmpw + 25, tmph + 17, &HFEFEFE
    SetPixel lHDC, tmpw + 26, tmph + 17, &HF6F6F6
    SetPixel lHDC, tmpw + 27, tmph + 17, &HF1F1F1
    SetPixel lHDC, tmpw + 28, tmph + 17, &H979797
    SetPixel lHDC, tmpw + 29, tmph + 17, &H6F6F6F
    SetPixel lHDC, tmpw + 30, tmph + 17, &HD2D2D2
    SetPixel lHDC, tmpw + 31, tmph + 17, &HF2F2F2
    SetPixel lHDC, tmpw + 32, tmph + 17, &HFEFEFE
    SetPixel lHDC, tmpw + 17, tmph + 18, &HFEFEFE
    SetPixel lHDC, tmpw + 18, tmph + 18, &HFEFEFE
    SetPixel lHDC, tmpw + 19, tmph + 18, &HFEFEFE
    SetPixel lHDC, tmpw + 20, tmph + 18, &HFEFEFE
    SetPixel lHDC, tmpw + 22, tmph + 18, &HFDFDFD
    SetPixel lHDC, tmpw + 23, tmph + 18, &HFEFEFE
    SetPixel lHDC, tmpw + 24, tmph + 18, &HFDFDFD
    SetPixel lHDC, tmpw + 25, tmph + 18, &HFCFCFC
    SetPixel lHDC, tmpw + 26, tmph + 18, &HC5C5C5
    SetPixel lHDC, tmpw + 27, tmph + 18, &H838383
    SetPixel lHDC, tmpw + 28, tmph + 18, &H6F6F6F
    SetPixel lHDC, tmpw + 29, tmph + 18, &HC8C8C8
    SetPixel lHDC, tmpw + 30, tmph + 18, &HEBEBEB
    SetPixel lHDC, tmpw + 31, tmph + 18, &HFCFCFC
    SetPixel lHDC, tmpw + 17, tmph + 19, &HFBFBFB
    SetPixel lHDC, tmpw + 18, tmph + 19, &HFBFBFB
    SetPixel lHDC, tmpw + 19, tmph + 19, &HFBFBFB
    SetPixel lHDC, tmpw + 20, tmph + 19, &HFBFBFB
    SetPixel lHDC, tmpw + 21, tmph + 19, &HFAFAFA
    SetPixel lHDC, tmpw + 22, tmph + 19, &HEFEFEF
    SetPixel lHDC, tmpw + 23, tmph + 19, &HD0D0D0
    SetPixel lHDC, tmpw + 24, tmph + 19, &HA3A3A3
    SetPixel lHDC, tmpw + 25, tmph + 19, &H7E7E7E
    SetPixel lHDC, tmpw + 26, tmph + 19, &H6A6A6A
    SetPixel lHDC, tmpw + 27, tmph + 19, &H8F8F8F
    SetPixel lHDC, tmpw + 28, tmph + 19, &HCDCDCD
    SetPixel lHDC, tmpw + 29, tmph + 19, &HE8E8E8
    SetPixel lHDC, tmpw + 30, tmph + 19, &HFAFAFA
    SetPixel lHDC, tmpw + 17, tmph + 20, &H545454
    SetPixel lHDC, tmpw + 18, tmph + 20, &H555555
    SetPixel lHDC, tmpw + 19, tmph + 20, &H525252
    SetPixel lHDC, tmpw + 20, tmph + 20, &H505050
    SetPixel lHDC, tmpw + 21, tmph + 20, &H535353
    SetPixel lHDC, tmpw + 22, tmph + 20, &H525252
    SetPixel lHDC, tmpw + 23, tmph + 20, &H616161
    SetPixel lHDC, tmpw + 24, tmph + 20, &H7A7A7A
    SetPixel lHDC, tmpw + 25, tmph + 20, &HA3A3A3
    SetPixel lHDC, tmpw + 26, tmph + 20, &HC5C5C5
    SetPixel lHDC, tmpw + 27, tmph + 20, &HDADADA
    SetPixel lHDC, tmpw + 28, tmph + 20, &HEDEDED
    SetPixel lHDC, tmpw + 29, tmph + 20, &HFAFAFA
    SetPixel lHDC, tmpw + 17, tmph + 21, &HC5C5C5
    SetPixel lHDC, tmpw + 18, tmph + 21, &HC5C5C5
    SetPixel lHDC, tmpw + 19, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 20, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 21, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 22, tmph + 21, &HC9C9C9
    SetPixel lHDC, tmpw + 23, tmph + 21, &HCECECE
    SetPixel lHDC, tmpw + 24, tmph + 21, &HD7D7D7
    SetPixel lHDC, tmpw + 25, tmph + 21, &HE1E1E1
    SetPixel lHDC, tmpw + 26, tmph + 21, &HECECEC
    SetPixel lHDC, tmpw + 27, tmph + 21, &HF6F6F6
    SetPixel lHDC, tmpw + 28, tmph + 21, &HFDFDFD
    SetPixel lHDC, tmpw + 17, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 18, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 19, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 20, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 21, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 22, tmph + 22, &HEDEDED
    SetPixel lHDC, tmpw + 23, tmph + 22, &HF0F0F0
    SetPixel lHDC, tmpw + 24, tmph + 22, &HF4F4F4
    SetPixel lHDC, tmpw + 25, tmph + 22, &HFAFAFA
    SetPixel lHDC, tmpw + 26, tmph + 22, &HFDFDFD
    'Vlines
    tmph = 11
    tmph1 = lh - 10
    tmpw = lw - 34
    APILine 0, tmph, 0, tmph1, &HF7F7F7
    APILine 1, tmph, 1, tmph1, &HA0A0A0
    APILine 2, tmph, 2, tmph1, &H999999
    APILine 3, tmph, 3, tmph1, &HC3C3C3
    APILine 4, tmph, 4, tmph1, &HC9C9C9
    APILine 5, tmph, 5, tmph1, &HD5D5D5
    APILine 6, tmph, 6, tmph1, &HD7D7D7
    APILine 7, tmph, 7, tmph1, &HDFDFDF
    APILine 8, tmph, 8, tmph1, &HE0E0E0
    APILine 9, tmph, 9, tmph1, &HE0E0E0
    APILine 10, tmph, 10, tmph1, &HE4E4E4
    APILine 11, tmph, 11, tmph1, &HE6E8E6
    APILine 12, tmph, 12, tmph1, &HE8E7E7
    APILine 13, tmph, 13, tmph1, &HEAE7E8
    APILine 14, tmph, 14, tmph1, &HEAE7E8
    APILine 15, tmph, 15, tmph1, &HEAE7E8
    APILine 16, tmph, 16, tmph1, &HEAE7E8
    APILine 17, tmph, 17, tmph1, &HEAE7E8
    APILine tmpw + 17, tmph, tmpw + 17, tmph1, &HEAE7E8
    APILine tmpw + 18, tmph, tmpw + 18, tmph1, &HEAE7E8
    APILine tmpw + 19, tmph, tmpw + 19, tmph1, &HEAE7E8
    APILine tmpw + 20, tmph, tmpw + 20, tmph1, &HEAE7E8
    APILine tmpw + 21, tmph, tmpw + 21, tmph1, &HE7E7E7
    APILine tmpw + 22, tmph, tmpw + 22, tmph1, &HE6E6E6
    APILine tmpw + 23, tmph, tmpw + 23, tmph1, &HE4E4E4
    APILine tmpw + 24, tmph, tmpw + 24, tmph1, &HE0E0E0
    APILine tmpw + 25, tmph, tmpw + 25, tmph1, &HE0E0E0
    APILine tmpw + 26, tmph, tmpw + 26, tmph1, &HDEDEDE
    APILine tmpw + 27, tmph, tmpw + 27, tmph1, &HD9D9D9
    APILine tmpw + 28, tmph, tmpw + 28, tmph1, &HD3D3D3
    APILine tmpw + 29, tmph, tmpw + 29, tmph1, &HCCCCCC
    APILine tmpw + 30, tmph, tmpw + 30, tmph1, &HC3C3C3
    APILine tmpw + 31, tmph, tmpw + 31, tmph1, &HA3A3A3
    APILine tmpw + 32, tmph, tmpw + 32, tmph1, &H9C9C9C
    APILine tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6
    'HLines

    APILine 17, 0, lw - 17, 0, &H67696A
    APILine 17, 1, lw - 17, 1, &HF5F4F6
    APILine 17, 2, lw - 17, 2, &HF3F2F4
    APILine 17, 3, lw - 17, 3, &HEEEEEE
    APILine 17, 4, lw - 17, 4, &HEBEBEB
    APILine 17, 5, lw - 17, 5, &HEBE8EA
    APILine 17, 6, lw - 17, 6, &HEBE8EA
    APILine 17, 7, lw - 17, 7, &HEAEAEA
    APILine 17, 8, lw - 17, 8, &HE1E1E1
    APILine 17, 9, lw - 17, 9, &HE5E2E3
    APILine 17, 10, lw - 17, 10, &HEAE7E8
    APILine 17, 11, lw - 17, 11, &HE8EBE9
    tmph = lh - 22
    APILine 17, tmph + 11, lw - 17, tmph + 11, &HE8EBE9
    APILine 17, tmph + 12, lw - 17, tmph + 12, &HEEF1EF
    APILine 17, tmph + 13, lw - 17, tmph + 13, &HF2F2F2
    APILine 17, tmph + 14, lw - 17, tmph + 14, &HF5F5F5
    APILine 17, tmph + 15, lw - 17, tmph + 15, &HFFFEFE
    APILine 17, tmph + 16, lw - 17, tmph + 16, &HFFFFFF
    APILine 17, tmph + 17, lw - 17, tmph + 17, &HFDFDFD
    APILine 17, tmph + 18, lw - 17, tmph + 18, &HFEFEFE
    APILine 17, tmph + 19, lw - 17, tmph + 19, &HFBFBFB
    APILine 17, tmph + 20, lw - 17, tmph + 20, &H545454
    APILine 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    APILine 17, tmph + 22, lw - 17, tmph + 22, &HECECEC

End Sub

Private Sub DrawMacOSXButtonHot()


Dim lHDC  As Long
Dim lh    As Long
Dim lw    As Long
Dim tmph  As Long
Dim tmpw  As Long
Dim tmph1 As Long


    With UserControl
        lHDC = .hdc
        lh = .ScaleHeight
        lw = .ScaleWidth
    End With
    APIFillRectByCoords hdc, 18, 11, lw - 34, lh - 19, &HE2A66A
    SetPixel lHDC, 6, 0, &HFEFEFE
    SetPixel lHDC, 7, 0, &HE6E5E5
    SetPixel lHDC, 8, 0, &HA9A5A5
    SetPixel lHDC, 9, 0, &H6C5E5E
    SetPixel lHDC, 10, 0, &H482729
    SetPixel lHDC, 11, 0, &H370D0C
    SetPixel lHDC, 12, 0, &H370706
    SetPixel lHDC, 13, 0, &H360605
    SetPixel lHDC, 14, 0, &H3A0606
    SetPixel lHDC, 15, 0, &H410807
    SetPixel lHDC, 16, 0, &H450707
    SetPixel lHDC, 17, 0, &H450608
    SetPixel lHDC, 5, 1, &HF0EFEF
    SetPixel lHDC, 6, 1, &HA38A8C
    SetPixel lHDC, 7, 1, &H6E342F
    SetPixel lHDC, 8, 1, &H661F1A
    SetPixel lHDC, 9, 1, &H9B6A63
    SetPixel lHDC, 10, 1, &HC9A29D
    SetPixel lHDC, 11, 1, &HE2BFBD
    SetPixel lHDC, 12, 1, &HE8C9C6
    SetPixel lHDC, 13, 1, &HEFD3CC
    SetPixel lHDC, 14, 1, &HEFD3CC
    SetPixel lHDC, 15, 1, &HF0D5C9
    SetPixel lHDC, 16, 1, &HF0D5C9
    SetPixel lHDC, 17, 1, &HF1D4C9
    SetPixel lHDC, 3, 2, &HFEFEFE
    SetPixel lHDC, 4, 2, &HE5E5E5
    SetPixel lHDC, 5, 2, &H755E5E
    SetPixel lHDC, 6, 2, &H41070C
    SetPixel lHDC, 7, 2, &H7F2D28
    SetPixel lHDC, 8, 2, &HEC9892
    SetPixel lHDC, 9, 2, &HECB6AF
    SetPixel lHDC, 10, 2, &HE3BBB6
    SetPixel lHDC, 11, 2, &HE3C0BD
    SetPixel lHDC, 12, 2, &HE1C2BF
    SetPixel lHDC, 13, 2, &HDFC3BC
    SetPixel lHDC, 14, 2, &HDFC3BC
    SetPixel lHDC, 15, 2, &HE4C9BD
    SetPixel lHDC, 16, 2, &HE4C9BD
    SetPixel lHDC, 17, 2, &HE5C8BD
    SetPixel lHDC, 3, 3, &HEEEEEE
    SetPixel lHDC, 4, 3, &H8A5A5A
    SetPixel lHDC, 5, 3, &H7A0702
    SetPixel lHDC, 6, 3, &H901501
    SetPixel lHDC, 7, 3, &HC38365
    SetPixel lHDC, 8, 3, &HE3B08F
    SetPixel lHDC, 9, 3, &HE1B394
    SetPixel lHDC, 10, 3, &HE5B798
    SetPixel lHDC, 11, 3, &HE6BC99
    SetPixel lHDC, 12, 3, &HE7BD9A
    SetPixel lHDC, 13, 3, &HE4BC99
    SetPixel lHDC, 14, 3, &HE7BF9C
    SetPixel lHDC, 15, 3, &HE9C1A1
    SetPixel lHDC, 16, 3, &HE8C0A1
    SetPixel lHDC, 17, 3, &HE8C0A1
    SetPixel lHDC, 2, 4, &HFBFBFB
    SetPixel lHDC, 3, 4, &H897879
    SetPixel lHDC, 4, 4, &H4D0909
    SetPixel lHDC, 5, 4, &H951905
    SetPixel lHDC, 6, 4, &HBF422E
    SetPixel lHDC, 7, 4, &HD49475
    SetPixel lHDC, 8, 4, &HD7A483
    SetPixel lHDC, 9, 4, &HDAAC8D
    SetPixel lHDC, 10, 4, &HDBAD8E
    SetPixel lHDC, 11, 4, &HD9AF8C
    SetPixel lHDC, 12, 4, &HDCB28F
    SetPixel lHDC, 13, 4, &HDDB592
    SetPixel lHDC, 14, 4, &HDCB491
    SetPixel lHDC, 15, 4, &HDFB797
    SetPixel lHDC, 16, 4, &HE0B898
    SetPixel lHDC, 17, 4, &HE0B898
    SetPixel lHDC, 1, 5, &HFEFEFE
    SetPixel lHDC, 2, 5, &HCDC9C9
    SetPixel lHDC, 3, 5, &H882517
    SetPixel lHDC, 4, 5, &H922100
    SetPixel lHDC, 5, 5, &HA13A00
    SetPixel lHDC, 6, 5, &HD57333
    SetPixel lHDC, 7, 5, &HDFA36F
    SetPixel lHDC, 8, 5, &HDDA876
    SetPixel lHDC, 9, 5, &HD8A573
    SetPixel lHDC, 10, 5, &HDFAE80
    SetPixel lHDC, 11, 5, &HDBAD7D
    SetPixel lHDC, 12, 5, &HDFB084
    SetPixel lHDC, 13, 5, &HDFB286
    SetPixel lHDC, 14, 5, &HDFB188
    SetPixel lHDC, 15, 5, &HE1B58D
    SetPixel lHDC, 16, 5, &HE3B58E
    SetPixel lHDC, 17, 5, &HE3B48E
    SetPixel lHDC, 1, 6, &HF9F9F9
    SetPixel lHDC, 2, 6, &H7B706E
    SetPixel lHDC, 3, 6, &H871405
    SetPixel lHDC, 4, 6, &HA5330E
    SetPixel lHDC, 5, 6, &HB34C0D
    SetPixel lHDC, 6, 6, &HD27030
    SetPixel lHDC, 7, 6, &HD89C68
    SetPixel lHDC, 8, 6, &HDAA573
    SetPixel lHDC, 9, 6, &HD9A674
    SetPixel lHDC, 10, 6, &HD9A87A
    SetPixel lHDC, 11, 6, &HDBAD7D
    SetPixel lHDC, 12, 6, &HDBAC80
    SetPixel lHDC, 13, 6, &HDCAF83
    SetPixel lHDC, 14, 6, &HDFB188
    SetPixel lHDC, 15, 6, &HDEB28A
    SetPixel lHDC, 16, 6, &HDFB18A
    SetPixel lHDC, 17, 6, &HE0B18B
    SetPixel lHDC, 1, 7, &HE8E8E7
    SetPixel lHDC, 2, 7, &H773F34
    SetPixel lHDC, 3, 7, &H9F2C00
    SetPixel lHDC, 4, 7, &HBA4B07
    SetPixel lHDC, 5, 7, &HC35E10
    SetPixel lHDC, 6, 7, &HCC7323
    SetPixel lHDC, 7, 7, &HDB8F46
    SetPixel lHDC, 8, 7, &HE8A763
    SetPixel lHDC, 9, 7, &HE3A76C
    SetPixel lHDC, 10, 7, &HE7AB70
    SetPixel lHDC, 11, 7, &HE8AE73
    SetPixel lHDC, 12, 7, &HE8AE73
    SetPixel lHDC, 13, 7, &HEDB17B
    SetPixel lHDC, 14, 7, &HEFB37D
    SetPixel lHDC, 15, 7, &HE9B57E
    SetPixel lHDC, 16, 7, &HE9B57E
    SetPixel lHDC, 17, 7, &HE9B47F
    SetPixel lHDC, 0, 8, &HFDFDFD
    SetPixel lHDC, 1, 8, &HCAC5C5
    SetPixel lHDC, 2, 8, &H682A1F
    SetPixel lHDC, 3, 8, &HB23E0C
    SetPixel lHDC, 4, 8, &HCC5D19
    SetPixel lHDC, 5, 8, &HCE691B
    SetPixel lHDC, 6, 8, &HCE7525
    SetPixel lHDC, 7, 8, &HCD8138
    SetPixel lHDC, 8, 8, &HC58440
    SetPixel lHDC, 9, 8, &HC5894E
    SetPixel lHDC, 10, 8, &HC98D52
    SetPixel lHDC, 11, 8, &HC88E53
    SetPixel lHDC, 12, 8, &HCC9257
    SetPixel lHDC, 13, 8, &HCF935D
    SetPixel lHDC, 14, 8, &HD0945E
    SetPixel lHDC, 15, 8, &HCE9963
    SetPixel lHDC, 16, 8, &HCE9963
    SetPixel lHDC, 17, 8, &HCE9963
    SetPixel lHDC, 0, 9, &HFAFAFA
    SetPixel lHDC, 1, 9, &HB9ADAB
    SetPixel lHDC, 2, 9, &H6E2B10
    SetPixel lHDC, 3, 9, &HB6580D
    SetPixel lHDC, 4, 9, &HCA6C20
    SetPixel lHDC, 5, 9, &HCE792B
    SetPixel lHDC, 6, 9, &HCE8132
    SetPixel lHDC, 7, 9, &HD08B42
    SetPixel lHDC, 8, 9, &HD3904B
    SetPixel lHDC, 9, 9, &HD3934C
    SetPixel lHDC, 10, 9, &HD89753
    SetPixel lHDC, 11, 9, &HDB9B5A
    SetPixel lHDC, 12, 9, &HDC9B5E
    SetPixel lHDC, 13, 9, &HDB9C60
    SetPixel lHDC, 14, 9, &HDB9C60
    SetPixel lHDC, 15, 9, &HDDA164
    SetPixel lHDC, 16, 9, &HDDA164
    SetPixel lHDC, 17, 9, &HDDA064
    SetPixel lHDC, 0, 10, &HF7F7F7
    SetPixel lHDC, 1, 10, &HB0A09E
    SetPixel lHDC, 2, 10, &H712E13
    SetPixel lHDC, 3, 10, &HBD5F14
    SetPixel lHDC, 4, 10, &HD17327
    SetPixel lHDC, 5, 10, &HD47F31
    SetPixel lHDC, 6, 10, &HD98C3D
    SetPixel lHDC, 7, 10, &HD9944B
    SetPixel lHDC, 8, 10, &HD7944F
    SetPixel lHDC, 9, 10, &HDC9C55
    SetPixel lHDC, 10, 10, &HDC9B57
    SetPixel lHDC, 11, 10, &HE3A362
    SetPixel lHDC, 12, 10, &HE3A265
    SetPixel lHDC, 13, 10, &HE2A367
    SetPixel lHDC, 14, 10, &HE0A165
    SetPixel lHDC, 15, 10, &HE3A66A
    SetPixel lHDC, 16, 10, &HE3A66A
    SetPixel lHDC, 17, 10, &HE2A66A
    tmph = lh - 22
    SetPixel lHDC, 0, tmph + 10, &HF7F7F7
    SetPixel lHDC, 1, tmph + 10, &HB0A09E
    SetPixel lHDC, 2, tmph + 10, &H712E13
    SetPixel lHDC, 3, tmph + 10, &HBD5F14
    SetPixel lHDC, 4, tmph + 10, &HD17327
    SetPixel lHDC, 5, tmph + 10, &HD47F31
    SetPixel lHDC, 6, tmph + 10, &HD98C3D
    SetPixel lHDC, 7, tmph + 10, &HD9944B
    SetPixel lHDC, 8, tmph + 10, &HD7944F
    SetPixel lHDC, 9, tmph + 10, &HDC9C55
    SetPixel lHDC, 10, tmph + 10, &HDC9B57
    SetPixel lHDC, 11, tmph + 10, &HE3A362
    SetPixel lHDC, 12, tmph + 10, &HE3A265
    SetPixel lHDC, 13, tmph + 10, &HE2A367
    SetPixel lHDC, 14, tmph + 10, &HE0A165
    SetPixel lHDC, 15, tmph + 10, &HE3A66A
    SetPixel lHDC, 16, tmph + 10, &HE3A66A
    SetPixel lHDC, 17, tmph + 10, &HE2A66A
    SetPixel lHDC, 0, tmph + 11, &HF5F5F5
    SetPixel lHDC, 1, tmph + 11, &HACA39E
    SetPixel lHDC, 2, tmph + 11, &H744421
    SetPixel lHDC, 3, tmph + 11, &HC56F1F
    SetPixel lHDC, 4, tmph + 11, &HD17A2A
    SetPixel lHDC, 5, tmph + 11, &HD58C42
    SetPixel lHDC, 6, tmph + 11, &HD7914B
    SetPixel lHDC, 7, tmph + 11, &HDF9854
    SetPixel lHDC, 8, tmph + 11, &HE4A05F
    SetPixel lHDC, 9, tmph + 11, &HE29F66
    SetPixel lHDC, 10, tmph + 11, &HE4A56B
    SetPixel lHDC, 11, tmph + 11, &HDDA467
    SetPixel lHDC, 12, tmph + 11, &HE0A76A
    SetPixel lHDC, 13, tmph + 11, &HE2A96C
    SetPixel lHDC, 14, tmph + 11, &HE3A870
    SetPixel lHDC, 15, tmph + 11, &HE6AC76
    SetPixel lHDC, 16, tmph + 11, &HE6AC76
    SetPixel lHDC, 17, tmph + 11, &HE6AC76
    SetPixel lHDC, 0, tmph + 12, &HF5F5F5
    SetPixel lHDC, 1, tmph + 12, &HB1AAA7
    SetPixel lHDC, 2, tmph + 12, &H825533
    SetPixel lHDC, 3, tmph + 12, &HCF792A
    SetPixel lHDC, 4, tmph + 12, &HE48D3D
    SetPixel lHDC, 5, tmph + 12, &HDD944A
    SetPixel lHDC, 6, tmph + 12, &HE49E58
    SetPixel lHDC, 7, tmph + 12, &HEBA460
    SetPixel lHDC, 8, tmph + 12, &HEEAA69
    SetPixel lHDC, 9, tmph + 12, &HF3B077
    SetPixel lHDC, 10, tmph + 12, &HEEAF75
    SetPixel lHDC, 11, tmph + 12, &HEBB275
    SetPixel lHDC, 12, tmph + 12, &HEFB679
    SetPixel lHDC, 13, tmph + 12, &HF1B87B
    SetPixel lHDC, 14, tmph + 12, &HF1B67E
    SetPixel lHDC, 15, tmph + 12, &HF2B781
    SetPixel lHDC, 16, tmph + 12, &HF1B681
    SetPixel lHDC, 17, tmph + 12, &HF1B681
    SetPixel lHDC, 0, tmph + 13, &HF7F7F7
    SetPixel lHDC, 1, tmph + 13, &HC2C2C1
    SetPixel lHDC, 2, tmph + 13, &H6B5D4E
    SetPixel lHDC, 3, tmph + 13, &HC27831
    SetPixel lHDC, 4, tmph + 13, &HDA8E46
    SetPixel lHDC, 5, tmph + 13, &HE7A05C
    SetPixel lHDC, 6, tmph + 13, &HEAA665
    SetPixel lHDC, 7, tmph + 13, &HE9AF6E
    SetPixel lHDC, 8, tmph + 13, &HEFB377
    SetPixel lHDC, 9, tmph + 13, &HF3B579
    SetPixel lHDC, 10, tmph + 13, &HF7B97D
    SetPixel lHDC, 11, tmph + 13, &HF2BB7E
    SetPixel lHDC, 12, tmph + 13, &HF4BB83
    SetPixel lHDC, 13, tmph + 13, &HF5BE85
    SetPixel lHDC, 14, tmph + 13, &HF4BB87
    SetPixel lHDC, 15, tmph + 13, &HF5BE8A
    SetPixel lHDC, 16, tmph + 13, &HF5BD8A
    SetPixel lHDC, 17, tmph + 13, &HF3BD8A
    SetPixel lHDC, 0, tmph + 14, &HFBFBFB
    SetPixel lHDC, 1, tmph + 14, &HE1E1E1
    SetPixel lHDC, 2, tmph + 14, &H85796E
    SetPixel lHDC, 3, tmph + 14, &HB76F2B
    SetPixel lHDC, 4, tmph + 14, &HDE924A
    SetPixel lHDC, 5, tmph + 14, &HE8A15D
    SetPixel lHDC, 6, tmph + 14, &HF2AE6D
    SetPixel lHDC, 7, tmph + 14, &HF1B776
    SetPixel lHDC, 8, tmph + 14, &HF2B67A
    SetPixel lHDC, 9, tmph + 14, &HFBBD81
    SetPixel lHDC, 10, tmph + 14, &HFFC286
    SetPixel lHDC, 11, tmph + 14, &HFAC386
    SetPixel lHDC, 12, tmph + 14, &HFBC28A
    SetPixel lHDC, 13, tmph + 14, &HFAC38A
    SetPixel lHDC, 14, tmph + 14, &HFAC18D
    SetPixel lHDC, 15, tmph + 14, &HFDC592
    SetPixel lHDC, 16, tmph + 14, &HFDC592
    SetPixel lHDC, 17, tmph + 14, &HFCC592
    SetPixel lHDC, 0, tmph + 15, &HFEFEFE
    SetPixel lHDC, 1, tmph + 15, &HEDEDED
    SetPixel lHDC, 2, tmph + 15, &HA2A0A0
    SetPixel lHDC, 3, tmph + 15, &H816753
    SetPixel lHDC, 4, tmph + 15, &HC09068
    SetPixel lHDC, 5, tmph + 15, &HEDA55F
    SetPixel lHDC, 6, tmph + 15, &HFAB26C
    SetPixel lHDC, 7, tmph + 15, &HFCBF7D
    SetPixel lHDC, 8, tmph + 15, &HF7C182
    SetPixel lHDC, 9, tmph + 15, &HF8C38A
    SetPixel lHDC, 10, tmph + 15, &HFACA90
    SetPixel lHDC, 11, tmph + 15, &HF7CB8E
    SetPixel lHDC, 12, tmph + 15, &HF8CC8F
    SetPixel lHDC, 13, tmph + 15, &HFACC96
    SetPixel lHDC, 14, tmph + 15, &HF9CB95
    SetPixel lHDC, 15, tmph + 15, &HF9CE97
    SetPixel lHDC, 16, tmph + 15, &HF8CD97
    SetPixel lHDC, 17, tmph + 15, &HF8CE97
    SetPixel lHDC, 1, tmph + 16, &HF6F6F6
    SetPixel lHDC, 2, tmph + 16, &HD6D6D6
    SetPixel lHDC, 3, tmph + 16, &H8E7C6F
    SetPixel lHDC, 4, tmph + 16, &H946843
    SetPixel lHDC, 5, tmph + 16, &HEEA762
    SetPixel lHDC, 6, tmph + 16, &HFFB771
    SetPixel lHDC, 7, tmph + 16, &HFEC17F
    SetPixel lHDC, 8, tmph + 16, &HFFC98A
    SetPixel lHDC, 9, tmph + 16, &HFFCE95
    SetPixel lHDC, 10, tmph + 16, &HFBCB91
    SetPixel lHDC, 11, tmph + 16, &HFFD396
    SetPixel lHDC, 12, tmph + 16, &HFFD396
    SetPixel lHDC, 13, tmph + 16, &HFFD29C
    SetPixel lHDC, 14, tmph + 16, &HFFD39D
    SetPixel lHDC, 15, tmph + 16, &HFFD49E
    SetPixel lHDC, 16, tmph + 16, &HFFD49E
    SetPixel lHDC, 17, tmph + 16, &HFED59E
    SetPixel lHDC, 1, tmph + 17, &HFDFDFD
    SetPixel lHDC, 2, tmph + 17, &HEDEDED
    SetPixel lHDC, 3, tmph + 17, &HBEBEBE
    SetPixel lHDC, 4, tmph + 17, &H6C6C6C
    SetPixel lHDC, 5, tmph + 17, &H7C684F
    SetPixel lHDC, 6, tmph + 17, &HD1AE81
    SetPixel lHDC, 7, tmph + 17, &HF1C284
    SetPixel lHDC, 8, tmph + 17, &HFDCE90
    SetPixel lHDC, 9, tmph + 17, &HF8D193
    SetPixel lHDC, 10, tmph + 17, &HFBD899
    SetPixel lHDC, 11, tmph + 17, &HF5DC9E
    SetPixel lHDC, 12, tmph + 17, &HF8DFA1
    SetPixel lHDC, 13, tmph + 17, &HF8DFA1
    SetPixel lHDC, 14, tmph + 17, &HF8DFA1
    SetPixel lHDC, 15, tmph + 17, &HF8DEA3
    SetPixel lHDC, 16, tmph + 17, &HF7DDA3
    SetPixel lHDC, 17, tmph + 17, &HF7DDA3
    SetPixel lHDC, 2, tmph + 18, &HF9F9F9
    SetPixel lHDC, 3, tmph + 18, &HE6E6E6
    SetPixel lHDC, 4, tmph + 18, &HBABABA
    SetPixel lHDC, 5, tmph + 18, &H827666
    SetPixel lHDC, 6, tmph + 18, &H836743
    SetPixel lHDC, 7, tmph + 18, &HBE935B
    SetPixel lHDC, 8, tmph + 18, &HF4C78B
    SetPixel lHDC, 9, tmph + 18, &HFDD79A
    SetPixel lHDC, 10, tmph + 18, &HFFDFA0
    SetPixel lHDC, 11, tmph + 18, &HFBE2A4
    SetPixel lHDC, 12, tmph + 18, &HFFE7A9
    SetPixel lHDC, 13, tmph + 18, &HFFE9AB
    SetPixel lHDC, 14, tmph + 18, &HFFE7A9
    SetPixel lHDC, 15, tmph + 18, &HFFE6AC
    SetPixel lHDC, 16, tmph + 18, &HFFE6AD
    SetPixel lHDC, 17, tmph + 18, &HFFE6AD
    SetPixel lHDC, 2, tmph + 19, &HFEFEFE
    SetPixel lHDC, 3, tmph + 19, &HF8F8F8
    SetPixel lHDC, 4, tmph + 19, &HE6E6E6
    SetPixel lHDC, 5, tmph + 19, &HC8C8C8
    SetPixel lHDC, 6, tmph + 19, &H8F8F8F
    SetPixel lHDC, 7, tmph + 19, &H686462
    SetPixel lHDC, 8, tmph + 19, &H6D655E
    SetPixel lHDC, 9, tmph + 19, &H918472
    SetPixel lHDC, 10, tmph + 19, &HB3A88E
    SetPixel lHDC, 11, tmph + 19, &HDAD1B2
    SetPixel lHDC, 12, tmph + 19, &HE3DBBA
    SetPixel lHDC, 13, tmph + 19, &HE7E0C0
    SetPixel lHDC, 14, tmph + 19, &HE9E2C1
    SetPixel lHDC, 15, tmph + 19, &HE9E2C5
    SetPixel lHDC, 16, tmph + 19, &HE9E1C5
    SetPixel lHDC, 17, tmph + 19, &HE9E2C5
    SetPixel lHDC, 3, tmph + 20, &HFEFEFE
    SetPixel lHDC, 4, tmph + 20, &HF9F9F9
    SetPixel lHDC, 5, tmph + 20, &HECECEC
    SetPixel lHDC, 6, tmph + 20, &HDADADA
    SetPixel lHDC, 7, tmph + 20, &HC2C2C1
    SetPixel lHDC, 8, tmph + 20, &H9F9D9B
    SetPixel lHDC, 9, tmph + 20, &H827D75
    SetPixel lHDC, 10, tmph + 20, &H6A6353
    SetPixel lHDC, 11, tmph + 20, &H5F5941
    SetPixel lHDC, 12, tmph + 20, &H5D553B
    SetPixel lHDC, 13, tmph + 20, &H595338
    SetPixel lHDC, 14, tmph + 20, &H5E5739
    SetPixel lHDC, 15, tmph + 20, &H5F5A3C
    SetPixel lHDC, 16, tmph + 20, &H635E3F
    SetPixel lHDC, 17, tmph + 20, &H635D40
    SetPixel lHDC, 5, tmph + 21, &HFCFCFC
    SetPixel lHDC, 6, tmph + 21, &HF5F5F5
    SetPixel lHDC, 7, tmph + 21, &HEBEBEB
    SetPixel lHDC, 8, tmph + 21, &HE1E1E1
    SetPixel lHDC, 9, tmph + 21, &HD6D6D6
    SetPixel lHDC, 10, tmph + 21, &HCECECE
    SetPixel lHDC, 11, tmph + 21, &HC9C9C9
    SetPixel lHDC, 12, tmph + 21, &HC7C7C7
    SetPixel lHDC, 13, tmph + 21, &HC7C7C7
    SetPixel lHDC, 14, tmph + 21, &HC6C6C6
    SetPixel lHDC, 15, tmph + 21, &HC6C6C6
    SetPixel lHDC, 16, tmph + 21, &HC5C5C5
    SetPixel lHDC, 17, tmph + 21, &HC5C5C5
    SetPixel lHDC, 7, tmph + 22, &HFDFDFD
    SetPixel lHDC, 8, tmph + 22, &HF9F9F9
    SetPixel lHDC, 9, tmph + 22, &HF4F4F4
    SetPixel lHDC, 10, tmph + 22, &HF0F0F0
    SetPixel lHDC, 11, tmph + 22, &HEEEEEE
    SetPixel lHDC, 12, tmph + 22, &HEDEDED
    SetPixel lHDC, 13, tmph + 22, &HECECEC
    SetPixel lHDC, 14, tmph + 22, &HECECEC
    SetPixel lHDC, 15, tmph + 22, &HECECEC
    SetPixel lHDC, 16, tmph + 22, &HECECEC
    SetPixel lHDC, 17, tmph + 22, &HECECEC

    tmpw = lw - 34
    SetPixel lHDC, tmpw + 17, 0, &H450608
    SetPixel lHDC, tmpw + 18, 0, &H450608
    SetPixel lHDC, tmpw + 19, 0, &H3B0707
    SetPixel lHDC, tmpw + 20, 0, &H370706
    SetPixel lHDC, tmpw + 21, 0, &H360507
    SetPixel lHDC, tmpw + 22, 0, &H3B0F10
    SetPixel lHDC, tmpw + 23, 0, &H442526
    SetPixel lHDC, tmpw + 24, 0, &H604E4E
    SetPixel lHDC, tmpw + 25, 0, &HA29D9E
    SetPixel lHDC, tmpw + 26, 0, &HEEEEEE
    SetPixel lHDC, tmpw + 34, 0, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 1, &HF1D4C9
    SetPixel lHDC, tmpw + 18, 1, &HF1D4C9
    SetPixel lHDC, tmpw + 19, 1, &HEDD3CD
    SetPixel lHDC, tmpw + 20, 1, &HEBD1CB
    SetPixel lHDC, tmpw + 21, 1, &HE9CEC4
    SetPixel lHDC, tmpw + 22, 1, &HE5C1B9
    SetPixel lHDC, tmpw + 23, 1, &HCFA89F
    SetPixel lHDC, tmpw + 24, 1, &HAA6E68
    SetPixel lHDC, tmpw + 25, 1, &H73211B
    SetPixel lHDC, tmpw + 26, 1, &H702924
    SetPixel lHDC, tmpw + 27, 1, &HAA9897
    SetPixel lHDC, tmpw + 28, 1, &HF7F7F7
    SetPixel lHDC, tmpw + 34, 1, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 2, &HE5C8BD
    SetPixel lHDC, tmpw + 18, 2, &HE5C8BD
    SetPixel lHDC, tmpw + 19, 2, &HDEC4BE
    SetPixel lHDC, tmpw + 20, 2, &HDCC2BC
    SetPixel lHDC, tmpw + 21, 2, &HE2C7BD
    SetPixel lHDC, tmpw + 22, 2, &HE2BEB6
    SetPixel lHDC, tmpw + 23, 2, &HE8C1B8
    SetPixel lHDC, tmpw + 24, 2, &HF0B4AE
    SetPixel lHDC, tmpw + 25, 2, &HF29C96
    SetPixel lHDC, tmpw + 26, 2, &H822D27
    SetPixel lHDC, tmpw + 27, 2, &H400807
    SetPixel lHDC, tmpw + 28, 2, &H71585A
    SetPixel lHDC, tmpw + 29, 2, &HEBEBEB
    SetPixel lHDC, tmpw + 34, 2, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 3, &HE8C0A1
    SetPixel lHDC, tmpw + 18, 3, &HE8C0A1
    SetPixel lHDC, tmpw + 19, 3, &HE5C09A
    SetPixel lHDC, tmpw + 20, 3, &HE4BF99
    SetPixel lHDC, tmpw + 21, 3, &HE4BA97
    SetPixel lHDC, tmpw + 22, 3, &HE9BF9C
    SetPixel lHDC, tmpw + 23, 3, &HDFB695
    SetPixel lHDC, tmpw + 24, 3, &HDFB695
    SetPixel lHDC, tmpw + 25, 3, &HE0AE90
    SetPixel lHDC, tmpw + 26, 3, &HCB8469
    SetPixel lHDC, tmpw + 27, 3, &H941600
    SetPixel lHDC, tmpw + 28, 3, &H830800
    SetPixel lHDC, tmpw + 29, 3, &H895253
    SetPixel lHDC, tmpw + 30, 3, &HF0EFEF
    SetPixel lHDC, tmpw + 34, 3, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 4, &HE0B898
    SetPixel lHDC, tmpw + 18, 4, &HE0B897
    SetPixel lHDC, tmpw + 19, 4, &HDAB58F
    SetPixel lHDC, tmpw + 20, 4, &HDBB690
    SetPixel lHDC, tmpw + 21, 4, &HDBB18E
    SetPixel lHDC, tmpw + 22, 4, &HD7AD8A
    SetPixel lHDC, tmpw + 23, 4, &HDAB190
    SetPixel lHDC, tmpw + 24, 4, &HD2A988
    SetPixel lHDC, tmpw + 25, 4, &HD6A486
    SetPixel lHDC, tmpw + 26, 4, &HDA9378
    SetPixel lHDC, tmpw + 27, 4, &HBF4129
    SetPixel lHDC, tmpw + 28, 4, &H991B03
    SetPixel lHDC, tmpw + 29, 4, &H500709
    SetPixel lHDC, tmpw + 30, 4, &H826F70
    SetPixel lHDC, tmpw + 31, 4, &HFAFAFA
    SetPixel lHDC, tmpw + 34, 4, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 5, &HE3B48E
    SetPixel lHDC, tmpw + 18, 5, &HE3B48D
    SetPixel lHDC, tmpw + 19, 5, &HE0B387
    SetPixel lHDC, tmpw + 20, 5, &HDEB185
    SetPixel lHDC, tmpw + 21, 5, &HE1B084
    SetPixel lHDC, tmpw + 22, 5, &HE3AE83
    SetPixel lHDC, tmpw + 23, 5, &HE1AF7B
    SetPixel lHDC, tmpw + 24, 5, &HE0A976
    SetPixel lHDC, tmpw + 25, 5, &HDCA473
    SetPixel lHDC, tmpw + 26, 5, &HDEA372
    SetPixel lHDC, tmpw + 27, 5, &HCC712E
    SetPixel lHDC, tmpw + 28, 5, &HA53900
    SetPixel lHDC, tmpw + 29, 5, &H9D2200
    SetPixel lHDC, tmpw + 30, 5, &H9E2114
    SetPixel lHDC, tmpw + 31, 5, &HC7C5C4
    SetPixel lHDC, tmpw + 32, 5, &HFEFEFE
    SetPixel lHDC, tmpw + 34, 5, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 6, &HE0B18B
    SetPixel lHDC, tmpw + 18, 6, &HE0B18A
    SetPixel lHDC, tmpw + 19, 6, &HDEB185
    SetPixel lHDC, tmpw + 20, 6, &HDEB185
    SetPixel lHDC, tmpw + 21, 6, &HDCAB7F
    SetPixel lHDC, tmpw + 22, 6, &HE1AC81
    SetPixel lHDC, tmpw + 23, 6, &HDCAA76
    SetPixel lHDC, tmpw + 24, 6, &HDCA572
    SetPixel lHDC, tmpw + 25, 6, &HDBA372
    SetPixel lHDC, tmpw + 26, 6, &HD79C6B
    SetPixel lHDC, tmpw + 27, 6, &HD17633
    SetPixel lHDC, tmpw + 28, 6, &HB74B0B
    SetPixel lHDC, tmpw + 29, 6, &HAC310D
    SetPixel lHDC, tmpw + 30, 6, &H961507
    SetPixel lHDC, tmpw + 31, 6, &H736D6A
    SetPixel lHDC, tmpw + 32, 6, &HF8F8F8
    SetPixel lHDC, tmpw + 34, 6, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 7, &HE9B47F
    SetPixel lHDC, tmpw + 18, 7, &HEAB47E
    SetPixel lHDC, tmpw + 19, 7, &HEFB67E
    SetPixel lHDC, tmpw + 20, 7, &HE8AF77
    SetPixel lHDC, tmpw + 21, 7, &HE7AF74
    SetPixel lHDC, tmpw + 22, 7, &HE4AC71
    SetPixel lHDC, tmpw + 23, 7, &HEAAD6F
    SetPixel lHDC, tmpw + 24, 7, &HE9A968
    SetPixel lHDC, tmpw + 25, 7, &HE7A564
    SetPixel lHDC, tmpw + 26, 7, &HD9904C
    SetPixel lHDC, tmpw + 27, 7, &HC5711F
    SetPixel lHDC, tmpw + 28, 7, &HC16010
    SetPixel lHDC, tmpw + 29, 7, &HBB4D05
    SetPixel lHDC, tmpw + 30, 7, &HA02D00
    SetPixel lHDC, tmpw + 31, 7, &H774033
    SetPixel lHDC, tmpw + 32, 7, &HE7E6E6
    SetPixel lHDC, tmpw + 34, 7, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 8, &HCE9963
    SetPixel lHDC, tmpw + 18, 8, &HCF9963
    SetPixel lHDC, tmpw + 19, 8, &HCE955D
    SetPixel lHDC, tmpw + 20, 8, &HCE955D
    SetPixel lHDC, tmpw + 21, 8, &HCA9257
    SetPixel lHDC, tmpw + 22, 8, &HC89055
    SetPixel lHDC, tmpw + 23, 8, &HCB8E50
    SetPixel lHDC, tmpw + 24, 8, &HCB8B4A
    SetPixel lHDC, tmpw + 25, 8, &HC58342
    SetPixel lHDC, tmpw + 26, 8, &HC87F3B
    SetPixel lHDC, tmpw + 27, 8, &HCA7624
    SetPixel lHDC, tmpw + 28, 8, &HCA6919
    SetPixel lHDC, tmpw + 29, 8, &HCC5E16
    SetPixel lHDC, tmpw + 30, 8, &HB23E07
    SetPixel lHDC, tmpw + 31, 8, &H682B1D
    SetPixel lHDC, tmpw + 32, 8, &HC7C2C2
    SetPixel lHDC, tmpw + 33, 8, &HFDFDFD
    SetPixel lHDC, tmpw + 34, 8, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 9, &HDDA064
    SetPixel lHDC, tmpw + 18, 9, &HDCA064
    SetPixel lHDC, tmpw + 19, 9, &HDA9D5D
    SetPixel lHDC, tmpw + 20, 9, &HD99C5C
    SetPixel lHDC, tmpw + 21, 9, &HDA9D5D
    SetPixel lHDC, tmpw + 22, 9, &HDA9A5A
    SetPixel lHDC, tmpw + 23, 9, &HD89753
    SetPixel lHDC, tmpw + 24, 9, &HD7914E
    SetPixel lHDC, tmpw + 25, 9, &HD38E49
    SetPixel lHDC, tmpw + 26, 9, &HD38B43
    SetPixel lHDC, tmpw + 27, 9, &HCD8430
    SetPixel lHDC, tmpw + 28, 9, &HCA7826
    SetPixel lHDC, tmpw + 29, 9, &HCE6C1E
    SetPixel lHDC, tmpw + 30, 9, &HB9560C
    SetPixel lHDC, tmpw + 31, 9, &H742E0D
    SetPixel lHDC, tmpw + 32, 9, &HB3A6A4
    SetPixel lHDC, tmpw + 33, 9, &HF9F9F9
    SetPixel lHDC, tmpw + 34, 9, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 10, &HE2A66A
    SetPixel lHDC, tmpw + 18, 10, &HE2A66A
    SetPixel lHDC, tmpw + 19, 10, &HE1A464
    SetPixel lHDC, tmpw + 20, 10, &HE0A363
    SetPixel lHDC, tmpw + 21, 10, &HE0A363
    SetPixel lHDC, tmpw + 22, 10, &HE1A161
    SetPixel lHDC, tmpw + 23, 10, &HE09F5B
    SetPixel lHDC, tmpw + 24, 10, &HDE9855
    SetPixel lHDC, tmpw + 25, 10, &HDC9752
    SetPixel lHDC, tmpw + 26, 10, &HDB934B
    SetPixel lHDC, tmpw + 27, 10, &HD68D39
    SetPixel lHDC, tmpw + 28, 10, &HD17F2D
    SetPixel lHDC, tmpw + 29, 10, &HD67426
    SetPixel lHDC, tmpw + 30, 10, &HC05D13
    SetPixel lHDC, tmpw + 31, 10, &H7C3514
    SetPixel lHDC, tmpw + 32, 10, &HAB9B98
    SetPixel lHDC, tmpw + 33, 10, &HF6F6F6
    SetPixel lHDC, tmpw + 34, 10, &HFFFFFFFF

    tmph = lh - 22
    tmpw = lw - 34
    SetPixel lHDC, tmpw + 17, tmph + 10, &HE2A66A
    SetPixel lHDC, tmpw + 18, tmph + 10, &HE2A66A
    SetPixel lHDC, tmpw + 19, tmph + 10, &HE1A464
    SetPixel lHDC, tmpw + 20, tmph + 10, &HE0A363
    SetPixel lHDC, tmpw + 21, tmph + 10, &HE0A363
    SetPixel lHDC, tmpw + 22, tmph + 10, &HE1A161
    SetPixel lHDC, tmpw + 23, tmph + 10, &HE09F5B
    SetPixel lHDC, tmpw + 24, tmph + 10, &HDE9855
    SetPixel lHDC, tmpw + 25, tmph + 10, &HDC9752
    SetPixel lHDC, tmpw + 26, tmph + 10, &HDB934B
    SetPixel lHDC, tmpw + 27, tmph + 10, &HD68D39
    SetPixel lHDC, tmpw + 28, tmph + 10, &HD17F2D
    SetPixel lHDC, tmpw + 29, tmph + 10, &HD67426
    SetPixel lHDC, tmpw + 30, tmph + 10, &HC05D13
    SetPixel lHDC, tmpw + 31, tmph + 10, &H7C3514
    SetPixel lHDC, tmpw + 32, tmph + 10, &HAB9B98
    SetPixel lHDC, tmpw + 33, tmph + 10, &HF6F6F6
    SetPixel lHDC, tmpw + 17, tmph + 11, &HE6AC76
    SetPixel lHDC, tmpw + 18, tmph + 11, &HE6AC76
    SetPixel lHDC, tmpw + 19, tmph + 11, &HE2A86D
    SetPixel lHDC, tmpw + 20, tmph + 11, &HE5A66C
    SetPixel lHDC, tmpw + 21, tmph + 11, &HE1A56A
    SetPixel lHDC, tmpw + 22, tmph + 11, &HE4A46A
    SetPixel lHDC, tmpw + 23, tmph + 11, &HE1A266
    SetPixel lHDC, tmpw + 24, tmph + 11, &HE6A364
    SetPixel lHDC, tmpw + 25, tmph + 11, &HE19F5E
    SetPixel lHDC, tmpw + 26, tmph + 11, &HDF9A55
    SetPixel lHDC, tmpw + 27, tmph + 11, &HD89048
    SetPixel lHDC, tmpw + 28, tmph + 11, &HD88A3E
    SetPixel lHDC, tmpw + 29, tmph + 11, &HCF7927
    SetPixel lHDC, tmpw + 30, tmph + 11, &HC87220
    SetPixel lHDC, tmpw + 31, tmph + 11, &H77481E
    SetPixel lHDC, tmpw + 32, tmph + 11, &HABA39E
    SetPixel lHDC, tmpw + 33, tmph + 11, &HF4F4F4
    SetPixel lHDC, tmpw + 17, tmph + 12, &HF1B681
    SetPixel lHDC, tmpw + 18, tmph + 12, &HF0B780
    SetPixel lHDC, tmpw + 19, tmph + 12, &HF2B87D
    SetPixel lHDC, tmpw + 20, tmph + 12, &HF5B67C
    SetPixel lHDC, tmpw + 21, tmph + 12, &HF1B57A
    SetPixel lHDC, tmpw + 22, tmph + 12, &HF2B278
    SetPixel lHDC, tmpw + 23, tmph + 12, &HF0B175
    SetPixel lHDC, tmpw + 24, tmph + 12, &HF3B071
    SetPixel lHDC, tmpw + 25, tmph + 12, &HECAA69
    SetPixel lHDC, tmpw + 26, tmph + 12, &HE9A45F
    SetPixel lHDC, tmpw + 27, tmph + 12, &HE8A058
    SetPixel lHDC, tmpw + 28, tmph + 12, &HE5974B
    SetPixel lHDC, tmpw + 29, tmph + 12, &HE38D3B
    SetPixel lHDC, tmpw + 30, tmph + 12, &HD37D2B
    SetPixel lHDC, tmpw + 31, tmph + 12, &H895A32
    SetPixel lHDC, tmpw + 32, tmph + 12, &HB4AFAC
    SetPixel lHDC, tmpw + 33, tmph + 12, &HF5F5F5
    SetPixel lHDC, tmpw + 17, tmph + 13, &HF3BD8A
    SetPixel lHDC, tmpw + 18, tmph + 13, &HF3BD8A
    SetPixel lHDC, tmpw + 19, tmph + 13, &HF2BD84
    SetPixel lHDC, tmpw + 20, tmph + 13, &HF5BC84
    SetPixel lHDC, tmpw + 21, tmph + 13, &HF3BC83
    SetPixel lHDC, tmpw + 22, tmph + 13, &HF4B981
    SetPixel lHDC, tmpw + 23, tmph + 13, &HF2B97C
    SetPixel lHDC, tmpw + 24, tmph + 13, &HF5B77B
    SetPixel lHDC, tmpw + 25, tmph + 13, &HF1B476
    SetPixel lHDC, tmpw + 26, tmph + 13, &HEFAF6E
    SetPixel lHDC, tmpw + 27, tmph + 13, &HE5A45F
    SetPixel lHDC, tmpw + 28, tmph + 13, &HE49F5A
    SetPixel lHDC, tmpw + 29, tmph + 13, &HDA8F4A
    SetPixel lHDC, tmpw + 30, tmph + 13, &HC57A35
    SetPixel lHDC, tmpw + 31, tmph + 13, &H736353
    SetPixel lHDC, tmpw + 32, tmph + 13, &HD6D5D5
    SetPixel lHDC, tmpw + 33, tmph + 13, &HF8F8F8
    SetPixel lHDC, tmpw + 17, tmph + 14, &HFCC592
    SetPixel lHDC, tmpw + 18, tmph + 14, &HFBC592
    SetPixel lHDC, tmpw + 19, tmph + 14, &HF7C289
    SetPixel lHDC, tmpw + 20, tmph + 14, &HFCC38B
    SetPixel lHDC, tmpw + 21, tmph + 14, &HFAC38A
    SetPixel lHDC, tmpw + 22, tmph + 14, &HFDC28A
    SetPixel lHDC, tmpw + 23, tmph + 14, &HFBC285
    SetPixel lHDC, tmpw + 24, tmph + 14, &HFBBD81
    SetPixel lHDC, tmpw + 25, tmph + 14, &HF6B97B
    SetPixel lHDC, tmpw + 26, tmph + 14, &HF6B675
    SetPixel lHDC, tmpw + 27, tmph + 14, &HF0AF6A
    SetPixel lHDC, tmpw + 28, tmph + 14, &HE8A35E
    SetPixel lHDC, tmpw + 29, tmph + 14, &HDD924D
    SetPixel lHDC, tmpw + 30, tmph + 14, &HBA702B
    SetPixel lHDC, tmpw + 31, tmph + 14, &H847A70
    SetPixel lHDC, tmpw + 32, tmph + 14, &HE8E8E8
    SetPixel lHDC, tmpw + 33, tmph + 14, &HFDFDFD
    SetPixel lHDC, tmpw + 17, tmph + 15, &HF8CE97
    SetPixel lHDC, tmpw + 18, tmph + 15, &HF9CD97
    SetPixel lHDC, tmpw + 19, tmph + 15, &HF9CE95
    SetPixel lHDC, tmpw + 20, tmph + 15, &HF7CC93
    SetPixel lHDC, tmpw + 21, tmph + 15, &HF6CB92
    SetPixel lHDC, tmpw + 22, tmph + 15, &HF9CA92
    SetPixel lHDC, tmpw + 23, tmph + 15, &HFCCD90
    SetPixel lHDC, tmpw + 24, tmph + 15, &HF8C488
    SetPixel lHDC, tmpw + 25, tmph + 15, &HF3BD80
    SetPixel lHDC, tmpw + 26, tmph + 15, &HFABD7D
    SetPixel lHDC, tmpw + 27, tmph + 15, &HF7B26D
    SetPixel lHDC, tmpw + 28, tmph + 15, &HEAA560
    SetPixel lHDC, tmpw + 29, tmph + 15, &HC0925D
    SetPixel lHDC, tmpw + 30, tmph + 15, &H896F54
    SetPixel lHDC, tmpw + 31, tmph + 15, &HBABABB
    SetPixel lHDC, tmpw + 32, tmph + 15, &HF1F1F1
    SetPixel lHDC, tmpw + 17, tmph + 16, &HFED59E
    SetPixel lHDC, tmpw + 18, tmph + 16, &HFFD59F
    SetPixel lHDC, tmpw + 19, tmph + 16, &HFED39A
    SetPixel lHDC, tmpw + 20, tmph + 16, &HFFD49B
    SetPixel lHDC, tmpw + 21, tmph + 16, &HFCD198
    SetPixel lHDC, tmpw + 22, tmph + 16, &HFFD098
    SetPixel lHDC, tmpw + 23, tmph + 16, &HFECF92
    SetPixel lHDC, tmpw + 24, tmph + 16, &HFFCB8F
    SetPixel lHDC, tmpw + 25, tmph + 16, &HFFC98C
    SetPixel lHDC, tmpw + 26, tmph + 16, &HFEC181
    SetPixel lHDC, tmpw + 27, tmph + 16, &HFBB671
    SetPixel lHDC, tmpw + 28, tmph + 16, &HF0AB66
    SetPixel lHDC, tmpw + 29, tmph + 16, &H9F733E
    SetPixel lHDC, tmpw + 30, tmph + 16, &H918478
    SetPixel lHDC, tmpw + 31, tmph + 16, &HE2E2E2
    SetPixel lHDC, tmpw + 32, tmph + 16, &HF9F9F9
    SetPixel lHDC, tmpw + 17, tmph + 17, &HF7DDA3
    SetPixel lHDC, tmpw + 18, tmph + 17, &HF8DDA4
    SetPixel lHDC, tmpw + 19, tmph + 17, &HF9E0A2
    SetPixel lHDC, tmpw + 20, tmph + 17, &HF5DC9E
    SetPixel lHDC, tmpw + 21, tmph + 17, &HF8DEA2
    SetPixel lHDC, tmpw + 22, tmph + 17, &HFBDDA2
    SetPixel lHDC, tmpw + 23, tmph + 17, &HF7D495
    SetPixel lHDC, tmpw + 24, tmph + 17, &HF8D193
    SetPixel lHDC, tmpw + 25, tmph + 17, &HFCCD90
    SetPixel lHDC, tmpw + 26, tmph + 17, &HF1C088
    SetPixel lHDC, tmpw + 27, tmph + 17, &HDBB186
    SetPixel lHDC, tmpw + 28, tmph + 17, &H8C7259
    SetPixel lHDC, tmpw + 29, tmph + 17, &H6D6B6B
    SetPixel lHDC, tmpw + 30, tmph + 17, &HD2D2D2
    SetPixel lHDC, tmpw + 31, tmph + 17, &HF2F2F2
    SetPixel lHDC, tmpw + 32, tmph + 17, &HFEFEFE
    SetPixel lHDC, tmpw + 17, tmph + 18, &HFFE6AD
    SetPixel lHDC, tmpw + 18, tmph + 18, &HFFE6AD
    SetPixel lHDC, tmpw + 19, tmph + 18, &HFFE7A9
    SetPixel lHDC, tmpw + 20, tmph + 18, &HFFEAAC
    SetPixel lHDC, tmpw + 21, tmph + 18, &HF7DDA1
    SetPixel lHDC, tmpw + 22, tmph + 18, &HFFE1A6
    SetPixel lHDC, tmpw + 23, tmph + 18, &HFFE1A2
    SetPixel lHDC, tmpw + 24, tmph + 18, &HFED799
    SetPixel lHDC, tmpw + 25, tmph + 18, &HFACC8F
    SetPixel lHDC, tmpw + 26, tmph + 18, &HC99A64
    SetPixel lHDC, tmpw + 27, tmph + 18, &H977048
    SetPixel lHDC, tmpw + 28, tmph + 18, &H817060
    SetPixel lHDC, tmpw + 29, tmph + 18, &HC8C8C8
    SetPixel lHDC, tmpw + 30, tmph + 18, &HEBEBEB
    SetPixel lHDC, tmpw + 31, tmph + 18, &HFCFCFC
    SetPixel lHDC, tmpw + 17, tmph + 19, &HE9E2C5
    SetPixel lHDC, tmpw + 18, tmph + 19, &HE9E2C5
    SetPixel lHDC, tmpw + 19, tmph + 19, &HEAE2C4
    SetPixel lHDC, tmpw + 20, tmph + 19, &HE7DFC1
    SetPixel lHDC, tmpw + 21, tmph + 19, &HEEE4C6
    SetPixel lHDC, tmpw + 22, tmph + 19, &HDBD1B4
    SetPixel lHDC, tmpw + 23, tmph + 19, &HB7AF93
    SetPixel lHDC, tmpw + 24, tmph + 19, &H8D8973
    SetPixel lHDC, tmpw + 25, tmph + 19, &H736D60
    SetPixel lHDC, tmpw + 26, tmph + 19, &H6A6660
    SetPixel lHDC, tmpw + 27, tmph + 19, &H8E9090
    SetPixel lHDC, tmpw + 28, tmph + 19, &HCDCDCD
    SetPixel lHDC, tmpw + 29, tmph + 19, &HE8E8E8
    SetPixel lHDC, tmpw + 30, tmph + 19, &HFAFAFA
    SetPixel lHDC, tmpw + 17, tmph + 20, &H635D40
    SetPixel lHDC, tmpw + 18, tmph + 20, &H615B3F
    SetPixel lHDC, tmpw + 19, tmph + 20, &H60583C
    SetPixel lHDC, tmpw + 20, tmph + 20, &H5D563A
    SetPixel lHDC, tmpw + 21, tmph + 20, &H61583D
    SetPixel lHDC, tmpw + 22, tmph + 20, &H605840
    SetPixel lHDC, tmpw + 23, tmph + 20, &H6A6556
    SetPixel lHDC, tmpw + 24, tmph + 20, &H7F7D75
    SetPixel lHDC, tmpw + 25, tmph + 20, &HA4A3A1
    SetPixel lHDC, tmpw + 26, tmph + 20, &HC5C5C5
    SetPixel lHDC, tmpw + 27, tmph + 20, &HDADADA
    SetPixel lHDC, tmpw + 28, tmph + 20, &HEDEDED
    SetPixel lHDC, tmpw + 29, tmph + 20, &HFAFAFA
    SetPixel lHDC, tmpw + 17, tmph + 21, &HC5C5C5
    SetPixel lHDC, tmpw + 18, tmph + 21, &HC5C5C5
    SetPixel lHDC, tmpw + 19, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 20, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 21, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 22, tmph + 21, &HC9C9C9
    SetPixel lHDC, tmpw + 23, tmph + 21, &HCECECE
    SetPixel lHDC, tmpw + 24, tmph + 21, &HD7D7D7
    SetPixel lHDC, tmpw + 25, tmph + 21, &HE1E1E1
    SetPixel lHDC, tmpw + 26, tmph + 21, &HECECEC
    SetPixel lHDC, tmpw + 27, tmph + 21, &HF6F6F6
    SetPixel lHDC, tmpw + 28, tmph + 21, &HFDFDFD
    SetPixel lHDC, tmpw + 17, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 18, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 19, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 20, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 21, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 22, tmph + 22, &HEDEDED
    SetPixel lHDC, tmpw + 23, tmph + 22, &HF0F0F0
    SetPixel lHDC, tmpw + 24, tmph + 22, &HF4F4F4
    SetPixel lHDC, tmpw + 25, tmph + 22, &HFAFAFA
    SetPixel lHDC, tmpw + 26, tmph + 22, &HFDFDFD

    tmph = 11
    tmph1 = lh - 10
    tmpw = lw - 34

    'Generar lineas intermedias
    APILine 0, tmph, 0, tmph1, &HF7F7F7
    APILine 1, tmph, 1, tmph1, &HB0A09E
    APILine 2, tmph, 2, tmph1, &H712E13
    APILine 3, tmph, 3, tmph1, &HBD5F14
    APILine 4, tmph, 4, tmph1, &HD17327
    APILine 5, tmph, 5, tmph1, &HD47F31
    APILine 6, tmph, 6, tmph1, &HD98C3D
    APILine 7, tmph, 7, tmph1, &HD9944B
    APILine 8, tmph, 8, tmph1, &HD7944F
    APILine 9, tmph, 9, tmph1, &HDC9C55
    APILine 10, tmph, 10, tmph1, &HDC9B57
    APILine 11, tmph, 11, tmph1, &HE3A362
    APILine 12, tmph, 12, tmph1, &HE3A265
    APILine 13, tmph, 13, tmph1, &HE2A367
    APILine 14, tmph, 14, tmph1, &HE0A165
    APILine 15, tmph, 15, tmph1, &HE3A66A
    APILine 16, tmph, 16, tmph1, &HE3A66A
    APILine 17, tmph, 17, tmph1, &HE2A66A
    APILine tmpw + 17, tmph, tmpw + 17, tmph1, &HE2A66A
    APILine tmpw + 18, tmph, tmpw + 18, tmph1, &HE2A66A
    APILine tmpw + 19, tmph, tmpw + 19, tmph1, &HE1A464
    APILine tmpw + 20, tmph, tmpw + 20, tmph1, &HE0A363
    APILine tmpw + 21, tmph, tmpw + 21, tmph1, &HE0A363
    APILine tmpw + 22, tmph, tmpw + 22, tmph1, &HE1A161
    APILine tmpw + 23, tmph, tmpw + 23, tmph1, &HE09F5B
    APILine tmpw + 24, tmph, tmpw + 24, tmph1, &HDE9855
    APILine tmpw + 25, tmph, tmpw + 25, tmph1, &HDC9752
    APILine tmpw + 26, tmph, tmpw + 26, tmph1, &HDB934B
    APILine tmpw + 27, tmph, tmpw + 27, tmph1, &HD68D39
    APILine tmpw + 28, tmph, tmpw + 28, tmph1, &HD17F2D
    APILine tmpw + 29, tmph, tmpw + 29, tmph1, &HD67426
    APILine tmpw + 30, tmph, tmpw + 30, tmph1, &HC05D13
    APILine tmpw + 31, tmph, tmpw + 31, tmph1, &H7C3514
    APILine tmpw + 32, tmph, tmpw + 32, tmph1, &HAB9B98
    APILine tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6

    'Lineas verticales
    APILine 17, 0, lw - 17, 0, &H450608
    APILine 17, 1, lw - 17, 1, &HF1D4C9
    APILine 17, 2, lw - 17, 2, &HE5C8BD
    APILine 17, 3, lw - 17, 3, &HE8C0A1
    APILine 17, 4, lw - 17, 4, &HE0B898
    APILine 17, 5, lw - 17, 5, &HE3B48E
    APILine 17, 6, lw - 17, 6, &HE0B18B
    APILine 17, 7, lw - 17, 7, &HE9B47F
    APILine 17, 8, lw - 17, 8, &HCE9963
    APILine 17, 9, lw - 17, 9, &HDDA064
    APILine 17, 10, lw - 17, 10, &HE2A66A
    APILine 17, 11, lw - 17, 11, &HE6AC76
    tmph = lh - 22
    APILine 17, tmph + 11, lw - 17, tmph + 11, &HE6AC76
    APILine 17, tmph + 12, lw - 17, tmph + 12, &HF1B681
    APILine 17, tmph + 13, lw - 17, tmph + 13, &HF3BD8A
    APILine 17, tmph + 14, lw - 17, tmph + 14, &HFCC592
    APILine 17, tmph + 15, lw - 17, tmph + 15, &HF8CE97
    APILine 17, tmph + 16, lw - 17, tmph + 16, &HFED59E
    APILine 17, tmph + 17, lw - 17, tmph + 17, &HF7DDA3
    APILine 17, tmph + 18, lw - 17, tmph + 18, &HFFE6AD
    APILine 17, tmph + 19, lw - 17, tmph + 19, &HE9E2C5
    APILine 17, tmph + 20, lw - 17, tmph + 20, &H635D40
    APILine 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    APILine 17, tmph + 22, lw - 17, tmph + 22, &HECECEC

End Sub

Private Sub DrawMacOSXButtonPressed()


Dim lHDC  As Long
Dim lh    As Long
Dim lw    As Long
Dim tmph  As Long
Dim tmpw  As Long
Dim tmph1 As Long


    With UserControl
        lHDC = .hdc
        lh = .ScaleHeight
        lw = .ScaleWidth
    End With

    APIFillRectByCoords hdc, 18, 11, lw - 34, lh - 19, &HCC9B6A
    SetPixel lHDC, 6, 0, &HFEFEFE
    SetPixel lHDC, 7, 0, &HE5E4E4
    SetPixel lHDC, 8, 0, &HA5A2A2
    SetPixel lHDC, 9, 0, &H675C5C
    SetPixel lHDC, 10, 0, &H422729
    SetPixel lHDC, 11, 0, &H300E0D
    SetPixel lHDC, 12, 0, &H300A09
    SetPixel lHDC, 13, 0, &H2F0908
    SetPixel lHDC, 14, 0, &H330909
    SetPixel lHDC, 15, 0, &H390A0A
    SetPixel lHDC, 16, 0, &H3C0A0A
    SetPixel lHDC, 17, 0, &H3C090A
    SetPixel lHDC, 5, 1, &HF0EEEE
    SetPixel lHDC, 6, 1, &H9D888A
    SetPixel lHDC, 7, 1, &H653531
    SetPixel lHDC, 8, 1, &H5A201D
    SetPixel lHDC, 9, 1, &H8D655F
    SetPixel lHDC, 10, 1, &HB99995
    SetPixel lHDC, 11, 1, &HD0B4B2
    SetPixel lHDC, 12, 1, &HD7BEBB
    SetPixel lHDC, 13, 1, &HDDC6C0
    SetPixel lHDC, 14, 1, &HDDC6C0
    SetPixel lHDC, 15, 1, &HDDC7BE
    SetPixel lHDC, 16, 1, &HDDC7BE
    SetPixel lHDC, 17, 1, &HDEC7BE
    SetPixel lHDC, 3, 2, &HFEFEFE
    SetPixel lHDC, 4, 2, &HE4E4E4
    SetPixel lHDC, 5, 2, &H6F5C5C
    SetPixel lHDC, 6, 2, &H390A0E
    SetPixel lHDC, 7, 2, &H712E2A
    SetPixel lHDC, 8, 2, &HD6928D
    SetPixel lHDC, 9, 2, &HD8ACA6
    SetPixel lHDC, 10, 2, &HD1B0AC
    SetPixel lHDC, 11, 2, &HD1B5B2
    SetPixel lHDC, 12, 2, &HD0B7B4
    SetPixel lHDC, 13, 2, &HCEB7B1
    SetPixel lHDC, 14, 2, &HCEB7B1
    SetPixel lHDC, 15, 2, &HD2BCB2
    SetPixel lHDC, 16, 2, &HD2BCB2
    SetPixel lHDC, 17, 2, &HD3BCB2
    SetPixel lHDC, 3, 3, &HEEEDED
    SetPixel lHDC, 4, 3, &H805858
    SetPixel lHDC, 5, 3, &H6A0D08
    SetPixel lHDC, 6, 3, &H7D1909
    SetPixel lHDC, 7, 3, &HB07B63
    SetPixel lHDC, 8, 3, &HCFA58A
    SetPixel lHDC, 9, 3, &HCDA78E
    SetPixel lHDC, 10, 3, &HD1AB92
    SetPixel lHDC, 11, 3, &HD2AF93
    SetPixel lHDC, 12, 3, &HD3B094
    SetPixel lHDC, 13, 3, &HD0AF93
    SetPixel lHDC, 14, 3, &HD3B296
    SetPixel lHDC, 15, 3, &HD4B49A
    SetPixel lHDC, 16, 3, &HD4B39A
    SetPixel lHDC, 17, 3, &HD4B39A
    SetPixel lHDC, 2, 4, &HFBFBFB
    SetPixel lHDC, 3, 4, &H837576
    SetPixel lHDC, 4, 4, &H440C0C
    SetPixel lHDC, 5, 4, &H821D0D
    SetPixel lHDC, 6, 4, &HA94433
    SetPixel lHDC, 7, 4, &HC08B72
    SetPixel lHDC, 8, 4, &HC49A7F
    SetPixel lHDC, 9, 4, &HC6A188
    SetPixel lHDC, 10, 4, &HC7A189
    SetPixel lHDC, 11, 4, &HC6A387
    SetPixel lHDC, 12, 4, &HC8A689
    SetPixel lHDC, 13, 4, &HC9A98C
    SetPixel lHDC, 14, 4, &HC8A88B
    SetPixel lHDC, 15, 4, &HCBAB91
    SetPixel lHDC, 16, 4, &HCCAC92
    SetPixel lHDC, 17, 4, &HCCAC92
    SetPixel lHDC, 1, 5, &HFEFEFE
    SetPixel lHDC, 2, 5, &HCAC8C7
    SetPixel lHDC, 3, 5, &H79281D
    SetPixel lHDC, 4, 5, &H7F2409
    SetPixel lHDC, 5, 5, &H8C3809
    SetPixel lHDC, 6, 5, &HBD6D39
    SetPixel lHDC, 7, 5, &HC9986E
    SetPixel lHDC, 8, 5, &HC89D74
    SetPixel lHDC, 9, 5, &HC49A71
    SetPixel lHDC, 10, 5, &HCAA17C
    SetPixel lHDC, 11, 5, &HC6A07A
    SetPixel lHDC, 12, 5, &HCAA480
    SetPixel lHDC, 13, 5, &HCAA582
    SetPixel lHDC, 14, 5, &HCBA584
    SetPixel lHDC, 15, 5, &HCDA989
    SetPixel lHDC, 16, 5, &HCFA98A
    SetPixel lHDC, 17, 5, &HCFA88A
    SetPixel lHDC, 1, 6, &HF9F9F9
    SetPixel lHDC, 2, 6, &H756C6B
    SetPixel lHDC, 3, 6, &H76190D
    SetPixel lHDC, 4, 6, &H913416
    SetPixel lHDC, 5, 6, &H9D4916
    SetPixel lHDC, 6, 6, &HBA6A36
    SetPixel lHDC, 7, 6, &HC39268
    SetPixel lHDC, 8, 6, &HC59A71
    SetPixel lHDC, 9, 6, &HC59B72
    SetPixel lHDC, 10, 6, &HC59C77
    SetPixel lHDC, 11, 6, &HC6A07A
    SetPixel lHDC, 12, 6, &HC6A07C
    SetPixel lHDC, 13, 6, &HC7A27F
    SetPixel lHDC, 14, 6, &HCBA584
    SetPixel lHDC, 15, 6, &HCAA686
    SetPixel lHDC, 16, 6, &HCBA586
    SetPixel lHDC, 17, 6, &HCCA587
    SetPixel lHDC, 1, 7, &HE8E7E7
    SetPixel lHDC, 2, 7, &H6C3E35
    SetPixel lHDC, 3, 7, &H8A2D09
    SetPixel lHDC, 4, 7, &HA34812
    SetPixel lHDC, 5, 7, &HAB591A
    SetPixel lHDC, 6, 7, &HB46B2B
    SetPixel lHDC, 7, 7, &HC3854A
    SetPixel lHDC, 8, 7, &HD19C64
    SetPixel lHDC, 9, 7, &HCD9C6C
    SetPixel lHDC, 10, 7, &HD1A070
    SetPixel lHDC, 11, 7, &HD2A272
    SetPixel lHDC, 12, 7, &HD2A272
    SetPixel lHDC, 13, 7, &HD6A57A
    SetPixel lHDC, 14, 7, &HD8A77C
    SetPixel lHDC, 15, 7, &HD2A87C
    SetPixel lHDC, 16, 7, &HD2A87C
    SetPixel lHDC, 17, 7, &HD2A77D
    SetPixel lHDC, 0, 8, &HFDFDFD
    SetPixel lHDC, 1, 8, &HC7C3C3
    SetPixel lHDC, 2, 8, &H5C2A21
    SetPixel lHDC, 3, 8, &H9C3E15
    SetPixel lHDC, 4, 8, &HB35A22
    SetPixel lHDC, 5, 8, &HB56324
    SetPixel lHDC, 6, 8, &HB66D2D
    SetPixel lHDC, 7, 8, &HB6783D
    SetPixel lHDC, 8, 8, &HB07B44
    SetPixel lHDC, 9, 8, &HB18050
    SetPixel lHDC, 10, 8, &HB58454
    SetPixel lHDC, 11, 8, &HB48554
    SetPixel lHDC, 12, 8, &HB78858
    SetPixel lHDC, 13, 8, &HBA895E
    SetPixel lHDC, 14, 8, &HBB8A5F
    SetPixel lHDC, 15, 8, &HB98E62
    SetPixel lHDC, 16, 8, &HB98E62
    SetPixel lHDC, 17, 8, &HB98E62
    SetPixel lHDC, 0, 9, &HFAFAFA
    SetPixel lHDC, 1, 9, &HB4ABA9
    SetPixel lHDC, 2, 9, &H612A14
    SetPixel lHDC, 3, 9, &HA05316
    SetPixel lHDC, 4, 9, &HB36628
    SetPixel lHDC, 5, 9, &HB67132
    SetPixel lHDC, 6, 9, &HB67738
    SetPixel lHDC, 7, 9, &HB98146
    SetPixel lHDC, 8, 9, &HBD864E
    SetPixel lHDC, 9, 9, &HBD894F
    SetPixel lHDC, 10, 9, &HC28D55
    SetPixel lHDC, 11, 9, &HC4905B
    SetPixel lHDC, 12, 9, &HC5905F
    SetPixel lHDC, 13, 9, &HC49161
    SetPixel lHDC, 14, 9, &HC49161
    SetPixel lHDC, 15, 9, &HC69564
    SetPixel lHDC, 16, 9, &HC69564
    SetPixel lHDC, 17, 9, &HC69464
    SetPixel lHDC, 0, 10, &HF7F7F7
    SetPixel lHDC, 1, 10, &HA99D9B
    SetPixel lHDC, 2, 10, &H632D17
    SetPixel lHDC, 3, 10, &HA65A1D
    SetPixel lHDC, 4, 10, &HB96C2E
    SetPixel lHDC, 5, 10, &HBC7738
    SetPixel lHDC, 6, 10, &HC18242
    SetPixel lHDC, 7, 10, &HC2894E
    SetPixel lHDC, 8, 10, &HC18A52
    SetPixel lHDC, 9, 10, &HC59157
    SetPixel lHDC, 10, 10, &HC59159
    SetPixel lHDC, 11, 10, &HCC9863
    SetPixel lHDC, 12, 10, &HCC9665
    SetPixel lHDC, 13, 10, &HCB9767
    SetPixel lHDC, 14, 10, &HC99565
    SetPixel lHDC, 15, 10, &HCC9A6A
    SetPixel lHDC, 16, 10, &HCC9A6A
    SetPixel lHDC, 17, 10, &HCC9B6A

    tmph = lh - 22
    SetPixel lHDC, 0, tmph + 10, &HF7F7F7
    SetPixel lHDC, 1, tmph + 10, &HA99D9B
    SetPixel lHDC, 2, tmph + 10, &H632D17
    SetPixel lHDC, 3, tmph + 10, &HA65A1D
    SetPixel lHDC, 4, tmph + 10, &HB96C2E
    SetPixel lHDC, 5, tmph + 10, &HBC7738
    SetPixel lHDC, 6, tmph + 10, &HC18242
    SetPixel lHDC, 7, tmph + 10, &HC2894E
    SetPixel lHDC, 8, tmph + 10, &HC18A52
    SetPixel lHDC, 9, tmph + 10, &HC59157
    SetPixel lHDC, 10, tmph + 10, &HC59159
    SetPixel lHDC, 11, tmph + 10, &HCC9863
    SetPixel lHDC, 12, tmph + 10, &HCC9665
    SetPixel lHDC, 13, tmph + 10, &HCB9767
    SetPixel lHDC, 14, tmph + 10, &HC99565
    SetPixel lHDC, 15, tmph + 10, &HCC9A6A
    SetPixel lHDC, 16, tmph + 10, &HCC9A6A
    SetPixel lHDC, 17, tmph + 10, &HCC9B6A
    SetPixel lHDC, 0, tmph + 11, &HF5F5F5
    SetPixel lHDC, 1, tmph + 11, &HA59F9A
    SetPixel lHDC, 2, tmph + 11, &H674024
    SetPixel lHDC, 3, tmph + 11, &HAE6827
    SetPixel lHDC, 4, tmph + 11, &HB97231
    SetPixel lHDC, 5, tmph + 11, &HBE8247
    SetPixel lHDC, 6, tmph + 11, &HC0874E
    SetPixel lHDC, 7, tmph + 11, &HC78E56
    SetPixel lHDC, 8, tmph + 11, &HCD9561
    SetPixel lHDC, 9, tmph + 11, &HCB9466
    SetPixel lHDC, 10, tmph + 11, &HCD9A6B
    SetPixel lHDC, 11, tmph + 11, &HC79867
    SetPixel lHDC, 12, tmph + 11, &HCA9B6A
    SetPixel lHDC, 13, tmph + 11, &HCC9D6C
    SetPixel lHDC, 14, tmph + 11, &HCD9D70
    SetPixel lHDC, 15, tmph + 11, &HD0A175
    SetPixel lHDC, 16, tmph + 11, &HD0A175
    SetPixel lHDC, 17, tmph + 11, &HD0A175
    SetPixel lHDC, 0, tmph + 12, &HF5F5F5
    SetPixel lHDC, 1, tmph + 12, &HACA7A4
    SetPixel lHDC, 2, tmph + 12, &H755035
    SetPixel lHDC, 3, tmph + 12, &HB77131
    SetPixel lHDC, 4, tmph + 12, &HCB8443
    SetPixel lHDC, 5, tmph + 12, &HC5894E
    SetPixel lHDC, 6, tmph + 12, &HCC935A
    SetPixel lHDC, 7, tmph + 12, &HD29962
    SetPixel lHDC, 8, tmph + 12, &HD69F6A
    SetPixel lHDC, 9, tmph + 12, &HDBA476
    SetPixel lHDC, 10, tmph + 12, &HD6A374
    SetPixel lHDC, 11, tmph + 12, &HD4A574
    SetPixel lHDC, 12, tmph + 12, &HD8A978
    SetPixel lHDC, 13, tmph + 12, &HDAAB7A
    SetPixel lHDC, 14, tmph + 12, &HDAAA7D
    SetPixel lHDC, 15, tmph + 12, &HDBAB7F
    SetPixel lHDC, 16, tmph + 12, &HDAAA7F
    SetPixel lHDC, 17, tmph + 12, &HDAAA7F
    SetPixel lHDC, 0, tmph + 13, &HF7F7F7
    SetPixel lHDC, 1, tmph + 13, &HC0C0BF
    SetPixel lHDC, 2, tmph + 13, &H63574B
    SetPixel lHDC, 3, tmph + 13, &HAC7036
    SetPixel lHDC, 4, tmph + 13, &HC2854A
    SetPixel lHDC, 5, tmph + 13, &HCF955E
    SetPixel lHDC, 6, tmph + 13, &HD29B66
    SetPixel lHDC, 7, tmph + 13, &HD1A26E
    SetPixel lHDC, 8, tmph + 13, &HD8A776
    SetPixel lHDC, 9, tmph + 13, &HDBA878
    SetPixel lHDC, 10, tmph + 13, &HDFAC7C
    SetPixel lHDC, 11, tmph + 13, &HDBAF7D
    SetPixel lHDC, 12, tmph + 13, &HDDAF81
    SetPixel lHDC, 13, tmph + 13, &HDEB183
    SetPixel lHDC, 14, tmph + 13, &HDDAF84
    SetPixel lHDC, 15, tmph + 13, &HDEB087
    SetPixel lHDC, 16, tmph + 13, &HDEB087
    SetPixel lHDC, 17, tmph + 13, &HDCB087
    SetPixel lHDC, 0, tmph + 14, &HFBFBFB
    SetPixel lHDC, 1, tmph + 14, &HE1E1E1
    SetPixel lHDC, 2, tmph + 14, &H7C7269
    SetPixel lHDC, 3, tmph + 14, &HA26830
    SetPixel lHDC, 4, tmph + 14, &HC6884E
    SetPixel lHDC, 5, tmph + 14, &HD0965F
    SetPixel lHDC, 6, tmph + 14, &HDAA26E
    SetPixel lHDC, 7, tmph + 14, &HD9AA75
    SetPixel lHDC, 8, tmph + 14, &HDBAA79
    SetPixel lHDC, 9, tmph + 14, &HE2AF7F
    SetPixel lHDC, 10, tmph + 14, &HE6B484
    SetPixel lHDC, 11, tmph + 14, &HE2B684
    SetPixel lHDC, 12, tmph + 14, &HE3B588
    SetPixel lHDC, 13, tmph + 14, &HE2B587
    SetPixel lHDC, 14, tmph + 14, &HE2B48A
    SetPixel lHDC, 15, tmph + 14, &HE5B78E
    SetPixel lHDC, 16, tmph + 14, &HE5B78E
    SetPixel lHDC, 17, tmph + 14, &HE4B88E
    SetPixel lHDC, 0, tmph + 15, &HFEFEFE
    SetPixel lHDC, 1, tmph + 15, &HEDEDED
    SetPixel lHDC, 2, tmph + 15, &H9E9C9C
    SetPixel lHDC, 3, tmph + 15, &H766051
    SetPixel lHDC, 4, tmph + 15, &HAD8666
    SetPixel lHDC, 5, tmph + 15, &HD49A61
    SetPixel lHDC, 6, tmph + 15, &HE0A66D
    SetPixel lHDC, 7, tmph + 15, &HE3B17C
    SetPixel lHDC, 8, tmph + 15, &HE0B380
    SetPixel lHDC, 9, tmph + 15, &HE0B587
    SetPixel lHDC, 10, tmph + 15, &HE2BC8C
    SetPixel lHDC, 11, tmph + 15, &HE0BB8B
    SetPixel lHDC, 12, tmph + 15, &HE0BC8B
    SetPixel lHDC, 13, tmph + 15, &HE3BD92
    SetPixel lHDC, 14, tmph + 15, &HE2BC91
    SetPixel lHDC, 15, tmph + 15, &HE2BF93
    SetPixel lHDC, 16, tmph + 15, &HE1BE93
    SetPixel lHDC, 17, tmph + 15, &HE1BF93
    SetPixel lHDC, 1, tmph + 16, &HF6F6F6
    SetPixel lHDC, 2, tmph + 16, &HD5D5D5
    SetPixel lHDC, 3, tmph + 16, &H86766C
    SetPixel lHDC, 4, tmph + 16, &H856144
    SetPixel lHDC, 5, tmph + 16, &HD59C63
    SetPixel lHDC, 6, tmph + 16, &HE5AB71
    SetPixel lHDC, 7, tmph + 16, &HE5B37E
    SetPixel lHDC, 8, tmph + 16, &HE7BB88
    SetPixel lHDC, 9, tmph + 16, &HE7BF91
    SetPixel lHDC, 10, tmph + 16, &HE3BC8D
    SetPixel lHDC, 11, tmph + 16, &HE7C392
    SetPixel lHDC, 12, tmph + 16, &HE7C392
    SetPixel lHDC, 13, tmph + 16, &HE8C398
    SetPixel lHDC, 14, tmph + 16, &HE8C499
    SetPixel lHDC, 15, tmph + 16, &HE8C599
    SetPixel lHDC, 16, tmph + 16, &HE8C599
    SetPixel lHDC, 17, tmph + 16, &HE7C699
    SetPixel lHDC, 1, tmph + 17, &HFDFDFD
    SetPixel lHDC, 2, tmph + 17, &HEDEDED
    SetPixel lHDC, 3, tmph + 17, &HBDBDBD
    SetPixel lHDC, 4, tmph + 17, &H676767
    SetPixel lHDC, 5, tmph + 17, &H71604C
    SetPixel lHDC, 6, tmph + 17, &HBEA17D
    SetPixel lHDC, 7, tmph + 17, &HDAB381
    SetPixel lHDC, 8, tmph + 17, &HE5BE8C
    SetPixel lHDC, 9, tmph + 17, &HE1C18F
    SetPixel lHDC, 10, tmph + 17, &HE4C895
    SetPixel lHDC, 11, tmph + 17, &HDFCA98
    SetPixel lHDC, 12, tmph + 17, &HE2CE9B
    SetPixel lHDC, 13, tmph + 17, &HE2CE9B
    SetPixel lHDC, 14, tmph + 17, &HE2CE9B
    SetPixel lHDC, 15, tmph + 17, &HE2CD9D
    SetPixel lHDC, 16, tmph + 17, &HE2CC9D
    SetPixel lHDC, 17, tmph + 17, &HE2CC9D
    SetPixel lHDC, 2, tmph + 18, &HF9F9F9
    SetPixel lHDC, 3, tmph + 18, &HE6E6E6
    SetPixel lHDC, 4, tmph + 18, &HB9B9B9
    SetPixel lHDC, 5, tmph + 18, &H7A7163
    SetPixel lHDC, 6, tmph + 18, &H776043
    SetPixel lHDC, 7, tmph + 18, &HAB885B
    SetPixel lHDC, 8, tmph + 18, &HDDB888
    SetPixel lHDC, 9, tmph + 18, &HE6C796
    SetPixel lHDC, 10, tmph + 18, &HE8CD9A
    SetPixel lHDC, 11, tmph + 18, &HE5D19E
    SetPixel lHDC, 12, tmph + 18, &HE9D6A3
    SetPixel lHDC, 13, tmph + 18, &HE9D6A5
    SetPixel lHDC, 14, tmph + 18, &HE9D6A3
    SetPixel lHDC, 15, tmph + 18, &HE9D5A6
    SetPixel lHDC, 16, tmph + 18, &HE9D5A6
    SetPixel lHDC, 17, tmph + 18, &HE9D5A6
    SetPixel lHDC, 2, tmph + 19, &HFEFEFE
    SetPixel lHDC, 3, tmph + 19, &HF8F8F8
    SetPixel lHDC, 4, tmph + 19, &HE6E6E6
    SetPixel lHDC, 5, tmph + 19, &HC8C8C8
    SetPixel lHDC, 6, tmph + 19, &H8C8C8C
    SetPixel lHDC, 7, tmph + 19, &H61605E
    SetPixel lHDC, 8, tmph + 19, &H656059
    SetPixel lHDC, 9, tmph + 19, &H857C6D
    SetPixel lHDC, 10, tmph + 19, &HA59C87
    SetPixel lHDC, 11, tmph + 19, &HC8C1A8
    SetPixel lHDC, 12, tmph + 19, &HD1CAB0
    SetPixel lHDC, 13, tmph + 19, &HD5CFB5
    SetPixel lHDC, 14, tmph + 19, &HD6D1B6
    SetPixel lHDC, 15, tmph + 19, &HD7D2BA
    SetPixel lHDC, 16, tmph + 19, &HD7D1BA
    SetPixel lHDC, 17, tmph + 19, &HD7D2BA
    SetPixel lHDC, 3, tmph + 20, &HFEFEFE
    SetPixel lHDC, 4, tmph + 20, &HF9F9F9
    SetPixel lHDC, 5, tmph + 20, &HECECEC
    SetPixel lHDC, 6, tmph + 20, &HDADADA
    SetPixel lHDC, 7, tmph + 20, &HC1C1C1
    SetPixel lHDC, 8, tmph + 20, &H9C9B99
    SetPixel lHDC, 9, tmph + 20, &H7D7A73
    SetPixel lHDC, 10, tmph + 20, &H635E50
    SetPixel lHDC, 11, tmph + 20, &H58533F
    SetPixel lHDC, 12, tmph + 20, &H554F39
    SetPixel lHDC, 13, tmph + 20, &H514D36
    SetPixel lHDC, 14, tmph + 20, &H554F37
    SetPixel lHDC, 15, tmph + 20, &H57523A
    SetPixel lHDC, 16, tmph + 20, &H5A563D
    SetPixel lHDC, 17, tmph + 20, &H5A563E
    SetPixel lHDC, 5, tmph + 21, &HFCFCFC
    SetPixel lHDC, 6, tmph + 21, &HF5F5F5
    SetPixel lHDC, 7, tmph + 21, &HEBEBEB
    SetPixel lHDC, 8, tmph + 21, &HE1E1E1
    SetPixel lHDC, 9, tmph + 21, &HD6D6D6
    SetPixel lHDC, 10, tmph + 21, &HCECECE
    SetPixel lHDC, 11, tmph + 21, &HC9C9C9
    SetPixel lHDC, 12, tmph + 21, &HC7C7C7
    SetPixel lHDC, 13, tmph + 21, &HC7C7C7
    SetPixel lHDC, 14, tmph + 21, &HC6C6C6
    SetPixel lHDC, 15, tmph + 21, &HC6C6C6
    SetPixel lHDC, 16, tmph + 21, &HC5C5C5
    SetPixel lHDC, 17, tmph + 21, &HC5C5C5
    SetPixel lHDC, 7, tmph + 22, &HFDFDFD
    SetPixel lHDC, 8, tmph + 22, &HF9F9F9
    SetPixel lHDC, 9, tmph + 22, &HF4F4F4
    SetPixel lHDC, 10, tmph + 22, &HF0F0F0
    SetPixel lHDC, 11, tmph + 22, &HEEEEEE
    SetPixel lHDC, 12, tmph + 22, &HEDEDED
    SetPixel lHDC, 13, tmph + 22, &HECECEC
    SetPixel lHDC, 14, tmph + 22, &HECECEC
    SetPixel lHDC, 15, tmph + 22, &HECECEC
    SetPixel lHDC, 16, tmph + 22, &HECECEC
    SetPixel lHDC, 17, tmph + 22, &HECECEC

    tmpw = lw - 34
    SetPixel lHDC, tmpw + 17, 0, &H3C090A
    SetPixel lHDC, tmpw + 18, 0, &H3C090A
    SetPixel lHDC, tmpw + 19, 0, &H340A0A
    SetPixel lHDC, tmpw + 20, 0, &H300A09
    SetPixel lHDC, tmpw + 21, 0, &H2F080A
    SetPixel lHDC, tmpw + 22, 0, &H341011
    SetPixel lHDC, tmpw + 23, 0, &H3E2526
    SetPixel lHDC, tmpw + 24, 0, &H5A4C4C
    SetPixel lHDC, tmpw + 25, 0, &H9E9B9B
    SetPixel lHDC, tmpw + 26, 0, &HEEEEEE
    SetPixel lHDC, tmpw + 34, 0, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 1, &HDEC7BE
    SetPixel lHDC, tmpw + 18, 1, &HDEC7BE
    SetPixel lHDC, tmpw + 19, 1, &HDBC6C1
    SetPixel lHDC, tmpw + 20, 1, &HD9C4BF
    SetPixel lHDC, tmpw + 21, 1, &HD7C1B9
    SetPixel lHDC, tmpw + 22, 1, &HD3B5AF
    SetPixel lHDC, tmpw + 23, 1, &HBE9F97
    SetPixel lHDC, tmpw + 24, 1, &H9B6A65
    SetPixel lHDC, tmpw + 25, 1, &H65231E
    SetPixel lHDC, tmpw + 26, 1, &H642A26
    SetPixel lHDC, tmpw + 27, 1, &HA59696
    SetPixel lHDC, tmpw + 28, 1, &HF7F7F7
    SetPixel lHDC, tmpw + 34, 1, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 2, &HD3BCB2
    SetPixel lHDC, tmpw + 18, 2, &HD3BCB2
    SetPixel lHDC, tmpw + 19, 2, &HCDB8B3
    SetPixel lHDC, tmpw + 20, 2, &HCBB6B1
    SetPixel lHDC, tmpw + 21, 2, &HD0BBB2
    SetPixel lHDC, tmpw + 22, 2, &HD0B2AC
    SetPixel lHDC, tmpw + 23, 2, &HD6B6AF
    SetPixel lHDC, tmpw + 24, 2, &HDCABA6
    SetPixel lHDC, tmpw + 25, 2, &HDC9691
    SetPixel lHDC, tmpw + 26, 2, &H732E29
    SetPixel lHDC, tmpw + 27, 2, &H380A0A
    SetPixel lHDC, tmpw + 28, 2, &H6A5556
    SetPixel lHDC, tmpw + 29, 2, &HEAEBEA
    SetPixel lHDC, tmpw + 34, 2, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 3, &HD4B39A
    SetPixel lHDC, tmpw + 18, 3, &HD4B39A
    SetPixel lHDC, tmpw + 19, 3, &HD1B294
    SetPixel lHDC, tmpw + 20, 3, &HD0B193
    SetPixel lHDC, tmpw + 21, 3, &HD0AE91
    SetPixel lHDC, tmpw + 22, 3, &HD4B296
    SetPixel lHDC, tmpw + 23, 3, &HCBAA8F
    SetPixel lHDC, tmpw + 24, 3, &HCBAA8F
    SetPixel lHDC, tmpw + 25, 3, &HCCA38B
    SetPixel lHDC, tmpw + 26, 3, &HB77E68
    SetPixel lHDC, tmpw + 27, 3, &H811B09
    SetPixel lHDC, tmpw + 28, 3, &H720E08
    SetPixel lHDC, tmpw + 29, 3, &H7D5051
    SetPixel lHDC, tmpw + 30, 3, &HEFEEEE
    SetPixel lHDC, tmpw + 34, 3, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 4, &HCCAC92
    SetPixel lHDC, tmpw + 18, 4, &HCCAC91
    SetPixel lHDC, tmpw + 19, 4, &HC6A889
    SetPixel lHDC, tmpw + 20, 4, &HC7A98A
    SetPixel lHDC, tmpw + 21, 4, &HC7A589
    SetPixel lHDC, tmpw + 22, 4, &HC4A185
    SetPixel lHDC, tmpw + 23, 4, &HC6A58A
    SetPixel lHDC, tmpw + 24, 4, &HBF9E83
    SetPixel lHDC, tmpw + 25, 4, &HC39A82
    SetPixel lHDC, tmpw + 26, 4, &HC58C76
    SetPixel lHDC, tmpw + 27, 4, &HA9432F
    SetPixel lHDC, tmpw + 28, 4, &H861F0C
    SetPixel lHDC, tmpw + 29, 4, &H460B0C
    SetPixel lHDC, tmpw + 30, 4, &H7B6B6C
    SetPixel lHDC, tmpw + 31, 4, &HFAFAFA
    SetPixel lHDC, tmpw + 34, 4, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 5, &HCFA88A
    SetPixel lHDC, tmpw + 18, 5, &HCFA889
    SetPixel lHDC, tmpw + 19, 5, &HCBA683
    SetPixel lHDC, tmpw + 20, 5, &HC9A481
    SetPixel lHDC, tmpw + 21, 5, &HCCA480
    SetPixel lHDC, tmpw + 22, 5, &HCEA280
    SetPixel lHDC, tmpw + 23, 5, &HCCA379
    SetPixel lHDC, tmpw + 24, 5, &HCA9E74
    SetPixel lHDC, tmpw + 25, 5, &HC69971
    SetPixel lHDC, tmpw + 26, 5, &HC89870
    SetPixel lHDC, tmpw + 27, 5, &HB46A34
    SetPixel lHDC, tmpw + 28, 5, &H90380A
    SetPixel lHDC, tmpw + 29, 5, &H892509
    SetPixel lHDC, tmpw + 30, 5, &H8A251B
    SetPixel lHDC, tmpw + 31, 5, &HC4C2C2
    SetPixel lHDC, tmpw + 32, 5, &HFEFEFE
    SetPixel lHDC, tmpw + 34, 5, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 6, &HCCA587
    SetPixel lHDC, tmpw + 18, 6, &HCCA586
    SetPixel lHDC, tmpw + 19, 6, &HC9A481
    SetPixel lHDC, tmpw + 20, 6, &HC9A481
    SetPixel lHDC, tmpw + 21, 6, &HC7A07C
    SetPixel lHDC, tmpw + 22, 6, &HCCA17E
    SetPixel lHDC, tmpw + 23, 6, &HC79F74
    SetPixel lHDC, tmpw + 24, 6, &HC69A70
    SetPixel lHDC, tmpw + 25, 6, &HC59870
    SetPixel lHDC, tmpw + 26, 6, &HC2926A
    SetPixel lHDC, tmpw + 27, 6, &HB96F39
    SetPixel lHDC, tmpw + 28, 6, &HA04814
    SetPixel lHDC, tmpw + 29, 6, &H973215
    SetPixel lHDC, tmpw + 30, 6, &H831A0F
    SetPixel lHDC, tmpw + 31, 6, &H6E6966
    SetPixel lHDC, tmpw + 32, 6, &HF8F8F8
    SetPixel lHDC, tmpw + 34, 6, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 7, &HD2A77D
    SetPixel lHDC, tmpw + 18, 7, &HD3A77C
    SetPixel lHDC, tmpw + 19, 7, &HD8AA7D
    SetPixel lHDC, tmpw + 20, 7, &HD2A376
    SetPixel lHDC, tmpw + 21, 7, &HD1A373
    SetPixel lHDC, tmpw + 22, 7, &HCEA070
    SetPixel lHDC, tmpw + 23, 7, &HD2A06F
    SetPixel lHDC, tmpw + 24, 7, &HD19D68
    SetPixel lHDC, tmpw + 25, 7, &HD09A65
    SetPixel lHDC, tmpw + 26, 7, &HC2864F
    SetPixel lHDC, tmpw + 27, 7, &HAE6927
    SetPixel lHDC, tmpw + 28, 7, &HA95A19
    SetPixel lHDC, tmpw + 29, 7, &HA44A10
    SetPixel lHDC, tmpw + 30, 7, &H8B2E09
    SetPixel lHDC, tmpw + 31, 7, &H6B3E34
    SetPixel lHDC, tmpw + 32, 7, &HE7E6E6
    SetPixel lHDC, tmpw + 34, 7, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 8, &HB98E62
    SetPixel lHDC, tmpw + 18, 8, &HBA8E62
    SetPixel lHDC, tmpw + 19, 8, &HB98B5E
    SetPixel lHDC, tmpw + 20, 8, &HB98B5E
    SetPixel lHDC, tmpw + 21, 8, &HB68858
    SetPixel lHDC, tmpw + 22, 8, &HB48656
    SetPixel lHDC, tmpw + 23, 8, &HB58452
    SetPixel lHDC, tmpw + 24, 8, &HB5814C
    SetPixel lHDC, tmpw + 25, 8, &HB07A46
    SetPixel lHDC, tmpw + 26, 8, &HB2773F
    SetPixel lHDC, tmpw + 27, 8, &HB36E2C
    SetPixel lHDC, tmpw + 28, 8, &HB26221
    SetPixel lHDC, tmpw + 29, 8, &HB35A20
    SetPixel lHDC, tmpw + 30, 8, &H9C3E11
    SetPixel lHDC, tmpw + 31, 8, &H5C2A1F
    SetPixel lHDC, tmpw + 32, 8, &HC4C1C0
    SetPixel lHDC, tmpw + 33, 8, &HFDFDFD
    SetPixel lHDC, tmpw + 34, 8, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 9, &HC69464
    SetPixel lHDC, tmpw + 18, 9, &HC69564
    SetPixel lHDC, tmpw + 19, 9, &HC3925E
    SetPixel lHDC, tmpw + 20, 9, &HC3915D
    SetPixel lHDC, tmpw + 21, 9, &HC3925E
    SetPixel lHDC, tmpw + 22, 9, &HC38F5B
    SetPixel lHDC, tmpw + 23, 9, &HC28D55
    SetPixel lHDC, tmpw + 24, 9, &HC08751
    SetPixel lHDC, tmpw + 25, 9, &HBC844C
    SetPixel lHDC, tmpw + 26, 9, &HBC8147
    SetPixel lHDC, tmpw + 27, 9, &HB57936
    SetPixel lHDC, tmpw + 28, 9, &HB3702D
    SetPixel lHDC, tmpw + 29, 9, &HB56626
    SetPixel lHDC, tmpw + 30, 9, &HA25115
    SetPixel lHDC, tmpw + 31, 9, &H662D12
    SetPixel lHDC, tmpw + 32, 9, &HAEA3A1
    SetPixel lHDC, tmpw + 33, 9, &HF9F9F9
    SetPixel lHDC, tmpw + 34, 9, &HFFFFFFFF
    SetPixel lHDC, tmpw + 17, 10, &HCC9B6A
    SetPixel lHDC, tmpw + 18, 10, &HCC9B6A
    SetPixel lHDC, tmpw + 19, 10, &HCA9864
    SetPixel lHDC, tmpw + 20, 10, &HC99763
    SetPixel lHDC, tmpw + 21, 10, &HC99763
    SetPixel lHDC, tmpw + 22, 10, &HCA9562
    SetPixel lHDC, tmpw + 23, 10, &HC9945D
    SetPixel lHDC, tmpw + 24, 10, &HC68E57
    SetPixel lHDC, tmpw + 25, 10, &HC48C55
    SetPixel lHDC, tmpw + 26, 10, &HC3884E
    SetPixel lHDC, tmpw + 27, 10, &HBE823E
    SetPixel lHDC, tmpw + 28, 10, &HB97634
    SetPixel lHDC, tmpw + 29, 10, &HBD6D2D
    SetPixel lHDC, tmpw + 30, 10, &HA8581C
    SetPixel lHDC, tmpw + 31, 10, &H6D3319
    SetPixel lHDC, tmpw + 32, 10, &HA49794
    SetPixel lHDC, tmpw + 33, 10, &HF6F6F6
    SetPixel lHDC, tmpw + 34, 10, &HFFFFFFFF

    tmph = lh - 22
    tmpw = lw - 34
    SetPixel lHDC, tmpw + 17, tmph + 10, &HCC9B6A
    SetPixel lHDC, tmpw + 18, tmph + 10, &HCC9B6A
    SetPixel lHDC, tmpw + 19, tmph + 10, &HCA9864
    SetPixel lHDC, tmpw + 20, tmph + 10, &HC99763
    SetPixel lHDC, tmpw + 21, tmph + 10, &HC99763
    SetPixel lHDC, tmpw + 22, tmph + 10, &HCA9562
    SetPixel lHDC, tmpw + 23, tmph + 10, &HC9945D
    SetPixel lHDC, tmpw + 24, tmph + 10, &HC68E57
    SetPixel lHDC, tmpw + 25, tmph + 10, &HC48C55
    SetPixel lHDC, tmpw + 26, tmph + 10, &HC3884E
    SetPixel lHDC, tmpw + 27, tmph + 10, &HBE823E
    SetPixel lHDC, tmpw + 28, tmph + 10, &HB97634
    SetPixel lHDC, tmpw + 29, tmph + 10, &HBD6D2D
    SetPixel lHDC, tmpw + 30, tmph + 10, &HA8581C
    SetPixel lHDC, tmpw + 31, tmph + 10, &H6D3319
    SetPixel lHDC, tmpw + 32, tmph + 10, &HA49794
    SetPixel lHDC, tmpw + 33, tmph + 10, &HF6F6F6
    SetPixel lHDC, tmpw + 17, tmph + 11, &HD0A175
    SetPixel lHDC, tmpw + 18, tmph + 11, &HD0A175
    SetPixel lHDC, tmpw + 19, tmph + 11, &HCC9D6D
    SetPixel lHDC, tmpw + 20, tmph + 11, &HCE9B6C
    SetPixel lHDC, tmpw + 21, tmph + 11, &HCB9A6A
    SetPixel lHDC, tmpw + 22, tmph + 11, &HCD996A
    SetPixel lHDC, tmpw + 23, tmph + 11, &HCA9666
    SetPixel lHDC, tmpw + 24, tmph + 11, &HCF9865
    SetPixel lHDC, tmpw + 25, tmph + 11, &HCA9460
    SetPixel lHDC, tmpw + 26, tmph + 11, &HC78F57
    SetPixel lHDC, tmpw + 27, tmph + 11, &HC1864B
    SetPixel lHDC, tmpw + 28, tmph + 11, &HC08143
    SetPixel lHDC, tmpw + 29, tmph + 11, &HB7712E
    SetPixel lHDC, tmpw + 30, tmph + 11, &HB16A28
    SetPixel lHDC, tmpw + 31, tmph + 11, &H694321
    SetPixel lHDC, tmpw + 32, tmph + 11, &HA59F9B
    SetPixel lHDC, tmpw + 33, tmph + 11, &HF4F4F4
    SetPixel lHDC, tmpw + 17, tmph + 12, &HDAAA7F
    SetPixel lHDC, tmpw + 18, tmph + 12, &HD9AB7E
    SetPixel lHDC, tmpw + 19, tmph + 12, &HDBAC7C
    SetPixel lHDC, tmpw + 20, tmph + 12, &HDDAA7B
    SetPixel lHDC, tmpw + 21, tmph + 12, &HDAA979
    SetPixel lHDC, tmpw + 22, tmph + 12, &HDAA677
    SetPixel lHDC, tmpw + 23, tmph + 12, &HD8A474
    SetPixel lHDC, tmpw + 24, tmph + 12, &HDBA471
    SetPixel lHDC, tmpw + 25, tmph + 12, &HD49F6A
    SetPixel lHDC, tmpw + 26, tmph + 12, &HD09861
    SetPixel lHDC, tmpw + 27, tmph + 12, &HD0955A
    SetPixel lHDC, tmpw + 28, tmph + 12, &HCC8D4F
    SetPixel lHDC, tmpw + 29, tmph + 12, &HCA8441
    SetPixel lHDC, tmpw + 30, tmph + 12, &HBB7532
    SetPixel lHDC, tmpw + 31, tmph + 12, &H7B5434
    SetPixel lHDC, tmpw + 32, tmph + 12, &HB1ACAA
    SetPixel lHDC, tmpw + 33, tmph + 12, &HF5F5F5
    SetPixel lHDC, tmpw + 17, tmph + 13, &HDCB087
    SetPixel lHDC, tmpw + 18, tmph + 13, &HDCB087
    SetPixel lHDC, tmpw + 19, tmph + 13, &HDBAF81
    SetPixel lHDC, tmpw + 20, tmph + 13, &HDEAF82
    SetPixel lHDC, tmpw + 21, tmph + 13, &HDCAF81
    SetPixel lHDC, tmpw + 22, tmph + 13, &HDDAD7F
    SetPixel lHDC, tmpw + 23, tmph + 13, &HDBAC7B
    SetPixel lHDC, tmpw + 24, tmph + 13, &HDDAA7A
    SetPixel lHDC, tmpw + 25, tmph + 13, &HD9A775
    SetPixel lHDC, tmpw + 26, tmph + 13, &HD7A26E
    SetPixel lHDC, tmpw + 27, tmph + 13, &HCE9961
    SetPixel lHDC, tmpw + 28, tmph + 13, &HCC945C
    SetPixel lHDC, tmpw + 29, tmph + 13, &HC2854D
    SetPixel lHDC, tmpw + 30, tmph + 13, &HAF7239
    SetPixel lHDC, tmpw + 31, tmph + 13, &H695C4F
    SetPixel lHDC, tmpw + 32, tmph + 13, &HD5D5D5
    SetPixel lHDC, tmpw + 33, tmph + 13, &HF8F8F8
    SetPixel lHDC, tmpw + 17, tmph + 14, &HE4B88E
    SetPixel lHDC, tmpw + 18, tmph + 14, &HE3B88E
    SetPixel lHDC, tmpw + 19, tmph + 14, &HE0B486
    SetPixel lHDC, tmpw + 20, tmph + 14, &HE4B689
    SetPixel lHDC, tmpw + 21, tmph + 14, &HE2B587
    SetPixel lHDC, tmpw + 22, tmph + 14, &HE5B588
    SetPixel lHDC, tmpw + 23, tmph + 14, &HE3B483
    SetPixel lHDC, tmpw + 24, tmph + 14, &HE2AF7F
    SetPixel lHDC, tmpw + 25, tmph + 14, &HDEAC7A
    SetPixel lHDC, tmpw + 26, tmph + 14, &HDEAA75
    SetPixel lHDC, tmpw + 27, tmph + 14, &HD8A36B
    SetPixel lHDC, tmpw + 28, tmph + 14, &HD09860
    SetPixel lHDC, tmpw + 29, tmph + 14, &HC58850
    SetPixel lHDC, tmpw + 30, tmph + 14, &HA56930
    SetPixel lHDC, tmpw + 31, tmph + 14, &H7B746C
    SetPixel lHDC, tmpw + 32, tmph + 14, &HE8E8E8
    SetPixel lHDC, tmpw + 33, tmph + 14, &HFDFDFD
    SetPixel lHDC, tmpw + 17, tmph + 15, &HE1BF93
    SetPixel lHDC, tmpw + 18, tmph + 15, &HE2BE93
    SetPixel lHDC, tmpw + 19, tmph + 15, &HE2BF91
    SetPixel lHDC, tmpw + 20, tmph + 15, &HE1BD8F
    SetPixel lHDC, tmpw + 21, tmph + 15, &HE0BC8E
    SetPixel lHDC, tmpw + 22, tmph + 15, &HE2BC8E
    SetPixel lHDC, tmpw + 23, tmph + 15, &HE4BD8C
    SetPixel lHDC, tmpw + 24, tmph + 15, &HE0B685
    SetPixel lHDC, tmpw + 25, tmph + 15, &HDCB07E
    SetPixel lHDC, tmpw + 26, tmph + 15, &HE1AF7C
    SetPixel lHDC, tmpw + 27, tmph + 15, &HDEA66E
    SetPixel lHDC, tmpw + 28, tmph + 15, &HD19962
    SetPixel lHDC, tmpw + 29, tmph + 15, &HAD875D
    SetPixel lHDC, tmpw + 30, tmph + 15, &H7D6851
    SetPixel lHDC, tmpw + 31, tmph + 15, &HB9B9B9
    SetPixel lHDC, tmpw + 32, tmph + 15, &HF1F1F1
    SetPixel lHDC, tmpw + 17, tmph + 16, &HE7C699
    SetPixel lHDC, tmpw + 18, tmph + 16, &HE8C69A
    SetPixel lHDC, tmpw + 19, tmph + 16, &HE7C496
    SetPixel lHDC, tmpw + 20, tmph + 16, &HE8C597
    SetPixel lHDC, tmpw + 21, tmph + 16, &HE5C294
    SetPixel lHDC, tmpw + 22, tmph + 16, &HE8C194
    SetPixel lHDC, tmpw + 23, tmph + 16, &HE6BF8E
    SetPixel lHDC, tmpw + 24, tmph + 16, &HE7BC8C
    SetPixel lHDC, tmpw + 25, tmph + 16, &HE7BB8A
    SetPixel lHDC, tmpw + 26, tmph + 16, &HE5B37F
    SetPixel lHDC, tmpw + 27, tmph + 16, &HE1A971
    SetPixel lHDC, tmpw + 28, tmph + 16, &HD79F67
    SetPixel lHDC, tmpw + 29, tmph + 16, &H8E6A40
    SetPixel lHDC, tmpw + 30, tmph + 16, &H8A8076
    SetPixel lHDC, tmpw + 31, tmph + 16, &HE2E2E2
    SetPixel lHDC, tmpw + 32, tmph + 16, &HF9F9F9
    SetPixel lHDC, tmpw + 17, tmph + 17, &HE2CC9D
    SetPixel lHDC, tmpw + 18, tmph + 17, &HE2CC9E
    SetPixel lHDC, tmpw + 19, tmph + 17, &HE3CF9C
    SetPixel lHDC, tmpw + 20, tmph + 17, &HDFCA98
    SetPixel lHDC, tmpw + 21, tmph + 17, &HE2CD9C
    SetPixel lHDC, tmpw + 22, tmph + 17, &HE4CC9C
    SetPixel lHDC, tmpw + 23, tmph + 17, &HE1C491
    SetPixel lHDC, tmpw + 24, tmph + 17, &HE1C18F
    SetPixel lHDC, tmpw + 25, tmph + 17, &HE4BD8C
    SetPixel lHDC, tmpw + 26, tmph + 17, &HDAB285
    SetPixel lHDC, tmpw + 27, tmph + 17, &HC7A582
    SetPixel lHDC, tmpw + 28, tmph + 17, &H806A56
    SetPixel lHDC, tmpw + 29, tmph + 17, &H676565
    SetPixel lHDC, tmpw + 30, tmph + 17, &HD2D2D2
    SetPixel lHDC, tmpw + 31, tmph + 17, &HF2F2F2
    SetPixel lHDC, tmpw + 32, tmph + 17, &HFEFEFE
    SetPixel lHDC, tmpw + 17, tmph + 18, &HE9D5A6
    SetPixel lHDC, tmpw + 18, tmph + 18, &HE9D5A6
    SetPixel lHDC, tmpw + 19, tmph + 18, &HE9D6A3
    SetPixel lHDC, tmpw + 20, tmph + 18, &HE9D7A6
    SetPixel lHDC, tmpw + 21, tmph + 18, &HE2CC9B
    SetPixel lHDC, tmpw + 22, tmph + 18, &HE8D0A0
    SetPixel lHDC, tmpw + 23, tmph + 18, &HE8CF9C
    SetPixel lHDC, tmpw + 24, tmph + 18, &HE7C795
    SetPixel lHDC, tmpw + 25, tmph + 18, &HE2BC8B
    SetPixel lHDC, tmpw + 26, tmph + 18, &HB68F63
    SetPixel lHDC, tmpw + 27, tmph + 18, &H886948
    SetPixel lHDC, tmpw + 28, tmph + 18, &H786A5D
    SetPixel lHDC, tmpw + 29, tmph + 18, &HC7C7C7
    SetPixel lHDC, tmpw + 30, tmph + 18, &HEBEBEB
    SetPixel lHDC, tmpw + 31, tmph + 18, &HFCFCFC
    SetPixel lHDC, tmpw + 17, tmph + 19, &HD7D2BA
    SetPixel lHDC, tmpw + 18, tmph + 19, &HD7D2BA
    SetPixel lHDC, tmpw + 19, tmph + 19, &HD7D1B9
    SetPixel lHDC, tmpw + 20, tmph + 19, &HD5CEB6
    SetPixel lHDC, tmpw + 21, tmph + 19, &HDBD3BB
    SetPixel lHDC, tmpw + 22, tmph + 19, &HC9C1AA
    SetPixel lHDC, tmpw + 23, tmph + 19, &HA9A28B
    SetPixel lHDC, tmpw + 24, tmph + 19, &H827E6C
    SetPixel lHDC, tmpw + 25, tmph + 19, &H6A665B
    SetPixel lHDC, tmpw + 26, tmph + 19, &H625F5A
    SetPixel lHDC, tmpw + 27, tmph + 19, &H8B8C8C
    SetPixel lHDC, tmpw + 28, tmph + 19, &HCDCDCD
    SetPixel lHDC, tmpw + 29, tmph + 19, &HE8E8E8
    SetPixel lHDC, tmpw + 30, tmph + 19, &HFAFAFA
    SetPixel lHDC, tmpw + 17, tmph + 20, &H5A563E
    SetPixel lHDC, tmpw + 18, tmph + 20, &H59543D
    SetPixel lHDC, tmpw + 19, tmph + 20, &H58513A
    SetPixel lHDC, tmpw + 20, tmph + 20, &H554F38
    SetPixel lHDC, tmpw + 21, tmph + 20, &H59513B
    SetPixel lHDC, tmpw + 22, tmph + 20, &H58513E
    SetPixel lHDC, tmpw + 23, tmph + 20, &H646053
    SetPixel lHDC, tmpw + 24, tmph + 20, &H7B7973
    SetPixel lHDC, tmpw + 25, tmph + 20, &HA2A19F
    SetPixel lHDC, tmpw + 26, tmph + 20, &HC5C5C5
    SetPixel lHDC, tmpw + 27, tmph + 20, &HDADADA
    SetPixel lHDC, tmpw + 28, tmph + 20, &HEDEDED
    SetPixel lHDC, tmpw + 29, tmph + 20, &HFAFAFA
    SetPixel lHDC, tmpw + 17, tmph + 21, &HC5C5C5
    SetPixel lHDC, tmpw + 18, tmph + 21, &HC5C5C5
    SetPixel lHDC, tmpw + 19, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 20, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 21, tmph + 21, &HC6C6C6
    SetPixel lHDC, tmpw + 22, tmph + 21, &HC9C9C9
    SetPixel lHDC, tmpw + 23, tmph + 21, &HCECECE
    SetPixel lHDC, tmpw + 24, tmph + 21, &HD7D7D7
    SetPixel lHDC, tmpw + 25, tmph + 21, &HE1E1E1
    SetPixel lHDC, tmpw + 26, tmph + 21, &HECECEC
    SetPixel lHDC, tmpw + 27, tmph + 21, &HF6F6F6
    SetPixel lHDC, tmpw + 28, tmph + 21, &HFDFDFD
    SetPixel lHDC, tmpw + 17, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 18, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 19, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 20, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 21, tmph + 22, &HECECEC
    SetPixel lHDC, tmpw + 22, tmph + 22, &HEDEDED
    SetPixel lHDC, tmpw + 23, tmph + 22, &HF0F0F0
    SetPixel lHDC, tmpw + 24, tmph + 22, &HF4F4F4
    SetPixel lHDC, tmpw + 25, tmph + 22, &HFAFAFA
    SetPixel lHDC, tmpw + 26, tmph + 22, &HFDFDFD

    tmph = 11
    tmph1 = lh - 10
    tmpw = lw - 34
    'Generar lineas intermedias
    APILine 0, tmph, 0, tmph1, &HF7F7F7
    APILine 1, tmph, 1, tmph1, &HA99D9B
    APILine 2, tmph, 2, tmph1, &H632D17
    APILine 3, tmph, 3, tmph1, &HA65A1D
    APILine 4, tmph, 4, tmph1, &HB96C2E
    APILine 5, tmph, 5, tmph1, &HBC7738
    APILine 6, tmph, 6, tmph1, &HC18242
    APILine 7, tmph, 7, tmph1, &HC2894E
    APILine 8, tmph, 8, tmph1, &HC18A52
    APILine 9, tmph, 9, tmph1, &HC59157
    APILine 10, tmph, 10, tmph1, &HC59159
    APILine 11, tmph, 11, tmph1, &HCC9863
    APILine 12, tmph, 12, tmph1, &HCC9665
    APILine 13, tmph, 13, tmph1, &HCB9767
    APILine 14, tmph, 14, tmph1, &HC99565
    APILine 15, tmph, 15, tmph1, &HCC9A6A
    APILine 16, tmph, 16, tmph1, &HCC9A6A
    APILine 17, tmph, 17, tmph1, &HCC9B6A
    APILine tmpw + 17, tmph, tmpw + 17, tmph1, &HCC9B6A
    APILine tmpw + 18, tmph, tmpw + 18, tmph1, &HCC9B6A
    APILine tmpw + 19, tmph, tmpw + 19, tmph1, &HCA9864
    APILine tmpw + 20, tmph, tmpw + 20, tmph1, &HC99763
    APILine tmpw + 21, tmph, tmpw + 21, tmph1, &HC99763
    APILine tmpw + 22, tmph, tmpw + 22, tmph1, &HCA9562
    APILine tmpw + 23, tmph, tmpw + 23, tmph1, &HC9945D
    APILine tmpw + 24, tmph, tmpw + 24, tmph1, &HC68E57
    APILine tmpw + 25, tmph, tmpw + 25, tmph1, &HC48C55
    APILine tmpw + 26, tmph, tmpw + 26, tmph1, &HC3884E
    APILine tmpw + 27, tmph, tmpw + 27, tmph1, &HBE823E
    APILine tmpw + 28, tmph, tmpw + 28, tmph1, &HB97634
    APILine tmpw + 29, tmph, tmpw + 29, tmph1, &HBD6D2D
    APILine tmpw + 30, tmph, tmpw + 30, tmph1, &HA8581C
    APILine tmpw + 31, tmph, tmpw + 31, tmph1, &H6D3319
    APILine tmpw + 32, tmph, tmpw + 32, tmph1, &HA49794
    APILine tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6

    'Lineas verticales
    APILine 17, 0, lw - 17, 0, &H3C090A
    APILine 17, 1, lw - 17, 1, &HDEC7BE
    APILine 17, 2, lw - 17, 2, &HD3BCB2
    APILine 17, 3, lw - 17, 3, &HD4B39A
    APILine 17, 4, lw - 17, 4, &HCCAC92
    APILine 17, 5, lw - 17, 5, &HCFA88A
    APILine 17, 6, lw - 17, 6, &HCCA587
    APILine 17, 7, lw - 17, 7, &HD2A77D
    APILine 17, 8, lw - 17, 8, &HB98E62
    APILine 17, 9, lw - 17, 9, &HC69464
    APILine 17, 10, lw - 17, 10, &HCC9B6A
    APILine 17, 11, lw - 17, 11, &HD0A175
    tmph = lh - 22
    APILine 17, tmph + 11, lw - 17, tmph + 11, &HD0A175
    APILine 17, tmph + 12, lw - 17, tmph + 12, &HDAAA7F
    APILine 17, tmph + 13, lw - 17, tmph + 13, &HDCB087
    APILine 17, tmph + 14, lw - 17, tmph + 14, &HE4B88E
    APILine 17, tmph + 15, lw - 17, tmph + 15, &HE1BF93
    APILine 17, tmph + 16, lw - 17, tmph + 16, &HE7C699
    APILine 17, tmph + 17, lw - 17, tmph + 17, &HE2CC9D
    APILine 17, tmph + 18, lw - 17, tmph + 18, &HE9D5A6
    APILine 17, tmph + 19, lw - 17, tmph + 19, &HD7D2BA
    APILine 17, tmph + 20, lw - 17, tmph + 20, &H5A563E
    APILine 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    APILine 17, tmph + 22, lw - 17, tmph + 22, &HECECEC

End Sub

Private Sub DrawPlastikButton(iState As isEstate)


Dim tmpcolor As Long


    Select Case iState
    Case statenormal, stateDefaulted
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H8)
        With UserControl
            .BackColor = tmpcolor
            DrawVGradient OffsetColor(tmpcolor, &HF), OffsetColor(tmpcolor, -&HF), 1, 1, .ScaleWidth - 2, .ScaleHeight - 3
            DrawVGradient OffsetColor(tmpcolor, &H15), OffsetColor(tmpcolor, -&H15), 1, 2, 2, .ScaleHeight - 5
            DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H20), .ScaleWidth - 2, 2, .ScaleWidth - 3, .ScaleHeight - 5
        End With 'UserControl
        tmpcolor = OffsetColor(tmpcolor, -&H60)

        With UserControl
            APILine 2, 0, .ScaleWidth - 2, 0, tmpcolor
            APILine 0, 2, 0, .ScaleHeight - 2, tmpcolor
            APILine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, tmpcolor
            APILine .ScaleWidth - 1, 0, .ScaleWidth - 1, .ScaleHeight, tmpcolor
            SetPixel .hdc, 1, 1, tmpcolor
            SetPixel .hdc, 1, .ScaleHeight - 2, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, .ScaleHeight - 2, tmpcolor
        End With
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H15)

        With UserControl
            SetPixel .hdc, 1, 0, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, 0, tmpcolor
            SetPixel .hdc, 0, 1, tmpcolor
            SetPixel .hdc, 0, .ScaleHeight - 2, tmpcolor
            SetPixel .hdc, 1, .ScaleHeight - 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, .ScaleHeight - 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 1, 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 1, .ScaleHeight - 2, tmpcolor
            APILine 2, 1, .ScaleWidth - 2, 1, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H15)
            APILine 2, .ScaleHeight - 2, .ScaleWidth - 2, .ScaleHeight - 2, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H25)
        End With

        If iState = stateDefaulted Or m_bFocused Then
            tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
            APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(tmpcolor, &H15)
            APILine 1, 2, UserControl.ScaleWidth - 1, 2, OffsetColor(tmpcolor, &H15)
            APILine 1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H5)
            APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H15)
        End If
        Exit Sub

    Case stateHot
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H18)

        With UserControl
            .BackColor = tmpcolor
            DrawVGradient OffsetColor(tmpcolor, &H10), OffsetColor(tmpcolor, -&H10), 1, 1, .ScaleWidth - 2, .ScaleHeight - 3
            DrawVGradient OffsetColor(tmpcolor, &H15), OffsetColor(tmpcolor, -&H15), 1, 2, 2, .ScaleHeight - 5
            DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H20), .ScaleWidth - 2, 2, .ScaleWidth - 3, .ScaleHeight - 5
        End With
        tmpcolor = OffsetColor(tmpcolor, -&H60)

        With UserControl
            APILine 2, 0, .ScaleWidth - 2, 0, tmpcolor
            APILine 0, 2, 0, .ScaleHeight - 2, tmpcolor
            APILine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, tmpcolor
            APILine .ScaleWidth - 1, 0, .ScaleWidth - 1, .ScaleHeight, tmpcolor
            SetPixel .hdc, 1, 1, tmpcolor
            SetPixel .hdc, 1, .ScaleHeight - 2, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, .ScaleHeight - 2, tmpcolor
        End With

        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H15)
        With UserControl
            SetPixel .hdc, 1, 0, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, 0, tmpcolor
            SetPixel .hdc, 0, 1, tmpcolor
            SetPixel .hdc, 0, .ScaleHeight - 2, tmpcolor
            SetPixel .hdc, 1, .ScaleHeight - 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, .ScaleHeight - 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 1, 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 1, .ScaleHeight - 2, tmpcolor
            APILine 2, 1, .ScaleWidth - 2, 1, OffsetColor(GetSysColor(COLOR_BTNFACE), &H15)
            APILine 2, .ScaleHeight - 2, .ScaleWidth - 2, .ScaleHeight - 2, OffsetColor(GetSysColor(COLOR_BTNFACE), -&H10)
        End With

        tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
        APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(tmpcolor, &H15)
        APILine 1, 2, UserControl.ScaleWidth - 1, 2, OffsetColor(tmpcolor, &H15)
        APILine 1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H5)
        APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H15)
        Exit Sub

    Case statePressed
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H9)
        DrawVGradient OffsetColor(tmpcolor, -&HF), tmpcolor, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
        DrawVGradient OffsetColor(tmpcolor, -&H15), OffsetColor(tmpcolor, &H15), UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5
        DrawVGradient OffsetColor(tmpcolor, -&H20), OffsetColor(tmpcolor, -&H5), 1, 2, 2, UserControl.ScaleHeight - 5
        tmpcolor = OffsetColor(tmpcolor, -&H60)

        With UserControl
            APILine 2, 0, .ScaleWidth - 2, 0, tmpcolor
            APILine 0, 2, 0, .ScaleHeight - 2, tmpcolor
            APILine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, tmpcolor
            APILine .ScaleWidth - 1, 0, .ScaleWidth - 1, .ScaleHeight, tmpcolor
            SetPixel .hdc, 1, 1, tmpcolor
            SetPixel .hdc, 1, .ScaleHeight - 2, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, .ScaleHeight - 2, tmpcolor
        End With

        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H15)
        SetPixel UserControl.hdc, 1, 0, tmpcolor
        SetPixel UserControl.hdc, UserControl.ScaleWidth - 2, 0, tmpcolor
        SetPixel UserControl.hdc, 0, 1, tmpcolor
        SetPixel UserControl.hdc, 0, UserControl.ScaleHeight - 2, tmpcolor
        SetPixel UserControl.hdc, 1, UserControl.ScaleHeight - 1, tmpcolor
        SetPixel UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
        SetPixel UserControl.hdc, UserControl.ScaleWidth - 1, 1, tmpcolor
        SetPixel UserControl.hdc, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2, tmpcolor
        APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H8)
        APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H15)
        Exit Sub

    Case statedisabled
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H12)
        With UserControl
            .BackColor = tmpcolor
            DrawVGradient OffsetColor(tmpcolor, &HF), OffsetColor(tmpcolor, -&HF), 1, 1, .ScaleWidth - 2, .ScaleHeight - 3
            DrawVGradient OffsetColor(tmpcolor, &H15), OffsetColor(tmpcolor, -&H15), 1, 2, 2, .ScaleHeight - 5
            DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H20), .ScaleWidth - 2, 2, .ScaleWidth - 3, .ScaleHeight - 5
        End With

        tmpcolor = OffsetColor(tmpcolor, -&H60)

        With UserControl
            APILine 2, 0, .ScaleWidth - 2, 0, tmpcolor
            APILine 0, 2, 0, .ScaleHeight - 2, tmpcolor
            APILine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, tmpcolor
            APILine .ScaleWidth - 1, 0, .ScaleWidth - 1, .ScaleHeight, tmpcolor
            SetPixel .hdc, 1, 1, tmpcolor
            SetPixel .hdc, 1, .ScaleHeight - 2, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, 1, tmpcolor
            SetPixel .hdc, .ScaleWidth - 2, .ScaleHeight - 2, tmpcolor
        End With

        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H5)
        SetPixel UserControl.hdc, 1, 0, tmpcolor
        SetPixel UserControl.hdc, UserControl.ScaleWidth - 2, 0, tmpcolor
        SetPixel UserControl.hdc, 0, 1, tmpcolor
        SetPixel UserControl.hdc, 0, UserControl.ScaleHeight - 2, tmpcolor
        SetPixel UserControl.hdc, 1, UserControl.ScaleHeight - 1, tmpcolor
        SetPixel UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
        SetPixel UserControl.hdc, UserControl.ScaleWidth - 1, 1, tmpcolor
        SetPixel UserControl.hdc, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2, tmpcolor
        APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H15)
        APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H25)
        Exit Sub
    End Select

End Sub

Private Sub DrawGalaxyButton(iState As isEstate)


Dim tmpcolor As Long


    If iState = statenormal Then
        tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
    Else
        tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, OffsetColor(GetSysColor(COLOR_BTNFACE), &HF))
    End If
    UserControl.BackColor = tmpcolor
    If iState = statePressed Then
        DrawVGradient OffsetColor(tmpcolor, -&HF), OffsetColor(tmpcolor, &HF), 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 6
        APILine 2, 1, UserControl.ScaleWidth - 3, 1, tmpcolor
        APILine 1, 2, 1, UserControl.ScaleHeight - 3, tmpcolor
        APILine 2, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, &H60)
        APILine UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, &H60)
    Else
        DrawVGradient OffsetColor(tmpcolor, &HF), OffsetColor(tmpcolor, -&HF), 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 6
        APILine 2, 1, UserControl.ScaleWidth - 3, 1, OffsetColor(tmpcolor, &H60)
        APILine 1, 2, 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, &H60)
        APILine 2, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, tmpcolor
        APILine UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, tmpcolor
    End If
    tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
    If iState <> statePressed Then
        tmpcolor = IIf(m_bFocused, OffsetColor(tmpcolor, -&H60), OffsetColor(tmpcolor, -&H30))
    Else
        tmpcolor = OffsetColor(tmpcolor, -&H30)
    End If

    With UserControl
        APILine 2, 0, .ScaleWidth - 3, 0, tmpcolor
        APILine 0, 2, 0, .ScaleHeight - 3, tmpcolor
        APILine 2, .ScaleHeight - 2, .ScaleWidth - 3, .ScaleHeight - 2, tmpcolor
        APILine .ScaleWidth - 2, 2, .ScaleWidth - 2, .ScaleHeight - 3, tmpcolor
        SetPixel .hdc, 1, 1, tmpcolor
        SetPixel .hdc, 1, .ScaleHeight - 3, tmpcolor
        SetPixel .hdc, .ScaleWidth - 3, 1, tmpcolor
        SetPixel .hdc, .ScaleWidth - 3, .ScaleHeight - 3, tmpcolor
    End With 'UserControl
    tmpcolor = OffsetColor(tmpcolor, &H15)
    APILine 3, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, tmpcolor
    APILine UserControl.ScaleWidth - 1, 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 4, tmpcolor
    APILine UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, UserControl.ScaleWidth, UserControl.ScaleHeight - 5, tmpcolor
    APILine UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 1, UserControl.ScaleWidth, UserControl.ScaleHeight - 4, OffsetColor(tmpcolor, &HF)

End Sub

Private Sub DrawKeramikButton(iState As isEstate)


Dim tmpcolor As Long


    If m_iState = statenormal Then
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&HF)
    ElseIf m_iState = stateHot Then
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE)), &H1)
    ElseIf m_iState = statePressed Then
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE)), &H18)
    Else
        tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &HF)
    End If
    With UserControl
        .BackColor = tmpcolor
        DrawVGradient OffsetColor(tmpcolor, &H20), OffsetColor(tmpcolor, -&H20), 5, 2, .ScaleWidth - 5, .ScaleHeight - 6
        'Left
        DrawVGradient OffsetColor(tmpcolor, &H80), OffsetColor(tmpcolor, -&H80), 0, 0, 1, .ScaleHeight - 4
        DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H25), 2, 4, 3, .ScaleHeight / 2 - 7
        DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H35), 3, 3, 4, .ScaleHeight / 2 - 7
        DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H25), 4, 2, 5, .ScaleHeight / 2 - 7
        DrawVGradient OffsetColor(tmpcolor, &H25), OffsetColor(tmpcolor, &H5), 5, 2, 6, .ScaleHeight / 2 - 4

        DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H20), 2, .ScaleHeight / 2 - 2, 3, .ScaleHeight / 2 - 5
        DrawVGradient OffsetColor(tmpcolor, -&H35), OffsetColor(tmpcolor, -&H20), 3, .ScaleHeight / 2 - 3, 4, .ScaleHeight / 2 - 3
        DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H20), 4, .ScaleHeight / 2 - 4, 5, .ScaleHeight / 2 - 1
        DrawVGradient OffsetColor(tmpcolor, &H5), OffsetColor(tmpcolor, -&H5), 5, .ScaleHeight / 2 - 4, 6, .ScaleHeight / 2 - 3

        DrawHGradient OffsetColor(tmpcolor, -&H35), OffsetColor(tmpcolor, -&H20), 5, .ScaleHeight - 5, 12, .ScaleHeight - 6
        DrawHGradient OffsetColor(tmpcolor, -&H38), OffsetColor(tmpcolor, -&H20), 4, .ScaleHeight - 6, 7, .ScaleHeight - 7
        DrawHGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H20), 3, .ScaleHeight - 7, 5, .ScaleHeight - 8
        'Right
        DrawVGradient OffsetColor(tmpcolor, &H80), OffsetColor(tmpcolor, -&H80), .ScaleWidth - 2, 0, .ScaleWidth - 1, .ScaleHeight - 4
        DrawVGradient OffsetColor(tmpcolor, &H80), OffsetColor(tmpcolor, -&H30), .ScaleWidth - 1, 0, .ScaleWidth, .ScaleHeight / 2
        DrawVGradient OffsetColor(tmpcolor, -&H40), OffsetColor(tmpcolor, -&H8), .ScaleWidth - 1, .ScaleHeight / 2, .ScaleWidth, .ScaleHeight / 2 - 4
        DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H25), .ScaleWidth - 4, 4, .ScaleWidth - 3, .ScaleHeight / 2 - 5
        DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H30), .ScaleWidth - 5, 3, .ScaleWidth - 4, .ScaleHeight / 2 - 4
        DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H25), .ScaleWidth - 6, 2, .ScaleWidth - 5, .ScaleHeight / 2 - 3
        DrawVGradient OffsetColor(tmpcolor, &H25), OffsetColor(tmpcolor, &H5), .ScaleWidth - 7, .ScaleWidth - 6, 6, .ScaleHeight / 2 - 4

        DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H15), .ScaleWidth - 4, .ScaleHeight / 2, .ScaleWidth - 3, .ScaleHeight / 2 - 5
        DrawVGradient OffsetColor(tmpcolor, -&H30), OffsetColor(tmpcolor, -&H25), .ScaleWidth - 5, .ScaleHeight / 2, .ScaleWidth - 4, .ScaleHeight / 2 - 5
        DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H35), .ScaleWidth - 6, .ScaleHeight / 2, .ScaleWidth - 5, .ScaleHeight / 2 - 4
        DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H25), .ScaleWidth - 7, .ScaleHeight / 2, .ScaleWidth - 6, .ScaleHeight / 2 - 4
        DrawHGradient OffsetColor(tmpcolor, -&H20), OffsetColor(tmpcolor, -&H35), .ScaleWidth - 15, .ScaleHeight - 4, .ScaleWidth - 7, .ScaleHeight - 3
        'top
        APILine 3, 0, .ScaleWidth - 3, 0, OffsetColor(tmpcolor, &H30)
        APILine 1, 1, .ScaleWidth - 1, 1, OffsetColor(tmpcolor, &H30)
        APILine 5, 1, .ScaleWidth - 5, 1, tmpcolor 'OffsetColor(tmpcolor, &H30)
        DrawHGradient OffsetColor(tmpcolor, &H20), tmpcolor, .ScaleWidth - 11, 2, .ScaleWidth - 4, 3
        DrawHGradient OffsetColor(tmpcolor, &H20), tmpcolor, .ScaleWidth - 10, 3, .ScaleWidth - 5, 4
        DrawHGradient OffsetColor(tmpcolor, &H20), tmpcolor, .ScaleWidth - 9, 4, .ScaleWidth - 6, 5
    End With 'UserControl
    APILine 6, 3, 7, 3, OffsetColor(tmpcolor, &H80)
    'bottom
    APILine 3, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 1, OffsetColor(tmpcolor, -&H10)
    APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H80)
    APILine 7, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 7, UserControl.ScaleHeight - 3, tmpcolor
    SetPixel UserControl.hdc, 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H70)
    SetPixel UserControl.hdc, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H70)
    APILine UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 4, OffsetColor(tmpcolor, -&H15)

End Sub

Private Sub CreateToolTip()


Dim lpRect    As RECT
Dim lWinStyle As Long


    ttip.lpStr = m_sToolTipText

    If m_lttHwnd <> 0 Then
        DestroyWindow m_lttHwnd
    End If

    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX

    If m_lToolTipType = TTBalloon Then
        lWinStyle = lWinStyle Or TTS_BALLOON
    End If

    m_lttHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hwnd, 0&, App.hInstance, 0&)
    SetWindowPos m_lttHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    GetClientRect UserControl.hwnd, lpRect

    With ttip
        If m_lttCentered Then
            .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
        Else
            .lFlags = TTF_SUBCLASS
        End If

        .lHwnd = UserControl.hwnd
        .lId = 0
        .hInstance = App.hInstance
        .lpRect = lpRect
    End With

    SendMessage m_lttHwnd, TTM_ADDTOOLA, 0&, ttip
    If m_sTooltiptitle <> vbNullString Or m_lToolTipIcon <> TTNoIcon Then
        SendMessage m_lttHwnd, TTM_SETTITLE, CLng(m_lToolTipIcon), ByVal m_sTooltiptitle
    End If
    If m_lttForeColor <> Empty Then
        SendMessage m_lttHwnd, TTM_SETTIPTEXTCOLOR, TranslateColor(m_lttForeColor), 0&
    End If

    If m_lttBackColor <> Empty Then
        SendMessage m_lttHwnd, TTM_SETTIPBKCOLOR, TranslateColor(m_lttBackColor), 0&
    End If

End Sub

Private Sub m_About_Click()

    m_About.Visible = False

End Sub

Private Sub m_About_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)


Dim tmprect  As RECT
Dim tmpcolor As Long

    With m_About
        SetRect tmprect, 290, 80, 360, 26
        tmpcolor = GetSysColor(COLOR_BTNFACE)
        DrawVGradientEx m_About.hdc, OffsetColor(tmpcolor, -&HF), OffsetColor(tmpcolor, &HF), tmprect.Left, tmprect.Top, tmprect.Right, tmprect.Bottom
        tmpcolor = GetSysColor(COLOR_BTNSHADOW)
        APILineEx .hdc, tmprect.Left + 2, tmprect.Top, tmprect.Right - 2, tmprect.Top, tmpcolor
        APILineEx .hdc, tmprect.Left + 2, tmprect.Top, tmprect.Left, tmprect.Top + 2, tmpcolor
        APILineEx .hdc, tmprect.Right - 2, tmprect.Top, tmprect.Right, tmprect.Top + 2, tmpcolor
        APILineEx .hdc, tmprect.Left, tmprect.Top + 2, tmprect.Left, tmprect.Top + tmprect.Bottom - 2, tmpcolor
        APILineEx .hdc, tmprect.Right, tmprect.Top + 2, tmprect.Right, tmprect.Top + tmprect.Bottom, tmpcolor
        APILineEx .hdc, tmprect.Left, tmprect.Top + tmprect.Bottom - 2, tmprect.Left + 2, tmprect.Top + tmprect.Bottom, tmpcolor
        APILineEx .hdc, tmprect.Right, tmprect.Top + tmprect.Bottom - 2, tmprect.Right - 3, tmprect.Top + tmprect.Bottom + 1, tmpcolor
        APILineEx .hdc, tmprect.Left + 2, tmprect.Top + tmprect.Bottom, tmprect.Right - 2, tmprect.Top + tmprect.Bottom, tmpcolor
        tmpcolor = GetSysColor(COLOR_BTNFACE)
        SetPixel .hdc, tmprect.Left, tmprect.Top, tmpcolor
        SetPixel .hdc, tmprect.Left + 1, tmprect.Top, tmpcolor
        SetPixel .hdc, tmprect.Left, tmprect.Top + 1, tmpcolor

        SetPixel .hdc, tmprect.Right, tmprect.Top, tmpcolor
        SetPixel .hdc, tmprect.Right - 1, tmprect.Top, tmpcolor
        SetPixel .hdc, tmprect.Right, tmprect.Top + 1, tmpcolor

        SetPixel .hdc, tmprect.Left, tmprect.Top + tmprect.Bottom, tmpcolor
        SetPixel .hdc, tmprect.Left + 1, tmprect.Top + tmprect.Bottom, tmpcolor
        SetPixel .hdc, tmprect.Left, tmprect.Top + tmprect.Bottom - 1, tmpcolor

        SetPixel .hdc, tmprect.Right, tmprect.Top + tmprect.Bottom, tmpcolor
        SetPixel .hdc, tmprect.Right - 1, tmprect.Top + tmprect.Bottom, tmpcolor
        SetPixel .hdc, tmprect.Right, tmprect.Top + tmprect.Bottom - 1, tmpcolor

        SetRect tmprect, 290, 80, 360, 106
        .FontSize = 8
        .ForeColor = vbBlue
        .FontUnderline = True
        DrawText .hdc, "Close", -1, tmprect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
    End With

End Sub

Private Sub m_About_Paint()


Dim lwformat As Long
Dim tmprect  As RECT
Dim tmpcolor As Long


    lwformat = DT_VCENTER Or DT_LEFT Or DT_SINGLELINE
    With m_About
        .ForeColor = GetSysColor(COLOR_BTNTEXT)
        .FontUnderline = False
        .FontSize = 18
        SetRect tmprect, 20, 10, 220, 40
        DrawText .hdc, "isButton", -1, tmprect, lwformat
        .FontSize = 10
        SetRect tmprect, 160, 70, 300, 20
        DrawText .hdc, "Version " & strCurrentVersion, -1, tmprect, lwformat
        .FontBold = True
        SetRect tmprect, 20, 110, 250, 20
        DrawText .hdc, "By Fred.cpp", -1, tmprect, lwformat
        .FontBold = False
        SetRect tmprect, 20, 140, 250, 20
        DrawText .hdc, "http://mx.geocities.com/fred_cpp/", -1, tmprect, lwformat
        'Draw button
        SetRect tmprect, 290, 80, 360, 26
        tmpcolor = GetSysColor(COLOR_BTNFACE)
        DrawVGradientEx m_About.hdc, OffsetColor(tmpcolor, &HF), OffsetColor(tmpcolor, -&HF), tmprect.Left, tmprect.Top, tmprect.Right, tmprect.Bottom
        tmpcolor = GetSysColor(COLOR_BTNSHADOW)
        APILineEx .hdc, tmprect.Left + 2, tmprect.Top, tmprect.Right - 2, tmprect.Top, tmpcolor
        APILineEx .hdc, tmprect.Left + 2, tmprect.Top, tmprect.Left, tmprect.Top + 2, tmpcolor
        APILineEx .hdc, tmprect.Right - 2, tmprect.Top, tmprect.Right, tmprect.Top + 2, tmpcolor
        APILineEx .hdc, tmprect.Left, tmprect.Top + 2, tmprect.Left, tmprect.Top + tmprect.Bottom - 2, tmpcolor
        APILineEx .hdc, tmprect.Right, tmprect.Top + 2, tmprect.Right, tmprect.Top + tmprect.Bottom, tmpcolor
        APILineEx .hdc, tmprect.Left, tmprect.Top + tmprect.Bottom - 2, tmprect.Left + 2, tmprect.Top + tmprect.Bottom, tmpcolor
        APILineEx .hdc, tmprect.Right, tmprect.Top + tmprect.Bottom - 2, tmprect.Right - 3, tmprect.Top + tmprect.Bottom + 1, tmpcolor
        APILineEx .hdc, tmprect.Left + 2, tmprect.Top + tmprect.Bottom, tmprect.Right - 2, tmprect.Top + tmprect.Bottom, tmpcolor
        tmpcolor = GetSysColor(COLOR_BTNFACE)
        SetPixel .hdc, tmprect.Left, tmprect.Top, tmpcolor
        SetPixel .hdc, tmprect.Left + 1, tmprect.Top, tmpcolor
        SetPixel .hdc, tmprect.Left, tmprect.Top + 1, tmpcolor
        SetPixel .hdc, tmprect.Right, tmprect.Top, tmpcolor
        SetPixel .hdc, tmprect.Right - 1, tmprect.Top, tmpcolor
        SetPixel .hdc, tmprect.Right, tmprect.Top + 1, tmpcolor
        SetPixel .hdc, tmprect.Left, tmprect.Top + tmprect.Bottom, tmpcolor
        SetPixel .hdc, tmprect.Left + 1, tmprect.Top + tmprect.Bottom, tmpcolor
        SetPixel .hdc, tmprect.Left, tmprect.Top + tmprect.Bottom - 1, tmpcolor
        SetPixel .hdc, tmprect.Right, tmprect.Top + tmprect.Bottom, tmpcolor
        SetPixel .hdc, tmprect.Right - 1, tmprect.Top + tmprect.Bottom, tmpcolor
        SetPixel .hdc, tmprect.Right, tmprect.Top + tmprect.Bottom - 1, tmpcolor
        SetRect tmprect, 290, 80, 360, 106
        .FontSize = 8
        .ForeColor = vbBlue
        .FontUnderline = True
        DrawText .hdc, "Close", -1, tmprect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
    End With

End Sub

Private Sub UserControl_Click()

    If lPrevButton = vbLeftButton Then
        If m_ButtonType = isbCheckBox Then
            m_Value = Not m_Value
            m_iState = stateHot
            Refresh
        End If
        RaiseEvent Click
    End If

End Sub
'/~ add this
Public Property Get BackDrop() As OLE_COLOR

    BackDrop = m_BDColor
    
End Property

Public Property Let BackDrop(ByVal m_BDColor As OLE_COLOR)

    Me.BackColor = m_BDColor
    PropertyChanged "BackDrop"
    UserControl_Resize
    Refresh
    
End Property

Private Sub UserControl_DblClick()

    If lPrevButton = vbLeftButton Then
        UserControl_MouseDown 1, 0, 1, 1
    End If

End Sub

Private Sub UserControl_ExitFocus()

    m_bFocused = False
    Refresh

End Sub

Private Sub UserControl_GotFocus()

    m_bFocused = True
    Refresh

End Sub

Private Sub UserControl_Hide()

    UserControl.Extender.ToolTipText = m_sToolTipText

End Sub

Private Sub UserControl_InitProperties()

    m_iStyle = 0
    m_sCaption = UserControl.Extender.Name
    m_IconSize = 16
    Set m_Icon = LoadPicture
    lwFontAlign = DT_CENTER Or DT_WORDBREAK
    m_bEnabled = True
    m_bShowFocus = False
    m_bUseCustomColors = False
    m_lBackColor = TranslateColor(vbButtonFace)
    m_lHighlightColor = TranslateColor(vbHighlight)
    m_lFontColor = TranslateColor(vbButtonText)
    m_lFontHighlightColor = TranslateColor(vbButtonText)
    lPrevStyle = GetWindowLong(m_About.hwnd, GWL_STYLE)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = 32 Then
        UserControl_MouseDown vbLeftButton, 0, 1, 1
    End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case 32
        RaiseEvent Click
        UserControl_Click
    End Select

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = 32 Then
        UserControl_MouseUp vbLeftButton, 0, 1, 1
    End If

End Sub

Private Sub UserControl_LostFocus()

    m_bFocused = False
    Refresh

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

    If Button = vbLeftButton Then
        m_iState = statePressed
        Refresh
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Public Sub Refresh()


Dim tmpcolor          As Long
Dim iStyleIconOffset  As Long
Dim bDrawThemeSuccess As Boolean
Dim ix                As Long
Dim iy                As Long
Dim ni                As Long
Dim nj                As Long


    UserControl.Cls
    iStyleIconOffset = 4
    If Not UserControl.Ambient.UserMode And m_iStyle <> isbWindowsXP Then
        m_iState = stateHot
    End If
    If m_ButtonType = isbCheckBox Then
        If m_Value Then
            m_iState = statePressed
        End If
    End If
    If Not m_bEnabled Then
        m_iState = statedisabled
        UserControl.BackColor = GetSysColor(COLOR_BTNFACE)
    End If
    Select Case m_iStyle
    Case isbNormal
        'Classic Style (Win98)
        If m_iState = statenormal Then
            tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        Else
            tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
        End If
        UserControl.BackColor = tmpcolor
        DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, IIf(m_iState = statePressed, EDGE_SUNKEN, EDGE_RAISED)

    Case isbSoft
        'Soft Style (I don't know where does It come, But I've seen this before)
        If m_iState = statenormal Or m_iState = statedisabled Then
            tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        Else
            tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
        End If
        UserControl.BackColor = tmpcolor
        Select Case m_iState
        Case statenormal, stateHot, stateDefaulted
            DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_RAISEDINNER
        Case statePressed
            DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER
        Case statedisabled
            APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
        End Select

    Case isbFlat
        'Flat Style (Office 2000 like)
        If m_iState = statenormal Or m_iState = statedisabled Then
            tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        Else
            tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
        End If
        UserControl.BackColor = tmpcolor
        If m_iState = statenormal Then
            'Normal (flat)
            tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
            UserControl.BackColor = tmpcolor
            APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
        ElseIf m_iState = stateHot Then
            'Hover
            'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
            'UserControl.BackColor = tmpColor
            DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_RAISEDINNER
        ElseIf m_iState = statePressed Then
            'Pushed
            'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
            'UserControl.BackColor = tmpColor
            DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER
        Else    'Disabled
            'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
            'UserControl.BackColor = tmpColor
            APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
        End If

    Case isbJava
        'Java Style
        UserControl.BackColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        Select Case m_iState
        Case statePressed
            tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNSHADOW))
        Case stateHot
            tmpcolor = IIf(m_bUseCustomColors, BlendColors(m_lHighlightColor, m_lBackColor), BlendColors(GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_BTNFACE)))
        Case Else
            tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        End Select
        APIFillRect UserControl.hdc, m_txtRect, tmpcolor
        DrawJavaBorder m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), tmpcolor

    Case isbOfficeXP
        'Redmon 2002 Office Suite ( ... )
        UserControl.BackColor = tmpcolor
        If m_iState = statenormal Then
            tmpcolor = MSOXPShiftColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H20)
            APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpcolor
        ElseIf m_iState = stateHot Then
            'Hover
            tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
            APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, MSOXPShiftColor(tmpcolor)
            APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
        ElseIf m_iState = statePressed Then
            'Pushed
            tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
            APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, MSOXPShiftColor(tmpcolor, &H80)
            APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
        Else
            'Disabled
            tmpcolor = MSOXPShiftColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H20)
            APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpcolor
        End If

    Case isbWindowsXP
        'WinXP (Emulated)
        If m_bUseCustomColors Then
            DrawCustomWinXPButton m_iState
        Else
            DrawWinXPButton m_iState
        End If
        iStyleIconOffset = 5

    Case isbWindowsTheme
        'Uses the current installed windows theme
        UserControl.BackColor = GetSysColor(COLOR_BTNFACE)
        If m_iState = statenormal And m_bFocused Then
            bDrawThemeSuccess = DrawTheme("Button", 1, stateDefaulted)
        Else
            bDrawThemeSuccess = DrawTheme("Button", 1, m_iState)
        End If
        If Not bDrawThemeSuccess Then
            m_iStyle = Me.NonThemeStyle
        End If

    Case isbPlastik
        DrawPlastikButton m_iState
        iStyleIconOffset = 5

    Case isbGalaxy
        DrawGalaxyButton m_iState

    Case isbKeramik
        DrawKeramikButton m_iState
        iStyleIconOffset = 7

    Case isbMacOSX
        DrawMacOSXButton
        iStyleIconOffset = 7

    End Select
    '''''DrawText
    DrawCaption

    ''''''Draw Icon
    If Not m_Icon Is Nothing Then
        If Icon <> 0 Then
            If m_IconAlign = isbCenter Then
                ix = (UserControl.ScaleWidth - m_IconSize) / 2
                iy = (UserControl.ScaleHeight - m_IconSize) / 2
            ElseIf m_IconAlign = isbBottom Then
                ix = (UserControl.ScaleWidth - m_IconSize) / 2
                iy = UserControl.ScaleHeight - m_IconSize - iStyleIconOffset
            ElseIf m_IconAlign = isbTop Then
                ix = (UserControl.ScaleWidth - m_IconSize) / 2
                iy = iStyleIconOffset
            ElseIf m_IconAlign = isbLeft Then
                ix = iStyleIconOffset
                iy = (UserControl.ScaleHeight - m_IconSize) / 2
            ElseIf m_IconAlign = isbRight Then
                ix = UserControl.ScaleWidth - m_IconSize - iStyleIconOffset
                iy = (UserControl.ScaleHeight - m_IconSize) / 2
            End If
            If m_iState = statePressed Then
                ix = ix + 1
                iy = iy + 1
            ElseIf m_iState = stateHot Then
                If m_iStyle = isbOfficeXP Then
                    'fDrawPicture m_Icon, ix, iy, True 'Doesn't support Scaled Images :(
                    'This is the cheap version. works fine even when should be sloooow
                    PaintPicture m_Icon, ix, iy, m_IconSize, m_IconSize
                    For nj = iy To iy + m_IconSize
                        For ni = ix To ix + m_IconSize
                            If GetPixel(UserControl.hdc, ni, nj) <> GetPixel(UserControl.hdc, 1, 1) Then
                                SetPixel UserControl.hdc, ni, nj, &H808080
                            End If
                        Next ni
                    Next nj
                    ix = ix - 2
                    iy = iy - 2
                End If
            End If
            PaintPicture m_Icon, ix, iy, m_IconSize, m_IconSize
        End If

    End If

End Sub

Private Sub BuildRegion()

    If m_lRegion Then
        DeleteObject m_lRegion
    End If
    Select Case m_iStyle
    Case isbMacOSX
        m_lRegion = CreateMacOSXButtonRegion
    Case isbWindowsXP, isbPlastik
        m_lRegion = CreateWinXPRegion
    Case isbGalaxy, isbKeramik
        m_lRegion = CreateGalaxyRegion
    Case Else
        m_lRegion = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
    End Select
    SetWindowRgn UserControl.hwnd, m_lRegion, True

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

    If Button = vbLeftButton Then
        If m_ButtonType = isbCheckBox And Not m_Value Then
            m_iState = statePressed
        Else
            m_iState = stateHot
        End If
        Refresh
    End If
    lPrevButton = Button

End Sub

Private Sub UserControl_Paint()

    Refresh

End Sub

Private Sub UserControl_Show()

    m_sToolTipText = UserControl.Extender.ToolTipText
    UserControl.Extender.ToolTipText = vbNullString
    UserControl_Paint

End Sub

Private Sub UserControl_Resize()

    If UserControl.Width < 300 Then
        UserControl.Width = 300
    End If
    If UserControl.Height < 300 Then
        UserControl.Height = 300
    End If

    SetRect m_btnRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    SetRect m_txtRect, 4, 4, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
    Select Case m_CaptionAlign
    Case isbCenter
        lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Case isbLeft
        lwFontAlign = DT_VCENTER Or DT_LEFT Or DT_SINGLELINE
    Case isbRight
        lwFontAlign = DT_VCENTER Or DT_RIGHT Or DT_SINGLELINE
    Case isbTop
        lwFontAlign = DT_CENTER Or DT_TOP Or DT_SINGLELINE
    Case isbBottom
        lwFontAlign = DT_CENTER Or DT_BOTTOM Or DT_SINGLELINE
    End Select
    BuildRegion
    Refresh

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_iState = statenormal
    With PropBag
        m_lFontColor = .ReadProperty("FontColor", GetSysColor(COLOR_BTNTEXT))
        m_lFontHighlightColor = .ReadProperty("FontHighlightColor", GetSysColor(COLOR_BTNTEXT))
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        m_IconSize = .ReadProperty("IconSize", 16)
        m_IconAlign = .ReadProperty("IconAlign", isbLeft)
        m_iNonThemeStyle = .ReadProperty("iNonThemeStyle", [isbWindowsXP])
        m_bEnabled = .ReadProperty("Enabled", True)
        m_bShowFocus = .ReadProperty("ShowFocus", False)
        m_bUseCustomColors = .ReadProperty("USeCustomColors", False)
        m_lBackColor = .ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
        m_lHighlightColor = .ReadProperty("HighlightColor", GetSysColor(COLOR_HIGHLIGHT))
        Set m_Icon = .ReadProperty("Icon", Nothing)
        m_iStyle = .ReadProperty("Style", 3)
        m_sToolTipText = .ReadProperty("ToolTipText", UserControl.Extender.ToolTipText)
        m_sTooltiptitle = .ReadProperty("Tooltiptitle", vbNullString)
        m_lToolTipIcon = .ReadProperty("ToolTipIcon", 0)
        m_lToolTipType = .ReadProperty("ToolTipType", TTBalloon)
        m_lttBackColor = .ReadProperty("ttBackColor", GetSysColor(COLOR_INFOTEXT))
        m_lttForeColor = .ReadProperty("ttForeColor", GetSysColor(COLOR_INFOBK))
        m_sCaption = .ReadProperty("Caption", "isButton")
        m_CaptionAlign = .ReadProperty("CaptionAlign", isbCenter)
        m_ButtonType = .ReadProperty("ButtonType", isbButton)
        m_Value = .ReadProperty("Value", False)
        m_BDColor = .ReadProperty("BackDrop", &HFFFFFF)
    End With
    If Ambient.UserMode Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If

        If bTrack Then
            With UserControl
                Subclass_Start .hwnd
                Subclass_AddMsg .hwnd, WM_MOUSEMOVE, MSG_AFTER
                Subclass_AddMsg .hwnd, WM_MOUSELEAVE, MSG_AFTER
                Subclass_AddMsg .hwnd, WM_THEMECHANGED, MSG_AFTER
                Subclass_AddMsg .hwnd, WM_SYSCOLORCHANGE, MSG_AFTER
            End With
        End If
    End If

End Sub

Private Sub UserControl_Terminate()

    On Error GoTo Catch
    Subclass_StopAll
Catch:

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Icon", m_Icon
        .WriteProperty "Style", m_iStyle, 3
        .WriteProperty "Caption", m_sCaption, ""
        .WriteProperty "CaptionAlign", m_CaptionAlign, 0
        .WriteProperty "IconSize", m_IconSize, 16
        .WriteProperty "IconAlign", m_IconAlign, 0
        .WriteProperty "iNonThemeStyle", m_iNonThemeStyle, isbWindowsXP
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "ShowFocus", m_bShowFocus, False
        .WriteProperty "USeCustomColors", m_bUseCustomColors, False
        .WriteProperty "BackColor", m_lBackColor, GetSysColor(COLOR_BTNFACE)
        .WriteProperty "HighlightColor", m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT)
        .WriteProperty "FontColor", m_lFontColor, GetSysColor(COLOR_BTNTEXT)
        .WriteProperty "FontHighlightColor", m_lFontHighlightColor, GetSysColor(COLOR_BTNTEXT)
        .WriteProperty "ToolTipText", m_sToolTipText, UserControl.Extender.ToolTipText
        .WriteProperty "Tooltiptitle", m_sTooltiptitle
        .WriteProperty "ToolTipIcon", m_lToolTipIcon
        .WriteProperty "ToolTipType", m_lToolTipType
        .WriteProperty "ttBackColor", m_lttBackColor, GetSysColor(COLOR_INFOTEXT)
        .WriteProperty "ttForeColor", m_lttForeColor, GetSysColor(COLOR_INFOBK)
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "ButtonType", m_ButtonType, isbButton
        .WriteProperty "Value", m_Value, False
        .WriteProperty "BackDrop", m_BDColor, &HFFFFFF
    End With

End Sub

Private Function IsFunctionExported(ByVal sFunction As String, _
                                    ByVal sModule As String) As Boolean


Dim hMod       As Long
Dim bLibLoaded As Boolean


    hMod = GetModuleHandleA(sModule)

    If hMod = 0 Then
        hMod = LoadLibraryA(sModule)
        If hMod Then
            bLibLoaded = True
        End If
    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If

    If bLibLoaded Then
        FreeLibrary hMod
    End If

End Function

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)


Dim tme As TRACKMOUSEEVENT_STRUCT


    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If bTrackUser32 Then
            TrackMouseEvent tme
        Else
            TrackMouseEventComCtl tme
        End If
    End If

End Sub

Public Property Get Style() As isbStyle

    Style = m_iStyle

End Property

Public Property Let Style(ByVal NewStyle As isbStyle)

    m_iStyle = NewStyle
    PropertyChanged "Style"
    UserControl_Resize
    UserControl_Paint

End Property

Public Property Get Caption() As String

    Caption = m_sCaption

End Property

Public Property Let Caption(ByVal NewCaption As String)

    m_sCaption = NewCaption
    PropertyChanged "Caption"
    UserControl_Resize
    Refresh

End Property

Public Property Get Icon() As StdPicture

    Set Icon = m_Icon

End Property

Public Property Set Icon(NewIcon As StdPicture)

    Set m_Icon = NewIcon
    PropertyChanged "Icon"
    UserControl_Resize
    UserControl_Paint

End Property

Public Property Get IconAlign() As isbAlign

    IconAlign = m_IconAlign

End Property

Public Property Let IconAlign(ByVal NewIconAlign As isbAlign)

    m_IconAlign = NewIconAlign
    PropertyChanged "IconAlign"
    UserControl_Resize
    UserControl_Paint

End Property

Public Property Get IconSize() As Integer

    IconSize = m_IconSize

End Property

Public Property Let IconSize(ByVal NewIconSize As Integer)

    m_IconSize = NewIconSize
    PropertyChanged "IconSize"
    UserControl_Resize
    UserControl_Paint

End Property

Public Property Get CaptionAlign() As isbAlign

    CaptionAlign = m_CaptionAlign

End Property

Public Property Let CaptionAlign(ByVal NewCaptionAlign As isbAlign)

    m_CaptionAlign = NewCaptionAlign
    PropertyChanged "CaptionAlign"
    UserControl_Resize
    UserControl_Paint

End Property

Public Property Get NonThemeStyle() As isbStyle

    NonThemeStyle = m_iNonThemeStyle

End Property

Public Property Let NonThemeStyle(ByVal NewNonThemeStyle As isbStyle)

    m_iNonThemeStyle = NewNonThemeStyle
    PropertyChanged "NonThemeStyle"
    UserControl_Resize
    UserControl_Paint

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_bEnabled
    Refresh

End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)

    m_bEnabled = bEnabled
    m_iState = statenormal
    Refresh
    PropertyChanged "Enabled"
    UserControl.Enabled = m_bEnabled

End Property

Public Property Get ShowFocus() As Boolean

    ShowFocus = m_bShowFocus

End Property

Public Property Let ShowFocus(ByVal bShowFocus As Boolean)

    m_bShowFocus = bShowFocus
    PropertyChanged "ShowFocus"
    Refresh

End Property

Public Property Get UseCustomColors() As Boolean

    UseCustomColors = m_bUseCustomColors

End Property

Public Property Let UseCustomColors(ByVal bUseCustomColors As Boolean)

    m_bUseCustomColors = bUseCustomColors
    PropertyChanged "UseCustomColors"
    Refresh

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = m_lBackColor

End Property

Public Property Let BackColor(lBackColor As OLE_COLOR)

    m_lBackColor = lBackColor
    PropertyChanged "BackColor"
    Refresh

End Property

Public Property Get HighlightColor() As OLE_COLOR

    HighlightColor = m_lHighlightColor

End Property

Public Property Let HighlightColor(lHighlightColor As OLE_COLOR)

    m_lHighlightColor = lHighlightColor
    PropertyChanged "HighlightColor"
    Refresh

End Property

Public Property Get FontColor() As OLE_COLOR

    FontColor = m_lFontColor

End Property

Public Property Let FontColor(lFontColor As OLE_COLOR)

    m_lFontColor = lFontColor
    PropertyChanged "FontColor"
    Refresh

End Property

Public Property Get FontHighlightColor() As OLE_COLOR

    FontHighlightColor = m_lFontHighlightColor

End Property

Public Property Let FontHighlightColor(lFontHighlightColor As OLE_COLOR)

    m_lFontHighlightColor = lFontHighlightColor
    PropertyChanged "FontHighlightColor"
    Refresh

End Property

Public Property Get ToolTip() As String

    ToolTip = m_sToolTipText

End Property

Public Property Let ToolTip(ByVal sToolTipText As String)

    m_sToolTipText = sToolTipText
    PropertyChanged "ToolTipText"
    Refresh

End Property

Public Property Get ToolTipTitle() As String

    ToolTipTitle = m_sTooltiptitle

End Property

Public Property Let ToolTipTitle(ByVal sTooltipTitle As String)

    m_sTooltiptitle = sTooltipTitle
    PropertyChanged "TooltipTitle"
    Refresh

End Property

Public Property Get ToolTipIcon() As ttIconType

    ToolTipIcon = m_lToolTipIcon

End Property

Public Property Let ToolTipIcon(lTooltipIcon As ttIconType)

    m_lToolTipIcon = lTooltipIcon
    PropertyChanged "TooltipIcon"
    Refresh

End Property

Public Property Get ToolTipType() As ttStyleEnum

    ToolTipType = m_lToolTipType

End Property

Public Property Let ToolTipType(lToolTipType As ttStyleEnum)

    m_lToolTipType = lToolTipType
    PropertyChanged "ToolTipType"
    Refresh

End Property

Public Property Get ToolTipBackColor() As OLE_COLOR

    ToolTipType = m_lttBackColor

End Property

Public Property Let ToolTipBackColor(lToolTipBackColor As OLE_COLOR)

    m_lttBackColor = lToolTipBackColor
    PropertyChanged "ToolTipBackColor"
    Refresh

End Property

Public Property Get ToolTipForeColor() As OLE_COLOR

    ToolTipForeColor = m_lttForeColor

End Property

Public Property Let ToolTipForeColor(lToolTipForeColor As OLE_COLOR)

    m_lttForeColor = lToolTipForeColor
    PropertyChanged "ToolTipForeColor"
    Refresh

End Property

Public Property Get Font() As StdFont

    Set Font = UserControl.Font

End Property

Public Property Set Font(newFont As StdFont)

    Set m_Font = newFont
    Set UserControl.Font = newFont
    Refresh
    PropertyChanged "Font"

End Property

Public Property Get ButtonType() As isbButtonType

    ButtonType = m_ButtonType

End Property

Public Property Let ButtonType(newType As isbButtonType)

    m_ButtonType = newType
    PropertyChanged "ButtonType"
    Refresh

End Property

Public Property Get Value() As Boolean

    Value = m_Value

End Property

Public Property Let Value(ByVal NewValue As Boolean)

    m_Value = NewValue
    PropertyChanged "Value"
    Refresh

End Property

Public Function Version() As String

    Version = strCurrentVersion

End Function

Public Function OpenLink(sLink As String) As Long

    OpenLink = ShellExecute(hwnd, "open", sLink, vbNull, vbNull, 1)

End Function

Public Sub About()

    With m_About
        .Visible = True
        SetWindowLong .hwnd, GWL_STYLE, lPrevStyle + WS_CAPTION + WS_THICKFRAME + WS_MINIMIZEBOX
        SetWindowPos .hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        SetWindowPos .hwnd, 0, 0, 0, 0, 0, SWP_REFRESH
        SetWindowText .hwnd, "About isButton " & strCurrentVersion
        SetWindowPos .hwnd, 0, 0, 0, 0, 0, SWP_REFRESH
        SetParent .hwnd, 0
    End With

End Sub
