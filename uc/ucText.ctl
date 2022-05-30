VERSION 5.00
Begin VB.UserControl ucText 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   OLEDropMode     =   1  'Manual
   PropertyPages   =   "ucText.ctx":0000
   ScaleHeight     =   18
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   177
   ToolboxBitmap   =   "ucText.ctx":0014
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   0
   End
End
Attribute VB_Name = "ucText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------
'Autor: Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 16/04/2022
'Version: 0.0.1
'Thanks: Latin Group of VB6
'-----------------------------
'23/04/2022 Revision 1.0.1
            'Ajust keypress filters in InputType=IT_Date and IT_Time
            'Fix same errors in Property page
               
'--------------------------SubClass by wqweto
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True

'--- for Modern Subclassing Thunk (MST)
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000
Private Const PTR_SIZE                      As Long = 4
Private Const EBMODE_DESIGN                 As Long = 0
'--- end MST

'--- for Modern Subclassing Thunk (MST)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
#If ImplSelfContained Then
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If
'--- end MST

'KERNEL32
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
'GDI32
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
'USER32
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function CopyImage Lib "user32.dll" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
'GDI+
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipFillEllipse Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipAddPathPolygon Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTF, ByVal mCount As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As CombineMode) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipAddPathArc Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathRectangle Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, ByRef Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipAddPathBezier Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single, ByVal mX3 As Single, ByVal mY3 As Single, ByVal mX4 As Single, ByVal mY4 As Single) As Long
Private Declare Function GdipAddPathEllipse Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mInterpolationMode As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByRef mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mRfType As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetClipPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipCloneImage Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mCloneImage As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mMode As Long, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipGetImageDimension Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Single, ByRef mHeight As Single) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixelOffSetMode As Long) As Long

'OTHERS
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByVal pcbBinary As Long, ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "Ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Private Const IDC_IBEAM                 As Long = 32513&
Private Const IDC_HAND                  As Long = 32649
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const LinearGradientModeVertical As Long = &H1
Private Const InterpolationModeHighQualityBicubic = &H7
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const OBJ_BITMAP                As Long = 7
Private Const RotateNoneFlipY           As Long = &H6
Private Const IMAGE_BITMAP              As Long = 0
Private Const LR_CREATEDIBSECTION       As Long = &H2000
Private Const DT_CALCRECT               As Long = &H400
Private Const DT_SINGLELINE             As Long = &H20
Private Const TME_LEAVE                 As Long = &H2&
Private Const RDW_INVALIDATE            As Long = &H1
Private Const TRANSPARENT               As Long = 1
Private Const KEYEVENTF_KEYDOWN         As Integer = &H0
Private Const KEYEVENTF_KEYUP           As Integer = &H2

'Mouse & Key Event
Private Const WM_SETFOCUS           As Long = &H7
Private Const WM_KILLFOCUS          As Long = &H8
Private Const WM_LBUTTONUP          As Long = &H202
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_RBUTTONDBLCLK      As Long = &H206
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_RBUTTONUP          As Long = &H205
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_CHAR               As Long = &H102
Private Const WM_VSCROLL            As Long = &H115
Private Const WM_KEYUP              As Long = &H101
Private Const WM_KEYDOWN            As Long = &H100
Private Const WM_SETTEXT            As Long = &HC
Private Const WM_GETTEXT            As Long = &HD
Private Const WM_GETTEXTLENGTH      As Long = &HE
Private Const WM_GETFONT            As Long = &H31
Private Const WM_SETFONT            As Long = &H30
Private Const WM_CONTEXTMENU        As Long = &H7B
Private Const WM_PASTE              As Long = &H302
Private Const WM_CUT                As Long = &H300
Private Const WM_COMMAND            As Long = &H111
Private Const WM_MOUSEWHEEL         As Long = &H20A
Private Const WM_ERASEBKGND         As Long = &H14
Private Const WM_HSCROLL            As Long = &H114
Private Const WM_CTLCOLORSTATIC     As Long = &H138
Private Const WM_CTLCOLOREDIT       As Long = &H133
Private Const WM_SIZE               As Long = &H5
Private Const WM_GETMINMAXINFO      As Long = &H24
Private Const WM_NCHITTEST          As Long = &H84
Private Const WM_SYSCOMMAND         As Long = &H112

Private Const EN_CHANGE As Long = &H300
Private Const EN_MAXTEXT As Long = &H501
Private Const EN_HSCROLL As Long = &H601
Private Const EN_VSCROLL As Long = &H602
Private Const EN_UPDATE As Long = &H400
Private Const EN_KILLFOCUS As Long = &H200
Private Const EN_SETFOCUS As Long = &H100

'Style
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_CLIPSIBLINGS = &H4000000

Private Const ES_MULTILINE As Long = &H4&
Private Const ES_WANTRETURN As Long = &H1000&
Private Const ES_AUTOVSCROLL As Long = &H40&
Private Const ES_AUTOHSCROLL As Long = &H80&
Private Const ES_READONLY As Long = &H800&
Private Const ES_CENTER& = &H1&
Private Const ES_LEFT& = &H0&
Private Const ES_RIGHT& = &H2&
Private Const ES_PASSWORD As Long = &H20&

'Private Const WS_MAXIMIZEBOX = &H10000
Private Const ES_NOHIDESEL As Long = &H100&
Private Const ES_NUMBER As Long = &H2000&
Private Const ES_LOWERCASE As Long = &H10&
Private Const ES_UPPERCASE As Long = &H8&
Private Const ECM_FIRST As Long = &H1500
Private Const EM_SETCUEBANNER As Long = (ECM_FIRST + 1)
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)
Private Const EM_LINELENGTH As Long = &HC1

'Private Const VK_RETURN As Long = &HD
'Private Const VK_SPACE As Long = &H20

Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_SETREADONLY As Long = &HCF
Private Const EM_GETSEL As Long = &HB0
'Private Const EM_GETLINE As Long = &HC4
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETSEL As Long = &HB1
Private Const EM_REPLACESEL As Long = &HC2
'Private Const WM_USER As Long = &H400
'Private Const EM_GETTEXTRANGE As Long = (WM_USER + 75)
Private Const SB_BOTTOM As Long = 7
Private Const SC_DRAGSIZE_SE As Long = &HF008&
Private Const HTBOTTOMRIGHT         As Long = 17

'[Events]
Public Event Change()
Public Event Click()
Public Event DbClick()
'Public Event SelChange()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event ContextMenu(ByRef Cancel As Boolean)
Public Event Paste(ByRef Cancel As Boolean)
Public Event Cut(ByRef Cancel As Boolean)
Public Event ImgLeftMouseDown(Button, Shift As Integer, X As Single, Y As Single)
Public Event ImgLeftMouseUp(Button, Shift As Integer, X As Single, Y As Single)
Public Event ImgRightMouseDown(Button, Shift As Integer, X As Single, Y As Single)
Public Event ImgRightMouseUp(Button, Shift As Integer, X As Single, Y As Single)
Public Event DropDown(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Public Event MouseWhell(Value As Integer)
Public Event Scroll()
Public Event Resize()
Public Event EndSize()
'---------------------------------- Enums -------------------------------
Private Type EDITBALLOONTIP
    cbStruct As Long
    pszTitle As Long
    pszText As Long
    iIcon As Long
End Type

Private Type POINTF
    X As Single
    Y As Single
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved                          As POINTAPI
    ptMaxSize                           As POINTAPI
    ptMaxPosition                       As POINTAPI
    ptMinTrackSize                      As POINTAPI
    ptMaxTrackSize                      As POINTAPI
End Type


Public Enum eTTInfo
    TTI_NONE = 0
    TTI_INFO = 1
    TTI_WARNING = 2
    TTI_ERROR = 3
End Enum

Private Type ICONINFO
    fIcon           As Long
    xHotspot        As Long
    yHotspot        As Long
    hbmMask         As Long
    hbmColor        As Long
End Type

Private Type BITMAP
  bmType                    As Long
  bmWidth                   As Long
  bmHeight                  As Long
  bmWidthBytes              As Long
  bmPlanes                  As Integer
  bmBitsPixel               As Integer
  bmBits                    As Long
End Type

Private Enum CombineMode
    CombineModeReplace = &H0
    CombineModeIntersect = &H1
    CombineModeUnion = &H2
    CombineModeXor = &H3
    CombineModeExclude = &H4
    CombineModeComplement = &H5
End Enum

Private Enum WrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Public Enum SCROLL_BAR
    None = 0
    Horizontal = 1
    Vertical = 2
    both = 3
End Enum

Public Enum TEXT_CONVERT
    [None] = 0
    [LowerCase] = 1
    [UpperCase] = 2
End Enum

Public Enum BORDER_STYLE
    [None] = 0
    [Fixed Single] = 1
    [Line Bottom]
End Enum

Public Enum TEXT_ALIGN
    [Left Justify] = 0
    [Right Justify] = 1
    [Center] = 2
End Enum

Public Enum eInputType
    [IT_Text] = 0
    [IT_Numeric] = 1
    [IT_Date] = 2
    [IT_Time] = 3
    [IT_PasswordChar] = 4
    [IT_MultiLine] = 5
    [IT_Desimal] = 6
    [IT_LettersOnly] = 7
    [IT_DropDown] = 8
End Enum

Public Enum eRightButtonStyle
    [RS_None] = 0
    [RS_Resizable] = 1
    [RS_Icon] = 2
    [RS_SpinButton] = 3
    [RS_ClearText] = 4
    [RS_ShowPassword] = 5
    [RS_DropDown] = 6
End Enum

Public Enum eFillStyle
    [FS_Solid] = 0
    [FS_Tansparent] = 1
End Enum

Private Enum RegionalConstant
  LOCALE_SCURRENCY = &H14
  LOCALE_SCOUNTRY = &H6
  LOCALE_SDATE = &H1D
  LOCALE_SDECIMAL = &HE
  LOCALE_SLANGUAGE = &H2
  LOCALE_SLONGDATE = &H20
  LOCALE_SMONDECIMALSEP = &H16
  LOCALE_SMONGROUPING = &H18
  LOCALE_SMONTHOUSANDSEP = &H17
  LOCALE_SNATIVECTRYNAME = &H8
  LOCALE_SNATIVECURRNAME = &H1008
  LOCALE_SNATIVEDIGITS = &H13
  LOCALE_SNEGATIVESIGN = &H51
  LOCALE_SSHORTDATE = &H1F
  LOCALE_STIME = &H1E
  LOCALE_STIMEFORMAT = &H1003
End Enum

Private m_Text              As String
Private m_ReadOnly          As Boolean
Private m_ScrollBar         As SCROLL_BAR
Private m_BorderStyle       As BORDER_STYLE
Private m_TextAlign         As TEXT_ALIGN
Private m_Enabled           As Boolean
Private m_hideSel           As Boolean
Private m_TextConvert       As TEXT_CONVERT
Private m_MaxLength            As Long
Private m_SelLength         As Long
Private m_SelStart          As Long
Private m_MarginLeft        As Long
Private m_MarginRight       As Long
Private m_BorderWidth       As Long
Private m_BorderRadius      As Long
Private m_ImgRightSize      As Long
Private m_InputType         As eInputType
Private m_RightButtonStyle  As eRightButtonStyle
Private m_BackColor         As OLE_COLOR
Private m_BackColorOnFocus  As OLE_COLOR
Private m_BorderColor       As OLE_COLOR
Private m_BorderColorOnFocus As OLE_COLOR
Private m_ParentBackColor   As OLE_COLOR
Private m_Caption           As String
Private m_OnFocusSelAll     As Boolean
Private m_OnFocusBigBorder  As Boolean
Private m_CueBanner         As String
Private m_ShortDateFormat   As String
Private m_Rect              As RECT
Private m_RectEdit          As RECT
Private m_EditGradient      As Boolean
Private m_ButtonsGradient   As Boolean
Private m_ImgLeftFillStyle  As eFillStyle
Private m_ImgRightFillStyle As eFillStyle
Private m_ImgLeftFillColor  As OLE_COLOR
Private m_ImgRightFillColor As OLE_COLOR
Private m_ImgLeftShowMouseEvents As Boolean
Private m_ImgRightShowMouseEvents As Boolean
Private m_OnKeyReturnTabulate As Boolean
Private m_MinSize           As POINTAPI
Private m_MaxSize           As POINTAPI
Private m_HotBorder         As Boolean
Private m_MinValue          As Variant
Private m_MaxValue          As Variant

Private hEdit               As Long
Private Margin              As Long
Private GdipToken           As Long
Private nScale              As Double
Private ArrImgLeft()        As Byte
Private hImgLeft            As Long
Private ImgLeftRealSize     As POINTF

Private m_ImgLeftSize       As Long
Private ArrImgRight()       As Byte
Private hImgRight           As Long
Private ImgRightRealSize    As POINTF
Private HotSpinButton       As Long
Private bMouseDown          As Boolean
Private bMouseEnter         As Boolean
Private bFocus              As Boolean
Private bShowPassword       As Boolean
Private bVisible            As Boolean
Private CaptionHeight       As Long
Private CaptionWidth        As Long
Private HotButton           As Long
Private BB                  As Single 'BigBorder
Private TextBoxChangeFrozen As Boolean
Private m_uIPAO             As IPAOHookStruct
Private m_pSubclassUC       As IUnknown
Private m_pSubclassEdit     As IUnknown
Private IsSubClased         As Boolean
Private ObjListPlus         As Object
Private hCurIBEAM           As Long
Private mDate               As Date
Private sDecimal            As String
Private sShortDate          As String
Private sThousand           As String
Private sDateDiv            As String
Private sMoney              As String
Private mBkBrush            As Long

'---------------------------------------------------
'Properties
'---------------------------------------------------
'*1
Public Property Get OnKeyReturnTabulate() As Boolean
    OnKeyReturnTabulate = m_OnKeyReturnTabulate
End Property

Public Property Let OnKeyReturnTabulate(New_Value As Boolean)
    m_OnKeyReturnTabulate = New_Value
    PropertyChanged "OnKeyReturnTabulate"
End Property

Public Property Get EditGradient() As Boolean
    EditGradient = m_EditGradient
End Property

Public Property Let EditGradient(New_Value As Boolean)
    m_EditGradient = New_Value
    PropertyChanged "EditGradient"
    UserControl_Resize
End Property

Public Property Get ButtonsGradient() As Boolean
    ButtonsGradient = m_ButtonsGradient
End Property

Public Property Let ButtonsGradient(New_Value As Boolean)
    m_ButtonsGradient = New_Value
    PropertyChanged "ButtonsGradient"
    Draw
End Property

Public Property Get TextConvert() As TEXT_CONVERT
    TextConvert = m_TextConvert
End Property

Public Property Let TextConvert(New_Value As TEXT_CONVERT)
    m_TextConvert = New_Value
    PropertyChanged "TextConvert"
    DoRefresh
End Property

Public Property Get SelLength() As Long
    Dim Pos As Long
    Pos = SendMessage(hEdit, EM_GETSEL, 0&, 0&)
    SelLength = WordHi(Pos) - WordLo(Pos)
End Property

Public Property Let SelLength(ByVal New_Value As Long)
    m_SelLength = New_Value
    Call SendMessage(hEdit, EM_SETSEL, m_SelStart, ByVal m_SelLength + m_SelStart)
    PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
    Dim Pos As Long
    Pos = SendMessage(hEdit, EM_GETSEL, 0&, 0&)
    SelStart = WordLo(Pos)
End Property

Public Property Let SelStart(ByVal New_Value As Long)
    m_SelStart = New_Value
    Call SendMessage(hEdit, EM_SETSEL, m_SelStart, ByVal m_SelLength + m_SelStart)
    PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
    m_Text = TextBox_GetText(hEdit)
    Dim Pos As Long
    Pos = SendMessage(hEdit, EM_GETSEL, 0&, 0&)
    SelText = Mid(m_Text, WordLo(Pos) + 1, WordHi(Pos) - WordLo(Pos))
End Property

Public Property Let SelText(ByVal New_Value As String)
    Call SendMessageW(hEdit, EM_REPLACESEL, 0&, StrPtr(New_Value))
    PropertyChanged "SelText"
End Property

Public Property Get MaxLength() As Long
    MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(New_Value As Long)
    m_MaxLength = New_Value
    Call SendMessage(hEdit, EM_LIMITTEXT, m_MaxLength, 0&)
    PropertyChanged "MaxLength"
End Property

Public Property Get HideSelection() As Boolean
    HideSelection = m_hideSel
End Property

Public Property Let HideSelection(New_Value As Boolean)
    m_hideSel = New_Value
    PropertyChanged "HideSelection"
    DoRefresh
End Property
    
Public Property Get Alignment() As TEXT_ALIGN
    Alignment = m_TextAlign
End Property

Public Property Let Alignment(New_Align As TEXT_ALIGN)
    m_TextAlign = New_Align
    PropertyChanged "Alignment"
    DoRefresh
End Property

Public Property Get Locked() As Boolean
     Locked = m_ReadOnly
End Property

Public Property Let Locked(New_Value As Boolean)
    m_ReadOnly = New_Value
    PropertyChanged "Locked"
    'DoRefresh
    SendMessage hEdit, EM_SETREADONLY, m_ReadOnly, 0&
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(New_Value As Boolean)
    m_Enabled = New_Value
    UserControl.Enabled = New_Value
    EnableWindow hEdit, New_Value
    PropertyChanged "Enabled"
    DoRefresh
End Property

Public Property Get Scrollbar() As SCROLL_BAR
    Scrollbar = m_ScrollBar
End Property

Public Property Let Scrollbar(New_ScrollBar As SCROLL_BAR)
    m_ScrollBar = New_ScrollBar
    PropertyChanged "Scrollbar"
    DoRefresh
End Property

Public Property Get HotBorder() As Boolean
    HotBorder = m_HotBorder
End Property

Public Property Let HotBorder(New_Value As Boolean)
    m_HotBorder = New_Value
    PropertyChanged "HotBorder"
End Property

Public Property Get TextLenght() As Long
    If hEdit <> 0 Then TextLenght = SendMessage(hEdit, WM_GETTEXTLENGTH, 0, ByVal 0&)
End Property

Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Texto"
Attribute Text.VB_UserMemId = -517
    m_Text = TextBox_GetText(hEdit)
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Value As String)
    If (m_InputType = IT_Date) Or (m_InputType = IT_Time) Then
        m_Text = Format$(New_Value, sShortDate)
    Else
        m_Text = New_Value
    End If
    Call TextBox_SetText(hEdit, m_Text)
    PropertyChanged "Text"
    RaiseEvent Change
    
    If m_InputType = IT_DropDown Then
        Draw
        Exit Property
    End If
    
    If Len(m_Text) = 0 Then
       If Len(m_CueBanner) And bFocus = False Then ShowWindow hEdit, vbHide: bVisible = False
    Else
       If bVisible = False Then ShowWindow hEdit, vbNormalFocus: bVisible = True
    End If

End Property

Public Property Get Font() As StdFont
     Set Font = UserControl.Font
End Property ' Get Font

Public Property Set Font(ByVal new_Font As StdFont)
    Set UserControl.Font = new_Font
    PropertyChanged "Font"
    SetFont
    UserControl_Resize
End Property ' Let Font

Public Property Get BorderStyle() As BORDER_STYLE
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(new_BorderStyle As BORDER_STYLE)
    m_BorderStyle = new_BorderStyle
    PropertyChanged "BorderStyle"
    DoRefresh
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal new_Color As OLE_COLOR)
    UserControl.BackColor = new_Color
    m_BackColor = new_Color
    PropertyChanged "BackColor"
    Draw
End Property

Public Property Get BackColorOnFocus() As OLE_COLOR
   BackColorOnFocus = m_BackColorOnFocus
End Property

Public Property Let BackColorOnFocus(ByVal new_Color As OLE_COLOR)
    m_BackColorOnFocus = new_Color
    PropertyChanged "BackColorOnFocus"
    Draw
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal new_Color As OLE_COLOR)
    m_BorderColor = new_Color
    PropertyChanged "BorderColor"
    Draw
End Property

Public Property Get BorderColorOnFocus() As OLE_COLOR
   BorderColorOnFocus = m_BorderColorOnFocus
End Property

Public Property Let BorderColorOnFocus(ByVal new_Color As OLE_COLOR)
    m_BorderColorOnFocus = new_Color
    PropertyChanged "BorderColorOnFocus"
    Draw
End Property

Public Property Get ParentBackColor() As OLE_COLOR
   ParentBackColor = m_ParentBackColor
End Property

Public Property Let ParentBackColor(ByVal new_Color As OLE_COLOR)
    m_ParentBackColor = new_Color
    PropertyChanged "ParentBackColor"
    Draw
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    SetFont
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    If Len(m_Caption) Then CalculateCaptionSize
    UserControl_Resize
End Property

Public Property Get InputType() As eInputType
    InputType = m_InputType
End Property

Public Property Let InputType(ByVal New_Value As eInputType)
    m_InputType = New_Value
    PropertyChanged "InputType"
    DoRefresh
End Property

Public Function LoadImgLeft(SrcImage As Variant) As Boolean
    If hImgLeft <> 0 Then GdipDisposeImage hImgLeft: hImgLeft = 0
    hImgLeft = LoadPictureEx(SrcImage)
    If hImgLeft Then
        GdipGetImageDimension hImgLeft, ImgLeftRealSize.X, ImgLeftRealSize.Y
        If VarType(SrcImage) = (vbArray Or vbByte) Then
            ArrImgLeft = SrcImage
            LoadImgLeft = True
        End If
    End If
    Call PropertyChanged("ArrImgLeft")
    UserControl_Resize
End Function

Public Function imgLeft() As Byte()
    imgLeft = ArrImgLeft
End Function

Public Sub DeleteImgLeft()
    If hImgLeft <> 0 Then GdipDisposeImage hImgLeft: hImgLeft = 0
    Erase ArrImgLeft
    Call PropertyChanged("ArrImgLeft")
    UserControl_Resize
End Sub

Public Function LoadImgRight(SrcImage As Variant) As Boolean
    If hImgRight <> 0 Then GdipDisposeImage hImgRight: hImgRight = 0
    hImgRight = LoadPictureEx(SrcImage)
    If hImgRight Then
        GdipGetImageDimension hImgRight, ImgRightRealSize.X, ImgRightRealSize.Y
        If VarType(SrcImage) = (vbArray Or vbByte) Then
            ArrImgRight = SrcImage
            LoadImgRight = True
        End If
    End If
    Call PropertyChanged("ArrImgRight")
    UserControl_Resize
End Function

Public Function ImgRight() As Byte()
    ImgRight = ArrImgRight
End Function

Public Sub DeleteImgRight()
    If hImgRight <> 0 Then GdipDisposeImage hImgRight: hImgRight = 0
    Erase ArrImgRight
    Call PropertyChanged("ArrImgRight")
    UserControl_Resize
End Sub

Public Property Get ImgLeftSize() As Long
    ImgLeftSize = m_ImgLeftSize
End Property

Public Property Let ImgLeftSize(ByVal New_Value As Long)
    m_ImgLeftSize = New_Value
    PropertyChanged "ImgLeftSize"
    UserControl_Resize
End Property

Public Property Get ImgRightSize() As Long
    ImgRightSize = m_ImgRightSize
End Property

Public Property Let ImgRightSize(ByVal New_Value As Long)
    m_ImgRightSize = New_Value
    PropertyChanged "ImgRightSize"
    UserControl_Resize
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_Value As Long)
    m_BorderWidth = New_Value
    PropertyChanged "BorderWidth"
    UserControl_Resize
End Property

Public Property Get BorderRadius() As Long
    BorderRadius = m_BorderRadius
End Property

Public Property Let BorderRadius(ByVal New_Value As Long)
    m_BorderRadius = New_Value
    PropertyChanged "BorderRadius"
    Draw
End Property

Public Property Get RightButtonStyle() As eRightButtonStyle
    RightButtonStyle = m_RightButtonStyle
End Property

Public Property Let RightButtonStyle(ByVal New_Value As eRightButtonStyle)
    m_RightButtonStyle = New_Value
    PropertyChanged "RightButtonStyle"
    UserControl_Resize
End Property

Public Property Get ImgLeftFillStyle() As eFillStyle
    ImgLeftFillStyle = m_ImgLeftFillStyle
End Property

Public Property Let ImgLeftFillStyle(ByVal New_Value As eFillStyle)
    m_ImgLeftFillStyle = New_Value
    PropertyChanged "ImgLeftFillStyle"
    Draw
End Property

Public Property Get ImgRightFillStyle() As eFillStyle
    ImgRightFillStyle = m_ImgRightFillStyle
End Property

Public Property Let ImgRightFillStyle(ByVal New_Value As eFillStyle)
    m_ImgRightFillStyle = New_Value
    PropertyChanged "ImgRightFillStyle"
    Draw
End Property

Public Property Get ImgLeftFillColor() As OLE_COLOR
    ImgLeftFillColor = m_ImgLeftFillColor
End Property

Public Property Let ImgLeftFillColor(ByVal New_Value As OLE_COLOR)
    m_ImgLeftFillColor = New_Value
    PropertyChanged "ImgLeftFillColor"
    Draw
End Property

Public Property Get ImgRightFillColor() As OLE_COLOR
    ImgRightFillColor = m_ImgRightFillColor
End Property

Public Property Let ImgRightFillColor(ByVal New_Value As OLE_COLOR)
    m_ImgRightFillColor = New_Value
    PropertyChanged "ImgRightFillColor"
    Draw
End Property

Public Property Get ImgLeftShowMouseEvents() As Boolean
    ImgLeftShowMouseEvents = m_ImgLeftShowMouseEvents
End Property

Public Property Let ImgLeftShowMouseEvents(ByVal New_Value As Boolean)
    m_ImgLeftShowMouseEvents = New_Value
    PropertyChanged "ImgLeftShowMouseEvents"
    Draw
End Property

Public Property Get ImgRightShowMouseEvents() As Boolean
    ImgRightShowMouseEvents = m_ImgRightShowMouseEvents
End Property

Public Property Let ImgRightShowMouseEvents(ByVal New_Value As Boolean)
    m_ImgRightShowMouseEvents = New_Value
    PropertyChanged "ImgRightShowMouseEvents"
    Draw
End Property

Public Property Get OnFocusSelAll() As Boolean
    OnFocusSelAll = m_OnFocusSelAll
End Property

Public Property Let OnFocusSelAll(ByVal New_Value As Boolean)
    m_OnFocusSelAll = New_Value
    PropertyChanged "OnFocusSelAll"
    Draw
End Property

Public Property Get OnFocusBigBorder() As Boolean
    OnFocusBigBorder = m_OnFocusBigBorder
End Property

Public Property Let OnFocusBigBorder(ByVal New_Value As Boolean)
    m_OnFocusBigBorder = New_Value
    PropertyChanged "OnFocusBigBorder"
    If m_OnFocusBigBorder Then
        BB = 3 * nScale
    Else
        BB = 0
    End If
    UserControl_Resize
End Property

Public Property Get CueBanner() As String
    CueBanner = m_CueBanner
End Property

Public Property Let CueBanner(ByVal Value As String)
    m_CueBanner = Value
    If hEdit <> 0 Then SendMessage hEdit, EM_SETCUEBANNER, 0, ByVal StrPtr(m_CueBanner)
    UserControl.PropertyChanged "CueBanner"
    Draw
End Property

Public Property Get hWnd() As Long
    hWnd = hEdit
End Property

Public Property Get hwndUC() As Long
    hwndUC = UserControl.hWnd
End Property
'
'Public Property Get Default() As String
'    Default = Me.Text
'End Property
'
'Public Property Let Default(ByVal Value As String)
'    Me.Text = Value
'End Property

Public Property Get ShortDateFormat() As String
    ShortDateFormat = m_ShortDateFormat
End Property

Public Property Let ShortDateFormat(ByVal Value As String)
    m_ShortDateFormat = Value
    UserControl.PropertyChanged "ShortDateFormat"
    
    If Len(m_ShortDateFormat) Then
        Dim i As Long, sChar As String
        sShortDate = m_ShortDateFormat
        For i = 1 To Len(sShortDate)
            sChar = Mid$(sShortDate, i, 1)
            Select Case UCase(sChar)
                Case "D", "M", "Y", "H", "N", "S"
                Case Else
                    sDateDiv = sChar
                    Exit For
            End Select
        Next
    Else
        If (m_InputType = IT_Date) Then
            sShortDate = fGetLocaleInfo(LOCALE_SSHORTDATE)
            sDateDiv = fGetLocaleInfo(LOCALE_SDATE)
        Else
            sShortDate = fGetLocaleInfo(LOCALE_STIMEFORMAT)
            sDateDiv = fGetLocaleInfo(LOCALE_STIME)
        End If
    End If
    
    If (m_InputType = IT_Date) Or (m_InputType = IT_Time) Then
        m_Text = Format$(m_Text, sShortDate)
    End If
    
End Property

Public Property Get ListPlus() As Object
    Set ListPlus = ObjListPlus
End Property

Public Property Let ListPlus(ByVal ListControl As Object)
    Set ObjListPlus = ListControl
    ObjListPlus.ParentToNotify = UserControl.hWnd
End Property

Public Property Get MinValue() As Variant
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal newValue As Variant)
    On Error Resume Next
    m_MinValue = newValue
    If m_InputType = IT_Numeric Then
        If Val(Me.Text) < m_MinValue Then Me.Text = m_MinValue
    ElseIf (m_InputType = IT_Date Or m_InputType = IT_Time) And Me.TextLenght > 0 Then
         If CDate(Me.Text) < m_MinValue Then Me.Text = m_MinValue
    End If
    PropertyChanged "MinValue"
End Property

Public Property Get MaxValue() As Variant
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal newValue As Variant)
    On Error Resume Next
    m_MaxValue = newValue
    If m_InputType = IT_Numeric Then
        If Val(Me.Text) > m_MaxValue Then Me.Text = m_MaxValue
    ElseIf (m_InputType = IT_Date Or m_InputType = IT_Time) And Me.TextLenght > 0 Then
         If CDate(Me.Text) > m_MaxValue Then Me.Text = m_MaxValue
    End If
    PropertyChanged "MaxValue"
End Property

'-----------------------------------------------------
'UserControl's Events
'-----------------------------------------------------
'*2
Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    nScale = GetWindowsDPI
    Margin = 4 * nScale
    m_MarginLeft = 16
    m_MarginRight = 16
   Call mIOleInPlaceActivate.InitIPAO(m_uIPAO, Me)
   hCurIBEAM = LoadCursor(0, IDC_IBEAM)

End Sub

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Text = Extender.Name
    m_BorderWidth = 1
    m_BorderRadius = 4
    m_ReadOnly = False
    m_ScrollBar = 0
    m_BorderStyle = [Fixed Single]
    'm_PasswordChar = ""
    m_TextAlign = [Left Justify]
    m_Enabled = True
    m_hideSel = False
    m_TextConvert = 0
    m_MaxLength = 0
    m_RightButtonStyle = RS_None
    m_BackColor = vbWindowBackground
    m_BackColorOnFocus = vbWindowBackground
    m_BorderColor = vbActiveBorder
    m_BorderColorOnFocus = vbHighlight
    m_ParentBackColor = Ambient.BackColor
    UserControl.ForeColor = Ambient.ForeColor
    m_ImgLeftFillStyle = FS_Tansparent
    m_ImgRightFillStyle = FS_Tansparent
    m_ImgLeftFillColor = Ambient.BackColor
    m_ImgRightFillColor = Ambient.BackColor
    m_ImgLeftShowMouseEvents = True
    m_ImgRightShowMouseEvents = True
    m_ImgLeftSize = 16
    m_ImgRightSize = 16
    
    DoRefresh
End Sub

Private Sub UserControl_DblClick()
    bMouseDown = True
    If (m_RightButtonStyle = RS_SpinButton) And (HotSpinButton > 0) Then
        Timer1_Timer
        Timer1.Interval = 500
        Timer1.Enabled = True
    End If
    Draw
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If Not ObjListPlus Is Nothing Then
        
        If KeyCode = vbKeyF4 Then
            Dim R As RECT
            GetWindowRect UserControl.hWnd, R
            ME_DropDown R
        End If
        ObjListPlus.SetKeyDown KeyCode
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
            
    If Not ObjListPlus Is Nothing Then
        ObjListPlus.SetKeyPress KeyAscii
    End If
End Sub

Private Sub UserControl_GotFocus()
    bFocus = True
    If m_EditGradient Then
        Draw
        If m_BackColorOnFocus <> m_BackColor And mBkBrush <> 0 Or Len(m_CueBanner) > 0 Then
            Call CreateBackBrush
            RedrawWindow hEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
        End If
    Else
        If m_BackColorOnFocus <> UserControl.BackColor Then UserControl.BackColor = m_BackColorOnFocus
        Draw
    End If

    If m_InputType <> IT_DropDown Then
        If hEdit <> 0 Then
            ShowWindow hEdit, vbNormalFocus
            bVisible = True
            If GetFocus = UserControl.hWnd Then
                PutFocus hEdit
            End If
        End If
    End If

    If m_OnFocusSelAll Then
        Me.SelStart = 0
        Me.SelLength = Me.TextLenght
    End If
End Sub

Private Sub UserControl_LostFocus()
    bFocus = False
    If m_EditGradient Then
        Draw
        If m_BackColorOnFocus <> m_BackColor And mBkBrush <> 0 Then
            Call CreateBackBrush
            RedrawWindow hEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
        End If
    Else
        If m_BackColor <> UserControl.BackColor Then UserControl.BackColor = m_BackColor
        Draw
    End If
    
    If m_InputType = IT_Date Or m_InputType = IT_Time Then
        Me.Text = Me.Text
        If Not IsDate(m_Text) And Len(m_Text) Then Me.Text = mDate
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If bMouseEnter = False Then
        If m_InputType = IT_DropDown Or m_HotBorder = True Then
            bMouseEnter = True
            Draw
        Else
            bMouseEnter = True
        End If
        RaiseEvent MouseEnter
    End If
    
    If m_RightButtonStyle = RS_SpinButton Then
        If X > UserControl.ScaleWidth - Margin * 4 - (m_MarginRight + m_BorderWidth) * nScale Then
            If Y > UserControl.ScaleHeight / 2 Then
                If HotSpinButton <> 1 Then
                    HotSpinButton = 1
                    Draw
                End If
            Else
                If HotSpinButton <> 2 Then
                    HotSpinButton = 2
                    Draw
                End If
            End If
        Else
            If HotSpinButton > 0 Then
                HotSpinButton = 0
                Draw
            End If
        End If
    End If
    
    If X < m_Rect.Left And X > 0 And Y > 0 And Y < m_Rect.Bottom Then
        If HotButton <> 1 Then
            HotButton = 1
            Draw
        End If
    ElseIf (X > m_Rect.Left + m_Rect.Right) And X < UserControl.ScaleWidth And Y > 0 And Y < m_Rect.Bottom And (m_RightButtonStyle <> RS_SpinButton) Then
        If HotButton <> 2 Then
            HotButton = 2
            Draw
        End If
    Else
        If HotButton <> 0 Then
            HotButton = 0
            Draw
        End If
    End If
    
    If X > m_Rect.Left And (X < m_Rect.Left + m_Rect.Right) And Y > 0 And Y < m_Rect.Bottom Then
        If Not m_InputType = IT_DropDown Then
            SetCursor hCurIBEAM
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As RECT
    
    If (X < m_Rect.Left) Then
        RaiseEvent ImgLeftMouseDown(Button, Shift, X, Y)
    ElseIf X > m_Rect.Left + m_Rect.Right Then
        RaiseEvent ImgRightMouseDown(Button, Shift, X, Y)
        If Button = vbLeftButton And m_RightButtonStyle = RS_DropDown Then
            GetWindowRect UserControl.hWnd, R
            RaiseEvent DropDown(R.Left, R.Top, R.Right, R.Bottom)
            If Not ObjListPlus Is Nothing Then ME_DropDown R
        End If
    Else
        If Button = vbLeftButton And m_InputType = IT_DropDown Then
            GetWindowRect UserControl.hWnd, R
            RaiseEvent DropDown(R.Left, R.Top, R.Right, R.Bottom)
            If Not ObjListPlus Is Nothing Then ME_DropDown R
        Else
            Exit Sub
        End If
    End If
    
    If (Button <> vbLeftButton) Then Exit Sub

    bMouseDown = True
    If m_RightButtonStyle = RS_SpinButton And HotSpinButton > 0 Then
        Timer1_Timer
        Timer1.Interval = 500
        Timer1.Enabled = True
    End If
 
    Draw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = False
    
    If m_RightButtonStyle = RS_SpinButton Then
        Timer1.Enabled = False
        Timer1.Interval = 0
        Draw
        Exit Sub
    End If
    
    If X > m_Rect.Left + m_Rect.Right And X < UserControl.ScaleWidth And Y > 0 And Y < m_Rect.Bottom Then
        If m_InputType = IT_PasswordChar And m_RightButtonStyle = RS_ShowPassword Then
            If bShowPassword = False Then
                Call SendMessage(hEdit, EM_SETPASSWORDCHAR, ByVal CLng(0), 0&)
                bShowPassword = True
            Else
                Call SendMessage(hEdit, EM_SETPASSWORDCHAR, ByVal CLng(Asc("•")), 0&)
                bShowPassword = False
            End If
            Draw
        ElseIf m_RightButtonStyle = RS_ClearText Then
            Me.Text = vbNullString
            Draw
        ElseIf m_RightButtonStyle = RS_DropDown Then
            Draw
            
        ElseIf m_RightButtonStyle = RS_Icon Then
            Draw
            RaiseEvent ImgRightMouseUp(Button, Shift, X, Y)
        End If
    
    ElseIf X < m_Rect.Left And X > 0 And Y > 0 And Y < m_Rect.Bottom Then
        Draw
        RaiseEvent ImgLeftMouseUp(Button, Shift, X, Y)
    ElseIf m_InputType = IT_DropDown Then
        Draw
    End If
End Sub

Private Sub UserControl_Resize()
    Dim R As RECT
    Dim BW As Single, ML As Single, MR As Single
    Dim Top As Single
    
    If m_RightButtonStyle = RS_Resizable Then
        MR = m_BorderRadius * nScale / 2
    ElseIf m_RightButtonStyle > RS_Resizable Then
        MR = 16 * nScale
    End If
   
    BW = m_BorderWidth * nScale
    
    If hImgLeft Then ML = m_ImgLeftSize * nScale
    If hImgRight And m_RightButtonStyle <> RS_None Then MR = m_ImgRightSize * nScale
    
    Top = CaptionHeight / 2
    If BB > Top Then Top = BB

    R.Left = BB + BW + Margin
    If ML > 0 Then R.Left = R.Left + ML + nScale + Margin * 2
    If m_InputType = IT_MultiLine Then
        R.Top = Top + BW + Margin
        R.Bottom = UserControl.ScaleHeight - R.Top - Margin - BB - BW '* 2
    Else
        R.Bottom = UserControl.TextHeight(m_Text)
        R.Top = Top + (UserControl.ScaleHeight - Top - BB) / 2 - R.Bottom / 2
    End If
    R.Right = UserControl.ScaleWidth - BB * 2 - BW * 2 - Margin * 2 - nScale
    'R.Right = UserControl.ScaleWidth - R.Left - Margin * 2 - BW '- BB
    If ML > 0 Then R.Right = R.Right - ML - Margin * 2 - nScale * 2
    If MR > 0 Then R.Right = R.Right - MR - Margin * 2

    With m_RectEdit
        .Left = R.Left
        .Top = R.Top
        .Right = R.Left + R.Right
        .Bottom = R.Top + R.Bottom
    End With
    
    SetWindowPos hEdit, 0, R.Left, R.Top, R.Right, R.Bottom, 0
    
    With m_Rect
        .Left = IIF(hImgLeft = 0, 0, R.Left - Margin)
        .Right = IIF(m_RightButtonStyle = RS_None, UserControl.ScaleWidth, R.Right + Margin * 2 + 1)
        .Top = CaptionHeight / 2
        .Bottom = UserControl.ScaleHeight - CaptionHeight / 2
    End With
    Draw
    
    If m_EditGradient Then Call CreateBackBrush
End Sub

Private Sub UserControl_Show()
    Draw
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Catch
    If hEdit <> 0 Then
        Call Subclass_StopAll
         DestroyWindow hEdit
    End If
    If mBkBrush Then DeleteObject mBkBrush
    If hCurIBEAM Then DestroyCursor hCurIBEAM
    If hImgLeft Then GdipDisposeImage hImgLeft
    If hImgRight Then GdipDisposeImage hImgRight
    Call GdiplusShutdown(GdipToken)
    Call mIOleInPlaceActivate.TerminateIPAO(m_uIPAO)
Catch:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        UserControl.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
        UserControl.BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_BackColor = .ReadProperty("BackColor", vbWindowBackground)
        m_BackColorOnFocus = .ReadProperty("BackColorOnFocus", vbWindowBackground)
        m_BorderColor = .ReadProperty("BorderColor", vbActiveBorder)
        m_BorderColorOnFocus = .ReadProperty("BorderColorOnFocus", vbHighlight)
        m_ParentBackColor = .ReadProperty("ParentBackColor", Ambient.BackColor)
        m_Text = VarToStr(.ReadProperty("Text", vbNullString))
        m_InputType = .ReadProperty("InputType", IT_Text)
        m_BorderWidth = .ReadProperty("BorderWidth", 1)
        m_BorderRadius = .ReadProperty("BorderRadius", 4)
        m_ReadOnly = .ReadProperty("Locked", False)
        m_Enabled = .ReadProperty("Enabled", True)
        UserControl.Enabled = m_Enabled
        m_BorderStyle = .ReadProperty("BorderStyle", [Fixed Single])
        m_ScrollBar = .ReadProperty("Scrollbar", 0)
        m_TextAlign = .ReadProperty("Alignment", [Left Justify])
        m_hideSel = .ReadProperty("HideSelection", False)
        m_TextConvert = .ReadProperty("TextConvert", 0)
        m_MaxLength = .ReadProperty("MaxLength", 0)
        m_ImgLeftSize = .ReadProperty("ImgLeftSize", 16)
        m_ImgRightSize = .ReadProperty("ImgRightSize", 16)
        m_RightButtonStyle = .ReadProperty("RightButtonStyle", RS_None)
        m_Caption = VarToStr(.ReadProperty("Caption", vbNullString))
        m_ImgLeftFillStyle = .ReadProperty("ImgLeftFillStyle", FS_Tansparent)
        m_ImgRightFillStyle = .ReadProperty("ImgRightFillStyle", FS_Tansparent)
        m_ImgLeftFillColor = .ReadProperty("ImgLeftFillColor", Ambient.BackColor)
        m_ImgRightFillColor = .ReadProperty("ImgRightFillColor", Ambient.BackColor)
        m_ImgLeftShowMouseEvents = .ReadProperty("ImgLeftShowMouseEvents", True)
        m_ImgRightShowMouseEvents = .ReadProperty("ImgRightShowMouseEvents", True)
        m_OnFocusSelAll = .ReadProperty("OnFocusSelAll", False)
        m_OnFocusBigBorder = .ReadProperty("OnFocusBigBorder", False)
        m_CueBanner = VarToStr(.ReadProperty("CueBanner", vbNullString))
        m_EditGradient = .ReadProperty("EditGradient", False)
        m_ButtonsGradient = .ReadProperty("ButtonsGradient", False)
        m_OnKeyReturnTabulate = .ReadProperty("OnKeyReturnTabulate", False)
        m_HotBorder = .ReadProperty("HotBorder", False)
        m_MinValue = .ReadProperty("MinValue", vbNullString)
        m_MaxValue = .ReadProperty("MaxValue", vbNullString)
        Me.ShortDateFormat = .ReadProperty("ShortDateFormat", vbNullString)
        
        On Error Resume Next
        ArrImgLeft = .ReadProperty("ImgLeft")
        hImgLeft = LoadPictureEx(ArrImgLeft)
        ArrImgRight = .ReadProperty("ImgRight")
        hImgRight = LoadPictureEx(ArrImgRight)
        On Error GoTo 0
    End With
    
    If hImgLeft Then GdipGetImageDimension hImgLeft, ImgLeftRealSize.X, ImgLeftRealSize.Y
    If hImgRight Then GdipGetImageDimension hImgRight, ImgRightRealSize.X, ImgRightRealSize.Y
    
    If Len(m_Caption) Then CalculateCaptionSize  'CaptionHeight = 10 * nScale
    
    If m_OnFocusBigBorder And m_BorderStyle = [Fixed Single] Then
        BB = 3 * nScale
    End If

    sDecimal = fGetLocaleInfo(LOCALE_SDECIMAL)
    sThousand = fGetLocaleInfo(LOCALE_SMONTHOUSANDSEP)
    'sDateDiv = fGetLocaleInfo(LOCALE_SDATE)
    sMoney = fGetLocaleInfo(LOCALE_SCURRENCY)

    DoRefresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "ForeColor", UserControl.ForeColor, Ambient.ForeColor
        .WriteProperty "BackColor", UserControl.BackColor, Ambient.BackColor
        .WriteProperty "BackColor", m_BackColor, vbWindowBackground
        .WriteProperty "BackColorOnFocus", m_BackColorOnFocus, vbWindowBackground
        .WriteProperty "BorderColor", m_BorderColor, vbActiveBorder
        .WriteProperty "BorderColorOnFocus", m_BorderColorOnFocus, vbHighlight
        .WriteProperty "ParentBackColor", m_ParentBackColor, Ambient.BackColor
        .WriteProperty "Text", StrToVar(m_Text), vbNullString
        .WriteProperty "InputType", m_InputType, IT_Text
        .WriteProperty "BorderWidth", m_BorderWidth, 1
        .WriteProperty "BorderRadius", m_BorderRadius, 4
        .WriteProperty "Locked", m_ReadOnly, False
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "BorderStyle", m_BorderStyle, [Fixed Single]
        .WriteProperty "Scrollbar", m_ScrollBar, 0
        .WriteProperty "Alignment", m_TextAlign, [Left Justify]
        .WriteProperty "HideSelection", m_hideSel, False
        .WriteProperty "TextConvert", m_TextConvert, 0
        .WriteProperty "MaxLength", m_MaxLength, 0
        .WriteProperty "ImgLeft", ArrImgLeft, 0
        .WriteProperty "ImgRight", ArrImgRight, 0
        .WriteProperty "ImgLeftSize", m_ImgLeftSize, 16
        .WriteProperty "ImgRightSize", m_ImgRightSize, 16
        .WriteProperty "RightButtonStyle", m_RightButtonStyle
        .WriteProperty "Caption", StrToVar(m_Caption), vbNullString
        .WriteProperty "ImgLeftFillStyle", m_ImgLeftFillStyle, FS_Tansparent
        .WriteProperty "ImgRightFillStyle", m_ImgRightFillStyle, FS_Tansparent
        .WriteProperty "ImgLeftFillColor", m_ImgLeftFillColor, Ambient.BackColor
        .WriteProperty "ImgRightFillColor", m_ImgRightFillColor, Ambient.BackColor
        .WriteProperty "ImgLeftShowMouseEvents", m_ImgLeftShowMouseEvents, True
        .WriteProperty "ImgRightShowMouseEvents", m_ImgRightShowMouseEvents, True
        .WriteProperty "OnFocusSelAll", m_OnFocusSelAll, False
        .WriteProperty "OnFocusBigBorder", m_OnFocusBigBorder, False
        .WriteProperty "CueBanner", StrToVar(m_CueBanner), vbNullString
        .WriteProperty "EditGradient", m_EditGradient, False
        .WriteProperty "ButtonsGradient", m_ButtonsGradient, False
        .WriteProperty "OnKeyReturnTabulate", m_OnKeyReturnTabulate, False
        .WriteProperty "ShortDateFormat", m_ShortDateFormat, vbNullString
        .WriteProperty "HotBorder", m_HotBorder, False
        .WriteProperty "MinValue", m_MinValue, vbNullString
        .WriteProperty "MaxValue", m_MaxValue, vbNullString
    End With
End Sub

Private Sub Timer1_Timer()
    If Timer1.Interval = 1000 Then
        'On Error Resume Next
        Dim Sel As Long
        Sel = Me.SelStart
        If m_MinValue <> vbNullString Then
            If m_InputType = IT_Numeric Then
                If VBA.Val(Me.Text) < m_MinValue Then
                    TextBoxChangeFrozen = True
                    Me.Text = m_MinValue
                    TextBoxChangeFrozen = False
                    Me.SelStart = Sel
                End If
            ElseIf m_InputType = IT_Date Or m_InputType = IT_Time Then
                If Me.TextLenght > 0 Then
                    If CDate(Me.Text) < m_MinValue Then
                        TextBoxChangeFrozen = True
                        Me.Text = m_MinValue
                        TextBoxChangeFrozen = False
                        Me.SelStart = Sel
                    End If
                End If
            End If
        End If
        If m_MaxValue <> vbNullString Then
            If m_InputType = IT_Numeric Then
                If VBA.Val(Me.Text) > m_MaxValue Then
                    TextBoxChangeFrozen = True
                    Me.Text = m_MaxValue
                    TextBoxChangeFrozen = False
                    Me.SelStart = Sel
                End If
            ElseIf m_InputType = IT_Date Or m_InputType = IT_Time Then
                If Me.TextLenght > 0 Then
                    If CDate(Me.Text) > m_MaxValue Then
                        TextBoxChangeFrozen = True
                        Me.Text = m_MaxValue
                        TextBoxChangeFrozen = False
                        Me.SelStart = Sel
                    End If
                End If
            End If
        End If
        Timer1.Interval = 0
        On Error GoTo 0
        Exit Sub
    End If


    If Not GetAsyncKeyState(1) <> 0 Then
        Timer1.Interval = 0
        Timer1.Enabled = False
    End If
    If Timer1.Interval = 500 Then Timer1.Interval = 50
    If (m_InputType = IT_Date) Or (m_InputType = IT_Time) Then
        If HotSpinButton = 1 Then
            Me_KeyDown vbKeyDown
        Else
            Me_KeyDown vbKeyUp
        End If
    Else
        If HotSpinButton = 1 Then
            If m_MinValue <> vbNullString Then
                If Val(Me.Text) > m_MinValue Then
                    Me.Text = Val(Me.Text) - 1
                End If
            Else
                Me.Text = Val(Me.Text) - 1
            End If
        ElseIf HotSpinButton = 2 Then
            If m_MaxValue <> vbNullString Then
                If Val(Me.Text) < m_MaxValue Then
                    Me.Text = Val(Me.Text) + 1
                End If
            Else
                Me.Text = Val(Me.Text) + 1
            End If
        End If
    End If

End Sub

'---------------------------------------------------
'Functions
'---------------------------------------------------
'*3

Public Sub SetMinSize(Optional Width As Long, Optional Height As Long)
    m_MinSize.X = Width
    m_MinSize.Y = Height
End Sub

Public Sub SetMaxSize(Optional Width As Long, Optional Height As Long)
    m_MaxSize.X = Width
    m_MaxSize.Y = Height
End Sub

Public Sub Refresh()
    UserControl_Resize
    RedrawWindow hEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
End Sub

Public Sub ScrollToBottom()
   SendMessage hEdit, WM_VSCROLL, SB_BOTTOM, 0&
End Sub

Public Function HideBalloonTip() As Boolean
    If hEdit <> 0 Then HideBalloonTip = CBool(SendMessageW(hEdit, EM_HIDEBALLOONTIP, 0, ByVal 0&) <> 0)
End Function

Public Function ShowBalloonTip(ByVal Text As String, Optional ByVal Title As String, Optional ByVal Icon As eTTInfo) As Boolean
    If hEdit <> 0 Then
        Dim EDITBT As EDITBALLOONTIP
        With EDITBT
            .cbStruct = LenB(EDITBT)
            .pszText = StrPtr(Text)
            .pszTitle = StrPtr(Title)
            Select Case Icon
                Case TTI_ERROR, TTI_INFO, TTI_NONE, TTI_WARNING
                    .iIcon = Icon
                Case Else
                    Err.Raise 380
            End Select
            If GetFocus() <> hEdit Then PutFocus UserControl.hWnd
            ShowBalloonTip = CBool(SendMessageW(hEdit, EM_SHOWBALLOONTIP, 0, ByVal VarPtr(EDITBT)) <> 0)
        End With
    End If
End Function

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Function LoadPictureEx(SrcImg As Variant, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional ByVal ImagesRadius As Integer) As Long
    Dim hImage1 As Long, hImage2 As Long, hGraphics As Long
    Dim DataArr() As Byte
    Dim lPictureRealWidth As Long, lPictureRealHeight As Long
    Dim X As Long, Y As Long, cx As Long, cy As Long
    Dim sngRatio1 As Single, sngRatio2 As Single
    Dim mPath As Long, Radius As Long, hPen As Long
    
    Select Case VarType(SrcImg)
        Case vbString
            If PathIsURL(SrcImg) Then
            
                If Left$(LCase(SrcImg), 5) = "data:" Then
                    Base64Decode Split(SrcImg, ",")(1), DataArr
                    Call LoadImageFromArray(DataArr, hImage1)
                Else
                    Dim oXMLHTTP As Object
                    Set oXMLHTTP = CreateObject("Microsoft.XMLHTTP")
                    
                    oXMLHTTP.Open "GET", SrcImg, True
                    oXMLHTTP.send
                    While oXMLHTTP.readyState <> 4
                        DoEvents
                    Wend
                    If oXMLHTTP.Status = 200 Then
                        DataArr() = oXMLHTTP.responseBody
                        Call LoadImageFromArray(DataArr, hImage1)
                    End If
                End If
            Else
                Call GdipLoadImageFromFile(ByVal StrPtr(SrcImg), hImage1)
            End If
        Case vbLong
            Dim hBmp As Long
            Dim IIF As ICONINFO
            Dim tBmp As BITMAP
            
            If GetObjectType(SrcImg) = OBJ_BITMAP Then
                If GetObject(SrcImg, Len(tBmp), tBmp) Then
                    
                    If tBmp.bmBitsPixel = 32 And tBmp.bmBits > 0 Then
                        GdipCreateBitmapFromScan0 tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidthBytes, PixelFormat32bppPARGB, tBmp.bmBits, hImage1
                        GdipImageRotateFlip hImage1, RotateNoneFlipY
                    Else
                        Call GdipCreateBitmapFromHBITMAP(SrcImg, 0, hImage1)
                    End If
                End If
            Else
        
                If TypeName(SrcImg) = "Long" Then 'StdPicture.handle
                    If GetIconInfo(SrcImg, IIF) Then
                    
                        hBmp = CopyImage(IIF.hbmColor, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION Or &H2)
                        Call GetObject(hBmp, Len(tBmp), tBmp)
                        GdipCreateBitmapFromScan0 tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidthBytes, PixelFormat32bppPARGB, tBmp.bmBits, hImage1
                        GdipImageRotateFlip hImage1, RotateNoneFlipY
                        DeleteObject hBmp
                        DeleteObject IIF.hbmColor
                        DeleteObject IIF.hbmMask
                    Else
                        On Error Resume Next
                        If IsBadCodePtr(SrcImg) = 0 Then
                            GdipCloneImage SrcImg, hImage1 ' OBJECT GDI PLUS BITMAP
                        End If
                        On Error GoTo 0
                    End If
                Else 'StdPicture
                    GdipCreateBitmapFromHICON SrcImg, hImage1
                End If
            End If
            
        Case vbDataObject
            Call GdipLoadImageFromStream(SrcImg, hImage1)
            
        Case (vbArray Or vbByte)
            DataArr() = SrcImg
            Call LoadImageFromArray(DataArr, hImage1)
    End Select

    If hImage1 <> 0 Then
        GdipGetImageWidth hImage1, lPictureRealWidth
        GdipGetImageHeight hImage1, lPictureRealHeight
        If Width = 0 Then Width = lPictureRealWidth
        If Height = 0 Then Height = lPictureRealHeight
        
        sngRatio1 = Width / lPictureRealWidth
        sngRatio2 = Height / lPictureRealHeight
        If sngRatio1 > sngRatio2 Then sngRatio1 = sngRatio2
        cx = lPictureRealWidth * sngRatio1: cy = lPictureRealHeight * sngRatio1
        X = (Width - cx) \ 2: Y = (Height - cy) \ 2

        GdipCreateBitmapFromScan0 Width, Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage2
        GdipGetImageGraphicsContext hImage2, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        
        If ImagesRadius > 0 Then
            Radius = ImagesRadius * nScale
            If GdipCreatePath(&H0, mPath) = 0 Then
                X = X + nScale: Y = Y + nScale
                cx = cx - nScale * 2: cy = cy - nScale * 2
                GdipAddPathArcI mPath, X, Y, Radius, Radius, 180, 90
                GdipAddPathArcI mPath, X + cx - Radius, Y, Radius, Radius, 270, 90
                GdipAddPathArcI mPath, X + cx - Radius, Y + cy - Radius, Radius, Radius, 0, 90
                GdipAddPathArcI mPath, X, Y + cy - Radius, Radius, Radius, 90, 90
                GdipClosePathFigure mPath
                GdipSetClipPath hGraphics, mPath, CombineModeIntersect
            End If
        End If
        
        GdipDrawImageRectI hGraphics, hImage1, X, Y, cx, cy
        
        If mPath Then
            GdipCreatePen1 RGBtoARGB(vbButtonFace, 100), 1 * nScale, &H2, hPen
            GdipResetClip hGraphics
            GdipDrawPath hGraphics, hPen, mPath
            GdipDeletePen hPen
            GdipDeletePath mPath
        End If
        
        GdipDeleteGraphics hGraphics
        GdipDisposeImage hImage1
        LoadPictureEx = hImage2
    End If
End Function

Private Function Base64Decode(ByVal sIn As String, ByRef bvOut() As Byte) As Boolean 'By Cocus
    Dim lLenOut As Long
    Const CRYPT_STRING_BASE64 = &H1
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, 0, VarPtr(lLenOut), 0, 0)
    If lLenOut = 0 Then Exit Function
    ReDim bvOut(lLenOut - 1)
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, VarPtr(bvOut(0)), VarPtr(lLenOut), 0, 0)
    Base64Decode = True
End Function

Private Function LoadImageFromArray(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromArray_Error
    Dim IStream     As IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromArray = True
        End If
    End If
    Set IStream = Nothing
    
LoadImageFromArray_Error:
End Function

Private Sub Me_KeyDown(KeyCode As Integer)
    On Error Resume Next
    Dim Val As Long
    Dim Sel As Long
    Dim TheDate   As Date
    Dim Part1() As String
    Dim Part2() As String
    Dim ConvDate As Date
    Dim d As Variant, M As Variant, Y As Variant
    Dim H As Variant, N As Variant, s As Variant
    Dim i As Long
  
    If KeyCode = vbKeyUp Then
       Val = 1
       KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
       Val = -1
       KeyCode = 0
    ElseIf (KeyCode = vbKeyLeft) Then
        KeyCode = 0
        Sel = Me.SelStart - 2
        Me.Text = Me.Text
        If Not IsDate(m_Text) Then Me.Text = mDate
        Me.SelStart = Sel
        AutoSelDatePart
        Exit Sub
    ElseIf KeyCode = vbKeyRight Then
        KeyCode = 0
        Sel = Me.SelStart + Me.SelLength + 1
        Me.Text = Me.Text
        If Not IsDate(m_Text) Then Me.Text = mDate
        Me.SelStart = Sel
        AutoSelDatePart
        Exit Sub
    Else
       Exit Sub
    End If
    
    If m_InputType = IT_Time Then
        Dim mText As String
        mText = Format(Me.Text, "hh:nn:ss")
    Else
        mText = Me.Text
    End If
    
    Part1 = Split(mText, sDateDiv)
    Part2 = Split(sShortDate, sDateDiv)
        
    For i = 0 To UBound(Part2)
        Dim sChar As String
        sChar = Left$(Part2(i), 1)
        
        If UCase(sChar) = "D" Then Part2(i) = sChar: d = Part1(i)
        If sChar = "M" Then Part2(i) = sChar: M = Part1(i)
        If UCase(sChar) = "Y" Then Part2(i) = "YYYY": Y = Part1(i)
        If UCase(sChar) = "H" Then Part2(i) = sChar: H = Part1(i)
        If sChar = "m" Then Part2(i) = "n": N = Part1(i)
        If UCase(sChar) = "N" Then Part2(i) = "n": N = Part1(i)
        If UCase(sChar) = "S" Then Part2(i) = sChar: s = Part1(i)
    Next
    
    If m_InputType = IT_Date Then
        If Not IsNumeric(M) Then
            ConvDate = CDate(Me.Text)
        Else
            ConvDate = VBA.DateSerial(Y, M, d)
        End If
    Else
        ConvDate = VBA.TimeSerial(H, N, s)
    End If
    
    Sel = Me.SelStart

    Select Case Sel
       Case Is < Len(Part1(0)) + 1
          TheDate = DateAdd(Part2(0), Val, ConvDate)
       Case Is < Len(Part1(0) & Len(Part1(1))) + 2
          TheDate = DateAdd(Part2(1), Val, ConvDate)
       Case Else
            If UBound(Part2) < 2 Then
                TheDate = ConvDate
            Else
                TheDate = DateAdd(Part2(2), Val, ConvDate)
            End If
    End Select
    
    If m_MinValue <> vbNullString Then
        If TheDate < m_MinValue Then TheDate = m_MinValue
    End If
    
    If m_MaxValue <> vbNullString Then
        If TheDate > m_MaxValue Then TheDate = m_MaxValue
    End If
    
    Me.Text = TheDate

    If IsDate(m_Text) Then mDate = m_Text
   
    Me.SelStart = Sel
    AutoSelDatePart
End Sub

Private Sub AutoSelDatePart()
    Dim Sel As Long
    Dim Part1() As String
    
    If Me.TextLenght < Len(sShortDate) Then Exit Sub
    
    Part1 = Split(Me.Text, sDateDiv)
    
    Sel = Me.SelStart

    Select Case Sel
        Case Is < Len(Part1(0)) + 1
            Me.SelStart = 0
            Me.SelLength = Len(Part1(0))
        Case Is < Len(Part1(0) & Part1(1)) + 2
            Me.SelStart = Len(Part1(0)) + 1
            Me.SelLength = Len(Part1(1))
        Case Else
            If UBound(Part1) < 2 Then
                Me.SelStart = 0
                Me.SelLength = Len(Part1(0))
            Else
                Me.SelStart = Me.TextLenght - Len(Part1(2))
                Me.SelLength = Len(Part1(2))
            End If
    End Select
End Sub
'*1

Private Sub Me_KeyPress(KeyAscii As Integer) ', oCtrl As TextBox)
    Dim lCurPos As Long
    Dim lLineLength As Long
    Dim i As Integer
    
    Dim DecimalDotB As Boolean

    If m_OnKeyReturnTabulate Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Call keybd_event(vbKeyTab, 0, KEYEVENTF_KEYDOWN, 0)
            Call keybd_event(vbKeyTab, 0, KEYEVENTF_KEYUP, 0)
        End If
    End If
    
    Select Case m_InputType
        Case IT_Date, IT_Time
            Dim DivCount1 As Long
            Dim DivCount2 As Long
            
        
            If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And KeyAscii <> Asc(sDateDiv) Or KeyAscii = vbKeySpace Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            DivCount1 = UBound(Split(Me.Text, sDateDiv))
            DivCount2 = UBound(Split(sShortDate, sDateDiv))
            
            If KeyAscii = Asc(sDateDiv) Then
                If Me.SelStart < Me.TextLenght Then
                    KeyAscii = 0
                    Me.SelStart = Me.SelStart + Me.SelLength + 1
                    AutoSelDatePart
                    Exit Sub
                Else
                    If DivCount1 >= DivCount2 Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
            
            If IsNumeric(Chr$(KeyAscii)) Then

                Dim LCP As Long 'len current part
                Dim PL As Long
                Dim IL  As Long
                Dim sParts() As String

                

                sParts = Split(Left(Me.Text, Me.SelStart) & Chr(KeyAscii), sDateDiv)
                
                LCP = UBound(sParts)
                
                PL = Len(Split(sShortDate, sDateDiv)(LCP))
                
                IL = Len(sParts(LCP))
                
                
                If (PL = 1 And IL > 2) Or (PL >= 2 And IL > PL) Then

                    If Me.SelStart = Me.TextLenght Then
                        If DivCount1 < DivCount2 Then
                            Me.Text = Me.Text & sDateDiv
                            Me.SelStart = Me.TextLenght
                        Else
                            KeyAscii = 0
                        End If
                    Else
                        KeyAscii = 0
                    End If
                End If
            End If
            
        Case IT_Desimal
            If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 44 And KeyAscii <> 46 And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
                'Beep
                Exit Sub
            Else
                If Chr(KeyAscii) = sThousand Then KeyAscii = Asc(sDecimal)
                ' Determine textbox length
                lLineLength = SendMessage(Me.hWnd, EM_LINELENGTH, lCurPos, 0)
        
                ' Determine existance of ","
                For i = 1 To lLineLength
                    If Mid$(Me.Text, i, 1) = sDecimal Then
                        DecimalDotB = True
                        Exit For
                    End If
                Next i
        
                ' Make sure Decimal separator is only typed once
                If sDecimal = Chr$(44) Then
                  If KeyAscii = 44 And DecimalDotB = False Then
                      DecimalDotB = True
                  ElseIf KeyAscii = 44 And DecimalDotB = True Then
                      KeyAscii = 0
                      'Beep
                      Exit Sub
                  End If
                ElseIf sDecimal = Chr$(46) Then
                  If KeyAscii = 46 And DecimalDotB = False Then
                      DecimalDotB = True
                  ElseIf KeyAscii = 46 And DecimalDotB = True Then
                      KeyAscii = 0
                      'Beep
                      Exit Sub
                  End If
                End If
            End If
        
        Case IT_LettersOnly
            If Not (KeyAscii > 64 And KeyAscii < 91) And Not (KeyAscii > 96 And KeyAscii < 123) And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Function CreateTextBox(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional Scroll As SCROLL_BAR = 0, _
                                Optional Align As TEXT_ALIGN = [Left Justify], _
                                Optional Convert As TEXT_CONVERT = 0, _
                                Optional Multiline As Boolean = False, _
                                Optional HideSel As Boolean = False, _
                                Optional NumberOnly As Boolean = False, _
                                Optional Password As Boolean = False, _
                                Optional ReadOnly As Boolean = False, _
                                Optional lAdditionalStyles As Long = 0) As Long
    
    Dim lExStyle As Long, lStyle As Long
    
    If hParent = 0 Then Exit Function

    lStyle = WS_CHILD Or WS_TABSTOP Or WS_VISIBLE Or WS_CLIPSIBLINGS 'Or ES_WANTRETURN  'Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL

    If Multiline Then
        Select Case Scroll
'        Case [None]
'            lStyle = lStyle
        Case [Horizontal]
            lStyle = lStyle Or WS_HSCROLL
        Case [Vertical]
            lStyle = lStyle Or WS_VSCROLL
        Case [both]
            lStyle = lStyle Or WS_VSCROLL Or WS_HSCROLL
        End Select
        lStyle = lStyle Or ES_MULTILINE Or ES_AUTOVSCROLL 'Or ES_WANTRETURN
    Else
        lStyle = lStyle Or ES_AUTOHSCROLL
    End If
    
    If HideSel Then lStyle = lStyle Or ES_NOHIDESEL
    
    If ReadOnly Then lStyle = lStyle Or ES_READONLY
    
    If NumberOnly Then lStyle = lStyle Or ES_NUMBER
    
    If Password Then lStyle = lStyle Or ES_PASSWORD

    Select Case Align
        Case [Left Justify]
            lStyle = lStyle Or ES_LEFT
        Case [Right Justify]
            lStyle = lStyle Or ES_RIGHT
        Case Center
            lStyle = lStyle Or ES_CENTER
    End Select

    Select Case Convert
'        Case [None]
'            lStyle = lStyle
        Case [LowerCase]
            lStyle = lStyle Or ES_LOWERCASE
        Case [UpperCase]
            lStyle = lStyle Or ES_UPPERCASE
    End Select
    
    If lAdditionalStyles > 0 Then lStyle = lStyle Or lAdditionalStyles

    Dim hTemp As Long
    hTemp = CreateWindowExW(lExStyle, StrPtr("edit"), StrPtr(strCaption), lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    'If hTemp <= 0 Then Exit Function
    CreateTextBox = hTemp
    
    If Password Then
        UserControl.Font.Name = "Tahoma"
        SendMessage hTemp, EM_SETPASSWORDCHAR, ByVal CLng(Asc("•")), 0&
    End If
End Function

Private Sub DoRefresh()
    If hEdit <> 0 Then m_Text = TextBox_GetText(hEdit)
    
    If hEdit <> 0 Then
        Call Subclass_StopAll
        DestroyWindow hEdit
    End If

    hEdit = CreateTextBox(UserControl.hWnd, m_Text, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_ScrollBar, m_TextAlign, m_TextConvert, m_InputType = IT_MultiLine, m_hideSel, m_InputType = IT_Numeric, m_InputType = IT_PasswordChar, m_ReadOnly)
    If Len(m_CueBanner) Then Me.CueBanner = m_CueBanner
    InitializeSubClassing
    TextBox_SetText hEdit, m_Text
    SetFont
    UserControl_Resize
    
    If Len(m_Text) = 0 And Len(m_CueBanner) And bFocus = False Or m_InputType = IT_DropDown Then
        ShowWindow hEdit, vbHide
        bVisible = False
    End If
    
    If m_Enabled = False Then EnableWindow hEdit, m_Enabled
    If m_MaxLength > 0 Then SendMessage hEdit, EM_LIMITTEXT, m_MaxLength, 0&
End Sub

Private Sub CreateBackBrush()
    Dim W As Long, H As Long
    Dim DC As Long, hdc As Long
    Dim hBmp As Long, OldBmp As Long

    If mBkBrush Then DeleteObject mBkBrush
    W = m_RectEdit.Right - m_RectEdit.Left
    H = m_RectEdit.Bottom - m_RectEdit.Top
    DC = GetDC(0)
    hdc = CreateCompatibleDC(DC)
    hBmp = CreateCompatibleBitmap(DC, W, H)
    ReleaseDC 0, DC
    OldBmp = SelectObject(hdc, hBmp)
    BitBlt hdc, 0, 0, W, H, UserControl.hdc, m_RectEdit.Left, m_RectEdit.Top, vbSrcCopy
    mBkBrush = CreatePatternBrush(hBmp)
    DeleteObject SelectObject(hdc, OldBmp)
    DeleteDC hdc
End Sub

Private Function TextBox_SetText(hWnd As Long, sText As String) As Long
    TextBox_SetText = SendMessageW(hWnd, WM_SETTEXT, 0&, ByVal StrPtr(sText))
End Function

Private Function TextBox_GetText(hWnd As Long) As String
    Dim sText As String
    Dim lLength As Long
    lLength = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
    sText = Space$(lLength + 1)
    Call SendMessageW(hWnd, WM_GETTEXT, lLength + 1, ByVal StrPtr(sText))
    TextBox_GetText = Left$(sText, lLength)
End Function

Private Sub SetFont()
    Dim hFont As Long
    hFont = SendMessage(UserControl.hWnd, WM_GETFONT, 0, 0)
    SendMessage hEdit, WM_SETFONT, hFont, 0
End Sub

Private Sub ME_DropDown(R As RECT)
    If ObjListPlus.IsListVisible Then
        ObjListPlus.HideList
        Exit Sub
    End If
    ObjListPlus.ShowList R.Left, R.Bottom
End Sub

Private Sub CalculateCaptionSize()
    Dim lSize As Long
    Dim bBold As Boolean
    Dim R As RECT
    lSize = UserControl.Font.Size
    bBold = UserControl.Font.Bold

    UserControl.Font.Size = 7
    UserControl.Font.Bold = True
    DrawTextW hdc, StrPtr(m_Caption), -1, R, DT_CALCRECT Or DT_SINGLELINE
    CaptionWidth = R.Right
    CaptionHeight = R.Bottom
    UserControl.Font.Size = lSize
    UserControl.Font.Bold = bBold
End Sub

Private Function WordHi(lngValue As Long) As Long
    If (lngValue And &H80000000) = &H80000000 Then
        WordHi = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        WordHi = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function

Private Function WordLo(lngValue As Long) As Long
    WordLo = (lngValue And &HFFFF&)
End Function

Private Function pvShiftState() As Integer
  Dim lS As Integer
    If (GetAsyncKeyState(vbKeyShift) < 0) Then
        lS = lS Or vbShiftMask
    End If
    If (GetAsyncKeyState(vbKeyMenu) < 0) Then
        lS = lS Or vbAltMask
    End If
    If (GetAsyncKeyState(vbKeyControl) < 0) Then
        lS = lS Or vbCtrlMask
    End If
    pvShiftState = lS
End Function

Private Function CreateRoundPath(ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, ByVal Radius As Single) As Long
    Dim hPath As Long
    If GdipCreatePath(&H0, hPath) = 0& Then
    
        If Radius > Width / 2 Then Radius = Width / 2
        If Radius > Height / 2 Then Radius = Height / 2
    
        If Radius = 0 Then
            GdipAddPathRectangle hPath, Left, Top, Width, Height
        Else
            Radius = Radius * 2
            GdipAddPathArc hPath, Left, Top, Radius, Radius, 180, 90
            GdipAddPathArc hPath, Left + Width - Radius, Top, Radius, Radius, 270, 90
            GdipAddPathArc hPath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
            GdipAddPathArc hPath, Left, Top + Height - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure hPath
        End If
        CreateRoundPath = hPath
    End If
    
End Function

Public Function RGBtoARGB(ByVal RGBColor As Long, Optional ByVal Opacity As Long = 100) As Long
    'By LaVople
    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
End Function

'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    
    If (clrFirst And &H80000000) Then clrFirst = GetSysColor(clrFirst And &HFF&)
    CopyMemory clrFore(0), clrFirst, 4&
    
    If (clrSecond And &H80000000) Then clrSecond = GetSysColor(clrSecond And &HFF&)
    CopyMemory clrBack(0), clrSecond, 4&

    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Sub DrawCross(hGraphics As Long, Color As Long, Left As Single, Top As Single, SizeBox As Single)
    Dim hPen As Long
    GdipCreatePen1 Color, SizeBox / 5, UnitPixel, hPen
    GdipDrawLine hGraphics, hPen, Left, Top, Left + SizeBox, Top + SizeBox
    GdipDrawLine hGraphics, hPen, Left, Top + SizeBox, Left + SizeBox, Top
    GdipDeletePen hPen
End Sub

Private Sub DrawEye(hGraphics As Long, Color As Long, Left As Single, Top As Single, SizeBox As Single, StyleHideText As Boolean)
    Dim hPath As Long, hBrush As Long, hPen As Long

    GdipCreatePath 0&, hPath
    GdipAddPathBezier hPath, Left, Top + SizeBox / 2, Left + SizeBox / 4, Top, Left + SizeBox * 0.75, Top, Left + SizeBox, Top + SizeBox / 2
    GdipAddPathBezier hPath, Left + SizeBox, Top + SizeBox / 2, Left + SizeBox * 0.75, Top + SizeBox, Left + SizeBox / 4, Top + SizeBox, Left, Top + SizeBox / 2
    GdipCreatePen1 Color, SizeBox / 6, &H2, hPen
    GdipDrawPath hGraphics, hPen, hPath
    If StyleHideText Then GdipDrawLine hGraphics, hPen, Left, Top + SizeBox, Left + SizeBox, Top

    GdipDeletePen hPen
    GdipResetPath hPath
    
    GdipAddPathEllipse hPath, Left + SizeBox / 3, Top + SizeBox / 3, SizeBox / 3, SizeBox / 3
    GdipCreateSolidFill Color, hBrush
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
End Sub

Private Sub DrawArrowUp(hGraphics As Long, Color As Long, Left As Single, Top As Single, SizeBox As Single)
    Dim SW As Single, SH As Single
    Dim L As Single, T As Single
    Dim hPath As Long, hBrush As Long
    
    Dim pt(6) As POINTF
    SW = SizeBox
    SH = SizeBox * 0.6
    
    L = Left
    T = Top + SizeBox / 2 - SH / 2

    pt(0).X = L + 0 * SW: pt(0).Y = T + 1 * SH
    pt(1).X = L + 0.5 * SW: pt(1).Y = T + 0.4 * SH
    pt(2).X = L + 1 * SW: pt(2).Y = T + 1 * SH
    pt(3).X = L + 1 * SW: pt(3).Y = T + 0.6 * SH
    pt(4).X = L + 0.5 * SW: pt(4).Y = T + 0 * SH
    pt(5).X = L + 0 * SW: pt(5).Y = T + 0.6 * SH
    pt(6).X = L + 0 * SW: pt(6).Y = T + 1 * SH
    
    GdipCreatePath 0&, hPath
    GdipAddPathPolygon hPath, pt(0), 7
    GdipCreateSolidFill Color, hBrush
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
End Sub

Private Sub DrawArrowDown(hGraphics As Long, Color As Long, Left As Single, Top As Single, SizeBox As Single)
    Dim SW As Single, SH As Single
    Dim L As Single, T As Single
    Dim hPath As Long, hBrush As Long
    
    Dim pt(6) As POINTF
    SW = SizeBox
    SH = SizeBox * 0.6

    L = Left
    T = Top + SizeBox / 2 - SH / 2

    pt(0).X = L + 0 * SW: pt(0).Y = T + 0
    pt(1).X = L + 0.5 * SW: pt(1).Y = T + 0.6 * SH
    pt(2).X = L + 1 * SW: pt(2).Y = T + 0 * SH
    pt(3).X = L + 1 * SW: pt(3).Y = T + 0.4 * SH
    pt(4).X = L + 0.5 * SW: pt(4).Y = T + 1 * SH
    pt(5).X = L + 0 * SW: pt(5).Y = T + 0.4 * SH
    pt(6).X = L + 0 * SW: pt(6).Y = T + 0 * SH
    
    GdipCreatePath 0&, hPath
    GdipAddPathPolygon hPath, pt(0), 7
    GdipCreateSolidFill Color, hBrush
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
End Sub

Private Function StrToVar(ByVal Text As String) As Variant
    If Text = vbNullString Then
        StrToVar = Empty
    Else
        Dim b() As Byte
        b() = Text
        StrToVar = b()
    End If
End Function

Private Function VarToStr(ByVal Bytes As Variant) As String
    If IsEmpty(Bytes) Then
        VarToStr = vbNullString
    Else
        Dim b() As Byte
        b() = Bytes
        VarToStr = b()
    End If
End Function

Private Function CUIntToInt(ByVal Value As Long) As Integer
    Const OFFSET_2 As Long = 65536
    Const MAXINT_2 As Integer = 32767
    If Value < 0 Or Value >= OFFSET_2 Then Err.Raise 6
    If Value <= MAXINT_2 Then
        CUIntToInt = Value
    Else
        CUIntToInt = Value - OFFSET_2
    End If
End Function

Private Function CIntToUInt(ByVal Value As Integer) As Long
    Const OFFSET_2 As Long = 65536
    If Value < 0 Then
        CIntToUInt = Value + OFFSET_2
    Else
        CIntToUInt = Value
    End If
End Function

Private Function fGetLocaleInfo(Valor As RegionalConstant) As String
   Dim Simbolo As String
   Dim r1 As Long
   Dim r2 As Long
   Dim p As Integer
   Dim Locale As Long
     
   Locale = GetUserDefaultLCID()
   r1 = GetLocaleInfo(Locale, Valor, vbNullString, 0)
   Simbolo = String$(r1, 0)
   r2 = GetLocaleInfo(Locale, Valor, Simbolo, r1)
   p = InStr(Simbolo, Chr$(0))
     
   If p > 0 Then
      fGetLocaleInfo = Left$(Simbolo, p - 1)
   End If
     
End Function

'*1
Public Sub Draw()
    Dim hGraphics As Long, hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim BorderColor As Long, BackColor As Long, IconBkColor As Long
    Dim BW As Single, ML As Single, MR As Single, BR As Single ', BB As Single
    Dim Color2 As Long
    Dim ArrowSize As Single, ArrowColor As Long
    Dim Left As Single, Top As Single
    Dim RL As RECTL
    Dim ImgScale As Single
    
    
    If bFocus Then
        BorderColor = RGBtoARGB(m_BorderColorOnFocus)
    Else
        If bMouseEnter And (m_HotBorder = True) Then
            BorderColor = RGBtoARGB(ShiftColor(m_BorderColorOnFocus, m_BorderColor, 100))
        Else
            BorderColor = RGBtoARGB(m_BorderColor)
        End If
    End If
    
    'BorderColor = RGBtoARGB(IIF(bFocus, m_BorderColorOnFocus, m_BorderColor), 100) ' RGBtoARGB(&HF77300, 100)
    
    
    If m_InputType = IT_DropDown Then
        If bMouseDown Then
            BackColor = ShiftColor(m_BackColorOnFocus, vbBlack, 240)
        Else
            If bMouseEnter Then
                If bFocus Then
                    BackColor = ShiftColor(m_BackColorOnFocus, vbWhite, 180)
                Else
                    BackColor = ShiftColor(m_BackColor, vbWhite, 180)
                End If
            Else
                BackColor = IIF(bFocus, m_BackColorOnFocus, m_BackColor)
            End If
        End If
    Else
        BackColor = IIF(bFocus, m_BackColorOnFocus, m_BackColor)
    End If

    BW = m_BorderWidth * nScale
    If hImgLeft Then ML = m_ImgLeftSize * nScale
    If hImgRight Then MR = m_ImgRightSize * nScale
    BR = m_BorderRadius * nScale
    If BR > (UserControl.ScaleHeight / 2 - CaptionHeight / 4) Then BR = (UserControl.ScaleHeight / 2 - CaptionHeight / 4)

    If GdipCreateFromHDC(UserControl.hdc, hGraphics) <> 0 Then Exit Sub

    GdipCreateSolidFill RGBtoARGB(m_ParentBackColor, 100), hBrush
    GdipFillRectangleI hGraphics, hBrush, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    GdipDeleteBrush hBrush
    
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

    Top = CLng(CaptionHeight / 2 + BW / 2)
    If BB > Top Then Top = BB
    hPath = CreateRoundPath(CLng(BB + BW / 2), Top, UserControl.ScaleWidth - BW - BB * 2 - 1, UserControl.ScaleHeight - Top - BW - BB - 1, BR)

    If bFocus And BB > 0 Then
        GdipCreatePen1 RGBtoARGB(m_BorderColorOnFocus, 30), BB * 2, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    End If

    If m_EditGradient Then
        RL.Width = UserControl.ScaleWidth
        RL.Height = UserControl.ScaleHeight
        
        Color2 = RGBtoARGB(ShiftColor(BackColor, vbBlack, 220), 100)
        BackColor = RGBtoARGB(BackColor, 100)
        
        GdipCreateLineBrushFromRectI RL, Color2, BackColor, LinearGradientModeVertical, WrapModeTileFlipX, hBrush
    Else
        BackColor = RGBtoARGB(BackColor, 100)
        GdipCreateSolidFill BackColor, hBrush
    End If
    
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush

    '-------------
    If (HotButton = 1) Or (m_ImgLeftFillStyle = FS_Solid) Then
        If bMouseDown And (HotButton = 1) Then
            IconBkColor = ShiftColor(m_ImgLeftFillColor, vbBlack, 240)
        Else
            If HotButton = 1 Then
                IconBkColor = ShiftColor(m_ImgLeftFillColor, vbWhite, 180)
            Else
                IconBkColor = m_ImgLeftFillColor
            End If
        End If
        
        If m_ImgLeftFillStyle = FS_Tansparent Then
            If m_ImgLeftShowMouseEvents Then
                GdipCreateSolidFill RGBtoARGB(IconBkColor, 100), hBrush
                Left = BB + BW + Margin / 2 '* 2
                Top = m_Rect.Top + m_Rect.Bottom / 2 - (ImgLeftSize * nScale) / 2
                GdipFillEllipse hGraphics, hBrush, Left, Top - Margin, ImgLeftSize * nScale + Margin * 2, ImgLeftSize * nScale + Margin * 2
                GdipDeleteBrush hBrush
            End If
        Else
            If m_Rect.Left > 0 Then
                If m_ButtonsGradient Then
                    RL.Width = UserControl.ScaleWidth
                    RL.Height = UserControl.ScaleHeight
                    
                    Color2 = RGBtoARGB(ShiftColor(IconBkColor, vbWhite, 200), 100)
                    GdipCreateLineBrushFromRectI RL, Color2, RGBtoARGB(IconBkColor, 100), _
                                        LinearGradientModeVertical, WrapModeTileFlipX, hBrush
                Else
                    GdipCreateSolidFill RGBtoARGB(IconBkColor, 100), hBrush
                End If
                GdipSetClipRectI hGraphics, m_Rect.Left, m_Rect.Top, UserControl.ScaleWidth, m_Rect.Bottom, CombineModeExclude
                GdipFillPath hGraphics, hBrush, hPath
                GdipResetClip hGraphics
                GdipDeleteBrush hBrush
                If hImgLeft Then
                    GdipCreatePen1 RGBtoARGB(ShiftColor(IconBkColor, vbBlack, 220), 100), nScale, UnitPixel, hPen
                    GdipDrawLine hGraphics, hPen, m_Rect.Left, Top, m_Rect.Left, UserControl.ScaleHeight - BW - BB - 1
                    GdipDeletePen hPen
                End If
            End If
        End If
    End If
    
    If (HotButton = 2) Or (m_ImgRightFillStyle = FS_Solid) Then
        If bMouseDown And (HotButton = 2) Then
            IconBkColor = ShiftColor(m_ImgRightFillColor, vbBlack, 240)
        Else
            If HotButton = 2 Then
                IconBkColor = ShiftColor(m_ImgRightFillColor, vbWhite, 180)
            Else
                IconBkColor = m_ImgRightFillColor
            End If
        End If
        

        If m_RightButtonStyle > RS_Resizable Then 'm_Rect.Right - UserControl.ScaleWidth
            If m_ImgRightShowMouseEvents Then
                If m_ButtonsGradient Then
                    RL.Width = UserControl.ScaleWidth
                    RL.Height = UserControl.ScaleHeight
                    
                    Color2 = RGBtoARGB(ShiftColor(IconBkColor, vbWhite, 200), 100)
                    GdipCreateLineBrushFromRectI RL, Color2, RGBtoARGB(IconBkColor, 100), _
                                        LinearGradientModeVertical, WrapModeTileFlipX, hBrush
                Else
                    GdipCreateSolidFill RGBtoARGB(IconBkColor, 100), hBrush
                End If
                GdipSetClipRectI hGraphics, 0, m_Rect.Top, m_Rect.Left + m_Rect.Right, m_Rect.Bottom, CombineModeExclude
                GdipFillPath hGraphics, hBrush, hPath
                GdipResetClip hGraphics
                GdipDeleteBrush hBrush
                GdipCreatePen1 RGBtoARGB(ShiftColor(IconBkColor, vbBlack, 220), 100), nScale, UnitPixel, hPen
                GdipDrawLine hGraphics, hPen, m_Rect.Left + m_Rect.Right, Top, m_Rect.Left + m_Rect.Right, UserControl.ScaleHeight - BW - BB - 1
                GdipDeletePen hPen
            End If
        End If
    End If

    ArrowSize = UserControl.ScaleHeight / 3
    If ArrowSize > 12 * nScale Then ArrowSize = 12 * nScale
    Left = UserControl.ScaleWidth - 8 * nScale - ArrowSize / 2 - Margin - BW - BB
    ArrowColor = RGBtoARGB(ShiftColor(UserControl.ForeColor, UserControl.BackColor, 100))
    
    If m_RightButtonStyle = RS_SpinButton Then

        If HotSpinButton > 0 Then
            If bMouseDown Then
                IconBkColor = ShiftColor(m_ImgRightFillColor, vbBlack, 240)
            Else
                IconBkColor = ShiftColor(m_ImgRightFillColor, vbWhite, 150)
            End If
            
            If m_ButtonsGradient Then
                RL.Width = UserControl.ScaleWidth
                RL.Height = UserControl.ScaleHeight
                
                Color2 = RGBtoARGB(ShiftColor(IconBkColor, vbWhite, 200), 100)
                GdipCreateLineBrushFromRectI RL, Color2, RGBtoARGB(IconBkColor, 100), _
                                    LinearGradientModeVertical, WrapModeTileFlipX, hBrush
            Else
                GdipCreateSolidFill RGBtoARGB(IconBkColor, 100), hBrush
            End If
            
            GdipSetClipRectI hGraphics, 0, 0, m_Rect.Right + m_Rect.Left + nScale, UserControl.ScaleHeight, CombineModeExclude
            If HotSpinButton = 1 Then
                GdipSetClipRectI hGraphics, m_Rect.Left, m_Rect.Top, UserControl.ScaleWidth, m_Rect.Bottom / 2, CombineModeExclude
            Else
                GdipSetClipRectI hGraphics, m_Rect.Left, m_Rect.Top + UserControl.ScaleHeight / 2 - CaptionHeight / 4 + BW, UserControl.ScaleWidth, m_Rect.Bottom / 2, CombineModeExclude
            End If
            GdipFillPath hGraphics, hBrush, hPath
            GdipResetClip hGraphics
            GdipDeleteBrush hBrush
        End If

        DrawArrowUp hGraphics, ArrowColor, Left, m_Rect.Top + m_Rect.Bottom / 4 - ArrowSize / 2 + BB / 2, ArrowSize
        DrawArrowDown hGraphics, ArrowColor, Left, (UserControl.ScaleHeight - BW) / 4 * 3 - ArrowSize / 2 - BB / 2, ArrowSize
    ElseIf m_RightButtonStyle = RS_ClearText Then
        DrawCross hGraphics, ArrowColor, Left, m_Rect.Top + m_Rect.Bottom / 2 - ArrowSize / 2 - CaptionHeight / 8, ArrowSize
    ElseIf m_RightButtonStyle = RS_ShowPassword Then
        DrawEye hGraphics, ArrowColor, Left, m_Rect.Top + m_Rect.Bottom / 2 - ArrowSize / 2 - CaptionHeight / 8 + nScale / 2, ArrowSize, bShowPassword
    ElseIf m_RightButtonStyle = RS_DropDown Then
        DrawArrowDown hGraphics, ArrowColor, Left, m_Rect.Top + m_Rect.Bottom / 2 - ArrowSize / 2 - CaptionHeight / 8, ArrowSize
    End If
            
    If m_BorderStyle = [Fixed Single] Then
        GdipCreatePen1 BorderColor, BW, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    ElseIf m_BorderStyle = [Line Bottom] Then
        GdipCreatePen1 BorderColor, BW, UnitPixel, hPen
        GdipDrawLine hGraphics, hPen, BR / 2 + BB + BW, CaptionHeight / 2 + m_Rect.Bottom - BW, UserControl.ScaleWidth - BW - BR / 2 - BB, CaptionHeight / 2 + m_Rect.Bottom - BW
        GdipDeletePen hPen
    End If
    
    If m_RightButtonStyle = RS_Resizable Then
        GdipCreatePen1 BorderColor, nScale, UnitPixel, hPen
        GdipDrawLine hGraphics, hPen, ScaleWidth - BW - BB - 1 - 10 * nScale, ScaleHeight - BW * 1.5 - BB - 1 - 2 * nScale, _
                    ScaleWidth - BW - BB - 1 - 2 * nScale, ScaleHeight - BW * 1.5 - BB - 1 - 10 * nScale
        
        GdipDrawLine hGraphics, hPen, ScaleWidth - BW - BB - 1 - 6 * nScale, ScaleHeight - BW * 1.5 - BB - 1 - 2 * nScale, _
                    ScaleWidth - BW - BB - 1 - 2 * nScale, ScaleHeight - BW * 1.5 - BB - 1 - 6 * nScale
        GdipDeletePen hPen
        GdipDeletePath hPath
    End If
  '-------------
    GdipSetInterpolationMode hGraphics, InterpolationModeHighQualityBicubic
    Call GdipSetPixelOffsetMode(hGraphics, 4&)
    
    If hImgLeft Then
        Dim MyScale As Double, Factor As Double
        Dim ReqWidth As Long, ReqHeight As Long
        
        ImgScale = IIF(ImgLeftRealSize.Y >= ImgLeftRealSize.X, ImgLeftRealSize.Y, ImgLeftRealSize.X)
        Factor = ImgLeftSize * nScale / ImgScale
        ReqWidth = ImgLeftRealSize.X * Factor
        ReqHeight = ImgLeftRealSize.Y * Factor
        
        Left = BB + BW + Margin + (ImgLeftSize * nScale / 2) - ReqWidth / 2
        If m_ImgLeftFillStyle = FS_Tansparent Then Left = Left + Margin / 2
        Top = m_Rect.Top + m_Rect.Bottom / 2 - ReqHeight / 2 - BB / 2  ' CaptionHeight / 8
        GdipDrawImageRectI hGraphics, hImgLeft, Left, Top, ReqWidth, ReqHeight
        'GdipDrawImageRectI hGraphics, hImgLeft, Left, Top, ImgLeftSize * nScale, ImgLeftSize * nScale
    End If
    
    If hImgRight And m_RightButtonStyle = RS_Icon Then
        ImgScale = IIF(ImgRightRealSize.Y >= ImgRightRealSize.X, ImgRightRealSize.Y, ImgRightRealSize.X)
        Factor = ImgRightSize * nScale / ImgScale
        ReqWidth = ImgRightRealSize.X * Factor
        ReqHeight = ImgRightRealSize.Y * Factor
        
        Left = m_Rect.Right + m_Rect.Left + Margin + (ImgRightSize * nScale / 2) - ReqWidth / 2
        Top = m_Rect.Top + m_Rect.Bottom / 2 - ReqHeight / 2 - BB / 2  ' CaptionHeight / 8
        GdipDrawImageRectI hGraphics, hImgRight, Left, Top, ReqWidth, ReqHeight
        
        'GdipDrawImageRectI hGraphics, hImgRight, m_Rect.Right + m_Rect.Left + Margin, m_Rect.Top + m_Rect.Bottom / 2 - ImgRightSize / 2 * nScale, ImgRightSize * nScale, ImgRightSize * nScale
    End If

    If Len(m_Caption) Then
        Dim mSize As Single
        Dim mBold As Boolean
        Dim mForeColor As OLE_COLOR
        Dim R As RECT
        
        With UserControl
            Left = BW + Margin + nScale * 2
            If Left < BR Then Left = BR + Margin
        
            hPath = CreateRoundPath(Left - nScale * 2, nScale, CaptionWidth + nScale * 3, CaptionHeight - nScale * 3, 3 * nScale)
            GdipCreateSolidFill RGBtoARGB(m_ParentBackColor, 100), hBrush
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath
            GdipDeleteGraphics hGraphics

            mSize = .FontSize
            mBold = .FontBold
            mForeColor = .ForeColor
            .FontSize = 7
            .FontBold = True
            With R
                .Left = Left
                .Top = -nScale
                .Right = Left + CaptionWidth
                .Bottom = .Top + CaptionHeight
            End With
            .ForeColor = IIF(bFocus, m_BorderColorOnFocus, m_BorderColor)
            DrawTextW hdc, StrPtr(m_Caption), -1, R, DT_SINGLELINE
            .FontBold = mBold
            .FontSize = mSize
            .ForeColor = mForeColor
        End With
    Else
        GdipDeleteGraphics hGraphics
    End If

    If Me.TextLenght = 0 And Len(m_CueBanner) And bFocus = False Then
        mForeColor = UserControl.ForeColor
        UserControl.ForeColor = ShiftColor(UserControl.ForeColor, m_BackColor, 100)
        DrawTextW hdc, StrPtr(m_CueBanner), -1, m_RectEdit, DT_SINGLELINE
        UserControl.ForeColor = mForeColor
    ElseIf m_InputType = IT_DropDown Then
        DrawTextW hdc, StrPtr(m_Text), -1, m_RectEdit, DT_SINGLELINE
    End If

    UserControl.Refresh
End Sub


'-----------------------------
'SubClass
'-----------------------------
'*-

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As ucText
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEBAAV1aLdCQUg8YIgz4AdC+L+oHHKBIQAIvCBQwREACri8IFSBEQAKuLwgVYERAAq4vCBYAREACruQkAAADzpYHCKBIQAFJqHP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlM8Crg+8cagBX/3IM/3cM/1IYi0QkGIk4Xl+4XBIQAC1wEBAAwhAADx8Ai0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1GIsKUv9xDP9yDP9RHItUJASLClL/URQzwMIEAJBVi+yLVRj/QgT/QhiLQhg7QgR0b4tCEIXAdGiLCotBLIXAdDdS/9BaiUIIg/gBd1OFwHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1E1KLCv91FP91EP91DP91CP9RIFr/ShhQUug4////WF3CGAAPHwA=" ' 9.6.2020 13:56:03
    Const THUNK_SIZE    As Long = 492
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitSubclassingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcByOrdinal(GetModuleHandle("comctl32"), 410)             '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcByOrdinal(GetModuleHandle("comctl32"), 412)             '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcByOrdinal(GetModuleHandle("comctl32"), 413)             '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitSubclassingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function TerminateSubclassingThunk(pSubclass As IUnknown, pObj As Object) As IUnknown
    If Not pSubclass Is Nothing Then
        Debug.Assert ThunkPrivateData(pSubclass, 2) = ObjPtr(pObj)
        ThunkPrivateData(pSubclass, 2) = 0
        Set pSubclass = Nothing
    End If
End Function

Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, PTR_SIZE)
    End If
End Property

Property Let ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long, ByVal lValue As Long)
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, lValue, PTR_SIZE)
    End If
End Property

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long

        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIF(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

Private Sub Subclass_StopAll()
    If IsSubClased Then
        TerminateSubclassingThunk m_pSubclassEdit, Me
        TerminateSubclassingThunk m_pSubclassUC, Me
        IsSubClased = False
    End If
End Sub

Private Sub InitializeSubClassing()
    If Ambient.UserMode And (IsSubClased = False) Then
        Set m_pSubclassEdit = InitSubclassingThunk(hEdit, Me, InitAddressOfMethod(Me, 5).EditProc(0, 0, 0, 0, 0))
        Set m_pSubclassUC = InitSubclassingThunk(UserControl.hWnd, Me, InitAddressOfMethod(Me, 5).UserControlProc(0, 0, 0, 0, 0))
        IsSubClased = True
    End If
End Sub

Public Function EditProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
    Dim ToadoX As Single, ToadoY As Single
    Static bAutoSel As Boolean
    'On Error Resume Next
    ToadoX = WordLo(lParam) * Screen.TwipsPerPixelX
    ToadoY = WordHi(lParam) * Screen.TwipsPerPixelY
    
    Select Case uMsg
        Case WM_CONTEXTMENU, WM_PASTE, WM_CUT
            Dim bCancel As Boolean
            If uMsg = WM_CONTEXTMENU Then
                RaiseEvent ContextMenu(bCancel)
            ElseIf uMsg = WM_PASTE Then
                RaiseEvent Paste(bCancel)
            Else
                RaiseEvent Cut(bCancel)
            End If
            
            If (m_InputType = IT_Date) Or (m_InputType = IT_Time) Then
                bCancel = True
            End If
            
            Handled = bCancel
            Exit Function
        Case WM_LBUTTONDOWN
            If (bFocus = False) Then bAutoSel = True

        Case WM_KEYDOWN
            Dim KeyCode As Integer
            KeyCode = wParam And &HFF&
            If (m_InputType = IT_Date) Or (m_InputType = IT_Time) Then Me_KeyDown KeyCode
            RaiseEvent KeyDown(KeyCode, pvShiftState())

            wParam = KeyCode

            If KeyCode = 0 Then
                Handled = True
                Exit Function
            End If
            
        Case WM_CHAR
            Dim KeyAscii As Integer

            KeyAscii = CUIntToInt(wParam And &HFFFF&)   '&H7FFF&

            Me_KeyPress KeyAscii

            RaiseEvent KeyPress(KeyAscii)

            wParam = KeyAscii

            If KeyAscii = 0 Then
                Handled = True
                Exit Function
            End If
            wParam = CIntToUInt(KeyAscii)
            
        Case WM_KEYUP
            RaiseEvent Change
            RaiseEvent KeyUp(wParam And &H7FFF&, pvShiftState())
            
        Case WM_ERASEBKGND
            If m_EditGradient Then
                Dim R As RECT
                R.Right = UserControl.ScaleWidth
                R.Bottom = UserControl.ScaleHeight
    
                FillRect wParam, R, mBkBrush
                EditProc = 1
                Handled = True
                Exit Function
            End If
            
        Case WM_HSCROLL, WM_VSCROLL
            If m_EditGradient Then RedrawWindow hEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
           
    End Select
    
 '----------BEFORE----------------------------------------
    EditProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
 '----------AFTER-----------------------------------------
 
    Select Case uMsg
   
    
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)

        Case WM_RBUTTONDOWN

            RaiseEvent MouseDown(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)

        Case WM_LBUTTONUP
            RaiseEvent MouseUp(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)
            RaiseEvent Click

        Case WM_LBUTTONDOWN
            If (m_InputType = IT_Date) Or (m_InputType = IT_Time) Then
                AutoSelDatePart
            Else
                If bAutoSel Then
                    bAutoSel = False
                    If (m_OnFocusSelAll = True) Then
                        Me.SelStart = 0
                        Me.SelLength = Me.TextLenght
                    End If
                End If
            End If

            RaiseEvent MouseDown(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)

        Case WM_MOUSEMOVE
            RaiseEvent MouseMove(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)
            If bMouseEnter = False Then
                bMouseEnter = True
                RaiseEvent MouseEnter
            End If

            
        Case WM_LBUTTONDBLCLK
            RaiseEvent DbClick

        Case WM_RBUTTONDBLCLK
            RaiseEvent DbClick

        Case WM_MOUSELEAVE
            Dim pt As POINTAPI
            GetCursorPos pt
            If WindowFromPoint(pt.X, pt.Y) <> UserControl.hWnd Then
                bMouseEnter = False
                RaiseEvent MouseLeave
            End If
            
        Case WM_SETFOCUS
            If wParam <> UserControl.hWnd Then
                PutFocus UserControl.hWnd
                Exit Function
            End If

            Call mIOleInPlaceActivate.SetIPAO(m_uIPAO, Me)

            If m_InputType = IT_Date Then AutoSelDatePart
            
        Case WM_KILLFOCUS
            If Me.TextLenght = 0 And Len(m_CueBanner) Then
                ShowWindow hEdit, vbHide
                bVisible = False
            End If
        Case WM_MOUSEWHEEL
            Dim Val As Integer
            Val = IIF(wParam < 0, -1, 1)
            RaiseEvent MouseWhell(Val)
    End Select
    
    Handled = True

End Function

Public Function UserControlProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
    'On Error Resume Next
    Select Case uMsg
        Case WM_COMMAND
            Select Case WordHi(wParam)
                Case EN_UPDATE, EN_SETFOCUS, EN_KILLFOCUS
                    If m_EditGradient Then RedrawWindow hEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                Case EN_CHANGE
                    
                    If lParam = hEdit Then
                        UserControl.Refresh
                     
                        If TextBoxChangeFrozen = False Then
                        
                            UserControl.PropertyChanged "Text"
                            If (m_InputType = IT_Date) Or (m_InputType = IT_Time) And (HotSpinButton = 0) Then
                                If ((m_MinValue <> vbNullString) Or (m_MaxValue <> vbNullString)) And (Me.TextLenght > 0) Then
                                    Timer1.Enabled = False
                                    Timer1.Interval = 1000
                                    Timer1.Enabled = True
                                End If
                                If IsDate(Me.Text) Then mDate = Me.Text
                            ElseIf m_InputType = IT_LettersOnly Then
                                Dim i As Long, iChar As Integer
                                Dim Sel As Long, sFilter As String
                                m_Text = Me.Text
                                Sel = Me.SelStart
                                For i = 1 To Me.TextLenght
                                    iChar = Asc(Mid(m_Text, i, 1))
                                    If (iChar > 64 And iChar < 91) Or (iChar > 96 And iChar < 123) Then
                                        sFilter = sFilter & Chr(iChar)
                                    End If
                                Next
                                TextBoxChangeFrozen = True
                                Me.Text = sFilter
                                TextBoxChangeFrozen = False
                                Me.SelStart = Sel
                            ElseIf m_InputType = IT_Desimal Then
                                m_Text = Me.Text
                                Sel = Me.SelStart
                                For i = 1 To Me.TextLenght
                                    iChar = Asc(Mid(m_Text, i, 1))
                                    If (iChar > 47 And iChar < 58) Or iChar = 44 Or iChar = 46 Then
                                        sFilter = sFilter & Chr(iChar)
                                    End If
                                Next
                                TextBoxChangeFrozen = True
                                Me.Text = sFilter
                                TextBoxChangeFrozen = False
                                Me.SelStart = Sel
                            ElseIf (m_InputType = IT_Numeric) And (HotSpinButton = 0) Then
                                If m_MinValue <> vbNullString Or m_MaxValue <> vbNullString Then
                                    Timer1.Enabled = False
                                    Timer1.Interval = 1000
                                    Timer1.Enabled = True
                                End If
                            End If
                            
                            RaiseEvent Change

                            If m_EditGradient Then RedrawWindow hEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                        End If
                    Else
                        
                        With ObjListPlus
                            Me.Text = .Text
                            If .ItemIcon(.ListIndex) > -1 Then
                                Me.LoadImgLeft .ImageListGetImage(.ItemIcon(.ListIndex) + 1)
                            End If
                        End With
                    End If

                Case EN_HSCROLL, EN_VSCROLL
                  
                    If m_EditGradient Then RedrawWindow hEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                    RaiseEvent Scroll
                Case Else
                   ' Debug.Print Hex(WordHi(wParam))
            End Select
    End Select
    
    UserControlProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
 
    Select Case uMsg
        Case WM_SETFOCUS

        Case WM_MOUSELEAVE
            Dim pt As POINTAPI
            GetCursorPos pt
            If WindowFromPoint(pt.X, pt.Y) <> hEdit Then
                bMouseEnter = False
                If HotSpinButton > 0 Or HotButton > 0 Or m_InputType = IT_DropDown Or m_HotBorder = True Then
                    HotSpinButton = 0
                    HotButton = 0
                    Draw
                End If
                RaiseEvent MouseLeave
            Else
                If HotSpinButton > 0 Or HotButton > 0 Then
                    HotSpinButton = 0
                    HotButton = 0
                    Draw
                End If
            End If
            
        Case WM_MOUSEMOVE
            Dim ET As TRACKMOUSEEVENTTYPE
            ET.cbSize = Len(ET)
            ET.hwndTrack = hWnd
            ET.dwFlags = TME_LEAVE
            TrackMouseEvent ET
            
        Case WM_MOUSEWHEEL
            Dim Val As Integer
            Val = IIF(wParam < 0, -1, 1)
            RaiseEvent MouseWhell(Val)
            
            If Not ObjListPlus Is Nothing Then
                If Val < 1 Then
                    ObjListPlus.SetKeyDown vbKeyDown
                Else
                    ObjListPlus.SetKeyDown vbKeyUp
                End If
            End If
            
            If m_RightButtonStyle = RS_SpinButton Then
                HotSpinButton = IIF(Val < 1, 1, 2)
                Timer1_Timer
                HotSpinButton = 0
            End If
            
         Case WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC
            If m_EditGradient Then
                SetBkMode wParam, TRANSPARENT
                UserControlProc = mBkBrush
            End If
            
         Case WM_SIZE
            If m_RightButtonStyle = RS_Resizable Then
                If GetAsyncKeyState(1) < 0 Then
                    UserControl_Resize
                    UserControl.Size ScaleWidth * Screen.TwipsPerPixelX, ScaleHeight * Screen.TwipsPerPixelY
                    RaiseEvent Resize
                End If
            End If
            
         Case WM_NCHITTEST
            
            If m_RightButtonStyle = RS_Resizable Then
                Dim R As RECT
                Dim GripSize As Long
                
                GripSize = Margin * 3 + m_BorderWidth * nScale + BB
                
                GetWindowRect hWnd, R
                If R.Right - WordLo(lParam) < GripSize And R.Bottom - WordHi(lParam) < GripSize Then
                    UserControlProc = HTBOTTOMRIGHT
                    Handled = True
                End If
            End If
            
    Case WM_GETMINMAXINFO
        If m_RightButtonStyle = RS_Resizable Then
            Dim tMINMAXINFO As MINMAXINFO
                    
            Call CopyMemory(tMINMAXINFO, ByVal lParam, LenB(tMINMAXINFO))
                tMINMAXINFO.ptMinTrackSize = m_MinSize
                If m_MaxSize.X > 0 Then tMINMAXINFO.ptMaxTrackSize.X = m_MaxSize.X
                If m_MaxSize.Y > 0 Then tMINMAXINFO.ptMaxTrackSize.Y = m_MaxSize.Y
            Call CopyMemory(ByVal lParam, tMINMAXINFO, LenB(tMINMAXINFO))
        End If
        
    Case WM_SYSCOMMAND
        If wParam = SC_DRAGSIZE_SE Then
            RaiseEvent EndSize
        End If
        
    Case Else
        'Debug.Print Hex(uMsg)
        
    End Select

    Handled = True
End Function














