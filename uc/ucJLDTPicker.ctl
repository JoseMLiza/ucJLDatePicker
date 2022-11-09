VERSION 5.00
Begin VB.UserControl ucJLDTPicker 
   AutoRedraw      =   -1  'True
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   ClipBehavior    =   0  'None
   Picture         =   "ucJLDTPicker.ctx":0000
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   64
   ToolboxBitmap   =   "ucJLDTPicker.ctx":0CCA
   Begin VB.Timer tmrMouseEvent 
      Left            =   480
      Top             =   0
   End
End
Attribute VB_Name = "ucJLDTPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------
'Autor: Jose Liza
'Date: 21/05/2022
'Version: 0.0.1
'Thanks: Leandro Ascierto (www.leandroascierto.com) And Latin Group of VB6
'-----------------------------
'--> APIS:
'---> USER32 (User32)
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECTL, ByVal hBrush As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As Any, ByVal wFormat As Long) As Long
'---> KERNEL32 (Kernel32)
'-> Rutinas:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
'-> Funciones:
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsSetValue Lib "kernel32" (ByVal dwTlsIndex As Long, ByVal lpTlsValue As Long) As Long
Private Declare Function TlsFree Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsAlloc Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
'---> OLEAUT32 (Oleaut32)
Private Declare Function OleTranslateColor Lib "Oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
'---> OLEPRO32 (Olepro32)
Private Declare Function OleCreatePictureIndirect Lib "Olepro32" (PicDesc As udtPicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
'---> GDI32 (Gdi32)
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function Polygon Lib "GDI32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function FillPath Lib "GDI32" (ByVal hdc As Long) As Long
'---> GDIPLUS (GdiPlus)
'-> Rutinas:
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
'-> Funciones:
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
'-> Graphics Functions
Private Declare Function GdipCreatePath Lib "gdiplus" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "gdiplus" (ByVal mGraphics As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipResetPath Lib "gdiplus" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal mPath As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipAddPathArcI Lib "gdiplus" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathPolygon Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTF, ByVal mCount As Long) As Long
Private Declare Function GdipClosePathFigures Lib "gdiplus" (ByVal mPath As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "gdiplus" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As enmWrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipCreateLineBrush Lib "gdiplus" (ByRef mPoints1 As POINTF, ByRef mPoints2 As POINTF, ByVal mColor1 As Long, ByVal nColor2 As Long, ByVal mWrapMode As enmWrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "gdiplus" (ByVal mGraphics As Long) As Long
Private Declare Function GdipSetPenMode Lib "gdiplus" (ByVal mPen As Long, ByVal mPenMode As enmPenAlignment) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal mBitmap As Long, ByRef mRect As RECTL, ByVal mFlags As enmImageLockMode, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As udtBitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal mBitmap As Long, ByRef mLockedBitmapData As udtBitmapData) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByRef pPoints As Any, ByVal count As Long, ByVal fillMode As Long) As Long
Private Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As Any, ByVal count As Long, ByVal tension As Single) As Long
Private Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, Points As Any, ByVal count As Long, ByVal tension As Single, ByVal fillMode As Long) As Long
Private Declare Function GdipAddPathLine2I Lib "gdiplus" (ByVal mPath As Long, ByRef mPoints As Any, ByVal mCount As Long) As Long
'-> String Functions
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal mFormat As Long, ByVal mFlags As enmStringFormatFlags) As Long
Private Declare Function GdipSetStringFormatHotkeyPrefix Lib "gdiplus" (ByVal mFormat As Long, ByVal mHotkeyPrefix As enmHotkeyPrefix) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "gdiplus" (ByVal mFormat As Long, ByVal mTrimming As enmStringTrimming) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As enmStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal mFormat As Long, ByVal mAlign As enmStringAlignment) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTS, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
Private Declare Function GdipAddPathString Lib "gdiplus" (ByVal mPath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal mFormat As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
'-> Font Functions
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal mFont As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
'-> Pen / Brush Functions
Private Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal mHatchStyle As HatchStyle, ByVal mForecol As Long, ByVal mBackcol As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePen2 Lib "gdiplus" (ByVal mBrush As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal mPen As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long

'--> Constantes
Private Const cs_CarNumeros             As String = "0123456789"
Private Const cs_ColsDay                As Integer = 7
Private Const cs_RowsPicker             As Integer = 7
Private Const cs_ColsMonthYear          As Integer = 4
Private Const cs_RowsMonthYear          As Integer = 3
Private Const cs_ItemsDay               As Integer = cs_ColsDay * (cs_RowsPicker - 1)
'--
Private Const SORT_DEFAULT              As Long = &H0& 'sorting default
Private Const LANG_NEUTRAL              As Long = &H0&
Private Const LANG_INVARIANT            As Long = &H7F&
Private Const SUBLANG_NEUTRAL           As Long = &H0&  'language neutral
Private Const SUBLANG_DEFAULT           As Long = &H1&  'user default
Private Const SUBLANG_SYS_DEFAULT       As Long = &H2&  'system default
Private Const LANG_SYSTEM_DEFAULT       As Long = LANG_NEUTRAL Or SUBLANG_SYS_DEFAULT * &H400&
Private Const LANG_USER_DEFAULT         As Long = LANG_NEUTRAL Or SUBLANG_DEFAULT * &H400&
'--
Private Const TME_LEAVE                 As Long = &H2&
Private Const UnitPixel                 As Long = &H2&
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const TLS_MINIMUM_AVAILABLE     As Long = 64
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const CombineModeExclude        As Long = &H4
Private Const IDC_HAND                  As Long = 32649
Private Const WS_CHILD                  As Long = &H40000000
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const GWL_HWNDPARENT            As Long = -8
Private Const HWND_TOPMOST              As Long = -1
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_SHOWWINDOW            As Long = &H40
'--
Private Const DT_LEFT                   As Long = &H0
Private Const DT_TOP                    As Long = &H0
Private Const DT_RIGHT                  As Long = &H2
Private Const DT_BOTTOM                 As Long = &H8
Private Const DT_CENTER                 As Long = &H1
Private Const DT_VCENTER                As Long = &H4
Private Const DT_WORDBREAK              As Long = &H10
Private Const DT_SINGLELINE             As Long = &H20
Private Const DT_CALCRECT               As Long = &H400
'--
Private Const WM_HOTKEY                 As Long = &H312
Private Const WM_MOUSELEAVE             As Long = &H2A3&
Private Const WM_CHAR                   As Long = &H102
Private Const WM_IME_SETCONTEXT         As Long = &H281&

'--> Tipos definidos por usuario.
Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type

Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type RECTL
    Left                                As Long
    Top                                 As Long
    Width                               As Long
    Height                              As Long
End Type

Private Type RECTS
    Left                                As Single
    Top                                 As Single
    Width                               As Single
    Height                              As Single
End Type

Private Type Radius
    TopLeft                             As Integer
    TopRight                            As Integer
    BottomLeft                          As Integer
    BottomRight                         As Integer
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type POINTF
    X As Single
    Y As Single
End Type

Private Type udtPicBmp
    Size As Long
    type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type udtBitmapData
    Width                               As Long
    Height                              As Long
    stride                              As Long
    PixelFormat                         As Long
    Scan0Ptr                            As Long
    ReservedPtr                         As Long
End Type
'--> Calendarios:

Private Type udtTitleMonthYear
    RECT                                As RECTL
    RECT2                               As RECT
    MouseState                          As enmMouseState
End Type

Private Type udtItemDatePicker
    RECT                                As RECTL    'RECTL Info.
    RECT2                               As RECT    'RECT Info.
    NumberMonth                         As Integer  'Número del mes
    MonthName                           As String   'Nombre del mes
    NumberYear                          As Integer  'Número del año
    DateInPicker                        As Date
    HeaderTitle                         As String
    TitleMonthYear                      As udtTitleMonthYear
    IndexCalendar                       As Integer  'Index del calendario al que pertenece el item.
    ViewNavigator                       As enmViewItemNavigator
End Type

'--> Items para meses o años segun navegacion rapida.
Private Type udtItemMonthYear
    RECT                                As RECTL
    RECT2                               As RECT
    Caption                             As String   'Text
    IsCurrentPart                       As Boolean  'Es parte de la fecha (mes o año)
    ValueItem                           As Long     'Numero del mes (1 to 12) o año (####)
    IndexCalendar                       As Integer  'Index del calendario al que pertenece el item.
    MouseState                          As enmMouseState
End Type

'--> Items de la cabecera del día. (Do, Lu, Ma, Mi, Ju, Vi, Sa):
Private Type udtItemHeaderDay
    RECT                                As RECTL    'RECT Info.
    Caption                             As String   'Texto
    DayName                             As String   'Nombre del día
    NumberWeekDay                       As Integer  'Numero del dia de la semana
    IndexCalendar                       As Integer  'Index del calendario al que pertenece el item.
End Type

'--> Items de los numeros de las semanas:
Private Type udtItemWeek
    RECT                                As RECTL    'RECT Info.
    Caption                             As String   'Texto
    NumberWeek                          As Integer  'Numero de la semana (del año)
    WeekInMonth                         As Integer  'Mes al que pertenece la semana
    WeekInYear                          As Integer  'Año al que pertenece la semana
    IndexCalendar                       As Integer  'Index del calendario al que pertenece el item.
End Type

'--> Items de los días del calendario.
Private Type udtItemDayCalendar
    RECT                                As RECTL    'RECT Info.
    RECT2                               As RECT     'REC Info.
    Caption                             As String   'Texto
    DateValue                           As String   'Fecha de dia
    DatePartDay                         As Integer  'Numero del día
    DatePartMonth                       As Integer  'Numero del mes
    DatePartYear                        As Integer  'Numero del año
    IsNow                               As Boolean  'Si es la fecha actual
    IsValueDate                         As Boolean  'Si el dia es igual a la propiedad value
    IsStartDate                         As Boolean  'Si el item de el dia inicial de la seleccion
    IsEndDate                           As Boolean  'Si el item de el dia final de la seleccion
    IsBetweenDate                       As Boolean  'Si el item es un dia entre el inicial y final de la seleccion
    IsDayInMonthCurrent                 As Boolean  'Fecha del día pertece al mes del calendario (para desactivar en la seleccion)
    NumberWeek                          As Integer  'Semana a la que pertenece el día.
    IsDaySaturday                       As Boolean  'Es día de fin de semana (Sabado).
    IsDaySunday                         As Boolean  'Es día de fin de semana (Domingo).
    IsFreeDay                           As Boolean  'Es día libre (es feriado)
    IndexCalendar                       As Integer  'Index del calendario al que pertenece el item.
    MouseState                          As enmMouseState
End Type

'--> Items de las horas, minutos y segundos del timePicker.
Private Type udtItemHourMinSecDay
    RECT                                As RECTL
    RECT2                               As RECT
    Caption                             As String
    Value                               As Integer
    IsCurrentPart                       As Boolean
    IndexCalendar                       As Integer
    MouseState                          As enmMouseState
End Type

'--> Items de los botones de accion "< >"
Private Type udtItemCalendarButton
    RECT                                As RECTL    'RECT Info.
    RECT2                               As RECT
    Caption                             As String
    IndexCalendar                       As Integer
    IconCharCode                        As Long
    MouseState                          As enmMouseState
    ButtonAction                        As enmButtonAction
    IsVisible                           As Boolean
End Type

'--> Enumaradores:
'---> Privados:
Private Enum enmViewItemNavigator
    ViewItemNavigatorSeconds
    ViewItemNavigatorMinutes
    ViewItemNavigatorHour
    ViewItemNavigatorDays
    ViewItemNavigatorMonths
    ViewItemNavigatorYears
End Enum

Private Enum enmHotkeyPrefix
    HotkeyPrefixNone = &H0
    HotkeyPrefixShow = &H1
    HotkeyPrefixHide = &H2
End Enum

Private Enum enmPenAlignment
    PenAlignmentCenter = &H0
    PenAlignmentInset = &H1
End Enum

Private Enum enmImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum

Private Enum enmWrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Private Enum enmStringTrimming
    StringTrimmingNone = &H0
    StringTrimmingCharacter = &H1
    StringTrimmingWord = &H2
    StringTrimmingEllipsisCharacter = &H3
    StringTrimmingEllipsisWord = &H4
    StringTrimmingEllipsisPath = &H5
End Enum

Private Enum enmStringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum

Private Enum enmStringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Private Enum enmGDIPFontStyle
    GDIPFontStyleRegular = 0
    GDIPFontStyleBold = 1
    GDIPFontStyleItalic = 2
    GDIPFontStyleBoldItalic = 3
    GDIPFontStyleUnderline = 4
    GDIPFontStyleStrikeout = 8
End Enum

Private Enum enmDrawSection
    DrawSectionParentBack
    DrawSectionControlContainer
    DrawSectionButtonsNav
    DrawSectionButtonsRange
    DrawSectionButtonsAction
    DrawSectionMonthYear
    DrawSectionHeaderDays
    DrawSectionWeeks
    DrawSectionDays
End Enum

Public Enum enmArrowDirectionConstant
    ArrowDirectionNone = 0
    ArrowDirectionDown = 1
    ArrowDirectionUp = 2
    ArrowDirectionLeft = 3
    ArrowDirectionRight = 4
End Enum

Private Enum enmMouseState
    Normal
    Hot
    Pressed
    Inherint
End Enum

Private Enum HatchStyle
    HatchStyleHorizontal = &H0
    HatchStyleVertical = &H1
    HatchStyleForwardDiagonal = &H2
    HatchStyleBackwardDiagonal = &H3
    HatchStyleCross = &H4
    HatchStyleDiagonalCross = &H5
    HatchStyle05Percent = &H6
    HatchStyle10Percent = &H7
    HatchStyle20Percent = &H8
    HatchStyle25Percent = &H9
    HatchStyle30Percent = &HA
    HatchStyle40Percent = &HB
    HatchStyle50Percent = &HC
    HatchStyle60Percent = &HD
    HatchStyle70Percent = &HE
    HatchStyle75Percent = &HF
    HatchStyle80Percent = &H10
    HatchStyle90Percent = &H11
    HatchStyleLightDownwardDiagonal = &H12
    HatchStyleLightUpwardDiagonal = &H13
    HatchStyleDarkDownwardDiagonal = &H14
    HatchStyleDarkUpwardDiagonal = &H15
    HatchStyleWideDownwardDiagonal = &H16
    HatchStyleWideUpwardDiagonal = &H17
    HatchStyleLightVertical = &H18
    HatchStyleLightHorizontal = &H19
    HatchStyleNarrowVertical = &H1A
    HatchStyleNarrowHorizontal = &H1B
    HatchStyleDarkVertical = &H1C
    HatchStyleDarkHorizontal = &H1D
    HatchStyleDashedDownwardDiagonal = &H1E
    HatchStyleDashedUpwardDiagonal = &H1F
    HatchStyleDashedHorizontal = &H20
    HatchStyleDashedVertical = &H21
    HatchStyleSmallConfetti = &H22
    HatchStyleLargeConfetti = &H23
    HatchStyleZigZag = &H24
    HatchStyleWave = &H25
    HatchStyleDiagonalBrick = &H26
    HatchStyleHorizontalBrick = &H27
    HatchStyleWeave = &H28
    HatchStylePlaid = &H29
    HatchStyleDivot = &H2A
    HatchStyleDottedGrid = &H2B
    HatchStyleDottedDiamond = &H2C
    HatchStyleShingle = &H2D
    HatchStyleTrellis = &H2E
    HatchStyleSphere = &H2F
    HatchStyleSmallGrid = &H30
    HatchStyleSmallCheckerBoard = &H31
    HatchStyleLargeCheckerBoard = &H32
    HatchStyleOutlinedDiamond = &H33
    HatchStyleSolidDiamond = &H34
    HatchStyleTotal = &H35
    HatchStyleLargeGrid = &H4
    HatchStyleMin = &H0
    HatchStyleMax = &H34
End Enum

Private Enum enmLCIDs
    'Add more as you need them here:
    LOCALE_SYSTEM_DEFAULT = LANG_SYSTEM_DEFAULT Or SORT_DEFAULT * &H10000
    LOCALE_USER_DEFAULT = LANG_USER_DEFAULT Or SORT_DEFAULT * &H10000
    LOCALE_NEUTRAL = (LANG_NEUTRAL Or SUBLANG_NEUTRAL * &H400&) Or SORT_DEFAULT * &H10000
    LOCALE_INVARIANT = (LANG_INVARIANT Or SUBLANG_NEUTRAL * &H400&) Or SORT_DEFAULT * &H10000

    LOCALE_ENUK = &H809&
    LOCALE_ENUS = &H409&
    LOCALE_DEDE = &H407& 'German, Germany.
End Enum

Private Enum enmLocaleTypes
    LOCALE_NOUSEROVERRIDE = &H80000000          'do not use user overrides
    LOCALE_USE_CP_ACP = &H40000000              'use the system ACP
    LOCALE_RETURN_NUMBER = &H20000000           'return number instead of string
    LOCALE_ILANGUAGE = &H1&                     'language id
    LOCALE_SLANGUAGE = &H2&                     'localized name of language
    LOCALE_SENGLANGUAGE = &H1001&               'English name of language
    LOCALE_SABBREVLANGNAME = &H3&               'abbreviated language name
    LOCALE_SNATIVELANGNAME = &H4&               'native name of language
    
    LOCALE_SDATE = &H1D&                        'date separator (derived from LOCALE_SSHORTDATE, use that instead)
    LOCALE_STIME = &H1E&                        'time separator (derived from LOCALE_STIMEFORMAT, use that instead)
    LOCALE_SSHORTDATE = &H1F&                   'short date format string
    LOCALE_SLONGDATE = &H20&                    'long date format string
    LOCALE_STIMEFORMAT = &H1003&                'time format string
    LOCALE_IDATE = &H21&                        'short date format ordering (derived from LOCALE_SSHORTDATE, use that instead)
    LOCALE_ILDATE = &H22&                       'long date format ordering (derived from LOCALE_SLONGDATE, use that instead)
    LOCALE_ITIME = &H23&                        'time format specifier (derived from LOCALE_STIMEFORMAT, use that instead)
    LOCALE_ITIMEMARKPOSN = &H1005&              'time marker position (derived from LOCALE_STIMEFORMAT, use that instead)
    LOCALE_ICENTURY = &H24&                     'century format specifier (short date, LOCALE_SSHORTDATE is preferred)
    LOCALE_ITLZERO = &H25&                      'leading zeros in time field (derived from LOCALE_STIMEFORMAT, use that instead)
    LOCALE_IDAYLZERO = &H26&                    'leading zeros in day field (short date, LOCALE_SSHORTDATE is preferred)
    LOCALE_IMONLZERO = &H27&                    'leading zeros in month field (short date, LOCALE_SSHORTDATE is preferred)
    LOCALE_S1159 = &H28&                        'AM designator
    LOCALE_S2359 = &H29&                        'PM designator
    
    LOCALE_IFIRSTDAYOFWEEK = &H100C&            'first day of week specifier
    LOCALE_IFIRSTWEEKOFYEAR = &H100D&           'first week of year specifier
End Enum

'---> Publicos:
Public Enum enmCallOutPosition
    [Position Left]
    [Position Top]
    [Position Right]
    [Position Bottom]
End Enum

Public Enum enmCallOutAlign
    [First Corner]
    [Middle]
    [Second Corner]
    [Custom Position]
End Enum

Public Enum enmHotLinePosition
    [HotLine Left]
    [HotLine Top]
    [HotLine Right]
    [HotLine Bottom]
End Enum

Public Enum enmBorderPosition
    [Border None]
    [Border Inside]
    [Border Center]
    [Border Outside]
End Enum

Public Enum enmSelectionStyle
    [None]
    [Corner No Between]
    [Corner Full]
End Enum

Public Enum enmButtonSection
    [Buttons Action]
    [Buttons Range]
End Enum

Public Enum enmButtonAction
    [Action None]
    [Action Apply]
    [Action Cancel]
    [Action Today]
End Enum

'--> Variables locales:
Dim nScale                              As Single
Dim hImgShadow                          As Long
Dim GdipToken                           As Long
Dim iCalendar                           As Integer
'Dim hFontCollection                     As Long
'--
Dim c_PT                                As POINTAPI
Dim c_hWnd                              As Long
Dim c_Left                              As Long
Dim c_Top                               As Long
Dim c_Width                             As Long
Dim c_Height                            As Long
Dim c_bIntercept                        As Boolean
'--
Dim c_EnterControl                      As Boolean
Dim c_EnterButton                       As Boolean
'--
Dim udtItemsMonthYear()                 As udtItemMonthYear
Dim udtItemsPicker()                    As udtItemDatePicker
Dim udtItemsNavButton()                 As udtItemCalendarButton
Dim udtItemsUpDownButton(1)             As udtItemCalendarButton    '(Siempre 1 picker por la navegacion rapida: 2 items(0 to 1))
Dim udtItemsRangeButton(5)              As udtItemCalendarButton    'Para botones de rangos ('Hoy', 'Este mes', 'Mes pasado', 'Ultimos 90 días', 'Este Año', 'Año pasado')
Dim udtItemsActionButton(2)             As udtItemCalendarButton    'Para botones de accion ('Hoy', 'Cancelar', 'Aplicar')
Dim udtItemsHeaderDay()                 As udtItemHeaderDay
Dim udtItemsWeek()                      As udtItemWeek
Dim udtItemsDay()                       As udtItemDayCalendar
'--
Dim c_EnBoton                           As Boolean
Dim c_IndexSelMove                      As Integer
'--
Dim c_Show                              As Boolean
Dim c_PhWnd                             As Long
Dim c_SubClass                          As clsSubClass
Dim c_CShadow                           As clsShadow
'--
Dim b_ShowFastNavigator                 As Boolean
'--
Dim d_ValueTemp                         As Date
'--

'--> Propiedades:
'---> Del control:
Private m_BackColor                     As OLE_COLOR
Private m_BackOpacity                   As Integer
Private m_Border                        As Boolean
Private m_BorderWidth                    As Integer
Private m_BorderColor                   As OLE_COLOR
Private m_BorderOpacity                 As Integer
Private m_BorderPosition                As enmBorderPosition
Private m_CornerTopLeft                 As Integer
Private m_CornerTopRight                As Integer
Private m_CornerBottomLeft              As Integer
Private m_CornerBottomRight             As Integer
Private m_PaddingX                      As Integer
Private m_PaddingY                      As Integer
Private m_Redraw                        As Boolean
Private m_Shadow                        As Boolean
Private m_ShadowSize                    As Integer
Private m_ShadowColor                   As OLE_COLOR
Private m_ShadowOpacity                 As Integer
Private m_ShadowOffsetX                 As Integer
Private m_ShadowOffsetY                 As Integer
Private m_ShowNumberWeek                As Boolean
Private m_ShowUseISOWeek                As Boolean
Private m_SpaceGrid                     As Integer
Private m_MouseToParent                 As Boolean
Private m_UseGDIPString                 As Boolean
'--
Private m_PT                            As POINTAPI
Private m_Left                          As Long
Private m_Top                           As Long
Private m_Over                          As Boolean
Private m_Enter                         As Boolean
'--
Private m_ColsPicker                    As Integer
Private m_NumberPickers                 As Integer
'Private m_ShowTimePicker                As Boolean 'Para trabajar el timepicker.
'Private m_UseTimePicker24Hrs            As Boolean 'Para formato de 24 horas en timepicker.
'Private m_TimerWithSecond               As Boolean 'Para usar el timer con segundos.
Private m_MaxRangeDays                  As Integer 'Para limitar el maximo de días en el selectionrange.
'Private m_AlwaysShowCalendars           As Boolean
Private m_LinkedCalendars               As Boolean
Private m_Value                         As Date
Private m_ValueStart                    As String
Private m_ValueEnd                      As String
Private m_MinDate                       As Date
Private m_MaxDate                       As Date
Private m_FirstDayOfWeek                As VbDayOfWeek
Private m_CountFreeDays                 As Boolean
Private m_CountReservedDay              As Boolean
Private m_CountSelDays                  As Integer
Private m_SinglePicker                  As Boolean
Private m_UseRangeValue                 As Boolean
Private m_ShowRangeButtons              As Boolean
Private m_ShowTodayButton               As Boolean
Private m_RightToLeft                   As Boolean
Private m_AutoApply                     As Boolean
Private m_IsChild                       As Boolean
Private m_BackColorParent               As OLE_COLOR
'--

'---> De los botones de navegacion.
Private m_ButtonNavBackColor            As OLE_COLOR
Private m_ButtonNavBorderWidth          As Integer
Private m_ButtonNavBorderColor          As OLE_COLOR
Private m_ButtonNavCornerRadius         As Integer
Private m_ButtonNavForeColor            As OLE_COLOR
Private m_ButtonNavIsIcoFont            As Boolean
Private m_ButtonNavIcoFont              As StdFont
Private m_ButtonNavCharCodeBack         As Long
Private m_ButtonNavCharCodeNext         As Long
Private m_ButtonNavWidth                As Long
Private m_ButtonNavHeight               As Long

'---> De los botones de accion.
Private m_ButtonsBackColor              As OLE_COLOR
Private m_ButtonsBorderWidth            As Integer
Private m_ButtonsBorderColor            As OLE_COLOR
Private m_ButtonsCornerRadius           As Integer
Private m_ButtonsFont                   As StdFont
Private m_ButtonsForeColor              As OLE_COLOR
Private m_ButtonsWidth                  As Long
Private m_ButtonsHeight                 As Long

'---> De los meses y años
Private m_MonthYearBackColor            As OLE_COLOR
Private m_MonthYearBorderWidth          As Integer
Private m_MonthYearBorderColor          As OLE_COLOR
Private m_MonthYearCornerRadius         As Integer
Private m_MonthYearFont                 As StdFont
Private m_MonthYearForeColor            As OLE_COLOR

'---> De las Semanas.
Private m_WeekBackColor                 As OLE_COLOR
Private m_WeekBorderWidth               As Integer
Private m_WeekBorderColor               As OLE_COLOR
Private m_WeekCornerRadius              As Boolean
Private m_WeekFont                      As StdFont
Private m_WeekFontHeaderBold            As Boolean
Private m_WeekForeColor                 As OLE_COLOR
Private m_WeekWidth                     As Long
Private m_WeekHeight                    As Long

'---> De los días.
Private m_DayBackColor                  As OLE_COLOR
Private m_DayBorderWidth                As Integer
Private m_DayBorderColor                As OLE_COLOR
Private m_DayCornerRadius               As Integer
Private m_DayFont                       As StdFont
Private m_DayForeColor                  As OLE_COLOR
Private m_DayHeaderFontBold             As Boolean
Private m_DayHeaderForeColor            As OLE_COLOR
Private m_DayHotColor                   As OLE_COLOR
Private m_DayOMForeColor                As OLE_COLOR
Private m_DayFreeArray                  As Variant
Private m_DayFreeForeColor              As OLE_COLOR
Private m_DayNowShow                    As Boolean
Private m_DayNowBorderWidth             As Integer
Private m_DayNowBorderColor             As OLE_COLOR
Private m_DayNowBackColor               As OLE_COLOR
Private m_DayNowForeColor               As OLE_COLOR
Private m_DaysPrePaintCount             As Long
Private m_DaySelCount                   As Long

'Mouse event (Los días no tendran color para el mousedown)
Private m_DayOverBackColor              As OLE_COLOR
Private m_DayOverForeColor              As OLE_COLOR
'--
Private m_DaySelBetweenColor            As OLE_COLOR
Private m_DaySelValuesColor             As OLE_COLOR
Private m_DaySelForeColor               As OLE_COLOR
Private m_DaySelFontBold                As Boolean
Private m_DaySelectionStyle             As enmSelectionStyle
'Private m_DayShowHotItem                As Boolean 'Siempre muestra el hot
Private m_DaySaturdayForeColor          As OLE_COLOR
Private m_DaySundayForeColor            As OLE_COLOR
Private m_DayWidth                      As Long
Private m_DayHeight                     As Long

'CallOut
Private m_CallOut                       As Boolean
Private m_CallOutWidth                  As Long
Private m_CallOutHight                  As Long
Private m_CallOutRightTriangle          As Boolean
Private m_CallOutPosition               As enmCallOutPosition
Private m_CallOutCustomPosPercent       As Long
Private m_CallOutAlign                  As enmCallOutAlign

Private m_UserFirstDayOfWeek            As Long

'Eventos del control
Public Event DayPrePaint(ByVal dDate As Date, BackColor As Long)
Public Event ChangeMinDate()
Public Event ChangeMaxDate()
Public Event ChangeDate(ByVal Value As Date)
Public Event ChangeStartDate(ByVal Value As String)
Public Event ChangeEndDate(ByVal Value As String)
Public Event ButtonRangeClick(ByVal Index, Caption As String)
Public Event ButtonActionClick(ByVal Index, Caption As String)
Public Event MouseEnter()
Public Event MouseOver()
Public Event MouseLeave()
Public Event MouseOut()
Public Event MouseMove()
Public Event Resize()
'---
'---> Del control:
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'm_BackColor                   As OLE_COLOR
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal Value As OLE_COLOR)
    m_BackColor = Value
    c_CShadow.BackColor = m_BackColor
    PropertyChanged "BackColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_BackOpacity                   As Integer
Public Property Get BackOpacity() As Integer
    BackOpacity = m_BackOpacity
End Property
Public Property Let BackOpacity(ByVal Value As Integer)
    m_BackOpacity = Value
    Call SafeRange(m_BackOpacity, 0, 100)
    c_CShadow.BackOpacity = m_BackOpacity
    PropertyChanged "BackOpacity"
    If m_IsChild Then Draw: Refresh
End Property

'm_Border                        As Boolean
Public Property Get Border() As Boolean
    Border = m_Border
End Property
Public Property Let Border(ByVal Value As Boolean)
    m_Border = Value
    c_CShadow.Border = m_Border
    PropertyChanged "Border"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_BorderWidth                    As Integer
Public Property Get BorderWidth() As Integer
    BorderWidth = m_BorderWidth
End Property
Public Property Let BorderWidth(ByVal Value As Integer)
    m_BorderWidth = Value
    c_CShadow.BorderWidth = m_BorderWidth
    PropertyChanged "BorderWidth"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_BorderColor                   As OLE_COLOR
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = Value
    c_CShadow.BorderColor = m_BorderColor
    PropertyChanged "BorderColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_BorderOpacity                 As Integer
Public Property Get BorderOpacity() As Integer
    BorderOpacity = m_BorderOpacity
End Property
Public Property Let BorderOpacity(ByVal Value As Integer)
    m_BorderOpacity = Value
    c_CShadow.BorderOpacity = m_BorderOpacity
    Call SafeRange(m_BorderOpacity, 0, 100)
    PropertyChanged "BorderOpacity"
    If m_IsChild Then Draw: Refresh
End Property

'm_CornerTopLeft                 As Integer
Public Property Get CornerTopLeft() As Integer
    CornerTopLeft = m_CornerTopLeft
End Property
Public Property Let CornerTopLeft(ByVal Value As Integer)
    m_CornerTopLeft = Value
    c_CShadow.CornerTopLeft = m_CornerTopLeft
    PropertyChanged "CornerTopLeft"
    If m_IsChild Then Draw: Refresh
End Property

'm_CornerTopRight                As Integer
Public Property Get CornerTopRight() As Integer
    CornerTopRight = m_CornerTopRight
End Property
Public Property Let CornerTopRight(ByVal Value As Integer)
    m_CornerTopRight = Value
    c_CShadow.CornerTopRight = m_CornerTopRight
    PropertyChanged "CornerTopRight"
    If m_IsChild Then Draw: Refresh
End Property

'm_CornerBottomLeft              As Integer
Public Property Get CornerBottomLeft() As Integer
    CornerBottomLeft = m_CornerBottomLeft
End Property
Public Property Let CornerBottomLeft(ByVal Value As Integer)
    m_CornerBottomLeft = Value
    c_CShadow.CornerTopLeft = m_CornerBottomLeft
    PropertyChanged "CornerBottomLeft"
    If m_IsChild Then Draw: Refresh
End Property

'm_CornerBottomRight             As Integer
Public Property Get CornerBottomRight() As Integer
    CornerBottomRight = m_CornerBottomRight
End Property
Public Property Let CornerBottomRight(ByVal Value As Integer)
    m_CornerBottomRight = Value
    c_CShadow.CornerBottomRight = m_CornerBottomRight
    PropertyChanged "CornerBottomRight"
    If m_IsChild Then Draw: Refresh
End Property

'm_PaddingX                      As Integer
Public Property Get PaddingX() As Integer
    PaddingX = m_PaddingX
End Property
Public Property Let PaddingX(ByVal Value As Integer)
    m_PaddingX = Value
    InitControl
    PropertyChanged "PaddingX"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_PaddingY                      As Integer
Public Property Get PaddingY() As Integer
    PaddingY = m_PaddingY
End Property
Public Property Let PaddingY(ByVal Value As Integer)
    m_PaddingY = Value
    InitControl
    PropertyChanged "PaddingY"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_Redraw                        As Boolean
Public Property Get Redraw() As Boolean
    Redraw = m_Redraw
End Property
Public Property Let Redraw(ByVal Value As Boolean)
    m_Redraw = Value
End Property

'm_Shadow                        As Boolean
Public Property Get Shadow() As Boolean
    Shadow = m_Shadow
End Property
Public Property Let Shadow(ByVal Value As Boolean)
    m_Shadow = Value
    c_CShadow.Shadow = m_Shadow
    PropertyChanged "Shadow"
    'Draw
    'Refresh
End Property

'm_ShadowSize                    As Integer
Public Property Get ShadowSize() As Integer
    ShadowSize = m_ShadowSize
End Property
Public Property Let ShadowSize(ByVal Value As Integer)
    m_ShadowSize = Value
    c_CShadow.ShadowSize = m_ShadowSize
    PropertyChanged "ShadowSize"
    'Draw
    'Refresh
End Property

'm_ShadowColor                   As OLE_COLOR
Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property
Public Property Let ShadowColor(ByVal Value As OLE_COLOR)
    m_ShadowColor = Value
    c_CShadow.ShadowColor = m_ShadowColor
    PropertyChanged "ShadowColor"
    'Draw
    'Refresh
End Property

'm_ShadowOpacity                 As Integer
Public Property Get ShadowOpacity() As Integer
    ShadowOpacity = m_ShadowOpacity
End Property

Public Property Let ShadowOpacity(ByVal Value As Integer)
    m_ShadowOpacity = Value
    Call SafeRange(m_ShadowOpacity, 0, 100)
    c_CShadow.ShadowOpacity = m_ShadowOpacity
    PropertyChanged "ShadowOpacity"
    'Draw
    'Refresh
End Property

'm_ShadowOffsetX                 As Integer
Public Property Get ShadowOffsetX() As Integer
    ShadowOffsetX = m_ShadowOffsetX
End Property

Public Property Let ShadowOffsetX(ByVal Value As Integer)
    m_ShadowOffsetX = Value
    c_CShadow.ShadowOffsetX = m_ShadowOffsetX
    PropertyChanged "ShadowOffsetX"
    'Draw
    'Refresh
End Property

'm_ShadowOffsetY                 As Integer
Public Property Get ShadowOffsetY() As Integer
    ShadowOffsetY = m_ShadowOffsetY
End Property
Public Property Let ShadowOffsetY(ByVal Value As Integer)
    m_ShadowOffsetY = Value
    c_CShadow.ShadowOffsetY = m_ShadowOffsetY
    PropertyChanged "ShadowOffsetY"
    'Draw
    'Refresh
End Property

'm_ShowNumberWeek                As Boolean
Public Property Get ShowNumberWeek() As Boolean
    ShowNumberWeek = m_ShowNumberWeek
End Property
Public Property Let ShowNumberWeek(ByVal Value As Boolean)
    m_ShowNumberWeek = Value
    PropertyChanged "ShowNumberWeek"
    InitControl
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_ShowUseISOWeek                As Boolean
Public Property Get ShowUseISOWeek() As Boolean
    ShowUseISOWeek = m_ShowUseISOWeek
End Property
Public Property Let ShowUseISOWeek(ByVal Value As Boolean)
    m_ShowUseISOWeek = Value
    PropertyChanged "ShowUseISOWeek"
    If m_IsChild Then Draw: Refresh
End Property

'm_SpaceGrid                     As Integer
Public Property Get SpaceGrid() As Integer
    SpaceGrid = m_SpaceGrid
End Property
Public Property Let SpaceGrid(ByVal Value As Integer)
    m_SpaceGrid = Value
    Call SafeRange(m_SpaceGrid, 1, m_DayWidth)
    InitControl
    PropertyChanged "SpaceGrid"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_MouseToParent                 As Boolean
Public Property Get MouseToParent() As Boolean
    MouseToParent = m_MouseToParent
End Property

Public Property Let MouseToParent(ByVal Value As Boolean)
    m_MouseToParent = Value
    PropertyChanged "MouseToParent"
End Property

'm_UseGDIPString                 As Boolean
Public Property Get UseGDIPString() As Boolean
    UseGDIPString = m_UseGDIPString
End Property

Public Property Let UseGDIPString(ByVal Value As Boolean)
    m_UseGDIPString = Value
    PropertyChanged "UseGDIPString"
    Draw
End Property

'm_Over
Public Property Get IsMouseOver() As Boolean
    IsMouseOver = m_Over
End Property

'm_Enter
Public Property Get IsMouseEnter() As Boolean
    IsMouseEnter = m_Enter
End Property

'--
'm_SinglePicker                  As Boolean
Public Property Get SinglePicker() As Boolean
    SinglePicker = m_SinglePicker
End Property
Public Property Let SinglePicker(ByVal Value As Boolean)
    m_SinglePicker = Value
    PropertyChanged "SinglePicker"
    If m_SinglePicker Then NumberPickers = 1
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_UseRangeValue                 As Boolean
Public Property Get UseRangeValue() As Boolean
    UseRangeValue = m_UseRangeValue
End Property
Public Property Let UseRangeValue(ByVal Value As Boolean)
    m_UseRangeValue = Value
    If Not m_UseRangeValue Then
        ValueStart = vbNullString
        ValueEnd = vbNullString
    End If
    PropertyChanged "UseRangeValue"
    If m_IsChild Then Draw: Refresh
End Property

'm_ShowRangeButtons              As Boolean
Public Property Get ShowRangeButtons() As Boolean
    ShowRangeButtons = m_ShowRangeButtons
End Property

Public Property Let ShowRangeButtons(ByVal Value As Boolean)
    m_ShowRangeButtons = Value
    PropertyChanged "ShowRangeButtons"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_ShowTodayButton               As Boolean
Public Property Get ShowTodayButton() As Boolean
    ShowTodayButton = m_ShowTodayButton
End Property

Public Property Let ShowTodayButton(ByVal Value As Boolean)
    m_ShowTodayButton = Value
    PropertyChanged "ShowTodayButton"
    If m_IsChild Then InitControl: Draw: Refresh
End Property


'm_RightToLeft                   As Boolean
Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal Value As Boolean)
    m_RightToLeft = Value
    PropertyChanged "RightToLeft"
    If m_IsChild Then Draw: Refresh
End Property

'm_AutoApply                     As Boolean
Public Property Get AutoApply() As Boolean
    AutoApply = m_AutoApply
End Property

Public Property Let AutoApply(ByVal Value As Boolean)
    m_AutoApply = Value
'    If IsChild Then
'        MsgBox "Cannot enable AutoApply if control is child", vbInformation + vbOKOnly, UserControl.Name
'        m_AutoApply = False
'    End If
    ShowTodayButton = IIF(m_AutoApply, True, False)
    PropertyChanged "AutoApply"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_IsChild                       As Boolean
Public Property Get IsChild() As Boolean
    IsChild = m_IsChild
End Property

Public Property Let IsChild(ByVal Value As Boolean)
    m_IsChild = Value
    PropertyChanged "IsChild"
    ResetControl
    '--
    If m_IsChild Then
        If c_Width > 0 Then UserControl.Width = c_Width
        If c_Height > 0 Then UserControl.Height = c_Height
        Draw
    Else
        UserControl.Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
        Cls
    End If
    Extender.Visible = m_IsChild
End Property

'm_BackColorParent               As OLE_COLOR
Public Property Get BackColorParent() As OLE_COLOR
    BackColorParent = m_BackColorParent
End Property
Public Property Let BackColorParent(ByVal Value As OLE_COLOR)
    m_BackColorParent = Value
    PropertyChanged "BackColorParent"
    If m_IsChild Then Draw: Refresh
End Property

'm_ColsPicker                    As Integer
Public Property Get ColsPicker() As Integer
    ColsPicker = m_ColsPicker
End Property
Public Property Let ColsPicker(ByVal Value As Integer)
    m_ColsPicker = Value
    Call SafeRange(m_ColsPicker, 0, IIF(m_NumberPickers > 6, 6, m_NumberPickers))
    InitControl
    PropertyChanged "ColsPicker"
    If m_IsChild Then Draw: Refresh
End Property


'm_NumberPickers                 As Integer
Public Property Get NumberPickers() As Integer
    NumberPickers = m_NumberPickers
End Property
Public Property Let NumberPickers(ByVal Value As Integer)
    m_NumberPickers = Value
    Call SafeRange(m_NumberPickers, 1, 12)
    If m_NumberPickers <= m_ColsPicker Then
        m_ColsPicker = m_NumberPickers
        PropertyChanged "ColsPicker"
    End If
    InitControl
    PropertyChanged "NumberPickers"
    If m_IsChild Then Draw: Refresh
End Property

''m_ShowTimePicker                As Boolean 'Para trabajar el timepicker
'Public Property Get ShowTimePicker() As Boolean
'    ShowTimePicker = m_ShowTimePicker
'End Property
'Public Property Let ShowTimePicker(ByVal Value As Boolean)
'    m_ShowTimePicker = Value
'    PropertyChanged "ShowTimePicker"
'End Property
'
''m_UseTimePicker24Hrs            As Boolean 'Para formato de 24 horas en timepicker
'Public Property Get UseTimePicker24Hrs() As Boolean
'    UseTimePicker24Hrs = m_UseTimePicker24Hrs
'End Property
'Public Property Let UseTimePicker24Hrs(ByVal Value As Boolean)
'    m_UseTimePicker24Hrs = Value
'    PropertyChanged "UseTimePicker24Hrs"
'End Property
'
''m_TimerWithSecond               As Boolean
'Public Property Get TimerWithSecond() As Boolean
'    TimerWithSecond = m_TimerWithSecond
'End Property
'Public Property Let TimerWithSecond(ByVal Value As Boolean)
'    m_TimerWithSecond = Value
'    PropertyChanged "TimerWithSecond"
'End Property

'm_MaxRangeDays                  As Integer 'Para limitar el maximo de días en el selectionrange.
Public Property Get MaxRangeDays() As Integer
    MaxRangeDays = m_MaxRangeDays
End Property
Public Property Let MaxRangeDays(ByVal Value As Integer)
    m_MaxRangeDays = Value
    PropertyChanged "MaxRangeDays"
    If m_IsChild Then Draw: Refresh
End Property

'm_AlwaysShowCalendars           As Boolean
'Public Property Get AlwaysShowCalendars() As Boolean
'    AlwaysShowCalendars = m_AlwaysShowCalendars
'End Property
'Public Property Let AlwaysShowCalendars(ByVal Value As Boolean)
'    m_AlwaysShowCalendars = Value
'    PropertyChanged "AlwaysShowCalendars"
'    'Call CreateShadow
'    Draw
'    Refresh
'End Property

'm_linkedCalendars               As Boolean
Public Property Get LinkedCalendars() As Boolean
    LinkedCalendars = m_LinkedCalendars
End Property
Public Property Let LinkedCalendars(ByVal Value As Boolean)
    m_LinkedCalendars = Value
    PropertyChanged "LinkedCalendars"
    InitControl
    If m_IsChild Then Draw: Refresh
End Property

'm_Value                         As Date
Public Property Get Value() As Date
    Value = m_Value
End Property
Public Property Let Value(ByVal Value As Date)
    m_Value = Value
    InitControl
    PropertyChanged "Value"
    If m_IsChild Then Draw: Refresh
End Property

'm_ValueStart                    As String to Date
Public Property Get ValueStart() As String
    ValueStart = m_ValueStart
End Property
Public Property Let ValueStart(ByVal Value As String)
On Error GoTo PropertyError
'---
    If CDate(m_ValueEnd) Then
        If CDate(Value) > CDate(m_ValueEnd) Then
            Err.Raise Number:="5000", Description:="Invalid start date, cannot be greater than the end date."
        End If
    End If
    m_ValueStart = IIF(IsDate(Value), Value, "")
    PropertyChanged "ValueStart"
    If m_IsChild Then Draw: Refresh
    RaiseEvent ChangeStartDate(m_ValueStart)
    Exit Property
'---
PropertyError:
    MsgBox "Error Nro.: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbExclamation
End Property

'm_ValueEnd                      As String to Date
Public Property Get ValueEnd() As String
    ValueEnd = m_ValueEnd
End Property
Public Property Let ValueEnd(ByVal Value As String)
On Error GoTo PropertyError
'---
    If CDate(m_ValueStart) Then
        If CDate(Value) < CDate(m_ValueStart) Then
            Err.Raise Number:="5000", Description:="Invalid end date, it cannot be less than the start date."
        End If
    End If
    m_ValueEnd = IIF(IsDate(Value), Value, "")
    PropertyChanged "ValueStart"
    If m_IsChild Then Draw: Refresh
    RaiseEvent ChangeEndDate(m_ValueEnd)
    Exit Property
'---
PropertyError:
    MsgBox "Error Nro.: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbExclamation
End Property

'm_MinDate                       As Date
Public Property Get MinDate() As Date
    MinDate = m_MinDate
End Property
Public Property Let MinDate(ByVal Value As Date)
    If Not IsDate(Value) Then Exit Property
    If Value >= m_MaxDate Then GoTo ExitPrperty
    m_MinDate = Value
    PropertyChanged "MinDate"
    If m_IsChild Then Draw
ExitPrperty:
    If Ambient.UserMode Then RaiseEvent ChangeMinDate
End Property

'm_MaxDate                       As Date
Public Property Get MaxDate() As Date
    MaxDate = m_MaxDate
End Property
Public Property Let MaxDate(ByVal Value As Date)
    If Not IsDate(Value) Then Exit Property
    If Value <= m_MinDate Then GoTo ExitProperty
    m_MaxDate = Value
    PropertyChanged "MaxDate"
    If m_IsChild Then Draw
ExitProperty:
    If Ambient.UserMode Then RaiseEvent ChangeMaxDate
End Property

'm_FirstDayOfWeek                As VbDayOfWeek
Public Property Get FirstDayOfWeek() As VbDayOfWeek
    FirstDayOfWeek = m_FirstDayOfWeek
End Property
Public Property Let FirstDayOfWeek(ByVal Value As VbDayOfWeek)
    m_FirstDayOfWeek = Value
    PropertyChanged "FirstDayOfWeek"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_CountFreeDays                 As Boolean
Public Property Get CountFreeDays() As Boolean
    CountFreeDays = m_CountFreeDays
End Property
Public Property Let CountFreeDays(ByVal Value As Boolean)
    m_CountFreeDays = Value
    PropertyChanged "CountFreeDays"
    'Draw
    'Refresh
End Property

'm_CountReservedDay              As Boolean
Public Property Get CountReservedDay() As Boolean
    CountReservedDay = m_CountReservedDay
End Property
Public Property Let CountReservedDay(ByVal Value As Boolean)
    m_CountReservedDay = Value
    PropertyChanged "CountReservedDay"
    'Draw
    'Refresh
End Property

'm_CountSelDays                  As Integer
Public Property Get CountSelDays() As Boolean
    CountSelDays = m_CountSelDays
End Property

'---> De los botones.
'm_ButtonNavBackColor            As OLE_COLOR
Public Property Get ButtonNavBackColor() As OLE_COLOR
    ButtonNavBackColor = m_ButtonNavBackColor
End Property
Public Property Let ButtonNavBackColor(ByVal Value As OLE_COLOR)
    m_ButtonNavBackColor = Value
    PropertyChanged "ButtonNavBackColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonNavBorderWidth          As Integer
Public Property Get ButtonNavBorderWidth() As Integer
    ButtonNavBorderWidth = m_ButtonNavBorderWidth
End Property
Public Property Let ButtonNavBorderWidth(ByVal Value As Integer)
    m_ButtonNavBorderWidth = Value
    PropertyChanged "ButtonNavBorderWidth"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_ButtonNavBorderColor          As OLE_COLOR
Public Property Get ButtonNavBorderColor() As OLE_COLOR
    ButtonNavBorderColor = m_ButtonNavBorderColor
End Property
Public Property Let ButtonNavBorderColor(ByVal Value As OLE_COLOR)
    m_ButtonNavBorderColor = Value
    PropertyChanged "ButtonNavBorderColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonNavCornerRadius         As Integer
Public Property Get ButtonNavCornerRadius() As Integer
    ButtonNavCornerRadius = m_ButtonNavCornerRadius
End Property
Public Property Let ButtonNavCornerRadius(ByVal Value As Integer)
    m_ButtonNavCornerRadius = Value
    PropertyChanged "ButtonNavCornerRadius"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonNavForeColor            As OLE_COLOR
Public Property Get ButtonNavForeColor() As OLE_COLOR
    ButtonNavForeColor = m_ButtonNavForeColor
End Property
Public Property Let ButtonNavForeColor(ByVal Value As OLE_COLOR)
    m_ButtonNavForeColor = Value
    PropertyChanged "ButtonNavForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonNavIsIcoFont            As Boolean
Public Property Get ButtonNavIsIcoFont() As Boolean
    ButtonNavIsIcoFont = m_ButtonNavIsIcoFont
End Property
Public Property Let ButtonNavIsIcoFont(ByVal Value As Boolean)
    m_ButtonNavIsIcoFont = Value
    PropertyChanged "ButtonNavIsIcoFont"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonNavIcoFont              As StdFont
Public Property Get ButtonNavIcoFont() As StdFont
    Set ButtonNavIcoFont = m_ButtonNavIcoFont
End Property
Public Property Set ButtonNavIcoFont(Value As StdFont)
    With m_ButtonNavIcoFont
        .Name = Value.Name
        .Size = Value.Size
        .Bold = Value.Bold
        .Italic = Value.Italic
        .Underline = Value.Underline
        .Strikethrough = Value.Strikethrough
        .Weight = Value.Weight
        .Charset = Value.Charset
    End With
    PropertyChanged "ButtonNavIcoFont"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonNavCharCodeBack         As Long (String)
Public Property Get ButtonNavCharCodeBack() As String
    ButtonNavCharCodeBack = "&H" & Hex(m_ButtonNavCharCodeBack)
End Property
Public Property Let ButtonNavCharCodeBack(ByVal Value As String)
    Value = UCase(Replace(Value, Space(1), vbNullString))
    Value = UCase(Replace(Value, "U+", "&H"))
    If Not Left(Value, 2) = "&H" And Not IsNumeric(Value) Then
        m_ButtonNavCharCodeBack = "&H" & Value
    Else
        m_ButtonNavCharCodeBack = Value
    End If
    PropertyChanged "ButtonNavCharCodeBack"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonNavCharcodeNext         As Long
Public Property Get ButtonNavCharcodeNext() As String
    ButtonNavCharcodeNext = "&H" & Hex(m_ButtonNavCharCodeNext)
End Property
Public Property Let ButtonNavCharcodeNext(ByVal Value As String)
    Value = UCase(Replace(Value, Space(1), vbNullString))
    Value = UCase(Replace(Value, "U+", "&H"))
    If Not Left(Value, 2) = "&H" And Not IsNumeric(Value) Then
        m_ButtonNavCharCodeNext = "&H" & Value
    Else
        m_ButtonNavCharCodeNext = Value
    End If
    PropertyChanged "ButtonNavCharcodeNext"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonNavWidth                As Long
Public Property Get ButtonNavWidth() As Long
    ButtonNavWidth = m_ButtonNavWidth
End Property
Public Property Let ButtonNavWidth(ByVal Value As Long)
    m_ButtonNavWidth = Value
    InitControl
    PropertyChanged "ButtonNavWidth"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_ButtonNavHeight               As Long
Public Property Get ButtonNavHeight() As Long
    ButtonNavHeight = m_ButtonNavHeight
End Property
Public Property Let ButtonNavHeight(ByVal Value As Long)
    m_ButtonNavHeight = Value
    InitControl
    PropertyChanged "ButtonNavHeight"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'---> De los botones de accion.
'm_ButtonsBackColor              As OLE_COLOR
Public Property Get ButtonsBackColor() As OLE_COLOR
    ButtonsBackColor = m_ButtonsBackColor
End Property
Public Property Let ButtonsBackColor(ByVal Value As OLE_COLOR)
    m_ButtonsBackColor = Value
    InitControl
    PropertyChanged "ButtonsBackColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonsBorderWidth            As Integer
Public Property Get ButtonsBorderWidth() As Integer
    ButtonsBorderWidth = m_ButtonsBorderWidth
End Property
Public Property Let ButtonsBorderWidth(ByVal Value As Integer)
    m_ButtonsBorderWidth = Value
    InitControl
    PropertyChanged "ButtonsBorderWidth"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonsBorderColor            As OLE_COLOR
Public Property Get ButtonsBorderColor() As OLE_COLOR
    ButtonsBorderColor = m_ButtonsBorderColor
End Property
Public Property Let ButtonsBorderColor(ByVal Value As OLE_COLOR)
    m_ButtonsBorderColor = Value
    InitControl
    PropertyChanged "ButtonsBorderColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonsCornerRadius           As Integer
Public Property Get ButtonsCornerRadius() As Integer
    ButtonsCornerRadius = m_ButtonsCornerRadius
End Property
Public Property Let ButtonsCornerRadius(ByVal Value As Integer)
    m_ButtonsCornerRadius = Value
    InitControl
    PropertyChanged "ButtonsCornerRadius"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonsFont                   As StdFont
Public Property Get ButtonsFont() As StdFont
    Set ButtonsFont = m_ButtonsFont
End Property
Public Property Set ButtonsFont(Value As StdFont)
    With m_ButtonsFont
        .Name = Value.Name
        .Size = Value.Size
        .Bold = Value.Bold
        .Italic = Value.Italic
        .Underline = Value.Underline
        .Strikethrough = Value.Strikethrough
        .Weight = Value.Weight
        .Charset = Value.Charset
    End With
    PropertyChanged "ButtonsFont"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonsForeColor              As OLE_COLOR
Public Property Get ButtonsForeColor() As OLE_COLOR
    ButtonsForeColor = m_ButtonsForeColor
End Property
Public Property Let ButtonsForeColor(ByVal Value As OLE_COLOR)
    m_ButtonsForeColor = Value
    InitControl
    PropertyChanged "ButtonsForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_ButtonsWidth                  As Long
Public Property Get ButtonsWidth() As Long
    ButtonsWidth = m_ButtonsWidth
End Property
Public Property Let ButtonsWidth(ByVal Value As Long)
    m_ButtonsWidth = Value
    InitControl
    PropertyChanged "ButtonsWidth"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_ButtonsHeight                 As Long
Public Property Get ButtonsHeight() As Long
    ButtonsHeight = m_ButtonsHeight
End Property
Public Property Let ButtonsHeight(ByVal Value As Long)
    m_ButtonsHeight = Value
    InitControl
    PropertyChanged "ButtonsHeight"
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'---> De los meses y años
'm_MonthYearBackColor            As OLE_COLOR
Public Property Get MonthYearBackColor() As OLE_COLOR
    MonthYearBackColor = m_MonthYearBackColor
End Property
Public Property Let MonthYearBackColor(ByVal Value As OLE_COLOR)
    m_MonthYearBackColor = Value
    PropertyChanged "MonthYearBackColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_MonthYearBorderWidth          As Integer
Public Property Get MonthYearBorderWidth() As Integer
    MonthYearBorderWidth = m_MonthYearBorderWidth
End Property
Public Property Let MonthYearBorderWidth(ByVal Value As Integer)
    m_MonthYearBorderWidth = Value
    PropertyChanged "MonthYearBorderWidth"
    If m_IsChild Then Draw: Refresh
End Property

'm_MonthYearBorderColor          As OLE_COLOR
Public Property Get MonthYearBorderColor() As OLE_COLOR
    MonthYearBorderColor = m_MonthYearBorderColor
End Property
Public Property Let MonthYearBorderColor(ByVal Value As OLE_COLOR)
    m_MonthYearBorderColor = Value
    PropertyChanged "MonthYearBorderColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_MonthYearCornerRadius         As Integer
Public Property Get MonthYearCornerRadius() As Integer
    MonthYearCornerRadius = m_MonthYearCornerRadius
End Property
Public Property Let MonthYearCornerRadius(ByVal Value As Integer)
    m_MonthYearCornerRadius = Value
    PropertyChanged "MonthYearCornerRadius"
    If m_IsChild Then Draw: Refresh
End Property

'm_MonthYearFont                 As StdFont
Public Property Get MonthYearFont() As StdFont
    Set MonthYearFont = m_MonthYearFont
End Property
Public Property Set MonthYearFont(Value As StdFont)
    With m_MonthYearFont
        .Name = Value.Name
        .Size = Value.Size
        .Bold = Value.Bold
        .Italic = Value.Italic
        .Underline = Value.Underline
        .Strikethrough = Value.Strikethrough
        .Weight = Value.Weight
        .Charset = Value.Charset
    End With
    PropertyChanged "MonthYearFont"
    If m_IsChild Then Draw: Refresh
End Property

'm_MonthYearForeColor            As OLE_COLOR
Public Property Get MonthYearForeColor() As OLE_COLOR
    MonthYearForeColor = m_MonthYearForeColor
End Property
Public Property Let MonthYearForeColor(ByVal Value As OLE_COLOR)
    m_MonthYearForeColor = Value
    PropertyChanged "MonthYearForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'---> De las Semanas.
'm_WeekBackColor                 As OLE_COLOR
Public Property Get WeekBackColor() As OLE_COLOR
    WeekBackColor = m_WeekBackColor
End Property
Public Property Let WeekBackColor(ByVal Value As OLE_COLOR)
    m_WeekBackColor = Value
    PropertyChanged "WeekBackColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_WeekBorderWidth               As Integer
Public Property Get WeekBorderWidth() As Integer
    WeekBorderWidth = m_WeekBorderWidth
End Property
Public Property Let WeekBorderWidth(ByVal Value As Integer)
    m_WeekBorderWidth = Value
    PropertyChanged "WeekBorderWidth"
    If m_IsChild Then Draw: Refresh
End Property

'm_WeekBorderColor               As OLE_COLOR
Public Property Get WeekBorderColor() As OLE_COLOR
    WeekBorderColor = m_WeekBorderColor
End Property
Public Property Let WeekBorderColor(ByVal Value As OLE_COLOR)
    m_WeekBorderColor = Value
    PropertyChanged "WeekBorderColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_WeekCornerRadius              As Integer
Public Property Get WeekCornerRadius() As Integer
    WeekCornerRadius = m_WeekCornerRadius
End Property
Public Property Let WeekCornerRadius(ByVal Value As Integer)
    m_WeekCornerRadius = Value
    PropertyChanged "WeekCornerRadius"
    If m_IsChild Then Draw: Refresh
End Property

'm_WeekFont                      As StdFont
Public Property Get WeekFont() As StdFont
    Set WeekFont = m_WeekFont
End Property
Public Property Set WeekFont(Value As StdFont)
    With m_WeekFont
        .Name = Value.Name
        .Size = Value.Size
        .Bold = Value.Bold
        .Italic = Value.Italic
        .Underline = Value.Underline
        .Strikethrough = Value.Strikethrough
        .Weight = Value.Weight
        .Charset = Value.Charset
    End With
    PropertyChanged "WeekFont"
    If m_IsChild Then Draw: Refresh
End Property

'm_WeekFontHeaderBold            As Boolean
Public Property Get WeekFontHeaderBold() As Boolean
    WeekFontHeaderBold = m_WeekFontHeaderBold
End Property
Public Property Let WeekFontHeaderBold(ByVal Value As Boolean)
    m_WeekFontHeaderBold = Value
    PropertyChanged "WeekFontHeaderBold"
    If m_IsChild Then Draw: Refresh
End Property

'm_WeekForeColor                 As OLE_COLOR
Public Property Get WeekForeColor() As OLE_COLOR
    WeekForeColor = m_WeekForeColor
End Property
Public Property Let WeekForeColor(ByVal Value As OLE_COLOR)
    m_WeekForeColor = Value
    PropertyChanged "WeekForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_WeekWidth                     As Long
Public Property Get WeekWidth() As Long
    WeekWidth = m_WeekWidth
End Property
Public Property Let WeekWidth(ByVal Value As Long)
    m_WeekWidth = Value
    PropertyChanged "WeekWidth"
    InitControl
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_WeekHeight                    As Long
Public Property Get WeekHeight() As Long
    WeekHeight = m_WeekHeight
End Property
Public Property Let WeekHeight(ByVal Value As Long)
    m_WeekHeight = Value
    PropertyChanged "WeekHeight"
    '---
    m_DayHeight = m_WeekHeight
    PropertyChanged "DayHeight"
    '---
    InitControl
    If m_IsChild Then InitControl: Draw: Refresh
End Property

''---> De los días.
'm_DayBackColor                  As OLE_COLOR
Public Property Get DayBackColor() As OLE_COLOR
    DayBackColor = m_DayBackColor
End Property
Public Property Let DayBackColor(ByVal Value As OLE_COLOR)
    m_DayBackColor = Value
    PropertyChanged "DayBackColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayBorderWidth                As Integer
Public Property Get DayBorderWidth() As Integer
    DayBorderWidth = m_DayBorderWidth
End Property
Public Property Let DayBorderWidth(ByVal Value As Integer)
    m_DayBorderWidth = Value
    PropertyChanged "DayBorderWidth"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayBorderColor                As OLE_COLOR
Public Property Get DayBorderColor() As OLE_COLOR
    DayBorderColor = m_DayBorderColor
End Property
Public Property Let DayBorderColor(ByVal Value As OLE_COLOR)
    m_DayBorderColor = Value
    PropertyChanged "DayBorderColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayCornerRadius               As Integer
Public Property Get DayCornerRadius() As Integer
    DayCornerRadius = m_DayCornerRadius
End Property
Public Property Let DayCornerRadius(ByVal Value As Integer)
    m_DayCornerRadius = Value
    PropertyChanged "DayCornerRadius"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayFont                       As StdFont
Public Property Get DayFont() As StdFont
    Set DayFont = m_DayFont
End Property
Public Property Set DayFont(Value As StdFont)
    With m_DayFont
        .Name = Value.Name
        .Size = Value.Size
        .Bold = Value.Bold
        .Italic = Value.Italic
        .Underline = Value.Underline
        .Strikethrough = Value.Strikethrough
        .Weight = Value.Weight
        .Charset = Value.Charset
    End With
    PropertyChanged "DayFont"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayHeaderFontBold             As Boolean
Public Property Get DayHeaderFontBold() As Boolean
    DayHeaderFontBold = m_DayHeaderFontBold
End Property
Public Property Let DayHeaderFontBold(ByVal Value As Boolean)
    m_DayHeaderFontBold = Value
    PropertyChanged "DayHeaderFontBold"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayHeaderForeColor            As OLE_COLOR
Public Property Get DayHeaderForeColor() As OLE_COLOR
    DayHeaderForeColor = m_DayHeaderForeColor
End Property
Public Property Let DayHeaderForeColor(ByVal Value As OLE_COLOR)
    m_DayHeaderForeColor = Value
    PropertyChanged "DayHeaderForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayHotColor                   As OLE_COLOR
Public Property Get DayHotColor() As OLE_COLOR
    DayHotColor = m_DayHotColor
End Property
Public Property Let DayHotColor(ByVal Value As OLE_COLOR)
    m_DayHotColor = Value
    PropertyChanged "DayHotColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayForeColor                  As OLE_COLOR
Public Property Get DayForeColor() As OLE_COLOR
    DayForeColor = m_DayForeColor
End Property
Public Property Let DayForeColor(ByVal Value As OLE_COLOR)
    m_DayForeColor = Value
    PropertyChanged "DayForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayOMForeColor                As OLE_COLOR
Public Property Get DayOMForeColor() As OLE_COLOR
    DayOMForeColor = m_DayOMForeColor
End Property
Public Property Let DayOMForeColor(ByVal Value As OLE_COLOR)
    m_DayOMForeColor = Value
    PropertyChanged "DayOMForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayFreeArray                  As Variant

'm_DayFreeForeColor              As OLE_COLOR
Public Property Get DayFreeForeColor() As OLE_COLOR
    DayFreeForeColor = m_DayFreeForeColor
End Property
Public Property Let DayFreeForeColor(ByVal Value As OLE_COLOR)
    m_DayFreeForeColor = Value
    PropertyChanged "DayFreeForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayNowShow                    As Boolean
Public Property Get DayNowShow() As Boolean
    DayNowShow = m_DayNowShow
End Property
Public Property Let DayNowShow(ByVal Value As Boolean)
    m_DayNowShow = Value
    PropertyChanged "DayNowShow"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayNowBorderWidth             As Integer
Public Property Get DayNowBorderWidth() As Integer
    DayNowBorderWidth = m_DayNowBorderWidth
End Property
Public Property Let DayNowBorderWidth(ByVal Value As Integer)
    m_DayNowBorderWidth = Value
    PropertyChanged "DayNowBorderWidth"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayNowBorderColor             As OLE_COLOR
Public Property Get DayNowBorderColor() As OLE_COLOR
    DayNowBorderColor = m_DayNowBorderColor
End Property
Public Property Let DayNowBorderColor(ByVal Value As OLE_COLOR)
    m_DayNowBorderColor = Value
    PropertyChanged "DayNowBorderColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayNowBackColor               As OLE_COLOR
Public Property Get DayNowBackColor() As OLE_COLOR
    DayNowBackColor = m_DayNowBackColor
End Property
Public Property Let DayNowBackColor(ByVal Value As OLE_COLOR)
    m_DayNowBackColor = Value
    PropertyChanged "DayNowBackColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayNowForeColor               As OLE_COLOR
Public Property Get DayNowForeColor() As OLE_COLOR
    DayNowForeColor = m_DayNowForeColor
End Property
Public Property Let DayNowForeColor(ByVal Value As OLE_COLOR)
    m_DayNowForeColor = Value
    PropertyChanged "DayNowForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DaysPrePaintCount             As Long
Public Property Get DaysPrePaintCount() As Long
    DaysPrePaintCount = m_DaysPrePaintCount
End Property

'm_DaySelCount                   As Long
Public Property Get DaySelCount() As Long
    If IsDate(m_ValueStart) And IsDate(m_ValueEnd) Then
        m_DaySelCount = DateDiff("d", m_ValueStart, m_ValueEnd) + 1
    Else
        m_DaySelCount = 0
    End If
    DaySelCount = m_DaySelCount
End Property

'Mouse event (Los días no tendran color para el mousedown)
'm_DayOverBackColor              As OLE_COLOR
Public Property Get DayOverBackColor() As OLE_COLOR
    DayOverBackColor = m_DayOverBackColor
End Property
Public Property Let DayOverBackColor(ByVal Value As OLE_COLOR)
    m_DayOverBackColor = Value
    PropertyChanged "DayOverBackColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayOverForeColor              As OLE_COLOR
Public Property Get DayOverForeColor() As OLE_COLOR
    DayOverForeColor = m_DayOverForeColor
End Property
Public Property Let DayOverForeColor(ByVal Value As OLE_COLOR)
    m_DayOverForeColor = Value
    PropertyChanged "DayOverForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'--
'm_DaySelBetweenColor            As OLE_COLOR
Public Property Get DaySelBetweenColor() As OLE_COLOR
    DaySelBetweenColor = m_DaySelBetweenColor
End Property
Public Property Let DaySelBetweenColor(ByVal Value As OLE_COLOR)
    m_DaySelBetweenColor = Value
    PropertyChanged "DaySelBetweenColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DaySelEndColor                As OLE_COLOR
'Public Property Get DaySelEndColor() As OLE_COLOR
'    DaySelEndColor = m_DaySelEndColor
'End Property
'Public Property Let DaySelEndColor(ByVal Value As OLE_COLOR)
'    m_DaySelEndColor = Value
'    PropertyChanged "DaySelEndColor"
'    Draw
'    Refresh
'End Property

'm_DaySelValuesColor              As OLE_COLOR
Public Property Get DaySelValuesColor() As OLE_COLOR
    DaySelValuesColor = m_DaySelValuesColor
End Property
Public Property Let DaySelValuesColor(ByVal Value As OLE_COLOR)
    m_DaySelValuesColor = Value
    PropertyChanged "DaySelValuesColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DaySelForeColor               As OLE_COLOR
Public Property Get DaySelForeColor() As OLE_COLOR
    DaySelForeColor = m_DaySelForeColor
End Property
Public Property Let DaySelForeColor(ByVal Value As OLE_COLOR)
    m_DaySelForeColor = Value
    PropertyChanged "DaySelForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DaySelFontBold                As Boolean
Public Property Get DaySelFontBold() As Boolean
    DaySelFontBold = m_DaySelFontBold
End Property
Public Property Let DaySelFontBold(ByVal Value As Boolean)
    m_DaySelFontBold = Value
    PropertyChanged "DaySelFontBold"
    If m_IsChild Then Draw: Refresh
End Property

'm_DaySelectionStyle             As enmSelectionStyle
Public Property Get DaySelectionStyle() As enmSelectionStyle
    DaySelectionStyle = m_DaySelectionStyle
End Property
Public Property Let DaySelectionStyle(ByVal Value As enmSelectionStyle)
    m_DaySelectionStyle = Value
    PropertyChanged "DaySelectionStyle"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayShowHotItem                As Boolean 'Siempre muestra el hot
'Public Property Get DayShowHotItem() As Boolean
'    DayShowHotItem = m_DayShowHotItem
'End Property
'Public Property Let DayShowHotItem(ByVal Value As Boolean)
'    m_DayShowHotItem = Value
'    PropertyChanged "DayShowHotItem"
'End Property


'm_DaySaturdayForeColor          As OLE_COLOR
Public Property Get DaySaturdayForeColor() As OLE_COLOR
    DaySaturdayForeColor = m_DaySaturdayForeColor
End Property
Public Property Let DaySaturdayForeColor(ByVal Value As OLE_COLOR)
    m_DaySaturdayForeColor = Value
    PropertyChanged "DaySaturdayForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DaySundayForeColor            As OLE_COLOR
Public Property Get DaySundayForeColor() As OLE_COLOR
    DaySundayForeColor = m_DaySundayForeColor
End Property
Public Property Let DaySundayForeColor(ByVal Value As OLE_COLOR)
    m_DaySundayForeColor = Value
    PropertyChanged "DaySundayForeColor"
    If m_IsChild Then Draw: Refresh
End Property

'm_DayWidth                      As Long
Public Property Get DayWidth() As Long
    DayWidth = m_DayWidth
End Property
Public Property Let DayWidth(ByVal Value As Long)
    m_DayWidth = Value
    PropertyChanged "DayWidth"
    InitControl
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_DayHeight                     As Long
Public Property Get DayHeight() As Long
    DayHeight = m_DayHeight
End Property
Public Property Let DayHeight(ByVal Value As Long)
    m_DayHeight = Value
    PropertyChanged "DayHeight"
    '---
    m_WeekHeight = m_DayHeight
    PropertyChanged "WeekHeight"
    '---
    InitControl
    If m_IsChild Then InitControl: Draw: Refresh
End Property

'm_CallOut                       as long
Public Property Get CallOut() As Boolean
    CallOut = m_CallOut
End Property
Public Property Let CallOut(Value As Boolean)
    m_CallOut = Value
    c_CShadow.CallOut = m_CallOut
    PropertyChanged "CallOut"
    'Draw
    'Refresh
End Property

'm_CallOutWidth                      As Long
Public Property Get CallOutWidth() As Long
    CallOutWidth = m_CallOutWidth
End Property
Public Property Let CallOutWidth(Value As Long)
    m_CallOutWidth = Value
    c_CShadow.CallOutWidth = m_CallOutWidth
    PropertyChanged "CallOutWidth"
    'Draw
    'Refresh
End Property

'm_CallOutHight                      As Long
Public Property Get CallOutHight() As Long
    CallOutHight = m_CallOutHight
End Property
Public Property Let CallOutHight(Value As Long)
    m_CallOutHight = Value
    c_CShadow.CallOutHight = m_CallOutHight
    PropertyChanged "CallOutHight"
    'Draw
    'Refresh
End Property

'm_CallOutRightTriangle              As Boolean
Public Property Get CallOutRightTriangle() As Boolean
    CallOutRightTriangle = m_CallOutRightTriangle
End Property
Public Property Let CallOutRightTriangle(Value As Boolean)
    m_CallOutRightTriangle = Value
    c_CShadow.CallOutRightTriangle = m_CallOutRightTriangle
    PropertyChanged "CallOutRightTriangle"
    'Draw
    'Refresh
End Property

'm_CallOutPosition                   As enmCallOutPosition
Public Property Get CallOutPosition() As enmCallOutPosition
    CallOutPosition = m_CallOutPosition
End Property
Public Property Let CallOutPosition(Value As enmCallOutPosition)
    m_CallOutPosition = Value
End Property

'm_CallOutCustomPos                   As Long
Public Property Get CallOutCustomPosPercent() As Long
    CallOutCustomPosPercent = m_CallOutCustomPosPercent
End Property
Public Property Let CallOutCustomPosPercent(Value As Long)
    m_CallOutCustomPosPercent = Value
    Call SafeRange(m_CallOutCustomPosPercent, 0, 100)
End Property

'm_CallOutAlign                      As enmCallOutAlign
Public Property Get CallOutAlign() As enmCallOutAlign
    CallOutAlign = m_CallOutAlign
End Property
Public Property Let CallOutAlign(Value As enmCallOutAlign)
    m_CallOutAlign = Value
End Property
'---

Private Sub tmrMouseMoveDays_Timer()
    
End Sub

Private Sub tmrMouseMove_Timer()
End Sub

Private Sub tmrMouseEvent_Timer()
    Dim hwin As Long
    Dim pt As POINTAPI
    Dim Left As Long, Top As Long
    Dim cRect As RECT
    '--
    If tmrMouseEvent.Interval = 2 Then
        Call Draw
        tmrMouseEvent.Interval = 1
    End If
    '--
    If Not m_IsChild Then
        If (GetAsyncKeyState(vbLeftButton) < 0) Or (GetAsyncKeyState(vbRightButton) < 0) Then
            GetCursorPos pt
            hwin = WindowFromPoint(pt.X, pt.Y)
           
            If hwin <> UserControl.hWnd Then
                'If GetCapture <> UserControl.hWnd Then HideList
                HideCalendar
            End If
        End If
    Else 'IsChild
        GetCursorPos pt
        ScreenToClient c_hWnd, pt
        '--
        Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
        Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)
        '--
        With cRect
            .Left = m_PT.X - (m_Left - Left)
            .Top = m_PT.Y - (m_Top - Top)
            .Right = .Left + UserControl.ScaleWidth
            .Bottom = .Top + UserControl.ScaleHeight
        End With
        '--
        SendMessage c_hWnd, WM_MOUSEMOVE, 0&, ByVal pt.X Or pt.Y * &H10000
        '--
        If m_Over Then
            m_Over = False
            RaiseEvent MouseOut
        End If
        '--
        If PtInRect(cRect, pt.X, pt.Y) = 0 Then
            m_Enter = False
            tmrMouseEvent.Interval = 0
            ResetControl
            Call Draw ': Refresh
            RaiseEvent MouseLeave
        End If
    End If
    '--
End Sub

Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    '---
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    nScale = GetWindowsDPI
    '---
    Set c_SubClass = New clsSubClass
    Set c_CShadow = New clsShadow
    '---
    iCalendar = -1
    '---
    udtItemsRangeButton(0).Caption = "Today"                'Hoy
    udtItemsRangeButton(1).Caption = "Current month"        'Este mes
    udtItemsRangeButton(2).Caption = "Last month"           'Mes pasado
    udtItemsRangeButton(3).Caption = "Last 90 days"         'Ultimos 90 días
    udtItemsRangeButton(4).Caption = "Current year"         'Este Año
    udtItemsRangeButton(5).Caption = "Last year"            'Año pasado
    '---
    udtItemsActionButton(0).Caption = "Today"               'Hoy
    udtItemsActionButton(0).ButtonAction = [Action Today]
    udtItemsActionButton(0).IsVisible = True
    
    If Not IsChild Then
        udtItemsActionButton(1).Caption = "Cancel"          'Cancelar
    Else
        udtItemsActionButton(1).Caption = "None"            'Cancelar
    End If
    udtItemsActionButton(1).ButtonAction = [Action Cancel]
    udtItemsActionButton(0).IsVisible = True
    
    udtItemsActionButton(2).Caption = "Apply"               'Aplicar
    udtItemsActionButton(2).ButtonAction = [Action Apply]
    If IsChild Then udtItemsActionButton(2).IsVisible = False
    '---
    m_Redraw = True
    '---
End Sub

Private Sub UserControl_InitProperties()
    'hFontCollection = ReadValue(&HFC)
    '---
    m_BackColor = &HFFFFFF
    m_BackOpacity = 100
    m_Border = True
    m_BorderWidth = 1
    m_BorderColor = SystemColorConstants.vbActiveBorder
    m_BorderOpacity = 100
    m_BorderPosition = [Border Center]
    m_PaddingX = 20
    m_PaddingY = 20
    m_ShowNumberWeek = True
    m_ShowUseISOWeek = True
    m_SpaceGrid = 2
    '--
    m_SinglePicker = False
    m_UseRangeValue = False
    m_ShowRangeButtons = False                             'Para mostrar los rangos predefinidos por el usuario(Today, Yesterday, Last 7 days, Last 30 days, This Month, Last Month, Custom Range)
    m_ShowTodayButton = False
    m_RightToLeft = False
    m_AutoApply = True
    m_IsChild = False
    
    m_BackColorParent = Ambient.BackColor
    m_ColsPicker = 0
    m_NumberPickers = IIF(Not m_SinglePicker, 2, 1)
    
    'TimerPicker
    'm_ShowTimePicker = False                        'Para trabajar el timepicker
    'm_UseTimePicker24Hrs = False                    'Para formato de 24 horas en timepicker
    'm_TimerWithSecond = False
    'TimerPicker
    
    m_MaxRangeDays = 0
    
    'm_AlwaysShowCalendars = True
    m_LinkedCalendars = True
    
    m_Value = Date
    m_ValueStart = ""
    m_ValueEnd = ""
    
    m_MinDate = DateSerial(1601, 1, 1)
    m_MaxDate = DateSerial(9999, 12, 31)
    
    m_FirstDayOfWeek = vbUseSystemDayOfWeek

    m_CountFreeDays = True
    m_CountReservedDay = True

    '---> De los botones de navegacion
    m_ButtonNavBackColor = &HFFFFFF
    m_ButtonNavBorderWidth = 1
    m_ButtonNavBorderColor = SystemColorConstants.vbActiveBorder
    m_ButtonNavCornerRadius = 12
    m_ButtonNavForeColor = SystemColorConstants.vbButtonText
    Set m_ButtonNavIcoFont = UserControl.Ambient.Font
    m_ButtonNavWidth = 24
    m_ButtonNavHeight = 24
    
    '---> De los botones de accion
    m_ButtonsBackColor = &HFFFFFF
    m_ButtonsBorderWidth = 1
    m_ButtonsBorderColor = SystemColorConstants.vbActiveBorder
    m_ButtonsCornerRadius = 5
    Set m_ButtonsFont = UserControl.Ambient.Font
    m_ButtonsForeColor = SystemColorConstants.vbButtonText
    m_ButtonsWidth = 24
    m_ButtonsHeight = 24
    
    '---> De los meses y años
    m_MonthYearBackColor = &HFFFFFF
    Set m_MonthYearFont = UserControl.Ambient.Font
    m_MonthYearForeColor = SystemColorConstants.vbButtonText
    '---> De las Semanas.
    m_WeekBackColor = &HFFFFFF
    Set m_WeekFont = UserControl.Ambient.Font
    m_WeekFontHeaderBold = True
    m_WeekForeColor = SystemColorConstants.vbGrayText
    m_WeekWidth = 26
    m_WeekHeight = 24

    '---> De los días.
    m_DayBackColor = &HFFFFFF
    Set m_DayFont = UserControl.Ambient.Font
    m_DayHeaderFontBold = True
    m_DayHeaderForeColor = SystemColorConstants.vbButtonText
    m_DayHotColor = SystemColorConstants.vbHighlight
    m_DayForeColor = SystemColorConstants.vbButtonText
    m_DayOMForeColor = SystemColorConstants.vbGrayText
    m_DayFreeForeColor = ColorConstants.vbRed
    '--
    m_DayNowShow = False
    m_DayNowBorderWidth = 0
    m_DayNowBorderColor = SystemColorConstants.vbActiveBorder
    m_DayNowBackColor = &HFFFFFF
    m_DayNowForeColor = SystemColorConstants.vbButtonText
    '--
    'Mouse event (Los días no tendran color para el mousedown)
    m_DayOverBackColor = &H999999
    m_DayOverForeColor = &HFFFFFF
    '--
    m_DaySelBetweenColor = &HE5D7CA
    'm_DaySelEndColor = &HB06D00
    m_DaySelValuesColor = &HB06D00
    m_DaySelForeColor = &HFFFFFF
    m_DaySelFontBold = True
    m_DaySelectionStyle = [Corner No Between]
    'm_DayShowHotItem = True 'Siempre muestra el hot
    m_DaySaturdayForeColor = &H999999
    m_DaySundayForeColor = &H999999
    m_DayWidth = 24
    m_DayHeight = 24
    '---
    m_CallOut = True
    m_CallOutWidth = 20
    m_CallOutHight = 10
    m_CallOutRightTriangle = False
    m_CallOutPosition = [Position Top]
    m_CallOutAlign = Middle
    m_CallOutCustomPosPercent = 0
    '---
    c_hWnd = UserControl.ContainerHwnd
    '---
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i, a As Integer
    Dim bTemp As Boolean
    '-> Botones del o los calendario(s)
    For i = 0 To UBound(udtItemsNavButton)
        With udtItemsNavButton(i)
            If PtInRect(.RECT2, X, Y) Then
                If .MouseState = Hot Then
                    .MouseState = Pressed
                    Call Draw ': Refresh
                    Exit For
                End If
            End If
        End With
    Next
    '---
    
    '-> Mes y año
    For i = 0 To UBound(udtItemsPicker)
        With udtItemsPicker(i)
            If PtInRect(.TitleMonthYear.RECT2, X, Y) Then
                If .TitleMonthYear.MouseState = Hot Then .TitleMonthYear.MouseState = Pressed
                '--
                Call Draw ': Refresh
            End If
            '--
            If .ViewNavigator <> ViewItemNavigatorDays And m_LinkedCalendars Then
                For a = 0 To UBound(udtItemsUpDownButton)
                    With udtItemsUpDownButton(a)
                        If PtInRect(.RECT2, X, Y) Then
                            .MouseState = Pressed
                            '--
                            Call Draw ': Refresh
                        End If
                    End With
                Next
            End If
        End With
    Next
    '---

    '-> Días de los meses.
    For i = 0 To UBound(udtItemsDay)
        With udtItemsDay(i)
            If PtInRect(.RECT2, X, Y) Then
                .MouseState = Pressed
            End If
        End With
    Next

    '-> Meses o Años de la navegacion rapida.
    For i = 0 To UBound(udtItemsMonthYear)
        With udtItemsMonthYear(i)
            If udtItemsPicker(.IndexCalendar).ViewNavigator <> ViewItemNavigatorDays Then
                If PtInRect(.RECT2, X, Y) Then
                    If udtItemsPicker(.IndexCalendar).ViewNavigator = ViewItemNavigatorMonths Then
                        bTemp = DateSerial(udtItemsPicker(.IndexCalendar).NumberYear, .ValueItem, 1) < m_MinDate Or DateSerial(udtItemsPicker(.IndexCalendar).NumberYear, .ValueItem, 1) > m_MaxDate
                    Else
                        'bTemp = .ValueItem < Year(m_MinDate) Or .ValueItem > Year(m_MinDate)
                    End If
                    If Not bTemp Then
                        .MouseState = Pressed
                        Call Draw ': Refresh
                    End If
                End If
            End If
        End With
    Next
    '---

    '-> Botones de los rangos de fecha.
    If m_ShowRangeButtons And m_UseRangeValue Then
        For i = 0 To UBound(udtItemsRangeButton)
            With udtItemsRangeButton(i)
                If PtInRect(.RECT2, X, Y) Then
                    .MouseState = Pressed
                    '--
                    Call Draw ': Refresh
                    '--
                End If
            End With
        Next
    End If
    '---

    '-> Botones de accion.
    If Not m_AutoApply Or m_ShowTodayButton Then
        For i = 0 To UBound(udtItemsActionButton)
            With udtItemsActionButton(i)
                If PtInRect(.RECT2, X, Y) Then
                    .MouseState = Pressed
                    '--
                    Call Draw ': Refresh
                    '--
                End If
            End With
        Next
    End If
    '---
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i, a As Integer
    Dim ET As TRACKMOUSEEVENTTYPE
    Dim newRect As RECT
    Dim iTmpValue As Integer
    Dim pt As POINTAPI
    Dim lHwnd As Long
    Dim bTemp As Boolean
    '---
    ET.cbSize = Len(ET)
    ET.hwndTrack = UserControl.hWnd
    ET.dwFlags = TME_LEAVE
    TrackMouseEvent ET
    
    If m_IsChild Then
        GetCursorPos pt
        lHwnd = WindowFromPoint(pt.X, pt.Y)
        '--
        If Not m_Enter Then
            ScreenToClient c_hWnd, pt
            m_PT.X = pt.X - X
            m_PT.Y = pt.Y - Y
            '--
            m_Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
            m_Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)
            '--
            m_Enter = True
            tmrMouseEvent.Interval = 1
            RaiseEvent MouseEnter
        End If
        '--
        If lHwnd = c_hWnd Then
            If Not m_Over Then
                m_Over = True
                RaiseEvent MouseOver
            End If
        Else
            If m_Over Then
                m_Over = False
                RaiseEvent MouseOut
            End If
        End If
    End If
    '---
    
    '-> Botones de navegacion del calendario.
    For i = 0 To UBound(udtItemsNavButton)
        With udtItemsNavButton(i)
            If PtInRect(.RECT2, X, Y) Then
                If .MouseState = Normal Then
                    .MouseState = IIF(Button = vbLeftButton, Pressed, Hot)
                    ShowHandPointer True
                    tmrMouseEvent.Interval = 2
                    Exit Sub
                End If
            Else
                If .MouseState <> Normal Then
                    .MouseState = Normal
                    ShowHandPointer False
                    tmrMouseEvent.Interval = 2
                    Exit Sub
                End If
            End If
        End With
    Next
    '---
    
    '-> Mes y años
    For i = 0 To UBound(udtItemsPicker)
        With udtItemsPicker(i)
            If PtInRect(.TitleMonthYear.RECT2, X, Y) Then
                If .TitleMonthYear.MouseState = Normal Then
                    .TitleMonthYear.MouseState = IIF(Button = vbLeftButton, Pressed, Hot)
                    ShowHandPointer True
                    tmrMouseEvent.Interval = 2
                    Exit Sub
                End If
            Else
                If .TitleMonthYear.MouseState <> Normal Then
                    .TitleMonthYear.MouseState = Normal
                    ShowHandPointer False
                    tmrMouseEvent.Interval = 2
                    Exit Sub
                End If
            End If
            '--
            If .ViewNavigator <> ViewItemNavigatorDays And m_LinkedCalendars Then
                For a = 0 To UBound(udtItemsUpDownButton)
                    With udtItemsUpDownButton(a)
                        If PtInRect(.RECT2, X, Y) Then
                            If .MouseState = Normal Then
                                .MouseState = IIF(Button = vbLeftButton, Pressed, Hot)
                                udtItemsPicker(.IndexCalendar).TitleMonthYear.MouseState = Inherint
                                ShowHandPointer True
                                tmrMouseEvent.Interval = 2
                                Exit Sub
                            End If
                        Else
                            If .MouseState <> Normal Then
                                .MouseState = Normal
                                udtItemsPicker(i).TitleMonthYear.MouseState = Normal
                                ShowHandPointer False
                                Exit Sub
                            End If
                        End If
                    End With
                Next
            End If
        End With
    Next
    '---

    '-> Capturar el calendario a actualizar
    If iCalendar = -1 Then
        For a = 0 To UBound(udtItemsPicker)
            With udtItemsPicker(a)
                If PtInRect(.RECT2, X, Y) Then
                    iCalendar = .IndexCalendar
                End If
            End With
        Next
    End If
    '-> Dias del mes.
    If iCalendar <> -1 Then
        If udtItemsPicker(iCalendar).ViewNavigator = ViewItemNavigatorDays Then
            For a = iCalendar To iCalendar
                With udtItemsPicker(a)
                    If PtInRect(.RECT2, X, Y) Then
                        For i = 0 To UBound(udtItemsDay)
                            With udtItemsDay(i)
                                If PtInRect(.RECT2, X, Y) Then
                                    If .DateValue >= m_MinDate And .DateValue <= m_MaxDate Then
                                        If .MouseState = Normal Then
                                            .MouseState = IIF(Button = vbLeftButton, Pressed, Hot)
                                            ShowHandPointer True
                                            '--
                                            If IsDate(m_ValueStart) Then
                                                If (m_MaxRangeDays And DateDiff("d", m_ValueStart, .DateValue, m_UserFirstDayOfWeek) + 1 <= m_MaxRangeDays) Or Not m_MaxRangeDays > 0 Then
                                                    If ((CDate(m_ValueStart) <= .DateValue) And m_ValueEnd = "") Then
                                                        c_IndexSelMove = i
                                                    End If
                                                End If
                                            End If
                                            tmrMouseEvent.Interval = 2
                                            Exit Sub
                                        End If
                                    Else
                                        Exit Sub
                                    End If
                                Else
                                    If .MouseState <> Normal Then
                                        .MouseState = Normal
                                        ShowHandPointer False
                                        ResetControl
                                        tmrMouseEvent.Interval = 2
                                        Exit Sub
                                    End If
                                End If
                            End With
                        Next
                        Exit Sub
                    Else
                        iCalendar = -1
                        ShowHandPointer False
                        ResetControl
                        tmrMouseEvent.Interval = 2
                    End If
                End With
            Next
        End If
    End If
    '---

    '-> Meses o Años de la navegacion rapida.
    If iCalendar <> -1 Then
        If udtItemsPicker(iCalendar).ViewNavigator <> ViewItemNavigatorDays Then
            For a = iCalendar To iCalendar
                With udtItemsPicker(a)
                    If PtInRect(.RECT2, X, Y) Then
                        For i = 0 To UBound(udtItemsMonthYear)
                            With udtItemsMonthYear(i)
                                If PtInRect(.RECT2, X, Y) Then
                                    If udtItemsPicker(a).ViewNavigator = ViewItemNavigatorMonths Then
                                        bTemp = DateSerial(udtItemsPicker(a).NumberYear, .ValueItem, 1) < m_MinDate Or DateSerial(udtItemsPicker(a).NumberYear, .ValueItem, 1) > m_MaxDate
                                    Else
                                        'bTemp = .ValueItem < Year(m_MinDate) Or .ValueItem > Year(m_MinDate)
                                    End If
                                    If Not bTemp Then
                                        If .MouseState = Normal Then
                                            .MouseState = IIF(Button = vbLeftButton, Pressed, Hot)
                                            '--
                                            If udtItemsPicker(.IndexCalendar).ViewNavigator = ViewItemNavigatorMonths Then
                                                iTmpValue = udtItemsPicker(.IndexCalendar).NumberMonth
                                            Else
                                                iTmpValue = udtItemsPicker(.IndexCalendar).NumberYear
                                            End If
                                            '--
                                            If .ValueItem <> iTmpValue Then
                                                c_IndexSelMove = i
                                                tmrMouseEvent.Interval = 2
                                            End If
                                            '--
                                            ShowHandPointer True
                                            Exit Sub
                                        End If
                                    Else
                                        ShowHandPointer False
                                        Exit Sub
                                    End If
                                Else
                                    If .MouseState <> Normal Then
                                        .MouseState = Normal
                                        ShowHandPointer False
                                        tmrMouseEvent.Interval = 2
                                        Exit Sub
                                    End If
                                End If
                            End With
                        Next
                    Else
                        iCalendar = -1
                        ShowHandPointer False
                        ResetControl
                        tmrMouseEvent.Interval = 2
                    End If
                End With
            Next
        End If
    End If
    '---

    '-> Botones de rangos.
    If m_ShowRangeButtons And m_UseRangeValue Then
        For i = 0 To UBound(udtItemsRangeButton)
            With udtItemsRangeButton(i)
                If PtInRect(.RECT2, X, Y) Then
                    If .MouseState = Normal Then
                        .MouseState = IIF(Button = vbLeftButton, Pressed, Hot)
                        ShowHandPointer True
                        tmrMouseEvent.Interval = 2
                        Exit Sub
                    End If
                Else
                    If .MouseState <> Normal Then
                        .MouseState = Normal
                        ShowHandPointer False
                        tmrMouseEvent.Interval = 2
                        Exit Sub
                    End If
                End If
            End With
        Next
    End If
    '---

    '-> Botones de accion.
    If Not m_AutoApply Or m_ShowTodayButton Then
        For i = 0 To UBound(udtItemsActionButton)
            With udtItemsActionButton(i)
                If PtInRect(.RECT2, X, Y) Then
                    If .MouseState = Normal Then
                        .MouseState = IIF(Button = vbLeftButton, Pressed, Hot)
                        ShowHandPointer True
                        tmrMouseEvent.Interval = 2
                        Exit Sub
                    End If
                Else
                    If .MouseState <> Normal Then
                        .MouseState = Normal
                        ShowHandPointer False
                        tmrMouseEvent.Interval = 2
                        Exit Sub
                    End If
                End If
            End With
        Next
    End If
    '---
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, a As Integer, M As Integer
    Dim dDate As Date
    Dim pi_Temp As Integer
    Dim J As Integer
    Dim bTemp As Boolean
    '-> Botones de navegacion del calendario(s)
    For i = 0 To UBound(udtItemsNavButton)
        With udtItemsNavButton(i)
            If PtInRect(.RECT2, X, Y) Then
                M = IIF(i Mod 2 = 0, -1, 1)
                '--
                .MouseState = Hot
                '--
                If m_LinkedCalendars And b_ShowFastNavigator Then
                    For a = 0 To UBound(udtItemsPicker)
                        With udtItemsPicker(a)
                            If .ViewNavigator <> ViewItemNavigatorDays Then .ViewNavigator = ViewItemNavigatorDays
                            .NumberYear = Year(.DateInPicker)
                            .NumberMonth = Month(.DateInPicker)
                            .MonthName = MonthName(Month(dDate))
                        End With
                    Next
                    b_ShowFastNavigator = False
                End If
                For a = IIF(m_LinkedCalendars, 0, .IndexCalendar) To IIF(m_LinkedCalendars, UBound(udtItemsPicker), .IndexCalendar)
                    With udtItemsPicker(a)
                        dDate = DateSerial(.NumberYear, .NumberMonth, 1)
                        If .ViewNavigator = ViewItemNavigatorDays Then
                            dDate = DateAdd("m", M, dDate)
                        End If
                        If Not m_LinkedCalendars Then
                            If .ViewNavigator = ViewItemNavigatorMonths Then
                                dDate = DateAdd("yyyy", M, dDate)
                            ElseIf .ViewNavigator = ViewItemNavigatorYears Then
                                pi_Temp = (udtItemsMonthYear(IIF(M = -1, 0, UBound(udtItemsMonthYear) - 1)).ValueItem - Year(dDate))
                                dDate = DateAdd("yyyy", pi_Temp, dDate)
                            End If
                        End If
                        '---
                        .NumberYear = Year(dDate)
                        .NumberMonth = Month(dDate)
                        .MonthName = MonthName(Month(dDate))
                        'If Not b_ShowFastNavigator Then .DateInPicker = DateSerial(.NumberYear, .NumberMonth, 1)
                        .DateInPicker = DateSerial(.NumberYear, .NumberMonth, 1)
                        '---
                        Call ChangeViewPicker
                    End With
                Next
                '---
                Call Draw ': Refresh
                '---
            End If
        End With
    Next
    '---

    '-> Mes y año
    For i = 0 To UBound(udtItemsPicker)
        With udtItemsPicker(i)
            If PtInRect(.TitleMonthYear.RECT2, X, Y) Then
                If .TitleMonthYear.MouseState = Pressed Then
                    .TitleMonthYear.MouseState = Hot
                    For J = 0 To UBound(udtItemsPicker)
                        With udtItemsPicker(J)
                            If .ViewNavigator <> ViewItemNavigatorDays And .IndexCalendar <> i Then
                                If DateSerial(.NumberYear, .NumberMonth, 1) <> .DateInPicker Then
                                    .NumberMonth = Month(.DateInPicker)
                                    .NumberYear = Year(.DateInPicker)
                                    .MonthName = MonthName(.NumberMonth)
                                End If
                                .ViewNavigator = ViewItemNavigatorDays
                            End If
                        End With
                    Next
                    '---
                    .TitleMonthYear.MouseState = Hot
                    '---
                    If .ViewNavigator = ViewItemNavigatorDays Then
                        .ViewNavigator = ViewItemNavigatorMonths
                    ElseIf .ViewNavigator = ViewItemNavigatorMonths Then
                        .ViewNavigator = ViewItemNavigatorYears
                    End If
                    Call ChangeViewPicker
                    '---
                    Call Draw ': Refresh
                End If
            End If
            If .ViewNavigator <> ViewItemNavigatorDays And m_LinkedCalendars Then
                For a = 0 To UBound(udtItemsUpDownButton)
                    With udtItemsUpDownButton(a)
                        If PtInRect(.RECT2, X, Y) Then
                            If .MouseState = Pressed Then
                                M = IIF(a Mod 2 = 0, -1, 1)
                                '--
                                dDate = DateSerial(udtItemsPicker(i).NumberYear, udtItemsPicker(i).NumberMonth, 1)
                                '--
                                If udtItemsPicker(i).ViewNavigator = ViewItemNavigatorMonths Then
                                    dDate = DateAdd("yyyy", M, dDate)
                                ElseIf udtItemsPicker(i).ViewNavigator = ViewItemNavigatorYears Then
                                    pi_Temp = (udtItemsMonthYear(IIF(M = -1, 0, UBound(udtItemsMonthYear) - 1)).ValueItem - Year(dDate))
                                    dDate = DateAdd("yyyy", pi_Temp, dDate)
                                End If
                                udtItemsPicker(i).NumberYear = Year(dDate)
                                udtItemsPicker(i).NumberMonth = Month(dDate)
                                udtItemsPicker(i).MonthName = MonthName(Month(dDate))
                                '---
                                Call ChangeViewPicker
                                '--
                                .MouseState = Hot
                                udtItemsPicker(i).TitleMonthYear.MouseState = Inherint
                                '--
                                Call Draw ': Refresh
                            End If
                        End If
                    End With
                Next
            End If
        End With
    Next

    '-> Días de los meses.
    For i = 0 To UBound(udtItemsDay)
        With udtItemsDay(i)
            If udtItemsPicker(.IndexCalendar).ViewNavigator = ViewItemNavigatorDays Then
                'If c_EnterButton = True And ((X >= .RECT.Left And X <= (.RECT.Left + .RECT.Width)) And (Y >= .RECT.Top And (Y <= .RECT.Top + .RECT.Height))) Then
                If PtInRect(.RECT2, X, Y) Then
                    If .MouseState = Pressed Then
                        .MouseState = Hot
                        '---
                        If m_UseRangeValue Then
                            If m_ValueStart <> "" And m_ValueEnd <> "" Then
                                m_ValueStart = "": m_ValueEnd = ""
                                c_IndexSelMove = -1
                            End If
                            '--
                            If .DateValue >= m_MinDate And .DateValue <= m_MaxDate Then
                                If Len(m_ValueStart) <= 0 Then
                                    m_ValueStart = .DateValue
                                    RaiseEvent ChangeStartDate(m_ValueStart)
                                    '---
                                    If m_AutoApply Then
                                        udtItemsPicker(.IndexCalendar).DateInPicker = DateSerial(.DatePartYear, .DatePartMonth, 1)
                                    End If
                                Else
                                    If .DateValue >= CDate(m_ValueStart) Then
                                        If (m_MaxRangeDays And DateDiff("d", m_ValueStart, .DateValue, m_UserFirstDayOfWeek) + 1 <= m_MaxRangeDays) Or Not m_MaxRangeDays > 0 Then
                                            m_ValueEnd = .DateValue
                                            RaiseEvent ChangeEndDate(m_ValueEnd)
                                            If m_AutoApply Then
                                                udtItemsPicker(.IndexCalendar).DateInPicker = DateSerial(.DatePartYear, .DatePartMonth, 1)
                                                Call ApplyChangeValues
                                            End If
                                        End If
                                    End If
                                End If
                                '---
                                Call Draw ': Refresh
                                '---
                            End If
                        Else
                            If .DateValue >= m_MinDate And .DateValue <= m_MaxDate Then
                                d_ValueTemp = .DateValue
                                If m_AutoApply Then
                                    udtItemsPicker(.IndexCalendar).DateInPicker = DateSerial(.DatePartYear, .DatePartMonth, 1)
                                    '--
                                    Call ApplyChangeValues
                                End If
                                '---
                                Call Draw ': Refresh
                                '---
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next
    '---

    '-> Meses o Años de la navegacion rapida.
    For i = 0 To UBound(udtItemsMonthYear)
        With udtItemsMonthYear(i)
            If udtItemsPicker(.IndexCalendar).ViewNavigator <> ViewItemNavigatorDays Then
                If PtInRect(.RECT2, X, Y) Then
                    If udtItemsPicker(.IndexCalendar).ViewNavigator = ViewItemNavigatorMonths Then
                        bTemp = DateSerial(udtItemsPicker(.IndexCalendar).NumberYear, .ValueItem, 1) < m_MinDate Or DateSerial(udtItemsPicker(.IndexCalendar).NumberYear, .ValueItem, 1) > m_MaxDate
                    Else
                        'bTemp = .ValueItem < Year(m_MinDate) Or .ValueItem > Year(m_MinDate)
                    End If
                    If Not bTemp Then
                        If .MouseState = Pressed Then
                            .MouseState = Hot
                            '---
                            Select Case udtItemsPicker(.IndexCalendar).ViewNavigator
                                Case ViewItemNavigatorMonths
                                    If (DateSerial(udtItemsPicker(.IndexCalendar).NumberYear, .ValueItem, 1) >= m_MinDate) And (DateSerial(udtItemsPicker(.IndexCalendar).NumberYear, .ValueItem, 1) <= m_MaxDate) Then
                                        udtItemsPicker(.IndexCalendar).NumberMonth = .ValueItem
                                        udtItemsPicker(.IndexCalendar).MonthName = MonthName(.ValueItem)
                                        udtItemsPicker(.IndexCalendar).DateInPicker = DateSerial(udtItemsPicker(.IndexCalendar).NumberYear, .ValueItem, 1)
                                        udtItemsPicker(.IndexCalendar).ViewNavigator = ViewItemNavigatorDays
                                        If m_LinkedCalendars Then UpdateLinkedCalendar (.IndexCalendar): b_ShowFastNavigator = False
                                    End If
                                Case ViewItemNavigatorYears
                                    If .ValueItem >= Year(m_MinDate) And .ValueItem <= Year(m_MaxDate) Then
                                        udtItemsPicker(.IndexCalendar).NumberYear = .ValueItem
                                        udtItemsPicker(.IndexCalendar).ViewNavigator = ViewItemNavigatorMonths
                                    End If
                            End Select
                            Call ChangeViewPicker
                            Call Draw ': Refresh
                        End If
                    End If
                End If
            End If
        End With
    Next
    '---

    '-> Botones de rangos.
    If m_ShowRangeButtons And m_UseRangeValue Then
        For i = 0 To UBound(udtItemsRangeButton)
            With udtItemsRangeButton(i)
                If PtInRect(.RECT2, X, Y) Then
                    If .MouseState = Pressed Then
                        .MouseState = Hot
                        RaiseEvent ButtonRangeClick(i, .Caption)
                        Call Draw ': Refresh
                    End If
                End If
            End With
        Next
    End If
    '---

    '-> Botones de accion.
    If Not m_AutoApply Or m_ShowTodayButton Then
        For i = 0 To UBound(udtItemsActionButton)
            With udtItemsActionButton(i)
                If PtInRect(.RECT2, X, Y) Then
                    If .MouseState = Pressed Then
                        .MouseState = Hot
                        Select Case .ButtonAction
                            Case [Action Today]
                                d_ValueTemp = Date
                                With udtItemsPicker(0)
                                    .NumberYear = Year(d_ValueTemp)
                                    .NumberMonth = Month(d_ValueTemp)
                                    .MonthName = MonthName(Month(d_ValueTemp))
                                    If m_UseRangeValue Then
                                        .DateInPicker = DateSerial(.NumberYear, .NumberMonth, 1)
                                        m_ValueStart = d_ValueTemp: RaiseEvent ChangeStartDate(d_ValueTemp)
                                        m_ValueEnd = d_ValueTemp: RaiseEvent ChangeEndDate(d_ValueTemp)
                                    Else
                                        .DateInPicker = d_ValueTemp
                                        RaiseEvent ChangeDate(d_ValueTemp)
                                    End If
                                End With
                                Call UpdateLinkedCalendar(0)
                            Case [Action Cancel]
                                If Not IsChild Then
                                    Call HideCalendar
                                    Exit Sub
                                Else
                                    ValueStart = ""
                                    RaiseEvent ChangeStartDate("")
                                    ValueEnd = ""
                                    RaiseEvent ChangeEndDate("")
                                End If
                            Case [Action Apply]
                                If ApplyChangeValues Then Exit Sub
                        End Select
                        RaiseEvent ButtonActionClick(i, .Caption)
                        Call Draw ': Refresh
                    End If
                End If
            End With
        Next
    End If
    '---
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '---
    'hFontCollection = ReadValue(&HFC)
    '---
    c_hWnd = UserControl.ContainerHwnd
    If Ambient.UserMode Then UserControl.Picture = Nothing
    '---
    With PropBag
        m_BackColor = .ReadProperty("BackColor", &HFFFFFF)
        m_BackOpacity = .ReadProperty("BackOpacity", 100)
        m_Border = .ReadProperty("Border", True)
        m_BorderWidth = .ReadProperty("BorderWidth", 1)
        m_BorderColor = .ReadProperty("BorderColor", SystemColorConstants.vbActiveBorder)
        m_BorderOpacity = .ReadProperty("BorderOpacity", 100)
        m_BorderPosition = .ReadProperty("BorderPosition", [Border Center])
        m_CornerTopLeft = .ReadProperty("CornerTopLeft", 0)
        m_CornerTopRight = .ReadProperty("CornerTopRight", 0)
        m_CornerBottomLeft = .ReadProperty("CornerBottomLeft", 0)
        m_CornerBottomRight = .ReadProperty("CornerBottomRight", 0)
        m_PaddingX = .ReadProperty("PaddingX", 20)
        m_PaddingY = .ReadProperty("PaddingY", 20)
        m_Shadow = .ReadProperty("Shadow", False)
        m_ShadowColor = .ReadProperty("ShadowColor", ColorConstants.vbBlack)
        m_ShadowSize = .ReadProperty("ShadowSize", 0)
        m_ShadowOpacity = .ReadProperty("ShadowOpacity", 50)
        m_ShadowOffsetX = .ReadProperty("ShadowOffsetX", 0)
        m_ShadowOffsetY = .ReadProperty("ShadowOffsetY", 0)
        m_ShowNumberWeek = .ReadProperty("ShowNumberWeek", True)
        m_ShowUseISOWeek = .ReadProperty("ShowUseISOWeek", True)
        m_SpaceGrid = .ReadProperty("SpaceGrid", 2)
        m_MouseToParent = .ReadProperty("MouseToParent", False)
        m_UseGDIPString = .ReadProperty("UseGDIPString", False)
        '--
        m_SinglePicker = .ReadProperty("SinglePicker", False)
        m_UseRangeValue = .ReadProperty("UseRangeValue", False)
        m_ShowRangeButtons = .ReadProperty("ShowRangeButtons", False)
        m_ShowTodayButton = .ReadProperty("ShowTodayButton", False)
        m_RightToLeft = .ReadProperty("RightToLeft", False)
        m_AutoApply = .ReadProperty("AutoApply", True)
        m_IsChild = .ReadProperty("IsChild", False)
        m_BackColorParent = .ReadProperty("BackColorParent", Ambient.BackColor)
        
        m_ColsPicker = .ReadProperty("ColsPicker", 0)
        m_NumberPickers = .ReadProperty("NumberPickers", IIF(Not m_SinglePicker, 2, 1))
        'm_ShowTimePicker = .ReadProperty("ShowTimePicker", False)                       'Para trabajar el timepicker
        'm_UseTimePicker24Hrs = .ReadProperty("UseTimePicker24Hrs", False)               'Para formato de 24 horas en timepicker
        'm_TimerWithSecond = .ReadProperty("TimerWithSecond", False)                     'Para usar el timer hasta los segundos.
        m_MaxRangeDays = .ReadProperty("MaxRangeDays", 0)

        'm_AlwaysShowCalendars = .ReadProperty("AlwaysShowCalendars", True)
        m_LinkedCalendars = .ReadProperty("LinkedCalendars", True)
        
        m_Value = .ReadProperty("Value", Date)
        m_ValueStart = .ReadProperty("ValueStart", "")
        m_ValueEnd = .ReadProperty("ValueEnd", "")
        
        m_MinDate = .ReadProperty("MinDate", DateSerial(1601, 1, 1))
        m_MaxDate = .ReadProperty("MaxDate", DateSerial(9999, 12, 31))
        m_FirstDayOfWeek = .ReadProperty("FirstDayOfWeek", vbUseSystemDayOfWeek)
        
        m_CountFreeDays = .ReadProperty("CountFreeDays", True)
        m_CountReservedDay = .ReadProperty("CountReservedDay", True)
        
        '---> De los botones de navegacion.
        m_ButtonNavBackColor = .ReadProperty("ButtonNavBackColor", &HFFFFFF)
        m_ButtonNavBorderWidth = .ReadProperty("ButtonNavBorderWidth", 1)
        m_ButtonNavBorderColor = .ReadProperty("ButtonNavBorderColor", SystemColorConstants.vbActiveBorder)
        m_ButtonNavCornerRadius = .ReadProperty("ButtonNavCornerRadius", 0)
        m_ButtonNavForeColor = .ReadProperty("ButtonNavForeColor", SystemColorConstants.vbButtonText)
        m_ButtonNavIsIcoFont = .ReadProperty("ButtonNavIsIcoFont", False)
        Set m_ButtonNavIcoFont = .ReadProperty("ButtonNavIcoFont", Ambient.Font)
        m_ButtonNavCharCodeBack = .ReadProperty("ButtonNavCharCodeBack", 0)
        m_ButtonNavCharCodeNext = .ReadProperty("ButtonNavCharCodeNext", 0)
        m_ButtonNavWidth = .ReadProperty("ButtonNavWidth", 24)
        m_ButtonNavHeight = .ReadProperty("ButtonNavHeight", 24)
        
        '---> De los botones de accion.
        m_ButtonsBackColor = .ReadProperty("ButtonsBackColor", &HFFFFFF)
        m_ButtonsBorderWidth = .ReadProperty("ButtonsBorderWidth", 1)
        m_ButtonsBorderColor = .ReadProperty("ButtonsBorderColor", SystemColorConstants.vbActiveBorder)
        m_ButtonsCornerRadius = .ReadProperty("ButtonsCornerRadius", 5)
        Set m_ButtonsFont = .ReadProperty("ButtonsFont", Ambient.Font)
        m_ButtonsForeColor = .ReadProperty("ButtonsForeColor", SystemColorConstants.vbButtonText)
        m_ButtonsWidth = .ReadProperty("ButtonsWidth", 24)
        m_ButtonsHeight = .ReadProperty("ButtonsHeight", 24)

        '---> De los meses y años
        m_MonthYearBackColor = .ReadProperty("MonthYearBackColor", &HFFFFFF)
        m_MonthYearBorderWidth = .ReadProperty("MonthYearBorderWidth", 0)
        m_MonthYearBorderColor = .ReadProperty("MonthYearBorderColor", SystemColorConstants.vbActiveBorder)
        m_MonthYearCornerRadius = .ReadProperty("MonthYearCornerRadius", 0)
        Set m_MonthYearFont = .ReadProperty("MonthYearFont", Ambient.Font)
        m_MonthYearForeColor = .ReadProperty("MonthYearForeColor", SystemColorConstants.vbButtonText)

        '---> De las Semanas.
        m_WeekBackColor = .ReadProperty("WeekBackColor", &HFFFFFF)
        m_WeekBorderWidth = .ReadProperty("WeekBorderWidth", 0)
        m_WeekBorderColor = .ReadProperty("WeekBorderColor", SystemColorConstants.vbActiveBorder)
        m_WeekCornerRadius = .ReadProperty("WeekCornerRadius", 0)
        Set m_WeekFont = .ReadProperty("WeekFont", Ambient.Font)
        m_WeekFontHeaderBold = .ReadProperty("WeekFontHeaderBold", True)
        m_WeekForeColor = .ReadProperty("WeekForeColor", SystemColorConstants.vbButtonText)
        m_WeekWidth = .ReadProperty("WeekWidth", 26)
        m_WeekHeight = .ReadProperty("WeekHeight", 24)

        '---> De los días.
        m_DayBackColor = .ReadProperty("DayBackColor", &HFFFFFF)
        m_DayBorderWidth = .ReadProperty("DayBorderWidth", 0)
        m_DayBorderColor = .ReadProperty("DayBorderColor", SystemColorConstants.vbActiveBorder)
        m_DayCornerRadius = .ReadProperty("DayCornerRadius", 0)
        Set m_DayFont = .ReadProperty("DayFont", Ambient.Font)
        m_DayHeaderFontBold = .ReadProperty("DayHeaderFontBold", True)
        m_DayHeaderForeColor = .ReadProperty("DayHeaderForeColor", SystemColorConstants.vbButtonText)
        m_DayHotColor = .ReadProperty("DayHotColor", SystemColorConstants.vbHighlight)
        m_DayForeColor = .ReadProperty("DayForeColor", SystemColorConstants.vbButtonText)
        m_DayOMForeColor = .ReadProperty("DayOMForeColor", SystemColorConstants.vbGrayText)
        m_DayFreeArray = .ReadProperty("DayFreeArray", Null)
        m_DayFreeForeColor = .ReadProperty("DayFreeForeColor", ColorConstants.vbRed)
        
        m_DayNowShow = .ReadProperty("DayNowShow", False)
        m_DayNowBorderWidth = .ReadProperty("DayNowBorderWidth", 0)
        m_DayNowBorderColor = .ReadProperty("DayNowBorderColor", SystemColorConstants.vbActiveBorder)
        m_DayNowBackColor = .ReadProperty("DayNowBackColor", &HFFFFFF)
        m_DayNowForeColor = .ReadProperty("DayNowForeColor", SystemColorConstants.vbButtonText)
        
        'Mouse event (Los días no tendran color para el mousedown)
        m_DayOverBackColor = .ReadProperty("DayOverBackColor", &H999999)
        m_DayOverForeColor = .ReadProperty("DayOverForeColor", &HFFFFFF)
        '--
        m_DaySelBetweenColor = .ReadProperty("DaySelBetweenColor", &HE5D7CA)
        'm_DaySelEndColor = .ReadProperty("DaySelEndColor", &HB06D00)
        m_DaySelValuesColor = .ReadProperty("DaySelValuesColor", &HB06D00)
        m_DaySelForeColor = .ReadProperty("DaySelForeColor", &HFFFFFF)
        m_DaySelFontBold = .ReadProperty("DaySelFontBold", True)
        m_DaySelectionStyle = .ReadProperty("DaySelectionStyle", [Corner No Between])
        'm_DayShowHotItem = .ReadProperty("DayShowHotItem", True) 'Siempre muestra el hot
        m_DaySaturdayForeColor = .ReadProperty("DaySaturdayForeColor", &H999999)
        m_DaySundayForeColor = .ReadProperty("DaySundayForeColor", &H999999)
        m_DayWidth = .ReadProperty("DayWidth", 24)
        m_DayHeight = .ReadProperty("DayHeight", 24)
        '---
        m_CallOut = .ReadProperty("CallOut", True)
        m_CallOutWidth = .ReadProperty("CallOutWidth", 20)
        m_CallOutHight = .ReadProperty("CallOutHight", 10)
        m_CallOutRightTriangle = .ReadProperty("CallOutRightTriangle", False)
        m_CallOutPosition = .ReadProperty("CallOutPosition", [Position Top])
        m_CallOutAlign = .ReadProperty("CallOutAlign", Middle)
        m_CallOutCustomPosPercent = .ReadProperty("CallOutCustomPosPercent", 0)
        '---
        UpdateScaleDPI
        'Call CreateShadow
    End With
    If m_IsChild Then
        InitControl
    Else
        Extender.Visible = False
    End If
    '---
End Sub

Private Sub UserControl_Resize()
    '---
    If Ambient.UserMode Or m_IsChild Then
        If c_Width > 0 Then UserControl.Width = c_Width
        If c_Height > 0 Then UserControl.Height = c_Height
        RaiseEvent Resize
    ElseIf Not m_IsChild Then
        UserControl.Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
    End If
    '---
End Sub

Private Sub UserControl_Show()
    '---
    InitControl
    '---
    'If Ambient.UserMode Then
    If Ambient.UserMode And Not m_IsChild Then
        With c_CShadow
            .BackColor = m_BackColor
            .Border = m_Border
            .BorderColor = m_BorderColor
            .BorderOpacity = m_BorderOpacity
            .BorderWidth = m_BorderWidth
            .CornerTopLeft = m_CornerTopLeft
            .CornerTopRight = m_CornerTopRight
            .CornerBottomLeft = m_CornerBottomLeft
            .CornerBottomRight = m_CornerBottomRight
            .CallOut = m_CallOut
            .CallOutWidth = m_CallOutWidth
            .CallOutHight = m_CallOutHight
            .CallOutPosition = m_CallOutPosition
            .CallOutAlign = m_CallOutAlign
            .CallOutCustomPos = m_CallOutCustomPosPercent
            .Shadow = m_Shadow
            .ShadowSize = m_ShadowSize
            .ShadowOpacity = m_ShadowOpacity
            .ShadowColor = m_ShadowColor
            .ShadowOffsetX = m_ShadowOffsetX
            .ShadowOffsetY = m_ShadowOffsetY
            '--
            .InitShadow UserControl.hWnd, UserControl.hdc
        End With
        '---
        UserControl.BackColor = m_BackColor
        Draw
        Refresh
    ElseIf Not Ambient.UserMode And Not m_IsChild Then
        UserControl.Cls
    ElseIf m_IsChild Then
        Draw
        Refresh
    End If
    '---
End Sub

Private Sub UserControl_Terminate()
    Set c_SubClass = Nothing
    Set c_CShadow = Nothing
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '---
    With PropBag
        Call .WriteProperty("BackColor", m_BackColor, &HFFFFFF)
        Call .WriteProperty("BackOpacity", m_BackOpacity, 100)
        Call .WriteProperty("Border", m_Border, True)
        Call .WriteProperty("BorderWidth", m_BorderWidth, 1)
        Call .WriteProperty("BorderColor", m_BorderColor, SystemColorConstants.vbActiveBorder)
        Call .WriteProperty("BorderOpacity", m_BorderOpacity, 100)
        Call .WriteProperty("BorderPosition", m_BorderPosition, [Border Center])
        Call .WriteProperty("CornerTopLeft", m_CornerTopLeft, 0)
        Call .WriteProperty("CornerTopRight", m_CornerTopRight, 0)
        Call .WriteProperty("CornerBottomLeft", m_CornerBottomLeft, 0)
        Call .WriteProperty("CornerBottomRight", m_CornerBottomRight, 0)
        Call .WriteProperty("PaddingX", m_PaddingX, 20)
        Call .WriteProperty("PaddingY", m_PaddingY, 20)
        Call .WriteProperty("Shadow", m_Shadow, False)
        Call .WriteProperty("ShadowColor", m_ShadowColor, ColorConstants.vbBlack)
        Call .WriteProperty("ShadowSize", m_ShadowSize, 0)
        Call .WriteProperty("ShadowOpacity", m_ShadowOpacity, 50)
        Call .WriteProperty("ShadowOffsetX", m_ShadowOffsetX, 0)
        Call .WriteProperty("ShadowOffsetY", m_ShadowOffsetY, 0)
        Call .WriteProperty("ShowNumberWeek", m_ShowNumberWeek, True)
        Call .WriteProperty("ShowUseISOWeek", m_ShowUseISOWeek, True)
        Call .WriteProperty("SpaceGrid", m_SpaceGrid, 2)
        Call .WriteProperty("MouseToParent", m_MouseToParent, False)
        Call .WriteProperty("UseGDIPString", m_UseGDIPString, False)
        '--
        Call .WriteProperty("SinglePicker", m_SinglePicker, False)
        Call .WriteProperty("UseRangeValue", m_UseRangeValue, False)
        Call .WriteProperty("ShowRangeButtons", m_ShowRangeButtons, False)
        Call .WriteProperty("ShowTodayButton", m_ShowTodayButton, False)
        Call .WriteProperty("RightToLeft", m_RightToLeft, False)
        Call .WriteProperty("AutoApply", m_AutoApply, True)
        Call .WriteProperty("IsChild", m_IsChild, False)
        Call .WriteProperty("BackColorParent", m_BackColorParent, False)
        
        Call .WriteProperty("ColsPicker", m_ColsPicker, 0)
        Call .WriteProperty("NumberPickers", m_NumberPickers, IIF(Not m_SinglePicker, 2, 1))
        'Call .WriteProperty("ShowTimePicker", m_ShowTimePicker, False)                          'Para trabajar el timepicker.
        'Call .WriteProperty("UseTimePicker24Hrs", m_UseTimePicker24Hrs, False)                  'Para formato de 24 horas en timepicker.
        'Call .WriteProperty("TimerWithSecond", m_TimerWithSecond, False)                        'Para usar el timer hasta los segundos.
        Call .WriteProperty("MaxRangeDays", m_MaxRangeDays, 0)

        'Call .WriteProperty("AlwaysShowCalendars", m_AlwaysShowCalendars, True)
        Call .WriteProperty("LinkedCalendars", m_LinkedCalendars, True)
        
        Call .WriteProperty("Value", m_Value, Date)
        Call .WriteProperty("ValueStart", m_ValueStart, "")
        Call .WriteProperty("ValueEnd", m_ValueEnd, "")
        
        Call .WriteProperty("MinDate", m_MinDate, DateSerial(1601, 1, 1))
        Call .WriteProperty("MaxDate", m_MaxDate, DateSerial(9999, 12, 31))
        
        Call .WriteProperty("FirstDayOfWeek", m_FirstDayOfWeek, vbUseSystemDayOfWeek)
        
        Call .WriteProperty("CountFreeDays", m_CountFreeDays, True)
        Call .WriteProperty("CountReservedDay", m_CountReservedDay, True)
        
        '---> De los botones de navegacion.
        Call .WriteProperty("ButtonNavBackColor", m_ButtonNavBackColor, &HFFFFFF)
        Call .WriteProperty("ButtonNavBorderWidth", m_ButtonNavBorderWidth, 1)
        Call .WriteProperty("ButtonNavBorderColor", m_ButtonNavBorderColor, SystemColorConstants.vbActiveBorder)
        Call .WriteProperty("ButtonNavCornerRadius", m_ButtonNavCornerRadius, 0)
        Call .WriteProperty("ButtonNavForeColor", m_ButtonNavForeColor, SystemColorConstants.vbButtonText)
        Call .WriteProperty("ButtonNavIsIcoFont", m_ButtonNavIsIcoFont, False)
        Call .WriteProperty("ButtonNavIcoFont", m_ButtonNavIcoFont, Ambient.Font)
        Call .WriteProperty("ButtonNavCharCodeBack", m_ButtonNavCharCodeBack, 0)
        Call .WriteProperty("ButtonNavCharCodeNext", m_ButtonNavCharCodeNext, 0)
        Call .WriteProperty("ButtonNavWidth", m_ButtonNavWidth, 24)
        Call .WriteProperty("ButtonNavHeight", m_ButtonNavHeight, 24)
        
        '---> De los botones de accion.
        Call .WriteProperty("ButtonsBackColor", m_ButtonsBackColor, &HFFFFFF)
        Call .WriteProperty("ButtonsBorderWidth", m_ButtonsBorderWidth, 1)
        Call .WriteProperty("ButtonsBorderColor", m_ButtonsBorderColor, SystemColorConstants.vbActiveBorder)
        Call .WriteProperty("ButtonsCornerRadius", m_ButtonsCornerRadius, 5)
        Call .WriteProperty("ButtonsFont", m_ButtonsFont, Ambient.Font)
        Call .WriteProperty("ButtonsForeColor", m_ButtonsForeColor, SystemColorConstants.vbButtonText)
        Call .WriteProperty("ButtonsWidth", m_ButtonsWidth, 24)
        Call .WriteProperty("ButtonsHeight", m_ButtonsHeight, 24)

        '---> De los meses y años
        Call .WriteProperty("MonthYearBackColor", m_MonthYearBackColor, &HFFFFFF)
        Call .WriteProperty("MonthYearBorderWidth", m_MonthYearBorderWidth, 0)
        Call .WriteProperty("MonthYearBorderColor", m_MonthYearBorderColor, SystemColorConstants.vbActiveBorder)
        Call .WriteProperty("MonthYearCornerRadius", m_MonthYearCornerRadius, 0)
        Call .WriteProperty("MonthYearFont", m_MonthYearFont, Ambient.Font)
        Call .WriteProperty("MonthYearForeColor", m_MonthYearForeColor, SystemColorConstants.vbButtonText)

        '---> De las Semanas.
        Call .WriteProperty("WeekBackColor", m_WeekBackColor, &HFFFFFF)
        
        Call .WriteProperty("WeekBorderWidth", m_WeekBorderWidth, 0)
        Call .WriteProperty("WeekBorderColor", m_WeekBorderColor, SystemColorConstants.vbActiveBorder)
        Call .WriteProperty("WeekCornerRadius", m_WeekCornerRadius, 0)
        Call .WriteProperty("WeekFont", m_WeekFont, Ambient.Font)
        Call .WriteProperty("WeekFontHeaderBold", m_WeekFontHeaderBold, True)
        Call .WriteProperty("WeekForeColor", m_WeekForeColor, SystemColorConstants.vbButtonText)
        Call .WriteProperty("WeekWidth", m_WeekWidth, 26)
        Call .WriteProperty("WeekHeight", m_WeekHeight, 24)

        '---> De los días.
        Call .WriteProperty("DayBackColor", m_DayBackColor, &HFFFFFF)
        Call .WriteProperty("DayBorderWidth", m_DayBorderWidth, 0)
        Call .WriteProperty("DayBorderColor", m_DayBorderColor, SystemColorConstants.vbActiveBorder)
        Call .WriteProperty("DayCornerRadius", m_DayCornerRadius, 0)
        Call .WriteProperty("DayFont", m_DayFont, Ambient.Font)
        Call .WriteProperty("DayHeaderFontBold", m_DayHeaderFontBold, True)
        Call .WriteProperty("DayHeaderForeColor", m_DayHeaderForeColor, SystemColorConstants.vbButtonText)
        Call .WriteProperty("DayHotColor", m_DayHotColor, SystemColorConstants.vbHighlight)
        Call .WriteProperty("DayForeColor", m_DayForeColor, SystemColorConstants.vbButtonText)
        Call .WriteProperty("DayOMForeColor", m_DayOMForeColor, SystemColorConstants.vbGrayText)
        Call .WriteProperty("DayFreeArray", m_DayFreeArray, Null)
        Call .WriteProperty("DayFreeForeColor", m_DayFreeForeColor, ColorConstants.vbRed)
        Call .WriteProperty("DayNowShow", m_DayNowShow, False)
        Call .WriteProperty("DayNowBorderWidth", m_DayNowBorderWidth, 0)
        Call .WriteProperty("DayNowBorderColor", m_DayNowBorderColor, SystemColorConstants.vbActiveBorder)
        Call .WriteProperty("DayNowBackColor", m_DayNowBackColor, &HFFFFFF)
        Call .WriteProperty("DayNowForeColor", m_DayNowForeColor, SystemColorConstants.vbButtonText)
        
        'Mouse event (Los días no tendran color para el mousedown)
        Call .WriteProperty("DayOverBackColor", m_DayOverBackColor, &H999999)
        Call .WriteProperty("DayOverForeColor", m_DayOverForeColor, &HFFFFFF)
        '--
        Call .WriteProperty("DaySelBetweenColor", m_DaySelBetweenColor, &HE5D7CA)
        'Call .WriteProperty("DaySelEndColor", m_DaySelEndColor, &H0)
        Call .WriteProperty("DaySelValuesColor", m_DaySelValuesColor, &HB06D00)
        Call .WriteProperty("DaySelForeColor", m_DaySelForeColor, &HFFFFFF)
        Call .WriteProperty("DaySelFontBold", m_DaySelFontBold, True)
        Call .WriteProperty("DaySelectionStyle", m_DaySelectionStyle, [Corner No Between])
        'Call .WriteProperty("DayShowHotItem", m_DayShowHotItem, True) ' Seimpre muestra el hot
        Call .WriteProperty("DaySaturdayForeColor", m_DaySaturdayForeColor, &H999999)
        Call .WriteProperty("DaySundayForeColor", m_DaySundayForeColor, &H999999)
        Call .WriteProperty("DayWidth", m_DayWidth, 24)
        Call .WriteProperty("DayHeight", m_DayHeight, 24)
        '---
        Call .WriteProperty("CallOut", m_CallOut, True)
        Call .WriteProperty("CallOutWidth", m_CallOutWidth, 20)
        Call .WriteProperty("CallOutHight", m_CallOutHight, 10)
        Call .WriteProperty("CallOutRightTriangle", m_CallOutRightTriangle, False)
        Call .WriteProperty("CallOutPosition", m_CallOutPosition, [Position Top])
        Call .WriteProperty("CallOutAlign", m_CallOutAlign, Middle)
        Call .WriteProperty("CallOutCustomPosPercent", m_CallOutCustomPosPercent, 0)
        '---
        'Call CreateShadow
    End With
    '---
End Sub

'***********
'* Rutinas *
'***********
'-> Publicas
Public Sub Refresh()
    Call InitControl
    Call Draw
End Sub
Public Sub ShowCalendar(Left As Long, Top As Long)
    If Not c_Show Then
        Extender.Visible = False
        Extender.Visible = True
        '---
        c_PhWnd = UserControl.ContainerHwnd
        '---
        SetParent UserControl.hWnd, 0
        SetWindowPos UserControl.hWnd, HWND_TOPMOST, Left, Top, UserControl.ScaleWidth, UserControl.ScaleHeight, SWP_NOACTIVATE Or SWP_SHOWWINDOW
        PutFocus UserControl.hWnd
        '--
        c_Show = True
        tmrMouseEvent.Interval = 1
        '---
        With c_SubClass
            If .ssc_Subclass(UserControl.hWnd, , , Me) Then
                .ssc_AddMsg UserControl.hWnd, WM_HOTKEY, MSG_AFTER
                .ssc_AddMsg UserControl.hWnd, WM_MOUSELEAVE, MSG_BEFORE
                .ssc_AddMsg UserControl.hWnd, WM_IME_SETCONTEXT, MSG_BEFORE
            End If
        End With
        '---
        InitControl
        Draw
        '---
    End If
End Sub

Public Sub HideCalendar()
    If Not c_Show Then Exit Sub
    '--
    c_SubClass.ssc_UnSubclass UserControl.hWnd
    ShowWindow UserControl.hWnd, 0
    SetParent UserControl.hWnd, c_PhWnd
    c_CShadow.EndShadow
    '--
    c_Show = False
    tmrMouseEvent.Interval = 0
End Sub

Public Sub SetRangeButtonsCaption(Index As Integer, strCaption As String)
    On Error GoTo ErrorRutina
    '---
    If Index > UBound(udtItemsRangeButton) Then
        Err.Raise Number:="5001", Description:="Invalid button index"
    Else
        udtItemsRangeButton(Index).Caption = strCaption
    End If
    Exit Sub
    '---
ErrorRutina:
    MsgBox "Error Nro.: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbExclamation
End Sub

Public Sub SetActionButtons(Index As Integer, strCaption As String, Action As enmButtonAction)
    On Error GoTo ErrorRutina
    '---
    If Index > UBound(udtItemsActionButton) Then
        Err.Raise Number:="5001", Description:="Invalid button index"
    Else
        With udtItemsActionButton(Index)
            If Len(Trim(strCaption)) Then .Caption = strCaption
            If Action <> [Action None] Then .ButtonAction = Action
        End With
    End If
    Exit Sub
    '---
ErrorRutina:
    MsgBox "Error Nro.: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbExclamation
End Sub


'-> Privadas:
Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

Private Sub InitControl()
    '---
    Dim hGraphics As Long, hImage As Long
    Dim i As Integer, a As Integer, jump As Integer, tmpW, tmph As Integer, countItem As Integer
    Dim sRECTL As RECTL
    Dim dDate() As Date
    Dim coWidth As Long
    Dim coHeight As Long
    Dim btnWidth As Long
    Dim PaddingX As Long
    Dim PaddingY As Long
    '---> Ininicar hGraphics para calular area de texto.
    GdipCreateBitmapFromScan0 UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    
    '---> Setear variable temporales de los values
    d_ValueTemp = m_Value
    'd_ValueStartTemp = m_ValueStart
    'd_ValueEndTemp = m_ValueEnd
    
    '---> Redimesionar las matrices.
    If m_SinglePicker Then m_NumberPickers = 1
    '---
    If m_CallOut Then
        Select Case m_CallOutPosition
            Case [Position Left], [Position Right]
                coWidth = m_CallOutHight * nScale
            Case [Position Top], [Position Bottom]
                coHeight = m_CallOutHight * nScale
        End Select
    End If
    
    'udtItemMonthYear (numero de items para el grid de los meses o años de la navegacion rapida).
    ReDim udtItemsMonthYear(cs_ColsMonthYear * cs_RowsMonthYear) As udtItemMonthYear
    
    'udtItemsPicker (1 picker equivale items(0))
    ReDim udtItemsPicker(m_NumberPickers - 1) As udtItemDatePicker
    ReDim dDate(m_NumberPickers - 1) As Date
    
    'udtItemsNavButton (1 picker: 2 items(0 to 1))
    ReDim udtItemsNavButton(IIF(m_LinkedCalendars, 1, (2 * m_NumberPickers) - 1)) As udtItemCalendarButton
    
    'udtItemsHeaderDays (1 picker: 7 items(0 to 6))
    ReDim udtItemsHeaderDay((cs_ColsDay * m_NumberPickers) - 1) As udtItemHeaderDay
    
    'udtItemsWeek (1 picker: 6 items(0 to 5))
    ReDim udtItemsWeek((cs_RowsPicker * m_NumberPickers) - 1) As udtItemWeek
    
    'udtItemsDay (1 picker: 42 items(0 to 41))
    ReDim udtItemsDay((cs_ItemsDay * m_NumberPickers) - 1) As udtItemDayCalendar
    
    m_UserFirstDayOfWeek = IIF(m_FirstDayOfWeek <> vbUseSystemDayOfWeek, m_FirstDayOfWeek, GetLocaleInfoAsLong(LOCALE_IFIRSTDAYOFWEEK))
    m_UserFirstDayOfWeek = IIF(m_FirstDayOfWeek = vbUseSystemDayOfWeek, IIF(m_UserFirstDayOfWeek = 6, 1, m_UserFirstDayOfWeek + 2), m_FirstDayOfWeek)
    
    '--
    PaddingX = m_PaddingX
    PaddingY = m_PaddingY
    'Cacular posicion y tamaños del o los calendarios
    jump = 0
    For i = 0 To UBound(udtItemsPicker)
        With udtItemsPicker(i)
            .RECT.Width = (m_SpaceGrid + ((m_WeekWidth + m_SpaceGrid) * Abs(m_ShowNumberWeek)) + ((m_DayWidth + m_SpaceGrid) * cs_ColsDay))
            .RECT.Height = (m_SpaceGrid + m_ButtonNavHeight + PaddingY + ((m_DayHeight + m_SpaceGrid) * cs_RowsPicker))
            '---
            If m_ColsPicker > 0 Then
                If a = m_ColsPicker Then a = 0: jump = jump + 1
                .RECT.Left = IIF(m_CallOutPosition = [Position Left] And Not m_IsChild, coWidth, 0) + PaddingX + ((.RECT.Width + PaddingX) * a)
                .RECT.Top = IIF(m_CallOutPosition = [Position Top] And Not m_IsChild, coHeight, 0) + PaddingY + ((.RECT.Height + PaddingY) * jump)
            Else
                .RECT.Left = IIF(m_CallOutPosition = [Position Left] And Not m_IsChild, coWidth, 0) + PaddingX + ((udtItemsPicker(0).RECT.Width + PaddingX) * i)
                .RECT.Top = IIF(m_CallOutPosition = [Position Top] And Not m_IsChild, coHeight, 0) + PaddingY
            End If
            .IndexCalendar = i
            .ViewNavigator = ViewItemNavigatorDays
            a = a + 1
            '--
            .RECT2 = RectLToRect(.RECT)
        End With
    Next
    
    'Mes y el año del calendario (Obtener tamaño del area para el texto)
    '->Obtener alto del texto para el área.
    GdiPlusGetMeasureString hGraphics, "MES", sRECTL.Width, sRECTL.Height, m_MonthYearFont
    'GetMeasureText hdc, "MES", sRECTL.Width, sRECTL.Height, m_MonthYearFont
    For i = 0 To UBound(udtItemsPicker)
        With udtItemsPicker(i)
            .TitleMonthYear.RECT.Left = .RECT.Left + m_ButtonNavWidth + PaddingX
            .TitleMonthYear.RECT.Width = .RECT.Width - ((m_ButtonNavWidth + PaddingX) * 2)
            .TitleMonthYear.RECT.Top = .RECT.Top + m_SpaceGrid
            .TitleMonthYear.RECT.Height = IIF(sRECTL.Height < m_ButtonNavHeight, m_ButtonNavHeight, sRECTL.Height)
            .TitleMonthYear.RECT2 = RectLToRect(.TitleMonthYear.RECT)
        End With
    Next
    
    'Botones de los calendarios "< >"
    'udtItemsPicker
    'udtItemsCalendarButton
    For i = 0 To UBound(udtItemsPicker)
        If m_LinkedCalendars Then
            If i = 0 Then
                udtItemsNavButton(i).RECT.Left = udtItemsPicker(i).RECT.Left + m_SpaceGrid
                udtItemsNavButton(i).RECT.Top = udtItemsPicker(i).RECT.Top + m_SpaceGrid
            End If
            If i = IIF(m_ColsPicker, m_ColsPicker - 1, UBound(udtItemsPicker)) Then
                udtItemsNavButton(UBound(udtItemsNavButton)).RECT.Left = (udtItemsPicker(i).RECT.Left + udtItemsPicker(i).RECT.Width) - (m_ButtonNavWidth + m_SpaceGrid)
                udtItemsNavButton(UBound(udtItemsNavButton)).RECT.Top = udtItemsPicker(i).RECT.Top + m_SpaceGrid
            End If
            '---
            For a = 0 To UBound(udtItemsNavButton)
            '---
                With udtItemsNavButton(a)
                    .IndexCalendar = 0
                    .RECT.Width = m_ButtonNavWidth
                    .RECT.Height = m_ButtonNavHeight
                    .RECT2 = RectLToRect(.RECT)
                End With
            '---
            Next
        Else
            For a = 0 + (i * 2) To (i * 2) + 1
                With udtItemsNavButton(a)
                    .IndexCalendar = i
                    .RECT.Left = udtItemsPicker(i).RECT.Left + (IIF(a = 0, m_SpaceGrid, 0)) + ((udtItemsPicker(i).RECT.Width - (m_ButtonNavWidth + m_SpaceGrid)) * (a Mod 2))
                    .RECT.Top = udtItemsPicker(i).RECT.Top + m_SpaceGrid + ((udtItemsPicker(i).TitleMonthYear.RECT.Height - m_ButtonNavHeight) / 2)
                    .RECT.Width = m_ButtonNavWidth
                    .RECT.Height = m_ButtonNavHeight
                    .RECT2 = RectLToRect(.RECT)
                End With
            Next
        End If
    Next
    
    'Botones de rangos.
    'udtItemsRangeButton(5) 'Para botones de rangos ('Hoy', 'Este mes', 'Mes pasado', 'Ultimos 90 días', 'Este Año', 'Año pasado')
    'Definir nombres en los botones de rangos:
    If m_ShowRangeButtons And m_UseRangeValue Then
        '--
        btnWidth = 0
        For i = 0 To UBound(udtItemsRangeButton)
            With udtItemsRangeButton(i)
                SetRect sRECTL, 0, 0, 0, 0
                GdiPlusGetMeasureString hGraphics, .Caption, sRECTL.Width, sRECTL.Height, m_ButtonsFont
                'GetMeasureText hdc, .Caption, sRECTL.Width, sRECTL.Height, m_ButtonsFont
                If sRECTL.Width > btnWidth Then btnWidth = sRECTL.Width
            End With
        Next
        For i = 0 To UBound(udtItemsRangeButton)
            With udtItemsRangeButton(i)
                If m_ColsPicker Then
                    a = m_ColsPicker - 1
                Else
                    a = m_NumberPickers - 1
                End If
                '--
                tmph = m_ButtonsHeight / 3
                '--
                .RECT.Left = udtItemsPicker(a).RECT.Left + udtItemsPicker(a).RECT.Width + (PaddingX * 2)
                '.RECT.Top = (udtItemsPicker(0).TitleMonthYear.RECT.Top + udtItemsPicker(0).TitleMonthYear.RECT.Height) + (i * m_ButtonNavHeight) + ((m_ButtonNavHeight / 3) * i)
                .RECT.Top = PaddingY + (IIF(m_ColsPicker, udtItemsPicker(m_NumberPickers - 1).RECT.Top + udtItemsPicker(m_NumberPickers - 1).RECT.Height - PaddingY, udtItemsPicker(a).RECT.Height) / 2) - (((m_ButtonsHeight * (UBound(udtItemsRangeButton) + 1)) + (tmph * UBound(udtItemsRangeButton))) / 2) + (m_ButtonsHeight * i) + (tmph * i)
                .RECT.Width = btnWidth + (btnWidth / 2)
                .RECT.Height = m_ButtonsHeight
                .RECT2 = RectLToRect(.RECT)
            End With
        Next
    End If
    '--
    
    'Botones de accion.
    'udtItemsActionButton(2) 'Para botones de rangos ('Hoy', 'Cancelar', 'Aplicar')
    'Definir nombres en los botones de accion:
    If Not m_AutoApply Or m_ShowTodayButton Then
        btnWidth = 0
        countItem = UBound(udtItemsActionButton) - IIF(m_ShowTodayButton, 2, IIF(IsChild, 1, 0))
        For i = 0 To countItem
            With udtItemsActionButton(i)
                SetRect sRECTL, 0, 0, 0, 0
                GdiPlusGetMeasureString hGraphics, .Caption, sRECTL.Width, sRECTL.Height, m_ButtonsFont
                'GetMeasureText hdc, .Caption, sRECTL.Width, sRECTL.Height, m_ButtonsFont
                If sRECTL.Width > btnWidth Then btnWidth = (sRECTL.Width + sRECTL.Width / 2)
            End With
        Next
        For i = 0 To countItem
            With udtItemsActionButton(i)
                If m_ColsPicker Then
                    a = m_ColsPicker - 1
                Else
                    a = m_NumberPickers - 1
                End If
                tmpW = btnWidth / 10
                .RECT.Left = ((udtItemsPicker(a).RECT.Left + udtItemsPicker(a).RECT.Width + PaddingX) / 2) - (((btnWidth * (countItem + 1)) + (tmpW * 2)) / 2) + (btnWidth * i) + (tmpW * i)
                .RECT.Top = udtItemsPicker(m_NumberPickers - 1).RECT.Top + udtItemsPicker(m_NumberPickers - 1).RECT.Height + PaddingY
                .RECT.Width = btnWidth
                .RECT.Height = m_ButtonsHeight
                .RECT2 = RectLToRect(.RECT)
            End With
        Next
    End If
    
    'Ajustar tamaño de control segun los calendarios.
    With UserControl
        i = m_NumberPickers - 1
        a = 1
        If m_ColsPicker Then
            a = RoundUp(m_NumberPickers / m_ColsPicker)
            i = 0
        End If
        c_Width = (((udtItemsPicker(i).RECT.Left + udtItemsPicker(i).RECT.Width) * IIF(m_ColsPicker, m_ColsPicker, 1)) + PaddingX + IIF(m_ShowRangeButtons And m_UseRangeValue, udtItemsRangeButton(0).RECT.Width + (PaddingX * 2), 0) + IIF(m_CallOutPosition = [Position Right] And Not m_IsChild, coWidth, 0)) * Screen.TwipsPerPixelX
        c_Height = (((udtItemsPicker(0).RECT.Top + udtItemsPicker(0).RECT.Height) * a) + PaddingY + IIF(Not m_AutoApply Or m_ShowTodayButton, udtItemsActionButton(0).RECT.Height + PaddingY, 0) + IIF(m_CallOutPosition = [Position Bottom] And Not m_IsChild, coHeight, 0)) * Screen.TwipsPerPixelX
        '---
        .Size c_Width, c_Height
    End With
    
    'Cacular posicion y tamaños de los objetos dentro del calendario.
    'Dias de la semana (Do, Lu, Ma, Mi, Ju, Vi, Sa):
    For i = 0 To UBound(udtItemsPicker)
        For a = (cs_ColsDay * i) To (cs_ColsDay * (i + 1)) - 1
            With udtItemsHeaderDay(a)
                If m_RightToLeft Then
                    .RECT.Left = (udtItemsPicker(i).RECT.Left + udtItemsPicker(i).RECT.Width) - ((m_WeekWidth * Abs(m_ShowNumberWeek)) + ((m_DayWidth + m_SpaceGrid) * (a Mod cs_ColsDay)))
                Else
                    .RECT.Left = udtItemsPicker(i).RECT.Left + (m_SpaceGrid + ((m_WeekWidth + m_SpaceGrid) * Abs(m_ShowNumberWeek)) + ((m_DayWidth + m_SpaceGrid) * (a Mod cs_ColsDay)))
                End If
                .RECT.Top = (udtItemsPicker(i).TitleMonthYear.RECT.Top + udtItemsPicker(i).TitleMonthYear.RECT.Height) + PaddingY
                .RECT.Width = m_DayWidth
                .RECT.Height = m_DayHeight
                .DayName = StrConv(WeekdayName(Weekday((a Mod cs_ColsDay) + m_UserFirstDayOfWeek, vbSunday), False, vbSunday), vbProperCase)
                .Caption = Left(.DayName, 2)
                .IndexCalendar = i
            End With
        Next
    Next

    'Los dias del mes:
    For i = 0 To UBound(udtItemsPicker)
        '---
        If m_LinkedCalendars Then
            If Not IsDate(m_ValueStart) Then dDate(i) = DateAdd("m", i, DateSerial(Year(m_Value), Month(m_Value), 1))
            If IsDate(m_ValueStart) Then dDate(i) = DateAdd("m", i, DateSerial(Year(m_ValueStart), Month(m_ValueStart), 1))
        Else
            If i = 0 Then
                If Not IsDate(m_ValueStart) Then dDate(i) = DateAdd("m", i, DateSerial(Year(m_Value), Month(m_Value), 1))
                If IsDate(m_ValueStart) Then dDate(i) = DateAdd("m", i, DateSerial(Year(m_ValueStart), Month(m_ValueStart), 1))
            Else
                If Not IsDate(m_ValueEnd) Then dDate(i) = DateAdd("m", i, DateSerial(Year(m_Value), Month(m_Value), 1))
                If IsDate(m_ValueStart) And IsDate(m_ValueEnd) Then dDate(i) = DateAdd("m", IIF(Month(m_ValueStart) <> Month(m_ValueEnd), 0, i), DateSerial(Year(m_ValueEnd), Month(m_ValueEnd), 1))
            End If
        End If
        '---
        With udtItemsPicker(i)
            .NumberMonth = Month(dDate(i))
            .NumberYear = Year(dDate(i))
            .MonthName = MonthName(.NumberMonth)
            .HeaderTitle = .MonthName & " " & .NumberYear
            .DateInPicker = DateSerial(.NumberYear, .NumberMonth, 1)
        End With
        '---
        jump = 0
        For a = (cs_ItemsDay * i) To (cs_ItemsDay * (i + 1)) - 1
            '---
            With udtItemsDay(a)
                If a Mod cs_ColsDay = 0 Then jump = jump + 1
                '---
                If m_RightToLeft Then
                    .RECT.Left = (udtItemsPicker(i).RECT.Left + udtItemsPicker(i).RECT.Width) - ((m_WeekWidth * Abs(m_ShowNumberWeek)) + ((m_DayWidth + m_SpaceGrid) * (a Mod cs_ColsDay)))
                Else
                    .RECT.Left = udtItemsPicker(i).RECT.Left + (m_SpaceGrid + ((m_WeekWidth + m_SpaceGrid) * Abs(m_ShowNumberWeek)) + ((m_DayWidth + m_SpaceGrid) * (a Mod cs_ColsDay)))
                End If
                .RECT.Top = (udtItemsPicker(i).TitleMonthYear.RECT.Top + udtItemsPicker(i).TitleMonthYear.RECT.Height) + PaddingY + ((m_DayHeight + m_SpaceGrid) * jump)
                .RECT.Width = m_DayWidth
                .RECT.Height = m_DayHeight
                .IndexCalendar = i
                .RECT2 = RectLToRect(.RECT)
            End With
            '---
        Next
        '---
    Next

    'Cabecera y números de las semanas:
    For i = 0 To UBound(udtItemsPicker)
        For a = (cs_RowsPicker * i) To (cs_RowsPicker * (i + 1)) - 1
            With udtItemsWeek(a)
                .RECT.Left = udtItemsPicker(i).RECT.Left + m_SpaceGrid
                .RECT.Top = (udtItemsPicker(i).TitleMonthYear.RECT.Top + udtItemsPicker(i).TitleMonthYear.RECT.Height) + PaddingY + ((m_WeekHeight + m_SpaceGrid) * (a Mod cs_RowsPicker))
                .RECT.Width = m_WeekWidth
                .RECT.Height = m_WeekHeight
                .IndexCalendar = i
            End With
        Next
    Next
    '---
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
    '---
End Sub

'Calcular posicion y tamaño de los items para el grid de los meses y años de navegacion.
Private Sub ChangeViewPicker()
    Dim i As Integer
    Dim J As Integer
    Dim jump As Integer
    Dim areaRect As RECTL
    Dim startYear As Integer
    '---
    jump = 0
    '---
    For i = 0 To UBound(udtItemsPicker)
        If udtItemsPicker(i).ViewNavigator <> ViewItemNavigatorDays Then
            b_ShowFastNavigator = True
            '--
            With areaRect
                .Left = udtItemsPicker(i).RECT.Left
                .Top = (udtItemsPicker(i).TitleMonthYear.RECT.Top + udtItemsPicker(i).TitleMonthYear.RECT.Height) + PaddingY
                .Width = udtItemsPicker(i).RECT.Width
                .Height = udtItemsPicker(i).RECT.Height - (m_SpaceGrid + udtItemsPicker(i).TitleMonthYear.RECT.Height + m_PaddingY)
            End With
            '---
            If udtItemsPicker(i).ViewNavigator = ViewItemNavigatorYears Then startYear = (Fix(Val(udtItemsPicker(i).NumberYear) / 10) * 10) - 1
            For J = 0 To UBound(udtItemsMonthYear) - 1
                If J > 0 And J Mod cs_ColsMonthYear = 0 Then jump = jump + 1
                With udtItemsMonthYear(J)
                    .RECT.Left = areaRect.Left + m_SpaceGrid + (Fix(areaRect.Width / cs_ColsMonthYear) * (J Mod cs_ColsMonthYear))
                    .RECT.Top = areaRect.Top + m_SpaceGrid + Fix(areaRect.Height / cs_RowsMonthYear) * jump
                    .RECT.Width = Fix(areaRect.Width / cs_ColsMonthYear) - m_SpaceGrid
                    .RECT.Height = Fix(areaRect.Height / cs_RowsMonthYear) - m_SpaceGrid
                    If udtItemsPicker(i).ViewNavigator = ViewItemNavigatorMonths Then
                        .Caption = MonthName(J + 1, True) & "."
                        .ValueItem = J + 1
                    ElseIf udtItemsPicker(i).ViewNavigator = ViewItemNavigatorYears Then
                        .ValueItem = startYear
                        .Caption = CStr(.ValueItem)
                        startYear = startYear + 1
                    End If
                    .IndexCalendar = i
                    .RECT2 = RectLToRect(.RECT)
                End With
            Next
            '---
            For J = 0 To UBound(udtItemsUpDownButton)
                With udtItemsUpDownButton(J)
                    .RECT.Left = IIF(J Mod 2 = 0, udtItemsPicker(i).TitleMonthYear.RECT.Left, udtItemsPicker(i).TitleMonthYear.RECT.Left + (udtItemsPicker(i).TitleMonthYear.RECT.Width - (udtItemsPicker(i).TitleMonthYear.RECT.Width / 6) - 2 * nScale))
                    .RECT.Top = udtItemsPicker(i).TitleMonthYear.RECT.Top
                    .RECT.Width = (udtItemsPicker(i).TitleMonthYear.RECT.Width / 6)
                    .RECT.Height = udtItemsPicker(i).TitleMonthYear.RECT.Height
                    .RECT2 = RectLToRect(.RECT)
                End With
            Next
            Exit For
        End If
    Next
    '---
End Sub

Private Sub UpdateLinkedCalendar(Index As Integer)
    Dim i As Integer
    Dim dDate As Date
    '--
    If Index > 0 Then
        dDate = DateSerial(udtItemsPicker(Index).NumberYear, udtItemsPicker(Index).NumberMonth, 1)
        With udtItemsPicker(0)
            dDate = DateAdd("m", -(Index), dDate)
            .NumberYear = Year(dDate)
            .NumberMonth = Month(dDate)
            .MonthName = MonthName(Month(dDate))
            .DateInPicker = DateSerial(.NumberYear, .NumberMonth, 1)
        End With
    End If
    '--
    dDate = DateSerial(udtItemsPicker(0).NumberYear, udtItemsPicker(0).NumberMonth, 1)
    For i = 1 To UBound(udtItemsPicker)
        dDate = DateAdd("m", 1, dDate)
        With udtItemsPicker(i)
            .NumberYear = Year(dDate)
            .NumberMonth = Month(dDate)
            .MonthName = MonthName(Month(dDate))
            .DateInPicker = DateSerial(.NumberYear, .NumberMonth, 1)
        End With
    Next
    '---
End Sub

Private Sub UpdateScaleDPI()
    'Escalar a DPI las propiedades de medidas.
    If Ambient.UserMode Then
        'Paddings
        m_PaddingX = m_PaddingX * nScale
        m_PaddingY = m_PaddingY * nScale
        
        'Espacios
        m_SpaceGrid = m_SpaceGrid * nScale
        
        'Botones
        m_ButtonNavWidth = m_ButtonNavWidth * nScale
        m_ButtonNavHeight = m_ButtonNavHeight * nScale
        m_ButtonNavBorderWidth = m_ButtonNavBorderWidth * nScale
        m_ButtonNavCornerRadius = m_ButtonNavCornerRadius * nScale
        m_ButtonsWidth = m_ButtonsWidth * nScale
        m_ButtonsHeight = m_ButtonsHeight * nScale
        m_ButtonsBorderWidth = m_ButtonsBorderWidth * nScale
        m_ButtonsCornerRadius = m_ButtonsCornerRadius * nScale
        
        'Titulos del calendario
        m_MonthYearBorderWidth = m_MonthYearBorderWidth * nScale
        m_MonthYearCornerRadius = m_MonthYearCornerRadius * nScale
        
        'Días
        m_DayWidth = m_DayWidth * nScale
        m_DayHeight = m_DayHeight * nScale
        m_DayBorderWidth = m_DayBorderWidth * nScale
        m_DayCornerRadius = m_DayCornerRadius * nScale
        
        'Semanas
        m_WeekWidth = m_WeekWidth * nScale
        m_WeekHeight = m_WeekHeight * nScale
        m_WeekBorderWidth = m_WeekBorderWidth * nScale
        m_WeekCornerRadius = m_WeekCornerRadius * nScale
        '---
    End If
End Sub

Private Sub ResetControl()
    Dim i, a As Integer
    'Reset cursor
    'ShowHandPointer False
    'Mes y año
    For i = 0 To UBound(udtItemsPicker)
        udtItemsPicker(i).TitleMonthYear.MouseState = Normal
    Next
    'Dias
    For i = 0 To UBound(udtItemsPicker)
        If udtItemsPicker(i).ViewNavigator = ViewItemNavigatorDays Then
            For a = 0 To UBound(udtItemsDay)
                udtItemsDay(a).MouseState = Normal
            Next
        Else
            For a = 0 To UBound(udtItemsMonthYear)
                udtItemsMonthYear(a).MouseState = Normal
            Next
        End If
    Next
    'Botones
    '-Ranges
    If m_ShowRangeButtons Then
        For i = 0 To UBound(udtItemsRangeButton)
            udtItemsRangeButton(i).MouseState = Normal
        Next
    End If
    '-AutoApply
    If Not m_AutoApply Or m_ShowTodayButton Then
        For i = 0 To UBound(udtItemsActionButton)
            udtItemsActionButton(i).MouseState = Normal
        Next
    End If
End Sub

Private Sub Draw()
    Dim hGraphics       As Long
    Dim i               As Integer
    Dim a               As Integer
    Dim countItem       As Integer
    Dim lRect           As RECTL
    Dim Corners         As Radius
    Dim IsBold          As Boolean
    Dim BackColor       As OLE_COLOR
    Dim TextColor       As OLE_COLOR
    Dim BorderWidth     As Integer
    Dim BorderColor     As OLE_COLOR
    Dim ArrowColor      As OLE_COLOR
    Dim IsLock          As Boolean
    '--
    Dim dDate           As Date
    Dim FirtsDay        As Integer
    Dim curDate         As Date
    '---
    Cls
    '---
    If hGraphics = 0 Then GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipTranslateWorldTransform hGraphics, 0, 0, &H1
    
    '-> Pintar ParentBackColor
    If m_IsChild Then
        SetRect lRect, -1, -1, UserControl.ScaleWidth + 2, UserControl.ScaleHeight + 2
        DrawRoundRect hGraphics, lRect, Corners, m_BackColorParent
    End If
    
    '-> Pintar area del control
    With Corners
        .TopLeft = IIF(m_IsChild, m_CornerTopLeft, 0)
        .TopRight = IIF(m_IsChild, m_CornerTopRight, 0)
        .BottomLeft = IIF(m_IsChild, m_CornerBottomLeft, 0)
        .BottomRight = IIF(m_IsChild, m_CornerBottomRight, 0)
    End With
    '---
    SetRect lRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    DrawRoundRect hGraphics, lRect, Corners, m_BackColor, IIF(m_IsChild, IIF(m_Border, m_BorderWidth, 0), 0), m_BorderColor

    '-> Pintar los meses y años.
    With Corners
        .TopLeft = m_MonthYearCornerRadius
        .TopRight = m_MonthYearCornerRadius
        .BottomLeft = m_MonthYearCornerRadius
        .BottomRight = m_MonthYearCornerRadius
    End With
    '--
    For i = 0 To UBound(udtItemsPicker)
        With udtItemsPicker(i)
            '--Pintar area del picker - pruebas
            'SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
            'DrawRoundRect hGraphics, lRect, Corners, vbCyan
            '--
            If .TitleMonthYear.MouseState = Normal Then BackColor = m_MonthYearBackColor
            If .TitleMonthYear.MouseState = Hot Then BackColor = ShiftColor(vbBlack, m_MonthYearBackColor, 20)
            If .TitleMonthYear.MouseState = Inherint Then BackColor = ShiftColor(vbBlack, m_MonthYearBackColor, 20)
            If .TitleMonthYear.MouseState = Pressed Then BackColor = ShiftColor(vbBlack, m_MonthYearBackColor, 50)
            '--
            If .ViewNavigator = ViewItemNavigatorDays Then .HeaderTitle = .MonthName & " " & Year(.DateInPicker)
            If .ViewNavigator = ViewItemNavigatorMonths Then .HeaderTitle = .NumberYear
            If .ViewNavigator = ViewItemNavigatorYears Then .HeaderTitle = udtItemsMonthYear(1).ValueItem & " - " & udtItemsMonthYear(10).ValueItem
            '--
            SetRect lRect, .TitleMonthYear.RECT.Left, .TitleMonthYear.RECT.Top, .TitleMonthYear.RECT.Width, .TitleMonthYear.RECT.Height
            DrawRoundRect hGraphics, lRect, Corners, BackColor, m_MonthYearBorderWidth, m_MonthYearBorderColor
            If m_UseGDIPString Then GdiPlusDrawString hGraphics, .HeaderTitle, .TitleMonthYear.RECT.Left, .TitleMonthYear.RECT.Top, .TitleMonthYear.RECT.Width, .TitleMonthYear.RECT.Height, m_MonthYearFont, m_MonthYearForeColor, StringAlignmentCenter, StringAlignmentCenter
            If Not m_UseGDIPString Then DrawText hdc, .HeaderTitle, .TitleMonthYear.RECT2.Left, .TitleMonthYear.RECT2.Top, .TitleMonthYear.RECT2.Right, .TitleMonthYear.RECT2.Bottom, m_MonthYearFont, m_MonthYearForeColor, StringAlignmentCenter, StringAlignmentCenter
            '---
            If .ViewNavigator <> ViewItemNavigatorDays And m_LinkedCalendars Then
                For a = 0 To UBound(udtItemsUpDownButton)
                    With udtItemsUpDownButton(a)
                        SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
                        If .MouseState = Normal Then ArrowColor = ShiftColor(vbBlack, m_MonthYearBackColor, 90)
                        If .MouseState = Hot Then ArrowColor = ShiftColor(vbWhite, m_DaySelValuesColor, 50)
                        If .MouseState = Pressed Then ArrowColor = ShiftColor(vbBlack, m_DaySelValuesColor, 80)
                        '---
                        Call DrawArrow(hGraphics, CSng(.RECT.Left), CSng(.RECT.Top), CSng(.RECT.Height), IIF(a Mod 2 = 0, ArrowDirectionDown, ArrowDirectionUp), ArrowColor)
                    End With
                Next
            End If
        End With
    Next

    '-> Pintar seccion de botones del calendario.
    '---
    With Corners
        .TopLeft = m_ButtonNavCornerRadius: .TopRight = m_ButtonNavCornerRadius
        .BottomLeft = m_ButtonNavCornerRadius: .BottomRight = m_ButtonNavCornerRadius
    End With
    '---
    For i = 0 To UBound(udtItemsNavButton)
        With udtItemsNavButton(i)
            '---
            If .MouseState = Normal Then BackColor = m_ButtonNavBackColor
            If .MouseState = Hot Then BackColor = ShiftColor(vbBlack, m_ButtonNavBackColor, 20)
            If .MouseState = Pressed Then BackColor = ShiftColor(vbBlack, m_ButtonNavBackColor, 50)
            ArrowColor = m_ButtonNavForeColor
            '---
            SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
            DrawRoundRect hGraphics, lRect, Corners, BackColor, m_ButtonNavBorderWidth, m_ButtonNavBorderColor
            '---
            Call DrawArrow(hGraphics, CSng(.RECT.Left), CSng(.RECT.Top), CSng(.RECT.Height), IIF(i Mod 2 = 0, ArrowDirectionLeft, ArrowDirectionRight), ArrowColor)
        End With
    Next

    '-> Pintar cabecera de los días.
    With Corners
        .TopLeft = 0: .TopRight = 0: .BottomLeft = 0: .BottomRight = 0
    End With
    '---
    For i = 0 To UBound(udtItemsHeaderDay)
        With udtItemsHeaderDay(i)
            'Pintar solo si el viewnavigator esta en días del mes
            If udtItemsPicker(.IndexCalendar).ViewNavigator = ViewItemNavigatorDays Then
                SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
                DrawRoundRect hGraphics, lRect, Corners, m_DayBackColor, m_DayBorderWidth, m_DayBorderColor
                If m_UseGDIPString Then GdiPlusDrawString hGraphics, .Caption, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height, m_DayFont, m_DayHeaderForeColor, StringAlignmentCenter, StringAlignmentCenter, m_DayHeaderFontBold
                If Not m_UseGDIPString Then DrawText hdc, .Caption, .RECT.Left, .RECT.Top, .RECT.Left + .RECT.Width, .RECT.Top + .RECT.Height, m_DayFont, m_DayHeaderForeColor, StringAlignmentCenter, StringAlignmentCenter, m_DayHeaderFontBold
            End If
        End With
    Next

    '-> Pintar los días del mes.
    Dim DayPre As Date, ColorPre As Long
    For i = 0 To UBound(udtItemsPicker)
        'For a = 0 To UBound(udtItemsDay)
        dDate = DateSerial(udtItemsPicker(i).NumberYear, udtItemsPicker(i).NumberMonth, 1)
        FirtsDay = Weekday(dDate, m_UserFirstDayOfWeek)
        If udtItemsPicker(i).ViewNavigator = ViewItemNavigatorDays Then
            '---
            For a = (cs_ItemsDay * i) To (cs_ItemsDay * (i + 1)) - 1
                With udtItemsDay(a)
                    '---
                    curDate = DateAdd("d", -FirtsDay + ((a + 1) - ((cs_ItemsDay * i))), dDate)
                    '--
                    .DateValue = curDate
                    .DatePartDay = Day(.DateValue)
                    .DatePartMonth = Month(.DateValue)
                    .DatePartYear = Year(.DateValue)
                    '--
                    .Caption = .DatePartDay
                    .NumberWeek = DatePart("ww", curDate, IIF(m_ShowUseISOWeek, vbMonday, m_UserFirstDayOfWeek), IIF(m_ShowUseISOWeek, vbFirstFourDays, vbUseSystem))
                    '--
                    IsLock = .DateValue < m_MinDate Or .DateValue > m_MaxDate
                    .IsDayInMonthCurrent = Month(dDate) = .DatePartMonth
                    .IsNow = curDate = Date
                    .IsValueDate = curDate = d_ValueTemp
                    If IsDate(m_ValueStart) Then .IsStartDate = curDate = m_ValueStart
                    If Not IsDate(m_ValueStart) Then .IsStartDate = False
                    If IsDate(m_ValueEnd) Then .IsEndDate = curDate = m_ValueEnd
                    If Not IsDate(m_ValueEnd) Then .IsEndDate = False
                    '-> Between
                    '--
                    If IsDate(m_ValueStart) And IsDate(m_ValueEnd) Then
                        .IsBetweenDate = (curDate > m_ValueStart And curDate < m_ValueEnd)
                    ElseIf IsDate(m_ValueStart) And Not (IsDate(m_ValueEnd)) And c_IndexSelMove > -1 Then
                        If curDate > CDate(m_ValueStart) And curDate < udtItemsDay(c_IndexSelMove).DateValue Then
                            .IsBetweenDate = True
                        Else
                            .IsBetweenDate = False
                        End If
                    Else
                        .IsBetweenDate = False
                    End If
                    '-> WeekEnd (Sabado y Domingo)
                    .IsDaySaturday = Weekday(curDate) = vbSaturday
                    .IsDaySunday = Weekday(curDate) = vbSunday
                    '---
                    BorderWidth = m_DayBorderWidth
                    BorderColor = m_DayBorderColor
                    '---
                    With Corners
                        .TopLeft = m_DayCornerRadius: .TopRight = m_DayCornerRadius: .BottomLeft = m_DayCornerRadius: .BottomRight = m_DayCornerRadius
                    End With
                    '---
                    'If .MouseState = Hot Then BackColor = ShiftColor(m_BackColor, m_DaySelValuesColor, 200) 'm_DayHotColor
                    If .MouseState = Hot Then BackColor = m_DayHotColor
                    If .MouseState = Normal Then BackColor = m_DayBackColor
                    '---
                    IsBold = False
                    '---
                    If .IsDayInMonthCurrent Then
                        TextColor = m_DayForeColor
                        If .IsDaySaturday Then TextColor = m_DaySaturdayForeColor
                        If .IsDaySunday Then TextColor = m_DaySundayForeColor
                        If .IsFreeDay Then TextColor = m_DayFreeForeColor
                        If IsLock Then TextColor = vbGrayText
                        If .IsNow And m_DayNowShow Then
                            BackColor = m_DayNowBackColor
                            If Not .IsDaySaturday And Not .IsDaySunday And Not .IsFreeDay And Not IsLock Then TextColor = m_DayNowForeColor
                            BorderWidth = m_DayNowBorderWidth: BorderColor = m_DayNowBorderColor
                        End If
                        'Start / End Date Selection
                        If .IsStartDate Or .IsEndDate Then
                            BackColor = m_DaySelValuesColor
                            If Not .IsFreeDay And Not IsLock Then TextColor = m_DaySelForeColor
                            IsBold = m_DaySelFontBold
                        End If
                        'Between date Selection
                        If .IsBetweenDate Then
                            BackColor = m_DaySelBetweenColor
                        End If
                        If .IsValueDate And (Not m_UseRangeValue) Then
                            BackColor = m_DaySelValuesColor
                            If Not .IsFreeDay And Not IsLock Then TextColor = m_DaySelForeColor
                            IsBold = m_DaySelFontBold
                        End If
                        If .IsBetweenDate Then
                            BackColor = m_DaySelBetweenColor
                            If Not .IsDaySaturday And Not .IsDaySunday And Not .IsFreeDay And Not IsLock Then TextColor = m_DayForeColor
                        End If
                        If .IsStartDate Or .IsBetweenDate Or .IsEndDate Or .IsValueDate Then
                            '---
                            If m_DaySelectionStyle = [Corner Full] Then
                                Corners.TopLeft = (m_DayWidth / 8): Corners.TopRight = (m_DayWidth / 8)
                                Corners.BottomLeft = (m_DayWidth / 8): Corners.BottomRight = (m_DayWidth / 8)
                            ElseIf m_DaySelectionStyle = [Corner No Between] Then
                                If m_RightToLeft Then
                                    Corners.TopLeft = IIF(.IsEndDate, (m_DayWidth / 8), (IIF(.IsStartDate, 0, IIF(.IsBetweenDate, 0, m_DayCornerRadius))))
                                    Corners.TopRight = IIF(.IsStartDate, (m_DayWidth / 8), (IIF(.IsEndDate, 0, IIF(.IsBetweenDate, 0, m_DayCornerRadius))))
                                    Corners.BottomLeft = IIF(.IsEndDate, (m_DayWidth / 8), (IIF(.IsStartDate, 0, IIF(.IsBetweenDate, 0, m_DayCornerRadius))))
                                    Corners.BottomRight = IIF(.IsStartDate, (m_DayWidth / 8), (IIF(.IsEndDate, 0, IIF(.IsBetweenDate, 0, m_DayCornerRadius))))
                                Else
                                    Corners.TopLeft = IIF(.IsStartDate, (m_DayWidth / 8), (IIF(.IsEndDate, 0, IIF(.IsBetweenDate, 0, m_DayCornerRadius))))
                                    Corners.TopRight = IIF(.IsEndDate, (m_DayWidth / 8), (IIF(.IsStartDate, 0, IIF(.IsBetweenDate, 0, m_DayCornerRadius))))
                                    Corners.BottomLeft = IIF(.IsStartDate, (m_DayWidth / 8), (IIF(.IsEndDate, 0, IIF(.IsBetweenDate, 0, m_DayCornerRadius))))
                                    Corners.BottomRight = IIF(.IsEndDate, (m_DayWidth / 8), (IIF(.IsStartDate, 0, IIF(.IsBetweenDate, 0, m_DayCornerRadius))))
                                End If
                            End If
                        End If
                        '---
                        RaiseEvent DayPrePaint(curDate, ColorPre)
                        If ColorPre <> 0 Then
                            m_DaysPrePaintCount = 0
                            BackColor = ColorPre
                            m_DaysPrePaintCount = m_DaysPrePaintCount + 1
                        End If
                    Else
                        TextColor = m_DayOMForeColor
                    End If
                    '---
                    SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
                    DrawRoundRect hGraphics, lRect, Corners, BackColor, BorderWidth, BorderColor, IsLock
                    If m_UseGDIPString Then GdiPlusDrawString hGraphics, .Caption, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height, m_DayFont, TextColor, StringAlignmentCenter, StringAlignmentCenter, IsBold
                    If Not m_UseGDIPString Then DrawText hdc, .Caption, .RECT2.Left, .RECT2.Top, .RECT2.Right, .RECT2.Bottom, m_DayFont, TextColor, StringAlignmentCenter, StringAlignmentCenter, IsBold
                    '---
                    If ColorPre <> 0 Then ColorPre = 0
                End With
            Next
            '---
        End If
    Next
    
    '-> Pintar columna de los numeros de la semana.
    With Corners
        .TopLeft = m_WeekCornerRadius: .TopRight = m_WeekCornerRadius
        .BottomLeft = m_WeekCornerRadius: .BottomRight = m_WeekCornerRadius
    End With
    '---
    If m_ShowNumberWeek Then
        For i = 0 To UBound(udtItemsPicker)
            If udtItemsPicker(i).ViewNavigator = ViewItemNavigatorDays Then
                For a = (cs_RowsPicker * i) To (cs_RowsPicker * (i + 1)) - 1
                    With udtItemsWeek(a)
                        If a Mod cs_RowsPicker > 0 Then
                            .NumberWeek = udtItemsDay(((a - i) * cs_ColsDay) - 1).NumberWeek
                        End If
                        IsBold = IIF(a Mod cs_RowsPicker = 0, True, False)
                        .Caption = IIF(a Mod cs_RowsPicker = 0, "#", .NumberWeek)
                        '---
                        SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
                        DrawRoundRect hGraphics, lRect, Corners, m_WeekBackColor, m_WeekBorderWidth, m_WeekBorderColor
                        If m_UseGDIPString Then GdiPlusDrawString hGraphics, .Caption, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height, m_WeekFont, Me.WeekForeColor, StringAlignmentCenter, StringAlignmentCenter, IsBold
                        If Not m_UseGDIPString Then DrawText hdc, .Caption, .RECT.Left, .RECT.Top, .RECT.Left + .RECT.Width, .RECT.Top + .RECT.Height, m_WeekFont, Me.WeekForeColor, StringAlignmentCenter, StringAlignmentCenter, IsBold
                    End With
                Next
            End If
        Next
    End If

    '-> Pintar navegacion rápida de meses o años.
    With Corners
        .TopLeft = udtItemsMonthYear(0).RECT.Width / 8
        .TopRight = udtItemsMonthYear(0).RECT.Width / 8
        .BottomLeft = udtItemsMonthYear(0).RECT.Width / 8
        .BottomRight = udtItemsMonthYear(0).RECT.Width / 8
    End With
    For i = 0 To UBound(udtItemsPicker)
        If udtItemsPicker(i).ViewNavigator <> ViewItemNavigatorDays Then
            For a = 0 To UBound(udtItemsMonthYear)
                With udtItemsMonthYear(a)
                    '---
                    If udtItemsPicker(i).ViewNavigator = ViewItemNavigatorMonths Then
                        IsLock = DateSerial(udtItemsPicker(i).NumberYear, .ValueItem, 1) < m_MinDate Or DateSerial(udtItemsPicker(i).NumberYear, .ValueItem, 1) > m_MaxDate
                    Else
                        IsLock = .ValueItem < Year(m_MinDate) Or .ValueItem > Year(m_MaxDate)
                    End If
                    '---
                    If (Month(udtItemsPicker(i).DateInPicker) & Year(udtItemsPicker(i).DateInPicker) = .ValueItem & udtItemsPicker(i).NumberYear) Or (Year(udtItemsPicker(i).DateInPicker) = .ValueItem) Then
                        'Debug.Print "i = " & i
                        BackColor = m_DaySelValuesColor
                        TextColor = m_DaySelForeColor
                    Else
                        If .MouseState = Normal Then BackColor = m_DayBackColor
                        If .MouseState = Hot Then BackColor = ShiftColor(vbWhite, m_DaySelValuesColor, 220)
                        If .MouseState = Pressed Then BackColor = ShiftColor(vbWhite, m_DaySelValuesColor, 120)
                        '---
                        If (a = 0 Or a = (cs_ColsMonthYear * cs_RowsMonthYear) - 1) And udtItemsPicker(i).ViewNavigator = ViewItemNavigatorYears Then
                            TextColor = m_DayOMForeColor
                        Else
                            TextColor = m_DayForeColor
                        End If
                    End If
                    '---
                    SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
                    DrawRoundRect hGraphics, lRect, Corners, BackColor, m_DayBorderWidth, m_DayBorderColor, IsLock
                    If m_UseGDIPString Then GdiPlusDrawString hGraphics, .Caption, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height, m_DayFont, TextColor, StringAlignmentCenter, StringAlignmentCenter
                    If Not m_UseGDIPString Then DrawText hdc, .Caption, .RECT2.Left, .RECT2.Top, .RECT2.Right, .RECT2.Bottom, m_DayFont, TextColor, StringAlignmentCenter, StringAlignmentCenter
                End With
            Next
        End If
    Next

    '-> Pintar Botones de rangos.
    If m_ShowRangeButtons And m_UseRangeValue Then
        With Corners
            .TopLeft = m_ButtonsCornerRadius: .TopRight = m_ButtonsCornerRadius
            .BottomLeft = m_ButtonsCornerRadius: .BottomRight = m_ButtonsCornerRadius
        End With
        '-> Pintar linea gradiente de separacion
        a = IIF(m_ColsPicker, m_ColsPicker - 1, m_NumberPickers - 1)
        '--
        DrawLineGradient hdc, udtItemsPicker(a).RECT.Left + udtItemsPicker(a).RECT.Width + m_PaddingX _
                            , udtItemsPicker(0).RECT.Top _
                            , m_BorderWidth * nScale _
                            , IIF(m_ColsPicker, udtItemsPicker(UBound(udtItemsPicker)).RECT.Top + udtItemsPicker(UBound(udtItemsPicker)).RECT.Height - m_PaddingY, udtItemsPicker(a).RECT.Height) _
                            , RGBtoARGB(m_BackColor, 100) _
                            , RGBtoARGB(m_BorderColor, 100) _
                            , True, True
        '--
        For i = 0 To UBound(udtItemsRangeButton)
            With udtItemsRangeButton(i)
                If .MouseState = Normal Then BackColor = m_ButtonsBackColor
                If .MouseState = Hot Then BackColor = ShiftColor(vbBlack, m_ButtonsBackColor, 20)
                If .MouseState = Pressed Then BackColor = ShiftColor(vbBlack, m_ButtonsBackColor, 50)
                '--
                SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
                DrawRoundRect hGraphics, lRect, Corners, BackColor, m_ButtonsBorderWidth, m_ButtonsBorderColor
                If m_UseGDIPString Then GdiPlusDrawString hGraphics, .Caption, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height, m_ButtonsFont, m_ButtonsForeColor, StringAlignmentCenter, StringAlignmentCenter
                If Not m_UseGDIPString Then DrawText hdc, .Caption, .RECT2.Left, .RECT2.Top, .RECT2.Right, .RECT2.Bottom, m_ButtonsFont, m_ButtonsForeColor, StringAlignmentCenter, StringAlignmentCenter
            End With
        Next
    End If

    '-> Pintar Botones de accion.
    If Not m_AutoApply Or m_ShowTodayButton Then
        countItem = UBound(udtItemsActionButton) - IIF(m_ShowTodayButton, 2, IIF(IsChild, 1, 0))
        With Corners
            .TopLeft = m_ButtonsCornerRadius: .TopRight = m_ButtonsCornerRadius
            .BottomLeft = m_ButtonsCornerRadius: .BottomRight = m_ButtonsCornerRadius
        End With
        '-> Pintar linea gradiente de separacion
        a = IIF(m_ColsPicker, m_ColsPicker - 1, m_NumberPickers - 1)
        '--
        DrawLineGradient hdc _
                        , CLng(m_PaddingX) _
                        , udtItemsActionButton(0).RECT.Top - (m_PaddingY / 2) _
                        , udtItemsPicker(a).RECT.Left + udtItemsPicker(a).RECT.Width - m_PaddingX _
                        , CLng(m_BorderWidth) _
                        , RGBtoARGB(m_BackColor, 100) _
                        , RGBtoARGB(m_BorderColor, 100) _
                        , False, True
        '--
        For i = 0 To countItem
            With udtItemsActionButton(i)
                '--
                If .ButtonAction = [Action Today] Then BackColor = m_ButtonsBackColor: TextColor = m_ButtonsForeColor: BorderColor = m_ButtonsBorderColor
                If .ButtonAction = [Action Cancel] Then BackColor = m_ButtonsBorderColor: TextColor = m_ButtonsForeColor: BorderColor = m_ButtonsBorderColor
                If .ButtonAction = [Action Apply] Then BackColor = m_DaySelValuesColor: TextColor = m_DaySelForeColor: BorderColor = m_DaySelValuesColor
                '--
                If .MouseState = Hot Then BackColor = ShiftColor(vbBlack, BackColor, 20)
                If .MouseState = Pressed Then BackColor = ShiftColor(vbBlack, BackColor, 50)
                '--
                SetRect lRect, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height
                DrawRoundRect hGraphics, lRect, Corners, BackColor, m_ButtonsBorderWidth, BorderColor
                If m_UseGDIPString Then GdiPlusDrawString hGraphics, .Caption, .RECT.Left, .RECT.Top, .RECT.Width, .RECT.Height, m_ButtonsFont, TextColor, StringAlignmentCenter, StringAlignmentCenter
                If Not m_UseGDIPString Then DrawText hdc, .Caption, .RECT2.Left, .RECT2.Top, .RECT2.Right, .RECT2.Bottom, m_ButtonsFont, TextColor, StringAlignmentCenter, StringAlignmentCenter
            End With
        Next
    End If
    
    '---
    If hGraphics <> 0 Then GdipDeleteGraphics hGraphics
End Sub

Private Sub DrawRoundRect(ByVal hGraphics, CanvasRect As RECTL, Corners As Radius, BackColor As OLE_COLOR, Optional BorderWidth As Integer, Optional BorderColor As OLE_COLOR = vbActiveBorder, Optional IsLock As Boolean = False)
    Dim hPath As Long, hBrush As Long, hPen As Long
    Dim X As Long, Y As Long, lWidth As Long, lHeight As Long
    Dim XX As Long, YY As Long, WW As Long, HH As Long
    '---
    X = CanvasRect.Left + BorderWidth
    Y = CanvasRect.Top + BorderWidth
    lWidth = CanvasRect.Width - BorderWidth
    lHeight = CanvasRect.Height - BorderWidth
    
    XX = X:         YY = Y
    WW = lWidth:    HH = lHeight
    
    hPath = CreatePathRoundRect(XX, YY, WW, HH, Corners)
    
    If IsLock Then
        GdipCreateHatchBrush HatchStyleForwardDiagonal, RGBtoARGB(ShiftColor(vbBlack, BackColor, 40), 100), RGBtoARGB(BackColor, 100), hBrush
    Else
        GdipCreateSolidFill RGBtoARGB(BackColor, 100), hBrush
    End If
    
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    
    If BorderWidth > 0 Then
        GdipCreatePen1 RGBtoARGB(BorderColor, 100), BorderWidth, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePath hPath
        GdipDeletePen hPen
    End If
    
    If hPath Then GdipDeletePath hPath
    
End Sub

Private Sub DrawArrow(hGraphics As Long, Left As Single, Top As Single, SizeBox As Single, ArrowDir As enmArrowDirectionConstant, Color As OLE_COLOR)
    Dim SW As Single, SH As Single
    Dim L As Single, T As Single, Size As Single
    Dim pt(6) As POINTF
    '---
    Size = SizeBox / 3
    SW = Size:   SH = Size
    '---
    L = Left + SizeBox / 2 - SW / 2
    T = Top + SizeBox / 2 - SH / 2
    '---
    Select Case ArrowDir
        Case ArrowDirectionDown 'ToDown
            pt(0).X = L + 0 * SW:           pt(0).Y = T + 0.2 * SH
            pt(1).X = L + 0.5 * SW:         pt(1).Y = T + 0.56 * SH
            pt(2).X = L + 1 * SW:           pt(2).Y = T + 0.2 * SH
            pt(3).X = L + 1 * SW:           pt(3).Y = T + 0.44 * SH
            pt(4).X = L + 0.5 * SW:         pt(4).Y = T + 0.8 * SH
            pt(5).X = L + 0 * SW:           pt(5).Y = T + 0.44 * SH
            pt(6).X = L + 0 * SW:           pt(6).Y = T + 0.2 * SH
        Case ArrowDirectionUp 'ToUp
            pt(0).X = L + 0 * SW:           pt(0).Y = T + 0.8 * SH
            pt(1).X = L + 0.5 * SW:         pt(1).Y = T + 0.44 * SH
            pt(2).X = L + 1 * SW:           pt(2).Y = T + 0.8 * SH
            pt(3).X = L + 1 * SW:           pt(3).Y = T + 0.56 * SH
            pt(4).X = L + 0.5 * SW:         pt(4).Y = T + 0.2 * SH
            pt(5).X = L + 0 * SW:           pt(5).Y = T + 0.56 * SH
            pt(6).X = L + 0 * SW:           pt(6).Y = T + 0.8 * SH
        Case ArrowDirectionLeft 'ToLeft
            pt(0).X = L + 0.8 * SW:         pt(0).Y = T + 0
            pt(1).X = L + 0.44 * SW:        pt(1).Y = T + 0.5 * SH
            pt(2).X = L + 0.8 * SW:         pt(2).Y = T + 1 * SH
            pt(3).X = L + 0.56 * SW:        pt(3).Y = T + 1 * SH
            pt(4).X = L + 0.2 * SW:         pt(4).Y = T + 0.5 * SH
            pt(5).X = L + 0.56 * SW:        pt(5).Y = T + 0 * SH
            pt(6).X = L + 0.8 * SW:         pt(6).Y = T + 0 * SH
        Case ArrowDirectionRight 'ToRight
            pt(0).X = L + 0.2 * SW:         pt(0).Y = T + 0
            pt(1).X = L + 0.56 * SW:        pt(1).Y = T + 0.5 * SH
            pt(2).X = L + 0.2 * SW:         pt(2).Y = T + 1 * SH
            pt(3).X = L + 0.44 * SW:        pt(3).Y = T + 1 * SH
            pt(4).X = L + 0.8 * SW:         pt(4).Y = T + 0.5 * SH
            pt(5).X = L + 0.44 * SW:        pt(5).Y = T + 0 * SH
            pt(6).X = L + 0.2 * SW:         pt(6).Y = T + 0 * SH
    End Select
    '---
    FillPolygon hGraphics, RGBtoARGB(Color, 100), pt
End Sub

Private Sub DrawLineGradient(ByVal hdc As Long, X As Long, Y As Long, Width As Long, Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, Optional IsVertical As Boolean, Optional IsSeparator As Boolean)
    Dim hGraphics As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim rRECTL As RECTL
    Dim lGradient As Long
    '---
    lGradient = 20 * nScale
    '---
    If GdipCreateFromHDC(hdc, hGraphics) = 0 Then
        '--
        If IsSeparator Then
            'Linea completa
            SetRect rRECTL, X, Y, Width, Height
            GdipCreateLineBrushFromRectWithAngleI rRECTL, Color2, Color2, 90, 0, WrapModeTileFlipXY, hBrush
            GdipFillRectangleI hGraphics, hBrush, X, Y, Width, Height
            If IsVertical Then
                'Degradado superior
                SetRect rRECTL, X, Y, Width, lGradient
                GdipCreateLineBrushFromRectWithAngleI rRECTL, Color1, Color2, 90, 0, WrapModeTileFlipXY, hBrush
                GdipFillRectangleI hGraphics, hBrush, X, Y, Width, lGradient
                'Degradado inferior
                SetRect rRECTL, X, Y + (Height - lGradient), Width, lGradient
                GdipCreateLineBrushFromRectWithAngleI rRECTL, Color2, Color1, 90, 0, WrapModeTileFlipXY, hBrush
                GdipFillRectangleI hGraphics, hBrush, X, Y + (Height - lGradient), Width, lGradient
            Else
                'Degradado izquierdo
                SetRect rRECTL, X, Y, lGradient, Height
                GdipCreateLineBrushFromRectWithAngleI rRECTL, Color2, Color1, 180, 0, WrapModeTileFlipXY, hBrush
                GdipFillRectangleI hGraphics, hBrush, X, Y, lGradient, Height
                'Degradado derecho
                SetRect rRECTL, X + (Width - lGradient), Y, lGradient, Height
                GdipCreateLineBrushFromRectWithAngleI rRECTL, Color1, Color2, 180, 0, WrapModeTileFlipXY, hBrush
                GdipFillRectangleI hGraphics, hBrush, X + (Width - lGradient), Y, lGradient, Height
            End If
        Else
            GdipCreateLineBrushFromRectWithAngleI rRECTL, Color1, Color2, IIF(IsVertical, 90, 180), 0, WrapModeTileFlipXY, hBrush
            GdipFillRectangleI hGraphics, hBrush, X, Y, Width, Height
        End If
        GdipDeleteBrush hBrush
        GdipDeletePen hPen
        '--
    End If
    '---
End Sub

Private Sub FillPolygon(ByVal hGraphics As Long, ByVal Color As Long, Points() As POINTF)
    Dim hPath As Long
    Dim hBrush As Long
    '---
    If GdipCreatePath(&H0, hPath) = 0 Then
        '---
        GdipAddPathPolygon hPath, Points(0), UBound(Points) + 1
        '---
        'GdipCreateSolidFill ConvertColor(Color, 100), hBrush
        GdipCreateSolidFill Color, hBrush
        GdipFillPath hGraphics, hBrush, hPath
        '---
        GdipDeleteBrush hBrush
        GdipDeletePath hPath
    End If
    '---
End Sub

Private Sub ShowHandPointer(bolHandPointer As Boolean)
    If bolHandPointer Then
        UserControl.MousePointer = MousePointerConstants.vbCustom
        UserControl.MouseIcon = GetSystemHandCursor
    Else
        UserControl.MousePointer = MousePointerConstants.vbDefault
        UserControl.MouseIcon = Nothing
    End If
End Sub

'*************
'* Funciones *
'*************
'-> Privadas

Private Function ApplyChangeValues() As Boolean
    ApplyChangeValues = False
    'Aplicar el Value
    If Not m_UseRangeValue Then
        If IsDate(d_ValueTemp) Then m_Value = d_ValueTemp
        RaiseEvent ChangeDate(m_Value)
        ApplyChangeValues = True
    End If
    'Aplicar StartValue y EndValue
    If m_UseRangeValue Then
        If IsDate(CDate(m_ValueStart)) Then
            'm_ValueStart = d_ValueStartTemp
            RaiseEvent ChangeStartDate(m_ValueStart)
        End If
        
        If IsDate(CDate(m_ValueEnd)) Then
            'm_ValueEnd = d_ValueEndTemp
            RaiseEvent ChangeEndDate(m_ValueEnd)
            ApplyChangeValues = True
        Else
            ApplyChangeValues = False
        End If
    End If
    '--
    If ApplyChangeValues Then Call HideCalendar
End Function

Private Function RoundUp(ByVal Number As Double) As Double
    Dim temp As Double
    '--
    temp = Int(Number)
    If temp <> Number Then
        temp = temp + 1
    End If
    RoundUp = temp
End Function

Private Function RectLToRect(Value As RECTL) As RECT
    Dim newRect As RECT
    With newRect
        .Left = Value.Left
        .Top = Value.Top
        .Right = Value.Left + Value.Width
        .Bottom = Value.Top + Value.Height
    End With
    RectLToRect = newRect
End Function

'ReadValue (Autor Cobein, extraido de labelplus Leandro)
Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim i       As Long
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            ReadValue = TlsGetValue(i + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function

Private Function GetSystemHandCursor() As Picture
    Dim Pic As udtPicBmp, ipic As IPicture, GUID(0 To 3) As Long
    Dim hCur As Long
    '---
    hCur = LoadCursor(ByVal 0&, IDC_HAND)
    '---
    GUID(0) = &H7BF80980
    GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B
    GUID(3) = &HAB0C3000
    '---
    With Pic
        .Size = Len(Pic)
        .type = vbPicTypeIcon
        .hBmp = hCur
        .hPal = 0
    End With
    '---
    Call OleCreatePictureIndirect(Pic, GUID(0), 1, ipic)
    '---
    Set GetSystemHandCursor = ipic
    '---
End Function

'GetFontStyleAndSize (Extraido de labelplus Leandro)
Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long, Optional IsBold As Boolean = False)
        lFontStyle = 0
        If IsBold Or oFont.Bold Then lFontStyle = lFontStyle Or GDIPFontStyleBold
        If oFont.Italic Then lFontStyle = lFontStyle Or GDIPFontStyleItalic
        If oFont.Underline Then lFontStyle = lFontStyle Or GDIPFontStyleUnderline
        If oFont.Strikethrough Then lFontStyle = lFontStyle Or GDIPFontStyleStrikeout
        
        lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
End Function

'ConvertColor (Extraido de labelplus Leandro)
Private Function RGBtoARGB(ByVal Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory RGBtoARGB, BGRA(0), 4&
End Function

'ShiftColor (De Leandro)
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    '--
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
    '--
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
    '--
    CopyMemory ShiftColor, clrFore(0), 4
End Function

'ChrW2 (Extraido de labelplus Leandro)
Private Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

'GetSafeRound (Extraido de labelplus Leandro)
Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function CreatePathRoundRect(X As Long, Y As Long, Width As Long, Height As Long, Corner As Radius) As Long
    Dim hPath As Long
    Dim BCLT, BCRT As Integer
    Dim BCBR, BCBL As Integer
    Dim XX, YY As Long
    Dim coLen As Long
    Dim coWidth As Long
    Dim lMax As Long
    Dim coAngle  As Long
    '---
    CreatePathRoundRect = 0
    Width = Width - 1 'Antialias pixel
    Height = Height - 1 'Antialias pixel
    '---
    BCLT = GetSafeRound((Corner.TopLeft + m_BorderWidth), Width, Height)
    BCRT = GetSafeRound((Corner.TopRight + m_BorderWidth), Width, Height)
    BCBL = GetSafeRound((Corner.BottomLeft + m_BorderWidth), Width, Height)
    BCBR = GetSafeRound((Corner.BottomRight + m_BorderWidth), Width, Height)
    '---
    If GdipCreatePath(&H0, hPath) = 0 Then
        '---
        If BCLT Then GdipAddPathArcI hPath, X, Y, BCLT * 2, BCLT * 2, 180, 90
        If BCLT = 0 Then GdipAddPathLineI hPath, X, Y, X + Width - BCRT, Y
        '---
        If BCRT Then GdipAddPathArcI hPath, X + Width - BCRT * 2, Y, BCRT * 2, BCRT * 2, 270, 90
        If BCRT = 0 Then GdipAddPathLineI hPath, X + Width, Y, X + Width, Y + Height - BCBR
        '---
        If BCBR Then GdipAddPathArcI hPath, X + Width - BCBR * 2, Y + Height - BCBR * 2, BCBR * 2, BCBR * 2, 0, 90
        If BCBR = 0 Then GdipAddPathLineI hPath, X + Width, Y + Height, X + BCBL, Y + Height
        '---
        If BCBL Then GdipAddPathArcI hPath, X, Y + Height - BCBL * 2, BCBL * 2, BCBL * 2, 90, 90
        If BCBL = 0 Then GdipAddPathLineI hPath, X, Y + Height, X, Y + BCLT
        '---
        GdipClosePathFigures hPath
        CreatePathRoundRect = hPath
    End If
End Function

Private Function DrawText(ByVal hdc As Long, ByVal Text As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal oFont As StdFont, ByVal ForeColor As OLE_COLOR, HAlign As enmStringAlignment, VAlign As enmStringAlignment, Optional IsBold As Boolean = False) As Long
    Dim RECT As RECT
    Dim Alignament As Long
    '---
    With Font
        .Name = oFont.Name
        .Size = oFont.Size * 0.8
        .Bold = IIF(IsBold, IsBold, oFont.Bold)
        .Underline = oFont.Underline
        .Italic = oFont.Italic
        .Strikethrough = oFont.Strikethrough
        .Charset = oFont.Charset
        If Not IsBold Then .Weight = oFont.Weight
    End With
    UserControl.ForeColor = ForeColor
    '---
    'Horizontal
    If HAlign = StringAlignmentNear Then Alignament = DT_LEFT
    If HAlign = StringAlignmentCenter Then Alignament = DT_CENTER
    If HAlign = StringAlignmentFar Then Alignament = DT_RIGHT
    'Vertical
    If VAlign = StringAlignmentNear Then Alignament = Alignament Or DT_TOP
    If VAlign = StringAlignmentCenter Then Alignament = Alignament Or DT_VCENTER
    If VAlign = StringAlignmentFar Then Alignament = Alignament Or DT_BOTTOM
    Alignament = Alignament Or DT_SINGLELINE
    '---
    SetRect RECT, X1, Y1, X2, Y2
    DrawTextW hdc, StrPtr(Text), -1, RECT, Alignament
    '---
End Function

Private Function GetMeasureText(ByVal hdc As Long, ByVal Text As String, OutWidth As Long, OutHeight As Long, ByVal oFont As StdFont, Optional IsBold As Boolean = False)
    Dim RECT As RECT
    '---
    With Font
        .Name = oFont.Name
        .Size = oFont.Size
        .Bold = IIF(IsBold, IsBold, oFont.Bold)
        .Underline = oFont.Underline
        .Italic = oFont.Italic
        .Strikethrough = oFont.Strikethrough
        .Charset = oFont.Charset
        If Not IsBold Then .Weight = oFont.Weight
    End With
    '---
    DrawTextW hdc, StrPtr(Text), -1, RECT, &H400 Or &H20
    '---
    OutWidth = RECT.Right
    OutHeight = RECT.Bottom
    '---
End Function

Private Function GdiPlusDrawString(ByVal hGraphics As Long, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal ForeColor As OLE_COLOR, Optional HAlign As enmStringAlignment, Optional VAlign As enmStringAlignment, Optional IsBold As Boolean = False) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTS
    Dim lFontSize As Long
    Dim lFontStyle As enmGDIPFontStyle
    Dim hFont As Long
    '---
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
    End If

    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
        GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisCharacter
    End If

    GetFontStyleAndSize oFont, lFontStyle, lFontSize, IsBold

    layoutRect.Left = X: layoutRect.Top = Y
    layoutRect.Width = Width: layoutRect.Height = Height

    GdipCreateSolidFill RGBtoARGB(ForeColor, 100), hBrush

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    GdipDrawString hGraphics, StrPtr(Text), -1, hFont, layoutRect, hFormat, hBrush

    GdipDeleteFont hFont
    GdipDeleteBrush hBrush
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
    '---
End Function

Private Function GdiPlusGetMeasureString(ByVal hGraphics As Long, ByVal Text As String, OutWidth As Long, OutHeight As Long, ByVal oFont As StdFont) As Long
    Dim hFontFamily
    Dim hFormat As Long
    Dim hFont As Long
    Dim lFontSize As Long
    Dim CF As Long
    Dim LF As Long
    Dim lFontStyle As enmGDIPFontStyle
    Dim layoutRect As RECTS
    Dim outRect As RECTS
    '---
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatTrimming hFormat, StringTrimmingNone
    End If
    
    GetFontStyleAndSize oFont, lFontStyle, lFontSize
       
    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    GdipMeasureString hGraphics, StrPtr(Text), -1, hFont, layoutRect, hFormat, outRect, CF, LF
        
    OutWidth = outRect.Width
    OutHeight = outRect.Height
    
    GdipDeleteFont hFont
    GdipDeleteFontFamily hFontFamily
    '---
End Function

Private Function GetLocaleInfoAsLong(ByVal Index As enmLocaleTypes, Optional ByVal Locale As enmLCIDs = LOCALE_USER_DEFAULT) As Long
    Dim TChars As Long

    Index = Index Or LOCALE_RETURN_NUMBER
    TChars = GetLocaleInfo(Locale, Index, 0, 0)
    If TChars Then
        If TChars = 2 Then
            GetLocaleInfo Locale, Index, VarPtr(GetLocaleInfoAsLong), TChars
        Else
            Err.Raise &H80049900, , "Index is not a Long LocaleInfo"
        End If
    Else
        Err.Raise &H80049904, , "GetLocaleInfo error " & CStr(Err.LastDllError)
    End If
End Function

Private Function GetLocaleInfoAsString(ByVal Index As enmLocaleTypes, Optional ByVal Locale As enmLCIDs = LOCALE_USER_DEFAULT) As String
    Dim TChars As Long
    TChars = GetLocaleInfo(Locale, Index, 0, 0)
    If TChars Then
        GetLocaleInfoAsString = Space$(TChars - 1)
        GetLocaleInfo Locale, Index, StrPtr(GetLocaleInfoAsString), TChars
    Else
        Err.Raise &H80049908, , "GetLocaleInfo error " & CStr(Err.LastDllError)
    End If
End Function

'-> Publicas
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

Private Sub WndProc(ByVal bBefore As Boolean, _
                    ByRef bHandled As Boolean, _
                    ByRef lReturn As Long, _
                    ByVal hWnd As Long, _
                    ByVal uMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long, _
                    ByRef lParamUser As Long)
    
    Dim i As Integer
    
    On Error Resume Next

    If hWnd = UserControl.hWnd Then
        Select Case uMsg
            'Case WM_HOTKEY
            Case WM_MOUSELEAVE
                '-> Reset estado del mouse en botones
                For i = 0 To UBound(udtItemsNavButton)
                    'udtItemsNavButton(i).mouseState = MouseStateLeave
                    udtItemsNavButton(i).MouseState = Normal
                Next
                '-> Reset estado del mouse en Mes y año
                For i = 0 To UBound(udtItemsPicker)
                    'udtItemsPicker(i).MonthYearMouseState = MouseStateLeave
                    udtItemsPicker(i).TitleMonthYear.MouseState = Normal
                Next
                c_EnterButton = False
                '--
                '-> Reset estado del mouse en los dias.
                For i = 0 To UBound(udtItemsDay)
                    'udtItemsDay(i).mouseState = MouseStateLeave
                    udtItemsDay(i).MouseState = Normal
                Next
                '--
                '-> Reset estado del mouse en la navegacion rapida.
                For i = 0 To UBound(udtItemsMonthYear)
                    'udtItemsMonthYear(i).mouseState = MouseStateLeave
                    udtItemsMonthYear(i).MouseState = Normal
                Next
                '--
                Call Draw
            Case WM_IME_SETCONTEXT
                HideCalendar
        End Select
    End If
End Sub

