VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------
'Based in   : clsShadow Leandro Ascierto
'Name       : clsShadow
'Author     : Jose Liza
'Created    : 21/02/2022
'------------------------------------
'KERNEL32
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'USER32
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

'OLEAUT32
Private Declare Function OleTranslateColor Lib "Oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long

'GDI32
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "GDI32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
'-> Recion Function
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRgn Lib "GDI32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

'GDI+
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
'-> Graphic Function
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetClip Lib "gdiplus" (ByVal mGraphics As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipClosePathFigures Lib "gdiplus" (ByVal mPath As Long) As Long
Private Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "gdiplus" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathPieI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipSetClipPath Lib "gdiplus" (ByVal mGraphics As Long, ByVal mPath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal mPath As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
'-> LinearGradientBrush Function
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "gdiplus" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As wrapMode, ByRef mLineGradient As Long) As Long
'-> SolidBrush Function
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
'-> Image Function
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
'-> Bitmap Function
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal mBitmap As Long, ByRef mRect As RECTL, ByVal mFlags As ImageLockMode, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
'-> Pen / Brush Function
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal mPen As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
'-> Region Function
Private Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal path As Long, region As Long) As Long
Private Declare Function GdipSetClipHrgn Lib "gdiplus" (ByVal mGraphics As Long, ByVal mHRgn As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipGetRegionHRgn Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, hRgn As Long) As Long
Private Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal region As Long) As Long

'User Defined Type
Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
End Type

Private Type RECTL
    Left                                As Long
    Top                                 As Long
    Width                               As Long
    Height                              As Long
End Type

Private Type BitmapData
    Width                       As Long
    Height                      As Long
    stride                      As Long
    PixelFormat                 As Long
    Scan0Ptr                    As Long
    ReservedPtr                 As Long
End Type

Private Type Size
   cx                                   As Long
   cy                                   As Long
End Type

Private Type BLENDFUNCTION
   BlendOp                              As Byte
   BlendFlags                           As Byte
   SourceConstantAlpha                  As Byte
   AlphaFormat                          As Byte
End Type

Private Type BITMAPINFOHEADER
   biSize                               As Long
   biWidth                              As Long
   biHeight                             As Long
   biPlanes                             As Integer
   biBitCount                           As Integer
   biCompression                        As Long
   biSizeImage                          As Long
   biXPelsPerMeter                      As Long
   biYPelsPerMeter                      As Long
   biClrUsed                            As Long
   biClrImportant                       As Long
End Type

Private Type BITMAPINFO
   bmiHeader                            As BITMAPINFOHEADER
   bmiColors(3)                         As Byte
End Type

'Enums
Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum

Private Enum wrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

'Constants
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const GWL_STYLE                 As Long = -16
Private Const GWL_HWNDPARENT            As Long = -8
Private Const AC_SRC_ALPHA              As Long = &H1
Private Const ULW_ALPHA                 As Long = &H2
Private Const UnitPixel                 As Long = &H2&
Private Const WS_EX_CLIENTEDGE          As Long = &H200&
Private Const CombineModeXor            As Long = &H3
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const DIB_RGB_COLORS            As Long = 0&
Private Const WS_POPUP                  As Long = &H80000000
Private Const WS_EX_LAYERED             As Long = &H80000
Private Const WS_EX_TOOLWINDOW          As Long = &H80&
Private Const LOGPIXELSX                As Long = 88

Private Const WM_PAINT                  As Long = &HF&
Private Const SWP_NOSIZE                As Long = &H1
Private Const WM_DESTROY                As Long = &H2
Private Const WM_MOVE                   As Long = &H3

Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_SHOWWINDOW            As Long = &H40


Dim GdipToken                           As Long
Dim DpiFactor                           As Single
Dim c_hDC                               As Long
Dim c_DIB                               As Long
Dim c_OldBmp                            As Long
Dim c_Pointer                           As Long
Dim hImgShadow                          As Long
Dim nSize                               As Long
Dim WidthWin                            As Long
Dim HeightWin                           As Long
Dim SizeWin                             As Size
Dim hWinShadow                          As Long
Dim hParent                             As Long
Dim tBLENDFUNCTION                      As BLENDFUNCTION
Dim o_Obj                               As UserControl
'Dim cSubClass                           As clsSubClass
'--
Dim m_hWnd                              As Long
Dim m_hDC                               As Long
Dim m_Width                             As Long
Dim m_Height                            As Long

'Properties
Dim m_BackColor                         As OLE_COLOR
Dim m_UseGradient                       As Boolean
Dim m_BackColorGradient1                As OLE_COLOR
Dim m_BackColorGradient2                As OLE_COLOR
Dim m_BackColorGradientAngle            As Integer
Dim m_BackOpacity                       As Integer
Dim m_CornerTopLeft                     As Integer
Dim m_CornerTopRight                    As Integer
Dim m_CornerBottomLeft                  As Integer
Dim m_CornerBottomRight                 As Integer
Dim m_Border                            As Boolean
Dim m_BorderWidth                       As Integer
Dim m_BorderColor                       As OLE_COLOR
Dim m_BorderOpacity                     As Integer
Dim m_Shadow                            As Boolean
Dim m_ShadowSize                        As Integer
Dim m_ShadowColor                       As OLE_COLOR
Dim m_ShadowOpacity                     As Integer
Dim m_ShadowOffsetX                     As Integer
Dim m_ShadowOffsetY                     As Integer
'CallOut
Dim m_CallOut                           As Boolean
Dim m_CallOutWidth                      As Long
Dim m_CallOutHight                      As Long
Dim m_CallOutRightTriangle              As Boolean
Dim m_CallOutPosition                   As enmCallOutPosition
Dim m_CallOutCustomPos                  As Long
Dim m_CallOutAlign                      As enmCallOutAlign

'm_BackColor                         As OLE_COLOR
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(Value As OLE_COLOR)
    m_BackColor = Value
End Property

'm_UseGradient                       As Boolean
Public Property Get UseGradient() As Boolean
    UseGradient = m_UseGradient
End Property
Public Property Let UseGradient(Value As Boolean)
    m_UseGradient = Value
End Property

'm_BackColorGradient1                As OLE_COLOR
Public Property Get BackColorGradient1() As OLE_COLOR
    BackColorGradient1 = m_BackColorGradient1
End Property
Public Property Let BackColorGradient1(Value As OLE_COLOR)
    m_BackColorGradient1 = Value
End Property

'm_BackColorGradient2                As OLE_COLOR
Public Property Get BackColorGradient2() As OLE_COLOR
    BackColorGradient2 = m_BackColorGradient2
End Property
Public Property Let BackColorGradient2(Value As OLE_COLOR)
    m_BackColorGradient2 = Value
End Property

'm_BackColorGradientAngle            As Integer
Public Property Get BackColorGradientAngle() As Integer
    BackColorGradientAngle = m_BackColorGradientAngle
End Property
Public Property Let BackColorGradientAngle(Value As Integer)
    m_BackColorGradientAngle = Value
End Property

'm_BackOpacity                       As Integer
Public Property Get BackOpacity() As Integer
    BackOpacity = m_BackOpacity
End Property
Public Property Let BackOpacity(Value As Integer)
    m_BackOpacity = Value
End Property

'm_CornerTopLeft                     As Integer
Public Property Get CornerTopLeft() As Integer
    CornerTopLeft = m_CornerTopLeft
End Property
Public Property Let CornerTopLeft(Value As Integer)
    m_CornerTopLeft = Value
End Property

'm_CornerTopRight                    As Integer
Public Property Get CornerTopRight() As Integer
    CornerTopRight = m_CornerTopRight
End Property
Public Property Let CornerTopRight(Value As Integer)
    m_CornerTopRight = Value
End Property

'm_CornerBottomLeft                  As Integer
Public Property Get CornerBottomLeft() As Integer
    CornerBottomLeft = m_CornerBottomLeft
End Property
Public Property Let CornerBottomLeft(Value As Integer)
    m_CornerBottomLeft = Value
End Property

'm_CornerBottomRight                 As Integer
Public Property Get CornerBottomRight() As Integer
    CornerBottomRight = m_CornerBottomRight
End Property
Public Property Let CornerBottomRight(Value As Integer)
    m_CornerBottomRight = Value
End Property

'm_Border                            As Boolean
Public Property Get Border() As Boolean
    Border = m_Border
End Property
Public Property Let Border(Value As Boolean)
    m_Border = Value
End Property

'm_BorderWidth                       As Integer
Public Property Get BorderWidth() As Integer
    BorderWidth = m_BorderWidth
End Property
Public Property Let BorderWidth(Value As Integer)
    m_BorderWidth = Value
End Property

'm_BorderColor                       As OLE_COLOR
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(Value As OLE_COLOR)
    m_BorderColor = Value
End Property

'm_BorderOpacity                     As Integer
Public Property Get BorderOpacity() As Integer
    BorderOpacity = m_BorderOpacity
End Property
Public Property Let BorderOpacity(Value As Integer)
    m_BorderOpacity = Value
End Property

'm_Shadow                            As Boolean
Public Property Get Shadow() As Boolean
    Shadow = m_Shadow
End Property
Public Property Let Shadow(Value As Boolean)
    m_Shadow = Value
End Property

'm_ShadowSize                        As Integer
Public Property Get ShadowSize() As Integer
    ShadowSize = m_ShadowSize
End Property
Public Property Let ShadowSize(Value As Integer)
    m_ShadowSize = Value
End Property

'm_ShadowColor                       As OLE_COLOR
Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property
Public Property Let ShadowColor(Value As OLE_COLOR)
    m_ShadowColor = Value
End Property

'm_ShadowOpacity                     As Integer
Public Property Get ShadowOpacity() As Integer
    ShadowOpacity = m_ShadowOpacity
End Property
Public Property Let ShadowOpacity(Value As Integer)
    m_ShadowOpacity = Value
End Property

'm_ShadowOffsetX                     As Integer
Public Property Get ShadowOffsetX() As Integer
    ShadowOffsetX = m_ShadowOffsetX
End Property
Public Property Let ShadowOffsetX(Value As Integer)
    m_ShadowOffsetX = Value
End Property

'm_ShadowOffsetY                     As Integer
Public Property Get ShadowOffsetY() As Integer
    ShadowOffsetY = m_ShadowOffsetY
End Property
Public Property Let ShadowOffsetY(Value As Integer)
    m_ShadowOffsetY = Value
End Property

'm_CallOut                           As Boolean
Public Property Get CallOut() As Boolean
    CallOut = m_CallOut
End Property
Public Property Let CallOut(Value As Boolean)
    m_CallOut = Value
End Property

'm_CallOutWidth                      As Long
Public Property Get CallOutWidth() As Long
    CallOutWidth = m_CallOutWidth
End Property
Public Property Let CallOutWidth(Value As Long)
    m_CallOutWidth = Value
End Property

'm_CallOutHight                      As Long
Public Property Get CallOutHight() As Long
    CallOutHight = m_CallOutHight
End Property
Public Property Let CallOutHight(Value As Long)
    m_CallOutHight = Value
End Property

'm_CallOutRightTriangle              As Boolean
Public Property Get CallOutRightTriangle() As Boolean
    CallOutRightTriangle = m_CallOutRightTriangle
End Property
Public Property Let CallOutRightTriangle(Value As Boolean)
    m_CallOutRightTriangle = Value
End Property

'm_CallOutPosition                   As enmCallOutPosition
Public Property Get CallOutPosition() As enmCallOutPosition
    CallOutPosition = m_CallOutPosition
End Property
Public Property Let CallOutPosition(Value As enmCallOutPosition)
    m_CallOutPosition = Value
End Property

'm_CallOutCustomPos                  As Long
Public Property Get CallOutCustomPos() As Long
    CallOutCustomPos = m_CallOutCustomPos
End Property
Public Property Let CallOutCustomPos(Value As Long)
    m_CallOutCustomPos = Value
End Property

'm_CallOutAlign                      As enmCallOutAlign
Public Property Get CallOutAlign() As enmCallOutAlign
    CallOutAlign = m_CallOutAlign
End Property
Public Property Let CallOutAlign(Value As enmCallOutAlign)
    m_CallOutAlign = Value
End Property

Private Sub Class_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    '--
    DpiFactor = GetWindowsDPI
    '---
    With tBLENDFUNCTION
        .AlphaFormat = AC_SRC_ALPHA
        .SourceConstantAlpha = 255
    End With
    '---
    m_BackColor = SystemColorConstants.vbWindowBackground
    m_UseGradient = False
    m_BackColorGradient1 = SystemColorConstants.vbWindowBackground
    m_BackColorGradient2 = SystemColorConstants.vbWindowBackground
    m_BackColorGradientAngle = 0
    m_BackOpacity = 100
    m_CornerTopLeft = 0
    m_CornerTopRight = 0
    m_CornerBottomLeft = 0
    m_CornerBottomRight = 0
    m_Border = True
    m_BorderWidth = 1
    m_BorderColor = SystemColorConstants.vbActiveBorder
    m_BorderOpacity = 100
    m_Shadow = True
    m_ShadowSize = 2
    m_ShadowColor = SystemColorConstants.vbButtonShadow
    m_ShadowOpacity = 10
    m_ShadowOffsetX = 0
    m_ShadowOffsetY = 0
End Sub

Private Sub Class_Terminate()
    'Set cSubClass = Nothing
    If m_hWnd <> 0 Then SetWindowLong m_hWnd, GWL_HWNDPARENT, 0
    If hWinShadow Then DestroyWindow hWinShadow
    CleanUp
    If hImgShadow Then GdipDisposeImage hImgShadow: hImgShadow = 0
    Call GdiplusShutdown(GdipToken)
End Sub

Private Function GetWindowsDPI() As Double
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

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function RGBtoARGB(ByVal Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory RGBtoARGB, BGRA(0), 4&
End Function

Private Function CreateDIB() As Boolean
    Dim tBITMAPINFO     As BITMAPINFO
    '---
    CleanUp
    '--
    With tBITMAPINFO.bmiHeader
        .biSize = Len(tBITMAPINFO.bmiHeader)
        .biBitCount = 32
        .biHeight = HeightWin
        .biWidth = WidthWin
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * 4
    End With
    '--
    c_hDC = CreateCompatibleDC(0)
    c_DIB = CreateDIBSection(c_hDC, tBITMAPINFO, DIB_RGB_COLORS, c_Pointer, 0&, 0&)
    c_OldBmp = SelectObject(c_hDC, c_DIB)
    '---
End Function

Private Sub CleanUp()
    If c_DIB Then Call DeleteObject(SelectObject(c_hDC, c_OldBmp)): c_DIB = 0
    If c_hDC Then Call DeleteDC(c_hDC): c_hDC = 0
End Sub

'Private Sub CreateShadow()
'    Dim hImage As Long
'    Dim hGraphics As Long
'    Dim hPath As Long, hBrush As Long
'    nSize = (m_ShadowSize * 4 + 10 * 2) * DpiFactor
'    If hImgShadow Then GdipDisposeImage hImgShadow
'    GdipCreateBitmapFromScan0 nSize, nSize, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
'    GdipGetImageGraphicsContext hImage, hGraphics
'
'    hPath = CreateRoundPath(0, 0, nSize, nSize, True)
'    GdipCreateSolidFill RGBToARGB(m_ShadowColor, m_ShadowOpacity), hBrush
'    GdipFillPath hGraphics, hBrush, hPath
'    GdipDeleteBrush hBrush
'
'    hImgShadow = CreateBlurShadowImage(hImage, m_ShadowColor, m_ShadowSize * DpiFactor, 0, 0, nSize, nSize)
'    GdipDeletePath hPath
'    GdipDeleteGraphics hGraphics
'    GdipDisposeImage hImage
'End Sub
Private Sub CreateShadow(Width As Long, Height As Long)
    Dim hImage As Long
    Dim hGraphics As Long
    Dim hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim lWidth As Long, lHeight As Long
    Dim ShadowSize As Integer
    
    If hImgShadow Then GdipDisposeImage hImgShadow: hImgShadow = 0
    If Me.Shadow = False Then Exit Sub
        
    If m_ShadowSize = 0 Then Exit Sub
    If Me.BackOpacity = 0 And m_Border = False Then Exit Sub
    If Me.ShadowOpacity = 0 Then Exit Sub
   
    ShadowSize = m_ShadowSize * DpiFactor
    lWidth = Width - (ShadowSize * 2)
    lHeight = Height - (ShadowSize * 2)
    
    GdipCreateBitmapFromScan0 lWidth, lHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    hPath = CreateRoundPath(0, 0, lWidth - 0, lHeight - 0, True)
    
    If Me.BackOpacity > 0 Then
        GdipCreateSolidFill RGBtoARGB(m_ShadowColor, m_ShadowOpacity), hBrush
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
    Else
        GdipCreatePen1 RGBtoARGB(m_ShadowColor, m_ShadowOpacity), (m_BorderWidth * DpiFactor) * 2, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    End If
        
    hImgShadow = CreateBlurShadowImage(hImage, m_ShadowColor, ShadowSize, 0, 0, lWidth, lHeight)
    
    GdipDeletePath hPath
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
    
End Sub

Private Function CreateBlurShadowImage(ByVal hImage As Long, ByVal Color As Long, blurDepth As Integer, _
                                        Optional ByVal Left As Long, Optional ByVal Top As Long, _
                                        Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
                                        
    Dim REC As RECTL
    Dim X As Long, Y As Long
    Dim hImgShadow As Long
    Dim bmpData1 As BitmapData
    Dim bmpData2 As BitmapData
    Dim t2xBlur As Long
    Dim R As Long, G As Long, b As Long
    Dim Alpha As Byte
    Dim lSrcAlpha As Long, lDestAlpha As Long
    Dim dBytes() As Byte
    Dim srcBytes() As Byte
    Dim vTally() As Long
    Dim tAlpha As Long, tColumn As Long, tAvg As Long
    Dim initY As Long, initYstop As Long, initYstart As Long
    Dim initX As Long, initXstop As Long
    
    If hImage = 0& Then Exit Function
 
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 
    t2xBlur = blurDepth * 2
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    b = (Color \ &H10000) And &HFF
 
    SetRect REC, Left, Top, Width, Height
 
    ReDim srcBytes(REC.Width * 4 - 1&, REC.Height - 1&)
  
    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
   
    Call GdipBitmapLockBits(hImage, REC, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppPARGB, bmpData1)
 
    SetRect REC, Left, Top, Width + t2xBlur, Height + t2xBlur
    
    Call GdipCreateBitmapFromScan0(REC.Width, REC.Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImgShadow)

    ReDim dBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    
    With bmpData2
        .Scan0Ptr = VarPtr(dBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
    
    Call GdipBitmapLockBits(hImgShadow, REC, ImageLockModeUserInputBuf Or ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppPARGB, bmpData2)
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    b = (Color \ &H10000) And &HFF
    
    tAvg = (t2xBlur + 1) * (t2xBlur + 1)    ' how many pixels are being blurred
    
    ReDim vTally(0 To t2xBlur)              ' number of blur columns per pixel
    
    For Y = 0 To Height + t2xBlur - 1     ' loop thru shadow dib
    
        FillMemory vTally(0), (t2xBlur + 1) * 4, 0  ' reset column totals
        
        If Y < t2xBlur Then         ' y does not exist in source
            initYstart = 0          ' use 1st row
        Else
            initYstart = Y - t2xBlur ' start n blur rows above y
        End If
        ' how may source rows can we use for blurring?
        If Y < Height Then initYstop = Y Else initYstop = Height - 1
        
        tAlpha = 0  ' reset alpha sum
        tColumn = 0    ' reset column counter
        
        ' the first n columns will all be zero
        ' only the far right blur column has values; tally them
        For initY = initYstart To initYstop
            tAlpha = tAlpha + srcBytes(3, initY)
        Next
        ' assign the right column value
        vTally(t2xBlur) = tAlpha
        
        For X = 3 To (Width - 2) * 4 - 1 Step 4
            ' loop thru each source pixel's alpha
            
            ' set shadow alpha using blur average
            dBytes(X, Y) = tAlpha \ tAvg
            ' and set shadow color
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = b
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = b * dBytes(X, Y) \ 255
            End Select
            ' remove the furthest left column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' count the next column of alphas
            vTally(tColumn) = 0&
            For initY = initYstart To initYstop
                vTally(tColumn) = vTally(tColumn) + srcBytes(X + 4, initY)
            Next
            ' add the new column's sum to the overall sum
            tAlpha = tAlpha + vTally(tColumn)
            ' set the next column to be recalculated
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
        
        ' now to finish blurring from right edge of source
        For X = X To (Width + t2xBlur - 1) * 4 - 1 Step 4
            dBytes(X, Y) = tAlpha \ tAvg
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = b
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = b * dBytes(X, Y) \ 255
            End Select
            ' remove this column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' set next column to be removed
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
    Next
 
    Call GdipBitmapUnlockBits(hImage, bmpData1)
    Call GdipBitmapUnlockBits(hImgShadow, bmpData2)
    
    CreateBlurShadowImage = hImgShadow
End Function

Private Sub RenderStretchPlus(ByVal hGraphics As Long, _
            ByVal DestX As Long, ByVal DestY As Long, ByVal DestW As Long, ByVal DestH As Long, ByVal hImage As Long, _
            ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Size As Long)

    Dim Sx2 As Long
  
    Sx2 = Size * 2

    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY, Size, Size, X, Y, Size, Size, &H2, 0&, 0&, 0& 'TOP_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY, DestW - Sx2, Size, X + Size, Y, Width - Sx2, Size, &H2, 0&, 0&, 0& 'TOP_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY, Size, Size, X + Width - Size, Y, Size, Size, &H2, 0&, 0&, 0& 'TOP_RIGHT
    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY + Size, Size, DestH - Sx2, X, Y + Size, Size, Height - Sx2, &H2, 0&, 0&, 0& 'MID_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY + Size, DestW - Sx2, DestH - Sx2, X + Size, Y + Size, Width - Sx2, Height - Sx2, &H2, 0&, 0&, 0& 'MID_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY + Size, Size, DestH - Sx2, X + Width - Size, Y + Size, Size, Height - Sx2, &H2, 0&, 0&, 0& 'MID_RIGHT
    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY + DestH - Size, Size, Size, X, Y + Height - Size, Size, Size, &H2, 0&, 0&, 0& 'BOTTOM_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY + DestH - Size, DestW - Sx2, Size, X + Size, Y + Height - Size, Width - Sx2, Size, &H2, 0&, 0&, 0& 'BOTTOM_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY + DestH - Size, Size, Size, X + Width - Size, Y + Height - Size, Size, Size, &H2, 0&, 0&, 0& 'BOTTOM_RIGHT
End Sub

Private Function CreateRoundPath(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional CallOut As Boolean) As Long
    Dim hPath As Long
    Dim BCTL As Integer
    Dim BCTR As Integer
    Dim BCBR As Integer
    Dim BCBL As Integer
    '--
    Dim XX As Long
    Dim YY As Long
    Dim coLen As Long
    Dim lMax As Long
    Dim coWidth As Long
    Dim coAngle  As Long
    '---
    '--> CallOut
    If CallOut Then
        coWidth = m_CallOutWidth * DpiFactor
        coLen = m_CallOutHight * DpiFactor
        coAngle = IIF(m_CallOutRightTriangle, 0, (m_CallOutWidth * DpiFactor) / 2)
    End If
    '<-- CallOut
    
    If GdipCreatePath(&H0, hPath) = 0 Then
        'Top Left:
        BCTL = GetSafeRound(m_CornerTopLeft * DpiFactor, Width, Height)
        'Top Right:
        BCTR = GetSafeRound(m_CornerTopRight * DpiFactor, Width, Height)
        'Bottom Right:
        BCBR = GetSafeRound(m_CornerBottomRight * DpiFactor, Width, Height)
        'Bottom Left:
        BCBL = GetSafeRound(m_CornerBottomLeft * DpiFactor, Width, Height)
        '---
        If m_CallOut Then
            Select Case m_CallOutPosition
                Case [Position Left]
                    Left = Left + coLen
                    Width = Width - coLen
                    lMax = Height - BCTL - BCBL
                    If coWidth > lMax Then coWidth = lMax
                Case [Position Top]
                    Top = Top + coLen
                    Height = Height - coLen
                    lMax = Width - BCTL - BCBL
                    If coWidth > lMax Then coWidth = lMax
                Case [Position Right]
                    Width = Width - coLen
                    lMax = Height - BCTR - BCBR
                    If coWidth > lMax Then coWidth = lMax
                Case [Position Bottom]
                    Height = Height - coLen
                    lMax = Width - BCBL - BCBR
                    If coWidth > lMax Then coWidth = lMax
            End Select
        End If
        '---
        'Top Left:
        If BCTL Then GdipAddPathArcI hPath, Left, Top, BCTL * 2, BCTL * 2, 180, 90
        '--
        If m_CallOutPosition = [Position Top] And m_CallOut Then
            Select Case m_CallOutAlign
                Case [First Corner]: XX = Left + BCTL
                Case Middle: XX = Left + BCTL + ((Width - BCTL - BCTR) \ 2) - (coWidth \ 2)
                Case [Second Corner]: XX = Left + Width - coWidth - BCTR
                Case [Custom Position]: XX = Left + BCTL + (lMax * (m_CallOutCustomPos / 100)) - (coWidth \ 2)
            End Select
            '--
            If (XX > Width / 2) And coAngle = 0 Then
                GdipAddPathLineI hPath, XX, Top, XX + coWidth, Top - coLen
                GdipAddPathLineI hPath, XX + coWidth, Top - coLen, XX + coWidth, Top
            Else
                If BCTL = 0 Then GdipAddPathLineI hPath, Left, Top, Left, Top
                GdipAddPathLineI hPath, XX, Top, XX + coAngle, Top - coLen
                GdipAddPathLineI hPath, XX + coAngle, Top - coLen, XX + coWidth, Top
            End If
        Else
            If BCTL = 0 Then GdipAddPathLineI hPath, Left, Top, Left + Width - BCTR, Top
        End If
        
        'Top Right:
        If BCTR Then GdipAddPathArcI hPath, Left + Width - BCTR * 2, Top, BCTR * 2, BCTR * 2, 270, 90
        If m_CallOutPosition = [Position Right] And m_CallOut Then
            Select Case m_CallOutAlign
                Case [First Corner]: YY = Top + BCTR
                Case Middle: YY = Top + BCTR + ((Height - BCTR - BCBR) \ 2) - (coWidth \ 2)
                Case [Second Corner]: YY = Top + Height - coWidth - BCBR
                Case [Custom Position]: YY = Top + BCTR + (lMax * (m_CallOutCustomPos / 100)) - (coWidth \ 2)
            End Select
            '--
            XX = Left + Width
            '--
            If (YY > Height / 2) And coAngle = 0 Then
                GdipAddPathLineI hPath, XX, YY, XX + coLen, YY + coWidth
                GdipAddPathLineI hPath, XX + coLen, YY + coWidth, XX, YY + coWidth
                
            Else
                If BCTR = 0 Then GdipAddPathLineI hPath, Left + Width, Top, Left + Width, Top
                GdipAddPathLineI hPath, XX, YY, XX + coLen, YY + coAngle
                GdipAddPathLineI hPath, XX + coLen, YY + coAngle, XX, YY + coWidth
            End If
        Else
            If BCTR = 0 Then GdipAddPathLineI hPath, Left + Width, Top, Left + Width, Top + Height - BCBR
        End If
        
        'Bottom Right:
        If BCBR Then GdipAddPathArcI hPath, Left + Width - BCBR * 2, Top + Height - BCBR * 2, BCBR * 2, BCBR * 2, 0, 90
        '--
        If m_CallOutPosition = [Position Bottom] And m_CallOut Then
            Select Case m_CallOutAlign
                Case [First Corner]: XX = Left + BCBL
                Case Middle: XX = Left + BCBL + ((Width - BCBR - BCBL) \ 2) - (coWidth \ 2)
                Case [Second Corner]: XX = Left + Width - coWidth - BCBR
                Case [Custom Position]: XX = Left + BCBR + (lMax * (m_CallOutCustomPos / 100)) - (coWidth \ 2)
            End Select
            '--
            YY = Top + Height
            '--
            If (XX > Width / 2) And coAngle = 0 Then
                GdipAddPathLineI hPath, XX + coWidth, YY, XX + coWidth, YY + coLen
                GdipAddPathLineI hPath, XX + coWidth, YY + coLen, XX, YY
            Else
                If BCBR = 0 Then GdipAddPathLineI hPath, Left + Width, Top + Height, Left + Width, Top + Height
                GdipAddPathLineI hPath, XX + coWidth, YY, XX + coAngle, YY + coLen
                GdipAddPathLineI hPath, XX + coAngle, YY + coLen, XX, YY
            End If
        Else
            If BCBR = 0 Then GdipAddPathLineI hPath, Left + Width, Top + Height, Left + BCBL, Top + Height
        End If
        
        'Bottom Left:
        If BCBL Then GdipAddPathArcI hPath, Left, Top + Height - BCBL * 2, BCBL * 2, BCBL * 2, 90, 90
        '--
        If m_CallOutPosition = [Position Left] And m_CallOut Then
            Select Case m_CallOutAlign
                Case [First Corner]: YY = Top + BCTL
                Case Middle: YY = Top + BCTL + ((Height - BCBL - BCTL) \ 2) - (coWidth \ 2)
                Case [Second Corner]: YY = Top + Height - coWidth - BCBL
                Case [Custom Position]: YY = Top + BCBL + (lMax * (m_CallOutCustomPos / 100)) - (coWidth \ 2)
            End Select
            
            If (YY > Height / 2) And coAngle = 0 Then
                GdipAddPathLineI hPath, Left, YY + coWidth, Left - coLen, YY + coWidth
                GdipAddPathLineI hPath, Left - coLen, YY + coWidth, Left, YY
            Else
                If BCBL = 0 Then GdipAddPathLineI hPath, Left, Top + Height, Left, Top + Height
                GdipAddPathLineI hPath, Left, YY + coWidth, Left - coLen, YY + coAngle
                GdipAddPathLineI hPath, Left - coLen, YY + coAngle, Left, YY
            End If
        Else
            If BCBL = 0 Then GdipAddPathLineI hPath, Left, Top + Height, Left, Top + BCTL
        End If
        'Closed path:
        GdipClosePathFigures hPath
    End If
    '---
    CreateRoundPath = hPath
    '---
End Function

Private Function CreateRoundRegion(hGraphics As Long, X As Long, Y As Long, Width As Long, Height As Long, Optional CallOut As Boolean)
    Dim hPath As Long
    Dim hRegion As Long
    Dim hRegion2 As Long
    Dim hBrush As Long
    '---
    hPath = CreateRoundPath(X, Y, Width, Height, CallOut)
    '---
    If hPath <> 0 Then
        
        GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
        GdipCreateRegionPath hPath, hRegion
        GdipGetRegionHRgn hRegion, hGraphics, hRegion2
        '---
        GdipDeleteRegion hRegion
        GdipDeleteBrush hBrush
        GdipDeletePath hPath
        '---
        CreateRoundRegion = hRegion2
    End If
    '---
End Function

'---
Public Sub InitShadow(ByVal hWnd As Long, ByVal hdc As Long)
    '---
    m_hWnd = hWnd
    m_hDC = hdc 'GetDC(hWnd)
    '---
    Update
End Sub

Public Sub EndShadow()
    ShowWindow hWinShadow, vbHide
    Call CleanUp
End Sub

Public Sub Update()
    Dim hGraphics       As Long
    Dim hPath           As Long
    Dim R               As RECT
    Dim pt              As POINTAPI
    Dim tSize           As Size
    Dim TotMargin       As Long
    Dim Left            As Long
    Dim Top             As Long
    Dim hRgn            As Long
    Dim hRes            As Long
    Dim hPen            As Long
    Dim hBrush          As Long
    Dim RL              As RECTL
    Dim ShadowOffsetX   As Integer
    Dim ShadowOffsetY   As Integer
    Dim XX              As Long
    Dim YY              As Long
    Dim m_TopMargin     As Long
    '---
    'If Not TypeOf o_Obj Is Form Then Exit Sub
    '--
    Call GetWindowRect(m_hWnd, R)
    SizeWin.cx = R.Right - R.Left
    SizeWin.cy = R.Bottom - R.Top
    '---
    'Debug.Print "SizeWin.cx = " & SizeWin.cx
    'Debug.Print "SizeWin.cy = " & SizeWin.cy
    '---
    m_Width = SizeWin.cx + Abs(m_ShadowOffsetX) * DpiFactor
    m_Height = SizeWin.cy + Abs(m_ShadowOffsetY) * DpiFactor
    '--
    m_TopMargin = 10 * DpiFactor
    '************************************************
    '** Rounded Object (Form / PictureBox / Frame) **
    '************************************************
    If hGraphics = 0 Then GdipCreateFromHDC m_hDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipTranslateWorldTransform hGraphics, 0, 0, &H1
    '--
    hPath = CreateRoundPath(-1, -1, (m_Width + 1) * DpiFactor, (m_Height + 1) * DpiFactor, True)
    '-- Gradient object
    If m_UseGradient Then
        SetRect RL, 0, 0, m_Width * DpiFactor, m_Height * DpiFactor
        GdipCreateLineBrushFromRectWithAngleI RL, RGBtoARGB(m_BackColorGradient1, m_BackOpacity), RGBtoARGB(m_BackColorGradient2, m_BackOpacity), m_BackColorGradientAngle + 90, 0, WrapModeTileFlipXY, hBrush
    Else
        GdipCreateSolidFill RGBtoARGB(m_BackColor, m_BackOpacity), hBrush
    End If
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
    '--
    hRgn = CreateRoundRegion(hGraphics, 0, 0, m_Width - m_BorderWidth, m_Height, True)
    hRes = SetWindowRgn(m_hWnd, hRgn, True)
    '--
    If hGraphics <> 0 Then GdipDeleteGraphics hGraphics
    '************************************************
    '** Rounded Object (Form / PictureBox / Frame) **
    '************************************************
    Call CreateShadow(SizeWin.cx, SizeWin.cy)
    '--
    TotMargin = (m_ShadowSize * 2 + m_BorderWidth) * DpiFactor
    '---
    WidthWin = SizeWin.cx + TotMargin * 2 + Abs(m_ShadowOffsetX) * DpiFactor
    HeightWin = SizeWin.cy + TotMargin * 2 + Abs(m_ShadowOffsetY) * DpiFactor
    '---
    If m_hWnd <> 0 Then
        hParent = GetParent(m_hWnd)
        'Set cSubClass = New clsSubClass
        hWinShadow = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_LAYERED, "static", vbNullString, WS_POPUP, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, 0&)
        'cSubClass.ssc_Subclass m_hWnd, , , Me
        'cSubClass.ssc_AddMsg m_hWnd, WM_MOVE, MSG_AFTER
    End If
    '---
    Left = R.Left - TotMargin
    Top = R.Top - TotMargin
    '---
    If m_ShadowOffsetX < 0 Then Left = Left + m_ShadowOffsetX * DpiFactor
    If m_ShadowOffsetY < 0 Then Top = Top + m_ShadowOffsetY * DpiFactor
    '---
    Call SetWindowPos(hWinShadow, m_hWnd, Left, Top, WidthWin, HeightWin, SWP_NOACTIVATE Or SWP_SHOWWINDOW)   'Or SWP_FRAMECHANGED
    Call SetWindowLong(m_hWnd, GWL_HWNDPARENT, hWinShadow)
    '---
    CreateDIB
    '--
    GdipCreateFromHDC c_hDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    '--
    Left = (m_ShadowSize * 2 + m_BorderWidth / 2) * DpiFactor
    Top = (m_ShadowSize * 2 + m_BorderWidth / 2) * DpiFactor
    '--
    If m_ShadowOffsetX < 0 Then Left = Left - m_ShadowOffsetX * DpiFactor
    If m_ShadowOffsetY < 0 Then Top = Top - m_ShadowOffsetY * DpiFactor
    '--
    If Left <= 0 Then Left = 0
    If Top < 0 Then Top = 0
    '--
    hPath = CreateRoundPath(Left, Top, SizeWin.cx + (m_BorderWidth) * DpiFactor, SizeWin.cy + (m_BorderWidth) * DpiFactor, True)
    '--
    GdipSetClipPath hGraphics, hPath, CombineModeXor
    '--
    Left = 0: Top = 0
    If m_ShadowOffsetX > 0 Then Left = m_ShadowOffsetX * DpiFactor
    If m_ShadowOffsetY > 0 Then Top = m_ShadowOffsetY * DpiFactor
    '--
    If m_ShadowSize > 0 And m_ShadowOpacity > 0 Then
        '--> Usando metodo de labelplus - Leandro
        XX = IIF(ShadowOffsetX > 0, ShadowOffsetX, 0) '+ PosX
        YY = IIF(ShadowOffsetY > 0, ShadowOffsetY, 0) '+ PosY
        GdipDrawImageRectI hGraphics, hImgShadow, XX, YY, SizeWin.cx + (10 * DpiFactor) - Abs(ShadowOffsetX), SizeWin.cy + (10 * DpiFactor) - Abs(ShadowOffsetY)
        'm_TopMargin
        '--> Estirando imagen metodo clsSahdow - Leandro
        'RenderStretchPlus hGraphics, Left, Top + m_TopMargin, WidthWin - Abs(m_ShadowOffsetX) * DpiFactor, HeightWin - Abs(m_ShadowOffsetY) * DpiFactor - m_TopMargin, hImgShadow, 0, 0, nSize + m_ShadowSize * 2 * DpiFactor, nSize + m_ShadowSize * 2 * DpiFactor, (m_ShadowSize * 2 + 10) * DpiFactor
    End If
    '---
    GdipResetClip hGraphics
    '--
    Left = (m_ShadowSize * 2 + m_BorderWidth) * DpiFactor
    Top = (m_ShadowSize * 2 + m_BorderWidth) * DpiFactor
    '--
'    hRgn = CreateRectRgn(0, 0, SizeWin.cx, SizeWin.cy)
'    Call GetWindowRgn(m_hWnd, hRgn)
'    OffsetRgn hRgn, Left, Top
'    GdipSetClipHrgn hGraphics, hRgn, CombineModeXor
'    DeleteObject hRgn
    '--
    'Gradient
    If m_UseGradient Then
        SetRect RL, Left, Top, WidthWin, HeightWin
        GdipCreateLineBrushFromRectWithAngleI RL, RGBtoARGB(m_BackColorGradient1, m_BackOpacity), RGBtoARGB(m_BackColorGradient2, m_BackOpacity), m_BackColorGradientAngle + 90, 0, WrapModeTileFlipXY, hBrush
    Else
        GdipCreateSolidFill RGBtoARGB(m_BackColor, m_BackOpacity), hBrush
    End If
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    '--
    If m_BorderWidth > 0 And m_BorderOpacity > 0 Then
        GdipCreatePen1 RGBtoARGB(m_BorderColor, m_BorderOpacity), m_BorderWidth * DpiFactor, &H2, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    End If
    '--
    GdipDeletePath hPath
    GdipDeleteGraphics hGraphics
    If hImgShadow Then GdipDisposeImage hImgShadow: hImgShadow = 0
    '--
    tSize.cx = WidthWin: tSize.cy = HeightWin
    Call UpdateLayeredWindow(hWinShadow, 0&, ByVal 0&, tSize, c_hDC, pt, 0&, tBLENDFUNCTION, ULW_ALPHA)
    '---
    Call ReleaseDC(m_hWnd, m_hDC)
    '---
End Sub

Private Function HiWord(dw As Long) As Integer
 If dw And &H80000000 Then
      HiWord = (dw \ 65535) - 1
 Else
    HiWord = dw \ 65535
 End If
End Function

Private Function LoWord(dw As Long) As Integer
  If dw And &H8000& Then
      LoWord = &H8000 Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
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

    If uMsg = WM_PAINT Then ' CHILD WINDOWS
        'Update
    ElseIf uMsg = WM_DESTROY Then
        'cSubClass.ssc_UnSubclass hWinShadow
    ElseIf uMsg = WM_MOVE Then 'PARENT WINDOWS
        Dim X As Long, Y As Long

        X = LoWord(lParam) - (m_ShadowSize * 2 + m_BorderWidth) * DpiFactor
        Y = HiWord(lParam) - (m_ShadowSize * 2 + m_BorderWidth) * DpiFactor
        If m_ShadowOffsetX < 0 Then X = X + m_ShadowOffsetX * DpiFactor
        If m_ShadowOffsetY < 0 Then Y = Y + m_ShadowOffsetY * DpiFactor

        SetWindowPos hWinShadow, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE
    End If
End Sub


