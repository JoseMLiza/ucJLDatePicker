VERSION 5.00
Begin VB.Form frmUserControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ucJLPicker"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9105
   Icon            =   "frmUserControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkProperties 
      Caption         =   "UseGDIPString"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
   End
   Begin Proyecto1.ucJLDTPicker ucJLDTPicker 
      Height          =   3960
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   6985
      CornerTopLeft   =   5
      CornerTopRight  =   5
      CornerBottomLeft=   5
      CornerBottomRight=   5
      PaddingX        =   10
      PaddingY        =   10
      Shadow          =   -1  'True
      ShadowSize      =   2
      ShadowOpacity   =   10
      SpaceGrid       =   1
      UseRangeValue   =   -1  'True
      AutoApply       =   0   'False
      IsChild         =   -1  'True
      BackColorParent =   -2147483633
      ColsPicker      =   2
      Value           =   44866
      MinDate         =   44562
      MaxDate         =   44926
      FirstDayOfWeek  =   2
      ButtonNavCornerRadius=   12
      BeginProperty ButtonNavIcoFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ButtonsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthYearBackColor=   -2147483643
      MonthYearBorderWidth=   1
      MonthYearCornerRadius=   5
      BeginProperty MonthYearFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WeekBorderColor =   0
      BeginProperty WeekFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WeekForeColor   =   8421504
      DayBorderColor  =   0
      DayCornerRadius =   5
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DayFreeArray    =   ""
      DayNowShow      =   -1  'True
      DayNowBorderWidth=   1
      DayNowBorderColor=   11562240
      DaySaturdayForeColor=   0
      DaySundayForeColor=   192
      CallOutAlign    =   3
      CallOutCustomPosPercent=   10
   End
   Begin VB.CheckBox chkProperties 
      Caption         =   "ShowRange"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin Proyecto1.ucText ucTxtProperties 
      Height          =   345
      Index           =   0
      Left            =   1440
      TabIndex        =   9
      Top             =   1515
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColor       =   16777215
      Text            =   "frmUserControl.frx":424A
      InputType       =   1
      ImgLeft         =   "frmUserControl.frx":426C
      ImgRight        =   "frmUserControl.frx":4284
      RightButtonStyle=   3
      MinValue        =   "0"
      MaxValue        =   "12"
   End
   Begin VB.CheckBox chkProperties 
      Caption         =   "AutoApply"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   2115
   End
   Begin VB.CheckBox chkProperties 
      Caption         =   "RightToLeft"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin Proyecto1.ucText ucTextValue 
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   75
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Text            =   "frmUserControl.frx":429C
      InputType       =   2
      Alignment       =   2
      ImgLeft         =   "frmUserControl.frx":42D0
      ImgRight        =   "frmUserControl.frx":42E8
      RightButtonStyle=   2
      ShortDateFormat =   "DD/MM/YYYY"
   End
   Begin VB.CheckBox chkProperties 
      Caption         =   "Use Range"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkProperties 
      Caption         =   "LinkedCalendar"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin Proyecto1.ucText ucTxtProperties 
      Height          =   345
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   1920
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColor       =   16777215
      Text            =   "frmUserControl.frx":4652
      InputType       =   1
      ImgLeft         =   "frmUserControl.frx":4674
      ImgRight        =   "frmUserControl.frx":468C
      RightButtonStyle=   3
      MinValue        =   "1"
      MaxValue        =   "12"
   End
   Begin Proyecto1.ucText ucTextValue 
      Height          =   315
      Index           =   1
      Left            =   840
      TabIndex        =   16
      Top             =   435
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Text            =   "frmUserControl.frx":46A4
      InputType       =   2
      Alignment       =   2
      ImgLeft         =   "frmUserControl.frx":46D8
      ImgRight        =   "frmUserControl.frx":46F0
      RightButtonStyle=   2
      ShortDateFormat =   "DD/MM/YYYY"
   End
   Begin Proyecto1.ucJLDTPicker ucJLDTPicker1 
      Height          =   480
      Left            =   1680
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      CornerTopLeft   =   5
      CornerTopRight  =   5
      CornerBottomLeft=   5
      CornerBottomRight=   5
      PaddingX        =   10
      PaddingY        =   10
      Shadow          =   -1  'True
      ShadowSize      =   2
      ShadowOpacity   =   10
      SpaceGrid       =   1
      AutoApply       =   0   'False
      BackColorParent =   -2147483633
      ColsPicker      =   1
      NumberPickers   =   1
      Value           =   44562
      FirstDayOfWeek  =   2
      ButtonNavCornerRadius=   12
      BeginProperty ButtonNavIcoFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ButtonsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthYearBorderWidth=   1
      MonthYearCornerRadius=   5
      BeginProperty MonthYearFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WeekBorderColor =   0
      BeginProperty WeekFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WeekForeColor   =   8421504
      DayBorderColor  =   0
      DayCornerRadius =   5
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DayFreeArray    =   ""
      DayNowShow      =   -1  'True
      DayNowBorderWidth=   1
      DayNowBorderColor=   11562240
      DaySaturdayForeColor=   0
      DaySundayForeColor=   192
      CallOutAlign    =   3
      CallOutCustomPosPercent=   10
   End
   Begin VB.Label Label4 
      Caption         =   "MaxDate:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "MinDate:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Calendarios:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1965
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Columnas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FirstDayOfWeek:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmUserControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim iText As Integer

Private Sub chkProperties_Click(Index As Integer)
    With ucJLDTPicker
        Select Case Index
            Case 1 'LinkedCalendars
                .LinkedCalendars = chkProperties(Index).Value
            Case 2 'UseRangeValue
                .UseRangeValue = chkProperties(Index).Value
            Case 3 'ShowRange
                .ShowRangeButtons = chkProperties(Index).Value
            Case 4 'RightToLeft
                .RightToLeft = chkProperties(Index).Value
            Case 5 'AutoApply
                .AutoApply = chkProperties(Index).Value
            Case 6 'UseGDIPString
                .UseGDIPString = chkProperties(Index).Value
        End Select
    End With
End Sub

Private Sub Combo1_Click()
    ucJLDTPicker.FirstDayOfWeek = Combo1.ListIndex + 1
End Sub

Private Sub Form_Load()
    ucTextValue(0).Text = ucJLDTPicker.MinDate
    ucTextValue(1).Text = ucJLDTPicker.MaxDate
    With Combo1
        .AddItem "Domingo"
        .AddItem "Lunes"
        .AddItem "Martes"
        .AddItem "Miercoles"
        .AddItem "Jueves"
        .AddItem "Viernes"
        .AddItem "Sábado"
        .ListIndex = 0
    End With
    '---
    ucTxtProperties(0).Text = ucJLDTPicker.ColsPicker
    ucTxtProperties(1).Text = ucJLDTPicker.NumberPickers
End Sub

Private Sub ucJLDTPicker_ButtonActionClick(ByVal Index As Variant, Caption As String)
    MsgBox "Button Index(" & Index & "): " & Caption
End Sub

Private Sub ucJLDTPicker_ButtonRangeClick(ByVal Index As Variant, Caption As String)
    MsgBox "Button Index(" & Index & "): " & Caption
End Sub

Private Sub ucJLDTPicker_ChangeMaxDate()
    ucTextValue(1).Text = ucJLDTPicker.MaxDate
End Sub

Private Sub ucJLDTPicker_ChangeMinDate()
    ucTextValue(0).Text = ucJLDTPicker.MinDate
End Sub

Private Sub ucJLDTPicker_DayPrePaint(ByVal dDate As Date, BackColor As Long)
    If dDate >= CDate("25/02/2022") And dDate <= CDate("15/03/2022") Then BackColor = &H99B418
    If dDate = CDate("27/02/2022") Then BackColor = &H94D9FF
End Sub

Private Sub ucText1_ImgRightMouseUp(Button As Variant, Shift As Integer, X As Single, Y As Single)
    Dim RECT As RECT
    If Not ucJLDTPicker.IsChild Then
        GetWindowRect ucText1.hWnd, RECT
        ucJLDTPicker.Visible = False
        ucJLDTPicker.Visible = True
        ucJLDTPicker.ShowCalendar RECT.Left, RECT.Bottom
        
        ucJLDTPicker.Value = CDate(ucText1.Text)
    End If
End Sub

Private Sub ucText1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then ucJLDTPicker.HideCalendar
End Sub

Private Sub ucJLDTPicker_Resize()
    Me.Width = ucJLDTPicker.Left + ucJLDTPicker.Width + 420
    Me.Height = ucJLDTPicker.Top + ucJLDTPicker.Height + 735
End Sub

Private Sub ucJLDTPicker1_ChangeDate(ByVal Value As Date)
    ucTextValue(iText).Text = Value
End Sub

Private Sub ucTextValue_ImgRightMouseUp(Index As Integer, Button As Variant, Shift As Integer, X As Single, Y As Single)
    Dim RECT As RECT
    If Not ucJLDTPicker1.IsChild Then
        GetWindowRect ucTextValue(Index).hWnd, RECT
        ucJLDTPicker1.ShowCalendar RECT.Left, RECT.Bottom
        ucJLDTPicker1.Value = CDate(ucTextValue(Index).Text)
        iText = Index
    End If
End Sub

Private Sub ucTextValue_Validate(Index As Integer, Cancel As Boolean)
    ucJLDTPicker.MinDate = CDate(ucTextValue(0).Text)
    ucJLDTPicker.MaxDate = CDate(ucTextValue(1).Text)
End Sub

Private Sub ucTxtProperties_Change(Index As Integer)
    Select Case Index
        Case 0 'Columnas
            ucJLDTPicker.ColsPicker = ucTxtProperties(Index).Text
        Case 1 'Calendarios
            ucJLDTPicker.NumberPickers = ucTxtProperties(Index).Text
            ucTxtProperties(0).Text = ucJLDTPicker.ColsPicker
            ucTxtProperties(0).MaxValue = ucJLDTPicker.NumberPickers
    End Select
End Sub
