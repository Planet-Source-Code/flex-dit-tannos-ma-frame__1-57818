VERSION 5.00
Begin VB.UserControl MaFrame 
   Alignable       =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "MaFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'===Declarations======================================================================================================


Public Enum Theme
    xThemeDarkBlue
    xThemeMoney
    xThemeGreen
    xThemeMetallic
    xThemeOffice2003Style1
    xThemeOffice2003Style2
    xThemeOrange
    xThemeTurquoise
    xThemeGray
    xThemeDarkBlue2
End Enum

Public Enum PictureAlign
    [xPictureAlignLeft] = &H0
    [xPictureAlignCenter] = &H1
    [xPictureAlignRight] = &H2
    [xPictureAlignTop] = &H3
    [xPictureAlignBottom] = &H4
End Enum

Public Enum CaptionAlign
    xAlignLeft
    xAlignRight
    xAlignCenter
End Enum

Public Enum FillStyle
    VerticalFading
    HorizontalFading
End Enum

Public Enum HeaderStyleSize
    Small
    Medium
    Large
End Enum
Public Enum TraceBorderStyle
    None
    Solid
    Dot
    Dash
    DashDot
End Enum


'=====================================================================================================================


'===Constantes=========================================================================================================

'valeurs par defaut des propriétés

Private Const m_def_sCaption As String = "MaFrame"

Private Const m_def_iTheme As Integer = xThemeDarkBlue

Private Const m_def_lHeaderTextColor As Long = vbWhite     'couleur texte header
'Private Const m_def_iHeaderHeight As Integer = 220               ' Height par defaut du header
Private Const m_def_lFadeColor As Long = vbButtonFace       'couleur utilisée pour le gradient (container)
Private Const m_def_UseCustomColors As Boolean = False
Private Const m_def_HeaderSize As Integer = Medium
'Private Const m_def_lPictureMaskColor As Long = &HC0C0C0        'Standard gray color as Transparany color
'Private Const m_def_bUseMaskColor As Boolean = True
Private Const m_def_BorderStyle As Long = Solid


Private Const m_def_iCornerRadius As Integer = 10   'angle utilise pour les angles arrondis
Private Const m_def_bRoundCorner As Boolean = False



Private Const CLR_INVALID As Integer = -1

'=====================================================================================================================

'===Propriétés========================================================================================================

Private m_HeaderTextColor As OLE_COLOR   'couleur texte header
Private m_lHeaderBackColor As OLE_COLOR   'couleur arr.plan du header
Private m_enmTheme As Theme               'theme
Private m_FadeColor As OLE_COLOR
Private m_HeaderSize As HeaderStyleSize
Private m_CaptionAlign As CaptionAlign
Private m_GradientStyle As FillStyle
Private m_sCaption As String                           'police du header
Private m_enmPictureAlign As PictureAlign
Private m_Picture As StdPicture
Private m_PictureAlign As PictureAlign
Private m_PictureSize As Long
Private m_bUseMaskColor As Boolean
Private m_BorderStyle As TraceBorderStyle
Private m_iCornerRadius As Integer
Private m_bRoundCorner As Boolean
Private m_bCornerBottomRight As Boolean
Private m_bCornerTopLeft As Boolean
Private m_bCornerTopRight As Boolean

Private m_UseCustomColors As Boolean


Private m_lColorOneNormal As OLE_COLOR
Private m_lColorTwoNormal As OLE_COLOR
Private m_lColorOneSelected As OLE_COLOR
Private m_lColorTwoSelected As OLE_COLOR
Private m_lColorHeaderColorOne As OLE_COLOR
Private m_lColorHeaderColorTwo As OLE_COLOR
Private m_lColorHeaderForeColor As OLE_COLOR
Private m_lColorHotOne As OLE_COLOR
Private m_lColorHotTwo As OLE_COLOR
Private m_lColorBorder As OLE_COLOR

Private m_CaptionFont As Font
Private m_hWnd As Long



'=====================================================================================================================

'=Public Events=======================================================================================================

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)



'=====================================================================================================================


'===Public Properties=================================================================================================


Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawControl
End Property

Public Property Get FadeColor() As OLE_COLOR
    FadeColor = m_FadeColor
End Property

Public Property Let FadeColor(ByVal New_FadeColor As OLE_COLOR)
    m_FadeColor = New_FadeColor
    PropertyChanged "FadeColor"
    DrawControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(bNewValue As Boolean)
    UserControl.Enabled() = bNewValue
    PropertyChanged "Enabled"
    DrawControl
End Property

Public Property Get FrameTheme() As Theme
    FrameTheme = m_enmTheme
End Property

Public Property Let FrameTheme(enmNewTheme As Theme)
    m_enmTheme = enmNewTheme
    GetGradientColors
    DrawControl
    PropertyChanged ("FrameTheme")
End Property


Public Property Get HeaderSize() As HeaderStyleSize
    HeaderSize = m_HeaderSize
End Property

Public Property Let HeaderSize(NewHeaderSize As HeaderStyleSize)
    m_HeaderSize = NewHeaderSize
    PropertyChanged (HeaderSize)
    DrawControl
End Property



Public Property Get BorderStyle() As TraceBorderStyle
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As TraceBorderStyle)
    m_BorderStyle = NewBorderStyle
    PropertyChanged ("BorderStyle")
    DrawControl
End Property


Public Property Get HeaderBackColor() As OLE_COLOR
    HeaderBackColor = m_lHeaderBackColor
End Property

Public Property Let HeaderBackColor(lNewValue As OLE_COLOR)
    m_lHeaderBackColor = lNewValue
    PropertyChanged ("HeaderBackColor")
End Property

Public Property Set Picture(NewPicture As StdPicture)

    Set m_Picture = NewPicture
    PropertyChanged "Picture"
    DrawControl
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property


Public Property Get PictureAlign() As PictureAlign
    PictureAlign = m_enmPictureAlign
End Property

Public Property Let PictureAlign(iNewValue As PictureAlign)
    m_enmPictureAlign = iNewValue
    PropertyChanged "PictureAlign"
End Property

Public Property Let PictureSize(ByVal NewPictureSize As Integer)

    m_PictureSize = NewPictureSize
    PropertyChanged "PictureSize"
    DrawControl
End Property

Public Property Get PictureSize() As Integer
    PictureSize = m_PictureSize
End Property


Public Property Get Caption() As String
    Caption = m_sCaption
End Property

Public Property Let Caption(sNewValue As String)
    m_sCaption = sNewValue
    DrawControl
    PropertyChanged "Caption"
End Property


Public Property Get CaptionFont() As Font
    Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal New_CaptionFont As Font)
    Set m_CaptionFont = New_CaptionFont
    PropertyChanged "CaptionFont"
    DrawControl
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property


Public Property Get CaptionAlign() As CaptionAlign
    CaptionAlign = m_CaptionAlign
End Property


Public Property Let CaptionAlign(ByVal NewCaptionAlign As CaptionAlign)
    m_CaptionAlign = NewCaptionAlign
    PropertyChanged "CaptionAlign"
    DrawControl
End Property


Public Property Get GradientStyle() As FillStyle
    GradientStyle = m_GradientStyle
End Property

Public Property Let GradientStyle(ByVal NewGradientStyle As FillStyle)
    m_GradientStyle = NewGradientStyle
    PropertyChanged "GradientStyle"
    DrawControl
End Property


Public Property Get PictureMaskColor() As OLE_COLOR
    PictureMaskColor = UserControl.MaskColor
End Property

Public Property Let PictureMaskColor(lNewColor As OLE_COLOR)
    UserControl.MaskColor = lNewColor
    PropertyChanged "PictureMaskColor"
End Property

Public Property Get UseMaskColor() As Boolean
    UseMaskColor = m_bUseMaskColor
End Property

Public Property Let UseMaskColor(bNewValue As Boolean)
    m_bUseMaskColor = bNewValue
    PropertyChanged "UseMaskColor"
End Property

Public Property Get CornerRadius() As Integer
    CornerRadius = m_iCornerRadius
End Property

Public Property Let CornerRadius(iNewValue As Integer)
    m_iCornerRadius = iNewValue
    Refresh
    PropertyChanged "CornerRadius"
End Property

Public Property Get CornerBottomLeft() As Boolean
    CornerBottomLeft = m_bRoundCorner
End Property

Public Property Let CornerBottomLeft(ByVal bCornerBottomLeft As Boolean)
    m_bRoundCorner = bCornerBottomLeft
    PropertyChanged ("CornerBottomLeft")
    DrawControl
End Property


Public Property Get CornerBottomRight() As Boolean
    CornerBottomRight = m_bCornerBottomRight
End Property

Public Property Let CornerBottomRight(ByVal bCornerBottomRight As Boolean)
    m_bCornerBottomRight = bCornerBottomRight
    PropertyChanged ("CornerBottomRight")
End Property


Public Property Get CornerTopLeft() As Boolean

    CornerTopLeft = m_bCornerTopLeft

End Property

Public Property Let CornerTopLeft(ByVal bCornerTopLeft As Boolean)
    m_bCornerTopLeft = bCornerTopLeft
    PropertyChanged ("CornerTopLeft")
End Property


Public Property Get CornerTopRight() As Boolean
    CornerTopRight = m_bCornerTopRight
End Property

Public Property Let CornerTopRight(ByVal bCornerTopRight As Boolean)
    m_bCornerTopRight = bCornerTopRight
    PropertyChanged ("CornerTopRight")
End Property


Public Property Get HeaderTextColor() As OLE_COLOR
    HeaderTextColor = m_HeaderTextColor
End Property

Public Property Let HeaderTextColor(ByVal New_HeaderTextColor As OLE_COLOR)
    m_HeaderTextColor = New_HeaderTextColor
    DrawControl
    PropertyChanged "HeaderTextColor"
End Property

Public Property Get UseCustomColors() As Boolean
    UseCustomColors = m_UseCustomColors
End Property

Public Property Let UseCustomColors(ByVal New_UseCustomColors As Boolean)
    m_UseCustomColors = New_UseCustomColors
    DrawControl
    PropertyChanged "UseCustomColors"
End Property



Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

' Convert Automation color to Windows color
'--------- Drawing
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function



Private Sub DrawControl()

    Dim pDC As Long
    Dim lHDC As Long
    Dim tR As RECT
    Dim border As Long
    Dim iStylePictureOffset As Long
    iStylePictureOffset = 4

    '        DoEvents
    UserControl.Cls
    picHeader.Cls


    Select Case m_HeaderSize
    Case Small
        picHeader.Height = 225
    Case Medium
        picHeader.Height = 335
    Case Large
        picHeader.Height = 600
    End Select


    Select Case m_BorderStyle
    Case None
        border = 5
    Case Solid
        border = 0
    Case Dot
        border = 2
    Case Dash
        border = 1
    Case DashDot
        border = 3
    End Select


    picHeader.Refresh

    '--parametres par defaut
    lHDC = UserControl.hdc
    pDC = picHeader.hdc


    '-- recupere userControl Rect
    GetItemWindowRect tR


    '-- UserControl Background
    If m_UseCustomColors Then
        UtilDrawBackground lHDC, Me.BackColor, m_FadeColor, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, m_GradientStyle    'False
    Else
        UtilDrawBackground lHDC, Me.BackColor, Me.BackColor, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, m_GradientStyle    'False
    End If

    '--gradient sur le header
    '--ensuite on dessine un cadre autour
    UtilDrawBackground pDC, m_lColorOneNormal, m_lColorTwoNormal, 0, 0, picHeader.Width, (picHeader.ScaleHeight / Screen.TwipsPerPixelY), 1
    UtilDrawBorderRectangle pDC, border, m_lColorBorder, 0, 0, (picHeader.ScaleWidth \ Screen.TwipsPerPixelX), (picHeader.ScaleHeight \ Screen.TwipsPerPixelY), False

    '-- dessin de la bordure du controle
    UtilDrawBorderRectangle lHDC, border, m_lColorBorder, tR.Left, tR.Top, tR.Right, tR.Bottom, False

    '--dessine le texte
    pDrawCaption


'    If Not m_Picture Is Nothing Then
'        If Picture <> 0 Then
'            Dim ix As Long, iy As Long
'            If m_PictureAlign = xPictureAlignCenter Then
'                ix = (picHeader.ScaleWidth - m_PictureSize) / 2
'                iy = (picHeader.ScaleHeight - m_PictureSize) / 2
'            ElseIf m_PictureAlign = xPictureAlignBottom Then
'                ix = (picHeader.ScaleWidth - m_PictureSize) / 2
'                iy = picHeader.ScaleHeight - m_PictureSize - iStylePictureOffset
'            ElseIf m_PictureAlign = xPictureAlignTop Then
'                ix = (picHeader.ScaleWidth - m_PictureSize) / 2
'                iy = iStylePictureOffset
'            ElseIf m_PictureAlign = xPictureAlignLeft Then
'                ix = iStylePictureOffset
'                iy = (picHeader.ScaleHeight - m_PictureSize) / 2
'            ElseIf m_PictureAlign = xPictureAlignRight Then
'                ix = picHeader.ScaleWidth - m_PictureSize - iStylePictureOffset
'                iy = (picHeader.ScaleHeight - m_PictureSize) / 2
'            End If
'
'            picHeader.PaintPicture m_Picture, ix, iy, m_PictureSize, m_PictureSize
'        End If
'    End If
End Sub


Private Sub pDrawCaption()
    On Error Resume Next
    '-- Caption
    Dim RCCap As RECT
    Dim oColor As OLE_COLOR
    Dim bChange As Boolean
    Dim lHorCaptionFont As Long
    Dim bFnt As StdFont
    Set bFnt = Font

    Set Font = Me.CaptionFont

    '--on recupere ForeColor
    oColor = UserControl.ForeColor
      
'    -- Caption Rectangle
    RCCap.Top = 2
    RCCap.Bottom = picHeader.ScaleHeight \ Screen.TwipsPerPixelY '(TextHeight("W") \ Screen.TwipsPerPixelY)
    RCCap.Right = picHeader.ScaleWidth \ Screen.TwipsPerPixelX
    RCCap.Left = 5

    Select Case m_CaptionAlign
    Case xAlignLeft
        lHorCaptionFont = DT_LEFT
    Case xAlignRight
        lHorCaptionFont = DT_RIGHT
    Case xAlignCenter
        lHorCaptionFont = DT_CENTER
    End Select
    
    UserControl.ForeColor = HeaderTextColor
    Set picHeader.Font = Me.CaptionFont
    UtilDrawText picHeader.hdc, m_sCaption, RCCap.Left, RCCap.Top, (RCCap.Right - RCCap.Left), RCCap.Bottom - RCCap.Top, Me.Enabled, HeaderTextColor, lHorCaptionFont
    UserControl.ForeColor = oColor
    Set Font = bFnt
End Sub


Private Sub GetItemWindowRect(tR As RECT)
    GetClientRect m_hWnd, tR

End Sub


Sub GetGradientColors()
    
    Select Case m_enmTheme
    Case xThemeDarkBlue
        m_lColorOneNormal = RGB(137, 170, 224)
        m_lColorTwoNormal = RGB(7, 33, 100)
        m_lColorBorder = RGB(1, 45, 150)
        m_lColorHeaderColorOne = RGB(81, 128, 208)
        m_lColorHeaderColorTwo = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
        Me.BackColor = RGB(142, 179, 231)
        HeaderTextColor = RGB(215, 230, 251)

    Case xThemeGreen
        m_lColorOneNormal = RGB(228, 235, 200)
        m_lColorTwoNormal = RGB(175, 194, 142)
        m_lColorBorder = RGB(100, 144, 88)
        m_lColorHeaderColorOne = RGB(165, 182, 121)
        m_lColorHeaderColorTwo = dBlendColor(RGB(99, 122, 68), vbBlack, 200)
        Me.BackColor = RGB(233, 244, 207)
        HeaderTextColor = RGB(100, 144, 88)

    Case xThemeOffice2003Style2
        m_lColorOneNormal = RGB(249, 249, 255)
        m_lColorTwoNormal = RGB(159, 157, 185)
        m_lColorBorder = RGB(124, 124, 148)
        m_lColorHeaderColorOne = RGB(81, 128, 208)
        m_lColorHeaderColorTwo = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
        Me.BackColor = RGB(253, 250, 255)
        HeaderTextColor = RGB(110, 109, 143)

    Case xThemeMetallic
        m_lColorOneNormal = RGB(219, 220, 232)
        m_lColorTwoNormal = RGB(149, 147, 177)
        m_lColorBorder = RGB(119, 118, 151)
        m_lColorHeaderColorOne = RGB(163, 162, 87)
        Me.BackColor = RGB(232, 232, 232)
        HeaderTextColor = RGB(119, 118, 151)

    Case xThemeOrange
        m_lColorOneNormal = RGB(255, 122, 0)
        m_lColorTwoNormal = RGB(130, 0, 0)
        m_lColorBorder = RGB(139, 0, 0)
        m_lColorHeaderColorOne = &HE0E0E0
        m_lColorHeaderColorTwo = dBlendColor(RGB(112, 111, 145), vbBlack, 200)
        Me.BackColor = RGB(255, 222, 173)
        HeaderTextColor = RGB(255, 222, 173)

    Case xThemeTurquoise
        m_lColorOneNormal = RGB(72, 209, 204)
        m_lColorTwoNormal = RGB(43, 103, 109)
        m_lColorBorder = RGB(65, 131, 111)
        m_lColorHeaderColorOne = &HE0E0E0
        m_lColorHeaderColorTwo = dBlendColor(RGB(112, 111, 145), vbBlack, 200)
        Me.BackColor = RGB(224, 255, 255)
        HeaderTextColor = RGB(233, 250, 248)

    Case xThemeGray
        m_lColorOneNormal = RGB(192, 192, 192)
        m_lColorTwoNormal = RGB(51, 51, 51)
        m_lColorBorder = RGB(51, 51, 51)
        m_lColorHeaderColorOne = &HE0E0E0
        m_lColorHeaderColorTwo = dBlendColor(RGB(112, 111, 145), vbBlack, 200)
        Me.BackColor = RGB(235, 235, 235)
        HeaderTextColor = RGB(235, 235, 235)

    Case xThemeDarkBlue2
        m_lColorOneNormal = RGB(81, 128, 208)
        m_lColorTwoNormal = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
        m_lColorBorder = RGB(0, 45, 150)
        m_lColorHeaderColorOne = RGB(81, 128, 208)
        m_lColorHeaderColorTwo = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
        Me.BackColor = RGB(142, 179, 231)
        HeaderTextColor = vbRed

    Case xThemeMoney
        m_lColorOneNormal = RGB(160, 160, 160)
        m_lColorTwoNormal = dBlendColor(RGB(90, 90, 90), vbBlack, 230)
        m_lColorBorder = RGB(68, 68, 68) 'RGB(65, 65, 65)
        m_lColorHeaderColorOne = RGB(81, 128, 208)
        m_lColorHeaderColorTwo = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
        Me.BackColor = RGB(112, 112, 112)
        HeaderTextColor = vbWhite 'RGB(255, 223, 127)
        
    Case xThemeOffice2003Style1
        m_lColorOneNormal = RGB(209, 227, 251)
        m_lColorTwoNormal = RGB(106, 140, 203)
        m_lColorBorder = RGB(0, 0, 128)
        m_lColorHeaderColorOne = RGB(81, 128, 208)
        m_lColorHeaderColorTwo = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
        Me.BackColor = RGB(255, 255, 255)
        HeaderTextColor = RGB(110, 109, 143)
                
    End Select
        End Sub


Private Sub UserControl_Initialize()
    m_hWnd = UserControl.hwnd
End Sub

Private Sub UserControl_InitProperties()

    Set UserControl.Font = Ambient.Font
    m_HeaderTextColor = m_def_lHeaderTextColor
    m_FadeColor = m_def_lFadeColor
    m_enmTheme = m_def_iTheme
    m_HeaderSize = m_def_HeaderSize
    m_BorderStyle = m_def_BorderStyle
    m_sCaption = UserControl.Extender.Name
    m_iCornerRadius = m_def_iCornerRadius
    m_bRoundCorner = m_def_bRoundCorner
    m_PictureSize = 16
    Set m_Picture = LoadPicture
    Set m_CaptionFont = Ambient.Font
    GetGradientColors
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        UserControl.ForeColor = .ReadProperty("ForeColor", &H80000008)
        UserControl.ScaleMode = .ReadProperty("ScaleMode", 1)
        UserControl.BackColor = .ReadProperty("BackColor", &HFFFFFF)
    End With
    m_UseCustomColors = PropBag.ReadProperty("UseCustomColors", m_def_UseCustomColors)
    m_enmTheme = PropBag.ReadProperty("FrameTheme", m_def_iTheme)
    m_HeaderSize = PropBag.ReadProperty("HeaderSize", m_def_HeaderSize)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_PictureSize = PropBag.ReadProperty("PictureSize", 16)
    m_sCaption = PropBag.ReadProperty("Caption", m_def_sCaption)
        m_iCornerRadius = PropBag.ReadProperty("CornerRadius", m_def_iCornerRadius)
        m_bRoundCorner = PropBag.ReadProperty("CornerBottomLeft", m_def_bRoundCorner)
    m_PictureAlign = PropBag.ReadProperty("PictureAlign", xPictureAlignLeft)
    m_HeaderTextColor = PropBag.ReadProperty("HeaderTextColor", m_def_lHeaderTextColor)
    m_FadeColor = PropBag.ReadProperty("FadeColor", m_def_lFadeColor)
    Set m_CaptionFont = PropBag.ReadProperty("CaptionFont", Ambient.Font)
    GetGradientColors
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
        Call .WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
        Call .WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
    End With
    Call PropBag.WriteProperty("Caption", m_sCaption, m_def_sCaption)
    Call PropBag.WriteProperty("UseCustomColors", m_UseCustomColors, m_def_UseCustomColors)
    Call PropBag.WriteProperty("FrameTheme", m_enmTheme, m_def_iTheme)
    Call PropBag.WriteProperty("HeaderSize", m_HeaderSize, m_def_HeaderSize)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("PictureSize", m_PictureSize, 16)
    Call PropBag.WriteProperty("PictureAlign", m_PictureAlign, 0)
    Call PropBag.WriteProperty("CornerRadius", m_iCornerRadius, m_def_iCornerRadius)
    Call PropBag.WriteProperty("CornerBottomLeft", m_bRoundCorner, m_def_bRoundCorner)
    Call PropBag.WriteProperty("HeaderTextColor", m_HeaderTextColor, m_def_lHeaderTextColor)
    Call PropBag.WriteProperty("FadeColor", m_FadeColor, m_def_lFadeColor)
    Call PropBag.WriteProperty("CaptionFont", m_CaptionFont, Ambient.Font)
End Sub


Private Sub UserControl_Paint()
    Call DrawControl
End Sub

Private Sub UserControl_Resize()
    DrawControl
End Sub

