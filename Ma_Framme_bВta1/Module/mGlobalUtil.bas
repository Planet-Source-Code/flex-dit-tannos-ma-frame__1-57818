Attribute VB_Name = "mGlobalUtil"
'Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
'
'Public Function AppThemed() As Boolean
'
'    On Error Resume Next
'    AppThemed = IsAppThemed()
'    On Error GoTo 0
'
'End Function




Public Sub UtilDrawBackground(ByVal lngHdc As Long, _
                              ByVal colorStart As Long, _
                              ByVal colorEnd As Long, _
                              ByVal lngLeft As Long, _
                              ByVal lngTop As Long, _
                              ByVal lngWidth As Long, _
                              ByVal lngHeight As Long, _
                              Optional ByVal horizontal As Long = 0)


    Dim tR As RECT

    With tR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngLeft + lngWidth
        .Bottom = lngTop + lngHeight
        ' gradient fill vertical:
    End With    'tR
    GradientFillRect lngHdc, tR, colorStart, colorEnd, IIf(horizontal = 0, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub

Public Sub UtilDrawBorderRectangle(ByVal lngHdc As Long, _
                                   ByVal lPenStyle As Long, _
                                   ByVal lColor As Long, _
                                   ByVal lngLeft As Long, _
                                   ByVal lngTop As Long, _
                                   ByVal lngWidth As Long, _
                                   ByVal lngHeight As Long, _
                                   ByVal bInset As Boolean)


    Dim tJ As POINTAPI
    Dim hPen As Long
    Dim hPenOld As Long

    hPen = CreatePen(lPenStyle, 1, lColor)
    hPenOld = SelectObject(lngHdc, hPen)
    MoveToEx lngHdc, lngLeft, lngTop + lngHeight - 1, tJ
    LineTo lngHdc, lngLeft, lngTop
    LineTo lngHdc, lngLeft + lngWidth - 1, lngTop
    LineTo lngHdc, lngLeft + lngWidth - 1, lngTop + lngHeight - 1
    LineTo lngHdc, lngLeft, lngTop + lngHeight - 1
    SelectObject lngHdc, hPenOld
    DeleteObject hPen

End Sub



Private Sub GradientFillRect(ByVal lHDC As Long, _
                             tR As RECT, _
                             ByVal oStartColor As OLE_COLOR, _
                             ByVal oEndColor As OLE_COLOR, _
                             ByVal eDir As GradientFillRectType)

    Dim tTV(0 To 1) As TRIVERTEX
    Dim tGR As GRADIENT_RECT
    Dim hBrush As Long
    Dim lStartColor As Long
    Dim lEndColor As Long

    'Dim lR As Long
    ' Use GradientFill:
    If Not (HasGradientAndTransparency) Then
        lStartColor = TranslateColor(oStartColor)
        lEndColor = TranslateColor(oEndColor)
        setTriVertexColor tTV(0), lStartColor
        tTV(0).x = tR.Left
        tTV(0).y = tR.Top
        setTriVertexColor tTV(1), lEndColor
        tTV(1).x = tR.Right
        tTV(1).y = tR.Bottom
        tGR.UpperLeft = 0
        tGR.LowerRight = 1
        GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
    Else
        ' Fill with solid brush:
        hBrush = CreateSolidBrush(TranslateColor(oEndColor))
        FillRect lHDC, tR, hBrush
        DeleteObject hBrush
    End If

End Sub


Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

' Convert Automation color to Windows color
'--------- Drawing

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub setTriVertexColor(tTV As TRIVERTEX, _
                              ByVal lColor As Long)


    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long

    lRed = (lColor And &HFF&) * &H100&
    lGreen = (lColor And &HFF00&)
    lBlue = (lColor And &HFF0000) \ &H100&
    With tTV
        setTriVertexColorComponent .Red, lRed
        setTriVertexColorComponent .Green, lGreen
        setTriVertexColorComponent .Blue, lBlue
    End With    'tTV

End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, _
                                       ByVal lComponent As Long)

    If (lComponent And &H8000&) = &H8000& Then
        iColor = (lComponent And &H7F00&)
        iColor = iColor Or &H8000
    Else
        iColor = lComponent
    End If

End Sub



Public Property Get dBlendColor(ByVal oColorFrom As OLE_COLOR, _
                                ByVal oColorTo As OLE_COLOR, _
                                Optional ByVal Alpha As Long = 128) As Long

    Dim lSrcR As Long

    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    dBlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property



Public Sub UtilDrawText(ByVal lngHdc As Long, _
                        ByVal sCaption As String, _
                        ByVal lTextX As Long, _
                        ByVal lTextY As Long, _
                        ByVal lTextX1 As Long, _
                        ByVal lTextY1 As Long, _
                        ByVal bEnabled As Boolean, _
                        ByVal color As Long, _
                        ByVal HorizontalAlign As Long)


    Dim rcText As RECT

    SetTextColor lngHdc, TranslateColor(color)
    'Dim lFlags As Long
    If Not bEnabled Then
        SetTextColor lngHdc, GetSysColor(vbGrayText And &H1F&)
    End If

    With rcText
        .Left = lTextX
        .Top = lTextY
        .Right = lTextX1
        .Bottom = lTextY1
    End With

    DrawTextA lngHdc, sCaption, Len(sCaption), rcText, HorizontalAlign Or DT_BOTTOM Or DT_END_ELLIPSIS

    If Not bEnabled Then
        SetTextColor lngHdc, GetSysColor(color)
    End If

End Sub
