Attribute VB_Name = "modAPIs"
'===Types=============================================================================================================
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type
Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type
Public Enum GradientFillRectType
    GRADIENT_FILL_RECT_H = 0
    GRADIENT_FILL_RECT_V = 1
End Enum
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                    lpRect As RECT) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                    lpRect As RECT) As Long

Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                              ByVal HPALETTE As Long, _
                                                              pccolorref As Long) As Long

Public Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, _
                                                    pVertex As TRIVERTEX, _
                                                    ByVal dwNumVertex As Long, _
                                                    pMesh As GRADIENT_RECT, _
                                                    ByVal dwNumMesh As Long, _
                                                    ByVal dwMode As Long) As Long

Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                               lpRect As RECT, _
                                               ByVal hBrush As Long) As Long



Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                               ByVal nWidth As Long, _
                                               ByVal crColor As Long) As Long


Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                              ByVal x As Long, _
                                              ByVal y As Long, _
                                              lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                            ByVal x As Long, _
                                            ByVal y As Long) As Long



Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                  ByVal crColor As Long) As Long




Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, _
                                                ByVal lpStr As String, _
                                                ByVal nCount As Long, _
                                                lpRect As RECT, _
                                                ByVal wFormat As Long) As Long
Public Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, _
                                                ByVal lpStr As Long, _
                                                ByVal nCount As Long, _
                                                lpRect As RECT, _
                                                ByVal wFormat As Long) As Long




Public Const DT_RIGHT = &H2
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1

Public Const DT_TOP = &H0
Public Const DT_BOTTOM = &H8
Public Const DT_VCENTER = &H4

Public Const DT_END_ELLIPSIS = &H8000&




