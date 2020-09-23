Attribute VB_Name = "ModAPI"
Option Explicit

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpstring As String, ByVal nCount As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
    
Public Type UDT_BarTextFont
    FontName As String
    FontSize As Single
    FontItalic As Boolean
    FontBold As Boolean
End Type

Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  'lfFacename(LF_FACESIZE) As Byte
  lfFacename As String * 33
End Type

Private Type Size
        cx As Long
        cy As Long
End Type

Private Const CLR_INVALID = &HFFFF
Private Const PS_SOLID = 0
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y


Public Function DrawBar(rHDC As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, rColor As OLE_COLOR) As Long
    Dim oldPen As Long
    Dim oldBrush As Long
    Dim rBrush As LOGBRUSH
    Dim rtn As Long
    
    rBrush.lbColor = rColor
    rBrush.lbStyle = vbSolid
    rBrush.lbHatch = 0

    oldPen = SelectObject(rHDC, CreatePen(PS_SOLID, 1, rColor))
    If oldPen = 0 Then GoTo Err_Handler
    oldBrush = SelectObject(rHDC, CreateBrushIndirect(rBrush))
    If oldBrush = 0 Then GoTo Err_Handler
    rtn = Rectangle(rHDC, X1, Y1, X2, Y2)
    If rtn = 0 Then GoTo Err_Handler
    rtn = DeleteObject(SelectObject(rHDC, oldPen))
    rtn = DeleteObject(SelectObject(rHDC, oldBrush))
    DrawBar = rtn
    Exit Function
Err_Handler:
    DrawBar = 0
End Function

Public Function DrawBarText(r_HDC As Long, sX As Long, sY As Long, rFont As UDT_BarTextFont, _
                            fColor As OLE_COLOR, Direction As Integer, OutText As String) As Long
    Dim LF As LOGFONT
    Dim hPrevFont As Long
    Dim hNewFont As Long
    Dim oldColor As OLE_COLOR
    Dim rtn As Long
    
    'Create the new font with selected font attributes
    With LF
        .lfFacename = rFont.FontName & vbNullChar
        .lfHeight = -MulDiv((rFont.FontSize), GetDeviceCaps(r_HDC, LOGPIXELSY), 72)
        .lfEscapement = Direction * 10
        .lfItalic = rFont.FontItalic
        .lfWeight = IIf(rFont.FontBold, 700, 400)
    End With
    'Create the new font
    hNewFont = CreateFontIndirect(LF)
    If hNewFont = 0 Then GoTo Err_Handler
    'Select the new font
    hPrevFont = SelectObject(r_HDC, hNewFont)
    If hPrevFont = 0 Then GoTo Err_Handler
    'Set the font Color and save the original Color
    oldColor = SetTextColor(r_HDC, fColor)
    If oldColor = CLR_INVALID Then GoTo Err_Handler
    'Output the text
    rtn = TextOut(r_HDC, sX, sY, OutText, Len(OutText))
    If rtn = 0 Then GoTo Err_Handler
    'Restore the old text Color
    rtn = SetTextColor(r_HDC, oldColor)
    'Restore the original font and delete the newly created font
    rtn = DeleteObject(SelectObject(r_HDC, hPrevFont))
    DrawBarText = rtn
    Exit Function
    
Err_Handler:
    DrawBarText = 0
End Function

Public Function GetTextSize(r_HDC As Long, rFont As UDT_BarTextFont, OutText As String, _
                            T_Width As Long, T_Height As Long) As Long
    Dim LF As LOGFONT
    Dim hPrevFont As Long
    Dim hNewFont As Long
    Dim ts As Size
    Dim rtn As Long
    
    'Create the new font with selected font attributes
    With LF
        .lfFacename = rFont.FontName & vbNullChar
        .lfHeight = -MulDiv((rFont.FontSize), GetDeviceCaps(r_HDC, LOGPIXELSY), 72)
        .lfItalic = rFont.FontItalic
        .lfWeight = IIf(rFont.FontBold, 700, 400)
    End With
    'Create the new font
    hNewFont = CreateFontIndirect(LF)
    If hNewFont = 0 Then GoTo Err_Handler
    'Select the new font
    hPrevFont = SelectObject(r_HDC, hNewFont)
    If hPrevFont = 0 Then GoTo Err_Handler
    'Get the text height and width
    rtn = GetTextExtentPoint32(r_HDC, OutText, Len(OutText), ts)
    'Restore the original font and delete the newly created font
    rtn = DeleteObject(SelectObject(r_HDC, hPrevFont))
    T_Width = ts.cx: T_Height = ts.cy
    GetTextSize = rtn
    Exit Function
    
Err_Handler:
GetTextSize = 0
End Function

Public Function MilsToPixels(rHDC As Long, mils As Single) As Integer
    MilsToPixels = mils * GetDeviceCaps(rHDC, LOGPIXELSY)
End Function
