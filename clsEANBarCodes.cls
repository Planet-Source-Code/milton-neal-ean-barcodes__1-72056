VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEANBarCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' *************************************************************************
' PROJECT:  EAN/UPC Barcode Module
' AUTHOR:   Milton Neal, Perth Western Australia. miltonneal@arach.net.au.
' DATE:     1st May 2009
' VERSION:  1.00
' UPDATES:
' *************************************************************************

Public Enum eSymbology
    EAN13 = 1
    EAN13_2 = 2
    EAN13_5 = 3
    EAN8 = 4
    EAN8_2 = 5
    EAN8_5 = 6
    UPCA = 7
    UPCA_2 = 8
    UPCA_5 = 9
    UPCE = 10
    UPCE_2 = 11
    UPCE_5 = 12
End Enum

'Dim HRbartext As String

Dim BC_ValidCharacters As String * 10
Dim EANUPC_LHO(9) As String
Dim EANUPC_LHE(9) As String
Dim EANUPC_RH(9) As String
Dim EANUPC_Parity(9) As String
Dim UPCe_Parity(9) As String
Dim Plus2_Parity(3) As String
Dim Plus5_Parity(9) As String

Private cHDC As Long                'Handle to the output device
Private cSymbology As eSymbology    'Barcode symbology to use
Private cBarX As Single             'Width of the bar in printer pixels
Private cBarMultiplier As Integer   'Bar width multiplier
Private cBarColor As OLE_COLOR      'Color to draw the barcode
Private cBarCodeData As String      'Barcode data
Private cBarRotation As Integer     'Direction to print/display barcode
Private cBarCodeTop As Long         'Top start position in pixels
Private cBarCodeLeft As Long        'Left start position in pixel
Private cBarCodeHeight As Long      'Height of the bars (not including text)in pixels
Private cBarTextGap As Long         'Gap between the barcode and the HR text in pixels

Private cHumanReadable As Boolean           'Human readable or barcode only
Private cBarTextFont As UDT_BarTextFont     'Holds the font information
Private cAutoFontSize As Boolean            'Auto adjest the HR test font size to fit

Dim BarStartX As Long           'X point to start barcode
Dim BarStartY As Long           'Y point to start barcode
Dim NarrowBar As Single         'Narrow bar
Dim WideBar As Single           'Wide bar
Dim GuardBarHeight As Long      'Height of the gaurd bars in pixels
Dim BarCodeWidth As Long        'Width of barcode in pixels
Dim SupplementWidth As Long     'Width of the supplemental barcode in pixels
Dim SupplementGap As Long       'Gap between the barcode and supplemental in pixels
Dim NumberOfBars As Integer     'Number of bars between the guard bars
Dim BarBit As String * 1        'Holds the "width" value of a bar

Dim UPCAOffset As Integer       'Offeset for drawing UPCA HR text
Dim UPCEOffset As Integer       'Offset for drawing UPCE HR check digit
Dim TextGap As Integer          'Size of the gap between the guard bars for the HR text(pixels)
Dim TextStartX As Long          'X point for text
Dim TextStartY As Long          'Y point for text
Dim HRTextHeight As Long        'HR text height in pixels
Dim HRTextWidth As Long         'HR text width in pixels
Dim LeftDigitWidth As Long      'Width of the left character in pixels
Dim RightDigitWidth As Long     'Width of the right character in pixels (for UPC only)
Dim HRLeftDigit As String       'Character left of start guard bar
Dim HRRightDigit As String      'Character right of end guard bar
Dim HRLeftText As String        'Left side HR text
Dim HRRightText As String       'Right side HR text
Dim HRSupplementText As String  'HR text for the supplemental
Dim SupplementData As String    'Holds the encoded data for the supplental barcode
Dim EncodedData As String       'Holds the encoded data for creating barcode

'Private Const BARSPERCHAR As Integer = 7

Private cErrorNumber As Long
Private cErrDescription As String
Const B_ErrNoData As Long = 6940
Const S_ErrNoData As String = "No barcode data specified."
Const B_ErrSymbology As Long = 6950
Const S_ErrSymbology As String = "Invalid or no symbology selected."
Const B_ErrDraw As Long = 6960
Const S_ErrDraw As String = "Error generating barcode."
Const B_ErrInvalidChar As Long = 6970
Const S_ErrInvalidChar As String = "Invalid character in the barcode."
Const B_ErrInvalidLength As Long = 6980
Const S_ErrInvalidLength As String = "Invalid code length for symbology."
Const B_ErrInvalidControl As Long = 6990
Const S_ErrInvalidControl As String = "Invalid control."

'===============================================
'PROPERTIES
'===============================================

Public Property Let BarcodeOutput(Dest As Object)
    On Error GoTo err
    cHDC = Dest.hdc
    On Error GoTo 0
    Exit Property
err:
    cErrorNumber = B_ErrInvalidControl
    cErrDescription = S_ErrInvalidControl
    err.Raise cErrorNumber, "clsEANBarCode.BarcodeOutput", cErrDescription
End Property

Public Property Get Symbology() As eSymbology
    Symbology = cSymbology
End Property
Public Property Let Symbology(bSym As eSymbology)
    cSymbology = bSym
End Property

Public Property Get BarXFactor() As Single
    BarXFactor = cBarX
End Property
Public Property Let BarXFactor(xbWidth As Single)
    cBarX = xbWidth
End Property

Public Property Get BarMultiplier() As Integer
    BarMultiplier = cBarMultiplier
End Property
Public Property Let BarMultiplier(bMulti As Integer)
    If bMulti > 10 Then bMulti = 10
    If bMulti < 1 Then bMulti = 1
    cBarMultiplier = bMulti
End Property

Public Property Get BarRotation() As Integer
    BarRotation = cBarRotation
End Property
Public Property Let BarRotation(bRotation As Integer)
    If bRotation <> 0 And bRotation <> 90 _
    And bRotation <> 180 And bRotation <> 270 Then
        cBarRotation = 0
    Else: cBarRotation = bRotation
    End If
End Property

Public Property Get BarCodeData() As String
    BarCodeData = cBarCodeData
End Property

Public Property Let BarCodeData(bcData As String)
    Dim iCntr As Integer
    Dim sDataChar As String
    Dim iAscValue As Integer
    
    bcData = Trim(bcData)
    
    For iCntr = 1 To Len(bcData)
        sDataChar = Mid(bcData, iCntr, 1)
        If InStr(1, BC_ValidCharacters, sDataChar) = 0 Then
            cErrorNumber = B_ErrInvalidChar
            cErrDescription = S_ErrInvalidChar
            err.Raise cErrorNumber, "clsEANBarCode.BarCodeData", cErrDescription
            Exit Property
        End If
    Next iCntr
    Select Case cSymbology
        Case EAN13
            If Len(bcData) <> 12 Then GoTo Raise_Error
            EncodeEAN13 bcData
            
        Case EAN13_2
            If Len(bcData) <> 14 Then GoTo Raise_Error
            EncodeEAN13 Mid(bcData, 1, 12)
            EncodePlus2 Right(bcData, 2)
            
        Case EAN13_5
            If Len(bcData) <> 17 Then GoTo Raise_Error
            EncodeEAN13 Mid(bcData, 1, 12)
            EncodePlus5 Right(bcData, 5)
            
        Case EAN8
            If Len(bcData) <> 7 Then GoTo Raise_Error
            EncodeEAN8 bcData
            
        Case EAN8_2
            If Len(bcData) <> 9 Then GoTo Raise_Error
            EncodeEAN8 Mid(bcData, 1, 7)
            EncodePlus2 Right(bcData, 2)
            
        Case EAN8_5
            If Len(bcData) <> 12 Then GoTo Raise_Error
            EncodeEAN8 Mid(bcData, 1, 7)
            EncodePlus5 Right(bcData, 5)
            
        Case UPCA
            If Len(bcData) <> 11 Then GoTo Raise_Error
            EncodeUPCA bcData
            
        Case UPCA_2
            If Len(bcData) <> 13 Then GoTo Raise_Error
            EncodeUPCA Mid(bcData, 1, 11)
            EncodePlus2 Right(bcData, 2)
            
        Case UPCA_5
            If Len(bcData) <> 16 Then GoTo Raise_Error
            EncodeUPCA Mid(bcData, 1, 11)
            EncodePlus5 Right(bcData, 5)
            
        Case UPCE
            If Len(bcData) <> 6 Then GoTo Raise_Error
            EncodeUPCE bcData
            
        Case UPCE_2
            If Len(bcData) <> 8 Then GoTo Raise_Error
            EncodeUPCE Mid(bcData, 1, 6)
            EncodePlus2 Right(bcData, 2)
            
        Case UPCE_5
            If Len(bcData) <> 11 Then GoTo Raise_Error
            EncodeUPCE Mid(bcData, 1, 6)
            EncodePlus5 Right(bcData, 5)
    End Select
    cBarCodeData = bcData
    Exit Property
    
Raise_Error:
    cErrorNumber = B_ErrInvalidLength
    cErrDescription = S_ErrInvalidLength
    err.Raise cErrorNumber, "clsEANBarCode.BarCodeData", cErrDescription
End Property

Public Property Get BarColor() As OLE_COLOR
    BarColor = cBarColor
End Property

Public Property Let BarColor(bColor As OLE_COLOR)
    cBarColor = bColor
End Property

Public Property Get BarCodeX() As Long
    BarCodeX = cBarCodeLeft
End Property

Public Property Let BarCodeX(X As Long)
    cBarCodeLeft = X
End Property

Public Property Get BarCodeY() As Long
    BarCodeY = cBarCodeTop
End Property

Public Property Let BarCodeY(Y As Long)
    cBarCodeTop = Y
End Property

Public Property Get BarcodeHeight() As Long
    BarcodeHeight = cBarCodeHeight
End Property

Public Property Let BarcodeHeight(bHeight As Long)
    If bHeight < 8 Then bHeight = 8
    cBarCodeHeight = bHeight
    GuardBarHeight = bHeight + (bHeight / 10)
End Property

Public Property Get BarTextGap() As Long
    BarTextGap = cBarTextGap
End Property

Public Property Let BarTextGap(bGap As Long)
    cBarTextGap = bGap
End Property

Public Property Get HRText() As Boolean
    HRText = cHumanReadable
End Property

Public Property Let HRText(bReadable As Boolean)
    cHumanReadable = bReadable
End Property

Public Property Get AutoTextFont() As Boolean
    AutoTextFont = cAutoFontSize
End Property
Public Property Let AutoTextFont(bAutoFS As Boolean)
    cAutoFontSize = bAutoFS
End Property

Public Property Get TotalBarWidth() As Long
    Dim BarH As Integer
    
    If cBarCodeData = "" Then
        cErrorNumber = B_ErrNoData
        cErrDescription = S_ErrNoData
        err.Raise cErrorNumber, "clsEANBarCode.BarWidth", cErrDescription
        Exit Property
    End If
   Call GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
    Select Case cBarRotation
        Case 0, 180
            TotalBarWidth = BarCodeWidth
        Case 90, 270
            If cHumanReadable = True Then
                BarH = cBarCodeHeight + HRTextHeight + cBarTextGap
                TotalBarWidth = IIf(BarH > GuardBarHeight, BarH, GuardBarHeight)
            Else: TotalBarWidth = GuardBarHeight
            End If
    End Select
End Property

Public Property Get TotalBarHeight() As Long
    Dim BarH As Integer
    
    If cBarCodeData = "" Then
        cErrorNumber = B_ErrNoData
        cErrDescription = S_ErrNoData
        err.Raise cErrorNumber, "clsEANBarCode.BarWidth", cErrDescription
        Exit Property
    End If
    Call GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
    Select Case cBarRotation
        Case 0, 180
            If cHumanReadable = True Then
                BarH = cBarCodeHeight + HRTextHeight + cBarTextGap
                TotalBarHeight = IIf(BarH > GuardBarHeight, BarH, GuardBarHeight)
            Else: TotalBarHeight = GuardBarHeight
            End If
        Case 90, 270
            TotalBarHeight = BarCodeWidth
    End Select
End Property

'=====================================
'PUBLIC FUNCTIONS
'=====================================

Public Function BarTextFont(bFont As String, bFontSize As Single, bFontBold As Boolean, bFontItalic As Boolean)
    If bFont = "" Then
        bFont = cBarTextFont.FontName
        bFontSize = cBarTextFont.FontSize
        bFontBold = cBarTextFont.FontBold
        bFontItalic = cBarTextFont.FontItalic
    Else
        cBarTextFont.FontName = bFont
        cBarTextFont.FontSize = bFontSize
        cBarTextFont.FontBold = bFontBold
        cBarTextFont.FontItalic = bFontItalic
    End If
End Function

Public Function DrawBarCode()
If cHDC = 0 Then Exit Function
    'Check that a symbology has been set
    If cSymbology < 1 Or cSymbology > 12 Then
        cErrorNumber = B_ErrSymbology
        cErrDescription = S_ErrSymbology
        err.Raise cErrorNumber, "clsEANBarCode.Draw_BarCode", cErrDescription
        Exit Function
    End If
    'Check that the barcode data has been set
    If cBarCodeData = "" Then
        cErrorNumber = B_ErrNoData
        cErrDescription = S_ErrNoData
        err.Raise cErrorNumber, "clsEANBarCode.BarWidth", cErrDescription
        Exit Function
    End If
    
    Select Case cBarRotation
        Case 0
            Draw_Normal
        Case 90
            Draw_90
        Case 180
            Draw_180
        Case 270
            Draw_270
    End Select
Exit Function

Err_Handler:
    cErrorNumber = B_ErrInvalidControl
    cErrDescription = S_ErrInvalidControl
    err.Raise cErrorNumber, "clsEANBarCode.Draw_BarCode", cErrDescription
End Function

'======================================
'PRIVATE ROUTINES
'======================================

Private Sub Draw_Normal()

    Dim NextBar As Single       'Postion to start the next bar
    Dim BarH As Integer         'Bar height
    Dim iCntr As Integer        'Loop counter
    Dim bColor As OLE_COLOR     'Current bar color
    Dim rtn As Long
    
    'DRAW THE BARCODE
    '================
    'Get the current height and width of the HR text
    rtn = GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
    If rtn = 0 Then GoTo Raise_Error
    BarH = cBarCodeHeight
    BarStartX = cBarCodeLeft: BarStartY = cBarCodeTop
    bColor = vbWhite
    NextBar = BarStartX + LeftDigitWidth
    For iCntr = 1 To Len(EncodedData)
        'Toggle the color
        If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
        BarBit = Mid(EncodedData, iCntr, 1)
        Select Case BarBit
            '"T" & "S" are embedded in the encoded data and are used to toggle between bar heights
            Case "T"
                BarH = GuardBarHeight
                If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
                
            Case "S"
                BarH = cBarCodeHeight
                If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
                
            'Narrow Bar
            Case "1"
                If bColor <> vbWhite Then
                    rtn = DrawBar(cHDC, CLng(NextBar), BarStartY, CLng(NextBar + NarrowBar), BarStartY + BarH, bColor)
                    If rtn = 0 Then GoTo Raise_Error
                End If
                NextBar = NextBar + NarrowBar

              'Wide Bar
            Case "2", "3", "4"
                WideBar = NarrowBar * BarBit
                If bColor <> vbWhite Then
                    rtn = DrawBar(cHDC, CLng(NextBar), BarStartY, CLng(NextBar + WideBar), BarStartY + BarH, bColor)
                    If rtn = 0 Then GoTo Raise_Error
                End If
                NextBar = NextBar + WideBar
            'Gap for supplement 2 or supplement 5
            Case "9"
                NextBar = NextBar + SupplementGap
                BarStartY = cBarCodeTop + HRTextHeight
                BarH = GuardBarHeight - HRTextHeight
        End Select
    Next iCntr
    
   'DRAW THE HUMAN READABLE TEXT
   '============================
    If cHumanReadable Then
        TextStartY = cBarCodeTop + cBarCodeHeight + cBarTextGap
        TextStartX = cBarCodeLeft
        'Draw the first digit left of the left guard bars
        If HRLeftDigit <> "" Then
            rtn = DrawBarText(cHDC, TextStartX, TextStartY - 3, cBarTextFont, cBarColor, cBarRotation, HRLeftDigit)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the left hand text right of the left guard bars
        TextStartX = TextStartX + LeftDigitWidth + ((3 + UPCAOffset) * NarrowBar)
        rtn = GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
        'If rtn = 0 Then GoTo Raise_Error
        rtn = DrawBarText(cHDC, TextStartX + ((TextGap - HRTextWidth) / 2), TextStartY, cBarTextFont, _
                            cBarColor, cBarRotation, HRLeftText)
        If rtn = 0 Then GoTo Raise_Error
        'Draw the right hand text right of the centre guard bars
        If HRRightText <> "" Then
            TextStartX = TextStartX + TextGap + (5 * NarrowBar)
            rtn = GetTextSize(cHDC, cBarTextFont, HRRightText, HRTextWidth, HRTextHeight)
            If rtn = 0 Then GoTo Raise_Error
            rtn = DrawBarText(cHDC, TextStartX + ((TextGap - HRTextWidth) / 2), TextStartY, cBarTextFont, _
                            cBarColor, cBarRotation, HRRightText)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the check digit to the right of the right hand guard bars (UPC only)
        If HRRightDigit <> "" Then
            TextStartX = TextStartX + TextGap + ((3 + UPCEOffset + UPCAOffset) * NarrowBar)
            rtn = DrawBarText(cHDC, TextStartX, TextStartY - 3, cBarTextFont, cBarColor, cBarRotation, HRRightDigit)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the supplemental text above the supplemental bars (if using supplementals)
        If HRSupplementText <> "" Then
            rtn = GetTextSize(cHDC, cBarTextFont, HRSupplementText, HRTextWidth, HRTextHeight)
            If rtn = 0 Then GoTo Raise_Error
            TextStartX = (BarCodeWidth + BarStartX - SupplementWidth) + ((SupplementWidth - HRTextWidth) / 2)
            TextStartY = BarStartY - HRTextHeight
            rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRSupplementText)
            If rtn = 0 Then GoTo Raise_Error
        End If
    End If
Exit Sub
    
Raise_Error:
    cErrorNumber = B_ErrDraw
    cErrDescription = S_ErrDraw
    err.Raise cErrorNumber, "clsEANBarCode.Draw_Normal", cErrDescription
End Sub

Private Sub Draw_90()

    Dim NextBar As Single       'Postion to start the next bar
    Dim BarH As Integer         'Bar height
    Dim iCntr As Integer        'Loop counter
    Dim bColor As OLE_COLOR     'Current bar color
    Dim rtn As Long
    
    'DRAW THE BARCODE
    '================
    'Get the current height and width of the HR text
    rtn = GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
    If rtn = 0 Then GoTo Raise_Error
    BarH = cBarCodeHeight
    BarStartX = cBarCodeLeft: BarStartY = cBarCodeTop + BarCodeWidth
    bColor = vbWhite
    NextBar = BarStartY - LeftDigitWidth
    For iCntr = 1 To Len(EncodedData)
        'Toggle the color
        If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
        BarBit = Mid(EncodedData, iCntr, 1)
        Select Case BarBit
            '"T" & "S" are embedded in the encoded data and are used to toggle between bar heights
            Case "T"
                BarH = GuardBarHeight
                If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
                
            Case "S"
                BarH = cBarCodeHeight
                If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
                
            'Narrow Bar
            Case "1"
                If bColor <> vbWhite Then
                    rtn = DrawBar(cHDC, BarStartX, CLng(NextBar - NarrowBar), BarStartX + BarH, CLng(NextBar), bColor)
                    If rtn = 0 Then GoTo Raise_Error
                End If
                NextBar = NextBar - NarrowBar
                
            'Wide Bar
            Case "2", "3", "4"
                WideBar = NarrowBar * BarBit
                If bColor <> vbWhite Then
                    rtn = DrawBar(cHDC, BarStartX, CLng(NextBar - WideBar), BarStartX + BarH, CLng(NextBar), bColor)
                    If rtn = 0 Then GoTo Raise_Error
                End If
                NextBar = NextBar - WideBar
                
            'Gap for supplement 2 or supplement 5
            Case "9"
                NextBar = NextBar - SupplementGap
                BarStartX = cBarCodeLeft + HRTextHeight
                BarH = GuardBarHeight - HRTextHeight
            End Select
    Next iCntr
    
    'DRAW THE HUMAN READABLE TEXT
    '============================
    If cHumanReadable Then
        TextStartY = BarStartY
        TextStartX = cBarCodeLeft + cBarCodeHeight + cBarTextGap
        'Draw the first digit left of the left guard bars
        If HRLeftDigit <> "" Then
            rtn = DrawBarText(cHDC, TextStartX - 3, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRLeftDigit)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the left hand text right of the left guard bars
        TextStartY = TextStartY - LeftDigitWidth - ((3 + UPCAOffset) * NarrowBar)
        rtn = GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
        If rtn = 0 Then GoTo Raise_Error
        rtn = DrawBarText(cHDC, TextStartX, TextStartY - ((TextGap - HRTextWidth) / 2), cBarTextFont, _
                            cBarColor, cBarRotation, HRLeftText)
        If rtn = 0 Then GoTo Raise_Error
        'Draw the right hand text right of the centre guard bars
        If HRRightText <> "" Then
            TextStartY = TextStartY - TextGap - (5 * NarrowBar)
            rtn = GetTextSize(cHDC, cBarTextFont, HRRightText, HRTextWidth, HRTextHeight)
            If rtn = 0 Then GoTo Raise_Error
            rtn = DrawBarText(cHDC, TextStartX, TextStartY - ((TextGap - HRTextWidth) / 2), cBarTextFont, _
                            cBarColor, cBarRotation, HRRightText)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the check digit to the right of the right hand guard bars (UPC only)
        If HRRightDigit <> "" Then
            TextStartY = TextStartY - TextGap - ((4 + UPCEOffset + UPCAOffset) * NarrowBar)
            rtn = DrawBarText(cHDC, TextStartX - 3, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRRightDigit)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the supplemental text above the supplemental bars (if using supplementals)
        If HRSupplementText <> "" Then
            rtn = GetTextSize(cHDC, cBarTextFont, HRSupplementText, HRTextWidth, HRTextHeight)
            If rtn = 0 Then GoTo Raise_Error
            TextStartY = (BarStartY - BarCodeWidth + SupplementWidth) - ((SupplementWidth - HRTextWidth) / 2)
            TextStartX = BarStartX - HRTextHeight
            rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRSupplementText)
            If rtn = 0 Then GoTo Raise_Error
        End If
    End If
Exit Sub

Raise_Error:
    cErrorNumber = B_ErrDraw
    cErrDescription = S_ErrDraw
    err.Raise cErrorNumber, "clsEANBarCode.Draw_90", cErrDescription
End Sub

Private Sub Draw_180()

    Dim NextBar As Single       'Postion to start the next bar
    Dim BarH As Integer         'Bar height
    Dim iCntr As Integer        'Loop counter
    Dim bColor As OLE_COLOR     'Current bar color
    Dim rtn As Long
    
    'DRAW THE BARCODE
    '================
    'Get the current height and width of the HR text
    rtn = GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
    If rtn = 0 Then GoTo Raise_Error
    BarH = cBarCodeHeight
    BarStartX = cBarCodeLeft + BarCodeWidth
    BarStartY = cBarCodeTop
    
    bColor = vbWhite
    NextBar = BarStartX - LeftDigitWidth
    For iCntr = 1 To Len(EncodedData)
        'Toggle the color
        If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
        BarBit = Mid(EncodedData, iCntr, 1)
        Select Case BarBit
            '"T" & "S" are embedded in the encoded data and are used to toggle between bar heights
            Case "T"
                BarH = GuardBarHeight
                If cHumanReadable Then
                    BarStartY = cBarCodeTop + HRTextHeight + cBarTextGap - (GuardBarHeight - cBarCodeHeight)
                Else
                    BarStartY = cBarCodeTop
                End If
                If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
                
            Case "S"
                BarH = cBarCodeHeight
                If cHumanReadable Then
                    BarStartY = cBarCodeTop + HRTextHeight + cBarTextGap
                Else
                    BarStartY = cBarCodeTop + (GuardBarHeight - cBarCodeHeight)
                End If
                If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite

            'Narrow Bar
            Case "1"
                If bColor <> vbWhite Then
                    rtn = DrawBar(cHDC, CLng(NextBar - NarrowBar), BarStartY, CLng(NextBar), BarStartY + BarH, bColor)
                    If rtn = 0 Then GoTo Raise_Error
                End If
                NextBar = NextBar - NarrowBar
                
            'Wide Bar
            Case "2", "3", "4"
                WideBar = NarrowBar * BarBit
                If bColor <> vbWhite Then
                    rtn = DrawBar(cHDC, CLng(NextBar - WideBar), BarStartY, CLng(NextBar), BarStartY + BarH, bColor)
                    If rtn = 0 Then GoTo Raise_Error
                End If
                NextBar = NextBar - WideBar
                
            'Gap for supplement 2 or supplement 5
            Case "9"
                NextBar = NextBar - SupplementGap
                If cHumanReadable Then
                    BarStartY = cBarCodeTop + HRTextHeight + cBarTextGap - (GuardBarHeight - cBarCodeHeight)
                Else
                    BarStartY = cBarCodeTop
                End If
                BarH = GuardBarHeight - HRTextHeight
            End Select
    Next iCntr
    
    'DRAW THE HUMAN READABLE TEXT
    '============================
    If cHumanReadable Then
        TextStartY = cBarCodeTop + HRTextHeight
        TextStartX = cBarCodeLeft + BarCodeWidth
        'Draw the first digit left of the left guard bars
        If HRLeftDigit <> "" Then
            rtn = DrawBarText(cHDC, TextStartX, TextStartY + 3, cBarTextFont, cBarColor, cBarRotation, HRLeftDigit)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the left hand text right of the left guard bars
        TextStartX = TextStartX - LeftDigitWidth - ((3 + UPCAOffset) * NarrowBar)
        rtn = GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
        If rtn = 0 Then GoTo Raise_Error
        rtn = DrawBarText(cHDC, TextStartX - ((TextGap - HRTextWidth) / 2), TextStartY, cBarTextFont, _
                            cBarColor, cBarRotation, HRLeftText)
        If rtn = 0 Then GoTo Raise_Error
        'Draw the right hand text right of the centre guard bars
        If HRRightText <> "" Then
            TextStartX = TextStartX - TextGap - (5 * NarrowBar)
            rtn = GetTextSize(cHDC, cBarTextFont, HRRightText, HRTextWidth, HRTextHeight)
            If rtn = 0 Then GoTo Raise_Error
            rtn = DrawBarText(cHDC, TextStartX - ((TextGap - HRTextWidth) / 2), TextStartY, cBarTextFont, _
                            cBarColor, cBarRotation, HRRightText)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the check digit to the right of the right hand guard bars (UPC only)
        If HRRightDigit <> "" Then
            TextStartX = TextStartX - TextGap - ((4 + UPCEOffset + UPCAOffset) * NarrowBar)
            rtn = DrawBarText(cHDC, TextStartX, TextStartY + 3, cBarTextFont, cBarColor, cBarRotation, HRRightDigit)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the supplemental text above the supplemental bars (if using supplementals)
        If HRSupplementText <> "" Then
            rtn = GetTextSize(cHDC, cBarTextFont, HRSupplementText, HRTextWidth, HRTextHeight)
            If rtn = 0 Then GoTo Raise_Error
            TextStartX = (cBarCodeLeft + SupplementWidth) - ((SupplementWidth - HRTextWidth) / 2)
            TextStartY = cBarCodeTop + cBarCodeHeight + HRTextHeight + cBarTextGap
            rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRSupplementText)
            If rtn = 0 Then GoTo Raise_Error
        End If
    End If
Exit Sub

Raise_Error:
    cErrorNumber = B_ErrDraw
    cErrDescription = S_ErrDraw
    err.Raise cErrorNumber, "clsEANBarCode.Draw_180", cErrDescription
End Sub

Private Sub Draw_270()

    Dim NextBar As Single       'Postion to start the next bar
    Dim BarH As Integer         'Bar height
    Dim iCntr As Integer        'Loop counter
    Dim bColor As OLE_COLOR     'Current bar color
    Dim rtn As Long
    
    'DRAW THE BARCODE
    '================
    'Get the current height and width of the HR text
    rtn = GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
    If rtn = 0 Then GoTo Raise_Error
    BarH = cBarCodeHeight
    BarStartX = cBarCodeLeft: BarStartY = cBarCodeTop + LeftDigitWidth
    bColor = vbWhite
    NextBar = BarStartY
    For iCntr = 1 To Len(EncodedData)
        'Toggle the color
        If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
        BarBit = Mid(EncodedData, iCntr, 1)
        Select Case BarBit
            '"T" & "S" are embedded in the encoded data and are used to toggle between bar heights
            Case "T"
                BarH = GuardBarHeight
                If cHumanReadable Then
                    BarStartX = cBarCodeLeft + HRTextHeight + cBarTextGap - (GuardBarHeight - cBarCodeHeight)
                Else
                    BarStartX = cBarCodeLeft
                End If
                If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
                
            Case "S"
                BarH = cBarCodeHeight
                If cHumanReadable Then
                    BarStartX = cBarCodeLeft + HRTextHeight + cBarTextGap
                Else
                    BarStartX = cBarCodeLeft + (GuardBarHeight - cBarCodeHeight)
                End If
                If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
                
            'Narrow Bar
            Case "1"
                If bColor <> vbWhite Then
                    rtn = DrawBar(cHDC, BarStartX, CLng(NextBar), BarStartX + BarH, CLng(NextBar + NarrowBar), bColor)
                    If rtn = 0 Then GoTo Raise_Error
                End If
                NextBar = NextBar + NarrowBar
                
            'Wide Bar
            Case "2", "3", "4"
                WideBar = NarrowBar * BarBit
                If bColor <> vbWhite Then
                    rtn = DrawBar(cHDC, BarStartX, CLng(NextBar), BarStartX + BarH, CLng(NextBar + WideBar), bColor)
                    If rtn = 0 Then GoTo Raise_Error
                End If
                NextBar = NextBar + WideBar
                
            'Gap for supplement 2 or supplement 5
            Case "9"
                NextBar = NextBar + SupplementGap
                If cHumanReadable Then
                    BarStartX = cBarCodeLeft + HRTextHeight + cBarTextGap - (GuardBarHeight - cBarCodeHeight)
                Else
                    BarStartX = cBarCodeLeft
                End If
                BarH = GuardBarHeight - HRTextHeight
            End Select
    Next iCntr
    
    'DRAW THE HUMAN READABLE TEXT
    '============================
    If cHumanReadable Then
        TextStartY = cBarCodeTop
        TextStartX = cBarCodeLeft + HRTextHeight
        'Draw the first digit left of the left guard bars
        If HRLeftDigit <> "" Then
            rtn = DrawBarText(cHDC, TextStartX + 3, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRLeftDigit)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the left hand text right of the left guard bars
        TextStartY = TextStartY + LeftDigitWidth + ((3 + UPCAOffset) * NarrowBar)
        rtn = GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
        If rtn = 0 Then GoTo Raise_Error
        rtn = DrawBarText(cHDC, TextStartX, TextStartY + ((TextGap - HRTextWidth) / 2), cBarTextFont, _
                            cBarColor, cBarRotation, HRLeftText)
        If rtn = 0 Then GoTo Raise_Error
        'Draw the right hand text right of the centre guard bars
        If HRRightText <> "" Then
            TextStartY = TextStartY + TextGap + (5 * NarrowBar)
            rtn = GetTextSize(cHDC, cBarTextFont, HRRightText, HRTextWidth, HRTextHeight)
            If rtn = 0 Then GoTo Raise_Error
            rtn = DrawBarText(cHDC, TextStartX, TextStartY + ((TextGap - HRTextWidth) / 2), cBarTextFont, _
                            cBarColor, cBarRotation, HRRightText)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the check digit to the right of the right hand guard bars (UPC only)
        If HRRightDigit <> "" Then
            TextStartY = TextStartY + TextGap + ((4 + UPCEOffset + UPCAOffset) * NarrowBar)
            rtn = DrawBarText(cHDC, TextStartX + 3, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRRightDigit)
            If rtn = 0 Then GoTo Raise_Error
        End If
        'Draw the HR text above the supplemental bars (if using supplementals)
        If HRSupplementText <> "" Then
            rtn = GetTextSize(cHDC, cBarTextFont, HRSupplementText, HRTextWidth, HRTextHeight)
            If rtn = 0 Then GoTo Raise_Error
            TextStartY = (cBarCodeTop + BarCodeWidth - SupplementWidth) + ((SupplementWidth - HRTextWidth) / 2)
            TextStartX = cBarCodeLeft + cBarCodeHeight + HRTextHeight + cBarTextGap
            rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRSupplementText)
            If rtn = 0 Then GoTo Raise_Error
        End If
    End If
Exit Sub

Raise_Error:
    cErrorNumber = B_ErrDraw
    cErrDescription = S_ErrDraw
    err.Raise cErrorNumber, "clsEANBarCode.Draw_270", cErrDescription
End Sub

Private Function EncodeEAN13(DataToEncode As String)
    Dim iCntr As Integer
    Dim iWeight As Integer
    Dim sChkDigit As String * 1
    Dim TotalWeight As Integer
    Dim sDataChar As String * 1
    Dim sParity As String
    Dim sParityBit As String * 1
    
    'Calculate the check digit
    iWeight = 3
    For iCntr = Len(DataToEncode) To 1 Step -1
        TotalWeight = TotalWeight + (CInt(Mid(DataToEncode, iCntr, 1)) * iWeight)
        If iWeight = 3 Then iWeight = 1 Else iWeight = 3
    Next iCntr
    TotalWeight = TotalWeight Mod 10
    sChkDigit = IIf(TotalWeight = 0, "0", CStr(10 - TotalWeight))
    DataToEncode = DataToEncode & sChkDigit
    'Add the start guard bars
    EncodedData = "T111S"
    For iCntr = 1 To Len(DataToEncode)
        sDataChar = Mid(DataToEncode, iCntr, 1)
        If iCntr = 1 Then
            sParity = EANUPC_Parity(CInt(sDataChar)) 'Get parity based on first character
        Else
            'Left hand data
            If iCntr < 8 Then
                sParityBit = Mid(sParity, iCntr - 1, 1)
                Select Case sParityBit
                    Case "0"    'EVEN
                        EncodedData = EncodedData & EANUPC_LHE(CInt(sDataChar))
                    Case "1"    'ODD
                        EncodedData = EncodedData & EANUPC_LHO(CInt(sDataChar))
                    End Select
            End If
            'Insert the centre guard bars
            If iCntr = 8 Then EncodedData = EncodedData & "T11111S"
            'Right hand data
            If iCntr > 7 Then EncodedData = EncodedData & EANUPC_RH(CInt(sDataChar))
        End If
    Next iCntr
    'Add the end guard bars
    EncodedData = EncodedData & "T111S"
    'Set some drawing values
    UPCAOffset = 0
    UPCEOffset = 0
    NumberOfBars = 42
    HRLeftDigit = Left(DataToEncode, 1)
    HRRightDigit = ""
    HRLeftText = Mid(DataToEncode, 2, 6)
    HRRightText = Mid(DataToEncode, 8, 6)
    HRSupplementText = ""
    GetBarWidth
End Function

Private Function EncodeEAN8(DataToEncode As String)
    Dim iCntr As Integer
    Dim iWeight As Integer
    Dim sChkDigit As String * 1
    Dim TotalWeight As Integer
    Dim sDataChar As String * 1
    
    'Calculate the check digit
    iWeight = 3
    For iCntr = Len(DataToEncode) To 1 Step -1
        TotalWeight = TotalWeight + (CInt(Mid(DataToEncode, iCntr, 1)) * iWeight)
        If iWeight = 3 Then iWeight = 1 Else iWeight = 3
    Next iCntr
    TotalWeight = TotalWeight Mod 10
    sChkDigit = IIf(TotalWeight = 0, "0", CStr(10 - TotalWeight))
    DataToEncode = DataToEncode & sChkDigit
    'Add the start guard bars
    EncodedData = "T111S"
    For iCntr = 1 To Len(DataToEncode)
        sDataChar = Mid(DataToEncode, iCntr, 1)
        'Left hand data
        If iCntr < 5 Then EncodedData = EncodedData & EANUPC_LHO(CInt(sDataChar))
        'Insert the centre guard bars
        If iCntr = 5 Then EncodedData = EncodedData & "T11111S"
        'Right hand data
        If iCntr > 4 Then EncodedData = EncodedData & EANUPC_RH(CInt(sDataChar))
    Next iCntr
    'Add the end gaurd bars
    EncodedData = EncodedData & "T111S"
    'Set some drawing values
    UPCAOffset = 0
    UPCEOffset = 0
    NumberOfBars = 28
    HRLeftDigit = ""
    HRRightDigit = ""
    HRLeftText = Mid(DataToEncode, 1, 4)
    HRRightText = Mid(DataToEncode, 5, 4)
    HRSupplementText = ""
    GetBarWidth
End Function

Private Function EncodeUPCA(DataToEncode As String)
    Dim iCntr As Integer
    Dim iWeight As Integer
    Dim TotalWeight As Integer
    Dim sChkDigit As String * 1
    Dim sDataChar As String * 1
    
    'Calculate the check digit
    iWeight = 3
    For iCntr = Len(DataToEncode) To 1 Step -1
        TotalWeight = TotalWeight + (CInt(Mid(DataToEncode, iCntr, 1)) * iWeight)
        If iWeight = 3 Then iWeight = 1 Else iWeight = 3
    Next iCntr
    TotalWeight = TotalWeight Mod 10
    sChkDigit = IIf(TotalWeight = 0, "0", CStr(10 - TotalWeight))
    DataToEncode = DataToEncode & sChkDigit
    'Add the start guard bars
    EncodedData = "T111"
    For iCntr = 1 To Len(DataToEncode)
        sDataChar = Mid(DataToEncode, iCntr, 1)
            'Left hand data
            If iCntr < 7 Then EncodedData = EncodedData & EANUPC_LHO(CInt(sDataChar))
            If iCntr = 1 Then EncodedData = EncodedData & "S"
            'Insert the centre guard bars
            If iCntr = 7 Then EncodedData = EncodedData & "T11111S"
            'Right hand data
            If iCntr > 6 Then EncodedData = EncodedData & EANUPC_RH(CInt(sDataChar))
            If iCntr = 11 Then EncodedData = EncodedData & "T"
    Next iCntr
    'Add the end gaurd bars
    EncodedData = EncodedData & "111S"
    'Set some drawing values
    UPCAOffset = 7
    UPCEOffset = 0
    NumberOfBars = 35
    HRLeftDigit = Left(DataToEncode, 1)
    HRRightDigit = Right(DataToEncode, 1)
    HRLeftText = Mid(DataToEncode, 2, 5)
    HRRightText = Mid(DataToEncode, 7, 5)
    HRSupplementText = ""
    GetBarWidth
End Function

Private Function EncodeUPCE(DataToEncode As String)
    Dim iCntr As Integer
    Dim iChkDigit As Integer
    Dim sDataChar As String * 1
    Dim sParity As String
    Dim sParityBit As String * 1
    
    iChkDigit = UPCeToUPCa(DataToEncode)
    sParity = UPCe_Parity(iChkDigit)
    'Add the start guard bars
    EncodedData = "T111S"
    For iCntr = 1 To Len(DataToEncode)
        sDataChar = Mid(DataToEncode, iCntr, 1)
        sParityBit = Mid(sParity, iCntr, 1)
        Select Case sParityBit
            Case "0"    'EVEN
                EncodedData = EncodedData & EANUPC_LHE(CInt(sDataChar))
            Case "1"    'ODD
                EncodedData = EncodedData & EANUPC_LHO(CInt(sDataChar))
        End Select
    Next iCntr
    'Add the end guard bars
    EncodedData = EncodedData & "T111111S"
    'Set some drawing values
    UPCAOffset = 0
    UPCEOffset = 3
    NumberOfBars = 42
    HRLeftDigit = "0"
    HRRightDigit = CStr(iChkDigit)
    HRLeftText = DataToEncode
    HRRightText = ""
    HRSupplementText = ""
    GetBarWidth
End Function

Private Function EncodePlus2(DataToEncode As String)
    Dim iCntr As Integer
    Dim iBarCnt As Integer
    Dim iBW As Integer
    Dim iChkDigit As Integer
    Dim sDataChar As String * 1
    Dim sParity As String
    Dim sParityBit As String * 1
    Dim iGap As Integer
    
    iChkDigit = CInt(DataToEncode) Mod 4
    sParity = Plus2_Parity(iChkDigit)
    'Add the start guard bars
    SupplementData = "112"
    For iCntr = 1 To Len(DataToEncode)
        sDataChar = Mid(DataToEncode, iCntr, 1)
        sParityBit = Mid(sParity, iCntr, 1)
        Select Case sParityBit
            Case "0"    'EVEN
                SupplementData = SupplementData & EANUPC_LHE(CInt(sDataChar))
            Case "1"    'ODD
                SupplementData = SupplementData & EANUPC_LHO(CInt(sDataChar))
        End Select
        If iCntr < Len(DataToEncode) Then SupplementData = SupplementData & "11"
    Next iCntr
    HRSupplementText = DataToEncode
    NarrowBar = cBarX * cBarMultiplier
    For iCntr = 1 To Len(SupplementData)
        iBW = CInt(Mid(SupplementData, iCntr, 1))
        If iBW > 1 Then iBW = iBW
        iBarCnt = iBarCnt + iBW
    Next iCntr
    SupplementWidth = iBarCnt * NarrowBar
    EncodedData = EncodedData & "9" & SupplementData
    iGap = MilsToPixels(cHDC, 0.125)
    SupplementGap = IIf(iGap > RightDigitWidth, iGap, RightDigitWidth)
    BarCodeWidth = BarCodeWidth + SupplementGap + SupplementWidth - RightDigitWidth
End Function

Private Function EncodePlus5(DataToEncode As String)
    Dim iCntr As Integer
    Dim iBarCnt As Integer
    Dim iBW As Integer
    Dim iWeight As Integer
    Dim TotalWeight As Integer
    Dim sParity As String
    Dim iChkDigit As Integer
    Dim sDataChar As String * 1
    Dim sParityBit As String * 1
    Dim iGap As Integer
    
    iWeight = 3
    For iCntr = Len(DataToEncode) To 1 Step -1
        TotalWeight = TotalWeight + (CInt(Mid(DataToEncode, iCntr, 1)) * iWeight)
        If iWeight = 3 Then iWeight = 9 Else iWeight = 3
    Next iCntr
    iChkDigit = TotalWeight Mod 10
    sParity = Plus5_Parity(iChkDigit)
    'Add the start guard bars
    SupplementData = "112"
    For iCntr = 1 To Len(DataToEncode)
        sDataChar = Mid(DataToEncode, iCntr, 1)
        sParityBit = Mid(sParity, iCntr, 1)
        Select Case sParityBit
            Case "0"    'EVEN
                SupplementData = SupplementData & EANUPC_LHE(CInt(sDataChar))
            Case "1"    'ODD
                SupplementData = SupplementData & EANUPC_LHO(CInt(sDataChar))
        End Select
        If iCntr < Len(DataToEncode) Then SupplementData = SupplementData & "11"
    Next iCntr
    HRSupplementText = DataToEncode
    NarrowBar = cBarX * cBarMultiplier
    For iCntr = 1 To Len(SupplementData)
        iBW = CInt(Mid(SupplementData, iCntr, 1))
        If iBW > 1 Then iBW = iBW
        iBarCnt = iBarCnt + iBW
    Next iCntr
    SupplementWidth = iBarCnt * NarrowBar
    EncodedData = EncodedData & "9" & SupplementData
    iGap = MilsToPixels(cHDC, 0.125)
    SupplementGap = IIf(iGap > RightDigitWidth, iGap, RightDigitWidth)
    BarCodeWidth = BarCodeWidth + SupplementGap + SupplementWidth - RightDigitWidth
End Function

Private Sub GetBarWidth()
    'Calculate the width of the barcode by first counting the total number of bars.
    'If "autofontsize" is true, detects the largest font size that will fit between
    'the guard bars, else use the current font size.
    'Then calculate the width of the character either side of the barcode (if any)
    'based on the font size. Sum these values to get the total barcode width.
    
    Dim iCntr As Integer
    Dim iBarCnt As Integer
    Dim c As String
    Dim w1 As Integer
    Dim w2 As Integer
    Dim fs As Integer       'Font size
    Dim rtn As Long
    
    'Count the bars to calculate the bar width
    NarrowBar = cBarX * cBarMultiplier
    For iCntr = 1 To Len(EncodedData)
        c = Mid(EncodedData, iCntr, 1)
        'Mask off the bar height switch markers "T" & "S"
        If c <> "T" And c <> "S" Then iBarCnt = iBarCnt + CInt(c)
    Next iCntr
    
    'Calculate the gap between the guard bars
    TextGap = NarrowBar * NumberOfBars
    If cAutoFontSize Then
        fs = 0
        Do
            fs = fs + 1     'Start with a fontsize of 1
            cBarTextFont.FontSize = fs
            'Check both the left and right side text in case one is wider than the other
            Call GetTextSize(cHDC, cBarTextFont, HRLeftText, HRTextWidth, HRTextHeight)
            w1 = HRTextWidth
            If HRRightText <> "" Then Call GetTextSize(cHDC, cBarTextFont, HRRightText, HRTextWidth, HRTextHeight)
            w2 = HRTextWidth
            HRTextWidth = IIf(w1 > w2, w1, w2)
        Loop Until HRTextWidth >= TextGap
        cBarTextFont.FontSize = fs - 1
    End If
    'Get the width of the left hand digit
    LeftDigitWidth = 0: RightDigitWidth = 0
    If HRLeftDigit <> "" Then Call GetTextSize(cHDC, cBarTextFont, HRLeftDigit, LeftDigitWidth, HRTextHeight)
    'Get the width of the right hand digit
    If HRRightDigit <> "" Then Call GetTextSize(cHDC, cBarTextFont, HRRightDigit, RightDigitWidth, HRTextHeight)
    If Not cHumanReadable Then LeftDigitWidth = 0: RightDigitWidth = 0
    BarCodeWidth = LeftDigitWidth + RightDigitWidth + (NarrowBar * iBarCnt)
End Sub

Private Function UPCeToUPCa(ByVal DataToConvert As String) As Integer
    'Expand the UCPe code to UPCa for the purpose of
    'calculating the check digit.
    'The check digit is used to determine the parity
    'for encoding the UPCe data.
    
    Dim iCntr As Integer
    Dim iWeight As Integer
    Dim TotalWeight As Integer
    Dim UPCaCode As String
    Dim D1 As String
    Dim D2 As String
    Dim D3 As String
    Dim D4 As String
    Dim D5 As String
    Dim D6 As String
    Dim D7 As String
    
    DataToConvert = "0" & DataToConvert
    D1 = Mid(DataToConvert, 1, 1)
    D2 = Mid(DataToConvert, 2, 1)
    D3 = Mid(DataToConvert, 3, 1)
    D4 = Mid(DataToConvert, 4, 1)
    D5 = Mid(DataToConvert, 5, 1)
    D6 = Mid(DataToConvert, 6, 1)
    D7 = Mid(DataToConvert, 7, 1)
    Select Case D7
        Case "0"
            UPCaCode = D1 & D2 & D3 & "00000" & D4 & D5 & D6
        Case "1", "2"
            UPCaCode = D1 & D2 & D3 & D7 & "0000" & D4 & D5 & D6
        Case "3"
            UPCaCode = D1 & D2 & D3 & D4 & "00000" & D5 & D6
        Case "4"
            UPCaCode = D1 & D2 & D3 & D4 & D5 & "00000" & D6
        Case "5", "6", "7", "8", "9"
            UPCaCode = D1 & D2 & D3 & D4 & D5 & D6 & "0000" & D7
    End Select
    iWeight = 3
    For iCntr = Len(UPCaCode) To 1 Step -1
        TotalWeight = TotalWeight + (CInt(Mid(UPCaCode, iCntr, 1)) * iWeight)
        If iWeight = 3 Then iWeight = 1 Else iWeight = 3
    Next iCntr
    TotalWeight = TotalWeight Mod 10
    UPCeToUPCa = IIf(TotalWeight = 0, 0, 10 - TotalWeight)
End Function

Private Sub Class_Initialize()
    'Initialize some default values
    'Barcode
    cSymbology = 0
    cBarX = 1
    cBarMultiplier = 1
    cBarRotation = 0
    cBarCodeTop = 1
    cBarCodeLeft = 1
    cBarCodeHeight = 45
    cBarTextGap = 0
    cBarCodeData = ""
    
    'Barcode text
    cHumanReadable = True
    cAutoFontSize = False
    cBarTextFont.FontName = "Tahoma"
    cBarTextFont.FontSize = 8
    cBarTextFont.FontBold = False
    cBarTextFont.FontItalic = False
    
    'Errors
    cErrorNumber = 0
    cErrDescription = ""
    
    'Barcode characters
    Initialise_BC
End Sub

Private Sub Initialise_BC()

    BC_ValidCharacters = "0123456789"
        
    'Left hand ODD encoding
    EANUPC_LHO(0) = "3211": EANUPC_LHO(1) = "2221"
    EANUPC_LHO(2) = "2122": EANUPC_LHO(3) = "1411"
    EANUPC_LHO(4) = "1132": EANUPC_LHO(5) = "1231"
    EANUPC_LHO(6) = "1114": EANUPC_LHO(7) = "1312"
    EANUPC_LHO(8) = "1213": EANUPC_LHO(9) = "3112"
    
    'Left hand EVEN
    EANUPC_LHE(0) = "1123": EANUPC_LHE(1) = "1222"
    EANUPC_LHE(2) = "2212": EANUPC_LHE(3) = "1141"
    EANUPC_LHE(4) = "2311": EANUPC_LHE(5) = "1321"
    EANUPC_LHE(6) = "4111": EANUPC_LHE(7) = "2131"
    EANUPC_LHE(8) = "3121": EANUPC_LHE(9) = "2113"
    
    'Right hand encoding
    EANUPC_RH(0) = "3211": EANUPC_RH(1) = "2221"
    EANUPC_RH(2) = "2122": EANUPC_RH(3) = "1411"
    EANUPC_RH(4) = "1132": EANUPC_RH(5) = "1231"
    EANUPC_RH(6) = "1114": EANUPC_RH(7) = "1312"
    EANUPC_RH(8) = "1213": EANUPC_RH(9) = "3112"
    
    'EAN UPCa Parity values "1" = Odd, "0" = Even
    EANUPC_Parity(0) = "111111": EANUPC_Parity(1) = "110100"
    EANUPC_Parity(2) = "110010": EANUPC_Parity(3) = "110001"
    EANUPC_Parity(4) = "101100": EANUPC_Parity(5) = "100110"
    EANUPC_Parity(6) = "100011": EANUPC_Parity(7) = "101010"
    EANUPC_Parity(8) = "101001": EANUPC_Parity(9) = "100101"
    
    'UPC-E parity values "1" = Odd, "0" = Even
    UPCe_Parity(0) = "000111": UPCe_Parity(1) = "001011"
    UPCe_Parity(2) = "001101": UPCe_Parity(3) = "001110"
    UPCe_Parity(4) = "010011": UPCe_Parity(5) = "011001"
    UPCe_Parity(6) = "011100": UPCe_Parity(7) = "010101"
    UPCe_Parity(8) = "010110": UPCe_Parity(9) = "011010"
    
    '2 digit extention parity values "1" = ODD, "0" = EVEN
    Plus2_Parity(0) = "11": Plus2_Parity(1) = "10"
    Plus2_Parity(2) = "01": Plus2_Parity(3) = "00"
    
    '5 digit extention parity values "1" = ODD, "0" = EVEN
    Plus5_Parity(0) = "00111": Plus5_Parity(1) = "01011"
    Plus5_Parity(2) = "01101": Plus5_Parity(3) = "01110"
    Plus5_Parity(4) = "10011": Plus5_Parity(5) = "11001"
    Plus5_Parity(6) = "11100": Plus5_Parity(7) = "10101"
    Plus5_Parity(8) = "10110": Plus5_Parity(9) = "11010"
End Sub



