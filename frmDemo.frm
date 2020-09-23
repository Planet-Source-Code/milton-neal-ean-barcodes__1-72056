VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   8550
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9075
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtXfactor 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6000
      TabIndex        =   29
      Text            =   ".005"
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Barcode Color ...."
      Height          =   315
      Left            =   6960
      TabIndex        =   26
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Human Readable"
      Height          =   1455
      Left            =   240
      TabIndex        =   23
      Top             =   6480
      Width           =   2175
      Begin VB.CheckBox chkAutoFont 
         Caption         =   "Auto Font Size"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optHR 
         Caption         =   "Barcode and Text"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optHR 
         Caption         =   "Barcode Only"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rotation"
      Height          =   1815
      Left            =   2640
      TabIndex        =   18
      Top             =   6480
      Width           =   2055
      Begin VB.OptionButton optRotation 
         Caption         =   "Upside Down"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optRotation 
         Caption         =   "Sideways Down"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optRotation 
         Caption         =   "Sideways UP"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optRotation 
         Caption         =   "Normal"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font ...."
      Height          =   315
      Left            =   7320
      TabIndex        =   17
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox txtGap 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6000
      TabIndex        =   16
      Text            =   "0"
      Top             =   5280
      Width           =   585
   End
   Begin VB.ComboBox cboMultiplier 
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   7080
      Width           =   900
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "&Draw"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   6600
      Width           =   900
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Text            =   "12"
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Text            =   "5"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Text            =   "5"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtBarcode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   5760
      Width           =   5775
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   28
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picWorkspace 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   120
      X2              =   8880
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar/Text Gap [mm]"
      Height          =   195
      Left            =   4320
      TabIndex        =   15
      Top             =   5280
      Width           =   1350
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BarX Factor [mils]:"
      Height          =   195
      Left            =   4320
      TabIndex        =   14
      Top             =   4800
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplier"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Height: [mm]"
      Height          =   195
      Left            =   2160
      TabIndex        =   12
      Top             =   4800
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y: [mm]"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   5280
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X: [mm]"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode Data:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveBMP 
         Caption         =   "Save as Bitmap"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuBarCode 
      Caption         =   "Barcode"
      Begin VB.Menu mnuCodeEAN13 
         Caption         =   "EAN 13"
      End
      Begin VB.Menu mnuCodeEAN_2 
         Caption         =   "EAN 13+2"
      End
      Begin VB.Menu mnuCodeEAN13_5 
         Caption         =   "EAN 13+5"
      End
      Begin VB.Menu mnuCodeSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeEAN8 
         Caption         =   "EAN 8"
      End
      Begin VB.Menu mnuCodeEAN8_2 
         Caption         =   "EAN 8+2"
      End
      Begin VB.Menu mnuCodeEAN8_5 
         Caption         =   "EAN 8+5"
      End
      Begin VB.Menu mnuCodeSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeUPCA 
         Caption         =   "UPC A"
      End
      Begin VB.Menu mnuCodeUPCA_2 
         Caption         =   "UPC A+2"
      End
      Begin VB.Menu mnuCodeUPCA_5 
         Caption         =   "UPC A+5"
      End
      Begin VB.Menu mnuCodeSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeUPCE 
         Caption         =   "UPC E"
      End
      Begin VB.Menu mnuCodeUPCE_2 
         Caption         =   "UPC E+2"
      End
      Begin VB.Menu mnuCodeUPCE_5 
         Caption         =   "UPC E+5"
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Dim BC As clsEANBarCodes
Dim Ratio As Single
Dim fname As String
Dim fsize As Single
Dim fbold As Boolean
Dim fitalic As Boolean

Private Sub cmdFont_Click()
    CommonDialog1.FontName = Me.txtBarcode.FontName
    ' Set Cancel to True
    CommonDialog1.CancelError = True
    On Error GoTo errHandler
    ' Set the Flags property
    CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
    ' Display the Font dialog box
    CommonDialog1.ShowFont
    fname = CommonDialog1.FontName
    fsize = CommonDialog1.FontSize
    fbold = CommonDialog1.FontBold
    fitalic = CommonDialog1.FontItalic
    Me.txtBarcode.Font.Name = CommonDialog1.FontName
    Me.txtBarcode.Font.Bold = CommonDialog1.FontBold
    Me.txtBarcode.Font.Italic = CommonDialog1.FontItalic
    cmdDraw_Click
Exit Sub
errHandler:
  ' User pressed the Cancel button
  Exit Sub

End Sub

Private Sub cmdColor_Click()
On Error GoTo errHandler
With CommonDialog1
    'Set Cancel to True
    .CancelError = True
    'Set the Flags property
    .Flags = cdlCCRGBInit
    'Display the Color Dialog box
    .ShowColor
    BC.BarColor = .Color
End With
cmdDraw_Click
errHandler:
' User pressed the Cancel button
End Sub

Private Sub cmdDraw_Click()
    Dim MilsToPixels As Single
    Dim th As Long
    Dim tw As Long
    Dim pr As Single
    Dim br As Single

    MilsToPixels = 1440 / Screen.TwipsPerPixelX
    'MilsToPixels = 1440 / Printer.TwipsPerPixelX
    pr = CSng(Me.txtXfactor.Text) * MilsToPixels
    On Error GoTo Err_Handler
    picWorkspace.Cls
    With BC
        .BarcodeOutput = Me.picWorkspace
        .BarCodeX = CLng(Me.txtX.Text) * (56.7 / Screen.TwipsPerPixelX)
        .BarCodeY = CLng(Me.txtY.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarXWidth = pr
        .BarMultiplier = CInt(Me.cboMultiplier.Text)
        .BarcodeHeight = CLng(Me.txtHeight.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarTextGap = CInt(Me.txtGap.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarTextFont fname, fsize, fbold, fitalic
        .BarCodeData = txtBarcode.Text
        .Draw_BarCode
        th = .TotalBarHeight
        tw = .TotalBarWidth
    Call Rectangle(Me.picWorkspace.hdc, .BarCodeX - 2, .BarCodeY - 2, .BarCodeX + tw + 2, .BarCodeY + th + 2)
    End With
    Me.picWorkspace.Refresh
    Exit Sub
    
Err_Handler:
MsgBox "Error Number: " & CStr(err.Number) & vbCrLf _
                        & "Description: " & err.Description & vbCrLf _
                        & "Error Source: " & err.Source, vbOKOnly + vbCritical, "Error"
End Sub

Private Sub cmdPrint_Click()
    Dim MilsToPixels As Single
    Dim pr As Single
    Dim br As Single
    
    MilsToPixels = 1440 / Printer.TwipsPerPixelX
    pr = CSng(Me.txtXfactor.Text) * MilsToPixels
    On Error GoTo Err_Handler
    With BC
        .BarcodeOutput = Printer
        .BarCodeX = CLng(Me.txtX.Text) * (56.7 / Printer.TwipsPerPixelX)
        .BarCodeY = CLng(Me.txtY.Text) * (56.7 / Printer.TwipsPerPixelY)
        .BarXWidth = pr
        .BarMultiplier = CInt(Me.cboMultiplier.Text)
        .BarcodeHeight = CLng(Me.txtHeight.Text) * 56.7 / Printer.TwipsPerPixelY
        .BarTextGap = CInt(Me.txtGap.Text) * (56.7 / Printer.TwipsPerPixelY)
        .BarTextFont fname, fsize, fbold, fitalic
        .BarCodeData = txtBarcode.Text
        Printer.Print ""
        .Draw_BarCode
    End With
    Printer.EndDoc
    Exit Sub
    
Err_Handler:
MsgBox "Error Number: " & CStr(err.Number) & vbCrLf _
                        & "Description: " & err.Description & vbCrLf _
                        & "Error Source: " & err.Source, vbOKOnly + vbCritical, "Error"
End Sub











Private Sub mnuFileSaveBMP_Click()
    Dim MilsToPixels As Single
    Dim th As Long
    Dim tw As Long
    Dim pr As Single
    Dim br As Single

    MilsToPixels = 1440 / Screen.TwipsPerPixelX
    On Error GoTo Err_Handler
    Me.picSave.Cls
    With BC
        .BarcodeOutput = Me.picSave
        .BarCodeX = 25
        .BarCodeY = 25
        .BarXWidth = 1
        .BarMultiplier = CInt(Me.cboMultiplier.Text)
        .BarcodeHeight = CLng(Me.txtHeight.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarTextGap = CInt(Me.txtGap.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarTextFont fname, fsize, fbold, fitalic
        .BarCodeData = txtBarcode.Text
        th = .TotalBarHeight
        tw = .TotalBarWidth
        Me.picSave.Width = (tw + 50) * Screen.TwipsPerPixelX
        Me.picSave.Height = (th + 50) * Screen.TwipsPerPixelY
        .BarcodeOutput = Me.picSave
        .Draw_BarCode
    End With
    Me.picSave.Refresh
    SavePicture Me.picSave.Image, "C:\barcode.bmp"
    Exit Sub
    
    
Err_Handler:
MsgBox "Error Number: " & CStr(err.Number) & vbCrLf _
                        & "Description: " & err.Description & vbCrLf _
                        & "Error Source: " & err.Source, vbOKOnly + vbCritical, "Error"
End Sub
Private Sub Form_Load()
    Dim X As Integer
    
    Set BC = New clsEANBarCodes
    Me.picWorkspace.ScaleMode = vbPixels
    Me.picSave.ScaleMode = vbPixels
    Me.lblHeading.Caption = "Select a Barcode"
    For X = 1 To 10
        Me.cboMultiplier.AddItem CStr(X), X - 1
    Next X
    Me.cboMultiplier.ListIndex = 0
    
    'Me.cboBarRatio.AddItem "3:1", 0
    'Me.cboBarRatio.AddItem "2:1", 1
    'Me.cboBarRatio.ListIndex = 0
    Me.mnuFilePrint.Enabled = False
    Me.mnuFileSaveBMP.Enabled = False
    Me.Show
    Me.txtBarcode.SetFocus
    CheckFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set BC = Nothing
End Sub

Private Sub mnuCodeEAN13_Click()
    Me.lblHeading.Caption = "EAN-13"
    BC.Symbology = EAN13
    CheckFields
End Sub

Private Sub mnuCodeEAN_2_Click()
    Me.lblHeading.Caption = "EAN-13+2"
    BC.Symbology = EAN13_2
    CheckFields
End Sub

Private Sub mnuCodeEAN13_5_Click()
    Me.lblHeading.Caption = "EAN-13+5"
    BC.Symbology = EAN13_5
    CheckFields
End Sub
Private Sub mnuCodeEAN8_Click()
    Me.lblHeading.Caption = "EAN-8"
    BC.Symbology = EAN8
    CheckFields
End Sub

Private Sub mnuCodeEAN8_2_Click()
    Me.lblHeading.Caption = "EAN-8+2"
    BC.Symbology = EAN8_2
    CheckFields
End Sub
Private Sub mnuCodeEAN8_5_Click()
    Me.lblHeading.Caption = "EAN-8+5"
    BC.Symbology = EAN8_5
    CheckFields
End Sub
Private Sub mnuCodeUPCA_Click()
    Me.lblHeading.Caption = "UPC-A"
    BC.Symbology = UPCA
    CheckFields
End Sub
Private Sub mnuCodeUPCA_2_Click()
    Me.lblHeading.Caption = "UPC-A+2"
    BC.Symbology = UPCA_2
    CheckFields
End Sub
Private Sub mnuCodeUPCA_5_Click()
    Me.lblHeading.Caption = "UPC-A+5"
    BC.Symbology = UPCA_5
    CheckFields
End Sub
Private Sub mnuCodeUPCE_Click()
    Me.lblHeading.Caption = "UPC-E"
    BC.Symbology = UPCE
    CheckFields
End Sub
Private Sub mnuCodeUPCE_2_Click()
    Me.lblHeading.Caption = "UPC-E+2"
    BC.Symbology = UPCE_2
    CheckFields
End Sub
Private Sub mnuCodeUPCE_5_Click()
    Me.lblHeading.Caption = "UPC-E+5"
    BC.Symbology = UPCE_5
    CheckFields
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End
End Sub



Private Sub optHR_Click(Index As Integer)
    BC.HRText = Index
    cmdDraw_Click
End Sub

Private Sub optRotation_Click(Index As Integer)
Select Case Index
    Case 0
        BC.BarRotation = 0
    Case 1
        BC.BarRotation = 90
    Case 2
        BC.BarRotation = 270
    Case 3
        BC.BarRotation = 180
End Select
cmdDraw_Click
End Sub

Private Sub chkAutoFont_Click()
    BC.AutoTextFont = Me.chkAutoFont.Value = Checked
    cmdDraw_Click
End Sub

Private Sub txtBarcode_Change()
    CheckFields
End Sub

Private Sub txtHeight_Change()
    CheckFields
End Sub

Private Sub txtX_Change()
  CheckFields
End Sub

Private Sub txtY_Change()
    CheckFields
End Sub


Private Sub CheckFields()
Dim sym As Integer

sym = BC.Symbology
If Me.txtX.Text = "" Or Me.txtY = "" Or Me.txtHeight = "" Or Me.txtBarcode = "" Or sym = 0 Then
        Me.cmdDraw.Enabled = False
        Me.cmdPrint.Enabled = False
        Me.mnuFilePrint.Enabled = False
        Me.mnuFileSaveBMP.Enabled = False
    Else
        Me.cmdDraw.Enabled = True
        Me.cmdPrint.Enabled = True
        Me.mnuFilePrint.Enabled = True
        Me.mnuFileSaveBMP.Enabled = True
    End If
End Sub

