VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FontEd 09"
   ClientHeight    =   3360
   ClientLeft      =   150
   ClientTop       =   510
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fonted.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Font"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.HScrollBar hsbFont 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      Max             =   253
      TabIndex        =   7
      Top             =   600
      Value           =   187
      Width           =   1455
   End
   Begin VB.TextBox txtFont 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.OptionButton optColor 
      Height          =   255
      Index           =   0
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   255
      Index           =   14
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   255
      Index           =   15
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox picCut 
      BackColor       =   &H00D69896&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2880
      Left            =   1200
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.HScrollBar hsbWidth 
      Height          =   255
      Left            =   120
      Max             =   8
      TabIndex        =   1
      Top             =   3000
      Value           =   6
      Visible         =   0   'False
      Width           =   1440
   End
   Begin FontEd.GBATileEditor tedEdit 
      Height          =   2880
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   5080
      DotSize         =   12
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   195
      Left            =   2280
      TabIndex        =   14
      Top             =   2280
      Width           =   225
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      Height          =   195
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label lblROM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   195
      Left            =   2280
      TabIndex        =   12
      Top             =   2040
      Width           =   225
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROM:"
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      Top             =   2040
      Width           =   405
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   615
      Left            =   1680
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00A6A6B4&
      X1              =   176
      X2              =   240
      Y1              =   216
      Y2              =   216
   End
   Begin VB.Label lblZodiacDaGreat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZodiacDaGreat"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1680
      TabIndex        =   9
      Top             =   3120
      Width           =   915
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   3165
      Left            =   105
      Top             =   105
      Width           =   1470
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font index"
      Height          =   195
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1695
      Left            =   1680
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open ROM"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save ROM"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "&Font"
      Begin VB.Menu mnuGrid 
         Caption         =   "Show Gridlines"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIEntire 
         Caption         =   "Import (A-Z)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEEntire 
         Caption         =   "Export (A-Z)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sHeader As String * 4
Private sFilePath As String
Private iFileNum As Integer
Private FontPalette As Long
Private FontGFX As Long
Private FontWidth As Long

Private Sub cmdSave_Click()
    mnuSave_Click
End Sub

Private Sub Form_Load()
    SetIcon Me.hWnd, "AAA"
    optColor(0).BackColor = RGB(254, 254, 254)
    optColor(14).BackColor = RGB(192, 192, 192)
    optColor(15).BackColor = RGB(64, 64, 64)
End Sub

Private Sub hsbFont_Change()
Dim b As Byte
    txtFont.Text = hsbFont.Value
    tedEdit.ROMAddress = FontGFX + (txtFont.Text * 64) ' Font Size = 0x40 bytes per font
    tedEdit.LoadTileData
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, FontWidth + 1 + txtFont.Text, b
    Close #iFileNum
    hsbWidth.Value = b
End Sub

Private Sub hsbFont_Scroll()
    hsbFont_Change
End Sub

Private Sub hsbWidth_Change()
    picCut.Left = 8 + (hsbWidth * 12)
    picCut.Width = 104 - picCut.Left
End Sub

Private Sub mnuAbout_Click()
    MsgBox "FontEd 09" & vbNewLine & "-----------" & vbNewLine & "Original Coding by Kawa & crew" & vbNewLine & "FontEd 09 by ZodiacDaGreat" & vbNewLine & vbNewLine, , "About"
End Sub

Private Sub mnuEEntire_Click()
Dim arrTemp() As Byte
Dim sExport As String
Dim cdgExport As clsCommonDialog
    Set cdgExport = New clsCommonDialog
    sExport = cdgExport.ShowSave(Me.hWnd, "Save Font...", "Font A-Z", , "Binary Files (*.bin)|*.bin", OVERWRITEPROMPT)
    If LenB(sExport) = 0 Then GoTo Endme
    
    ReadByteArray sFilePath, arrTemp, FontGFX + (187 * 64), &HD00
    WriteByteArray sExport, arrTemp, 0
    
    Erase arrTemp
Endme:
    Set cdgExport = Nothing
End Sub

Private Sub mnuExport_Click()
Dim arrTemp() As Byte
Dim sExport As String
Dim cdgExport As clsCommonDialog
    Set cdgExport = New clsCommonDialog
    sExport = cdgExport.ShowSave(Me.hWnd, "Save Font...", "Font " & txtFont.Text, , "Binary Files (*.bin)|*.bin", OVERWRITEPROMPT)
    If LenB(sExport) = 0 Then GoTo Endme
    
    ReadByteArray sFilePath, arrTemp, FontGFX + (txtFont.Text * 64), &H40
    WriteByteArray sExport, arrTemp, 0
    
    Erase arrTemp
Endme:
    Set cdgExport = Nothing
End Sub

Private Sub mnuGrid_Click()
    If mnuGrid.Checked = True Then
        mnuGrid.Checked = False
        tedEdit.ShowGrid = False
    Else
        tedEdit.ShowGrid = True
        mnuGrid.Checked = True
    End If
    LoadFont
End Sub

Private Sub mnuIEntire_Click()
Dim arrTemp() As Byte
Dim sImport As String
Dim cdgImport As clsCommonDialog
    Set cdgImport = New clsCommonDialog
    sImport = cdgImport.ShowOpen(Me.hWnd, "Import Font...", , "Binary Files (*.bin)|*.bin")
    If LenB(sImport) = 0 Then GoTo Endme
    
    ReadByteArray sImport, arrTemp, 0
    WriteByteArray sFilePath, arrTemp, FontGFX + (187 * 64)
    hsbFont_Change
    
    Erase arrTemp
Endme:
    Set cdgImport = Nothing
End Sub

Private Sub mnuImport_Click()
Dim arrTemp() As Byte
Dim sImport As String
Dim cdgImport As clsCommonDialog
    Set cdgImport = New clsCommonDialog
    sImport = cdgImport.ShowOpen(Me.hWnd, "Import Font...", , "Binary Files (*.bin)|*.bin")
    If LenB(sImport) = 0 Then GoTo Endme
    
    ReadByteArray sImport, arrTemp, 0
    WriteByteArray sFilePath, arrTemp, FontGFX + (txtFont.Text * 64)
    hsbFont_Change
    
    Erase arrTemp
Endme:
    Set cdgImport = Nothing
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String
Dim cdgOpen As clsCommonDialog
    iFileNum = FreeFile
    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, "Open ROM...", , "GameBoy Advance ROMs (*.gba,*.agb,*.bin)|*.gba;*.agb;*.bin")
    If LenB(sResult) = 0 Then GoTo Endme
    sFilePath = sResult
        
    Open sFilePath For Binary As #iFileNum
    Get #iFileNum, &HAC + 1, sHeader
    Close #iFileNum
        
    Select Case sHeader
        Case "AXVE"
            lblROM.Caption = "Ruby Version"
            FontGFX = &HEA2C44
            FontWidth = &H1E6594
            FontPalette = &H1E66B2
        
        Case "AXPE"
            lblROM.Caption = "Sapphire Version"
            FontGFX = &HEA2DF0
            FontWidth = &H1E6524
            FontPalette = &H1E66B2
            
        Case Else
            MsgBox "Error! Unsupported ROM!", vbCritical, "Error"
            GoTo Endme
            
    End Select
    
    mnuExport.Enabled = True
    mnuImport.Enabled = True
    mnuEEntire.Enabled = True
    mnuIEntire.Enabled = True
    mnuGrid.Enabled = True
    mnuSave.Enabled = True
    cmdSave.Enabled = True
    txtFont.Enabled = True
    hsbFont.Enabled = True
    txtFont.Text = 187
    lblCode.Caption = sHeader
    lblROM.ToolTipText = sFilePath
    hsbWidth.Visible = True
    tedEdit.Visible = True
    picCut.Visible = True

    LoadFont
Endme:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()
    tedEdit.SaveTileData
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Put #iFileNum, FontWidth + 1 + txtFont.Text, CByte(hsbWidth.Value)
    Close #iFileNum
End Sub

Private Sub optColor_Click(Index As Integer)
    tedEdit.PenColor = Index
End Sub

Function LoadFont()
    tedEdit.FileName = sFilePath
    
    tedEdit.Colors(0) = RGB(254, 254, 254)
    tedEdit.Colors(14) = RGB(192, 192, 192)
    tedEdit.Colors(15) = RGB(64, 64, 64)
    
    optColor(15).Value = True
    hsbFont_Change
End Function

Private Sub txtFont_keypress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
    If txtFont.Text = "" Then Exit Sub
    If Val(txtFont.Text) > 253 Then Exit Sub
    
    hsbFont.Value = txtFont.Text
End Sub
