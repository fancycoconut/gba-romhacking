VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tileset Manager"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frBytes 
      Caption         =   "Amount of Bytes Needed"
      Height          =   1575
      Left            =   3840
      TabIndex        =   10
      Tag             =   "11"
      Top             =   720
      Width           =   2415
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   2175
         TabIndex        =   11
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   12
            Tag             =   "14"
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblOffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "???"
            Height          =   195
            Left            =   960
            TabIndex        =   18
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lbl9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "At Offset:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Tag             =   "17"
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblBytesDecimal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "???"
            Height          =   195
            Left            =   600
            TabIndex        =   16
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lbl7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dec:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   330
         End
         Begin VB.Label lbl8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hex:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   345
         End
         Begin VB.Label lblBytesHex 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "???"
            Height          =   195
            Left            =   600
            TabIndex        =   13
            Top             =   480
            Width           =   225
         End
      End
   End
   Begin VB.Frame frTileset 
      Caption         =   "Insert Tileset"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Tag             =   "10"
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   2
         Top             =   240
         Width           =   3375
         Begin VB.TextBox txt1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Text            =   "0x"
            Top             =   120
            Width           =   270
         End
         Begin VB.CheckBox chSubColor 
            Caption         =   "Sub Color 0"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Tag             =   "37"
            Top             =   1410
            Width           =   1575
         End
         Begin VB.TextBox txtOffset 
            Height          =   285
            Left            =   360
            MaxLength       =   6
            TabIndex        =   7
            Text            =   "800000"
            Top             =   120
            Width           =   780
         End
         Begin VB.CommandButton cmdInsert 
            Caption         =   "Insert"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            TabIndex        =   6
            Tag             =   "13"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtBlocks 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "100"
            Top             =   480
            Width           =   1020
         End
         Begin VB.CheckBox chCompression 
            Caption         =   "Compressed Tileset"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Tag             =   "19"
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox chMajor 
            Caption         =   "Major Tileset"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Tag             =   "20"
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label lbl4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset"
            Height          =   195
            Left            =   1320
            TabIndex        =   9
            Tag             =   "15"
            Top             =   120
            Width           =   465
         End
         Begin VB.Label lbl5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount of Blocks"
            Height          =   195
            Left            =   1320
            TabIndex        =   8
            Tag             =   "18"
            Top             =   510
            Width           =   1230
         End
      End
   End
   Begin VB.Frame frInfo 
      Caption         =   "Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Tag             =   "12"
      Top             =   2400
      Width           =   6135
      Begin VB.PictureBox picFlag 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5520
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   27
         Top             =   720
         Width           =   495
         Begin VB.Shape shpFlag 
            BorderColor     =   &H00C0C0C0&
            Height          =   225
            Left            =   150
            Top             =   120
            Width           =   300
         End
         Begin VB.Image imgFlag 
            Height          =   165
            Left            =   180
            Top             =   150
            Width           =   240
         End
      End
      Begin VB.TextBox txtFilePath 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "???"
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox txtTilesetPath 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "???"
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   960
         TabIndex        =   25
         Top             =   840
         Width           =   225
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Tag             =   "16"
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROM:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tileset:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   525
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   1
      Begin VB.Menu mnuOpen 
         Caption         =   "Open ROM"
         HelpContextID   =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpenTileset 
         Caption         =   "Open Tileset"
         HelpContextID   =   3
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         HelpContextID   =   4
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTileset 
      Caption         =   "&Tileset"
      HelpContextID   =   5
      Begin VB.Menu mnuCompress 
         Caption         =   "Compress Tileset"
         Enabled         =   0   'False
         HelpContextID   =   6
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuUncompress 
         Caption         =   "Uncompress Tileset"
         Enabled         =   0   'False
         HelpContextID   =   7
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview Tileset"
         Enabled         =   0   'False
         HelpContextID   =   38
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPalette 
         Caption         =   "Palette Inserter"
         Enabled         =   0   'False
         HelpContextID   =   39
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   8
      Begin VB.Menu mnuReadMe 
         Caption         =   "Readme"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         HelpContextID   =   9
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

Private Sub chMajor_Click()
    If chMajor.Value = vbChecked Then
        Select Case Left(sHeader, 3)
            Case "BPR", "BPG"
                txtBlocks.Text = 640
            Case Else
                txtBlocks.Text = 512
        End Select
        
        txtBlocks.Enabled = False
        chSubColor.Enabled = False
        chSubColor.Value = vbChecked
    Else
        txtBlocks.Text = 100
        txtBlocks.Enabled = True
        chSubColor.Enabled = True
        chSubColor.Value = vbUnchecked
    End If
End Sub

Private Sub cmdFind_Click()
Dim var2 As Long
Dim var3 As Long
    Select Case Left(sHeader, 3)
        Case "BPR", "BPG"
            var3 = 4 * CLng(txtBlocks.Text)
        Case Else
            var3 = 2 * CLng(txtBlocks.Text)
    End Select
    FindOffset
    var2 = 16 * CLng(frmMain.txtBlocks.Text)
    CountFileBytes sTilesetPath, BytesAmount
    BytesAmount = BytesAmount + 24 + 196 + var2 + var3 + 2 ' 24 = TilesetHeader, 196 = Palettes
    
    lblOffset.Caption = "0x" & Hex(Ofst)
    lblBytesDecimal.Caption = BytesAmount & " bytes"
    lblBytesHex.Caption = "0x" & Hex(BytesAmount) & " bytes"
End Sub

Private Sub cmdInsert_Click()
    Calculation
    InsertTileset
    InsertPalette
    InsertBlockData
    InsertBehaviourData
    InsertTilesetHeader
    If EnlargedROM = 1 Then
        FixPointers
    End If
    GenerateLogFile
End Sub

Private Sub FixPointers()
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Select Case Left(sHeader, 3)
            Case "AXV", "AXP", "BPE"
                Put #iFileNum, RepointOffset + 1 + 19, &H9
            Case "BPR", "BPG"
                Put #iFileNum, RepointOffset + 1 + 23, &H9
        End Select
        
        Put #iFileNum, RepointOffset + 1 + 7, &H9
        Put #iFileNum, RepointOffset + 1 + 11, &H9
        Put #iFileNum, RepointOffset + 1 + 15, &H9
    Close #iFileNum
End Sub

Private Sub Form_Load()
    SetIcon Me.hWnd, "AAA"
    Localize Me
    imgFlag.Picture = LoadResPicture("NULL", 0)
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuPalette_Click()
    frmPalette.Show vbModal, Me
End Sub

Private Sub mnuPreview_Click()
    frmPreview.Show , Me
End Sub

Private Sub mnuReadMe_Click()
Dim arrReadme() As Byte
    arrReadme = LoadResData("README", 100)
    WriteByteArray App.Path & "\Readme.txt", arrReadme, 0
    Shell "notepad.exe " & App.Path & "\Readme.txt", vbNormalFocus
    Kill App.Path & "\Readme.txt"
    Erase arrReadme
End Sub

Private Sub mnuUncompress_Click()
Dim cdgSaveUncompress As clsCommonDialog
Dim iExport As Integer
Dim sExport As String
Dim Export() As Byte
Dim tmp() As Byte 'Original size was 32768/0x8000. I putted 131072/0x20000

    Set cdgSaveUncompress = New clsCommonDialog
    sExport = cdgSaveUncompress.ShowSave(Me.hWnd, "Save Tileset" & "...", , , "Tilesets (*.raw)|*.raw|Binary Files & Dumps (*.bin,*.dmp)|*.bin;*.dmp", OVERWRITEPROMPT)
    If LenB(sExport) = 0 Then GoTo EndMe
    
    GetFileData sTilesetPath, tmp
    LZ77UnComp tmp, Export
    
    CleanFile sExport
    WriteByteArray sExport, Export, 0
    
    MsgBox LoadResString(21), vbInformation
    sTilesetPath = sExport
    txtTilesetPath.Text = sTilesetPath
    mnuUncompress.Enabled = False
    mnuCompress.Enabled = True
    
    Erase tmp, Export
EndMe:
    Set cdgSaveUncompress = Nothing
End Sub

Private Sub mnuCompress_Click()
Dim cdgSaveCompress As clsCommonDialog
Dim sTileset As String
Dim sReturn As String
Dim tmp() As Byte
Dim DecmpSize As Long
Dim Export() As Byte
Dim iExport As Integer
Dim sExport As String
    
    Set cdgSaveCompress = New clsCommonDialog
    sExport = cdgSaveCompress.ShowSave(Me.hWnd, "Save Tileset" & "...", , , "Tilesets (*.raw)|*.raw|Binary Files & Dumps (*.bin,*.dmp)|*.bin;*.dmp", OVERWRITEPROMPT)
    If LenB(sExport) = 0 Then GoTo EndMe
    
    GetFileData sTilesetPath, tmp
    CountFileBytes sTilesetPath, DecmpSize
    ReDim Export(DecmpSize - 1)
    DecmpSize = LZ77Comp(DecmpSize, tmp, Export)
    ReDim Preserve Export(DecmpSize - 1)

    CleanFile sExport
    WriteByteArray sExport, Export, 0
    
    MsgBox LoadResString(22), vbInformation
    sTilesetPath = sExport
    txtTilesetPath.Text = sTilesetPath
    mnuCompress.Enabled = False
    mnuUncompress.Enabled = True
    
    Erase Export, tmp
EndMe:
    Set cdgSaveCompress = Nothing
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String
Dim cdgOpen As clsCommonDialog

    iFileNum = FreeFile
    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, LoadResString(2) & "...", , "GameBoy Advance ROMs (*.gba,*.agb,*.bin)|*.gba;*.agb;*.bin")
    If LenB(sResult) = 0 Then GoTo EndMe
    sFilePath = sResult
       
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, &HAC + 1, sHeader
    
        If LOF(iFileNum) > 16777216 Then
            EnlargedROM = 1
            txtOffset.MaxLength = 7
        Else
            EnlargedROM = 0
            txtOffset.MaxLength = 6
        End If
        
    Close #iFileNum
        
    If OpenROM(sHeader) = False Then GoTo Unsupported
    
    Select Case Right(sHeader, 1)
        Case "D"
            imgFlag.Picture = LoadResPicture("GERMAN", 0)
        Case "E"
            imgFlag.Picture = LoadResPicture("US", 0)
        Case "F"
            imgFlag.Picture = LoadResPicture("FRANCE", 0)
        Case "I"
            imgFlag.Picture = LoadResPicture("ITALY", 0)
        Case "J"
            imgFlag.Picture = LoadResPicture("JAPAN", 0)
        Case "S"
            imgFlag.Picture = LoadResPicture("SPAIN", 0)
    End Select
    
    chMajor_Click
    txtFilePath.Text = sFilePath
    lblHeader.Caption = sHeader
    mnuOpenTileset.Enabled = True
    mnuPalette.Enabled = True
    If LenB(sTilesetPath) <> 0 Then cmdInsert.Enabled = True
    GoTo EndMe
    
Unsupported:
    sFilePath = vbNullString
    txtFilePath.Text = "???"
    lblHeader.Caption = "???"
    mnuPalette.Enabled = False
    imgFlag.Picture = LoadResPicture("NULL", 0)
    
EndMe:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuOpenTileset_Click()
Dim cdgOpenTileset As clsCommonDialog
Dim iFile As Integer
Dim bHeaderByte As Byte
    
    Set cdgOpenTileset = New clsCommonDialog
    sTilesetPath = cdgOpenTileset.ShowOpen(Me.hWnd, "Open Tileset...", , "Tilesets (*.raw)|*.raw|Binary Files (*.bin)|*.bin|Dumps (*.dmp)|*.dmp")
    If LenB(sTilesetPath) = 0 Then GoTo EndMe
    txtTilesetPath.Text = sTilesetPath
        
    iFile = FreeFile
    Open sTilesetPath For Binary As #iFile
    Get #iFile, &H0 + 1, bHeaderByte
        If bHeaderByte = &H10 Then
            mnuUncompress.Enabled = True
            mnuCompress.Enabled = False
            chCompression.Value = vbChecked
        Else
            mnuUncompress.Enabled = False
            mnuCompress.Enabled = True
            chCompression.Value = vbUnchecked
        End If
    Close #iFile
    
    cmdFind.Enabled = True
    mnuPreview.Enabled = True
    If LenB(sFilePath) <> 0 Then cmdInsert.Enabled = True
    
EndMe:
    Set cdgOpenTileset = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Function OpenROM(sHeader As String) As Boolean
    Select Case sHeader
        Case "AXVE"
            TilesetHeader = &H286CF4
        Case "AXPE"
            TilesetHeader = &H286C84
        Case "BPRE"
            TilesetHeader = &H2D4A94
        Case "BPGE"
            TilesetHeader = &H2D4A74
        Case "BPEE"
            TilesetHeader = &H3DF704
        Case Else
            MsgBox LoadResString(23) & vbNewLine & LoadResString(24), vbExclamation
            OpenROM = False
            Exit Function
    End Select
    OpenROM = True
End Function
