VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trainer Tool"
   ClientHeight    =   3720
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   7440
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
   ScaleWidth      =   496
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTrainerImage 
      Caption         =   "Trainer Image Table"
      Height          =   1095
      Left            =   4560
      TabIndex        =   18
      Top             =   360
      Width           =   2775
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   2535
         TabIndex        =   19
         Top             =   240
         Width           =   2535
         Begin VB.TextBox txtTrainerIndex 
            ForeColor       =   &H80000011&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtUnknown 
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblTrainerIndex 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trainer #"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   675
         End
         Begin VB.Label lblUnknown 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unknown"
            Height          =   195
            Left            =   1440
            TabIndex        =   21
            Top             =   120
            Width           =   660
         End
      End
   End
   Begin VB.PictureBox picTrainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   2520
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1800
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2400
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Frame fraROM 
      Caption         =   "ROM Information"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   7215
      Begin VB.PictureBox picFlag 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6600
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   24
         Top             =   360
         Width           =   495
         Begin VB.Image imgFlag 
            Height          =   165
            Left            =   150
            Top             =   150
            Width           =   240
         End
         Begin VB.Shape shpFlag 
            BorderColor     =   &H00C0C0C0&
            Height          =   225
            Left            =   120
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   225
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROM:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Frame fraPointers 
      Caption         =   "Pointers"
      Height          =   1095
      Left            =   4560
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   2535
         TabIndex        =   4
         Top             =   240
         Width           =   2535
         Begin VB.TextBox txtPalette 
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txt2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   9
            Text            =   "0x"
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txt1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Text            =   "0x"
            Top             =   360
            Width           =   270
         End
         Begin VB.TextBox txtGraphics 
            Height          =   285
            Left            =   360
            MaxLength       =   6
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblPalette 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Palette"
            Height          =   195
            Left            =   1440
            TabIndex        =   6
            Top             =   120
            Width           =   510
         End
         Begin VB.Label lblGraphics 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Graphics"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   615
         End
      End
   End
   Begin VB.Frame fraTrainers 
      Caption         =   "Trainers"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.ListBox lstTrainers 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
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
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTrainers 
      Caption         =   "&Trainers"
      Begin VB.Menu mnuRepoint 
         Caption         =   "Edit New Trainer Amount"
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
Private iFileNum As Integer
Private sHeader As String * 4

Public sFilePath As String
Public TrainerPics As Long
Public TrainerPals As Long
Public iTrainerAmount As Integer

Public Sub DrawTrainer(iFile As Integer, GraphicsPointer As Long, PalettePointer As Long, PictureBox As Control)
Dim i As Integer
Dim SomeCounter As Long
Dim arrPalette() As Byte
Dim arrGraphics() As Byte
Dim arrTemp(32192) As Byte
Dim arrPal(0 To 256) As Integer
Dim PCPalette(0 To 16, 0 To 16) As Long
    Get #iFile, GraphicsPointer + 1, arrTemp
    LZ77UnComp arrTemp, arrGraphics
    Erase arrTemp
    
    Get #iFile, PalettePointer + 1, arrTemp
    LZ77UnComp arrTemp, arrPalette
    
    SomeCounter = 0
    For i = 0 To 15
        On Error GoTo EndMe
        arrPal(i) = CInt(arrPalette(SomeCounter + 1) * &H100 + arrPalette(SomeCounter))
        SomeCounter = SomeCounter + 2
    Next i
    
    UnPackPalette arrPal, PCPalette
    
    PictureBox.Cls
    For i = 0 To 63
        DrawTile8 PictureBox.hDC, i, arrGraphics, PCPalette
    Next i
   
EndMe:
    Erase arrPalette, arrGraphics, arrTemp, arrPal, PCPalette
End Sub

Public Sub LoadTrainers(iFile As Integer)
Dim i As Integer
Dim bTemp As Byte
    i = 0
    lstTrainers.Clear
    iTrainerAmount = 0
    Do
        Get #iFile, TrainerPics + 1 + 6 + i, bTemp
        If bTemp <> iTrainerAmount Then Exit Do
        iTrainerAmount = iTrainerAmount + 1
        i = i + 8
        lstTrainers.AddItem "Trainer" & " #" & Right("00" & iTrainerAmount, 3)
    Loop
End Sub

Private Sub cmdSave_Click()
    mnuSave_Click
End Sub

Private Sub Form_Load()
    imgFlag.Picture = LoadResPicture("NULL", 0)
End Sub

Private Sub lstTrainers_Click()
Dim Temp As Long
Dim iTemp As Integer
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, TrainerPics + 1 + 4 + (lstTrainers.ListIndex * 8), iTemp
        txtUnknown.Text = Hex(iTemp)
        
        Get #iFileNum, TrainerPics + 1 + 6 + (lstTrainers.ListIndex * 8), iTemp
        txtTrainerIndex.Text = Hex(iTemp)
    
        Get #iFileNum, TrainerPics + 1 + (lstTrainers.ListIndex * 8), Temp
        Temp = Temp - &H8000000
        txtGraphics.Text = Hex(Temp)
        
        Get #iFileNum, TrainerPals + 1 + (lstTrainers.ListIndex * 8), Temp
        Temp = Temp - &H8000000
        txtPalette.Text = Hex(Temp)
        
        If Temp = 0 Then GoTo EndMe
        
        DrawTrainer iFileNum, CLng("&H" & txtGraphics.Text), CLng("&H" & txtPalette.Text), picTemp
        StretchBlt picTrainer.hDC, 0, 0, 120, 120, picTemp.hDC, 0, 0, 64, 64, vbSrcCopy
        
EndMe:
        picTrainer.Refresh
        picTemp.Cls
    Close #iFileNum
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String
Dim cdgOpen As clsCommonDialog
    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, "Open ROM", , "Gameboy Advance ROMs (*.gba,*.agb,*.bin)|*.gba;*.agb;*.bin")
    If LenB(sResult) = 0 Then GoTo EndMe
    sFilePath = sResult
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, &HAC + 1, sHeader
    
        Select Case sHeader
            Case "AXVE"
                lblName.Caption = "Pokémon Ruby"
                Get #iFileNum, &H31ADC + 1, TrainerPics
                Get #iFileNum, &H31AF0 + 1, TrainerPals
                
            Case "BPRE"
                lblName.Caption = "Pokémon Fire Red"
                Get #iFileNum, &H3473C + 1, TrainerPics
                Get #iFileNum, &H3474C + 1, TrainerPals
                
            Case "BPEE"
                lblName.Caption = "Pokémon Emerald"
                Get #iFileNum, &H5DF78 + 1, TrainerPics
                Get #iFileNum, &H5B784 + 1, TrainerPals
                
            Case Else
                picTrainer.Cls
                Close #iFileNum
                lstTrainers.Clear
                cmdSave.Enabled = False
                mnuSave.Enabled = False
                sFilePath = vbNullString
                mnuRepoint.Enabled = False
                lstTrainers.Enabled = False
                GoTo EndMe
        End Select
        
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
                imgFlag.Picture = LoadResPicture("SPANISH", 0)
        End Select
        
        TrainerPics = TrainerPics - &H8000000
        TrainerPals = TrainerPals - &H8000000
        
        LoadTrainers iFileNum
        
        cmdSave.Enabled = True
        mnuSave.Enabled = True
        lstTrainers.Enabled = True
        mnuRepoint.Enabled = True
        lstTrainers.ListIndex = 0
        lblHeader.Caption = sHeader
    
    Close #iFileNum
EndMe:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuRepoint_Click()
    frmRepoint.Show vbModal, Me
End Sub

Private Sub mnuSave_Click()
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Put #iFileNum, TrainerPics + 1 + 4 + (lstTrainers.ListIndex * 8), CInt("&H" & txtUnknown.Text)
        Put #iFileNum, TrainerPics + 1 + (lstTrainers.ListIndex * 8), CLng("&H" & txtGraphics.Text) + &H8000000
        Put #iFileNum, TrainerPals + 1 + (lstTrainers.ListIndex * 8), CLng("&H" & txtPalette.Text) + &H8000000
    Close #iFileNum
End Sub
