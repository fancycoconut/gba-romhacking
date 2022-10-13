VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance IntroEd"
   ClientHeight    =   3585
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   7215
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
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraIntro 
      Caption         =   "Game Intro Settings"
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   4335
      Begin VB.PictureBox picIntro 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   273
         TabIndex        =   18
         Top             =   240
         Width           =   4095
         Begin VB.PictureBox picPoke 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   975
            Left            =   2760
            ScaleHeight     =   975
            ScaleWidth      =   1215
            TabIndex        =   28
            Top             =   0
            Width           =   1215
         End
         Begin VB.ComboBox cmbPoke 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":000C
            Left            =   960
            List            =   "frmMain.frx":04E4
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtIntroSong 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            MaxLength       =   3
            TabIndex        =   21
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblPokemon 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pokemon"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lblIntroSong 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Song"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   360
         End
      End
   End
   Begin VB.Frame fraStartMap 
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   6735
      Begin VB.PictureBox picStartMap 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   433
         TabIndex        =   8
         Top             =   240
         Width           =   6495
         Begin VB.CheckBox chkRemove 
            Caption         =   "Remove Item"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtAmount 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   26
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cmbItems 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdDefault 
            Caption         =   "Default"
            Enabled         =   0   'False
            Height          =   300
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtMap 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtBank 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   14
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.Frame fraMoney 
            Caption         =   "Start Money"
            Height          =   855
            Left            =   4440
            TabIndex        =   9
            Top             =   0
            Width           =   1935
            Begin VB.PictureBox picMoney 
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   120
               ScaleHeight     =   33
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   113
               TabIndex        =   10
               Top             =   240
               Width           =   1695
               Begin VB.TextBox txtMoney 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   240
                  MaxLength       =   9
                  TabIndex        =   11
                  Top             =   120
                  Width           =   1215
               End
            End
         End
         Begin VB.Label lblAmount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   195
            Left            =   3480
            TabIndex        =   25
            Top             =   0
            Width           =   555
         End
         Begin VB.Label lblPCItem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PC Item"
            Height          =   195
            Left            =   1800
            TabIndex        =   23
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lblMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map"
            Height          =   195
            Left            =   960
            TabIndex        =   13
            Top             =   0
            Width           =   300
         End
         Begin VB.Label lblBank 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   0
            Width           =   345
         End
      End
   End
   Begin VB.Frame fraROMInformation 
      Caption         =   "ROM Information"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox picFlag 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6360
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   34
         Top             =   240
         Width           =   495
         Begin VB.Image imgFlag 
            Height          =   165
            Left            =   150
            Top             =   30
            Width           =   240
         End
         Begin VB.Shape shpFlag 
            BorderColor     =   &H00C0C0C0&
            Height          =   225
            Left            =   120
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing Loaded..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame fraTruck 
      Caption         =   "Truck Settings"
      Height          =   1335
      Left            =   4560
      TabIndex        =   0
      Top             =   840
      Width           =   2535
      Begin VB.PictureBox picTruck 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   153
         TabIndex        =   1
         Top             =   240
         Width           =   2295
         Begin VB.TextBox txtDefaultSong 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "0"
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove Animation"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblSong 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Song"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   930
         End
      End
   End
   Begin VB.Frame fraTitlescreen 
      Caption         =   "Titlescreen"
      Height          =   1335
      Left            =   4560
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
      Begin VB.PictureBox picTitlescreen 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   153
         TabIndex        =   30
         Top             =   240
         Width           =   2295
         Begin VB.CommandButton cmdCry 
            Caption         =   "Default"
            Height          =   300
            Left            =   600
            TabIndex        =   33
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cmbCry 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":13F3
            Left            =   360
            List            =   "frmMain.frx":18CB
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblCry 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pokemon Cry"
            Height          =   195
            Left            =   600
            TabIndex        =   31
            Top             =   0
            Width           =   945
         End
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
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuReadme 
         Caption         =   "Readme"
         Shortcut        =   {F1}
      End
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
Private sFilePath As String
Private sHeader As String * 4

Private Cry As Long
Private StartMap As Long
Private StartItem As Long
Private StartMoney As Long
Private IntroSong As Long

Private ItemNames As Long
Private PokemonPics As Long
Private PokemonPalettes As Long

Private StartPokePic As Long
Private StartPokePal As Long
Private StartPokeIndex As Long

Private Sub cmbPoke_Click()
    LoadPokeImage cmbPoke.ListIndex, picPoke
End Sub

Private Sub cmdCry_Click()
If Left(sHeader, 3) = "BPR" Then
    cmbCry.ListIndex = 6
Else
    cmbCry.ListIndex = 3
End If
End Sub

Private Sub cmdDefault_Click()
If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
    txtBank.Text = 4
    txtMap.Text = 1
Else
    txtBank.Text = 25
    txtMap.Text = 40
End If
End Sub

Private Function LeftShift(ByVal Value As Long, ByVal iShift As Integer)
    LeftShift = Value * (2 ^ iShift)
End Function

Private Sub LoadItems()
Dim i As Integer
Dim bSapp2Asc() As Byte
Dim iItemCount As Integer
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Select Case Left(sHeader, 3)
            Case "AXV", "AXP"
                iItemCount = &H15C
            Case "BPR", "BPG"
                iItemCount = &H176
            Case "BPE"
                iItemCount = &H178
        End Select
        
        ReDim bSapp2Asc(43) As Byte ' Getting Item Names
        Seek #iFileNum, ItemNames + 1
        For i = 0 To iItemCount
            Get #iFileNum, , bSapp2Asc
            cmbItems.AddItem Sapp2Asc(bSapp2Asc)
        Next i
    Close #iFileNum
    Erase bSapp2Asc
End Sub

Private Sub LoadPokeImage(ByVal Index As Integer, PictureBox As Control)
Dim i As Integer
Dim arrGFX() As Byte
Dim iFile As Integer
Dim TempOffset As Long
Dim SomeCounter As Long
Dim arrPalette() As Byte
Dim arrTemp(32192) As Byte
Dim arrPal(0 To 256) As Integer
Dim PCPalette(0 To 16, 0 To 16) As Long
    
    iFile = FreeFile
    Open sFilePath For Binary As #iFile
        Get #iFile, PokemonPics + 1 + (8 * Index), TempOffset ' Getting pointer to graphics data
        Get #iFile, TempOffset - &H8000000 + 1, arrTemp ' Getting graphics data to an array
        LZ77UnComp arrTemp, arrGFX ' Uncompressing graphics
        Erase arrTemp
        
        Get #iFile, PokemonPalettes + 1 + (8 * Index), TempOffset ' Getting pointer to palette data
        Get #iFile, TempOffset - &H8000000 + 1, arrTemp ' Getting palette data to an array
        LZ77UnComp arrTemp, arrPalette ' Uncompressing palettes
        
        SomeCounter = 0 ' Converting the palette array into a 2D array
        For i = 0 To 15 ' 16 Colors ^^
            'On Error Resume Next
            arrPal(i) = CInt(arrPalette(SomeCounter + 1) * &H100 + arrPalette(SomeCounter))
            SomeCounter = SomeCounter + 2
        Next i
    
        UnPackPalette arrPal, PCPalette ' Unpacking palette into a displayable form
        
        PictureBox.Cls
        For i = 0 To 63 ' ploting image - 0 to 64 since its height and width is 64
            DrawTile8 PictureBox.hDC, i, arrGFX, PCPalette
        Next i
    
    Close #iFile
    Erase arrGFX, arrPalette, arrTemp, arrPal, PCPalette
End Sub

Private Sub LoadStuff()
Dim Temp As Long
Dim bTemp As Byte
Dim bTemp2 As Byte
Dim iTemp As Integer
    Select Case Left(sHeader, 3)
        Case "AXV", "AXP", "BPE"
            fraTruck.Visible = True
            fraTitlescreen.Visible = False
        Case Else
            fraTruck.Visible = False
            fraTitlescreen.Visible = True
    End Select
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, StartMap + 1, bTemp
        txtBank.Text = bTemp
        
        Get #iFileNum, StartMap + 2 + 1, bTemp
        txtMap.Text = bTemp
        
        Get #iFileNum, StartMoney + 1, Temp
        txtMoney = Temp
        
        Get #iFileNum, StartItem + 1, iTemp
        cmbItems.ListIndex = iTemp
        
        Get #iFileNum, StartItem + 2 + 1, iTemp
        txtAmount.Text = iTemp
        
        Get #iFileNum, IntroSong + 1, bTemp
        txtIntroSong.Text = Hex(LeftShift(bTemp, 1))
        
        If Left(sHeader, 3) = "AXV" Or Left(sHeader, 3) = "AXP" Or Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
            Get #iFileNum, StartPokePic + 1, Temp
            Temp = Temp - &H8000000
            Get #iFileNum, Temp + 6 + 1, iTemp
        Else ' Emerald Only
            Get #iFileNum, StartPokeIndex + 1, iTemp
        End If
        cmbPoke.ListIndex = iTemp
        
        If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
            Get #iFileNum, Cry + 3 + 1, bTemp2
            If bTemp2 = &H21 Then
                Get #iFileNum, Cry + 1, bTemp
                cmbCry.ListIndex = bTemp
            Else
                Get #iFileNum, Cry + 2 + 1, bTemp
                cmbCry.ListIndex = bTemp + 255
            End If
        End If
    Close #iFileNum
End Sub

Private Sub cmdRemove_Click()
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Select Case sHeader
            Case "AXVE"
                PutFiller sFilePath, &HC757E, 2
                PutFiller sFilePath, &HC759E, 4
                PutFiller sFilePath, &HC75B4, 2
                PutFiller sFilePath, &HC75D8, 2
                PutFiller sFilePath, &HC75E2, 2
                PutFiller sFilePath, &HC75F0, 4
                PutFiller sFilePath, &HC7600, 4
                PutFiller sFilePath, &HC7624, 4
                PutFiller sFilePath, &HC7640, 2
                PutFiller sFilePath, &HC7644, 4
                PutFiller sFilePath, &HC765E, 2
                PutFiller sFilePath, &HC7668, 4
                PutFiller sFilePath, &HC7674, 4
                PutFiller sFilePath, &HC7680, 14
                PutFiller sFilePath, &HC768A, 4
                
                If Val(txtDefaultSong.Text) = 0 Then GoTo EndMe ' No Song patch
                Put #iFileNum, &HC7680 + 6 + 1, CByte(RightShift(CInt("&H" & txtDefaultSong.Text), 1))
                Put #iFileNum, &HC7680 + 7 + 1, CByte(&H20)
                Put #iFileNum, &HC7680 + 8 + 1, CInt(&H40)
                Put #iFileNum, &HC7680 + 10 + 1, &HFEF3F7AD ' Branch link data
                
            Case "AXPE"
                PutFiller sFilePath, &HC757E, 2
                PutFiller sFilePath, &HC759E, 4
                PutFiller sFilePath, &HC75B4, 2
                PutFiller sFilePath, &HC75D8, 2
                PutFiller sFilePath, &HC75E2, 2
                PutFiller sFilePath, &HC75F0, 4
                PutFiller sFilePath, &HC7600, 4
                PutFiller sFilePath, &HC7624, 4
                PutFiller sFilePath, &HC7640, 2
                PutFiller sFilePath, &HC7644, 4
                PutFiller sFilePath, &HC765E, 2
                PutFiller sFilePath, &HC7668, 4
                PutFiller sFilePath, &HC7674, 4
                PutFiller sFilePath, &HC7680, 14
                PutFiller sFilePath, &HC768A, 4
                
                If Val(txtDefaultSong.Text) = 0 Then GoTo EndMe ' No Song patch
                Put #iFileNum, &HC7680 + 6 + 1, CByte(RightShift(CInt("&H" & txtDefaultSong.Text), 1))
                Put #iFileNum, &HC7680 + 7 + 1, CByte(&H20)
                Put #iFileNum, &HC7680 + 8 + 1, CInt(&H40)
                Put #iFileNum, &HC7680 + 10 + 1, &HFEF3F7AD ' Branch link data
                
            Case "BPEE"
                PutFiller sFilePath, &HFB3BE, 2
                PutFiller sFilePath, &HFB3DE, 4
                PutFiller sFilePath, &HFB3F4, 2
                PutFiller sFilePath, &HFB418, 2
                PutFiller sFilePath, &HFB422, 2
                PutFiller sFilePath, &HFB430, 4
                PutFiller sFilePath, &HFB440, 4
                PutFiller sFilePath, &HFB464, 4
                PutFiller sFilePath, &HFB480, 2
                PutFiller sFilePath, &HFB484, 4
                PutFiller sFilePath, &HFB49E, 2
                PutFiller sFilePath, &HFB4A8, 4
                PutFiller sFilePath, &HFB4B4, 4
                PutFiller sFilePath, &HFB4BC, 8
                PutFiller sFilePath, &HFB4CA, 4
                
                If Val(txtDefaultSong.Text) = 0 Then GoTo EndMe ' No Song patch
                Put #iFileNum, &HFB4BC + 1, CByte(RightShift(CInt("&H" & txtDefaultSong.Text), 1))
                Put #iFileNum, &HFB4BC + 1 + 1, CByte(&H20)
                Put #iFileNum, &HFB4BE + 1, CInt(&H40)
                Put #iFileNum, &HFB4C0 + 1, &HF95AF7A8 ' Branch link data
        End Select
EndMe:
    Close #iFileNum
    MsgBox "Truck Animation removed successfully!", vbInformation
End Sub

Private Sub Form_Load()
    SetIcon Me.hWnd, "AAA"
    imgFlag.Picture = LoadResPicture("NULL", 0)
    picPoke.BackColor = vbButtonFace
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
Dim cnt As Control
Dim sResult As String
Dim cdgOpen As clsCommonDialog
    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, "Open ROM" & "...", , "Gameboy Advance ROMs (*agb,*.gba,*.bin)|*.agb;*.gba;*.bin")
    If LenB(sResult) = 0 Then GoTo EndMe
    
    iFileNum = FreeFile
    Open sResult For Binary As #iFileNum
        Get #iFileNum, &HAC + 1, sHeader
        
        Select Case sHeader
            Case "AXVE"
                lblROM.Caption = "Pokémon Ruby" & " - " & sHeader
                StartMap = &H52E0E
                IntroSong = 41630 ' &HA29E
                ItemNames = &H3C5564
                StartItem = &H4062F0
                StartMoney = &H52F4C
                StartPokePic = 45752 ' &HB2B8
                StartPokePal = 45764 ' &HB2C4
                StartPokeIndex = 45702 ' &HB286
                PokemonPics = &H1E8354
                PokemonPalettes = &H1EA5B4
            
            Case "AXPE"
                lblROM.Caption = "Pokémon Sapphire" & " - " & sHeader
                StartMap = &H52E0E
                IntroSong = 41630 ' &HA29E
                ItemNames = &H3C55BC
                StartItem = &H406348
                StartMoney = &H52F4C
                StartPokePic = 45752 ' &HB2B8
                StartPokePal = 45764 ' &HB2C4
                StartPokeIndex = 45702 ' &HB286
                PokemonPics = &H1E82E4
                PokemonPalettes = &H1EA544
                
            Case "BPRE"
                lblROM.Caption = "Pokémon Fire Red" & " - " & sHeader
                Cry = &H791EE
                StartMap = &H54A04
                IntroSong = &H12F836
                ItemNames = &H3DB028
                StartItem = &H402220
                StartMoney = &H54B60
                StartPokePic = &H130FA0
                StartPokePal = &H130FA4
                StartPokeIndex = &H130F4C
                PokemonPics = &H2350AC
                PokemonPalettes = &H23730C
            
            Case "BPGE"
                lblROM.Caption = "Pokémon Leaf Green" & " - " & sHeader
                Cry = &H791EE
                StartMap = &H54A04
                IntroSong = &H12F80E
                ItemNames = &H3DAE64
                StartItem = &H40205C
                StartMoney = &H54B60
                StartPokePic = &H130F78
                StartPokePal = &H130F7C
                StartPokeIndex = &H130F24
                PokemonPics = &H235088
                PokemonPalettes = &H2372E8
                
            Case "BPRI"
                Cry = &H7913E
                StartMap = &H54A10
                IntroSong = &H12F8C6
                ItemNames = &H3D1EE8
                StartItem = &H3F99B0
                StartMoney = &H54B6C
                StartPokePic = &H131030
                StartPokePal = &H131034
                StartPokeIndex = &H130FDC
                PokemonPics = &H22E150
                PokemonPalettes = &H2303B0
            
            Case "BPEE"
                lblROM.Caption = "Pokémon Emerald" & " - " & sHeader
                StartMap = &H84456
                IntroSong = &H30872
                ItemNames = &H5839A0
                StartItem = &H5DFEFC
                StartMoney = &H845BC
                StartPokeIndex = &H31924
                PokemonPics = &H30A18C
                PokemonPalettes = &H303678
                
            Case Else
                MsgBox "Error - Unsupported ROM" & vbNewLine & "Sorry :P", vbExclamation
                ResetForm
                GoTo EndMe
                
        End Select
    Close #iFileNum
    
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
    
    sFilePath = sResult
    mnuSave.Enabled = True
    lblROM.ToolTipText = sFilePath
    
    For Each cnt In Me.Controls
        On Error Resume Next
        If cnt.Enabled = False Then cnt.Enabled = True
    Next
    
    LoadItems
    LoadStuff
EndMe:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Function RightShift(ByVal Value As Long, ByVal iShift As Integer)
    RightShift = Value \ (2 ^ iShift)
End Function

Private Sub mnuReadme_Click()
Dim arrTemp() As Byte
    arrTemp = LoadResData("README", 100)
    WriteByteArray App.Path & "\Readme.txt", arrTemp, 0
    Shell "notepad.exe " & App.Path & "\Readme.txt", vbNormalFocus
    Kill App.Path & "\Readme.txt"
    Erase arrTemp
End Sub

Private Sub mnuSave_Click()
Dim Temp As Long
Dim Opcode As Byte
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Put #iFileNum, StartMap + 1, CByte(txtBank.Text)
        Put #iFileNum, StartMap + 2 + 1, CByte(txtMap.Text)
        Put #iFileNum, StartMoney + 1, CLng(txtMoney.Text)
        Put #iFileNum, StartItem + 1, cmbItems.ListIndex
        Put #iFileNum, StartItem + 2 + 1, CInt(txtAmount.Text)
        Put #iFileNum, IntroSong + 1, CByte(RightShift(CInt("&H" & txtIntroSong.Text), 1))
        
        If chkRemove.Value = vbChecked Then Put #iFileNum, StartItem + 1, CLng(&H0)
        
        ' Changing Intro Pokemon
        If Left(sHeader, 3) = "AXV" Or Left(sHeader, 3) = "AXP" Or Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
            ' Ruby, Sapphire, Fire Red & Leaf Green Only
            Temp = PokemonPics + (cmbPoke.ListIndex * 8)
            Put #iFileNum, StartPokePic + 1, Temp + &H8000000
            Temp = PokemonPalettes + (cmbPoke.ListIndex * 8)
            Put #iFileNum, StartPokePal + 1, Temp + &H8000000
            
            If Left(sHeader, 3) = "AXV" Or Left(sHeader, 3) = "AXP" Then Opcode = &H34 ' 34YY add r4, r4, #0xYY
            If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then Opcode = &H30 ' 30YY add r0, r0, #0xYY
            
            If cmbPoke.ListIndex > 255 Then
                Put #iFileNum, StartPokeIndex + 1, CByte(255)
                Put #iFileNum, StartPokeIndex + 2 + 1, CByte(cmbPoke.ListIndex - 255)
                Put #iFileNum, StartPokeIndex + 3 + 1, CByte(Opcode)
            Else
                Put #iFileNum, StartPokeIndex + 1, CByte(cmbPoke.ListIndex)
                Put #iFileNum, StartPokeIndex + 2 + 1, CInt(&H0)
            End If
            
        Else ' Emerald Only
            Put #iFileNum, StartPokeIndex + 1, cmbPoke.ListIndex
        End If

        ' Changing Pokemon Cry in titlescreen
        If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
            Opcode = &H30 ' 30YY add r0, r0, #0xYY
            If cmbCry.ListIndex > 255 Then
                Put #iFileNum, Cry + 1, CByte(255)
                Put #iFileNum, Cry + 2 + 1, CByte(cmbCry.ListIndex - 255)
                Put #iFileNum, Cry + 3 + 1, CByte(Opcode)
            Else
                Put #iFileNum, Cry + 1, CByte(cmbCry.ListIndex)
                Put #iFileNum, Cry + 2 + 1, CInt(&H2100)
            End If
        End If
    Close #iFileNum
    LoadStuff ' Reload values
End Sub

Private Sub ResetForm()
    txtMap.Enabled = False
    txtBank.Enabled = False
    mnuSave.Enabled = False
    cmbPoke.Enabled = False
    fraTruck.Visible = True
    sFilePath = vbNullString
    cmbItems.Enabled = False
    txtMoney.Enabled = False
    txtAmount.Enabled = False
    chkRemove.Enabled = False
    cmdRemove.Enabled = False
    cmdDefault.Enabled = False
    txtIntroSong.Enabled = False
    txtDefaultSong.Enabled = False
    fraTitlescreen.Visible = False
    lblROM.Caption = "Nothing Loaded..."
    imgFlag.Picture = LoadResPicture("NULL", 0)
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub txtBank_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub txtMap_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub txtMoney_(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
