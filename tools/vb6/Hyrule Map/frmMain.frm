VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hyrule Map"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   12000
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
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPalette 
      Caption         =   "Palettes"
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   2775
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2535
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.PictureBox picCurrentTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   3120
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   17
      Top             =   5400
      Width           =   840
   End
   Begin VB.Frame fraTilesets 
      Caption         =   "Tilesets"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   2775
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2535
         TabIndex        =   13
         Top             =   240
         Width           =   2535
         Begin VB.ComboBox cmbTilesets 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Text            =   "Tilesets"
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin VB.Frame fraROMInformation 
      Caption         =   "ROM Information"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   2775
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   225
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROM:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   405
      End
   End
   Begin HyruleMap.xpWellsStatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      Top             =   6360
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   529
      BackColor       =   15790320
      ForeColor       =   -2147483630
      ForeColorDissabled=   -2147483631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfPanels  =   2
      MaskColor       =   16711935
      PWidth1         =   300
      pText1          =   "Welcome to Hyrule Map ^^"
      pTTText1        =   ""
      pEnabled1       =   -1  'True
      PWidth2         =   100
      pText2          =   "Current Tile: ????"
      pTTText2        =   ""
      pEnabled2       =   -1  'True
   End
   Begin VB.HScrollBar hsbMap 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   5040
      Width           =   6135
   End
   Begin VB.VScrollBar vsbMap 
      Enabled         =   0   'False
      Height          =   4935
      Left            =   9240
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picMapBox 
      Height          =   4920
      Left            =   3120
      ScaleHeight     =   324
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   4
      Top             =   120
      Width           =   6120
      Begin VB.PictureBox picMap 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   7680
         Left            =   0
         ScaleHeight     =   512
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   512
         TabIndex        =   15
         Top             =   0
         Width           =   7680
         Begin VB.Shape shpMapCursor 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   120
            Left            =   0
            Top             =   0
            Width           =   120
         End
      End
   End
   Begin VB.PictureBox picTilesetBox 
      Height          =   5160
      Left            =   9600
      ScaleHeight     =   340
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   132
      TabIndex        =   3
      Top             =   120
      Width           =   2040
      Begin VB.PictureBox picTileset 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5040
         Left            =   0
         ScaleHeight     =   336
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   16
         Top             =   0
         Width           =   1920
         Begin VB.Shape shpTilesetCursor 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   120
            Left            =   0
            Top             =   0
            Width           =   120
         End
      End
   End
   Begin VB.VScrollBar vsbTileset 
      Enabled         =   0   'False
      Height          =   5175
      Left            =   11640
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame fraMaps 
      Caption         =   "Maps"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.ListBox lstMaps 
         Enabled         =   0   'False
         Height          =   2205
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
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
Private sFilePath As String
Private iFileNum As Integer
Private sHeader As String * 4

Private Map As Long
Private MapTable As Long
Private MapHeader As Long

Private Tileset As Long
Private TilesetTable As Long
Private TilesetHeader As Long

Private Const Tilesize = 8
Private iCurrentTile As Integer

Private Function LoadStuff(cnt As Control, ItemDescription As String, HeaderOffset As Long) As Long
Dim bTemp As Byte
Dim Table As Long
Dim Counter As Integer
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, HeaderOffset + 1, Table
        Table = Table - &H8000000
        
        Counter = 0
        cnt.Clear
        Do
            Get #iFileNum, Table + 1 + 3 + (Counter * 4), bTemp
            If bTemp <> &H8 Then Exit Do
            Counter = Counter + 1
            cnt.AddItem ItemDescription & "  " & Right("000" & Counter, 3)
        Loop
    Close #iFileNum
    LoadStuff = Table
End Function

Private Sub cmbTilesets_Click()
Dim i As Integer
Dim Temp As Long
Dim arrPalette() As Byte
Dim arrTemp(32192) As Byte
Dim arrtileset() As Byte
Dim arrPal(0 To 256) As Integer
Dim PCPalette(0 To 16, 0 To 16) As Long
    arrPalette = LoadResData("RGB", "PAL")
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, TilesetTable + 1 + (cmbTilesets.ListIndex * 4), Temp
        Temp = Temp - &H8000000
        Get #iFileNum, Temp + 1, arrTemp
        LZ77UnComp arrTemp, arrtileset
        
        Temp = 0
        For i = 0 To 15 'PalSize / 2
            arrPal(i) = CInt(arrPalette(Temp + 1) * &H100 + arrPalette(Temp))
            Temp = Temp + 2
        Next i
        UnPackPalette arrPal, PCPalette
        
        picTileset.Cls
        For i = 0 To 255
            On Error Resume Next
            DrawTile8 picTileset.hDC, i, arrtileset, PCPalette
        Next i

    Close #iFileNum
    Erase arrPalette, arrTemp, arrtileset, arrPal, PCPalette
End Sub

Private Sub mnuOpen_Click()
Dim cnt As Control
Dim sResult As String
Dim cdgOpen As clsCommonDialog
    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, "Open ROM...", , "Gameboy Advance ROMs (*.gba;*.agb;*.bin)|*.gba;*.agb;*.bin")
    If LenB(sResult) = 0 Then GoTo EndMe
    sFilePath = sResult
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, &HAC + 1, sHeader
    Close #iFileNum
    
    Select Case sHeader
        Case "AZLP"
            MapHeader = &H80B74
            TilesetHeader = &H12FD8C
            lblName.Caption = "A Link to the Past"
            
        Case Else
            MsgBox "Error - Unsupported ROM :P", vbExclamation
            GoTo EndMe
            
    End Select
    
    lblHeader.Caption = sHeader
    MapTable = LoadStuff(lstMaps, "Map", MapHeader)
    TilesetTable = LoadStuff(cmbTilesets, "Tileset", TilesetHeader)
    
    For Each cnt In Me.Controls
        On Error Resume Next
        If cnt.Enabled = False Then cnt.Enabled = True
    Next
    
EndMe:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim MyX As Integer
Dim MyY As Integer
    MyX = x \ Tilesize
    MyY = Y \ Tilesize
    If Button <> vbLeftButton Then Exit Sub
    
    shpMapCursor.Move MyX * Tilesize, MyY * Tilesize
    shpMapCursor.Visible = True
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picMap_MouseDown Button, Shift, x, Y
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim MyX As Integer
Dim MyY As Integer
    MyX = x \ Tilesize
    MyY = Y \ Tilesize
    If Button <> vbLeftButton Then Exit Sub
    
    iCurrentTile = (MyY * Tilesize) + MyX
    StatusBar.PanelCaption(2) = "Current Tile: " & iCurrentTile
    StretchBlt picCurrentTile.hDC, 0, 0, 56, 56, picTileset.hDC, MyX * Tilesize, MyY * Tilesize, 8, 8, vbSrcCopy
    picCurrentTile.Refresh
    shpTilesetCursor.Move MyX * Tilesize, MyY * Tilesize
    shpTilesetCursor.Visible = True
End Sub

Private Sub picTileset_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picTileset_MouseDown Button, Shift, x, Y
End Sub
