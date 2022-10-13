VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Door Manager - Editor Mode On"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10320
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   688
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNavigator 
      Caption         =   "Door Navigator"
      Height          =   1215
      Left            =   7080
      TabIndex        =   43
      Top             =   2280
      Width           =   1935
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1695
         TabIndex        =   44
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "<"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1320
            TabIndex        =   47
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtFrame 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblFrame 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frame"
            Height          =   195
            Left            =   600
            TabIndex        =   45
            Top             =   120
            Width           =   450
         End
      End
   End
   Begin VB.Frame fraPalette 
      Caption         =   "Palette"
      Height          =   3375
      Left            =   9120
      TabIndex        =   26
      Top             =   120
      Width           =   1095
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   15
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2880
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   14
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2520
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   13
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2160
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   12
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1800
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   11
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1440
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   10
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1080
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   720
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2880
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2520
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2160
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1800
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1440
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1080
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   720
         WhatsThisHelpID =   3
         Width           =   255
      End
      Begin VB.OptionButton optColors 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         WhatsThisHelpID =   3
         Width           =   255
      End
   End
   Begin VB.Frame fraCanvas 
      Caption         =   "Drawing Canvas"
      Height          =   2055
      Left            =   7080
      TabIndex        =   24
      Top             =   120
      Width           =   1935
      Begin DoorManager.GBATileEditor TileEdit 
         Height          =   1440
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   2540
         DotSize         =   12
      End
   End
   Begin VB.Frame fraROM 
      Caption         =   "ROM Information"
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   6735
      Begin VB.PictureBox picFlag 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6120
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   49
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
         TabIndex        =   23
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   225
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROM:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.TextBox txtDoorType 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   17
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox picDoor 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   480
      Left            =   2760
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      Top             =   120
      Width           =   3840
      Begin VB.Shape shpCursor 
         BorderColor     =   &H00000FF0&
         Height          =   480
         Left            =   0
         Shape           =   1  'Square
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox cmbPal 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":000C
      Left            =   5520
      List            =   "frmMain.frx":0040
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame fraPointers 
      Caption         =   "Pointers"
      Height          =   1095
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   4
         Top             =   240
         Width           =   2655
         Begin VB.TextBox txtDesign2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0x"
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtDesign 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0x"
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtPalette 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtGraphics 
            Enabled         =   0   'False
            Height          =   285
            Left            =   360
            MaxLength       =   6
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblPalette 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Palette Data"
            Height          =   195
            Left            =   1440
            TabIndex        =   7
            Top             =   120
            Width           =   900
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
   Begin VB.Frame fraDoors 
      Caption         =   "Doors"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.ListBox lstDoors 
         Enabled         =   0   'False
         Height          =   1815
         ItemData        =   "frmMain.frx":0104
         Left            =   240
         List            =   "frmMain.frx":0106
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox txtTileNumber 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblDoorType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Door Type"
      Height          =   195
      Left            =   4200
      TabIndex        =   18
      Top             =   720
      Width           =   750
   End
   Begin VB.Label lblTileNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tile Number"
      Height          =   195
      Left            =   2880
      TabIndex        =   16
      Top             =   720
      Width           =   840
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palette Index"
      Height          =   195
      Left            =   5520
      TabIndex        =   9
      Top             =   1440
      Width           =   975
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
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuEditorMode 
         Caption         =   "EditorMode"
      End
      Begin VB.Menu mnuGridlines 
         Caption         =   "Gridlines"
      End
      Begin VB.Menu mnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPalettes 
         Caption         =   "Palettes"
         Begin VB.Menu mnuPal 
            Caption         =   "RSE Palette 0"
            Index           =   1
         End
         Begin VB.Menu mnuPal 
            Caption         =   "RSE Palette 1"
            Index           =   2
         End
         Begin VB.Menu mnuPal 
            Caption         =   "RSE Palette 2"
            Index           =   3
         End
         Begin VB.Menu mnuPal 
            Caption         =   "RSE Palette 3"
            Index           =   4
         End
         Begin VB.Menu mnuPal 
            Caption         =   "RSE Palette 4"
            Index           =   5
         End
         Begin VB.Menu mnuPal 
            Caption         =   "RSE Palette 5"
            Index           =   6
         End
         Begin VB.Menu mnuPal 
            Caption         =   "FRLG Palette 0"
            Index           =   7
         End
         Begin VB.Menu mnuPal 
            Caption         =   "FRLG Palette 1"
            Index           =   8
         End
         Begin VB.Menu mnuPal 
            Caption         =   "FRLG Palette 2"
            Index           =   9
         End
         Begin VB.Menu mnuPal 
            Caption         =   "FRLG Palette 3"
            Index           =   10
         End
         Begin VB.Menu mnuPal 
            Caption         =   "FRLG Palette 4"
            Index           =   11
         End
         Begin VB.Menu mnuPal 
            Caption         =   "FRLG Palette 5"
            Index           =   12
         End
         Begin VB.Menu mnuPal 
            Caption         =   "RGB Standard"
            Index           =   13
         End
      End
   End
   Begin VB.Menu mnuDoors 
      Caption         =   "&Doors"
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepoint 
         Caption         =   "Edit Door Amount"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuReadme 
         Caption         =   "Readme"
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
Public DoorTable As Long
Public sFilePath As String
Public iDoorAmount As Integer

Private iFileNum As Integer
Private PaletteData As Long
Private arrCustomPal() As Byte
Private sHeader As String * 4

Private Type Doors
    TileNumber As Integer
    DoorType As Integer
    GraphicsPointer As Long
    PaletteDataPointer As Long
End Type

Private Type PaletteData
    Data1 As Byte
    Data2 As Byte
    Data3 As Byte
    Data4 As Byte
    Data5 As Byte
    Data6 As Byte
    Data7 As Byte
    Data8 As Byte
End Type

Private DoorData As Doors
Private LocalPalettes As PaletteData

Private Sub DrawDoor(ByVal hDC As Long, ByRef arrTiles() As Byte, sPaletteIndex As String, Optional CustPal As Boolean)
Dim i As Integer
Dim arrPalData() As Byte
Dim SomeCounter As Integer
Dim arrPal(0 To 256) As Integer
Dim PCPalette(0 To 16, 0 To 16) As Long
    
    If sPaletteIndex = 0 Then sPaletteIndex = 1 ' Loading Palette Data
    arrPalData = LoadResData(Val(sPaletteIndex), "PAL")
    
    SomeCounter = 0
    For i = 0 To 15 ' 16 palettes ^^
        arrPal(i) = CInt(arrPalData(SomeCounter + 1) * &H100 + arrPalData(SomeCounter))
        SomeCounter = SomeCounter + 2
    Next i
    UnPackPalette arrPal, PCPalette
    
    If mnuEditorMode.Checked = True Then
        For i = 0 To 15
            TileEdit.Colors(i) = UnPackPaletteRGB(arrPal(i))
            optColors(i).BackColor = UnPackPaletteRGB(arrPal(i))
        Next i
    End If
    
    picDoor.Cls
    For i = 0 To 7
        DrawTile8 hDC, i, arrTiles, PCPalette
    Next i
    
    Erase arrPalData
End Sub

Private Sub cmbPal_Click()
If shpCursor.Left \ 32 = 0 Then
    LocalPalettes.Data1 = cmbPal.ListIndex
ElseIf shpCursor.Left \ 32 = 1 Then
    LocalPalettes.Data2 = cmbPal.ListIndex
ElseIf shpCursor.Left \ 32 = 2 Then
    LocalPalettes.Data3 = cmbPal.ListIndex
ElseIf shpCursor.Left \ 32 = 3 Then
    LocalPalettes.Data4 = cmbPal.ListIndex
ElseIf shpCursor.Left \ 32 = 4 Then
    LocalPalettes.Data5 = cmbPal.ListIndex
ElseIf shpCursor.Left \ 32 = 5 Then
    LocalPalettes.Data6 = cmbPal.ListIndex
ElseIf shpCursor.Left \ 32 = 6 Then
    LocalPalettes.Data7 = cmbPal.ListIndex
ElseIf shpCursor.Left \ 32 = 7 Then
    LocalPalettes.Data8 = cmbPal.ListIndex
End If
End Sub

Private Sub cmdNext_Click()
Dim MaxFrame As Byte
MaxFrame = 23
If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then MaxFrame = 11

If txtFrame.Text = MaxFrame Then Exit Sub
txtFrame.Text = Val(txtFrame.Text) + 1
End Sub

Private Sub cmdPrevious_Click()
If txtFrame.Text = "0" Then Exit Sub
txtFrame.Text = Val(txtFrame.Text) - 1
End Sub

Private Sub cmdSave_Click()
    mnuSave_Click
End Sub

Private Sub Form_Load()
Dim Index As Integer
    SetIcon Me.hWnd, "AAA", True ' Load Icon
    
    ' Loading Palette
    Index = CInt(GetFromINI("Settings", "Palette", 0, App.Path & "\Settings.ini"))
    If Index = 0 Then Index = 1
    mnuPal(Index).Checked = True
    
    ' Loading EditorMode
    If GetFromINI("Settings", "EditorMode", 0, App.Path & "\Settings.ini") = 1 Then mnuEditorMode.Checked = True
    If mnuEditorMode.Checked = True Then
        frmMain.Width = 10410
        frmMain.Caption = "Door Manager - Editor Mode On"
    Else
        frmMain.Width = 7050
        frmMain.Caption = "Door Manager - Editor Mode Off"
    End If
    
    ' Loading Gridlines
    If GetFromINI("Settings", "Gridlines", 0, App.Path & "\Settings.ini") = 1 Then
        mnuGridlines.Checked = True
        TileEdit.ShowGrid = True
    End If
    
    imgFlag.Picture = LoadResPicture("NULL", 0)
    Localize Me
End Sub

Public Function LoadDoors(iFile As Integer)
Dim i As Integer
Dim bTemp As Byte
    i = 0
    lstDoors.Clear
    iDoorAmount = 0
    Do
        Get #iFile, DoorTable + 1 + 11 + i, bTemp
        If bTemp <> &H8 Then Exit Do
        iDoorAmount = iDoorAmount + 1
        i = i + 12
        lstDoors.AddItem "Door" & " #" & Right("00" & iDoorAmount, 2)
    Loop
End Function

Private Sub lstDoors_Click()
Dim arrTemp(32768) As Byte
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, DoorTable + 1 + (lstDoors.ListIndex * 12), DoorData
        
        txtTileNumber.Text = Hex(DoorData.TileNumber)
        txtDoorType.Text = Hex(DoorData.DoorType)
        
        txtGraphics.Text = Hex(DoorData.GraphicsPointer - &H8000000)
        txtPalette.Text = Hex(DoorData.PaletteDataPointer - &H8000000)
        
        ' Drawing the door
        If DoorData.GraphicsPointer = &H8000000 Then GoTo ContinueReading ' For newly made doors
        
        Get #iFileNum, (DoorData.GraphicsPointer - &H8000000) + 1, arrTemp
        DrawDoor picTemp.hDC, arrTemp, GetFromINI("Settings", "Palette", 0, App.Path & "\Settings.ini")
        StretchBlt picDoor.hDC, 0, 0, 256, 32, picTemp.hDC, 0, 0, 64, 8, vbSrcCopy
        picTemp.Cls
        
ContinueReading:
        ' Loading Palette Data
        If DoorData.PaletteDataPointer = &H8000000 Then GoTo ContinueReadingNext ' For newly made doors
        
        PaletteData = (DoorData.PaletteDataPointer - &H8000000)
        Get #iFileNum, PaletteData + 1, LocalPalettes
        picDoor_MouseDown vbLeftButton, 0, 0, 0
    
ContinueReadingNext:
        ' Loads data to the editor only if edit mode is on
        If mnuEditorMode.Checked = True Then
            txtFrame.Text = 0
            txtFrame_Change
        End If
    
    Erase arrTemp
    Close #iFileNum
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuEditorMode_Click()
    If mnuEditorMode.Checked = True Then
        mnuEditorMode.Checked = False
        AddToINI "Settings", "EditorMode", 0, App.Path & "\Settings.ini"
        frmMain.Width = 7050
        frmMain.Caption = "Door Manager - Editor Mode Off"
    Else
        mnuEditorMode.Checked = True
        AddToINI "Settings", "EditorMode", 1, App.Path & "\Settings.ini"
        frmMain.Width = 10410
        frmMain.Caption = "Door Manager - Editor Mode On"
    End If
    ' Only load the data when a ROM's loaded
    If LenB(sFilePath) <> 0 Then
        lstDoors_Click
        txtFrame.Text = 0
        txtFrame_Change
    End If
End Sub

Private Sub mnuExport_Click()
Dim arrTemp() As Byte
Dim sExport As String
Dim cdgSave As clsCommonDialog
    If DoorData.GraphicsPointer - &H8000000 = 0 Then Exit Sub
    
    Set cdgSave = New clsCommonDialog
    sExport = cdgSave.ShowSave(Me.hWnd, "Save As", "Door " & lstDoors.ListIndex + 1, , "Binary Files (*.bin)|*.bin")
    If LenB(sExport) = 0 Then GoTo EndMe
    
    ReadByteArray sFilePath, arrTemp, DoorData.GraphicsPointer - &H8000000, 768
    WriteByteArray sExport, arrTemp, 0
    
    Erase arrTemp
EndMe:
    Set cdgSave = Nothing
End Sub

Private Sub mnuGridlines_Click()
    If mnuGridlines.Checked = True Then
        mnuGridlines.Checked = False
        TileEdit.ShowGrid = False
        AddToINI "Settings", "Gridlines", 0, App.Path & "\Settings.ini"
    Else
        TileEdit.ShowGrid = True
        mnuGridlines.Checked = True
        AddToINI "Settings", "Gridlines", 1, App.Path & "\Settings.ini"
    End If
    ' Only load the data when a ROM's loaded
    If LenB(sFilePath) <> 0 Then
        txtFrame.Text = 0
        txtFrame_Change
    End If
End Sub

Private Sub mnuImport_Click()
Dim arrTemp() As Byte
Dim sImport As String
Dim cdgOpen As clsCommonDialog
    If DoorData.GraphicsPointer - &H8000000 = 0 Then Exit Sub

    Set cdgOpen = New clsCommonDialog
    sImport = cdgOpen.ShowOpen(Me.hWnd, "Open", , "Binary Files (*.bin)|*.bin")
    If LenB(sImport) = 0 Then GoTo EndMe
    
    ReadByteArray sImport, arrTemp, 0, 768
    WriteByteArray sFilePath, arrTemp, DoorData.GraphicsPointer - &H8000000
    
    Erase arrTemp
    lstDoors_Click
EndMe:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuOpen_Click()
Dim cnt As Control
Dim sResult As String
Dim cdgOpen As clsCommonDialog
    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, "Open ROM...", , "Gameboy Advance ROMs (*.gba,*agb,*.bin)|*.gba;*.agb;*.bin")
    If LenB(sResult) = 0 Then GoTo EndMe
    sFilePath = sResult
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
    Get #iFileNum, &HAC + 1, sHeader
    
        Select Case sHeader
            Case "AXVE"
                lblName.Caption = "Pokémon Ruby"
                Get #iFileNum, &H586B0 + 1, DoorTable
            
            Case "AXPE"
                lblName.Caption = "Pokémon Sapphire"
                Get #iFileNum, &H586B4 + 1, DoorTable
                
            Case "BPRE"
                lblName.Caption = "Pokémon Fire Red"
                Get #iFileNum, &H5B298 + 1, DoorTable
                
            Case "BPGE"
                lblName.Caption = "Pokémon Fire Red"
                Get #iFileNum, &H5B298 + 1, DoorTable
                
            Case "BPEE"
                lblName.Caption = "Pokémon Emerald"
                Get #iFileNum, &H8A850 + 1, DoorTable
                
            Case Else
                MsgBox "Error - Your ROM is unsupported", vbExclamation
                ResetForm
                Close #iFileNum
                sHeader = vbNullString
                sFilePath = vbNullString
                GoTo EndMe
                
        End Select
    
        DoorTable = DoorTable - &H8000000
        LoadDoors iFileNum
        
        If LOF(iFileNum) > 16777216 Then ' Support for 32MB ROMs
            txtGraphics.MaxLength = 7
            txtPalette.MaxLength = 7
        Else
            txtGraphics.MaxLength = 6
            txtPalette.MaxLength = 6
        End If
        
        For Each cnt In Me.Controls ' Enabling Stuff
            On Error Resume Next
            If cnt.Enabled = False Then cnt.Enabled = True
        Next
        
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
        
        TileEdit.Filename = sFilePath
        TileEdit.PenColor = 0
        TileEdit.Visible = True
        
        lstDoors.ListIndex = 0
        shpCursor.Visible = True
        txtDesign.Enabled = False
        txtDesign2.Enabled = False
        lblHeader.Caption = sHeader
    
    Close #iFileNum
EndMe:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuPal_Click(Index As Integer)
Dim iPrevious As Integer
    ' Previous value - Uncheck the old menu
    iPrevious = CInt(GetFromINI("Settings", "Palette", 0, App.Path & "\Settings.ini"))
    If iPrevious = 0 Then iPrevious = 1
    mnuPal(iPrevious).Checked = False
    
    ' New value - write to Ini and check the menu
    AddToINI "Settings", "Palette", CStr(Index), App.Path & "\Settings.ini"
    mnuPal(Index).Checked = True
    
    If LenB(sFilePath) <> 0 Then lstDoors_Click ' Reload doors in new palette only if ROM is loaded
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuReadme_Click()
Dim arrTemp() As Byte
    arrTemp = LoadResData("README", 100)
    WriteByteArray App.Path & "\Readme.txt", arrTemp, 0
    Shell "notepad.exe " & App.Path & "\Readme.txt", vbNormalFocus
    Kill App.Path & "\Readme.txt"
    Erase arrTemp
End Sub

Private Sub mnuRepoint_Click()
    frmRepoint.Show vbModal, Me
End Sub

Private Sub mnuSave_Click()
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        DoorData.TileNumber = CInt("&H" & txtTileNumber.Text)
        DoorData.DoorType = CInt("&H" & txtDoorType.Text)
        DoorData.GraphicsPointer = CLng("&H" & txtGraphics.Text) + &H8000000
        DoorData.PaletteDataPointer = CLng("&H" & txtPalette.Text) + &H8000000
        
        If mnuEditorMode.Checked = True Then ' Saves the data only if edit mode is on
            TileEdit.SaveTileData
        End If
        
        Put #iFileNum, DoorTable + 1 + (lstDoors.ListIndex * 12), DoorData
        Put #iFileNum, PaletteData + 1, LocalPalettes
    Close #iFileNum
End Sub

Private Sub ResetForm()
    mnuSave.Enabled = False ' Disabling controls
    mnuExport.Enabled = False
    mnuImport.Enabled = False
    mnuRepoint.Enabled = False
    
    picDoor.Cls
    lstDoors.Clear
    lblName.Caption = "???"
    cmdSave.Enabled = False
    cmdNext.Enabled = False
    lstDoors.Enabled = False
    TileEdit.Visible = False
    shpCursor.Visible = False
    cmdPrevious.Enabled = False
    lblHeader.Caption = lblName.Caption
    imgFlag.Picture = LoadResPicture("NULL", 0)
End Sub

Private Sub optColors_Click(Index As Integer)
    TileEdit.PenColor = Index
End Sub

Private Sub picDoor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    shpCursor.Move Fix(x \ 32) * 32, Fix(y \ 32) * 32
    If Button = vbRightButton Then PopupMenu mnuPalettes, vbPopupMenuRightButton
    
    If shpCursor.Left \ 32 = 0 Then
        cmbPal.ListIndex = LocalPalettes.Data1
    ElseIf shpCursor.Left \ 32 = 1 Then
        cmbPal.ListIndex = LocalPalettes.Data2
    ElseIf shpCursor.Left \ 32 = 2 Then
        cmbPal.ListIndex = LocalPalettes.Data3
    ElseIf shpCursor.Left \ 32 = 3 Then
        cmbPal.ListIndex = LocalPalettes.Data4
    ElseIf shpCursor.Left \ 32 = 4 Then
        cmbPal.ListIndex = LocalPalettes.Data5
    ElseIf shpCursor.Left \ 32 = 5 Then
        cmbPal.ListIndex = LocalPalettes.Data6
    ElseIf shpCursor.Left \ 32 = 6 Then
        cmbPal.ListIndex = LocalPalettes.Data7
    ElseIf shpCursor.Left \ 32 = 7 Then
        cmbPal.ListIndex = LocalPalettes.Data8
    End If
End Sub

Private Sub txtFrame_Change()
    TileEdit.ROMAddress = DoorData.GraphicsPointer - &H8000000 + (txtFrame.Text * &H20)
    TileEdit.LoadTileData
End Sub
