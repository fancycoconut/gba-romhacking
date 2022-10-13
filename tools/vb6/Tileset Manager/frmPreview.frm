VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preview"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "40"
   Begin VB.PictureBox picTileset 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   120
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   120
      Width           =   3000
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   0
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer
Dim arrTemp() As Byte
Dim SomeCounter As Long
Dim arrTileset() As Byte
Dim arrPal(0 To 256) As Integer
Dim PCPalette(0 To 16, 0 To 16) As Long

Dim arrRGBPal(31) As Byte
'arrRGBPal(0) = &H0
'arrRGBPal(1) = &H0
arrRGBPal(2) = &H4C
arrRGBPal(3) = &H33
arrRGBPal(4) = &HC7
arrRGBPal(5) = &H7C
arrRGBPal(6) = &HFC
arrRGBPal(7) = &H7F
arrRGBPal(8) = &H6
arrRGBPal(9) = &H44
arrRGBPal(10) = &HD8
arrRGBPal(11) = &H1
arrRGBPal(12) = &H40
arrRGBPal(13) = &H7F
arrRGBPal(14) = &HBE
arrRGBPal(15) = &H5B
arrRGBPal(16) = &H12
'arrRGBPal(17) = &H0
arrRGBPal(18) = &HEA
arrRGBPal(19) = &H3
arrRGBPal(20) = &HDF
arrRGBPal(21) = &H2
arrRGBPal(22) = &H58
arrRGBPal(23) = &H50
'arrRGBPal(24) = &H0
'arrRGBPal(25) = &H0
arrRGBPal(26) = &H4B
arrRGBPal(27) = &H7A
arrRGBPal(28) = &H17
arrRGBPal(29) = &H50
arrRGBPal(30) = &H90
arrRGBPal(31) = &H3

    GetFileData sTilesetPath, arrTileset
    
    If frmMain.chCompression.Value = vbChecked Then
        LZ77UnComp arrTileset, arrTemp
    Else
        arrTemp = arrTileset
    End If
    
    SomeCounter = 0
    For i = 0 To 15 'PalSize / 2
        arrPal(i) = CInt(arrRGBPal(SomeCounter + 1) * &H100 + arrRGBPal(SomeCounter))
        SomeCounter = SomeCounter + 2
    Next i
    
    UnPackPalette arrPal, PCPalette
    
    picTemp.Cls
    For i = 0 To 255
        On Error Resume Next
        DrawTile8 picTemp.hdc, i, arrTemp, PCPalette
    Next i
    
    StretchBlt picTileset.hdc, 0, 0, 200, 200, picTemp.hdc, 0, 0, 64, 144, vbSrcCopy
    picTileset.Refresh
    picTemp.Cls
    
    Erase arrRGBPal, arrTileset, arrTemp
End Sub
