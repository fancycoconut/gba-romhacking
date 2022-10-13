VERSION 5.00
Begin VB.Form frmEditAmount 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Trade Amount"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditAmount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "9"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Tag             =   "30"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRepoint 
      Caption         =   "Repoint"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Tag             =   "29"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame fraOffset 
      Caption         =   "Offset"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Tag             =   "26"
      Top             =   1200
      Width           =   2655
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   2415
         TabIndex        =   9
         Top             =   240
         Width           =   2415
         Begin VB.TextBox txtOldOffset 
            ForeColor       =   &H80000011&
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNewOffset 
            Height          =   285
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txt1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Text            =   "0x"
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txt2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            Text            =   "0x"
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblOld 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Tag             =   "27"
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblNew 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   14
            Tag             =   "28"
            Top             =   0
            Width           =   315
         End
      End
   End
   Begin VB.Frame fraAmount 
      Caption         =   "Amount"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Tag             =   "25"
      Top             =   120
      Width           =   2655
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   2415
         TabIndex        =   16
         Top             =   240
         Width           =   2415
         Begin VB.TextBox txtOld 
            ForeColor       =   &H80000011&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtNewAmount 
            Height          =   285
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   18
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblOld 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Tag             =   "27"
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblNew 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            Height          =   195
            Index           =   0
            Left            =   1320
            TabIndex        =   17
            Tag             =   "28"
            Top             =   0
            Width           =   315
         End
      End
   End
   Begin VB.Label lblDec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   195
      Left            =   3480
      TabIndex        =   8
      Top             =   720
      Width           =   225
   End
   Begin VB.Label lblHex 
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   480
      Width           =   225
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex:"
      Height          =   195
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   345
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dec:"
      Height          =   195
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   330
   End
   Begin VB.Label lblNeeded 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Needed Bytes"
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Tag             =   "24"
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "frmEditAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRepoint_Click()
Dim i As Integer
Dim arrTemp() As Byte
Dim TempVariable As Long
Dim arrTradeData() As Byte
' Former trade data array disabled - they are now stored in the resource file
'Dim TradeData(59) As Byte
'TradeData(0) = &H0
'TradeData(1) = &H0
'TradeData(2) = &H0
'TradeData(3) = &H0
'TradeData(4) = &H0
'TradeData(5) = &H0
'TradeData(6) = &H0
'TradeData(7) = &H0
'TradeData(8) = &H0
'TradeData(9) = &H0 '0-9 = Pokemon Nickname
'TradeData(10) = &H0
'TradeData(11) = &H0 '11-12 = Filler Word
'TradeData(12) = &H0
'TradeData(13) = &H0 '12-13 = Pokemon Offered
'TradeData(14) = &H5
'TradeData(15) = &H5 '14-15 = 0x0505 {Unknown Word}
'TradeData(16) = &H4
'TradeData(17) = &H4 '16-17 = 0x0404 {Unknown Word}
'TradeData(18) = &H4
'TradeData(19) = &H4 '18-19 = 0x0404 {Unknown Word}
'TradeData(20) = &H1
'TradeData(21) = &H0 '20-21 = Ability
'TradeData(22) = &H0
'TradeData(23) = &H0 '22-23 = Filler Word
'TradeData(24) = &H0
'TradeData(25) = &H0 '24-25 = OTID {Reversed}
'TradeData(26) = &H0
'TradeData(27) = &H0 '26-27 = Filler Word
'TradeData(28) = &H5
'TradeData(29) = &H5 '28-29 = 0x0505 {Unknown Word}
'TradeData(30) = &H5
'TradeData(31) = &H5 '30-31 = 0x0505 {Unknown Word}
'TradeData(32) = &H1E
'TradeData(33) = &H0 '32-33 = 0x1E00 {Unknown Word}
'TradeData(34) = &H0
'TradeData(35) = &H0 '34-35 = Filler Word
'TradeData(36) = &H40
'TradeData(37) = &H9C
'TradeData(38) = &H0
'TradeData(39) = &H0 '36-39 = Personality Byte
'TradeData(40) = &H0
'TradeData(41) = &H0 '40-41 = Held Item
'TradeData(42) = &HFF '42 = Byte 0xFF
'TradeData(43) = &H0
'TradeData(44) = &H0
'TradeData(45) = &H0
'TradeData(46) = &H0
'TradeData(47) = &H0
'TradeData(48) = &H0
'TradeData(49) = &H0 '43-49 = Trainer Name
'TradeData(50) = &H0
'TradeData(51) = &H0 '50-51 = Filler Word
'TradeData(52) = &H0
'TradeData(53) = &H0 '52-53 = Filler Word
'TradeData(54) = &H0
'TradeData(55) = &HA '54-55 = 0x000A*
'TradeData(56) = &H0
'TradeData(57) = &H0 '56-57 = Pokemon Wanted
'TradeData(58) = &H0
'TradeData(59) = &H0 '58-59 = Filler
'* 0xA is used to determine Trade Amount in loop

    arrTradeData = LoadResData("TRADEDATA", 100)

    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
    
        If txtNewOffset.Text = vbNullString Then Exit Sub
    
        'Step 1 Replacing calling offsets to trade table
        Select Case sHeader
            Case "AXVE", "AXPE"
                Put #iFileNum, &H4C284 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H4D8D4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H4D930 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H4DAA4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
            Case "BPRE", "BPGE"
                Put #iFileNum, &H50EFC + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H53AD4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H53B30 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H53CA4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
            Case "BPEE"
                Put #iFileNum, &H7BBB0 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H7E774 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H7E7D0 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H7E944 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
            Case "BPEF"
                Put #iFileNum, &H7BBAC + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H7E770 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H7E7CC + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H7E940 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
        End Select
    
        'Step 2 Repointing existing data
        TempVariable = iTradeAmount * 60
        ReadByteArray sFilePath, arrTemp, TradeDataOffset, TempVariable
        WriteByteArray sFilePath, arrTemp, CLng("&H" & txtNewOffset.Text)
        
        If Val(txtNewAmount.Text) < Val(txtOld.Text) Then GoTo RemoveTradeData
        If TradeDataOffset = CLng("&H" & txtNewOffset.Text) Then GoTo ContinueReading
        PutFreeSpace sFilePath, TradeDataOffset, TempVariable

ContinueReading:
        'Step 3A Inserting new trade data
        For i = 1 To (txtNewAmount - txtOld)
            Put #iFileNum, CLng("&H" & txtNewOffset.Text) + 1 + TempVariable, arrTradeData
            TempVariable = TempVariable + 60
        Next i
        GoTo EndMe
    
RemoveTradeData:
        'Step 3B Removing extra trade data
        PutFreeSpace sFilePath, CLng("&H" & txtOldOffset.Text), (txtOld * 60)
        WriteByteArray sFilePath, arrTemp, CLng("&H" & txtNewOffset.Text)
        PutFreeSpace sFilePath, CLng("&H" & txtNewOffset.Text) + txtNewAmount * 60, (txtOld - txtNewAmount) * 60
    
EndMe:
        frmMain.GetTradeAmount
        Erase arrTemp
        Erase arrTradeData
    Close #iFileNum
    Unload Me
End Sub

Private Sub Form_Load()
    Localize Me
    
    If EnlargedROM = 1 Then
        txtOldOffset.MaxLength = 7
        txtNewOffset.MaxLength = 7
    End If
    
    txtOld.Text = iTradeAmount
    txtNewAmount.Text = iTradeAmount
    txtOldOffset.Text = Hex(TradeDataOffset)
End Sub

Private Sub txtNewAmount_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub txtNewOffset_Change()
    If txtNewAmount = 0 Then Exit Sub
    cmdRepoint.Enabled = True
    lblDec.Caption = txtNewAmount.Text * 60
    lblHex.Caption = Hex(txtNewAmount.Text * 60)
End Sub
