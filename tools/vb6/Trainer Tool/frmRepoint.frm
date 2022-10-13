VERSION 5.00
Begin VB.Form frmRepoint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit New Trainer Amount"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepoint.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRepoint 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame fraOffsets 
      Caption         =   "Offsets"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   9
         Top             =   240
         Width           =   2535
         Begin VB.TextBox txtOldOffset 
            ForeColor       =   &H80000011&
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNewOffset 
            Height          =   285
            Left            =   1680
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
            Left            =   1440
            TabIndex        =   10
            Text            =   "0x"
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblOldOffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblNewOffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            Height          =   195
            Left            =   1440
            TabIndex        =   14
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
      Top             =   120
      Width           =   2775
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   4
         Top             =   240
         Width           =   2535
         Begin VB.TextBox txtOldAmount 
            ForeColor       =   &H80000011&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtNewAmount 
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblOldAmount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblNewAmount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            Height          =   195
            Left            =   1440
            TabIndex        =   7
            Top             =   0
            Width           =   315
         End
      End
   End
   Begin VB.Label lblHex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   195
      Left            =   3600
      TabIndex        =   20
      Top             =   600
      Width           =   225
   End
   Begin VB.Label lblDec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   195
      Left            =   3600
      TabIndex        =   19
      Top             =   360
      Width           =   225
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex:"
      Height          =   195
      Left            =   3120
      TabIndex        =   18
      Top             =   600
      Width           =   345
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dec:"
      Height          =   195
      Left            =   3120
      TabIndex        =   17
      Top             =   360
      Width           =   330
   End
   Begin VB.Label lblNeededBytes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Needed Bytes"
      Height          =   195
      Left            =   3120
      TabIndex        =   16
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmRepoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iFileNum As Integer
Private sHeader As String * 4

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRepoint_Click()
Dim i As Integer
Dim Temp As Long
Dim Temp2 As Long
Dim arrTemp() As Byte
Dim arrData(7) As Byte
Dim TrainerIndex As Byte
    If txtNewOffset.Text = vbNullString Then Exit Sub
    
    iFileNum = FreeFile
    Open frmMain.sFilePath For Binary As #iFileNum
        ' Step 1 - Repointing trainer image table
        arrData(3) = &H8
        arrData(5) = &H8
        
        Temp = frmMain.iTrainerAmount * 8
        ReadByteArray frmMain.sFilePath, arrTemp, frmMain.TrainerPics, Temp
        WriteByteArray frmMain.sFilePath, arrTemp, CLng("&H" & txtNewOffset.Text)
        
        If Val(txtNewAmount.Text) < Val(txtOldAmount.Text) Then GoTo RemoveTrainerPics
        If frmMain.TrainerPics = CLng("&H" & txtNewOffset.Text) Then GoTo ContinueReading
        PutFreeSpace frmMain.sFilePath, frmMain.TrainerPics, Temp
        
ContinueReading:
        ' Inserting new trainer images
        TrainerIndex = CByte(txtOldAmount.Text)
        If (txtNewAmount - txtOldAmount) = 0 Then GoTo TrainerPalettes
        For i = 1 To (txtNewAmount - txtOldAmount)
            arrData(6) = TrainerIndex
            Put #iFileNum, CLng("&H" & txtNewOffset.Text) + 1 + Temp, arrData
            TrainerIndex = TrainerIndex + 1
            Temp = Temp + 8
        Next i
        GoTo TrainerPalettes
        
RemoveTrainerPics:
        ' Removing extra trainer images
        PutFreeSpace frmMain.sFilePath, CLng("&H" & txtOldOffset.Text), (txtOldAmount * 8)
        WriteByteArray frmMain.sFilePath, arrTemp, CLng("&H" & txtNewOffset.Text)
        PutFreeSpace frmMain.sFilePath, CLng("&H" & txtNewOffset.Text) + txtNewAmount * 8, (txtOldAmount - txtNewAmount) * 8
        Temp = Temp - (txtOldAmount - txtNewAmount) * 8
        
TrainerPalettes:
        ' Step 2 - Repointing trainer palette table
        Erase arrTemp, arrData
        arrData(3) = &H8
        
        Temp2 = frmMain.iTrainerAmount * 8
        ReadByteArray frmMain.sFilePath, arrTemp, frmMain.TrainerPals, Temp2
        WriteByteArray frmMain.sFilePath, arrTemp, CLng("&H" & txtNewOffset.Text) + Temp
        
        If Val(txtNewAmount.Text) < Val(txtOldAmount.Text) Then GoTo RemoveTrainerPals
        If frmMain.TrainerPics = CLng("&H" & txtNewOffset.Text) Then GoTo ContinueExecution
        PutFreeSpace frmMain.sFilePath, frmMain.TrainerPals, Temp2

ContinueExecution:
        ' Inserting new trainer palettes
        TrainerIndex = CByte(txtOldAmount.Text)
        If (txtNewAmount - txtOldAmount) = 0 Then GoTo EndMe
        For i = 1 To (txtNewAmount - txtOldAmount)
            arrData(4) = TrainerIndex
            Put #iFileNum, CLng("&H" & txtNewOffset.Text) + 1 + Temp + Temp2, arrData
            TrainerIndex = TrainerIndex + 1
            Temp2 = Temp2 + 8
        Next i
        GoTo EndMe
        
RemoveTrainerPals:
        ' Removing extra trainer palettes
        PutFreeSpace frmMain.sFilePath, CLng("&H" & txtOldOffset.Text) + Temp + (txtOldAmount - txtNewAmount) * 8, (txtOldAmount * 8)
        WriteByteArray frmMain.sFilePath, arrTemp, CLng("&H" & txtNewOffset.Text) + Temp
        PutFreeSpace frmMain.sFilePath, CLng("&H" & txtNewOffset.Text) + Temp + txtNewAmount * 8, (txtOldAmount - txtNewAmount) * 8

EndMe:
        ' Step 3 - Repointing trainer image and palette calling offsets
        Select Case sHeader
            Case "AXVE"
                ' Trainer Images - 10 offsets
                Put #iFileNum, &H31ADC + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H31B9C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H34DA8 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H34F6C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H3988C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H85A48 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H85A8C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H91AE4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &HF6E88 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H143854 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                
                ' Trainer Palettes - 9 offsets
                Put #iFileNum, &H31AF0 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
                Put #iFileNum, &H31B98 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
                Put #iFileNum, &H34DA4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
                Put #iFileNum, &H34F68 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
                Put #iFileNum, &H39888 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
                Put #iFileNum, &H85A44 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
                Put #iFileNum, &H85A90 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
                Put #iFileNum, &HF6E98 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
                Put #iFileNum, &H143860 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000 + Temp
            Case "BPRE"
                ' Trainer Images - 9 offsets
                Put #iFileNum, &H3473C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H347A4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H37E8C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H38060 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H3CAE8 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H838E4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H83928 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H10BC3C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H158528 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                
                'Trainer Palettes - 11 offsets
            Case "BPEE"
        End Select
     
        Erase arrTemp, arrData
        frmMain.TrainerPics = CLng("&H" & txtNewOffset.Text)
        frmMain.TrainerPals = CLng("&H" & txtNewOffset.Text) + Temp
        frmMain.LoadTrainers iFileNum
    Close #iFileNum
    Unload Me
End Sub

Private Sub Form_Load()
    sHeader = frmMain.lblHeader
    txtOldAmount.Text = frmMain.iTrainerAmount
    txtNewAmount.Text = txtOldAmount.Text
    txtOldOffset.Text = Hex(frmMain.TrainerPics)
End Sub

Private Sub txtNewAmount_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub txtNewOffset_Change()
    If txtNewAmount.Text = 0 Then Exit Sub
    cmdRepoint.Enabled = True
    lblDec.Caption = 2 * (8 * txtNewAmount)
    lblHex.Caption = "0x" & Hex(2 * (8 * txtNewAmount))
End Sub
