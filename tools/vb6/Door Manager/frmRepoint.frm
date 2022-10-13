VERSION 5.00
Begin VB.Form frmRepoint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Door Amount"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4080
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame fraAmount 
      Caption         =   "Amount"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2535
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   7
         Top             =   240
         Width           =   2295
         Begin VB.TextBox txtNewAmount 
            Height          =   285
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtOldAmount 
            ForeColor       =   &H80000011&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblNewAmount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            Height          =   195
            Left            =   1200
            TabIndex        =   10
            Top             =   0
            Width           =   315
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
      End
   End
   Begin VB.Frame fraOffset 
      Caption         =   "Offsets"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   1
         Top             =   240
         Width           =   2295
         Begin VB.TextBox txt2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   20
            Text            =   "0x"
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txt1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Text            =   "0x"
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtNewOffset 
            Height          =   285
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtOldOffset 
            ForeColor       =   &H80000011&
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblNewOffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            Height          =   195
            Left            =   1200
            TabIndex        =   4
            Top             =   0
            Width           =   315
         End
         Begin VB.Label lblOldOffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   240
         End
      End
   End
   Begin VB.Label lblHex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   195
      Left            =   3360
      TabIndex        =   16
      Top             =   720
      Width           =   225
   End
   Begin VB.Label lblDec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   195
      Left            =   3360
      TabIndex        =   15
      Top             =   480
      Width           =   225
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex:"
      Height          =   195
      Left            =   2880
      TabIndex        =   14
      Top             =   720
      Width           =   345
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dec:"
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   480
      Width           =   330
   End
   Begin VB.Label lblNeededBytes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Needed Bytes"
      Height          =   195
      Left            =   2880
      TabIndex        =   12
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "frmRepoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sHeader As String * 4
Private iFileNum As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
Dim Temp As Long
Dim arrTemp() As Byte
Dim arrData(11) As Byte ' 12 bytes
    If txtNewOffset.Text = vbNullString Then Exit Sub
    
    arrData(7) = &H8
    arrData(11) = &H8
    
    iFileNum = FreeFile
    Open frmMain.sFilePath For Binary As #iFileNum
        
        ' Step 1 - Changing the calling offsets to the table
        Select Case sHeader
            Case "AXVE"
                Put #iFileNum, &H586B0 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H586DC + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H58708 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H58734 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H5876C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H587A8 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
            Case "AXPE"
            
            Case "BPRE"
                Put #iFileNum, &H5B298 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H5B2CC + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H5B300 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H5B340 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H5B37C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
            Case "BPGE"
            
            Case "BPEE"
                Put #iFileNum, &H8A850 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H8A87C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H8A8A8 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H8A8D4 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H8A90C + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
                Put #iFileNum, &H8A950 + 1, CLng("&H" & txtNewOffset.Text) + &H8000000
        End Select
        
        ' Step 2 - Repointing Existing Data
        Temp = frmMain.iDoorAmount * 12
        ReadByteArray frmMain.sFilePath, arrTemp, frmMain.DoorTable, Temp
        WriteByteArray frmMain.sFilePath, arrTemp, CLng("&H" & txtNewOffset.Text)
        
        If Val(txtNewAmount.Text) < Val(txtOldAmount.Text) Then GoTo RemoveDoorData
        If frmMain.DoorTable = CLng("&H" & txtNewOffset.Text) Then GoTo ContinueReading
        PutFreeSpace frmMain.sFilePath, frmMain.DoorTable, Temp

ContinueReading:
        ' Step 3A - Inserting new door data
        For i = 1 To (txtNewAmount - txtOldAmount)
            Put #iFileNum, CLng("&H" & txtNewOffset.Text) + 1 + Temp, arrData
            Temp = Temp + 12
        Next i
        GoTo EndMe
        
RemoveDoorData:
        ' Step 3B - Removing extra door data
        PutFreeSpace frmMain.sFilePath, CLng("&H" & txtOldOffset.Text), (txtOldAmount * 12)
        WriteByteArray frmMain.sFilePath, arrTemp, CLng("&H" & txtNewOffset.Text)
        PutFreeSpace frmMain.sFilePath, CLng("&H" & txtNewOffset.Text) + txtNewAmount * 12, (txtOldAmount - txtNewAmount) * 12
EndMe:
    frmMain.DoorTable = CLng("&H" & txtNewOffset.Text)
    frmMain.LoadDoors iFileNum
    Erase arrTemp, arrData
    Close #iFileNum
    Unload Me
End Sub

Private Sub Form_Load()
    Localize Me
    
    sHeader = frmMain.lblHeader.Caption
    txtNewOffset.MaxLength = frmMain.txtGraphics.MaxLength
    txtOldOffset.Text = Hex(frmMain.DoorTable)
    txtOldAmount.Text = frmMain.iDoorAmount
    txtNewAmount.Text = txtOldAmount.Text
End Sub

Private Sub txtNewAmount_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub txtNewOffset_Change()
    If txtNewAmount.Text = 0 Then Exit Sub
    cmdOK.Enabled = True
    lblDec.Caption = txtNewAmount.Text * 12
    lblHex.Caption = "0x" & Hex(txtNewAmount.Text * 12)
End Sub
