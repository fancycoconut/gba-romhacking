VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASM Manager"
   ClientHeight    =   4770
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Trebuchet MS"
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
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frCopy 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   5415
      Begin VB.Label lblMX 
         AutoSize        =   -1  'True
         Caption         =   "Copyright © 2008 ZodiacDaGreat"
         Enabled         =   0   'False
         Height          =   240
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   2490
      End
   End
   Begin VB.Frame frStuff 
      Caption         =   "Routine Infomation:"
      Enabled         =   0   'False
      Height          =   2415
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   2775
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         Begin VB.TextBox txtOffset 
            Enabled         =   0   'False
            Height          =   360
            Left            =   120
            MaxLength       =   6
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblOffset 
            AutoSize        =   -1  'True
            Caption         =   "Offset:"
            Enabled         =   0   'False
            Height          =   240
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   540
         End
      End
   End
   Begin VB.ListBox listRoutines 
      Enabled         =   0   'False
      Height          =   1260
      ItemData        =   "frmMain.frx":1272
      Left            =   120
      List            =   "frmMain.frx":127C
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "ROM Information:"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   2295
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         Caption         =   "Code:"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   225
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   465
      End
   End
   Begin MSComDlg.CommonDialog openfd 
      Left            =   5160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "frmMain.frx":129B
      Top             =   0
      Width           =   5655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open ROM"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert Routine"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
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
Private iFileNum As Integer
Private sFilePath As String
Private sHeader As String * 4
Private VarA As String
Private VarB As String
Private VarC As String
Private VarD As String

Private Sub mnuAbout_Click()
    frmAbout.Show , Me
End Sub

Private Sub mnuInsert_Click()
    'ReadIVs
    If listRoutines.Selected(0) = True Then
        
    'Different Music
    ElseIf listRoutines.Selected(1) = True Then
        VarA = Val(InputBox("Enter a value for variable set A & B", "Variable A", "7000"))
            If Len(VarA) > 4 Then
            Exit Sub
            ElseIf VarA = 0 Then
            Exit Sub
            End If
        VarB = Val(InputBox("Enter a value for variable set A & B", "Variable B", "7002"))
            If Len(VarB) > 4 Then
            Exit Sub
            ElseIf VarB = 0 Then
            Exit Sub
            End If
        VarC = Val(InputBox("Enter a value for variable set C & D", "Variable C", "7001"))
            If Len(VarC) > 4 Then
            Exit Sub
            ElseIf VarC = 0 Then
            Exit Sub
            End If
        VarD = Val(InputBox("Enter a value for variable set C & D", "Variable D", "7003"))
            If Len(VarD) > 4 Then
            Exit Sub
            ElseIf VarD = 0 Then
            Exit Sub
            End If
        DifferentMusic
        MsgBox "Routine Inserted Successfully", vbInformation
    End If
    
End Sub

Private Sub listRoutines_Click()
    'Read IVs Routine
    If listRoutines.Selected(0) = True Then
    
    'Different Music Routine
    ElseIf listRoutines.Selected(1) = True Then
        lblOffset.Enabled = True
        txtOffset.Enabled = True
        txtOffset.Text = "800000"
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String

iFileNum = FreeFile

    With openfd
        .Filter = "GBA Roms(*.gba,*.agb,*.bin)|*.gba;*.agb;*.bin"
        .DialogTitle = "Open ROM..."
        .ShowOpen
    End With
    
    sResult = openfd.FileName
    
    If LenB(sResult) > 0 Then
    
        sFilePath = sResult
       
        Open sResult For Binary As #iFileNum
            Get #iFileNum, &HAC + 1, sHeader
    
        Select Case sHeader
            Case "AXVE"
                lblROM.Caption = "POKEMON RUBY"
                lblHeader.Caption = sHeader
            Case "BPRE"
                lblROM.Caption = "POKEMON FIRE RED"
                lblHeader.Caption = sHeader
            Case "BPEE"
                lblROM.Caption = "POKEMON EMERALD"
                lblHeader.Caption = sHeader
            Case Else
                MsgBox "Error - UnSupported ROM", vbCritical
                Exit Sub
        End Select
            listRoutines.Enabled = True
            frStuff.Enabled = True
            lblMX.Enabled = True
            mnuInsert.Enabled = True
    End If
End Sub

'Inserting the Different Music Routine
Function DifferentMusic()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum

Dim Routine(135) As Byte
Routine(0) = &HFE
Routine(1) = &HB4
Routine(2) = &H5
Routine(3) = &H1C
Routine(4) = &H1B
Routine(5) = &H48
Routine(6) = &H20
Routine(7) = &H49
'Routine(8) = &H0
Routine(9) = &HF0
Routine(10) = &H33
Routine(11) = &HF8
'Routine(12) = &H0
Routine(13) = &H22
Routine(14) = &H90
Routine(15) = &H42
Routine(16) = &HB
Routine(17) = &HD0
Routine(18) = &H3
Routine(19) = &H88
Routine(20) = &H1B
Routine(21) = &H49
Routine(22) = &H8B
Routine(23) = &H42
Routine(24) = &H7
Routine(25) = &HD0
Routine(26) = &HAB
Routine(27) = &H42
Routine(28) = &H5
Routine(29) = &HD1
Routine(30) = &H17
Routine(31) = &H48
Routine(32) = &H19
Routine(33) = &H49
'Routine(34) = &H0
Routine(35) = &HF0
Routine(36) = &H26
Routine(37) = &HF8
Routine(38) = &H5
Routine(39) = &H88
Routine(40) = &H12
Routine(41) = &HE0
Routine(42) = &H13
Routine(43) = &H48
Routine(44) = &H16
Routine(45) = &H49
'Routine(46) = &H0
Routine(47) = &HF0
Routine(48) = &H20
Routine(49) = &HF8
'Routine(50) = &H0
Routine(51) = &H22
Routine(52) = &H90
Routine(53) = &H42
Routine(54) = &HB
Routine(55) = &HD0
Routine(56) = &H3
Routine(57) = &H88
Routine(58) = &H12
Routine(59) = &H49
Routine(60) = &H8B
Routine(61) = &H42
Routine(62) = &H7
Routine(63) = &HD0
Routine(64) = &HAB
Routine(65) = &H42
Routine(66) = &H5
Routine(67) = &HD1
Routine(68) = &HE
Routine(69) = &H48
Routine(70) = &H10
Routine(71) = &H49
'Routine(72) = &H0
Routine(73) = &HF0
Routine(74) = &H13
Routine(75) = &HF8
Routine(76) = &H5
Routine(77) = &H88
Routine(78) = &HFF
Routine(79) = &HE7
Routine(80) = &H28
Routine(81) = &H4
Routine(82) = &HFE
Routine(83) = &HBC
Routine(84) = &HD
Routine(85) = &H4A
Routine(86) = &H4
Routine(87) = &HB4
Routine(88) = &HD
Routine(89) = &H4A
Routine(90) = &HE
Routine(91) = &H49
Routine(92) = &H40
Routine(93) = &HB
Routine(94) = &H40
Routine(95) = &H18
Routine(96) = &H83
Routine(97) = &H88
Routine(98) = &H59
'Routine(99) = &H0
Routine(100) = &HC9
Routine(101) = &H18
Routine(102) = &H89
'Routine(103) = &H0
Routine(104) = &H89
Routine(105) = &H18
Routine(106) = &HA
Routine(107) = &H68
Routine(108) = &H1
Routine(109) = &H68
Routine(110) = &H10
Routine(111) = &H1C
'Routine(112) = &H0
Routine(113) = &HBD
Routine(114) = &H8
Routine(115) = &H47

Dim Replace1(5) As Byte
Replace1(0) = &H1
Replace1(1) = &H49
Replace1(2) = &H8
Replace1(3) = &H47
'Replace1(4) = &H0
'Replace1(5) = &H0

Dim Filler(1) As Byte
Dim Nil(15) As Byte
Dim Nil2(7) As Byte
        
          
            Put #iFileNum, CLng("&H" & txtOffset) + 1, Routine
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 116, CLng("&H" & VarA)
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 118, Filler
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 120, CLng("&H" & VarC)
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 122, Filler
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 124, CLng("&H" & VarB)
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 126, Filler
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 128, CLng("&H" & VarD)
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 130, Filler
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 132, &HFFFF
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 134, Filler
            Put #iFileNum, CLng("&H" & txtOffset) + 1 + 152, Nil2
        
        Select Case sHeader
            Case "BPRE"
                Put #iFileNum, &H1DD0F6 + 1, Replace1
                Put #iFileNum, &H1DD0F6 + 1 + 6, CLng(("&H" & txtOffset) + 1)
                Put #iFileNum, &H1DD0FF + 1, &H8
                Put #iFileNum, &H1DD100 + 1, Nil
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 136, &H806E455
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 140, &H81DD10F
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 144, &H84A329C
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 148, &H84A32CC
            
            Case "AXVE"
                Put #iFileNum, &H1DDEFA + 1, Replace1
                Put #iFileNum, &H1DDEFA + 1 + 6, CLng(("&H" & txtOffset) + 1)
                Put #iFileNum, &H1DDF03 + 1, &H8
                Put #iFileNum, &H1DDF04 + 1, Nil
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 136, &H8069211
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 140, &H81DDF13
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 144, &H845545C
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 148, &H845548C
            
            Case "BPEE"
                Put #iFileNum, &H2E0132 + 1, Replace1
                Put #iFileNum, &H2E0132 + 1 + 6, CLng(("&H" & txtOffset) + 1)
                Put #iFileNum, &H2E013B + 1, &H8
                Put #iFileNum, &H2E013C + 1, Nil
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 136, &H809D649
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 140, &H82E014B
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 144, &H86B4930
                Put #iFileNum, CLng("&H" & txtOffset) + 1 + 148, &H86B4960
        End Select
End Function

Function ReadIVs()
Dim Routine() As Byte
Routine(0) = &H7E
Routine(0) = &H7E

iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
End Function
