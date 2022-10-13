VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Harvest Moon Advance Editor - Alpha"
   ClientHeight    =   3705
   ClientLeft      =   150
   ClientTop       =   525
   ClientWidth     =   7905
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
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox listPeople 
      Enabled         =   0   'False
      Height          =   1740
      ItemData        =   "frmMain.frx":06C2
      Left            =   120
      List            =   "frmMain.frx":06C4
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame frData 
      Enabled         =   0   'False
      Height          =   2895
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   5175
      Begin VB.PictureBox Picture2 
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
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   4935
         TabIndex        =   5
         Top             =   240
         Width           =   4935
         Begin VB.TextBox txtName 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            MaxLength       =   4
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblppl2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Enabled         =   0   'False
            Height          =   240
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Data: "
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
      Begin VB.TextBox txtROM 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Nothing Loaded..."
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblsHeader 
         AutoSize        =   -1  'True
         Caption         =   "Nothing Loaded..."
         Height          =   240
         Left            =   720
         TabIndex        =   12
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         Caption         =   "Code:"
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblInfo 
         Caption         =   "ROM Information: "
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         Caption         =   "ROM:"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame frCopy1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   7695
      Begin VB.Label frCopy2 
         AutoSize        =   -1  'True
         Caption         =   "Copyright 2008 ZodiacDaGreat"
         Enabled         =   0   'False
         Height          =   240
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   2325
      End
   End
   Begin MSComDlg.CommonDialog openfd 
      Left            =   7200
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Begin VB.Menu mnuEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOthers 
      Caption         =   "&Others"
      Begin VB.Menu mnuHorse 
         Caption         =   "Prize/Mart Editor"
         Enabled         =   0   'False
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuTxtEditor 
         Caption         =   "Text Editor"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuReadMe 
         Caption         =   "ReadMe"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Enabled         =   0   'False
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
Private sHeader As String * 4
Private iFileNum As Integer
Private sName As String * 7
Private sName2 As String * 11
Private sName3 As String * 3

Private Sub mnuAbout_Click()
    frmAbout.Show , Me
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String
Dim Temp As String

    iFileNum = FreeFile

    With openfd
        .Filter = "Gameboy Advance ROMs (*.gba, *.agb, *.bin)|*.gba;*.agb;*.bin"
        .DialogTitle = "Open ROM"
        .ShowOpen
    End With
    
    sResult = openfd.FileName
    
    If LenB(sResult) > 0 Then
    
    sFilePath = sResult
        
        Open sResult For Binary As #iFileNum
            Get #iFileNum, &HAC + 1, sHeader
            
    Select Case sHeader
        Case "A4NE"
            txtROM.Text = sFilePath
            lblsHeader.Caption = sHeader
        
            lblppl2.Enabled = True
            txtName.Enabled = True
            frData.Enabled = True
            listPeople.Enabled = True
            frCopy2.Enabled = True
            frCopy1.Enabled = True
            mnuSave.Enabled = True
            mnuHorse.Enabled = True
            mnuTxtEditor.Enabled = True
        
        Case Else
            MsgBox "Error 1: Unsupported ROM", vbCritical
            listPeople.Enabled = False
            mnuSave.Enabled = False
        Exit Sub
        End Select
    End If
    
    Close #iFileNum
End Sub

Private Sub listPeople_Click()
Dim sOffset As String
Dim i As Integer

iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum

    'Rick's Data
    If listPeople.Selected(0) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Rick", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Popuri's Data
    If listPeople.Selected(1) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Popuri", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Barley's Data
    If listPeople.Selected(2) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Barley", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'May's Data
    If listPeople.Selected(3) = True Then
        'Name
        txtName.MaxLength = 3
        sOffset = GetFromINI("May", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName3
        txtName.Text = RTrim$(sName3)
        '
    End If
    
    'Saibara's Data
    If listPeople.Selected(4) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Saibara", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Gray's Data
    If listPeople.Selected(5) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Gray", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Duke's Data
    If listPeople.Selected(6) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Duke", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Manna's Data
    If listPeople.Selected(7) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Manna", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Basil's Data
    If listPeople.Selected(8) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Basil", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Anna's Data
    If listPeople.Selected(9) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Anna", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Mary's Data
    If listPeople.Selected(10) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Mary", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Thomas's Data
    If listPeople.Selected(11) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Thomas", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
        
    'Harris's Data
    If listPeople.Selected(12) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Harris", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Ellen's Data
    If listPeople.Selected(13) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Ellen", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Stu's Data
    If listPeople.Selected(14) = True Then
        'Name
        txtName.MaxLength = 3
        sOffset = GetFromINI("Stu", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName3
        txtName.Text = RTrim$(sName3)
        '
    End If
    
    'Jeff's Data
    If listPeople.Selected(15) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Jeff", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Sasha's Data
    If listPeople.Selected(16) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Sasha", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Karen's Data
    If listPeople.Selected(17) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Karen", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Doctor's Data
    If listPeople.Selected(18) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Doctor", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Elli's Data
    If listPeople.Selected(19) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Elli", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Carter's Data
    If listPeople.Selected(20) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Carter", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Cliff's Data
    If listPeople.Selected(21) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Cliff", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Doug's Data
    If listPeople.Selected(22) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Doug", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Ann's Data
    If listPeople.Selected(23) = True Then
        'Name
        txtName.MaxLength = 3
        sOffset = GetFromINI("Ann", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName3
        txtName.Text = RTrim$(sName3)
        '
    End If
    
    'Kai's Data
    If listPeople.Selected(24) = True Then
        'Name
        txtName.MaxLength = 3
        sOffset = GetFromINI("Kai", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName3
        txtName.Text = RTrim$(sName3)
        '
    End If
    
    'Gotz's Data
    If listPeople.Selected(25) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Gotz", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Zack's Data
    If listPeople.Selected(26) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Zack", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Won's Data
    If listPeople.Selected(27) = True Then
        'Name
        txtName.MaxLength = 3
        sOffset = GetFromINI("Won", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName3
        txtName.Text = RTrim$(sName3)
        '
    End If
    
    'Gourmet's Data
    If listPeople.Selected(28) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Gourmet", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'H. Goddess's Data
    If listPeople.Selected(29) = True Then
        'Name
        txtName.MaxLength = 11
        sOffset = GetFromINI("H. Goddess", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName2
        txtName.Text = RTrim$(sName2)
        '
    End If
    
    'Kappa's Data
    If listPeople.Selected(30) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Kappa", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Lou's Data
    If listPeople.Selected(31) = True Then
        'Name
        txtName.MaxLength = 3
        sOffset = GetFromINI("Lou", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName3
        txtName.Text = RTrim$(sName3)
        '
    End If
    
    'Lu's Data
    If listPeople.Selected(32) = True Then
        'Name
        txtName.MaxLength = 3
        sOffset = GetFromINI("Lu", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName3
        txtName.Text = RTrim$(sName3)
        '
    End If
    
    'Staid's Data
    If listPeople.Selected(33) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Staid", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Nappy's Data
    If listPeople.Selected(34) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Nappy", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Bold's Data
    If listPeople.Selected(35) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Bold", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Chef's Data
    If listPeople.Selected(36) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Chef", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Aqua's Data
    If listPeople.Selected(37) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Aqua", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Hoggy's Data
    If listPeople.Selected(38) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Hoggy", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Timid's Data
    If listPeople.Selected(39) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Timid", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
    
    'Lillia's Data
    If listPeople.Selected(40) = True Then
        'Name
        txtName.MaxLength = 7
        sOffset = GetFromINI("Lillia", "Name", 0, App.Path & "\\Data.ini")
        Get #iFileNum, sOffset + 1, sName
        txtName.Text = RTrim$(sName)
        '
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHorse_Click()
    frmMart.Show , Me
End Sub

Private Sub mnuSave_Click()

iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
    
    'Saving Rick's Data
    If listPeople.Selected(0) = True Then
        Put #iFileNum, &H104130 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
        
    End If
    
    'Saving Popuri's Data
    If listPeople.Selected(1) = True Then
        Put #iFileNum, &H104138 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Barleys's Data
    If listPeople.Selected(2) = True Then
        Put #iFileNum, &H104140 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving May's Data
    If listPeople.Selected(3) = True Then
        Put #iFileNum, &H104148 + 1, Left$(txtName.Text & String$(7, 0), 3) 'Writing Back Name
    
    End If
    
    'Saving Saibara's Data
    If listPeople.Selected(4) = True Then
        Put #iFileNum, &H10414C + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Gray's Data
    If listPeople.Selected(5) = True Then
        Put #iFileNum, &H104154 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Duke's Data
    If listPeople.Selected(6) = True Then
        Put #iFileNum, &H10415C + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Manna's Data
    If listPeople.Selected(7) = True Then
        Put #iFileNum, &H104164 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Basil's Data
    If listPeople.Selected(8) = True Then
        Put #iFileNum, &H10416C + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Anna's Data
    If listPeople.Selected(9) = True Then
        Put #iFileNum, &H104174 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Mary's Data
    If listPeople.Selected(10) = True Then
        Put #iFileNum, &H10417C + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Thomas's Data
    If listPeople.Selected(11) = True Then
        Put #iFileNum, &H104184 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Harris's Data
    If listPeople.Selected(12) = True Then
        Put #iFileNum, &H10418C + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Ellen's Data
    If listPeople.Selected(13) = True Then
        Put #iFileNum, &H104194 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Stu's Data
    If listPeople.Selected(14) = True Then
        Put #iFileNum, &H10419C + 1, Left$(txtName.Text & String$(7, 0), 3) 'Writing Back Name
    
    End If
    
    'Saving Jeff's Data
    If listPeople.Selected(15) = True Then
        Put #iFileNum, &H1041A0 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Sasha's Data
    If listPeople.Selected(16) = True Then
        Put #iFileNum, &H1041A8 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Karen's Data
    If listPeople.Selected(17) = True Then
        Put #iFileNum, &H1041B0 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Doctor's Data
    If listPeople.Selected(18) = True Then
        Put #iFileNum, &H1041B8 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Elli's Data
    If listPeople.Selected(19) = True Then
        Put #iFileNum, &H1041C0 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Carter's Data
    If listPeople.Selected(20) = True Then
        Put #iFileNum, &H1041C8 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Cliff's Data
    If listPeople.Selected(21) = True Then
        Put #iFileNum, &H1041D0 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Doug's Data
    If listPeople.Selected(22) = True Then
        Put #iFileNum, &H1041D8 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Ann's Data
    If listPeople.Selected(23) = True Then
        Put #iFileNum, &H1041E0 + 1, Left$(txtName.Text & String$(7, 0), 3) 'Writing Back Name
    
    End If
    
    'Saving Kai's Data
    If listPeople.Selected(24) = True Then
        Put #iFileNum, &H1041E4 + 1, Left$(txtName.Text & String$(7, 0), 3) 'Writing Back Name
    
    End If
    
    'Saving Gotz's Data
    If listPeople.Selected(25) = True Then
        Put #iFileNum, &H1041E8 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Zack's Data
    If listPeople.Selected(26) = True Then
        Put #iFileNum, &H1041F0 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Won's Data
    If listPeople.Selected(27) = True Then
        Put #iFileNum, &H1041F8 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Gourmet's Data
    If listPeople.Selected(28) = True Then
        Put #iFileNum, &H1041FC + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving H. Goddess's Data
    If listPeople.Selected(29) = True Then
        Put #iFileNum, &H104204 + 1, Left$(txtName.Text & String$(7, 0), 11) 'Writing Back Name
    
    End If
    
    'Saving Kappa's Data
    If listPeople.Selected(30) = True Then
        Put #iFileNum, &H104210 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Lou's Data
    If listPeople.Selected(31) = True Then
        Put #iFileNum, &H104218 + 1, Left$(txtName.Text & String$(7, 0), 3) 'Writing Back Name
    
    End If
    
    'Saving Lu's Data
    If listPeople.Selected(32) = True Then
        Put #iFileNum, &H10421C + 1, Left$(txtName.Text & String$(7, 0), 3) 'Writing Back Name
    
    End If
    
    'Saving Staid's Data
    If listPeople.Selected(33) = True Then
        Put #iFileNum, &H104220 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Nappy's Data
    If listPeople.Selected(34) = True Then
        Put #iFileNum, &H104228 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Bold's Data
    If listPeople.Selected(35) = True Then
        Put #iFileNum, &H104230 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Chef's Data
    If listPeople.Selected(36) = True Then
        Put #iFileNum, &H104238 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Aqua's Data
    If listPeople.Selected(37) = True Then
        Put #iFileNum, &H104240 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Hoggy's Data
    If listPeople.Selected(38) = True Then
        Put #iFileNum, &H104248 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Timid's Data
    If listPeople.Selected(39) = True Then
        Put #iFileNum, &H104250 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
    
    'Saving Lillia's Data
    If listPeople.Selected(40) = True Then
        Put #iFileNum, &H104128 + 1, Left$(txtName.Text & String$(7, 0), 7) 'Writing Back Name
    
    End If
MsgBox "Data Saved", vbInformation
Close #iFileNum
End Sub

Private Sub mnuTxtEditor_Click()
    frmTxtEditor.Show , Me
End Sub
