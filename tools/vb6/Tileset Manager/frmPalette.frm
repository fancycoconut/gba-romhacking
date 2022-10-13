VERSION 5.00
Begin VB.Form frmPalette 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Palette Inserter"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
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
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "39"
   Begin VB.ComboBox cmbPal 
      Height          =   315
      ItemData        =   "frmPalette.frx":0000
      Left            =   2640
      List            =   "frmPalette.frx":0016
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Tag             =   "45"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsertPalette 
      Caption         =   "Insert Palette"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "44"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtPalOffset 
      Height          =   285
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
   Begin VB.Frame fraPalette 
      Caption         =   "Palette"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Tag             =   "43"
      Top             =   720
      Width           =   6135
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   5895
         TabIndex        =   8
         Top             =   240
         Width           =   5895
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   0
            Left            =   120
            MaxLength       =   4
            TabIndex        =   40
            Top             =   120
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   39
            Top             =   480
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   1
            Left            =   840
            MaxLength       =   4
            TabIndex        =   38
            Top             =   120
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   2
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   37
            Top             =   120
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   3
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   36
            Top             =   120
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   4
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   35
            Top             =   120
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   5
            Left            =   3720
            MaxLength       =   4
            TabIndex        =   34
            Top             =   120
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   6
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   33
            Top             =   120
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   7
            Left            =   5160
            MaxLength       =   4
            TabIndex        =   32
            Top             =   120
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   840
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   31
            Top             =   480
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1560
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   30
            Top             =   480
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2280
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   29
            Top             =   480
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   3000
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   28
            Top             =   480
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   3720
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   27
            Top             =   480
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   4440
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   26
            Top             =   480
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   5160
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   25
            Top             =   480
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   8
            Left            =   120
            MaxLength       =   4
            TabIndex        =   24
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   9
            Left            =   840
            MaxLength       =   4
            TabIndex        =   23
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   10
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   22
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   11
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   21
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   12
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   20
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   13
            Left            =   3720
            MaxLength       =   4
            TabIndex        =   19
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   14
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   18
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtPal 
            Height          =   300
            Index           =   15
            Left            =   5160
            MaxLength       =   4
            TabIndex        =   17
            Top             =   840
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   16
            Top             =   1200
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   840
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   15
            Top             =   1200
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   1560
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   14
            Top             =   1200
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   2280
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   13
            Top             =   1200
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   3000
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   12
            Top             =   1200
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   3720
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   11
            Top             =   1200
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   4440
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   10
            Top             =   1200
            Width           =   600
         End
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   5160
            ScaleHeight     =   225
            ScaleWidth      =   570
            TabIndex        =   9
            Top             =   1200
            Width           =   600
         End
      End
   End
   Begin VB.Label lblIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index"
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Tag             =   "42"
      Top             =   240
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000013&
      X1              =   72
      X2              =   416
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Label lblZodiacDaGreat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZodiacDaGreat"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label lblPalOffset 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palette Offset"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Tag             =   "41"
      Top             =   240
      Width           =   960
   End
   Begin VB.Menu mnuRMB 
      Caption         =   "RMB"
      Visible         =   0   'False
      Begin VB.Menu mnuGetLog 
         Caption         =   "Get from log.."
         HelpContextID   =   46
      End
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnClearing As Boolean, blnClearing2 As Boolean
Private Const vbUnSafeColor As Long = &HC0C0FF

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsertPalette_Click()
Dim Data As Byte, Data2 As Byte
Dim i As Long, x As Long, iSum As Integer, Offset As Long
If Len(txtPalOffset.Text) = 0 Then Exit Sub

iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
x = 1
iSum = &H20 * cmbPal.ListIndex
Offset = CLng("&H" & txtPalOffset.Text)
For i = txtPal.LBound To txtPal.UBound
    Data = Val("&H" & Left$(txtPal(i).Text, 2))
    Data2 = Val("&H" & Right$(txtPal(i).Text, 2))
    Put #iFileNum, Offset + iSum + x + i, Data
    Put #iFileNum, Offset + iSum + x + i + 1, Data2
    x = x + 1
Next
Close #iFileNum
MsgBox LoadResString(47) & " 0x" & txtPalOffset.Text, vbInformation
Unload Me
End Sub

Private Sub Form_Load()
    Localize Me
    
    If EnlargedROM = 1 Then txtPalOffset.MaxLength = 7
    cmbPal.ListIndex = 0
End Sub

Private Sub txtPal_Change(Index As Integer)
If Not IsHex(txtPal(Index).Text) And Not blnClearing2 Then
    txtPal(Index).Text = vbNullString
    picPal(Index).BackColor = vbWhite
    Exit Sub
End If

If Len(txtPal(Index).Text) = 4 Then
    picPal(Index).BackColor = "&H" & GBA2RGB(txtPal(Index).Text, True)
Else
    picPal(Index).BackColor = vbWhite
End If
End Sub

Private Function SafeCheck(sString As String) As Boolean
Const IntMax As Integer = 32767 '&H7FFF

If Len(sString) < 4 Then SafeCheck = True: Exit Function

sString = Right$(sString, 2) & Left$(sString, 2)

If CLng("&H" & sString) <= IntMax And CLng("&H" & sString) >= 0 Then
    SafeCheck = True
Else
    SafeCheck = False
End If
End Function

Private Sub txtPalOffset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbMiddleButton Then PopupMenu mnuRMB, vbPopupMenuRightButton
End Sub

Private Sub mnuGetLog_Click()
txtPalOffset.Text = GetFromINI("Log", "Tileset Palette Offset", vbNullString, App.Path & "\log.txt")
End Sub
