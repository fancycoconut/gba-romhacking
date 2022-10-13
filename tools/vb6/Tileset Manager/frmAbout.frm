VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   192
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "9"
   Begin VB.Frame frCredits 
      Caption         =   "Greetz"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4455
      Begin VB.Label lblNames 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LU-HO"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Yay"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblZodiacDaGreat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZodiacDaGreat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   915
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      X1              =   240
      X2              =   72
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Image imgVisitSite 
      Height          =   225
      Left            =   4320
      Picture         =   "frmAbout.frx":000C
      ToolTipText     =   "Visit AHP website"
      Top             =   1080
      Width           =   225
   End
   Begin VB.Shape shpHilt 
      BorderColor     =   &H80000000&
      Height          =   315
      Left            =   4275
      Shape           =   5  'Rounded Square
      Top             =   1035
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   0
      X2              =   312
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   1080
   End
   Begin VB.Image imgAboutBanner 
      Height          =   975
      Left            =   0
      Picture         =   "frmAbout.frx":02EA
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Tilesets ^ ^"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1770
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tileset Manager"
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
      TabIndex        =   0
      Top             =   1080
      Width           =   1365
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Localize Me
    lblVersion.Caption = LoadResString(36) & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblNames.Caption = "LU-HO                          HackMew                       D-Trogh"
    shpHilt.Left = 350
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    shpHilt.Left = 350
End Sub

Private Sub imgVisitSite_Click()
    ShellExecute 0, vbNullString, "http://ahp.freebyte.us", vbNullString, "", 1
End Sub

Private Sub imgVisitSite_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    shpHilt.Left = 285
End Sub
