VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
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
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrEffects 
      Interval        =   1000
      Left            =   1080
      Top             =   2400
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Door Manager"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":000C
      Height          =   1215
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ZodiacDaGreat 2009"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Width           =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   72
      X2              =   160
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":00FE
      Top             =   240
      Width           =   480
   End
   Begin VB.Shape shpBanner 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Localize Me
    lblVersion.Caption = "Version" & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub tmrEffects_Timer()
Dim i As Integer
    i = i + 1
    imgIcon.Top = imgIcon.Top + 1
    If imgIcon.Top = 152 Then imgIcon.Top = 16
End Sub
