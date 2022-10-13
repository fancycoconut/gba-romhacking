VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2760
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblHackMew 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HackMew"
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label lblMX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mastermind_X"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label lblGreetz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greetz"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   570
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance IntroEd"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   960
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version" & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
