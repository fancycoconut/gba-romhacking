VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2640
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
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "15"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblAeonos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aeonos"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label lblTutti 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tutti"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label lblDTrogh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D-Trogh"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label lblThanks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label lblMastermindX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mastermind_X - Original programming and code"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label lblStaff 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Credits"
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000007&
      X1              =   8
      X2              =   304
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DayAndNight"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      Height          =   195
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   960
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
    Localize Me
    lblVersion.Caption = "Version" & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
