VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2985
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
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblThethethethe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "thethethethe - Updated Source Code"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2280
   End
   Begin VB.Label lblKyoufu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kyoufu Kawa - Original Source Code and Item Data Structure"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   3870
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "~Description here~"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4530
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   0
      X2              =   304
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Image Image1 
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

Private Sub cmdOkay_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version" & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = "~" & App.FileDescription & "~"
End Sub
