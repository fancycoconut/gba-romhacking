VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3105
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
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "11"
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Okay"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Tag             =   "34"
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame fraCredits 
      Caption         =   "Credits and Greetz"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Tag             =   "32"
      Top             =   1440
      Width           =   4455
      Begin VB.Label lblNames 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ash2000..."
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblEnjoy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enjoy!"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Tag             =   "33"
         Top             =   840
         Width           =   465
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0.0"
         Height          =   195
         Left            =   3120
         TabIndex        =   1
         Tag             =   "31"
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Picture         =   "frmAbout.frx":000C
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUnload_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Localize Me
    lblVersion.Caption = LoadResString(31) & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblNames.Caption = "Ash2000                  HackMew                  Swampert22" & vbNewLine & "D-Trogh                   Zel"
End Sub
