VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Trebuchet MS"
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
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frCopy 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4455
      Begin VB.Label lblCopy 
         AutoSize        =   -1  'True
         Caption         =   "Copyright © 2008 ZodiacDaGreat"
         Enabled         =   0   'False
         Height          =   240
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   2490
      End
   End
   Begin VB.CommandButton cmdAlright 
      Caption         =   "Alright!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line2 
      DrawMode        =   5  'Not Copy Pen
      X1              =   120
      X2              =   4560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      DrawMode        =   5  'Not Copy Pen
      X1              =   120
      X2              =   4560
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://sfc.pokemon-inside.net"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For more information please visit: "
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2580
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v0.0.0"
      Height          =   240
      Left            =   3960
      TabIndex        =   3
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Ultimate ASM Tool For Inserting Routines."
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3405
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAlright_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub
