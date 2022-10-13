VERSION 5.00
Begin VB.Form frmTxtEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text Editor - Demo"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox listOffsets 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmTxtEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

