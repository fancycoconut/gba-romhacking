VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "$tart Money Ed"
   ClientHeight    =   1200
   ClientLeft      =   150
   ClientTop       =   525
   ClientWidth     =   3735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMoney 
      Enabled         =   0   'False
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin MSComDlg.CommonDialog openfd 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Am&ount($):"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   765
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ROM:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open ROM"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save ROM"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other"
      Begin VB.Menu mnuTruck 
         Caption         =   "Remove Truck Ani."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
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

Private Sub mnuAbout_Click()
 frmAbout.Show , Me
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String
Dim Money As Long

iFileNum = FreeFile

    With openfd
        .Filter = "GBA Roms|*.gba"
        .DialogTitle = "Open Rom"
        .ShowOpen
    End With
    
    sResult = openfd.FileName
    
    If LenB(sResult) > 0 Then
        
        sFilePath = sResult
        
        Open sResult For Binary As #iFileNum
            Get #iFileNum, &HAD, sHeader
        
        Select Case sHeader
            Case "BPRE"
                lblName.Caption = "Fire Red Version {" & sHeader & "}"
                Get #iFileNum, &H54B60 + 1, Money
                txtMoney.Text = Money
                
            Case "BPRF"
                lblName.Caption = "Fire Red Version {" & sHeader & "}"
                Get #iFileNum, &H54C40 + 1, Money
                txtMoney.Text = Money
            
            Case "BPRI"
                lblName.Caption = "Fire Red Version {" & sHeader & "}"
                Get #iFileNum, &H54B6C + 1, Money
                txtMoney.Text = Money
                
            Case "BPGE"
                lblName.Caption = "Fire Red Version {" & sHeader & "}"
                Get #iFileNum, &H54B60 + 1, Money
                txtMoney.Text = Money
            
            Case "BPGF"
                lblName.Caption = "Fire Red Version {" & sHeader & "}"
                Get #iFileNum, &H54C40 + 1, Money
                txtMoney.Text = Money
                
            Case "AXVE"
                lblName.Caption = "Ruby Version {" & sHeader & "}"
                Get #iFileNum, &H52F4C + 1, Money
                txtMoney.Text = Money
                
            Case "AXVF"
                lblName.Caption = "Ruby Version {" & sHeader & "}"
                Get #iFileNum, &H53378 + 1, Money
                txtMoney.Text = Money
                
            Case "AXPF"
                lblName.Caption = "Ruby Version {" & sHeader & "}"
                Get #iFileNum, &H53378 + 1, Money
                txtMoney.Text = Money
                
            Case "BPEE"
                lblName.Caption = "Emerald Version {" & sHeader & "}"
                Get #iFileNum, &H845BC + 1, Money
                txtMoney.Text = Money
            
            Case "BPED"
                lblName.Caption = "Emerald Version {" & sHeader & "}"
                Get #iFileNum, &H845D8 + 1, Money
                txtMoney.Text = Money
            
            Case "BPEF"
                lblName.Caption = "Emerald Version {" & sHeader & "}"
                Get #iFileNum, &H845CC + 1, Money
                txtMoney.Text = Money
                
            Case "AXPE"
                lblName.Caption = "Sapphire Version {" & sHeader & "}"
                Get #iFileNum, &H52F4C + 1, Money
                txtMoney.Text = Money
                
            Case Else
                MsgBox "Error 1: Non-Pokemon ROM/Unsupported Rom", vbExclamation, "$tart Money Ed"
                Exit Sub
            End Select
                            
            txtMoney.Enabled = True
            mnuSave.Enabled = True
            mnuTruck.Enabled = True
             
    End If
    
    Close #iFileNum
End Sub

Private Sub mnuTruck_Click()
Select Case sHeader
    Case "AXVE"
        PutFiller sFilePath, &HC757E, 2
        PutFiller sFilePath, &HC759E, 4
        PutFiller sFilePath, &HC75B4, 2
        PutFiller sFilePath, &HC75D8, 2
        PutFiller sFilePath, &HC75E2, 2
        PutFiller sFilePath, &HC75F0, 4
        PutFiller sFilePath, &HC7600, 4
        PutFiller sFilePath, &HC7624, 4
        PutFiller sFilePath, &HC7640, 2
        PutFiller sFilePath, &HC7644, 4
        PutFiller sFilePath, &HC765E, 2
        PutFiller sFilePath, &HC7668, 4
        PutFiller sFilePath, &HC7674, 4
        PutFiller sFilePath, &HC7680, 4
        PutFiller sFilePath, &HC768A, 4
    Case "BPEE"
        PutFiller sFilePath, &HFB3BE, 2
        PutFiller sFilePath, &HFB3DE, 4
        PutFiller sFilePath, &HFB3F4, 2
        PutFiller sFilePath, &HFB418, 2
        PutFiller sFilePath, &HFB422, 2
        PutFiller sFilePath, &HFB430, 4
        PutFiller sFilePath, &HFB440, 4
        PutFiller sFilePath, &HFB464, 4
        PutFiller sFilePath, &HFB480, 2
        PutFiller sFilePath, &HFB484, 4
        PutFiller sFilePath, &HFB49E, 2
        PutFiller sFilePath, &HFB4A8, 4
        PutFiller sFilePath, &HFB4B4, 4
        PutFiller sFilePath, &HFB4C0, 4
        PutFiller sFilePath, &HFB4CA, 4
End Select
MsgBox "Truck animation removed successfully!" & vbNewLine & "Credits to MX{Emerald} & HackMew{Ruby}", vbInformation
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub mnuSave_Click()
iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
    

    
    Select Case sHeader
        Case "AXVF"
        Put #iFileNum, &H53378 + 1, CLng(txtMoney)
        
        Case "AXPF"
        Put #iFileNum, &H53378 + 1, CLng(txtMoney)
    
        Case "AXVE"
        Put #iFileNum, &H52F4C + 1, CLng(txtMoney)
        
        Case "BPRE"
        Put #iFileNum, &H54B60 + 1, CLng(txtMoney)
        
        Case "BPRF"
        Put #iFileNum, &H54C40 + 1, CLng(txtMoney)
        
        Case "BPRI"
        Put #iFileNum, &H54B6C + 1, CLng(txtMoney)
        
        Case "BPEE"
        Put #iFileNum, &H845BC + 1, CLng(txtMoney)
        
        Case "BPED"
        Put #iFileNum, &H845D8 + 1, CLng(txtMoney)
        
        Case "BPEF"
        Put #iFileNum, &H845CC + 1, CLng(txtMoney)
        
        Case "AXPE"
        Put #iFileNum, &H52F4C + 1, CLng(txtMoney)
        
        Case "BPGE"
        Put #iFileNum, &H54B60 + 1, CLng(txtMoney)
        
        Case "BPGF"
        Put #iFileNum, &H54C40 + 1, CLng(txtMoney)
        
    End Select
    
    MsgBox "Saved", vbInformation, "$tart Money Ed"
    
    
    Close #iFileNum
End Sub
