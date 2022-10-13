VERSION 5.00
Begin VB.Form frmMart 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Prize/Mart Editor"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMart2 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4695
         TabIndex        =   19
         Top             =   240
         Width           =   4695
         Begin VB.Frame frGotz 
            Caption         =   "Frame1"
            Height          =   1335
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   4695
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   4455
               TabIndex        =   30
               Top             =   240
               Width           =   4455
               Begin VB.CommandButton cmdGotz 
                  Caption         =   "Save"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   3480
                  TabIndex        =   34
                  Top             =   720
                  Width           =   975
               End
               Begin VB.TextBox txtGotz 
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   5
                  TabIndex        =   33
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.ListBox listWood 
                  Height          =   780
                  ItemData        =   "frmMart.frx":000C
                  Left            =   0
                  List            =   "frmMart.frx":004A
                  TabIndex        =   32
                  Top             =   120
                  Width           =   1815
               End
               Begin VB.TextBox txtLumber 
                  Height          =   375
                  Left            =   3360
                  MaxLength       =   5
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label lblCost1 
                  AutoSize        =   -1  'True
                  Caption         =   "Cost:"
                  Height          =   240
                  Left            =   1920
                  TabIndex        =   36
                  Top             =   0
                  Width           =   405
               End
               Begin VB.Label lblLumber 
                  AutoSize        =   -1  'True
                  Caption         =   "Lumber:"
                  Height          =   240
                  Left            =   3240
                  TabIndex        =   35
                  Top             =   0
                  Width           =   615
               End
            End
            Begin VB.Label lblGotz 
               AutoSize        =   -1  'True
               Caption         =   "Gotz The WoodCutter's: "
               ForeColor       =   &H80000002&
               Height          =   240
               Left            =   120
               TabIndex        =   37
               Top             =   0
               Width           =   1860
            End
         End
         Begin VB.CommandButton cmdNext2 
            Caption         =   "Next"
            Height          =   375
            Left            =   3840
            TabIndex        =   28
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton cmdPrevious1 
            Caption         =   "Previous"
            Height          =   375
            Left            =   2880
            TabIndex        =   27
            Top             =   2880
            Width           =   855
         End
         Begin VB.Frame frBarley 
            Height          =   1335
            Left            =   0
            TabIndex        =   20
            Top             =   1440
            Width           =   4695
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   4455
               TabIndex        =   21
               Top             =   240
               Width           =   4455
               Begin VB.CommandButton cmdBarley 
                  Caption         =   "Save"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   3360
                  TabIndex        =   24
                  Top             =   360
                  Width           =   975
               End
               Begin VB.TextBox txtBarley 
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   5
                  TabIndex        =   23
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.ListBox listBarley 
                  Height          =   780
                  ItemData        =   "frmMart.frx":014B
                  Left            =   0
                  List            =   "frmMart.frx":0164
                  TabIndex        =   22
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.Label lblCost2 
                  AutoSize        =   -1  'True
                  Caption         =   "Cost:"
                  Height          =   240
                  Left            =   1920
                  TabIndex        =   25
                  Top             =   120
                  Width           =   405
               End
            End
            Begin VB.Label lblBarley 
               AutoSize        =   -1  'True
               Caption         =   "Yodel Farm: "
               ForeColor       =   &H80000002&
               Height          =   240
               Left            =   120
               TabIndex        =   26
               Top             =   0
               Width           =   930
            End
         End
      End
   End
   Begin VB.Frame frMart1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4695
         TabIndex        =   2
         Top             =   240
         Width           =   4695
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next"
            Height          =   375
            Left            =   3840
            TabIndex        =   17
            Top             =   2880
            Width           =   855
         End
         Begin VB.Frame frHorse 
            Height          =   1335
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   4695
            Begin VB.PictureBox pic1 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   4455
               TabIndex        =   11
               Top             =   240
               Width           =   4455
               Begin VB.CommandButton cmdItems 
                  Caption         =   "Save"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   3360
                  TabIndex        =   14
                  Top             =   360
                  Width           =   975
               End
               Begin VB.TextBox txtMedals 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   5
                  TabIndex        =   13
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.ListBox listItems 
                  Height          =   780
                  ItemData        =   "frmMart.frx":01C4
                  Left            =   0
                  List            =   "frmMart.frx":01EE
                  TabIndex        =   12
                  Top             =   120
                  Width           =   1815
               End
               Begin VB.Label lblMedals 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Medals Required:"
                  Height          =   240
                  Left            =   1920
                  TabIndex        =   15
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.Label lblHorse 
               AutoSize        =   -1  'True
               Caption         =   "Horse Race Prize Editor: "
               ForeColor       =   &H80000002&
               Height          =   240
               Left            =   120
               TabIndex        =   16
               Top             =   0
               Width           =   1890
            End
         End
         Begin VB.Frame frSaibara 
            Caption         =   "Frame1"
            Height          =   1335
            Left            =   0
            TabIndex        =   3
            Top             =   1440
            Width           =   4695
            Begin VB.PictureBox pic2 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   4455
               TabIndex        =   4
               Top             =   240
               Width           =   4455
               Begin VB.CommandButton cmdSaibara 
                  Caption         =   "Save"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   3360
                  TabIndex        =   7
                  Top             =   360
                  Width           =   975
               End
               Begin VB.TextBox txtSaibara 
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   5
                  TabIndex        =   6
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.ListBox listSaibara 
                  Height          =   780
                  ItemData        =   "frmMart.frx":0273
                  Left            =   0
                  List            =   "frmMart.frx":0295
                  TabIndex        =   5
                  Top             =   120
                  Width           =   1815
               End
               Begin VB.Label lblPrice1 
                  AutoSize        =   -1  'True
                  Caption         =   "Price:"
                  Height          =   240
                  Left            =   1920
                  TabIndex        =   8
                  Top             =   120
                  Width           =   450
               End
            End
            Begin VB.Label lblSaibara 
               AutoSize        =   -1  'True
               Caption         =   "Saibara The Blacksmith: "
               ForeColor       =   &H80000002&
               Height          =   240
               Left            =   120
               TabIndex        =   9
               Top             =   0
               Width           =   1815
            End
         End
      End
   End
   Begin VB.Frame frMart3 
      Height          =   3615
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4695
         TabIndex        =   38
         Top             =   240
         Width           =   4695
         Begin VB.CommandButton cmdNext3 
            Caption         =   "Next"
            Height          =   375
            Left            =   3840
            TabIndex        =   42
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton cmdPrevious2 
            Caption         =   "Previous"
            Height          =   375
            Left            =   2880
            TabIndex        =   41
            Top             =   2880
            Width           =   855
         End
         Begin VB.Frame frPoutry 
            Height          =   1335
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   4695
            Begin VB.PictureBox Picture6 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   4455
               TabIndex        =   43
               Top             =   240
               Width           =   4455
               Begin VB.ListBox listPoutry 
                  Height          =   780
                  ItemData        =   "frmMart.frx":0301
                  Left            =   0
                  List            =   "frmMart.frx":030E
                  TabIndex        =   46
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.TextBox txtPoutry 
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   5
                  TabIndex        =   45
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.CommandButton cmdPoutry 
                  Caption         =   "Save"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   3360
                  TabIndex        =   44
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label lblCost3 
                  AutoSize        =   -1  'True
                  Caption         =   "Cost:"
                  Height          =   240
                  Left            =   1920
                  TabIndex        =   47
                  Top             =   120
                  Width           =   405
               End
            End
            Begin VB.Label lblPoutry 
               AutoSize        =   -1  'True
               Caption         =   "Poutry Farm:"
               ForeColor       =   &H80000002&
               Height          =   240
               Left            =   120
               TabIndex        =   40
               Top             =   0
               Width           =   990
            End
         End
      End
   End
End
Attribute VB_Name = "frmMart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iFileNum As Integer
Private sFilePath As String

Private Sub listPoutry_Click()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
    
    If listPoutry.Selected(0) = True Then
        txtPoutry.Text = GetWord(sFilePath, &HFD990 + 1)
    ElseIf listPoutry.Selected(1) = True Then
        txtPoutry.Text = GetWord(sFilePath, &HFD9A4 + 1)
    ElseIf listPoutry.Selected(2) = True Then
        txtPoutry.Text = GetWord(sFilePath, &HFD9B8 + 1)
    End If
End Sub

Private Sub cmdPoutry_click()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
    If CLng(txtPoutry.Text) > "65535" Then
    MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
    Exit Sub
    End If
    
    If listPoutry.Selected(0) = True Then
        PutWord sFilePath, &HFD990, txtPoutry.Text
    ElseIf listPoutry.Selected(1) = True Then
        PutWord sFilePath, &HFD9A4, txtPoutry.Text
    ElseIf listPoutry.Selected(2) = True Then
        PutWord sFilePath, &HFD9B8, txtPoutry.Text
    End If
    MsgBox "Poutry Data Saved!", vbInformation
End Sub

Private Sub listBarley_Click()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
    
    If listBarley.Selected(0) = True Then
        txtBarley.Text = GetWord(sFilePath, &HFFB98 + 1)
    ElseIf listBarley.Selected(1) = True Then
        txtBarley.Text = GetWord(sFilePath, &HFFBAC + 1)
    ElseIf listBarley.Selected(2) = True Then
        txtBarley.Text = GetWord(sFilePath, &HFFBC0 + 1)
    ElseIf listBarley.Selected(3) = True Then
        txtBarley.Text = GetWord(sFilePath, &HFFBD4 + 1)
    ElseIf listBarley.Selected(4) = True Then
        txtBarley.Text = GetWord(sFilePath, &HFFBE8 + 1)
    ElseIf listBarley.Selected(5) = True Then
        txtBarley.Text = GetWord(sFilePath, &HFFBFC + 1)
    ElseIf listBarley.Selected(6) = True Then
        txtBarley.Text = GetWord(sFilePath, &HFFC10 + 1)
    End If
End Sub

Private Sub cmdBarley_Click()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
    If CLng(txtBarley.Text) > "65535" Then
    MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
    Exit Sub
    End If
    
    If listBarley.Selected(0) = True Then
        PutWord sFilePath, &HFFB98, txtBarley.Text
    ElseIf listBarley.Selected(1) = True Then
        PutWord sFilePath, &HFFBAC, txtBarley.Text
    ElseIf listBarley.Selected(2) = True Then
        PutWord sFilePath, &HFFBC0, txtBarley.Text
    ElseIf listBarley.Selected(3) = True Then
        PutWord sFilePath, &HFFBD4, txtBarley.Text
    ElseIf listBarley.Selected(4) = True Then
        PutWord sFilePath, &HFFBE8, txtBarley.Text
    ElseIf listBarley.Selected(5) = True Then
        PutWord sFilePath, &HFFBFC, txtBarley.Text
    ElseIf listBarley.Selected(6) = True Then
        PutWord sFilePath, &HFFC10, txtBarley.Text
    End If
    MsgBox "Yodel Farm Data Saved!", vbInformation
End Sub

Private Sub listWood_Click()
Dim Cost As Long

iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
    
    If listWood.Selected(0) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF6AC + 1)
        txtLumber.Text = ""
    ElseIf listWood.Selected(1) = True Then
        Get #iFileNum, &HFF6C0 + 1, Cost
        txtGotz.MaxLength = 8
        txtGotz.Text = Cost
        txtLumber.Text = ""
    ElseIf listWood.Selected(2) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF6D4 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF6D8 + 1)
    ElseIf listWood.Selected(3) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF6E8 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF6EC + 1)
    ElseIf listWood.Selected(4) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF6FC + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF700 + 1)
    ElseIf listWood.Selected(5) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF710 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF714 + 1)
    ElseIf listWood.Selected(6) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF724 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF728 + 1)
    ElseIf listWood.Selected(7) = True Then
        Get #iFileNum, &HFF738 + 1, Cost
        txtGotz.MaxLength = 9
        txtGotz.Text = Cost
        txtLumber.Text = GetWord(sFilePath, &HFF73C + 1)
    ElseIf listWood.Selected(8) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF788 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF78C + 1)
    ElseIf listWood.Selected(9) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF79C + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF7A0 + 1)
    ElseIf listWood.Selected(10) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF7B0 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF7B4 + 1)
    ElseIf listWood.Selected(11) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF7C4 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF7C8 + 1)
    ElseIf listWood.Selected(12) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF7D8 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF7DC + 1)
    ElseIf listWood.Selected(13) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF7EC + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF7F0 + 1)
    ElseIf listWood.Selected(14) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF800 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF804 + 1)
    ElseIf listWood.Selected(15) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF814 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF818 + 1)
    ElseIf listWood.Selected(16) = True Then
        txtGotz.MaxLength = 5
        txtGotz.Text = GetWord(sFilePath, &HFF828 + 1)
        txtLumber.Text = GetWord(sFilePath, &HFF82C + 1)
    End If
End Sub

Private Sub cmdGotz_Click()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum

    If listWood.Selected(0) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF6AC, txtGotz.Text
    ElseIf listWood.Selected(1) = True Then
        If CLng(txtGotz.Text) > "16777215" Then
            MsgBox "Error - Invalid price value which should be between 0 and 16777215", vbCritical
            Exit Sub
        End If
        Put #iFileNum, &HFF6C0 + 1, CLng(txtGotz.Text)
    ElseIf listWood.Selected(2) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF6D4, txtGotz.Text
        PutWord sFilePath, &HFF6D8, txtLumber.Text
    ElseIf listWood.Selected(3) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF6E8, txtGotz.Text
        PutWord sFilePath, &HFF6EC, txtLumber.Text
    ElseIf listWood.Selected(4) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF6FC, txtGotz.Text
        PutWord sFilePath, &HFF700, txtLumber.Text
    ElseIf listWood.Selected(5) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF710, txtGotz.Text
        PutWord sFilePath, &HFF714, txtLumber.Text
    ElseIf listWood.Selected(6) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF724, txtGotz.Text
        PutWord sFilePath, &HFF728, txtLumber.Text
    ElseIf listWood.Selected(7) = True Then
        If CLng(txtGotz.Text) > "999999999" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        Put #iFileNum, &HFF738 + 1, CLng(txtGotz.Text)
        PutWord sFilePath, &HFF728, txtLumber.Text
    ElseIf listWood.Selected(8) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF788, txtGotz.Text
        PutWord sFilePath, &HFF78C, txtLumber.Text
    ElseIf listWood.Selected(9) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF79C, txtGotz.Text
        PutWord sFilePath, &HFF7A0, txtLumber.Text
    ElseIf listWood.Selected(10) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF7B0, txtGotz.Text
        PutWord sFilePath, &HFF7B4, txtLumber.Text
    ElseIf listWood.Selected(11) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF7C4, txtGotz.Text
        PutWord sFilePath, &HFF7C8, txtLumber.Text
    ElseIf listWood.Selected(12) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF7D8, txtGotz.Text
        PutWord sFilePath, &HFF7DC, txtLumber.Text
    ElseIf listWood.Selected(13) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF7EC, txtGotz.Text
        PutWord sFilePath, &HFF7F0, txtLumber.Text
    ElseIf listWood.Selected(14) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF800, txtGotz.Text
        PutWord sFilePath, &HFF804, txtLumber.Text
    ElseIf listWood.Selected(15) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF814, txtGotz.Text
        PutWord sFilePath, &HFF818, txtLumber.Text
    ElseIf listWood.Selected(16) = True Then
        If CLng(txtGotz.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        If CLng(txtLumber.Text) > "65535" Then
            MsgBox "Error - Invalid lumber value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        PutWord sFilePath, &HFF828, txtGotz.Text
        PutWord sFilePath, &HFF82C, txtLumber.Text
    End If
    MsgBox "Gotz The WoodCutter's Data Saved!", vbInformation
End Sub

Private Sub cmdNext_Click()
    frMart1.Visible = False
    frMart2.Visible = True
End Sub

Private Sub cmdNext2_Click()
    frMart2.Visible = False
    frMart3.Visible = True
End Sub

Private Sub cmdPrevious1_Click()
    frMart1.Visible = True
    frMart2.Visible = False
End Sub

Private Sub cmdPrevious2_Click()
    frMart3.Visible = False
    frMart2.Visible = True
End Sub

Private Sub cmdSaibara_Click()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum

        If CLng(txtSaibara.Text) > "65535" Then
            MsgBox "Error - Invalid price value which should be between 0 and 65535", vbCritical
            Exit Sub
        End If
        
        If listSaibara.Selected(0) = True Then
            PutWord sFilePath, &HFEDA8, txtSaibara.Text
        ElseIf listSaibara.Selected(1) = True Then
            PutWord sFilePath, &HFEDBC, txtSaibara.Text
        ElseIf listSaibara.Selected(2) = True Then
            PutWord sFilePath, &HFEDD0, txtSaibara.Text
        ElseIf listSaibara.Selected(3) = True Then
            PutWord sFilePath, &HFEDE4, txtSaibara.Text
        ElseIf listSaibara.Selected(4) = True Then
            PutWord sFilePath, &HFEDF8, txtSaibara.Text
        ElseIf listSaibara.Selected(5) = True Then
            PutWord sFilePath, &HFEE0C, txtSaibara.Text
        ElseIf listSaibara.Selected(6) = True Then
            PutWord sFilePath, &HFEE20, txtSaibara.Text
        ElseIf listSaibara.Selected(7) = True Then
            PutWord sFilePath, &HFEE34, txtSaibara.Text
        ElseIf listSaibara.Selected(8) = True Then
            PutWord sFilePath, &HFEE48, txtSaibara.Text
        ElseIf listSaibara.Selected(9) = True Then
            PutWord sFilePath, &HFEE5C, txtSaibara.Text
        End If
    MsgBox "Saibara's Blacksmith Data saved!", vbInformation
End Sub

Private Sub listSaibara_Click()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum

    If listSaibara.Selected(0) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEDA8 + 1)
    ElseIf listSaibara.Selected(1) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEDBC + 1)
    ElseIf listSaibara.Selected(2) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEDD0 + 1)
    ElseIf listSaibara.Selected(3) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEDE4 + 1)
    ElseIf listSaibara.Selected(4) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEDF8 + 1)
    ElseIf listSaibara.Selected(5) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEE0C + 1)
    ElseIf listSaibara.Selected(6) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEE20 + 1)
    ElseIf listSaibara.Selected(7) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEE34 + 1)
    ElseIf listSaibara.Selected(8) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEE48 + 1)
    ElseIf listSaibara.Selected(9) = True Then
        txtSaibara.Text = GetWord(sFilePath, &HFEE5C + 1)
    End If
End Sub

Private Sub txtPoutry_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
cmdPoutry.Enabled = True
End Sub

Private Sub txtMedals_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
cmdItems.Enabled = True
End Sub

Private Sub txtGotz_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub txtLumber_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
cmdGotz.Enabled = True
End Sub

Private Sub txtSaibara_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
cmdSaibara.Enabled = True
End Sub

Private Sub txtBarley_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
cmdBarley.Enabled = True
End Sub

Private Sub cmdItems_Click()
Dim pMedals As Integer

iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
    
    If CLng(txtMedals.Text) > "65535" Then
    MsgBox "Error - Invalid medal value which should be between 0 and 65535", vbCritical
    Exit Sub
    End If
    
    If listItems.Selected(0) = True Then
        Put #iFileNum, &HFB010 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(1) = True Then
        Put #iFileNum, &HFB024 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(2) = True Then
        Put #iFileNum, &HFB038 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(3) = True Then
        Put #iFileNum, &HFB04C + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(4) = True Then
        Put #iFileNum, &HFB060 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(5) = True Then
        Put #iFileNum, &HFB074 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(6) = True Then
        Put #iFileNum, &HFB088 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(7) = True Then
        Put #iFileNum, &HFB09C + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(8) = True Then
        Put #iFileNum, &HFB0B0 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(9) = True Then
        Put #iFileNum, &HFB0C4 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(10) = True Then
        Put #iFileNum, &HFB0D8 + 1, CLng(txtMedals.Text)
    ElseIf listItems.Selected(11) = True Then
        Put #iFileNum, &HFB0EC + 1, CLng(txtMedals.Text)
    End If
    
    MsgBox "Horse Race Prize Data saved!", vbInformation
End Sub

Private Sub listItems_Click()
Dim pMedals As Long

iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum

    If listItems.Selected(0) = True Then
        Get #iFileNum, &HFB010 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(1) = True Then
        Get #iFileNum, &HFB024 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(2) = True Then
        Get #iFileNum, &HFB038 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(3) = True Then
        Get #iFileNum, &HFB04C + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(4) = True Then
        Get #iFileNum, &HFB060 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(5) = True Then
        Get #iFileNum, &HFB074 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(6) = True Then
        Get #iFileNum, &HFB088 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(7) = True Then
        Get #iFileNum, &HFB09C + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(8) = True Then
        Get #iFileNum, &HFB0B0 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(9) = True Then
        Get #iFileNum, &HFB0C4 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(10) = True Then
        Get #iFileNum, &HFB0D8 + 1, pMedals
        txtMedals.Text = pMedals
    ElseIf listItems.Selected(11) = True Then
        Get #iFileNum, &HFB0EC + 1, pMedals
        txtMedals.Text = pMedals
    End If
End Sub

Private Sub Form_Load()
    sFilePath = frmMain.openfd.FileName
End Sub
