VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trader Advanced"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   44
      Tag             =   "16"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   43
      Top             =   4440
      Width           =   495
   End
   Begin VB.Frame fraTrades 
      Caption         =   "Trades"
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Tag             =   "15"
      Top             =   1680
      Width           =   2535
      Begin VB.ListBox listTrades 
         Enabled         =   0   'False
         Height          =   1425
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame fraROMInfo 
      Caption         =   "ROM Information"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Tag             =   "12"
      Top             =   3840
      Width           =   2535
      Begin VB.PictureBox picFlag 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2040
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   49
         Top             =   480
         Width           =   375
         Begin VB.Image imgFlag 
            Height          =   165
            Left            =   60
            Top             =   30
            Width           =   240
         End
         Begin VB.Shape shpFlag 
            BorderColor     =   &H00C0C0C0&
            Height          =   225
            Left            =   30
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   255
         Width           =   225
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Tag             =   "14"
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Tag             =   "13"
         Top             =   255
         Width           =   465
      End
   End
   Begin VB.Frame fraCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   7935
      Begin VB.Label lblCopy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2009 ZodiacDaGreat"
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   2760
         TabIndex        =   1
         Top             =   180
         Width           =   2415
      End
   End
   Begin VB.Frame fraTradeData 
      Caption         =   "Trade Data"
      Height          =   3015
      Left            =   2760
      TabIndex        =   7
      Tag             =   "18"
      Top             =   1680
      Width           =   5295
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   5055
         TabIndex        =   8
         Top             =   240
         Width           =   5055
         Begin VB.PictureBox picRecPoke 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1000
            Left            =   3720
            ScaleHeight     =   67
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   67
            TabIndex        =   46
            Top             =   1200
            Width           =   1000
         End
         Begin VB.PictureBox picGivePoke 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1000
            Left            =   2040
            ScaleHeight     =   67
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   67
            TabIndex        =   45
            Top             =   1200
            Width           =   1000
         End
         Begin VB.TextBox txtOTID 
            Enabled         =   0   'False
            Height          =   320
            Left            =   3840
            MaxLength       =   5
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtPokeName 
            Enabled         =   0   'False
            Height          =   320
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtTrainerName 
            Enabled         =   0   'False
            Height          =   320
            Left            =   120
            MaxLength       =   7
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin VB.Frame fra1 
            Caption         =   "Trade"
            Height          =   1425
            Left            =   120
            TabIndex        =   15
            Tag             =   "6"
            Top             =   960
            Width           =   4815
            Begin VB.PictureBox pic2 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   240
               ScaleHeight     =   975
               ScaleWidth      =   1455
               TabIndex        =   16
               Top             =   360
               Width           =   1455
               Begin VB.ComboBox cmbGivePoke 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "frmMain.frx":000C
                  Left            =   0
                  List            =   "frmMain.frx":04E4
                  TabIndex        =   47
                  Top             =   0
                  Width           =   1455
               End
               Begin VB.ComboBox cmbRecPoke 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "frmMain.frx":13F3
                  Left            =   0
                  List            =   "frmMain.frx":18CB
                  TabIndex        =   18
                  Top             =   600
                  Width           =   1455
               End
               Begin VB.Label lbl3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "For"
                  Height          =   195
                  Left            =   600
                  TabIndex        =   17
                  Tag             =   "17"
                  Top             =   360
                  Width           =   240
               End
            End
            Begin VB.Label lbl4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "For"
               Height          =   195
               Left            =   3120
               TabIndex        =   48
               Tag             =   "17"
               Top             =   600
               Visible         =   0   'False
               Width           =   240
            End
         End
         Begin VB.Label lblOTID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trainer ID"
            Height          =   195
            Left            =   3840
            TabIndex        =   13
            Tag             =   "22"
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lblPokeName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pokemon Name"
            Height          =   195
            Left            =   1920
            TabIndex        =   11
            Tag             =   "20"
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblTrainerName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trainer Name"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Tag             =   "19"
            Top             =   120
            Width           =   960
         End
      End
   End
   Begin VB.Frame fraTradeData2 
      Caption         =   "Trade Data"
      Height          =   3015
      Left            =   2760
      TabIndex        =   21
      Tag             =   "18"
      Top             =   1680
      Visible         =   0   'False
      Width           =   5295
      Begin VB.PictureBox pic3 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   5055
         TabIndex        =   22
         Top             =   240
         Width           =   5055
         Begin VB.ComboBox cmbItems 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   39
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cmbNatures 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":27DA
            Left            =   120
            List            =   "frmMain.frx":27DC
            TabIndex        =   38
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox cmbAbility 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":27DE
            Left            =   3480
            List            =   "frmMain.frx":27E8
            TabIndex        =   37
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox cmbAbilities 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":27F6
            Left            =   3480
            List            =   "frmMain.frx":27F8
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   720
            Width           =   1575
         End
         Begin VB.Frame fraStats 
            Caption         =   "Increase Stats (IVs)"
            Height          =   1335
            Left            =   120
            TabIndex        =   23
            Tag             =   "39"
            Top             =   1080
            Width           =   4815
            Begin VB.TextBox txtIV 
               Height          =   285
               Index           =   5
               Left            =   720
               MaxLength       =   3
               TabIndex        =   35
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox txtIV 
               Height          =   285
               Index           =   4
               Left            =   4200
               MaxLength       =   3
               TabIndex        =   34
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtIV 
               Height          =   285
               Index           =   3
               Left            =   3120
               MaxLength       =   3
               TabIndex        =   33
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtIV 
               Height          =   285
               Index           =   2
               Left            =   2160
               MaxLength       =   3
               TabIndex        =   32
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtIV 
               Height          =   285
               Index           =   1
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   31
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtIV 
               Height          =   285
               Index           =   0
               Left            =   480
               MaxLength       =   3
               TabIndex        =   30
               Top             =   360
               Width           =   375
            End
            Begin VB.Label lblStats 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HP"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   29
               Top             =   360
               Width           =   195
            End
            Begin VB.Label lblStats 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AT"
               Height          =   195
               Index           =   1
               Left            =   1080
               TabIndex        =   28
               Top             =   360
               Width           =   195
            End
            Begin VB.Label lblStats 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DF"
               Height          =   195
               Index           =   2
               Left            =   1920
               TabIndex        =   27
               Top             =   360
               Width           =   195
            End
            Begin VB.Label lblStats 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SPD"
               Height          =   195
               Index           =   3
               Left            =   2760
               TabIndex        =   26
               Top             =   360
               Width           =   285
            End
            Begin VB.Label lblStats 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SP.DF"
               Height          =   195
               Index           =   5
               Left            =   240
               TabIndex        =   25
               Top             =   840
               Width           =   435
            End
            Begin VB.Label lblStats 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SP.AT"
               Height          =   195
               Index           =   4
               Left            =   3720
               TabIndex        =   24
               Top             =   360
               Width           =   435
            End
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Held Item"
            Height          =   195
            Left            =   1680
            TabIndex        =   42
            Tag             =   "21"
            Top             =   120
            Width           =   690
         End
         Begin VB.Label lblNatures 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nature"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Tag             =   "23"
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblAbility 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ability"
            Height          =   195
            Left            =   3480
            TabIndex        =   40
            Tag             =   "38"
            Top             =   120
            Width           =   435
         End
      End
   End
   Begin VB.Image imgBanner 
      Height          =   1575
      Left            =   0
      Picture         =   "frmMain.frx":27FA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   1
      Begin VB.Menu mnuOpen 
         Caption         =   "Open ROM"
         HelpContextID   =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save ROM"
         Enabled         =   0   'False
         HelpContextID   =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         HelpContextID   =   4
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTrade 
      Caption         =   "&Trade"
      HelpContextID   =   5
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Enabled         =   0   'False
         HelpContextID   =   7
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         HelpContextID   =   8
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepoint 
         Caption         =   "Edit Trade Amount"
         Enabled         =   0   'False
         HelpContextID   =   9
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   10
      Begin VB.Menu mnuReadme 
         Caption         =   "Readme"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         HelpContextID   =   11
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TradeStructure
    Nickname(9) As Byte
    FillerWord As Integer
    PokemonOffered As Integer
    HP As Byte
    Attack As Byte
    Defense As Byte
    Speed As Byte
    SPAttack As Byte
    SPDefense As Byte
    Ability As Byte
    Filler As Byte
    FillerWord2 As Integer
    OTIDFirst As Byte
    OTIDSecond As Byte
    FillerWord3 As Integer
    UnknownWord(1) As Byte
    UnknownWord2(1) As Byte
    UnknownWord3(1) As Byte
    FillerWord4 As Integer
    Personality As Long
    Item As Integer
    SeparationByte As Byte
    TrainerName(7) As Byte
    FillerWord5(2) As Byte
    UnknownByte As Byte
    Sheen As Byte ' Trade determination byte
    PokemonWanted As Integer
    FillerWord6 As Integer
End Type

Private TradeData As TradeStructure

Private Function CheckPokeName(Textbox As Control) As String
If Len(Textbox.Text) = 10 Then
    CheckPokeName = ""
Else
    CheckPokeName = "\x"
End If
End Function

Private Sub cmbGivePoke_Click()
    LoadPokeImage cmbGivePoke.ListIndex, picGivePoke
End Sub

Private Sub cmbRecPoke_Click()
    LoadPokeImage cmbRecPoke.ListIndex, picRecPoke
End Sub

Private Sub cmdNavi_Click()
    If cmdNavi.Caption = ">>" Then
        fraTradeData.Visible = False
        fraTradeData2.Visible = True
        cmdNavi.Caption = "<<"
    Else
        fraTradeData.Visible = True
        fraTradeData2.Visible = False
        cmdNavi.Caption = ">>"
    End If
End Sub

Private Sub cmdSave_Click()
    mnuSave_Click
End Sub

Private Function FixTrades(Cnt As Control)
Dim Temp As Long
Dim arrTemp() As Byte
Dim PointerBank As Long

    Temp = Val(InputBox(LoadResString(42) & vbNewLine & LoadResString(43), App.Title, "&H"))
    If LenB(Temp) = 0 Then Exit Function
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        arrTemp = LoadResData("TRADEDATA", 100)
        
        ' Fixing header first
        Select Case sHeader
            Case "AXVE", "AXPE"
                Put #iFileNum, &H4C284 + 1, Temp + &H8000000
                Put #iFileNum, &H4D8D4 + 1, Temp + &H8000000
                Put #iFileNum, &H4D930 + 1, Temp + &H8000000
                Put #iFileNum, &H4DAA4 + 1, Temp + &H8000000
            Case "BPRE", "BPGE"
                Put #iFileNum, &H50EFC + 1, Temp + &H8000000
                Put #iFileNum, &H53AD4 + 1, Temp + &H8000000
                Put #iFileNum, &H53B30 + 1, Temp + &H8000000
                Put #iFileNum, &H53CA4 + 1, Temp + &H8000000
            Case "BPEE"
                Put #iFileNum, &H7BBB0 + 1, Temp + &H8000000
                Put #iFileNum, &H7E774 + 1, Temp + &H8000000
                Put #iFileNum, &H7E7D0 + 1, Temp + &H8000000
                Put #iFileNum, &H7E944 + 1, Temp + &H8000000
            Case "BPEF"
                Put #iFileNum, &H7BBAC + 1, Temp + &H8000000
                Put #iFileNum, &H7E770 + 1, Temp + &H8000000
                Put #iFileNum, &H7E7CC + 1, Temp + &H8000000
                Put #iFileNum, &H7E940 + 1, Temp + &H8000000
        End Select
    Close #iFileNum
    
    ' Inserting only 3 trades
    WriteByteArray sFilePath, arrTemp, Temp
    WriteByteArray sFilePath, arrTemp, Temp + 60
    WriteByteArray sFilePath, arrTemp, Temp + 120
    
    ' Enable controls
    For Each Cnt In frmMain.Controls
        On Error Resume Next
        If Cnt.Enabled = False Then Cnt.Enabled = True
    Next
    
    Erase arrTemp
    GetTradeAmount
End Function

Private Sub Form_Load()
    Localize Me
    SetIcon Me.hWnd, "AAA"
    imgFlag.Picture = LoadResPicture("NULL", 0)
End Sub

Private Function GetNames()
Dim TempPointer As Long
Dim bSapp2Asc() As Byte
Dim i As Integer
    cmbAbilities.Clear
    cmbItems.Clear
    cmbNatures.Clear
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        ReDim bSapp2Asc(43) As Byte ' Getting Item Names
        Seek #iFileNum, ItemNames + 1
        For i = 0 To iItemCount
            Get #iFileNum, , bSapp2Asc
            cmbItems.AddItem Sapp2Asc(bSapp2Asc)
        Next i
        
        ReDim bSapp2Asc(12) As Byte ' Getting Abilities
        Seek #iFileNum, AbilityNames + 1
        For i = 0 To 77
            Get #iFileNum, , bSapp2Asc
            cmbAbilities.AddItem Sapp2Asc(bSapp2Asc)
        Next i
        
        ReDim bSapp2Asc(7) As Byte ' Getting Natures
        For i = 0 To 24
            Get #iFileNum, NatureNames + 1, TempPointer
            TempPointer = TempPointer - &H8000000
            Get #iFileNum, TempPointer + 1, bSapp2Asc
            cmbNatures.AddItem Sapp2Asc(bSapp2Asc)
            NatureNames = NatureNames + 4
        Next i
    Erase bSapp2Asc
    Close #iFileNum
End Function

Public Function GetTradeAmount()
Dim bTemp As Byte
Dim TradeHeader As Long
Dim iIncValue As Integer
    listTrades.Clear
    iTradeAmount = 0
    iIncValue = 0
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Select Case sHeader
            Case "AXVE", "AXPE"
                TradeHeader = &H4C284
            Case "BPRE", "BPGE"
                TradeHeader = &H50EFC
            Case "BPEE"
                TradeHeader = &H7BBB0
            Case "BPEF"
                TradeHeader = &H7BBAC
        End Select
        
        Get #iFileNum, TradeHeader + 1, TradeDataOffset
        TradeDataOffset = TradeDataOffset - &H8000000
        
        Do
            Get #iFileNum, TradeDataOffset + 1 + iIncValue + 55, bTemp
            If bTemp <> &HA Then Exit Do
            iIncValue = iIncValue + 60
            iTradeAmount = iTradeAmount + 1
            listTrades.AddItem LoadResString(6) & " #" & Right("00" & Hex(iTradeAmount), 2)
        Loop
    Close #iFileNum
End Function

Private Function GetWord(sFilePath As String, lOffset As Long) As Long
Dim iFileNum As Integer
Dim bFirstByte As Byte
Dim bSecondByte As Byte
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, lOffset, bFirstByte
        Get #iFileNum, lOffset + 1, bSecondByte
    Close #iFileNum
    
    GetWord = CLng("&H" & Hex$(bSecondByte) & Right$("0" & Hex$(bFirstByte), 2))
End Function

Private Sub listTrades_Click()
Dim bTemp As Byte
Dim lTemp As Long
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, TradeDataOffset + 1 + (60 * listTrades.ListIndex), TradeData
        
        txtPokeName.Text = Sapp2Asc(TradeData.Nickname)
        cmbRecPoke.ListIndex = TradeData.PokemonOffered
        
        txtIV(0).Text = TradeData.HP
        txtIV(1).Text = TradeData.Attack
        txtIV(2).Text = TradeData.Defense
        txtIV(3).Text = TradeData.Speed
        txtIV(4).Text = TradeData.SPAttack
        txtIV(5).Text = TradeData.SPDefense
        
        cmbAbility.ListIndex = TradeData.Ability
        lTemp = cmbRecPoke.ListIndex * &H1C
        If TradeData.Ability = 0 Then
            Get #iFileNum, PokemonStats + 1 + 22 + lTemp, bTemp
            cmbAbilities.ListIndex = bTemp
        Else
            Get #iFileNum, PokemonStats + 1 + 23 + lTemp, bTemp
            cmbAbilities.ListIndex = bTemp
        End If
        
        txtOTID.Text = CLng("&H" & Hex$(TradeData.OTIDSecond) & Right$("0" & Hex$(TradeData.OTIDFirst), 2))
        cmbNatures.ListIndex = TradeData.Personality Mod 25
        cmbItems.ListIndex = TradeData.Item
        txtTrainerName.Text = Sapp2Asc(TradeData.TrainerName)
        cmbGivePoke.ListIndex = TradeData.PokemonWanted
    Close #iFileNum
End Sub

Public Sub LoadPokeImage(ByVal Index As Integer, PictureBox As Control)
Dim i As Integer
Dim arrGFX() As Byte
Dim iFile As Integer
Dim TempOffset As Long
Dim SomeCounter As Long
Dim arrPalette() As Byte
Dim arrTemp(32192) As Byte
Dim arrPal(0 To 256) As Integer
Dim PCPalette(0 To 16, 0 To 16) As Long
    iFile = FreeFile
    Open sFilePath For Binary As #iFile
        Get #iFile, PokemonPics + 1 + (8 * Index), TempOffset ' Getting pointer to graphics data
        Get #iFile, TempOffset - &H8000000 + 1, arrTemp ' Getting graphics data to an array
        LZ77UnComp arrTemp, arrGFX ' Uncompressing graphics
        Erase arrTemp
        
        Get #iFile, PokemonPalettes + 1 + (8 * Index), TempOffset ' Getting pointer to palette data
        Get #iFile, TempOffset - &H8000000 + 1, arrTemp ' Getting palette data to an array
        LZ77UnComp arrTemp, arrPalette ' Uncompressing palettes
        
        SomeCounter = 0 ' Converting the palette array into a 2D array
        For i = 0 To 15 ' 16 Colors ^^
            arrPal(i) = CInt(arrPalette(SomeCounter + 1) * &H100 + arrPalette(SomeCounter))
            SomeCounter = SomeCounter + 2
        Next i
    
        UnPackPalette arrPal, PCPalette ' Unpacking palette into a displayable form
        
        PictureBox.Cls
        For i = 0 To 63 ' ploting image - 0 to 64 since its height and width is 64
            DrawTile8 PictureBox.hDC, i, arrGFX, PCPalette
        Next i
    
    Close #iFile
    Erase arrGFX, arrPalette, arrTemp, arrPal, PCPalette
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuExport_Click()
Dim arrTemp() As Byte
Dim cdgExport As clsCommonDialog
Dim sExport As String
    Set cdgExport = New clsCommonDialog
    sExport = cdgExport.ShowSave(Me.hWnd, "Save As", "Trade " & listTrades.ListIndex + 1, , "Trader Advanced Files (*.taf)|*.taf", OVERWRITEPROMPT)
    If LenB(sExport) = 0 Then GoTo EndMe
    
    ReadByteArray sFilePath, arrTemp, TradeDataOffset + (60 * listTrades.ListIndex), 60
    WriteByteArray sExport, arrTemp, 0
    
    Erase arrTemp
EndMe:
    Set cdgExport = Nothing
End Sub

Private Sub mnuImport_Click()
Dim arrTemp() As Byte
Dim cdgImport As clsCommonDialog
Dim sImport As String
    Set cdgImport = New clsCommonDialog
    sImport = cdgImport.ShowOpen(Me.hWnd, "Open", , "Trader Advanced Files (*.taf))|*.taf")
    If LenB(sImport) = 0 Then GoTo EndMe
    
    ReadByteArray sImport, arrTemp, 0, 60
    WriteByteArray sFilePath, arrTemp, TradeDataOffset + (60 * listTrades.ListIndex)
    
    Erase arrTemp
    listTrades_Click
EndMe:
    Set cdgImport = Nothing
End Sub

Private Sub mnuOpen_Click()
Dim ctl As Control
Dim sResult As String
Dim cdgOpen As clsCommonDialog

    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, LoadResString(2) & "...", , "GameBoy Advance ROMs (*.gba,*.agb,*.bin)|*.gba;*.agb;*.bin")
    If LenB(sResult) = 0 Then GoTo EndMe
    sFilePath = sResult
        
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, &HAC + 1, sHeader
    
        If LOF(iFileNum) > 16777216 Then
            EnlargedROM = 1
        Else
            EnlargedROM = 0
        End If
    Close #iFileNum
    
    If OpenROM(sHeader) = False Then GoTo UnsupportedROM
    
    GetNames ' Loading control contents
    GetTradeAmount
    
    On Error GoTo DisplayMsg
    listTrades.ListIndex = 0
    
    For Each ctl In frmMain.Controls  ' New way of enabling disabled stuff in one go :)
        On Error Resume Next
        If ctl.Enabled = False Then ctl.Enabled = True
    Next
    lbl4.Visible = True
    
    GoTo EndMe

UnsupportedROM:
    listTrades.Clear
    cmbGivePoke.Enabled = False
    cmbRecPoke.Enabled = False
    fraTradeData.Visible = True
    fraTradeData2.Visible = False
    cmdNavi.Enabled = False
    listTrades.Enabled = False
    cmdSave.Enabled = False
    mnuSave.Enabled = False
    mnuImport.Enabled = False
    mnuExport.Enabled = False
    mnuRepoint.Enabled = False
    lblROM.Caption = "???"
    lblHeader.Caption = "???"
    sHeader = vbNullString
    sFilePath = vbNullString
    imgFlag.Picture = LoadResPicture("NULL", 0)
    GoTo EndMe
    
DisplayMsg:
    If MsgBox(LoadResString(40) & vbNewLine & LoadResString(41), vbYesNo) = vbYes Then FixTrades ctl
    
EndMe:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub
    
Private Sub mnuReadme_Click()
Dim arrReadme() As Byte
    arrReadme = LoadResData("README", 100)
    WriteByteArray App.Path & "\Readme.txt", arrReadme, 0
    Shell "notepad.exe " & App.Path & "\Readme.txt", vbNormalFocus
    Kill App.Path & "\Readme.txt"
    Erase arrReadme
End Sub

Private Sub mnuRepoint_Click()
    frmEditAmount.Show vbModal, Me
End Sub

Private Sub mnuSave_Click()
Dim Index As Integer
    For Index = 0 To 5
        If LenB(txtIV(Index)) = 0 Then Exit Sub
        If txtIV(Index) > 255 Then Exit Sub ' &HFF
    Next Index
    
    If cmbGivePoke.ListIndex < 0 Then Exit Sub ' Fix for keypress on the combo boxes
    If cmbRecPoke.ListIndex < 0 Then Exit Sub ' So 0xFFFF won't be written
    If cmbNatures.ListIndex < 0 Then Exit Sub
    If cmbItems.ListIndex < 0 Then Exit Sub
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        ' Assigning changed values to private type
        Erase TradeData.Nickname
        Asc2Sapp txtPokeName.Text & CheckPokeName(txtPokeName), TradeData.Nickname
        TradeData.PokemonOffered = cmbRecPoke.ListIndex
        
        TradeData.HP = txtIV(0).Text
        TradeData.Attack = txtIV(1).Text
        TradeData.Defense = txtIV(2).Text
        TradeData.Speed = txtIV(3).Text
        TradeData.SPAttack = txtIV(4).Text
        TradeData.SPDefense = txtIV(5).Text
        
        TradeData.Ability = cmbAbility.ListIndex
        TradeData.Personality = (TradeData.Personality \ 25) * 25 + cmbNatures.ListIndex
        TradeData.Item = cmbItems.ListIndex
        
        Erase TradeData.TrainerName
        Asc2Sapp txtTrainerName.Text & CheckPokeName(txtTrainerName), TradeData.TrainerName
        TradeData.PokemonWanted = cmbGivePoke.ListIndex
        
        Put #iFileNum, TradeDataOffset + 1 + (60 * listTrades.ListIndex), TradeData ' Writing the trade data to ROM
        PutWord sFilePath, TradeDataOffset + 24 + (60 * listTrades.ListIndex), txtOTID.Text ' Writing the OTID to ROM
    Close #iFileNum
    listTrades_Click
End Sub

Private Function OpenROM(sHeader As String) As Boolean
    Select Case sHeader
        Case "AXVE"
            iItemCount = &H15C
            ItemNames = &H3C5564
            NatureNames = &H3C1004
            AbilityNames = &H1FA248
            PokemonStats = &H1FEC18
            PokemonPics = &H1E8354
            PokemonPalettes = &H1EA5B4
            lblROM.Caption = "Ruby Version"
            
        Case "AXPE"
            iItemCount = &H15C
            ItemNames = &H3C55BC
            NatureNames = &H3C105C
            AbilityNames = &H1FA1D8
            PokemonStats = &H1FEBA8
            PokemonPics = &H1E82E4
            PokemonPalettes = &H1EA544
            lblROM.Caption = "Sapphire Version"
            
        Case "BPRE"
            iItemCount = &H176
            ItemNames = &H3DB028
            NatureNames = &H463E60
            AbilityNames = &H24FC40
            PokemonStats = &H254784
            PokemonPics = &H2350AC
            PokemonPalettes = &H23730C
            lblROM.Caption = "Fire Red Version"
            
        Case "BPGE"
            iItemCount = &H176
            ItemNames = &H3DAE64
            NatureNames = &H463880
            AbilityNames = &H24FC1C
            PokemonStats = &H254760
            PokemonPics = &H235088
            PokemonPalettes = &H2372E8
            lblROM.Caption = "Leaf Green Version"
            
        Case "BPEE"
            iItemCount = &H178
            ItemNames = &H5839A0
            NatureNames = &H61CB50
            AbilityNames = &H31B6DB
            PokemonStats = &H3203CC
            PokemonPics = &H30A18C
            PokemonPalettes = &H303678
            lblROM.Caption = "Emerald Version"
            
        Case "BPEF"
            iItemCount = &H178
            ItemNames = &H587D6C
            NatureNames = &H620F54
            AbilityNames = &H32324E
            PokemonStats = &H327F3C
            PokemonPics = &H311CBC
            PokemonPalettes = &H30B1A8
            lblROM.Caption = "Emeraude Version"
            
        Case Else
            MsgBox LoadResString(35) & vbNewLine & LoadResString(36) & vbNewLine & LoadResString(37), vbExclamation
            OpenROM = False
            Exit Function
            
    End Select
    
    Select Case Right(sHeader, 1)
        Case "D"
            imgFlag.Picture = LoadResPicture("GERMAN", 0)
        Case "E"
            imgFlag.Picture = LoadResPicture("US", 0)
        Case "F"
            imgFlag.Picture = LoadResPicture("FRANCE", 0)
        Case "I"
            imgFlag.Picture = LoadResPicture("ITALY", 0)
        Case "J"
            imgFlag.Picture = LoadResPicture("JAPAN", 0)
        Case "S"
            imgFlag.Picture = LoadResPicture("SPAIN", 0)
    End Select
    
    lblHeader.Caption = sHeader
    lblROM.ToolTipText = sFilePath
    
    OpenROM = True
End Function

Public Sub PutWord(sFilePath As String, lOffset As Long, sValue As String)
Dim iFileNum As Integer
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Put #iFileNum, lOffset + 1, CInt("&H" & Hex$(sValue))
    Close #iFileNum
End Sub

Private Sub txtIV_Change(Index As Integer)
    If Val(txtIV(Index)) > 255 Then txtIV(Index).Text = 255
End Sub

Private Sub txtOTID_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
