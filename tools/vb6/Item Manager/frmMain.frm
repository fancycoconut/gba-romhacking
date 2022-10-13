VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Manager"
   ClientHeight    =   4185
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   8400
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
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   29
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   2880
      Width           =   495
   End
   Begin VB.Frame fraROMInformation 
      Caption         =   "ROM Information"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   8175
      Begin VB.PictureBox picFlag 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7560
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   45
         Top             =   360
         Width           =   495
         Begin VB.Image imgFlag 
            Height          =   165
            Left            =   180
            Top             =   150
            Width           =   240
         End
         Begin VB.Shape shpFlag 
            BorderColor     =   &H00C0C0C0&
            Height          =   225
            Left            =   150
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   225
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame fraItems 
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.ListBox lstItems 
         Enabled         =   0   'False
         Height          =   2400
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraItemData 
      Caption         =   "Item Data"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   5655
      Begin VB.PictureBox picItemData 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   361
         TabIndex        =   8
         Top             =   240
         Width           =   5415
         Begin VB.TextBox txtMystery1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4635
            MaxLength       =   1
            TabIndex        =   30
            Top             =   840
            Width           =   255
         End
         Begin VB.ComboBox cmbPocket 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":000C
            Left            =   3840
            List            =   "frmMain.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtMystery2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   1
            TabIndex        =   23
            Top             =   840
            Width           =   255
         End
         Begin VB.TextBox txtSpecial2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4680
            MaxLength       =   2
            TabIndex        =   21
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtSpecial1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4680
            MaxLength       =   2
            TabIndex        =   19
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtDescription 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   16
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtPrice 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   14
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtIndexNumber 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   12
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtItemName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   14
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblPocket 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pocket"
            Height          =   195
            Left            =   3120
            TabIndex        =   24
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label lblMystery 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mystery"
            Height          =   195
            Left            =   3120
            TabIndex        =   22
            Top             =   960
            Width           =   585
         End
         Begin VB.Label lblSpecial2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Special2"
            Height          =   195
            Left            =   3120
            TabIndex        =   20
            Top             =   600
            Width           =   585
         End
         Begin VB.Label lblSpecial1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Special1"
            Height          =   195
            Left            =   3120
            TabIndex        =   18
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblDescriptionText 
            BackStyle       =   0  'Transparent
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   240
            TabIndex        =   17
            Top             =   1680
            Width           =   2805
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label lblPrice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Price"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   345
         End
         Begin VB.Label lblIndexNumber 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index Number"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   405
         End
      End
   End
   Begin VB.Frame fraItemData2 
      Caption         =   "Item Data"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2640
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.PictureBox picItemData2 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   361
         TabIndex        =   28
         Top             =   240
         Width           =   5415
         Begin VB.ComboBox cmbAttacks 
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtExtra 
            Height          =   285
            Left            =   4800
            MaxLength       =   2
            TabIndex        =   42
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtBattleUsageCode 
            Height          =   285
            Left            =   4320
            MaxLength       =   7
            TabIndex        =   40
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtBattleUsage 
            Height          =   285
            Left            =   4800
            MaxLength       =   1
            TabIndex        =   38
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtFieldUsage 
            Height          =   285
            Left            =   1560
            MaxLength       =   7
            TabIndex        =   36
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmMain.frx":0010
            Left            =   840
            List            =   "frmMain.frx":0023
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txtPokeBallIndex 
            Height          =   285
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   32
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblTMHMData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Attack"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lblExtra 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Extra"
            Height          =   195
            Left            =   3000
            TabIndex        =   41
            Top             =   960
            Width           =   390
         End
         Begin VB.Label lblBattleUsageCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            Height          =   195
            Left            =   3000
            TabIndex        =   39
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblBattleUsage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Battle Usage"
            Height          =   195
            Left            =   3000
            TabIndex        =   37
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblFieldUsage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field Usage"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblPokeBallIndex 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PokeBall Index"
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   1050
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open ROM"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save ROM"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
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
Private iFileNum As Integer
Private sFilePath As String
Private sHeader As String * 4

Private Items As Long
Private TMHMData As Long
Private ItemHeader As Long
Private AttackNames As Long
Private TMHMDataHeader As Long

Private Type ItemDataStructure
    ItemName(13) As Byte
    IndexNumber As Integer
    Price As Integer
    Special1 As Byte
    Special2 As Byte
    Description As Long
    Mystery1 As Byte
    Mystery2 As Byte
    Pocket As Byte
    Type As Byte
    FieldUsageCode As Long
    BattleUsage As Integer
    Filler As Integer
    BattleUsageCode As Long
    Extra As Integer
    Filler2 As Integer
End Type

Private ItemData As ItemDataStructure

Private Function CheckItemName(Textbox As Control) As String
If Len(Textbox.Text) = 14 Then
    CheckItemName = ""
Else
    CheckItemName = "\x"
End If
End Function

Private Sub LoadStuff(sFile As String)
Dim i As Integer
Dim iCounter As Integer
Dim bSapp2Asc() As Byte
Dim iItemCount As Integer
Dim DescriptionPointer As Long
    iFileNum = FreeFile
    Open sFile For Binary As #iFileNum
        ReDim bSapp2Asc(43) As Byte
        
        If Left(sHeader, 3) = "AXV" Then iItemCount = &H15C
        If Left(sHeader, 3) = "AXP" Then iItemCount = &H15C
        If Left(sHeader, 3) = "BPG" Then iItemCount = &H176
        If Left(sHeader, 3) = "BPR" Then iItemCount = &H176
        If Left(sHeader, 3) = "BPE" Then iItemCount = &H178
        
        lstItems.Clear
        Seek #iFileNum, Items + 1
        For i = 0 To iItemCount
            Get #iFileNum, , bSapp2Asc
            lstItems.AddItem Sapp2Asc(bSapp2Asc)
        Next i
        
        ReDim bSapp2Asc(12) As Byte
        
        cmbAttacks.Clear
        For i = 0 To 354 - 1
            Get #iFileNum, AttackNames + (i * 13) + 1, bSapp2Asc
            cmbAttacks.AddItem Sapp2Asc(bSapp2Asc)
        Next i
    Close #iFileNum
    
    cmbPocket.Clear ' Loading Pockets
    cmbPocket.AddItem "01 Misc", 0
    Select Case Left(sHeader, 3)
        Case "AXV", "AXP", "BPE"
            cmbPocket.AddItem "02 PokéBalls", 1
            cmbPocket.AddItem "03 TMs/HMs", 2
            cmbPocket.AddItem "04 Berries", 3
            cmbPocket.AddItem "05 Key Items", 4
        Case "BPR", "BPG"
            cmbPocket.AddItem "02 Key Items", 1
            cmbPocket.AddItem "03 PokéBalls", 2
            cmbPocket.AddItem "04 TMs/HMs", 3
            cmbPocket.AddItem "05 Berries", 4
    End Select
    Erase bSapp2Asc
End Sub

Private Sub cmdNavi_Click()
If cmdNavi.Caption = ">>" Then
    cmdNavi.Caption = "<<"
    fraItemData.Visible = False
    fraItemData2.Visible = True
Else
    cmdNavi.Caption = ">>"
    fraItemData.Visible = True
    fraItemData2.Visible = False
End If
End Sub

Private Sub cmdSave_Click()
    mnuSave_Click
End Sub

Private Sub Form_Load()
    SetIcon Me.hWnd, "AAA"
    imgFlag.Picture = LoadResPicture("NULL", 0)
End Sub

Private Sub InitTool()
Dim cnt As Control
    lblHeader.Caption = sHeader
    lblROM.ToolTipText = sFilePath

    For Each cnt In frmMain.Controls
        On Error Resume Next
        If cnt.Enabled = False Then cnt.Enabled = True
    Next
    
    lstItems.ListIndex = 0
End Sub

Private Sub lstItems_Click()
Dim AttackData As Integer
Dim arrTemp(255) As Byte
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, Items + (lstItems.ListIndex * 44) + 1, ItemData
        Get #iFileNum, (ItemData.Description - &H8000000) + 1, arrTemp
        If lstItems.ListIndex >= 289 Then
            Get #iFileNum, TMHMData + ((lstItems.ListIndex - 289) * 2) + 1, AttackData
        End If
    Close #iFileNum
    
    txtItemName.Text = Sapp2Asc(ItemData.ItemName)
    txtIndexNumber.Text = Hex(ItemData.IndexNumber)
    txtPrice.Text = ItemData.Price
    txtDescription.Text = Hex(ItemData.Description - &H8000000)
    lblDescriptionText.Caption = Replace(Sapp2Asc(arrTemp), "\n", vbCrLf)
    txtSpecial1.Text = ItemData.Special1
    txtSpecial2.Text = ItemData.Special2
    txtMystery1.Text = Hex(ItemData.Mystery1)
    txtMystery2.Text = Hex(ItemData.Mystery2)
    cmbPocket.ListIndex = ItemData.Pocket - 1
    txtFieldUsage.Text = Hex(ItemData.FieldUsageCode)
    txtBattleUsage.Text = Hex(ItemData.BattleUsage)
    txtBattleUsageCode.Text = Hex(ItemData.BattleUsageCode)
    txtExtra.Text = ItemData.Extra
    
    Select Case Left(sHeader, 3)
        Case "AXV", "AXP", "BPE"
            If cmbPocket.ListIndex = 1 Then ' Check for pokeball
                lblType.Visible = False
                lblPokeBallIndex.Visible = True
                cmbType.Visible = False
                txtPokeBallIndex.Visible = True
                txtPokeBallIndex.Text = Hex(ItemData.Type) ' Displays Pokeball index
            Else
                lblType.Visible = True
                lblPokeBallIndex.Visible = False
                cmbType.Visible = True
                txtPokeBallIndex.Visible = False
                cmbType.ListIndex = ItemData.Type ' Displays Item Type
            End If
            
            If cmbPocket.ListIndex = 2 Then ' Check for TM/HM
                cmbAttacks.Enabled = True
                cmbAttacks.ListIndex = AttackData - 1
            Else
                cmbAttacks.Enabled = False
            End If
            
        Case "BPR", "BPG"
            If cmbPocket.ListIndex = 2 Then ' Check for pokeball
                lblType.Visible = False
                lblPokeBallIndex.Visible = True
                cmbType.Visible = False
                txtPokeBallIndex.Visible = True
                txtPokeBallIndex.Text = Hex(ItemData.Type) ' Displays Pokeball index
            Else
                lblType.Visible = True
                lblPokeBallIndex.Visible = False
                cmbType.Visible = True
                txtPokeBallIndex.Visible = False
                cmbType.ListIndex = ItemData.Type ' Displays Item Type
            End If
            
            If cmbPocket.ListIndex = 3 Then ' Check for TM/HM
                cmbAttacks.Enabled = True
                cmbAttacks.ListIndex = AttackData - 1
            Else
                cmbAttacks.Enabled = False
            End If
    End Select

    Erase arrTemp
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String
Dim cdgOpen As clsCommonDialog
    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, "Open ROM" & "...", , "Gameboy Advance ROMs (*agb,*.gba,*.bin)|*.agb;*.gba;*.bin")
    If LenB(sResult) = 0 Then GoTo EndMe
    
    iFileNum = FreeFile
    Open sResult For Binary As #iFileNum
        Get #iFileNum, &HAC + 1, sHeader
        
        Select Case sHeader
            Case "AXVE"
                lblROM.Caption = "Pokémon Ruby"
                ItemHeader = &HA99E4
                AttackNames = &H1F832D
                TMHMDataHeader = &H6F038
            
            Case "AXPE"
                lblROM.Caption = "Pokémon Sapphire"
                ItemHeader = &HA99E4
                AttackNames = &H1F82BD
                TMHMDataHeader = &H6F03C
                
            Case "BPRE"
                lblROM.Caption = "Pokémon Fire Red"
                ItemHeader = &H1C8
                AttackNames = &H2470A1
                TMHMDataHeader = &H125A8C
                
            Case "BPGE"
                lblROM.Caption = "Pokémon Leaf Green"
                ItemHeader = &H1C8
                AttackNames = &H24707D
                TMHMDataHeader = &H125A64
                
            Case "BPEE"
                lblROM.Caption = "Pokémon Emerald"
                ItemHeader = &H1C8
                AttackNames = &H319789
            
            Case Else
                MsgBox "Error - Unsupported ROM :P", vbExclamation
                Unsupported
                GoTo EndMe
            
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
        
        Get #iFileNum, ItemHeader + 1, Items
        Get #iFileNum, TMHMDataHeader + 1, TMHMData
        
        Items = Items - &H8000000
        TMHMData = TMHMData - &H8000000
        If Left(sHeader, 3) = "BPE" Then TMHMData = &H615B94 ' No header for emerald :(
    Close #iFileNum
    sFilePath = sResult
    LoadStuff sFilePath
    
    InitTool
EndMe:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()
    Asc2Sapp txtItemName.Text & CheckItemName(txtItemName), ItemData.ItemName ' Saves the name to item structure
    lstItems.List(lstItems.ListIndex) = txtItemName.Text ' Updates the name in the list
    
    ItemData.IndexNumber = CInt("&H" & txtIndexNumber.Text)
    ItemData.Price = CInt(txtPrice.Text)
    ItemData.Description = CLng("&H" & txtDescription.Text) + &H8000000
    ItemData.Special1 = txtSpecial1.Text
    ItemData.Special2 = txtSpecial2.Text
    ItemData.Mystery1 = CInt("&H" & txtMystery1.Text)
    ItemData.Mystery2 = CInt("&H" & txtMystery2.Text)
    ItemData.Pocket = cmbPocket.ListIndex + 1
    ItemData.FieldUsageCode = CLng("&H" & txtFieldUsage.Text)
    ItemData.BattleUsage = CInt("&H" & txtBattleUsage.Text)
    ItemData.BattleUsageCode = CLng("&H" & txtBattleUsageCode.Text)
    ItemData.Extra = CInt(txtExtra.Text)
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Select Case Left(sHeader, 3)
            Case "AXV", "AXP", "BPE"
                If cmbPocket.ListIndex = 1 Then ' Check for pokeball
                    ItemData.Type = CInt("&H" & txtPokeBallIndex.Text)
                Else
                    ItemData.Type = cmbType.ListIndex
                End If
                
                If lstItems.ListIndex < 289 Then GoTo Continue
                If cmbPocket.ListIndex = 2 Then ' Check for TM/HM
                    Put #iFileNum, TMHMData + ((lstItems.ListIndex - 289) * 2) + 1, cmbAttacks.ListIndex + 1 ' Writes TM/HM data
                End If
            Case "BPR", "BPG"
                If cmbPocket.ListIndex = 2 Then ' Check for pokeball
                    ItemData.Type = CInt("&H" & txtPokeBallIndex.Text)
                Else
                    ItemData.Type = cmbType.ListIndex
                End If
                
                If lstItems.ListIndex < 289 And lstItems.ListIndex > 346 Then GoTo Continue
                If cmbPocket.ListIndex = 3 Then ' Check for TM/HM
                    Put #iFileNum, TMHMData + ((lstItems.ListIndex - 289) * 2) + 1, cmbAttacks.ListIndex + 1 ' Writes TM/HM data
                End If
        End Select

Continue:
        Put #iFileNum, Items + (lstItems.ListIndex * 44) + 1, ItemData
    Close #iFileNum
    lstItems_Click
End Sub

Private Sub Unsupported()
    lstItems.Clear
    lstItems.Enabled = False
    txtItemName.Enabled = False
    txtIndexNumber.Enabled = False
    txtPrice.Enabled = False
    txtDescription.Enabled = False
    txtSpecial1.Enabled = False
    txtSpecial2.Enabled = False
    txtMystery1.Enabled = False
    txtMystery2.Enabled = False
    cmbPocket.Enabled = False
    cmdSave.Enabled = False
    mnuSave.Enabled = False
    cmdNavi.Enabled = False
    fraItemData.Visible = True
    fraItemData2.Visible = False
    lblDescriptionText.Caption = ""
    imgFlag.Picture = LoadResPicture("NULL", 0)
End Sub
