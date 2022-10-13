Attribute VB_Name = "modFunctions"
Option Explicit

Public Function CheckROMLanguage(sHeader As String, Label As Control)
Select Case Right$(sHeader, 1)
    Case "D"
        Label.Caption = LoadResString(42)
    Case "E"
        Label.Caption = LoadResString(43)
    Case "F"
        Label.Caption = LoadResString(44)
    Case "I"
        Label.Caption = LoadResString(45)
    Case "J"
        Label.Caption = LoadResString(46)
    Case "S"
        Label.Caption = LoadResString(47)
End Select
End Function

Public Sub EditArray(ByRef arrArray() As Byte, Position As Long, Value As Long, Length As Long)
Dim i As Integer
    For i = 0 To Length - 1
        arrArray(Position + i) = CByte(Int(Value / (2 ^ (8 * i))) And 255)
    Next i
End Sub

Public Function LeftShift(ByVal Value As Long, ByVal iShift As Integer)
    LeftShift = Value * (2 ^ iShift)
End Function

Public Function LoadPalette(sHeader As String)
Dim i As Integer
    Select Case Left$(sHeader, 3) ' Palettes are stored in resource file ;)
        Case "AXV", "AXP"
            arrPalette = LoadResData("AXVE", "PAL")
            
        Case "BPR", "BPG"
            arrPalette = LoadResData("BPRE", "PAL")
            
        Case "BPE"
            arrPalette = LoadResData("BPEE", "PAL")
            
        Case Else ' Default palette when no ROM's loaded
            arrPalette = LoadResData("BPEE", "PAL")
    End Select
    
    frmMain.palOBJ.ShowSpriteColors = True
    frmMain.palBG.ShowBackgroundColors = True
    
    If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
        For i = 0 To 12
            If frmMain.chkPal(i).Value = 1 Then frmMain.chkPal(i).Value = 0
            frmMain.chkPal(i).Value = 1
        Next i
    Else
        For i = 0 To 11
            If frmMain.chkPal(i).Value = 1 Then frmMain.chkPal(i).Value = 0
            frmMain.chkPal(i).Value = 1
        Next i
    End If
    
    For i = 0 To 15
        If frmMain.chkPalOBJ(i).Value = 1 Then frmMain.chkPalOBJ(i).Value = 0
        frmMain.chkPalOBJ(i).Value = 1
    Next i
    
    Erase arrPalette
End Function

Public Function OpenROM(sHeader As String) As Boolean
    Select Case sHeader
        Case "AXVD"
            frmMain.lblROM.Caption = "POKÈMON RUBIN"
            SWI01Pos = &H386
            Hook1 = &H5461A
            Hook2 = &H73EB4
            RTC = &H300404A
            BattleFlag = &H3004DCB
            MenuFlag = &H202000E
            IndoorFlag = &H202E83F
            RTCReturnAddress = &H8054633
            PaletteOriginal = &H202EEC8
            DMA3ReturnAddress = &H8073EC7
            NintendoFix = 0
            
        Case "AXVE"
            frmMain.lblROM.Caption = "POKÈMON RUBY"
            SWI01Pos = &H252
            Hook1 = &H542DA
            Hook2 = &H73AF4
            RTC = &H300403A
            BattleFlag = &H3004DCB
            MenuFlag = &H202000E
            IndoorFlag = &H202E83F
            RTCReturnAddress = &H80542F3
            PaletteOriginal = &H202EEC8
            DMA3ReturnAddress = &H8073B07
            NintendoFix = 0
        
        Case "AXVS"
            frmMain.lblROM.Caption = "POKÈMON RUBÕ"
            SWI01Pos = &H386
            Hook1 = &H54716
            Hook2 = &H73FB0
            RTC = &H300403A
            BattleFlag = &H3004DCB
            MenuFlag = &H202000E
            IndoorFlag = &H202E83F
            RTCReturnAddress = &H805472F
            PaletteOriginal = &H202EEC8
            DMA3ReturnAddress = &H8073FC3
            NintendoFix = 0
        
        Case "AXPE"
            frmMain.lblROM.Caption = "POKÈMON SAPPHIRE"
            SWI01Pos = &H252
            Hook1 = &H542DE
            Hook2 = &H73AF8
            RTC = &H300403A
            BattleFlag = &H3004DCB
            MenuFlag = &H202000E
            IndoorFlag = &H202E83F
            RTCReturnAddress = &H80542F7
            PaletteOriginal = &H202EEC8
            DMA3ReturnAddress = &H8073B0B
            NintendoFix = 0
            
        Case "AXPF"
            frmMain.lblROM.Caption = "POKÈMON SAPHIR"
            SWI01Pos = &H386
            Hook1 = &H5470A
            Hook2 = &H73FA8
            RTC = &H300403A
            BattleFlag = &H3004DCB
            MenuFlag = &H202000E
            IndoorFlag = &H202E83F
            RTCReturnAddress = &H8054723
            PaletteOriginal = &H202EEC8
            DMA3ReturnAddress = &H8073FBB
            NintendoFix = 0
            
        Case "AXPS"
            frmMain.lblROM.Caption = "POKÈMON ZAFIRO"
            SWI01Pos = &H386
            Hook1 = &H5471A
            Hook2 = &H73FB4
            RTC = &H300403A
            BattleFlag = &H3004DCB
            MenuFlag = &H202000E
            IndoorFlag = &H202E83F
            RTCReturnAddress = &H8054733
            PaletteOriginal = &H202EEC8
            DMA3ReturnAddress = &H8073FC7
            NintendoFix = 0
        
        Case "BPED"
            frmMain.lblROM.Caption = "POKÈMON SMARAGD"
            SWI01Pos = &H3AA
            Hook1 = &H876AE
            Hook2 = &HA19F0
            RTC = &H3005CFA
            BattleFlag = &H20244D0
            MenuFlag = &H2020006
            IndoorFlag = &H203732F
            RTCReturnAddress = &H80876D5
            PaletteOriginal = &H2037B14
            DMA3ReturnAddress = &H80A1A03
            NintendoFix = &H16C986
            
        Case "BPEE"
            frmMain.lblROM.Caption = "POKÈMON EMERALD"
            SWI01Pos = &H3AA
            Hook1 = &H87692
            Hook2 = &HA19D4
            RTC = &H3005CFA
            BattleFlag = &H20244D0
            MenuFlag = &H2020006
            IndoorFlag = &H203732F
            RTCReturnAddress = &H80876B9
            PaletteOriginal = &H2037B14
            DMA3ReturnAddress = &H80A19E7
            NintendoFix = &H16CD1E
    
        Case "BPEF"
            frmMain.lblROM.Caption = "POKÈMON EMERAUDE"
            SWI01Pos = &H3AA
            Hook1 = &H876A2
            Hook2 = &HA19E8
            RTC = &H3005CFA
            BattleFlag = &H20244D0
            MenuFlag = &H2020006
            IndoorFlag = &H203732F
            RTCReturnAddress = &H80876C9
            PaletteOriginal = &H2037B14
            DMA3ReturnAddress = &H80A19FB
            NintendoFix = &H16CAAE
    
        Case "BPEJ"
            frmMain.lblROM.Caption = "POKÈMON EMERALD"
            SWI01Pos = &H3AA
            Hook1 = &H86FF6
            Hook2 = &HA129C
            RTC = &H3005A52
            BattleFlag = &H20244D0
            MenuFlag = &H2020006
            IndoorFlag = &H2036FCF
            RTCReturnAddress = &H808701D
            PaletteOriginal = &H20377B4
            DMA3ReturnAddress = &H80A12AF
            NintendoFix = &H16CB2E
        
        Case "BPRD"
            frmMain.lblROM.Caption = "POKÈMON FEUERROTE"
            SWI01Pos = &H3AA
            HookRTC = &H428
            Hook2 = &H703EC
            BattleFlag = &H200E724
            MenuFlag = &H2020655
            IndoorFlag = &H2036E13
            PaletteOriginal = &H20375F8
            DMA3ReturnAddress = &H807049B
            NintendoFix = 0
            
        Case "BPRE"
            frmMain.lblROM.Caption = "POKÈMON FIRE RED"
            SWI01Pos = &H3AA
            HookRTC = &H41E
            Hook2 = &H70488
            BattleFlag = &H200E724
            MenuFlag = &H2020655
            IndoorFlag = &H2036E13
            PaletteOriginal = &H20375F8
            DMA3ReturnAddress = &H807049B
            NintendoFix = 0
        
        Case "BPRI"
            frmMain.lblROM.Caption = "POKÈMON FIRE RED"
            SWI01Pos = &H3AA
            HookRTC = &H41E
            Hook2 = &H70488
            BattleFlag = &H200E724
            MenuFlag = &H2020655
            IndoorFlag = &H2036E13
            PaletteOriginal = &H20375F8
            DMA3ReturnAddress = &H807049B
            NintendoFix = 0
            
        Case "BPGE"
            frmMain.lblROM.Caption = "POKÈMON LEAF GREEN"
            SWI01Pos = &H3AA
            HookRTC = &H41E
            Hook2 = &H70488
            BattleFlag = &H200E724
            MenuFlag = &H2020655
            IndoorFlag = &H2036E13
            PaletteOriginal = &H20375F8
            DMA3ReturnAddress = &H807049B
            NintendoFix = 0
            
        Case Else
            MsgBox LoadResString(34), vbExclamation
            OpenROM = False
            Exit Function
            
    End Select
    OpenROM = True
End Function

Public Function RightShift(ByVal Value As Long, ByVal iShift As Integer)
    RightShift = Value \ (2 ^ iShift)
End Function
