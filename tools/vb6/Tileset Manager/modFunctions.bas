Attribute VB_Name = "modFunctions"
Option Explicit
Public sHeader As String * 4
Public iFileNum As Integer
Public sFilePath As String
Public sTilesetPath As String
Public TilesetData() As Byte
Public RepointOffset As Long
Public TilesetOffset As Long
Public TilesetNumber As Long
Public TilesetHeader As Long
Public PaletteOffset As Long
Public BlockData As Long
Public BehaviourData As Long
Public Temp As Long
Public Blocks As Long
Public Behaviours As Long
Public Ofst As Long
Public BytesAmount As Long
Public EnlargedROM As Byte

Function GenerateLogFile()
    LogWrite "/log", "Log", "ROM", sFilePath
    LogWrite "/log", "Log", "Tileset", sTilesetPath
    LogWrite "/log", "Log", LoadResString(28), TilesetNumber
    LogWrite "/log", "Log", LoadResString(29), Hex(RepointOffset)
    LogWrite "/log", "Log", LoadResString(30), Hex(TilesetOffset)
    LogWrite "/log", "Log", LoadResString(31), Hex(PaletteOffset)
    LogWrite "/log", "Log", LoadResString(32), Hex(BlockData)
    LogWrite "/log", "Log", LoadResString(33), Hex(BehaviourData)
    LogWrite "/log", "Log", LoadResString(34), LoadResString(35)
End Function

Function InsertTilesetHeader()
iFileNum = FreeFile
Open sFilePath For Binary As #iFileNum
    
    If frmMain.chCompression.Value = vbChecked Then
        Put #iFileNum, RepointOffset + 1, &H1
        Put #iFileNum, RepointOffset + 1 + 1, &H1
    Else
        Put #iFileNum, RepointOffset + 1, &H0
        Put #iFileNum, RepointOffset + 1 + 1, &H1
    End If
        
    If frmMain.chMajor.Value = vbChecked Then
        Put #iFileNum, RepointOffset + 1 + 1, &H0
    Else
        If frmMain.chSubColor.Value = vbChecked Then
            Put #iFileNum, RepointOffset + 1 + 1, &H0
        Else
            PaletteOffset = PaletteOffset - 192
        End If
    End If
    
    PutFiller sFilePath, RepointOffset + 2, 2
    Put #iFileNum, RepointOffset + 1 + 4, TilesetOffset
    Put #iFileNum, RepointOffset + 1 + 7, &H8
    Put #iFileNum, RepointOffset + 1 + 8, PaletteOffset
    Put #iFileNum, RepointOffset + 1 + 11, &H8
    Put #iFileNum, RepointOffset + 1 + 12, BlockData
    Put #iFileNum, RepointOffset + 1 + 15, &H8
    
    Select Case Left(sHeader, 3)
        Case "AXV", "AXP", "BPE"
            Put #iFileNum, RepointOffset + 1 + 16, BehaviourData
            Put #iFileNum, RepointOffset + 1 + 19, &H8
            PutFiller sFilePath, RepointOffset + 20, 4
        Case "BPR", "BPG"
            PutFiller sFilePath, RepointOffset + 16, 4
            Put #iFileNum, RepointOffset + 1 + 20, BehaviourData
            Put #iFileNum, RepointOffset + 1 + 23, &H8
    End Select
    
    MsgBox LoadResString(25) & vbNewLine & LoadResString(26) & TilesetNumber & vbNewLine & LoadResString(27), vbInformation
Close #iFileNum
End Function

Function InsertBehaviourData()
BehaviourData = RepointOffset + 24 + Temp + 196 + Blocks
Select Case sHeader
    Case "BPRE", "BPGE"
        Behaviours = 4 * CLng(frmMain.txtBlocks.Text)
    Case Else
        Behaviours = 2 * CLng(frmMain.txtBlocks.Text)
End Select
    PutFiller sFilePath, BehaviourData, Behaviours
End Function

Function InsertBlockData()
    Blocks = 16 * CLng(frmMain.txtBlocks.Text)
    BlockData = RepointOffset + 24 + Temp + 192 + 2
    PutFiller sFilePath, BlockData, Blocks
End Function

Function InsertPalette()
    CountFileBytes sTilesetPath, Temp
    Temp = Temp + 2
    PaletteOffset = RepointOffset + 24 + Temp
    PutFiller sFilePath, PaletteOffset, 192
End Function

Function InsertTileset()
    TilesetOffset = RepointOffset + 24
    GetFileData sTilesetPath, TilesetData
    WriteByteArray sFilePath, TilesetData, TilesetOffset
End Function

Function Calculation()
Select Case sHeader
    Case "BPEE"
        RepointOffset = (CLng("&H" & frmMain.txtOffset.Text) - TilesetHeader) / &H18
        TilesetNumber = RepointOffset
        RepointOffset = (RepointOffset * &H18) + TilesetHeader
        RepointOffset = RepointOffset + 8
    Case Else
        RepointOffset = (CLng("&H" & frmMain.txtOffset.Text) - TilesetHeader) / &H18
        TilesetNumber = RepointOffset
        RepointOffset = (RepointOffset * &H18) + TilesetHeader
End Select
End Function

Function FindOffset()
    Select Case Left(sHeader, 3)
        Case "BPE"
            Ofst = (CLng("&H" & frmMain.txtOffset.Text) - TilesetHeader) / &H18
            Ofst = (Ofst * &H18) + TilesetHeader
            Ofst = Ofst + 8
        Case Else
            Ofst = (CLng("&H" & frmMain.txtOffset.Text) - TilesetHeader) / &H18
            Ofst = (Ofst * &H18) + TilesetHeader
    End Select
End Function

Public Sub CleanFile(ByRef sFilePath As String)
Dim iFileNum As Integer
    iFileNum = FreeFile
    Open sFilePath For Output As #iFileNum
    Close #iFileNum
End Sub

Public Sub CountFileBytes(ByVal sFilePath As String, Amount As Long)
Dim iFileNum As Integer
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Amount = LOF(iFileNum)
    Close #iFileNum
End Sub

Public Sub GetFileData(ByRef sFilePath As String, ByRef TilesetData() As Byte)
Dim iFileNum As Integer
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        ReDim TilesetData(LOF(iFileNum) - 1) As Byte
        Get #iFileNum, 1, TilesetData
    Close #iFileNum
End Sub
