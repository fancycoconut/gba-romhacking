Attribute VB_Name = "modMain"
Option Explicit
Private sFilePath As String
Private iFileNum As Integer
Private sHeader As String * 4
Private MapHack As Long
Private MapBankHeader As Long

Private Sub Main()
Dim bTemp As Byte
Dim sResult As String
Dim cdgOpen As clsCommonDialog
    Set cdgOpen = New clsCommonDialog
    
    If Val(InputBox("This is a private build. Please enter the password.")) <> 2148 Then Exit Sub
    
    sResult = cdgOpen.ShowOpen(0, "Open ROM...", , "Gameboy Advance ROMs (*.gba,*.agb,*.bin)|*.gba;*.agb;*.bin")
    If LenB(sResult) = 0 Then GoTo EndMe
    sFilePath = sResult
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
    Get #iFileNum, &HAC + 1, sHeader
    
        Select Case sHeader
            Case "AXVE"
                MapHack = &H53314
                MapBankHeader = &H8308588
                'MapBankHeader = &H86D7030 ' My hack currently
            
            Case "BPRE"
                MapHack = &H5523C
                MapBankHeader = &H83526A8
            
            Case "BPEE"
                MapHack = &H84A94
                MapBankHeader = &H8486578
                
            Case Else
                MsgBox "Error - Unsupported ROM", vbExclamation
                GoTo EndMe
            
        End Select
        
        MapBankHeader = Val(InputBox("Is this the right Map Bank Header offset?", "Map Bank Header Offset", "&H" & Hex(MapBankHeader)))
        If Val(MapBankHeader) = 0 Then GoTo EndMe
        
        Get #iFileNum, MapHack + 1, bTemp
        If bTemp <> &H3 Then
            If MsgBox("ROM is currently locked. Do you want to unlock it?", vbYesNo) = vbNo Then Exit Sub
            
            Put #iFileNum, MapHack + 1, CByte(&H3)
            Put #iFileNum, MapHack + 1 + 6, &HB896800
            Put #iFileNum, MapHack + 1 + 16, MapBankHeader
            Put #iFileNum, MapHack + 1 + 38, &HC090409
            Put #iFileNum, MapHack + 1 + 42, &HFFE7F7FF
            Put #iFileNum, MapHack + 1 + 46, &H4708BC02
            Put #iFileNum, MapHack + 1 + 48, CLng(&H4708)
            
            MsgBox "ROM unlocked successfully.", vbInformation
        Else
            If MsgBox("ROM is not locked. Lock it?", vbYesNo) = vbNo Then Exit Sub
            
            Put #iFileNum, MapHack + 1, CByte(&HB)
            Put #iFileNum, MapHack + 1 + 6, &H68000B89
            Put #iFileNum, MapHack + 1 + 16, &H80000C0
            Put #iFileNum, MapHack + 1 + 38, &HFFE9F7FF
            Put #iFileNum, MapHack + 1 + 42, &H4708BC02
            Put #iFileNum, MapHack + 1 + 46, CInt(&H0)
            Put #iFileNum, MapHack + 1 + 48, MapBankHeader
            
            MsgBox "ROM locked successfully.", vbInformation
        End If
        
    Close #iFileNum
EndMe:
    Set cdgOpen = Nothing
End Sub
