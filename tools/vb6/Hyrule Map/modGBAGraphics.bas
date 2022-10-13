Attribute VB_Name = "modGBAGraphics"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'BitBlts
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'Pixels
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public Function BinToDecI(Bin As String) As Integer
Dim i As Integer
Dim nDec As Integer
Dim nPos As Integer

If Len(Bin) > 16 Then
    Err.Raise 6
Else
    For i = Len(Bin) To 1 Step -1
        If Mid$(Bin, i, 1) = "1" Then
            SetBitI nDec, nPos
        End If
        nPos = nPos + 1
    Next
    BinToDecI = nDec
End If
End Function

Private Function ByteSwap(Data As Byte, Data2 As Byte) As String
    ByteSwap = Right$("0" & Hex$(Data), 2) & Right$("0" & Hex$(Data2), 2)
End Function

Public Function DecompressGold(ByVal Offset As Long, gfxBuffer() As Byte) As Long
  Dim gfxPointer As Long
  Dim byteIn As Byte
  Dim a As Byte
  Dim c As Byte
  Dim n As Byte
  Dim s As Byte
  Dim w As Byte
  Dim x As Byte
  Dim Y As Byte
  Dim z As Byte
  Dim i As Integer
  Get #256, Offset, byteIn
  Do
    Get #256, , byteIn
    If byteIn = &HFF Then Exit Do
    c = byteIn And &HE0
    x = byteIn And &H1F
recalc:
    Select Case c
      Case 0
        For i = 0 To x
          Get #256, , byteIn
          gfxBuffer(gfxPointer) = byteIn
          gfxPointer = gfxPointer + 1
        Next i
      Case &H20
        Get #256, , byteIn
          For i = 0 To x
          gfxBuffer(gfxPointer) = byteIn
          gfxPointer = gfxPointer + 1
        Next i
      Case &H40
        Get #256, , byteIn
        Y = byteIn
        Get #256, , byteIn
        z = byteIn
        For i = 0 To x
          gfxBuffer(gfxPointer) = IIf(i Mod 2 = 0, Y, z)
          gfxPointer = gfxPointer + 1
        Next i
      Case &H60
        For i = 0 To x
          gfxBuffer(gfxPointer) = 0
          gfxPointer = gfxPointer + 1
        Next i
      Case &H80
        Get #256, , byteIn
        a = byteIn
        If a And &H80 = 0 Then
          Get #256, , byteIn
          n = byteIn
          For i = 0 To x
            gfxBuffer(gfxPointer) = gfxBuffer((a * &H100) + (n + 1) + i)
            gfxPointer = gfxPointer + 1
          Next i
        Else
          a = a And &H7F
          s = gfxPointer - a - 1
          For i = 0 To x
            gfxBuffer(gfxPointer) = gfxBuffer(s + i)
            gfxPointer = gfxPointer + 1
          Next i
        End If
      Case &HA0
        Get #256, , byteIn
        a = byteIn
        If a And &H80 = 0 Then
          Get #256, , byteIn
          n = byteIn
          For i = 0 To x 'reverse bit order
            gfxBuffer(gfxPointer) = gfxBuffer((a * &H100) + (n + 1) + i)
            gfxPointer = gfxPointer + 1
          Next i
        Else
          a = a And &H7F
          s = gfxPointer - a - 1
          For i = 0 To x 'reverse bit order..
            gfxBuffer(gfxPointer) = gfxBuffer(s + i)
            gfxPointer = gfxPointer + 1
          Next i
        End If
      Case &HC0
        Get #256, , byteIn
        a = byteIn
        If a And &H80 = 0 Then
          Get #256, , byteIn
          n = byteIn
          For i = 0 To x
            gfxBuffer(gfxPointer) = gfxBuffer((a * &H100) + (n + 1) - i)
            gfxPointer = gfxPointer + 1
          Next i
        Else
          a = a And &H7F
          s = gfxPointer - a - 1
          For i = 0 To x
            gfxBuffer(gfxPointer) = gfxBuffer(s - i)
            gfxPointer = gfxPointer + 1
          Next i
        End If
      Case &HE0
        c = x And &H1C
        Get #256, , byteIn
        w = byteIn
        x = ((x And 3) * &H100) + w
        GoTo recalc
    End Select
  Loop
  Close #256
  DecompressGold = gfxPointer
End Function

Public Sub DrawTile8(ByVal hDC As Long, ByVal Map16n As Long, ByRef bGfx() As Byte, ByRef lPal() As Long)
Dim bTileData(31) As Byte
Dim iThisPal As Integer
Dim i As Long
Dim FlipX As Boolean
Dim FlipY As Boolean
Dim X1 As Long
Dim X2 As Long
Dim Y As Long
Dim bColA As Byte
Dim bColB As Byte
Dim lDestX As Long
Dim lDestY As Long

    lDestX = (Map16n Mod 8&) * 8&
    lDestY = (Map16n \ 8&) * 8&

    CopyMemory bTileData(0), bGfx((Map16n And &H3FF&) * (UBound(bTileData) + 1)), UBound(bTileData) + 1

    Select Case (Map16n And &H400&)
        Case &H400&
            FlipX = True
        Case &H800&
            FlipY = True
        Case &HC00&
            FlipX = True
            FlipY = True
    End Select

    iThisPal = Map16n \ 65536

    For i = 0 To UBound(bTileData)

        X1 = (i * 2&) Mod 8&
        X2 = X1 + 1
        Y = (i \ 4&)

        If FlipX = True Then
            X1 = 7& - X1
            X2 = 7& - X2
        End If

        If FlipY = True Then
            Y = 7& - Y
        End If

        bColA = bTileData(i) Mod 16&
        bColB = bTileData(i) \ 16&

        If bColA <> 0 Then
            SetPixel hDC, lDestX + X1, lDestY + Y, lPal(iThisPal, bColA)
        End If

        If bColB <> 0 Then
            SetPixel hDC, lDestX + X2, lDestY + Y, lPal(iThisPal, bColB)
        End If

    Next i

End Sub

Public Sub DrawTiles(ByVal hDC As Long, ByRef bGfx() As Byte, ByRef lPal() As Long)
Dim i As Long
End Sub

Function Dual(ByVal vValue As Long) As String
    vValue = CVar(vValue)
    If Not IsNumeric(vValue) Then
        Dual = "Value is non-numerical!"
        Exit Function
    ElseIf vValue > 999999999 Then
        Dual = "Number is too high!"
        Exit Function
    End If
    Do
        If vValue Mod 2 = 0 Then
            Dual = "0" & Dual
        Else
            Dual = "1" & Dual
        End If
        vValue = vValue \ 2
    Loop While vValue > 0
    Dual = Format(Dual, "000000000000000")
End Function

Public Function LZ77UnComp(Source() As Byte, Dest() As Byte) As Long
On Error Resume Next

Dim i As Long, j As Long
Dim XIn As Long, XOut As Long
    XIn = 4
    XOut = 0
Dim Length As Long
Dim Offset As Long
Dim WindowOffset As Long
Dim retLen As Long
Dim xLen As Long
Dim d As Byte
Dim Data As Long

xLen = (Source(0) Or (Source(1) * 256&) Or (Source(2) * 65536) Or (Source(3) * 16777216)) \ 256&
retLen = xLen
ReDim Dest(0 To xLen - 1) As Byte
  
    Do While xLen > 0&
        d = Source(XIn)
        XIn = XIn + 1&
      
        For i = 0& To 7&
        
            If (d And &H80&) <> 0 Then
                Data = ((Source(XIn) * 256&) Or Source(XIn + 1))
                XIn = XIn + 2&
                Length = (Data \ &H1000&) + 3&
                Offset = (Data And &HFFF)
                WindowOffset = XOut - Offset - 1&
                
                For j = 0 To Length - 1&
                    Dest(XOut) = Dest(WindowOffset)
                    XOut = XOut + 1&
                    WindowOffset = WindowOffset + 1&
                    xLen = xLen - 1&
                    
                    If xLen = 0 Then
                        LZ77UnComp = retLen
                        Exit Function
                    End If
                Next j
            Else
                Dest(XOut) = Source(XIn)
                XOut = XOut + 1&
                XIn = XIn + 1&
                xLen = xLen - 1&
                If xLen = 0 Then
                    LZ77UnComp = retLen
                    Exit Function
                End If
            End If
          d = (d * 2) And &HFF&
        Next i
    Loop
    
    LZ77UnComp = retLen
    
End Function

Public Function LZ77Comp(DecmpSize As Long, Source() As Byte, Dest() As Byte) As Long
Dim i As Long, j As Long
Dim XIn As Long, TmpXIn As Long
Dim XOut As Long, TmpXOut As Long           'XOut = Poss./Length in new array
Dim Length As Long, TmpLength As Long
Dim Offset As Long, TmpOffset As Long
Dim Ctrl As Byte
Dim XData(0 To 7, 0 To 1) As Byte

XOut = 4&
Dest(0) = &H10  'Unknown byte?
Dest(1) = DecmpSize And &HFF&
Dest(2) = (DecmpSize \ &H100&) And &HFF&
Dest(3) = (DecmpSize \ &H10000) And &HFF&
  
      Do While ((DecmpSize - 1&) >= TmpXIn)
            Ctrl = 0
            
            For i = 7& To 0& Step -1&

                If (XIn < &H1000) Then
                    j = XIn
                Else
                    j = &H1000&
                End If
        
                Length = 0  'Reset Length
                Offset = 0  'Reset Offset
              
                Do While (j > 1&)
                    TmpXIn = XIn + 1&
                    TmpXOut = (XIn - j) + 1&
                    
                    Do While Source(TmpXIn - 1&) = Source(TmpXOut - 1&)
                        TmpXIn = TmpXIn + 1&
                        TmpXOut = TmpXOut + 1&
                    
                        If (TmpXIn >= (DecmpSize - 1&)) Then
                            Exit Do
                        End If
                    Loop
                    
                    TmpLength = (TmpXIn - XIn - 1&)
                    TmpOffset = (TmpXIn - TmpXOut - 1&)
                    
                    If (TmpLength > Length) Then
                        Length = TmpLength
                        Offset = TmpOffset
                          
                        If (Length >= &H12&) Then
                            Length = &H12&
                            Exit Do
                        End If
                    End If
                    
                    j = j - 1&
                Loop
                
                If (Length < 3&) Then
                    XData(i, 0) = Source(XIn)
                    
                    XIn = XIn + 1
                    If (XIn > (DecmpSize - 1&)) Then
                        Exit For
                    End If
                Else
                    Ctrl = Ctrl Or (2& ^ i)
                    XData(i, 0) = ((Length - 3&) * 16&) Or (Offset \ 256&)
                    XData(i, 1) = Offset And &HFF&
                    
                    XIn = XIn + Length
                    If (XIn > (DecmpSize - 1&)) Then
                        Exit For
                    End If
                End If
            Next i
            
            Dest(XOut) = Ctrl
            XOut = XOut + 1&
            
            For i = 7& To 0& Step -1&
                
                Dest(XOut) = XData(i, 0)
                XOut = XOut + 1&
                
                If ((Ctrl And &H80&) <> 0) Then
                    Dest(XOut) = XData(i, 1)
                    XOut = XOut + 1&
                End If
                
                Ctrl = (Ctrl * 2&) And &HFF&
                
            Next i
      Loop
      
EndMe:
      LZ77Comp = XOut
      Exit Function
      
End Function

Public Sub UnPackPalette(ByRef GBAPalette() As Integer, ByRef PCPalette() As Long)
Dim Red As Integer
Dim Green As Integer
Dim Blue As Integer
Dim i, ii As Integer
Dim Index As Long
For ii = 0 To 15
    Index = &H10 * ii

    For i = 0 To 16 '16
        Red = ((GBAPalette(Index + i)) And 31) * 8
        Green = (((GBAPalette(Index + i)) \ 32) And 31) * 8
        Blue = ((GBAPalette(Index + i) \ 1024) And 31) * 8
        PCPalette(ii, i) = RGB(Red, Green, Blue)
    Next i
    
Next ii
End Sub

Public Function UnPackPaletteRGB(ByVal GBAPalette As Integer) As Long
Dim Red As Integer
Dim Green As Integer
Dim Blue As Integer
    Red = (GBAPalette And 31) * 8
    Green = (((GBAPalette) \ 32) And 31) * 8
    Blue = ((GBAPalette \ 1024) And 31) * 8
    UnPackPaletteRGB = RGB(Red, Green, Blue)
End Function

Public Function Colour15To24(ByVal ColourData As Integer) As Long
Dim r As Byte, G As Byte, B As Byte
r = ((ColourData And 31) / 31) * &HFF
G = (((ColourData \ 32) And 31) / 31) * &HFF
B = (((ColourData \ 1024) And 31) / 31) * &HFF
Colour15To24 = CLng(B) + (256 * CLng(G)) + (65536 * CLng(r))
End Function

Function GBA2RGB(GBPalette As String, Optional bSwitch As Boolean = False) As String
Dim Red As String, Green As String, Blue As String
Dim memory As String, memoryL As Long
    memoryL = Val("&H" & Right$(GBPalette, 2) & Left$(GBPalette, 2))
    memory = Dual(memoryL)
    Blue = Left$(memory, 5)
    memory = Right$(memory, 10)
    Green = Left$(memory, 5)
    Red = Right$(memory, 5)
    Red = BinToDecI(Red)
    Red = Round((Red * 255) / 31)
    Green = BinToDecI(Green)
    Green = Round((Green * 255) / 31)
    Blue = BinToDecI(Blue)
    Blue = Round((Blue * 255) / 31)
    Red = Right$("0" & Hex$(Red), 2)
    Green = Right$("0" & Hex$(Green), 2)
    Blue = Right$("0" & Hex$(Blue), 2)
      
    If Not bSwitch Then
        GBA2RGB = Red & Green & Blue
    Else
        GBA2RGB = Blue & Green & Red
    End If
End Function

Public Sub SetBitI(Value As Integer, ByVal Position As Byte)
Select Case Position
    Case 0 To 14
        Value = Value Or 2 ^ Position
    Case 15
        Value = Value Or &H8000
    Case Else
        Err.Raise 6
End Select
End Sub

Function RGB2GBA(RGBPalette As String) As String
Dim Red As String, Green As String, Blue As String, sTemp As String
Red = Val("&H" & Left$(RGBPalette, 2))
RGBPalette = Right$(RGBPalette, 4)
Green = Val("&H" & Left$(RGBPalette, 2))
Blue = Val("&H" & Right$(RGBPalette, 2))
sTemp = Hex$((Blue \ 8) * 1024 + (Green \ 8) * 32 + (Red \ 8))
sTemp = Right$("000" & sTemp, 4)
RGB2GBA = Right$(sTemp, 2) & Left$(sTemp, 2)
End Function
