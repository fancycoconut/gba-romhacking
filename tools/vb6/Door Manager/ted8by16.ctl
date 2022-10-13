VERSION 5.00
Begin VB.UserControl GBATileEditor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   133
   ToolboxBitmap   =   "ted8by16.ctx":0000
End
Attribute VB_Name = "GBATileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim TileData(8, 16) As Byte

Public Enum teBorder
    teNone = 0
    teInset = 1
End Enum

Event Changed()
Attribute Changed.VB_Description = "Sent whenever the user draws a pixel."
Const m_def_ShowGrid = 0
Const m_def_Filename = ""
Const m_def_PenColor = 15
Const m_def_ROMAddress = 0
Const m_def_DotSize = 16
Dim m_ShowGrid As Boolean
Dim m_Filename As String
Dim m_PenColor As Byte
Dim m_ROMAddress As Long
Dim m_DotSize As Integer
Dim Palette(0 To 15) As Long
 
Public Property Get Colors(ByVal iPals As Integer) As OLE_COLOR
Attribute Colors.VB_Description = "An array of colors to draw with."
    If iPals > 15 Then Err.Raise 1601, "Kawa's Tile Editor", "You only have 16 colors!"
    Colors = Palette(iPals)
End Property

Public Property Let Colors(ByVal iPals As Integer, newColor As OLE_COLOR)
    If iPals > 15 Then Err.Raise 1601, "Kawa's Tile Editor", "You only have 16 colors!"
    Palette(iPals) = newColor
    PropertyChanged "Colors"
End Property
 
Public Property Get BorderStyle() As teBorder
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As teBorder)
    If New_BorderStyle > teInset Then Err.Raise 1600, "Kawa's Tile Editor", "Borderstyle can't be higher than 1."
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get DotSize() As Integer
Attribute DotSize.VB_Description = "Returns/sets the scale of drawing."
    DotSize = m_DotSize
End Property

Public Property Let DotSize(ByVal New_DotSize As Integer)
    m_DotSize = New_DotSize
    PropertyChanged "DotSize"
    UserControl_Resize
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Dim x, y As Integer
    For x = 0 To 7
        For y = 0 To 15
            Line (x * m_DotSize, y * m_DotSize)-((x + 1) * m_DotSize, (y + 1) * m_DotSize), Palette(TileData(x, y)), BF
        Next y
    Next x
    
    DoGrid
End Sub

Private Sub DoGrid()
Dim x, y As Integer
    If m_ShowGrid = True Then
        For x = 1 To 7
            For y = 1 To 15
                Line (x * m_DotSize, 0)-(x * m_DotSize, 10024), QBColor(15)
                Line (0, y * m_DotSize)-(10024, y * m_DotSize), QBColor(15)
            Next y
        Next x
    End If
End Sub

Public Sub SetPalette(NewPal() As Long)
Attribute SetPalette.VB_Description = "Set all 16 colors at once."
Dim i As Integer
    For i = 0 To 15
        Palette(i) = NewPal(i)
    Next i
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Tag = "^_^"
    UserControl_MouseMove Button, Shift, x, y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim px As Integer, py As Integer
Dim ax As Integer, ay As Integer

    If Tag <> "^_^" Then Exit Sub
    px = Int(x / m_DotSize)
    py = Int(y / m_DotSize)
    If px < 0 Or px > 7 Then Exit Sub
    If py < 0 Or py > 15 Then Exit Sub
    ax = px * m_DotSize
    ay = ay * m_DotSize
    TileData(px, py) = IIf(Button = 2, 0, PenColor)
    
    Line (px * m_DotSize, py * m_DotSize)-((px + 1) * m_DotSize, (py + 1) * m_DotSize), Palette(TileData(px, py)), BF
    DoGrid
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Tag = "-_-"
End Sub

Private Sub UserControl_Paint()
    Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ROMAddress = PropBag.ReadProperty("ROMAddress", m_def_ROMAddress)
    m_PenColor = PropBag.ReadProperty("PenColor", m_def_PenColor)
    m_Filename = PropBag.ReadProperty("Filename", m_def_Filename)
    m_ShowGrid = PropBag.ReadProperty("ShowGrid", m_def_ShowGrid)
    m_DotSize = PropBag.ReadProperty("DotSize", m_def_DotSize)
End Sub

Private Sub UserControl_Resize()
    Width = ((m_DotSize * 8) + IIf(UserControl.BorderStyle = 1, 4, 0)) * Screen.TwipsPerPixelX
    Height = ((m_DotSize * 8) + IIf(UserControl.BorderStyle = 1, 4, 0)) * Screen.TwipsPerPixelY
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ROMAddress", m_ROMAddress, m_def_ROMAddress)
    Call PropBag.WriteProperty("PenColor", m_PenColor, m_def_PenColor)
    Call PropBag.WriteProperty("Filename", m_Filename, m_def_Filename)
    Call PropBag.WriteProperty("ShowGrid", m_ShowGrid, m_def_ShowGrid)
    Call PropBag.WriteProperty("DotSize", m_DotSize, m_def_DotSize)
End Sub

Public Sub LoadTileData()
Attribute LoadTileData.VB_Description = "Load data from the file and address specified and show it."
Dim RawData(31) As Byte ' Needs to be edited if used in another project
Dim iFileNum As Integer
Dim i As Integer
Dim iColor1 As Integer
Dim iColor2 As Integer
'Dim s As Integer
Dim x As Integer, y As Integer

    iFileNum = FreeFile
    Open m_Filename For Binary As iFileNum
        Get #iFileNum, m_ROMAddress + 1, RawData
    Close #iFileNum
  
    For i = 0 To 31 ' Same as RawData's size
        iColor1 = RawData(i)
        'For s = 0 To 3: iColor1 = iColor1 \ 2: Next s 'simulates color1 >> 4
        iColor1 = iColor1 \ (2 ^ 4) ' iColor1 >> 4
        iColor1 = iColor1 And &HF
        iColor2 = RawData(i) And &HF

        TileData(x + 1, y) = iColor1
        TileData(x, y) = iColor2
        x = x + 2
        If x > 7 Then
            x = 0
            y = y + 1
        End If
    Next i
  
    Refresh
End Sub

Public Sub SaveTileData()
Attribute SaveTileData.VB_Description = "Save the tile as-is at the specified location in the specified file."
Dim RawData(31) As Byte ' Needs to be edited if used in another project
Dim i As Integer
Dim iColor1 As Integer
Dim iColor2 As Integer
Dim x As Integer, y As Integer
Dim iFileNum As Integer
  
    For i = 0 To 31 ' Same as RawData's size
        iColor1 = TileData(x + 1, y)
        iColor2 = TileData(x, y)
        RawData(i) = CByte(Val("&H" & Hex(iColor1) & Hex(iColor2)))
        'Debug.Print Hex(RawData(i))
        x = x + 2
        
        If x > 7 Then
             x = 0
            y = y + 1
        End If
        
    Next i
  
    iFileNum = FreeFile
    Open m_Filename For Binary As iFileNum
        Put #iFileNum, m_ROMAddress + 1, RawData
    Close #iFileNum
  
End Sub

Public Property Get ROMAddress() As Long
Attribute ROMAddress.VB_Description = "The address in the file stored in Filename to read from."
    ROMAddress = m_ROMAddress
End Property
Public Property Let ROMAddress(ByVal New_ROMAddress As Long)
    m_ROMAddress = New_ROMAddress
    PropertyChanged "ROMAddress"
End Property

Private Sub UserControl_InitProperties()
    m_ROMAddress = m_def_ROMAddress
    m_PenColor = m_def_PenColor
    m_Filename = m_def_Filename
    m_ShowGrid = m_def_ShowGrid
End Sub

Public Property Get PenColor() As Byte
Attribute PenColor.VB_Description = "The palette index to draw with."
Attribute PenColor.VB_UserMemId = 0
    PenColor = m_PenColor
End Property
Public Property Let PenColor(ByVal New_PenColor As Byte)
    m_PenColor = New_PenColor
    PropertyChanged "PenColor"
End Property

Public Property Get Filename() As String
Attribute Filename.VB_Description = "The file name to read from."
    Filename = m_Filename
End Property
Public Property Let Filename(ByVal New_Filename As String)
    m_Filename = New_Filename
    PropertyChanged "Filename"
End Property

Public Property Get ShowGrid() As Boolean
Attribute ShowGrid.VB_Description = "Toggle drawing a white grid."
    ShowGrid = m_ShowGrid
End Property
Public Property Let ShowGrid(ByVal New_ShowGrid As Boolean)
    m_ShowGrid = New_ShowGrid
    PropertyChanged "ShowGrid"
End Property

