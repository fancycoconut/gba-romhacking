Attribute VB_Name = "modExtras"
Option Explicit
' Globals
Public iFileNum As Integer
Public sFilePath As String
Public sHeader As String * 4
Public EnlargedROM As Byte
Public arrPalette() As Byte

' Day Night Time
Public Type Time
    Hours As Byte
    Minutes As Byte
End Type

Public tmMorning As Time
Public tmDay As Time
Public tmAfternoon As Time
Public tmEvening As Time
Public tmNight As Time

' Palette Shades
Public Type PaletteShade
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Public MorningEffect As PaletteShade
Public AfternoonEffect As PaletteShade
Public EveningEffect As PaletteShade
Public NightEffect As PaletteShade

Public Hook1 As Long
Public Hook2 As Long
Public HookRTC As Long
Public SWI01Pos As Long
Public NintendoFix As Long

' Check Time Offsets
Public RTC As Long
Public MenuFlag As Long
Public IndoorFlag As Long
Public BattleFlag As Long
Public RTCReturnAddress As Long

' DMA3 Offsets
Public PaletteOriginal As Long
Public DMA3ReturnAddress As Long
