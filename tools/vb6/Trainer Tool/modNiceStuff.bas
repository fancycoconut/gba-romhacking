Attribute VB_Name = "modNiceStuff"
Option Explicit
' For Web Browser or others
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' For Read/Write Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' XP Style Stuff
Public Type tagInitCommonControlsEx
   lSize As Long
   lICC As Long
End Type

Public m_hMod As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

' LocaliseStrings Stuff
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

'SetIcon Stuff
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50

Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Const LR_SHARED = &H8000&
Private Const IMAGE_ICON = 1

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

Public Function AddToINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
Dim lRet As Long
    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    AddToINI = (lRet)
End Function

Public Function FileExists(FileName As String) As Boolean
    On Error GoTo ExitFunc
    If LenB(FileName) Then
        If LenB(Dir$(FileName)) Then
            FileExists = True
        End If
    End If
ExitFunc:
End Function

Public Function GetFromINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)
Dim sBuffer As String, lRet As Long
    sBuffer = Space$(255)
    lRet = GetPrivateProfileString(sSection, sKey, vbNullString, sBuffer, Len(sBuffer), sIniFile)
    If lRet = 0 Then
        If LenB(sDefault) <> 0 Then AddToINI sSection, sKey, sDefault, sIniFile
        GetFromINI = sDefault
    Else
        GetFromINI = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
End Function

Public Sub Localize(frm As Form)
Dim ctl As Control
Dim lVal As Long
    
    ' set the form's caption
    If Val(frm.Tag) > 0 Then
        frm.Caption = LoadResString(Val(frm.Tag))
    End If
    
    For Each ctl In frm.Controls
        If TypeName(ctl) <> "Menu" Then
            lVal = Val(ctl.Tag)
        Else
            lVal = Val(ctl.HelpContextID)
        End If
        
        If lVal > 0 Then
            ctl.Caption = LoadResString(lVal)
        End If
    Next
    
End Sub

Public Sub LogWrite(ByVal FileName As String, ByVal INIHeader As String, ByVal variable As String, ByVal TheValue As String)
  Dim AppPath As String
  Dim TempReturn As String
  AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
  TempReturn = WritePrivateProfileString(INIHeader, variable, TheValue, AppPath & FileName & ".txt")
End Sub

Public Sub Main()
Const ICC_USEREX_CLASSES = &H200
Dim iccex As tagInitCommonControlsEx

   On Error Resume Next
   iccex.lSize = LenB(iccex)
   iccex.lICC = ICC_USEREX_CLASSES
   m_hMod = LoadLibrary("shell32.dll")
   InitCommonControlsEx iccex
   
   On Error GoTo 0
   frmMain.Show
   Exit Sub
   
End Sub

Public Sub PutFiller(ByRef sFilePath As String, lOffset As Long, lLength As Long)
Dim Filler() As Byte
Dim iFileNum As Integer
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        ReDim Filler(lLength - 1)
        Put #iFileNum, lOffset + 1, Filler()
    Close #iFileNum
End Sub

Public Sub PutFreeSpace(ByRef sFilePath As String, lOffset As Long, lLength As Long)
Dim FreeSpace() As Byte
Dim iFileNum As Integer
    iFileNum = FreeFile
    
    ReDim FreeSpace(0 To lLength - 1)
    For lLength = LBound(FreeSpace) To UBound(FreeSpace)
        FreeSpace(lLength) = &HFF
    Next lLength
    
    Open sFilePath For Binary As #iFileNum
    Put #iFileNum, lOffset + 1, FreeSpace
    Close #iFileNum
End Sub

Public Sub ReadByteArray(ByVal sPath As String, ByRef arrData() As Byte, Optional lOffset As Long = 0, Optional lLenght As Long = 0)
Dim lFile As Long
    lFile = FreeFile
    Open sPath For Binary Access Read As lFile
        If lLenght = 0 Then ReDim arrData(1 To LOF(lFile)) As Byte
        If lLenght > 0 Then ReDim arrData(1 To lLenght) As Byte
        Get lFile, lOffset + 1, arrData
    Close lFile
End Sub

Public Sub SetIcon(ByVal hWnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long

    If (bSetAsAppIcon) Then
        ' Find VB's hidden parent window:
        lhWnd = hWnd
        lhWndTop = lhWnd
        Do While Not (lhWnd = 0)
            lhWnd = GetWindow(lhWnd, GW_OWNER)
            If Not (lhWnd = 0) Then
                lhWndTop = lhWnd
            End If
        Loop
    End If
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    End If
    SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    End If
    SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall

End Sub

Public Sub WriteByteArray(ByVal sPath As String, ByRef arrData() As Byte, Optional lOffset As Long = 0)
Dim lFile As Long
    lFile = FreeFile
    Open sPath For Binary Access Write As lFile
        Put lFile, lOffset + 1, arrData
    Close lFile
End Sub

