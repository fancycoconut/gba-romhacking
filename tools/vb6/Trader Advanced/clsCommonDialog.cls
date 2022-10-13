VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private cdlg As OPENFILENAME
Private LastFileName As String, LastFilter As Long

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpsFilter       As String
    lpsCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpsFile         As String
    nMaxFile          As Long
    lpsFileTitle    As String
    nMaxFileTitle     As Long
    lpsInitialDir   As String
    lpsTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpsDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Public Enum DialogFlags
    FileMustExist = OFN_FILEMUSTEXIST
    HideReadOnly = OFN_HIDEREADONLY
    OVERWRITEPROMPT = OFN_OVERWRITEPROMPT
    PATHMUSTEXIST = OFN_PATHMUSTEXIST
End Enum

Private Const sPipe As String = "|"

Public Function ShowOpen(ByVal hWnd As Long, ByVal Title As String, Optional ByVal InitDir As String = "", Optional ByVal Filter As String = "All Files (*.*)|*.*|", Optional flags As DialogFlags = FileMustExist Or HideReadOnly Or PATHMUSTEXIST) As String
    
    Filter = Replace(Filter, sPipe, vbNullChar)
    
    If Right$(Filter, 1) <> vbNullChar Then Filter = Filter & vbNullChar
    If LenB(InitDir) = 0 Then InitDir = LastFileName
    
    With cdlg
        .lStructSize = Len(cdlg)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpsFilter = Filter
        .lpsFile = Space$(254)
        .nMaxFile = 255
        .lpsFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpsInitialDir = InitDir
        .lpsTitle = Title
        .flags = flags
    End With
    
    If GetOpenFileName(cdlg) Then
        ShowOpen = Left$(cdlg.lpsFile, InStr(cdlg.lpsFile, vbNullChar) - 1)
        LastFileName = ShowOpen
        LastFilter = cdlg.nFilterIndex
    End If

End Function

Public Function ShowSave(ByVal hWnd As Long, ByVal Title As String, Optional ByVal sFileName As String = "", Optional InitDir As String = "", Optional ByVal Filter As String, Optional flags As DialogFlags = FileMustExist Or OVERWRITEPROMPT Or PATHMUSTEXIST) As String
Dim Buffer As String
Dim Result As Long

    Filter = Replace(Filter, sPipe, vbNullChar)
    
    If Right$(Filter, 1) <> vbNullChar Then Filter = Filter & vbNullChar
    Buffer = sFileName & String$(254 - Len(sFileName), 0)
    If LenB(InitDir) = 0 Then InitDir = LastFileName
    
    With cdlg
        .lStructSize = Len(cdlg)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .flags = flags
        .nMaxFile = Len(Buffer)
        .lpsFile = Buffer
        .lpsFilter = Filter
        .lpsInitialDir = InitDir
        .lpsTitle = Title
        .lpsFilter = Filter
        .lpsDefExt = ".bin"
    End With
    
    Result = GetSaveFileName(cdlg)
    
    If Result <> 0 Then
        ShowSave = Left$(cdlg.lpsFile, InStr(cdlg.lpsFile, vbNullChar) - 1)
        LastFilter = cdlg.nFilterIndex
        LastFileName = ShowSave
    End If

End Function

Public Property Get FilterIndex() As Long
    FilterIndex = LastFilter
End Property
