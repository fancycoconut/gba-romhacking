VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDest As Any, pSource As Any, ByVal cb As Long)

Private sString As String
Private iChunkSize As Long
Private iPos As Long
Private iLen As Long

Public Property Get Length() As Long
   Length = iPos \ 2
End Property

Public Property Get ToString() As String
   ' The internal string:
   If iPos > 0 Then
      ToString = LeftB$(sString, iPos)
   End If
End Property

Public Property Let TheString(ByRef sThis As String)
Dim lLen As Long

    ' Setting the string:
    lLen = LenB(sThis)
    If lLen <> 0 Then
        
        If iLen < lLen Then
            ' Need to expand string to accommodate:
            Do
                sString = sString & String$(iChunkSize \ 2, vbNullChar)
                iLen = iLen + iChunkSize
            Loop While iLen < lLen
        End If
    
        RtlMoveMemory ByVal StrPtr(sString), ByVal StrPtr(sThis), lLen
        iPos = lLen
     
    Else
        ' Clear
        sString = vbNullString
        iPos = 0
        iLen = 0
    End If
   
End Property

Public Sub Clear()
    sString = vbNullString
    iPos = 0
    iLen = 0
End Sub

Public Sub Append(ByRef sThis As String)
Dim lLen As Long

   ' Append an item to the string:
   lLen = LenB(sThis)
   If (iPos + lLen) > iLen Then
      sString = sString & String$(iChunkSize \ 2, vbNullChar)
      iLen = iLen + iChunkSize
   End If
   
   RtlMoveMemory ByVal UnsignedAdd(StrPtr(sString), iPos), ByVal StrPtr(sThis), lLen
   iPos = iPos + lLen
   
End Sub

Public Sub Insert(ByVal iIndex As Long, ByRef sThis As String)
Dim lLen As Long
Dim lPos As Long
Dim lSize As Long

   ' is iIndex within bounds?
   If (iIndex * 2& > iPos) Then
      Err.Raise 9
   Else

      lLen = LenB(sThis)
      If (iPos + lLen) > iLen Then
         sString = sString & String$(iChunkSize \ 2, vbNullChar)
         iLen = iLen + iChunkSize
      End If

      ' Move existing characters from current position
      lPos = UnsignedAdd(StrPtr(sString), iIndex * 2&)
      lSize = iPos - iIndex * 2&

      ' moving from iIndex to iIndex + lLen
      RtlMoveMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize

      ' Insert new characters:
      RtlMoveMemory ByVal lPos, ByVal StrPtr(sThis), lLen

      iPos = iPos + lLen

   End If

End Sub

Public Sub Overwrite(ByVal iIndex As Long, ByRef sThis As String)
Dim lLen As Long
Dim lPos As Long

    ' is iIndex within bounds?
    If (iIndex * 2& > iPos) Then
        Err.Raise 9
    Else
        
        lLen = LenB(sThis)
        If (iPos + lLen) > iLen Then
            sString = sString & String$(iChunkSize \ 2, vbNullChar)
            iLen = iLen + iChunkSize
        End If
        
        lPos = UnsignedAdd(StrPtr(sString), iIndex * 2&)
        RtlMoveMemory ByVal lPos, ByVal StrPtr(sThis), lLen
        
    End If
    
End Sub

Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
Dim lSrc As Long
Dim lDst As Long
Dim lSize As Long

   ' is iIndex within bounds?
   If (iIndex * 2& > iPos) Then
      Err.Raise 9
   Else
      ' is there sufficient length?
      If ((iIndex + lLen) * 2& > iPos) Then
         Err.Raise 9
      Else
         ' Need to copy characters from iIndex*2 to iPos back by lLen chars:
         lSrc = UnsignedAdd(StrPtr(sString), (iIndex + lLen) * 2&)
         lDst = UnsignedAdd(StrPtr(sString), iIndex * 2&)
         lSize = (iPos - (iIndex + lLen) * 2&)
         RtlMoveMemory ByVal lDst, ByVal lSrc, lSize
         iPos = iPos - lLen * 2&
      End If
   End If
   
End Sub

Public Function Find(ByVal sToFind As String, Optional ByVal lStartIndex As Long = 1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
Dim lIns As Long

    If LenB(sToFind) <= iLen Then
    
        If (lStartIndex > 0) Then
            lIns = InStr(lStartIndex, sString, sToFind, Compare)
        Else
            If lStartIndex <= Length Then
                lIns = InStr(sString, sToFind, Compare)
            End If
        End If
        
        If (lIns < iPos \ 2&) Then
            Find = lIns
        End If
    
    End If
    
End Function

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   ElseIf Start And &H80000000 Then 'Start < 0
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If

End Function

Private Sub Class_Initialize()
   iChunkSize = 16384
End Sub
