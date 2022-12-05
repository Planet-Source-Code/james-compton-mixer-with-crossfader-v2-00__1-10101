Attribute VB_Name = "Module1"
Option Explicit


Public Function InStrRevVB5(ByVal StringToCheck As String, ByVal StringToMatch As String, Optional ByVal StartAt As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
 
Dim lPos        As Long
Dim lSavePos    As Long
 
    ' -1 means search entire string. A positive number
    ' means search only up to that position from the left.
    If StartAt = -1 Then StartAt = Len(StringToCheck)
    
    ' Find the last instance of StringToMatch within StringToCheck.
    lPos = InStr(1, StringToCheck, StringToMatch, Compare)
    While lPos > 0 And lPos < StartAt
        lSavePos = lPos
        lPos = InStr(lPos + 1, StringToCheck, StringToMatch, Compare)
    Wend
    
    InStrRevVB5 = lSavePos
        
End Function

Public Function BasePath(ByVal fname As String, Optional delim As String = "\", Optional keeplast As Boolean = True) As String
    Dim outstr As String
    Dim llen As Long
    llen = InStrRevVB5(fname, delim)


    If (Not keeplast) Then
        llen = llen - 1
    End If


    If (llen > 0) Then
        BasePath = Mid(fname, 1, llen)
    Else
        BasePath = fname
    End If
End Function


