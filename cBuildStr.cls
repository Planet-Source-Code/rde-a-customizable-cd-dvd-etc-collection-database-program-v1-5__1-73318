VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBuildStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Declare some RtlMoveMemory Alias's (thanks Bruce :)
Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)
Private Declare Function AllocStrBPtr Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal lAddrPtr As Long, ByVal lCount As Long) As Long

Private Const lChunk As Long = &H2000
Private Const lZero As Long = &H0

Private lTotal As Long
Private lUBound As Long
Private aBuffer() As Byte

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Property Get Value(Optional ByVal lStart As Long, Optional ByVal lLength As Long) As String
    Dim lCount As Long
    lStart = ValidStartPos(lStart)
    lCount = lTotal - lStart
    If Not (lCount > lZero) Then Exit Property
    lCount = ValidByteLength(lLength, lCount)
    CopyMemByR ByVal VarPtr(Value), AllocStrBPtr(VarPtr(aBuffer(lZero)) + lStart, lCount), 4&
End Property

Public Sub Append(sSubStr As String, Optional ByVal lLength As Long)
    Dim lCount As Long, lSeek As Long
    lCount = ValidByteLength(lLength, LenB(sSubStr))
    If (lCount = lZero) Then Exit Sub
    lSeek = lTotal
    lTotal = lSeek + lCount
    If lTotal > lUBound Then
        lUBound = lTotal + lChunk
        ReDim Preserve aBuffer(lUBound) As Byte
    End If
    ' Preserve Unicode by passing StrPtr and byte count
    CopyMemByV VarPtr(aBuffer(lSeek)), StrPtr(sSubStr), lCount
End Sub

'----------------------------------------------------------
' Optionally specify the delimiter character(s) to insert
' between the appended substrings. It will work correctly
' when arguments are omitted or passed empty:
'    Appends "s1",,, "s2", "", "", "s3",, vbCrLf
'    Appends "", "", "s4",,, "s5",, "", vbCrLf
'    Appends "",, "s6", "", "", "",, "s7" vbCrLf
'----------------------------------------------------------

Public Sub Appends(sSubStr As String, Optional sSubStr2 As String, Optional sSubStr3 As String, Optional sSubStr4 As String, Optional sSubStr5 As String, Optional sSubStr6 As String, Optional sSubStr7 As String, Optional sSubStr8 As String, Optional sDelim As String)
    Dim lDelim As Long, cDelim As Long
    Dim Len1 As Long, Len2 As Long
    Dim Len3 As Long, Len4 As Long
    Dim Len5 As Long, Len6 As Long
    Dim Len7 As Long, Len8 As Long
    Dim lBytes As Long

    Len1 = LenB(sSubStr):  Len2 = LenB(sSubStr2)
    Len3 = LenB(sSubStr3): Len4 = LenB(sSubStr4)
    Len5 = LenB(sSubStr5): Len6 = LenB(sSubStr6)
    Len7 = LenB(sSubStr7): Len8 = LenB(sSubStr8)
    lDelim = LenB(sDelim)

    If (lDelim) Then
        If (lTotal = lZero) Then cDelim = -lDelim
        If (Len1) Then cDelim = cDelim + lDelim
        If (Len2) Then cDelim = cDelim + lDelim
        If (Len3) Then cDelim = cDelim + lDelim
        If (Len4) Then cDelim = cDelim + lDelim
        If (Len5) Then cDelim = cDelim + lDelim
        If (Len6) Then cDelim = cDelim + lDelim
        If (Len7) Then cDelim = cDelim + lDelim
        If (Len8) Then cDelim = cDelim + lDelim
    End If

    lBytes = Len1 + Len2 + Len3 + Len4 + Len5 + Len6 + Len7 + Len8 + cDelim

    If lBytes + lTotal > lUBound Then
        lUBound = lBytes + lTotal + lChunk
        ReDim Preserve aBuffer(lUBound) As Byte
    End If

    ' Preserve Unicode by passing StrPtr and byte count
    If (Len1) Then
        If (lDelim) Then If (lTotal) Then GoSub InsDelim
        CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sSubStr), Len1
        lTotal = lTotal + Len1
    End If
    If (Len2) Then
        If (lDelim) Then If (lTotal) Then GoSub InsDelim
        CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sSubStr2), Len2
        lTotal = lTotal + Len2
    End If
    If (Len3) Then
        If (lDelim) Then If (lTotal) Then GoSub InsDelim
        CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sSubStr3), Len3
        lTotal = lTotal + Len3
    End If
    If (Len4) Then
        If (lDelim) Then If (lTotal) Then GoSub InsDelim
        CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sSubStr4), Len4
        lTotal = lTotal + Len4
    End If
    If (Len5) Then
        If (lDelim) Then If (lTotal) Then GoSub InsDelim
        CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sSubStr5), Len5
        lTotal = lTotal + Len5
    End If
    If (Len6) Then
        If (lDelim) Then If (lTotal) Then GoSub InsDelim
        CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sSubStr6), Len6
        lTotal = lTotal + Len6
    End If
    If (Len7) Then
        If (lDelim) Then If (lTotal) Then GoSub InsDelim
        CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sSubStr7), Len7
        lTotal = lTotal + Len7
    End If
    If (Len8) Then
        If (lDelim) Then If (lTotal) Then GoSub InsDelim
        CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sSubStr8), Len8
        lTotal = lTotal + Len8
    End If
    Exit Sub

InsDelim:
    CopyMemByV VarPtr(aBuffer(lTotal)), StrPtr(sDelim), lDelim
    lTotal = lTotal + lDelim
    Return
End Sub

Public Property Get Length() As Long
    Length = lTotal \ 2&
End Property

Public Sub Reset(Optional ByVal FreeMemory As Boolean)
    lTotal = lZero
    If FreeMemory Then
        lUBound = lChunk
        ReDim aBuffer(lUBound) As Byte
    End If
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Class_Initialize()
    Call Reset(True)
End Sub

Private Function ValidStartPos(ByVal lStart As Long) As Long
    If (lStart > lZero) Then
       ValidStartPos = (lStart + lStart) - 2 ' Byte count, zero based
       If Not (ValidStartPos < lTotal) Then ValidStartPos = lTotal
    End If
End Function

Private Function ValidByteLength(ByVal lLength As Long, ByVal LenB_SubStr As Long) As Long
    ValidByteLength = LenB_SubStr
    If (lLength > lZero) Then
       lLength = lLength + lLength ' Unicode byte count
       If Not (lLength > LenB_SubStr) Then ValidByteLength = lLength
    End If
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Rd - crYptic but cRaZy!
