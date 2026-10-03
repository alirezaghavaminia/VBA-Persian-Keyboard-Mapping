Attribute VB_Name = "modPersianMapping"
' ------------------------------------------------------------------
' By Alireza Ghavaminia:
' https://github.com/alirezaghavaminia/VBA-Persian-Keyboard-Mapping
' ------------------------------------------------------------------

Option Compare Database
Option Explicit

' Generates a VBA code snippet (as text) that builds the Farsi string via ChrW,
' so you can paste it directly into your source without Unicode corruption.
Public Function GenerateChrWCode(ByVal sText As String) As String
    Dim faString As String
    Dim i As Long
    Dim code As Long
    Dim parts As String

    faString = EnglishToFarsi(sText)

    parts = ""
    For i = 1 To Len(faString)
        code = AscW(Mid$(faString, i, 1))
        If code < 0 Then code = code + 65536  ' handle negative AscW values
        If i > 1 Then parts = parts & " & "
        parts = parts & "ChrW(" & code & ")"
    Next i

    GenerateChrWCode = parts
End Function


' ==============================================================
'                       Complete mapping
' ==============================================================
' Converts text typed using US/English key *positions* into real Farsi
' Unicode text, per the Windows "Persian (Standard)" / ISIRI 9147 layout.
'
' Level 1 = unshifted keys (lowercase letters, bare punctuation)
' Level 2 = shifted keys   (uppercase letters, shifted punctuation)
'
' Half-space (ZWNJ, U+200C):
'   On the real keyboard, ZWNJ is produced by BOTH Shift+B and Shift+Space.
'   A plain space looks identical whether or not Shift was held, so this
'   function can't tell those apart -- Shift+Space is therefore NOT
'   converted. Use Shift+B ("B") to insert a half-space. A literal ZWNJ
'   character already present in sText passes through unchanged.
Public Function EnglishToFarsi(ByVal sText As String) As String
    Dim keysL1 As String, keysL2 As String
    Dim codesL1 As Variant, codesL2 As Variant
    Dim i As Long, p As Long
    Dim ch As String, result As String

    ' ---- Level 1: unshifted keys (order must match codesL1) ----
    keysL1 = "`1234567890-=" & _
             "qwertyuiop[]" & _
             "asdfghjkl;'" & _
             "zxcvbnm,./" & _
             "\"
    codesL1 = Array( _
        8205, 1777, 1778, 1779, 1780, 1781, 1782, 1783, 1784, 1785, 1776, 45, 61, _
        1590, 1589, 1579, 1602, 1601, 1594, 1593, 1607, 1582, 1581, 1580, 1670, _
        1588, 1587, 1740, 1576, 1604, 1575, 1578, 1606, 1605, 1705, 1711, _
        1592, 1591, 1586, 1585, 1584, 1583, 1662, 1608, 46, 47, _
        92)
        ' ` 1 2 3 4 5 6 7 8 9 0 - =
        ' q w e r t y u i o p [ ]
        ' a s d f g h j k l ; '
        ' z x c v b n m , . /
        ' \
        
        
    ' ---- Level 2: shifted keys (order must match codesL2) ----
    keysL2 = "~!@#$%^&*()_+" & _
             "QWERTYUIOP{}" & _
             "ASDFGHJKL:" & Chr$(34) & _
             "ZXCVBNM<>?" & _
             "|"
    codesL2 = Array( _
        247, 33, 1644, 1643, 65020, 1642, 215, 1548, 42, 41, 40, 1600, 43, _
        1618, 1612, 1613, 1611, 1615, 1616, 1614, 1617, 93, 91, 125, 123, _
        1572, 1574, 1610, 1573, 1571, 1570, 1577, 187, 171, 58, 1563, _
        1603, 1619, 1688, 1648, 8204, 1620, 1569, 62, 60, 1567, _
        124)
        
        ' ~ ! @ # $ % ^ & * ( ) _ +
        ' Q W E R T Y U I O P { }
        ' A S D F G H J K L : "
        ' Z X C V B N M < > ?
        ' |
        
    result = ""
    For i = 1 To Len(sText)
        ch = Mid$(sText, i, 1)
        p = InStr(1, keysL1, ch, vbBinaryCompare)
        If p > 0 Then
            result = result & FaChr(codesL1(p - 1))
        Else
            p = InStr(1, keysL2, ch, vbBinaryCompare)
            If p > 0 Then
                result = result & FaChr(codesL2(p - 1))
            Else
                ' anything not on the layout (spaces, other symbols): keep as-is
                result = result & ch
            End If
        End If
    Next i
    EnglishToFarsi = result
End Function

' ChrW$ only accepts values in the signed-Integer range (-32768 to 32767).
' Code points above &H7FFF (e.g. the Rial sign, U+FDFC = 65020) must be
' passed as their negative twos-complement equivalent, or VBA raises
' "Run-time error 5: Invalid procedure call or argument".
Private Function FaChr(ByVal codePoint As Long) As String
    If codePoint > 32767 Then codePoint = codePoint - 65536
    FaChr = ChrW$(codePoint)
End Function


Public Function Farsi(ByVal sText As String) As String
    Farsi = EnglishToFarsi(sText)
End Function
