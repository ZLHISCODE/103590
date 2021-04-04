Attribute VB_Name = "mdlFunctions"
Option Explicit
'Ä£¿é¹¦ÄÜ£º³¬´óÊý×Ö10½øÖÆ16½øÖÆ2½øÖÆ¼äµÄÏà»¥×ª»»
Public Const HEX_TO_DEC    As Long = 1
Public Const HEX_TO_BIN    As Long = 2
Public Const DEC_TO_HEX    As Long = 3
Public Const DEC_TO_BIN    As Long = 4
Public Const BIN_TO_DEC    As Long = 5
Public Const BIN_TO_HEX    As Long = 6

Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Const CP_UTF8 = 65001
Public Const m_bIsNt = True

'Purpose:Convert Utf8 to Unicode
Public Function UTF8_Decode(ByVal sUTF8 As String) As String
   Dim lngUtf8Size      As Long
   Dim strBuffer        As String
   Dim lngBufferSize    As Long
   Dim lngResult        As Long
   Dim bytUtf8()        As Byte
   Dim n                As Long

   If LenB(sUTF8) = 0 Then Exit Function

   If m_bIsNt Then
      On Error GoTo EndFunction
      bytUtf8 = StrConv(sUTF8, vbFromUnicode)
      lngUtf8Size = UBound(bytUtf8) + 1
      On Error GoTo 0
      'Set buffer for longest possible string i.e. each byte is
      'ANSI, thus 1 unicode(2 bytes)for every utf-8 character.
      lngBufferSize = lngUtf8Size * 2
      strBuffer = String$(lngBufferSize, vbNullChar)
      'Translate using code page 65001(UTF-8)
      lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), _
         lngUtf8Size, StrPtr(strBuffer), lngBufferSize)
      'Trim result to actual length
      If lngResult Then
         UTF8_Decode = Left$(strBuffer, lngResult)
      End If
   Else
      Dim i                As Long
      Dim TopIndex         As Long
      Dim TwoBytes(1)      As Byte
      Dim ThreeBytes(2)    As Byte
      Dim AByte            As Byte
      Dim TStr             As String
      Dim BArray()         As Byte

      'Resume on error in case someone inputs text with accents
      'that should have been encoded as UTF-8
      On Error Resume Next

      TopIndex = Len(sUTF8)  ' Number of bytes equal TopIndex+1
      If TopIndex = 0 Then Exit Function ' get out if there's nothing to convert
      BArray = StrConv(sUTF8, vbFromUnicode)
      i = 0 ' Initialise pointer
      TopIndex = TopIndex - 1
      ' Iterate through the Byte Array
      Do While i <= TopIndex
         AByte = BArray(i)
         If AByte < &H80 Then
            ' Normal ANSI character - use it as is
            TStr = TStr & Chr$(AByte): i = i + 1 ' Increment byte array index
         ElseIf AByte >= &HE0 Then         'was = &HE1 Then
            ' Start of 3 byte UTF-8 group for a character
            ' Copy 3 byte to ThreeBytes
            ThreeBytes(0) = BArray(i): i = i + 1
            ThreeBytes(1) = BArray(i): i = i + 1
            ThreeBytes(2) = BArray(i): i = i + 1
            ' Convert Byte array to UTF-16 then Unicode
            TStr = TStr & ChrW$((ThreeBytes(0) And &HF) * &H1000 + (ThreeBytes(1) And &H3F) * &H40 + (ThreeBytes(2) And &H3F))
         ElseIf (AByte >= &HC2) And (AByte <= &HDB) Then
            ' Start of 2 byte UTF-8 group for a character
            TwoBytes(0) = BArray(i): i = i + 1
            TwoBytes(1) = BArray(i): i = i + 1
            ' Convert Byte array to UTF-16 then Unicode
            TStr = TStr & ChrW$((TwoBytes(0) And &H1F) * &H40 + (TwoBytes(1) And &H3F))
         Else
            ' Normal ANSI character - use it as is
            TStr = TStr & Chr$(AByte): i = i + 1 ' Increment byte array index
         End If
      Loop
      UTF8_Decode = TStr    ' Return the resultant string
      Erase BArray
   End If

EndFunction:

End Function

'Purpose:Convert Unicode string to UTF-8.
Public Function UTF8_Encode(ByVal strUnicode As String, Optional ByVal bHTML As Boolean = True) As String
   Dim i                As Long
   Dim TLen             As Long
   Dim lPtr             As Long
   Dim UTF16            As Long
   Dim UTF8_EncodeLong  As String

   TLen = Len(strUnicode)
   If TLen = 0 Then Exit Function

   If m_bIsNt Then
      Dim lngBufferSize    As Long
      Dim lngResult        As Long
      Dim bytUtf8()        As Byte
      'Set buffer for longest possible string.
      lngBufferSize = TLen * 3 + 1
      ReDim bytUtf8(lngBufferSize - 1)
      'Translate using code page 65001(UTF-8).
      lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), _
         TLen, bytUtf8(0), lngBufferSize, vbNullString, 0)
      'Trim result to actual length.
      If lngResult Then
         lngResult = lngResult - 1
         ReDim Preserve bytUtf8(lngResult)
         'CopyMemory StrPtr(UTF8_Encode), bytUtf8(0&), lngResult
         UTF8_Encode = StrConv(bytUtf8, vbUnicode)
         ' For i = 0 To lngResult
         '    UTF8_Encode = UTF8_Encode & Chr$(bytUtf8(i))
         ' Next
      End If
   Else
      For i = 1 To TLen
         ' Get UTF-16 value of Unicode character
         lPtr = StrPtr(strUnicode) + ((i - 1) * 2)
         CopyMemory UTF16, ByVal lPtr, 2
         'Convert to UTF-8
         If UTF16 < &H80 Then                                      ' 1 UTF-8 byte
            UTF8_EncodeLong = Chr$(UTF16)
         ElseIf UTF16 < &H800 Then                                 ' 2 UTF-8 bytes
            UTF8_EncodeLong = Chr$(&H80 + (UTF16 And &H3F))              ' Least Significant 6 bits
            UTF16 = UTF16 \ &H40                                   ' Shift right 6 bits
            UTF8_EncodeLong = Chr$(&HC0 + (UTF16 And &H1F)) & UTF8_EncodeLong  ' Use 5 remaining bits
         Else                                                      ' 3 UTF-8 bytes
            UTF8_EncodeLong = Chr$(&H80 + (UTF16 And &H3F))              ' Least Significant 6 bits
            UTF16 = UTF16 \ &H40                                   ' Shift right 6 bits
            UTF8_EncodeLong = Chr$(&H80 + (UTF16 And &H3F)) & UTF8_EncodeLong  ' Use next 6 bits
            UTF16 = UTF16 \ &H40                                   ' Shift right 6 bits
            UTF8_EncodeLong = Chr$(&HE0 + (UTF16 And &HF)) & UTF8_EncodeLong   ' Use 4 remaining bits
         End If
         UTF8_Encode = UTF8_Encode & UTF8_EncodeLong
      Next
   End If

   'Substitute vbCrLf with HTML line breaks if requested.
   If bHTML Then
      UTF8_Encode = Replace$(UTF8_Encode, vbCrLf, "<br/>")
   End If

End Function


'Ê®½øÖÆ  ¡ú  Ê®Áù½øÖÆ
Function ToHex(DecStr As String) As String
    Dim i   As Long, j    As Long, tmp    As String
    Do While Len(DecStr) > 9
        ToHex = Hex(Val(Right$(DecStr, 4)) Mod 16) & ToHex
        For i = 1 To 4
            tmp = "0" & DecStr:      DecStr = ""
            For j = 2 To Len(tmp)
                DecStr = DecStr & CStr(Val(Mid$(tmp, j, 1)) \ 2 + _
                    IIf(Val(Mid$(tmp, j - 1, 1)) Mod 2, 5, 0))
            Next j
            If Left$(DecStr, 1) = "0" Then DecStr = Right$(DecStr, Len(DecStr) - 1)
        Next i
    Loop
    ToHex = Hex(Val(DecStr)) & ToHex
End Function

'Ê®Áù½øÖÆ  ¡ú  ¶þ½øÖÆ
Function ToBin(HexStr As String) As String
    Dim i   As Long
    Const tmp   As String = "0000000100100011010001010110011110001001101010111100110111101111"
    For i = 1 To Len(HexStr)
        ToBin = ToBin & Mid$(tmp, (Val("&H" & Mid$(HexStr, i, 1)) + 1) * 4 - 3, 4)
    Next i
    Dim P1   As Long:   P1 = InStr(ToBin, "1")
    If P1 Then ToBin = Right$(ToBin, Len(ToBin) - P1 + 1) Else ToBin = "0"
End Function

'¶þ½øÖÆ  ¡ú  Ê®½øÖÆ
Function ToDec(BinStr As String) As String
    Dim i   As Long, j    As Long, tmp    As String
    ToDec = "0"
    For i = 1 To Len(BinStr)
        ToDec = "0" & ToDec:      tmp = "0"
        For j = 2 To Len(ToDec)
            If Val(Mid$(ToDec, j, 1)) >= 5 Then tmp = Left$(tmp, Len(tmp) - 1) & CStr(Val(Right$(tmp, 1)) + 1)
            tmp = tmp & (Val(Mid$(ToDec, j, 1)) Mod 5) * 2
        Next j
        If Left$(tmp, 1) = "0" Then tmp = Right$(tmp, Len(tmp) - 1)
        ToDec = tmp
        If Mid$(BinStr, i, 1) = "1" Then ToDec = Left$(ToDec, Len(ToDec) - 1) & CStr(Val(Right$(ToDec, 1)) + 1)
    Next i
End Function

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§¡¡¡¡¡¡¡¡¡¡¡¡¡¡10¡ú16¡ú2  ¡¡¡¡¡¡¡¡¡¡¡¡¡¡16¡¡¡¡¡¡¡¡¡¡¡¡¡¡©§
'©§¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡ü            ©¦  ¡¡¡¡¡¡¡¡¡¡¡¡¨L¨I  ¡¡¡¡¡¡¡¡¡¡¡¡©§
'©§¡¡¡¡¡¡¡¡¡¡¡¡¡¡©¸©¤©¤©¤©¼¡¡¡¡¡¡¡¡¡¡¡¡2  ¡ú10¡¡¡¡¡¡¡¡¡¡¡¡©§
'©Ä©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ì
'©§Í¨¹ýÒÔÉÏ3¸öº¯Êý£¬ÒÑ¾­¿ÉÒÔÔÚ2½øÖÆ10½øÖÆ16½øÖÆ¼ä×ÔÓÉ×ª»»©§
'©§µ«2½øÖÆ×ª16½øÖÆÊ±µÄÐ§ÂÊ¼«µÍ£¬ÓÚÊÇÓÖÐ´ÁËÒ»¸öToHex_Bº¯Êý©§
'©§ÔÚ×ª»»³¬´óÊý×ÖÊ±£¬ToHex_B()Òª±ÈToHex(ToDec())¿ìºÜ¶à±¶  ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿
Public Function NumConv(ByVal NumStr As String, Mode As Long) As String
    Select Case Mode
        Case 1:   NumConv = ToDec(ToBin(NumStr))        '  HexToDec
        Case 2:   NumConv = ToBin(NumStr)                      '  HexToBin
        Case 3:   NumConv = ToHex(NumStr)                      '  DecToHex
        Case 4:   NumConv = ToBin(ToHex(NumStr))        '  DecToBin
        Case 5:   NumConv = ToDec(NumStr)                      '  BinToDec
        Case 6:   NumConv = ToHex_B(NumStr)                  '  BinToHex
        Case Else:   NumConv = NumStr
    End Select
End Function

'¶þ½øÖÆ  ¡ú  Ê®Áù½øÖÆ
Function ToHex_B(BinStr As String) As String
    Dim i   As Long
    BinStr = String((Len(BinStr) \ 4 + IIf(Len(BinStr) Mod 4, 1, 0)) * 4 - Len(BinStr), "0") & BinStr
    For i = 0 To Len(BinStr) \ 4 - 1
        Select Case Mid$(BinStr, i * 4 + 1, 4)
            Case "0000":   ToHex_B = ToHex_B & "0"
            Case "0001":   ToHex_B = ToHex_B & "1"
            Case "0010":   ToHex_B = ToHex_B & "2"
            Case "0011":   ToHex_B = ToHex_B & "3"
            Case "0100":   ToHex_B = ToHex_B & "4"
            Case "0101":   ToHex_B = ToHex_B & "5"
            Case "0110":   ToHex_B = ToHex_B & "6"
            Case "0111":   ToHex_B = ToHex_B & "7"
            Case "1000":   ToHex_B = ToHex_B & "8"
            Case "1001":   ToHex_B = ToHex_B & "9"
            Case "1010":   ToHex_B = ToHex_B & "A"
            Case "1011":   ToHex_B = ToHex_B & "B"
            Case "1100":   ToHex_B = ToHex_B & "C"
            Case "1101":   ToHex_B = ToHex_B & "D"
            Case "1110":   ToHex_B = ToHex_B & "E"
            Case "1111":   ToHex_B = ToHex_B & "F"
           End Select
    Next i
End Function

Public Function ASCToStr(strIn As String) As String
'Ä¿Ç°»¹ÓÐÎÊÌâ£¬Ö÷ÒªÊÇ²»ÄÜÓÃReplace¡£²»¹ýÕâ¸öº¯ÊýÄ¿Ç°²»»áÊ¹ÓÃµ½¡£Ö»ÓÃÓÚ²âÊÔ·ÖÎöÎÄ±¾ÓÃ£¡
On Error Resume Next
    'ÏÈ´¦ÀíÌØÊâ×Ö·û×ªÒåÐòÁÐ
    strIn = Replace(strIn, "\}", "}")
    strIn = Replace(strIn, "\\", "\")
    strIn = Replace(strIn, "\{", "{")
    strIn = Replace(strIn, "\TAB", vbTab)
    strIn = Replace(strIn, "\par", vbCrLf)

    '½«ÖÐÎÄ×Ö·û´®×ª»»ÎªASC´®£¨°üÀ¨Ó¢ÎÄÒ»Æð£©
    Dim i As Long, CurPos As Long, strResult As String, strASC As String, strStr As String, strTMP As String
    Dim Part1 As String, Part2 As String, InPos As String
    CurPos = InStr(1, strIn, "\'", vbTextCompare)
    Do While CurPos > 0
        strTMP = Mid(strIn, CurPos, 4)
        strASC = Replace(strTMP, "\'", "")
        strStr = Chr(NumConv(strASC, HEX_TO_DEC))
        If NumConv(strASC, HEX_TO_DEC) < 32 Then  '¿ØÖÆ×Ö·û
            'ÕâÀï²»ÄÜÓÃReplace£¬ÔÝÊ±²ÉÓÃ£¬ÒÔºóÒª¸ÄÕý¡£
            strIn = Replace(strIn, strTMP, ChrW(NumConv(strASC, HEX_TO_DEC)))
            CurPos = InStr(CurPos + 1, strIn, "\'", vbTextCompare)
        ElseIf NumConv(strASC, HEX_TO_DEC) < 128 Then  '²»Ðè×ªÒåµÄ×Ö·û
            'ÕâÀï²»ÄÜÓÃReplace£¬ÔÝÊ±²ÉÓÃ£¬ÒÔºóÒª¸ÄÕý¡£
            strIn = Replace(strIn, strTMP, strStr)
            CurPos = InStr(CurPos + 1, strIn, "\'", vbTextCompare)
        Else
            'ÕâÀï²»ÄÜÓÃReplace£¬ÔÝÊ±²ÉÓÃ£¬ÒÔºóÒª¸ÄÕý¡£
            '¸ßÎ»ÔòÁ½¸ö×Ö½Ú±íÊ¾Ò»¸ö×Ö·û
            strTMP = Mid(strIn, CurPos, 8)
            If Mid(strTMP, 5, 2) <> "\'" Then
                strIn = Replace(strIn, strTMP, ChrW(NumConv(strASC, HEX_TO_DEC)) & Mid(strTMP, 5, Len(strTMP) - 4))
                CurPos = InStr(CurPos + Len(strTMP) - 3, strIn, "\'", vbTextCompare)
            Else
                strASC = Replace(strTMP, "\'", "")
                strStr = Chr(NumConv(strASC, HEX_TO_DEC))
                strIn = Replace(strIn, strTMP, strStr)
                CurPos = InStr(CurPos + 1, strIn, "\'", vbTextCompare)
            End If
        End If
    Loop
    ASCToStr = strIn
End Function


Public Function StrToASC(ByVal strIn As String) As String
    '½«ÖÐÎÄ×Ö·û´®×ª»»ÎªASC´®£¨°üÀ¨Ó¢ÎÄÒ»Æð£©
    'ÏÈ½«ÌØÊâ×Ö·û½øÐÐ×ªÒå£º
    strIn = Replace(strIn, Chr(9), "\TAB ")
    strIn = Replace(strIn, Chr(13) + Chr(10), "\par ")
    
    Dim i As Long, s As String, lsChar As String, lsPart1 As String, lsPart2 As String
    Dim lsCharHex As String
    For i = 1 To Len(strIn)
        lsChar = Mid(strIn, i, 1)
        If lsChar = "·"" Then
            lsCharHex = LCase(Hex(Asc(lsChar)))
            If Len(lsCharHex) = 4 Then
                lsCharHex = "\'" + Mid(lsCharHex, 1, 2) + "\'" + Mid(lsCharHex, 3, 2)
            Else
                lsCharHex = lsChar
            End If
            s = s + lsCharHex
        Else
            lsCharHex = LCase(Hex(Asc(lsChar)))
            If Len(lsCharHex) = 4 Then
                lsCharHex = "\'" + Mid(lsCharHex, 1, 2) + "\'" + Mid(lsCharHex, 3, 2)
            Else
                lsCharHex = lsChar
            End If
            s = s + lsCharHex
        End If
    Next
    StrToASC = s
End Function

Public Function PicToASC(ByVal strFileName As String) As String
    '»ñÈ¡Í¼Æ¬Êý¾Ý£¬×ª»»ÎªASC×Ö·û´®£¨Ìá¹©Í¼Æ¬µÄÎÄ¼þÃû£©¡£
    Dim bData() As Byte
    Dim i As Long
    Dim lNum As Long
    
    
    Dim strData As String, strTMP As String
    
    lNum = FreeFile
    Open strFileName For Binary As #lNum
    ReDim bData(LOF(lNum) - 1)
    Get #lNum, , bData
    Close #lNum

    strData = Space((UBound(bData) + 1) * 2)    'ÏÈ·ÖÅä¿Õ¼ä£¬È»ºóÔÙ´¦Àí£¡£¡£¡
    For i = 0 To UBound(bData)
        strTMP = Hex$(bData(i))
        If Len(strTMP) = 1 Then
            strTMP = "0" + strTMP
        End If
        Mid(strData, i * 2 + 1) = strTMP
    Next
    PicToASC = strData
    
End Function

Public Function LinkRTF(ByVal strFirst As String, ByVal strMid As String, Optional ByVal strEnd As String = "}") As String
'Á¬½Ó×Ö·û´®
    LinkRTF = strFirst + strMid + strEnd
End Function




