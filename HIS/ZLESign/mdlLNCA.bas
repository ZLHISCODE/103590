Attribute VB_Name = "mdlLNCA"
Option Explicit
' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

Public gobjLNCA As Object   '辽宁CA控件升级到3.1

Public Function LNCAGetGIF() As String

    Dim objWeb As Object
    Dim strBASE64 As String
    
    On Error GoTo hErr
    
    Set objWeb = CreateObject("WEBSIGNATURE.WebSignatureCtrl.1")
    Call objWeb.AddAboutLicense("g2v8Bes5QQ/bVTG1j/MkqJCy/2lYacbqb42K9fkW8t2arAJX3hzGvwBbbSgsthNOVwTVHt+OgbwnRmXWWhTVVAaYvh7I8+mcnJIgIU3zIS0o9c4lwa2xMFtcRalxj2tSHm1ZzLJwfSOBeT1YpOPurl8ygcb7KiBviFxAAi2FuLf11m+rMY7lRbSIsFQa/dUkov5dF4iHEXw=")
    
    strBASE64 = objWeb.ShowSealDialog
    
    If strBASE64 <> "" Then
        LNCAGetGIF = SaveBase64File("BMP", Format(Now, "ddhhmmss"), strBASE64)
    End If
    Exit Function
hErr:
    '--错误，可能是未安装部件
End Function
 
 
Private Function SaveBase64File(ByVal strType As String, ByVal strFileName As String, ByVal str2Decode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
'    Dim frmPic As New frmGraph
    Dim ByteData() As Byte, lngCount As Long, strFileBmp As String, lngFileNum
     
    lngCount = Len(str2Decode)
    ReDim ByteData(lngCount / 4 * 3)
    lngCount = 0
    ' for each 4 character group....
    For lPtr = 1 To Len(str2Decode) Step 4
        iLen = 4
        For iCtr = 0 To 3
            ' retrive the base 64 value, 4 at a time
            iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                ' not found
                Case 0: Exit Function
            End Select
        Next

        ' convert the 4, 6 bit values into 3, 8 bit values
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)

        ' add the three new characters to the output string
        For iCtr = 1 To iLen - 1
            ByteData(lngCount) = bits(iCtr)
            lngCount = lngCount + 1
        Next
    Next
    
    strFileBmp = App.Path & "\" & strFileName & "." & strType
    lngFileNum = FreeFile
    Open strFileBmp For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBase64File = strFileBmp

End Function

