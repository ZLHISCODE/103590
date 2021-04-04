Attribute VB_Name = "mdlBase64Code"
Option Explicit
Option Base 0

Private mbytB64Enc(63) As Byte
Private mbytBuffer1() As Byte
Private mbytBuffer2() As Byte
Private mstrContent   As String

Public Function getBase64Code(ByVal strPicturePath As String) As String
          '将图片转换为base64编码
          'strPicturePath                 图片路径
          Dim objFSO As New FileSystemObject
          
          '获取图片类型
1         On Error GoTo getBase64Code_Error

2         If objFSO.FileExists(strPicturePath) = False Then
3             objFSO.CreateTextFile (strPicturePath)
              
4         End If
      '    strFileType = Replace(objFile.Type, " 文件", "")
          
5         Call Class_Initialize
6         Call LoadPicture(strPicturePath, mbytBuffer1)
7         Call Encode(mbytBuffer1, mbytBuffer2)
8         Call ByteArrayToString(mbytBuffer2, mstrContent)
9         getBase64Code = mstrContent


10        Exit Function
getBase64Code_Error:
11        Call writeErrLog("ZL9LabWork", "mdlBase64Code", "执行(getBase64Code)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
12        Err.Clear

End Function

Private Sub LoadPicture(ByVal strPathName As String, ByRef pbArrayOutput() As Byte)
          '读取图片
          Dim lngSize     As Long
          Dim intFreeFile As Integer
1         On Error GoTo LoadPicture_Error

2         lngSize = FileLen(strPathName)
3         intFreeFile = FreeFile
4         ReDim pbArrayOutput(lngSize - 1)
5         Open strPathName For Binary As intFreeFile
6         Get intFreeFile, , pbArrayOutput
7         Close intFreeFile


8         Exit Sub
LoadPicture_Error:
9         Call writeErrLog("ZL9LabWork", "mdlBase64Code", "执行(LoadPicture)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
10        Err.Clear

End Sub

Private Sub ByteArrayToString(ByRef pbArrayInput() As Byte, ByRef strOut As String)
   strOut = StrConv(pbArrayInput, vbUnicode)
End Sub

Private Sub Encode(ByRef pbArrayInput() As Byte, ByRef pbArrayOutput() As Byte)
          '生成base64编码
          Dim intSizeMod As Integer
          Dim lngSizeIn  As Long
          Dim lngSizeOut As Long
          Dim lngIndex    As Long
          Dim lngIndex2  As Long
          Dim lngotal   As Long
          Dim bBuffer(2) As Byte
1         On Error GoTo Encode_Error

2         lngSizeIn = UBound(pbArrayInput) + 1
3         intSizeMod = lngSizeIn Mod 3
4         lngSizeOut = ((lngSizeIn - intSizeMod) \ 3) * 4
5         If intSizeMod > 0 Then lngSizeOut = lngSizeOut + 4


6         ReDim pbArrayOutput(lngSizeOut - 1)


7         If lngSizeIn >= 3 Then

8             lngotal = lngSizeIn - intSizeMod - 1
9             For lngIndex = 0 To lngotal Step 3

10                bBuffer(0) = pbArrayInput(lngIndex)
11                bBuffer(1) = pbArrayInput(lngIndex + 1)
12                bBuffer(2) = pbArrayInput(lngIndex + 2)
13                pbArrayOutput(lngIndex2) = mbytB64Enc((bBuffer(0) And &HFC) \ 4)
14                pbArrayOutput(lngIndex2 + 1) = mbytB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
15                pbArrayOutput(lngIndex2 + 2) = mbytB64Enc((bBuffer(1) And &HF) * 4 Or (bBuffer(2) And &HC0) \ 64)
16                pbArrayOutput(lngIndex2 + 3) = mbytB64Enc((bBuffer(2) And &H3F))
17                lngIndex2 = lngIndex2 + 4
18            Next
19        End If


20        Select Case intSizeMod
          Case 1
21            bBuffer(0) = pbArrayInput(lngSizeIn - 1)
22            bBuffer(1) = 0
23            pbArrayOutput(lngIndex2) = mbytB64Enc((bBuffer(0) And &HFC) \ 4)
24            pbArrayOutput(lngIndex2 + 1) = mbytB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
25            pbArrayOutput(lngIndex2 + 2) = 61
26            pbArrayOutput(lngIndex2 + 3) = 61
27        Case 2
28            bBuffer(0) = pbArrayInput(lngSizeIn - 2)
29            bBuffer(1) = pbArrayInput(lngSizeIn - 1)
30            bBuffer(2) = 0
31            pbArrayOutput(lngIndex2) = mbytB64Enc((bBuffer(0) And &HFC) \ 4)
32            pbArrayOutput(lngIndex2 + 1) = mbytB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
33            pbArrayOutput(lngIndex2 + 2) = mbytB64Enc((bBuffer(1) And &HF) * 4 Or (bBuffer(2) And &HC0) \ 64)
34            pbArrayOutput(lngIndex2 + 3) = 61
35        End Select


36        Exit Sub
Encode_Error:
37        Call writeErrLog("ZL9LabWork", "mdlBase64Code", "执行(Encode)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
38        Err.Clear

End Sub


Private Sub Class_Initialize()
          Dim iIndex As Integer
1         On Error GoTo Class_Initialize_Error

2         For iIndex = 65 To 90
3             mbytB64Enc(iIndex - 65) = iIndex
4         Next
5         For iIndex = 97 To 122
6             mbytB64Enc(iIndex - 71) = iIndex
7         Next
8         For iIndex = 48 To 57
9             mbytB64Enc(iIndex + 4) = iIndex
10        Next
11        mbytB64Enc(62) = 43
12        mbytB64Enc(63) = 47


13        Exit Sub
Class_Initialize_Error:
14        Call writeErrLog("ZL9LabWork", "mdlBase64Code", "执行(Class_Initialize)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
15        Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' 编    码:蔡青松
' 编码日期:2017-2-27
' 功    能:将base64编码解析成图片
' 入    参:
'           strImgPath      图片保存路径
'           strBase64Code   base64编码串
' 出    参:
' 返    回:完整图片路径
' 修 改 人:
' 修改日期:
'---------------------------------------------------------------------------------------
Public Function getBase64Img(ByVal strFolder As String, ByVal strImgPath As String, ByVal strBase64Code As String) As String
          '写文件
          Dim objFile As New FileSystemObject
          Dim bBase64Code() As Byte

1         On Error GoTo getBase64Img_Error

2         bBase64Code = Base64Decode(strBase64Code)
3         If Not objFile.FolderExists(strFolder) Then
4             objFile.CreateFolder (strFolder)
5         End If
6         If objFile.FileExists(strFolder) = True Then
7             objFile.DeleteFile (strImgPath)
8         End If
9         objFile.CreateTextFile (strImgPath)
10        Open strImgPath For Binary As #1
11        Put #1, , bBase64Code
12        Close #1
          
13        getBase64Img = strImgPath


14        Exit Function
getBase64Img_Error:
15        Call writeErrLog("ZL9LabWork", "mdlBase64Code", "执行(getBase64Img)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
16        Err.Clear

End Function


'VB Base64 解码/解密函数：
Function Base64Decode(B64 As String) As Byte()                                  'Base64 解码
                                                                 '排错
          Dim length As Long, mods As Long
          Dim OutStr() As Byte, i As Long, j As Long
          Dim buf(3) As Byte
          
          
          Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
1         On Error GoTo Base64Decode_Error

2         If InStr(1, B64, "=") <> 0 Then B64 = Left(B64, InStr(1, B64, "=") - 1)     '判断Base64真实长度,除去补位
3         mods = Len(B64) Mod 4
4         length = Len(B64) - mods
5         ReDim OutStr(length / 4 * 3 - 1 + Switch(mods = 0, 0, mods = 2, 1, mods = 3, 2))
6         For i = 1 To length Step 4
              
7             For j = 0 To 3
8                 buf(j) = InStr(1, B64_CHAR_DICT, Mid(B64, i + j, 1)) - 1            '根据字符的位置取得索引值
9             Next
10            OutStr((i - 1) / 4 * 3) = buf(0) * &H4 + (buf(1) And &H30) / &H10
11            OutStr((i - 1) / 4 * 3 + 1) = (buf(1) And &HF) * &H10 + (buf(2) And &H3C) / &H4
12            OutStr((i - 1) / 4 * 3 + 2) = (buf(2) And &H3) * &H40 + buf(3)
13        Next
14        If mods = 2 Then
15            OutStr(length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &H30) / 16
16        ElseIf mods = 3 Then
17            OutStr(length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &H30) / 16
18            OutStr(length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 3, 1)) - 1) And &H3C) / &H4
19        End If
20        Base64Decode = OutStr                                                       '读取解码结果


21        Exit Function
Base64Decode_Error:
22        Call writeErrLog("ZL9LabWork", "mdlBase64Code", "执行(Base64Decode)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
23        Err.Clear

End Function


