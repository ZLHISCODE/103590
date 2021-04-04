VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Ku2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'用户：龙泉驿区第一人民医院
'厂家：珠海科域生物工程有限公司
'类型：尿沉渣
'日期：2015/9/24   敬秦

Private mstrCode As String
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    Dim strDate As String, strNo As String, strStat As String, strQC As String, strItem As String
    Dim lngS As Long, lngE As Long, i As Integer
    Dim strCode As String, arrValue() As String, strLine As String
    Dim strPicData As String, strPic As String, strPicPath As String
    On Error GoTo hErr
    
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    lngS = InStr(mstrCode, Chr(&HB))
    lngE = InStr(mstrCode, Chr(&H1C))
    Do While lngS > 0 And lngE > 0  '获取一个完整的标本
        strCode = Mid(mstrCode, lngS, lngE - lngS)
        mstrCode = Mid(mstrCode, lngE + 1)
        lngS = InStr(mstrCode, Chr(&HB))
        lngE = InStr(mstrCode, Chr(&H1C))
        Call WriteLog("开始处理数据", strCode, "")
        arrValue = Split(strCode, Chr(&HD))
        For i = 0 To UBound(arrValue)
            strLine = arrValue(i)
            If InStr(strLine, "MSH") > 0 Then
                strDate = Split(strLine, "|")(6)
                strDate = Format(Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Right(strDate, 2), "yyyy-mm-dd hh:mm:ss")
            End If
            If InStr(strLine, "PID") > 0 Then
                strStat = IIf(Split(strLine, "|")(1) = "N", "", "1")
            End If
            If InStr(strLine, "OBR") Then
                strQC = IIf(Split(strLine, "|")(1) = "1", "0", "1")
                strNo = Split(strLine, "|")(3)
            End If
            If InStr(strLine, "OBX") And UBound(Split(strLine, "|")) > 6 Then
                strItem = strItem & "|" & Split(strLine, "|")(3) & "|" & IIf(Trim(Split(strLine, "|")(5)) = "", "无", Trim(Split(strLine, "|")(5)))
                strItem = Replace(strItem, "umol/L", "")
                strItem = Replace(strItem, "mmol/L", "")
                strItem = Replace(strItem, "g/L", "")
                strItem = Replace(strItem, "Cell/uL", "")
                strItem = Trim(strItem)
            End If
            If UBound(Split(strLine, "|")) <= 5 Then        'base64画图
                If InStr(strLine, "OBX") > 0 Then
                    strPicData = Split(Split(strLine, "|")(5), "^")(2)
                Else
                    strPicData = strPicData & strLine
                End If
            End If
        Next
        If strPicData <> "" Then '获取最后一段
            strPicData = strPicData & Split(arrValue(UBound(arrValue) - 1), "|")(0)
            strPicData = Replace(strPicData, vbCr, "")
            strPicData = Replace(strPicData, vbLf, "")
            strPicPath = SaveBase64Bmp("1", strNo, strPicData)
            strPicData = ""
            strPic = "Image;103;1;" & strPicPath
        End If
        If strNo <> "" And strItem <> "" Then
            strResult = strResult & "||" & strDate & "|" & strNo & "^" & strStat & "^| |尿液|" & strQC & strItem
            strItem = ""
        End If
        If strPic <> "" Then
            strItem = ""
            If strPic <> "" Then
                strResult = strResult & vbCrLf & strPic
            End If
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Call WriteLog("接口返回", strResult, "")
    Exit Sub
hErr:
    Call WriteLog("Ku2000", Err.Description, "")
    Err.Clear
End Sub

Private Function SaveBase64Bmp(ByVal strType As String, ByVal str标本号 As String, ByVal str2Decode As String) As String

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
    Dim strDecode       As String
    Dim str             As String
    Dim ByteData()      As Byte
    Dim OutPut()        As Byte
    Dim lFrom           As Long
    Dim lTo             As Long
    Dim i               As Long
    Dim lngCount As Long, strFileName As String, lngFileNum
    
    '除去回车
    str = Replace(str2Decode, vbCrLf, "")

    '每4个字符一组（4个字符表示3个字）
    For lPtr = 1 To Len(str) Step 4
        iLen = 4
        For iCtr = 0 To 3
            '查找字符在BASE64字符串中的位置
            iValue = InStr(1, BASE64CHR, Mid$(str, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                '没有发现
                Case 0: Exit Function
            End Select
        Next

        '转换4个6比特数成为3个8比特数
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)
        
        '计算数组的起始位置
        lFrom = lTo
        lTo = lTo + (iLen - 1) - 1

        '重新定义输出数组
        ReDim Preserve OutPut(0 To lTo)

        For lngCount = lFrom To lTo
            OutPut(lngCount) = bits(lngCount - lFrom + 1)
        Next
        lTo = lTo + 1
    Next
   
    ReDim ByteData(UBound(OutPut))
    For i = 0 To UBound(OutPut)
        ByteData(i) = OutPut(i)
    Next
'
'        For i = 1078 To UBound(OutPut) - 2 Step 2
'            If OutPut(18) = &H18 And OutPut(19) = &H1 Then '直方图
''               If Output(i) <> byt1 Or byte2 <> Output(i + 1) Then
''                   byt1 = Output(i): byte2 = Output(i + 1)
'''                   Debug.Print byt1, byte2
''               End If
'                If OutPut(i) = OutPut(i - 2) And OutPut(i + 1) = OutPut(i - 1) And OutPut(i) = OutPut(i + 2) And OutPut(i + 1) = OutPut(i + 3) And ((OutPut(i) = 0 And OutPut(i + 1) = 124) Or (OutPut(i) = 255 And OutPut(i + 1) = 127) Or (OutPut(i) = 224 And OutPut(i + 1) = 3)) Then
'                    If ((i - 54) \ 560 > 0 And (i - 54) \ 560 <> 159) Then
'                        'If Output(i) = Output(i - 560) And Output(i + 1) = Output(i - 559) Then 'And Output(i) = Output(i + 560) And Output(i + 1) = Output(i + 561)
'                        If OutPut(i) = OutPut(i + 560) And OutPut(i + 1) = OutPut(i + 561) Then
'                            ByteData(i) = 0: ByteData(i + 1) = 0
'    '                        Debug.Print (i - 54) \ 560
'                        End If
'                    End If
'                End If
'            End If
'            If ByteData(i) = 0 Then
'               ByteData(i) = 255
'            End If
'            If ByteData(i + 1) = 0 Then
'               ByteData(i + 1) = 255
'            End If
'            'output2(i) = 255 - output2(i): output2(i + 1) = 255 - output2(i + 1)
'        Next i
 
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    
    lngFileNum = FreeFile
    Open strFileName For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
'    Call intsetbackgroundcolor("C:\Users\Administrator\Desktop\qwe.jpg", 1)
    SaveBase64Bmp = strFileName

End Function




