Attribute VB_Name = "mdlPublic"
Option Explicit
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63)     As String

Public Enum DataEnum
    responseText = 1
    responseBody = 2
End Enum

'本地日志模块
Private mobjFso As New FileSystemObject '文件对象
'公共调用方法
'*****************************************************************************************************************************
'将14位时间字符串转换为日期：YYYY-MM-DD HH:mm:ss
Public Function String14ToDate(ByVal strData As String, Optional ByRef strErr As String = "0") As String
    '获取时间戳
    Dim strTimeStamp As String
    If strData = "" Then
        If strErr = "0" Then
            MsgBoxEx "有效时间不能为空！", vbExclamation, gstrSysName
        Else
            strErr = "有效时间不能为空！"
        End If
        String14ToDate = ""
        Exit Function
    End If
    If Len(strData) = 14 Then
            Dim year As String, mouth As String, day As String, hour As String, mm As String, ss As String
            year = Mid(strData, 1, 4)
            mouth = Mid(strData, 5, 2)
            day = Mid(strData, 7, 2)
            hour = Mid(strData, 9, 2)
            mm = Mid(strData, 11, 2)
            ss = Mid(strData, 13, 2)
            strTimeStamp = year & "-" & mouth & "-" & day & " " & hour & ":" & mm & ":" & ss
            If Not IsDate(strTimeStamp) Then
                If strErr = "0" Then
                    MsgBoxEx "获取的时间不是一个日期！" & strTimeStamp, vbExclamation, gstrSysName
                Else
                    strErr = "获取的时间不是一个日期！" & strTimeStamp
                End If
                String14ToDate = ""
                Exit Function
            End If
    End If
    String14ToDate = strTimeStamp
End Function

'==========================================================
'| 模 块 名 | [BASE64]
'| 说    明 | BASE64编码及解码常用接口
'---------------------------------------------------------------------------《《Begin》》---------------------------------------------------------------------------------------
'==========================================================
Private Sub InitBase()
'功能:初始化 BASE64数组
     Dim iPtr     As Integer
     For iPtr = 0 To 63
         psBase64Chr(iPtr) = Mid$(BASE64CHR, iPtr + 1, 1)
     Next
End Sub

Public Function SaveBase64ToFile(ByVal strType As String, ByVal strSN As String, ByVal str2Decode As String) As String
'功能:保存Base64为图片文件
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
    
    Dim ByteData() As Byte, lngCount As Long, strFileName As String, lngFileNum
    
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
    
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strSN & "." & strType
    lngFileNum = FreeFile
    Open strFileName For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBase64ToFile = strFileName

End Function

Public Function EncodeBase64String(str2Encode As String) As String
'功能:对字符串进行Base64编码并返回字符串
     Dim sValue()             As Byte
     sValue = StrConv(str2Encode, vbFromUnicode)
     EncodeBase64String = EncodeBase64Byte(sValue)
End Function

Public Function EncodeBase64Byte(sValue() As Byte) As String
'功能:将一个字节数组进行Base64编码，并返回字符串
     Dim lCtr                 As Long
     Dim lPtr                 As Long
     Dim lLen                 As Long
     Dim sEncoded             As String
     Dim Bits8(1 To 3)        As Byte
     Dim Bits6(1 To 4)        As Byte
     Dim i As Integer
     InitBase
     For lCtr = 1 To UBound(sValue) + 1 Step 3
         For i = 1 To 3
             If lCtr + i - 2 <= UBound(sValue) Then
                 Bits8(i) = sValue(lCtr + i - 2)
                 lLen = 3
             Else
                 Bits8(i) = 0
                 lLen = lLen - 1
             End If
         Next

         '//转换字符串为数组，然后转换为4个6位(0-63)
         Bits6(1) = (Bits8(1) And &HFC) \ 4
         Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
         Bits6(3) = (Bits8(2) And &HF) * 4 + (Bits8(3) And &HC0) \ &H40
         Bits6(4) = Bits8(3) And &H3F

         '//添加4个新字符
         For lPtr = 1 To lLen + 1
             sEncoded = sEncoded & psBase64Chr(Bits6(lPtr))
         Next
     Next

     '//不足4位，以=填充
     Select Case lLen + 1
         Case 2: sEncoded = sEncoded & "=="
         Case 3: sEncoded = sEncoded & "="
         Case 4:
     End Select

     EncodeBase64Byte = sEncoded
End Function

Public Function EncodFileToBase64String(strFileSource As String)
'功能：对文件进行Base64编码并返回编码后的Base64字符串
     Dim lpdata() As Byte, _
         i As Long, _
         n As Long, _
         fso As New Scripting.FileSystemObject

     If Not fso.FileExists(strFileSource) Then Exit Function

     i = FreeFile

     Open strFileSource For Binary Access Read Lock Write As i

     n = LOF(i) - 1

     ReDim lpdata(0 To n)
     Get i, , lpdata
     Close i

     EncodFileToBase64String = EncodeBase64Byte(lpdata)
End Function

Public Sub EncodFileToBase64File(strFileSource As String, strFileBase64Desti As String)
'功能：对文件进行Base64编码，并将编码后的内容直接写入一个文本文件中
     Dim fso As New FileSystemObject, _
         ts As TextStream
    
     Set ts = fso.CreateTextFile(strFileBase64Desti, True)
     ts.Write (EncodFileToBase64String(strFileSource))
     ts.Close
     Set ts = Nothing
     Set fso = Nothing
End Sub

Public Function DecodeBase64Byte(str2Decode As String) As Byte()
'功能：从一个经过Base64的字符串中解码到源字节数组
     Dim lPtr             As Long
     Dim iValue           As Integer
     Dim iLen             As Integer
     Dim iCtr             As Integer
     Dim bits(1 To 4)     As Byte
     Dim strDecode        As String
     Dim str              As String
     Dim Output()         As Byte
    
     Dim iIndex           As Long

     Dim lFrom As Long
     Dim lTo As Long
    
     InitBase
    
     '//除去回车
     str = Replace(str2Decode, vbCrLf, "")

     '//每4个字符一组（4个字符表示3个字）
     For lPtr = 1 To Len(str) Step 4
         iLen = 4
         For iCtr = 0 To 3
             '//查找字符在BASE64字符串中的位置
             iValue = InStr(1, BASE64CHR, Mid$(str, lPtr + iCtr, 1), vbBinaryCompare)
             Select Case iValue   'A~Za~z0~9+/
                 Case 1 To 64:
                     bits(iCtr + 1) = iValue - 1
                 Case 65          '=
                     iLen = iCtr

 Exit For
                     '//没有发现
                 Case 0: Exit Function
             End Select
         Next

         '//转换4个6比特数成为3个8比特数
         bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
         bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
         bits(3) = (bits(3) And &H3) * &H40 + bits(4)

         '//计算数组的起始位置
         lFrom = lTo
         lTo = lTo + (iLen - 1) - 1
                
         '//重新定义输出数组
         ReDim Preserve Output(0 To lTo)
        
         For iIndex = lFrom To lTo
             Output(iIndex) = bits(iIndex - lFrom + 1)
         Next

         lTo = lTo + 1
        
     Next
     DecodeBase64Byte = Output
End Function

Public Function DecodeBase64String(str2Decode As String) As String
'功能：从一个经过Base64的字符串中解码到源字符串
     DecodeBase64String = StrConv(DecodeBase64Byte(str2Decode), vbUnicode)
End Function

Public Sub DecodeBase64StringToFile(strBase64 As String, strFilePath As String)
'功能:将一个Base64字符串解码，并写入二进制文件
     Dim fso As New Scripting.FileSystemObject
     Dim i As Long

     If fso.FileExists(strFilePath) Then
         fso.DeleteFile strFilePath, True
     End If

     i = FreeFile
     Open strFilePath For Binary Access Write As i
     Put i, , DecodeBase64Byte(strBase64)
     Close i
     Set fso = Nothing
End Sub

Public Sub DecodeBase64FileToFile(strBase64FilePath As String, strFilePath As String)
'功能:将一个Base64编码文件解码，并写入二进制文件
     Dim fso As New Scripting.FileSystemObject
     Dim ts As TextStream

     If Not fso.FileExists(strBase64FilePath) Then Exit Sub

     Set ts = fso.OpenTextFile(strBase64FilePath)
     
     DecodeBase64StringToFile ts.ReadAll, strFilePath
End Sub
'==========================================================
'| 模 块 名 | [BASE64]
'| 说    明 | BASE64编码及解码常用接口
'---------------------------------------------------------------------------《《End》》---------------------------------------------------------------------------------------
'==========================================================

'==========================================================
'| 模 块 名 | XMLHTTP
'| 说    明 | 替代Inet控件，实现数据通讯
'---------------------------------------------------------------------------《《Begin》》---------------------------------------------------------------------------------------
'==========================================================
Public Function HttpGet(ByVal Url As String, ByVal DataStic As DataEnum) As Variant
    Dim xmlHttp As Object
    Dim DataS As String
    Dim DataB() As Byte

    On Error GoTo errH:

100 Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
102 xmlHttp.Open "get", Url, True
104 xmlHttp.send

106 Do While xmlHttp.readyState <> 4
108     DoEvents
    Loop

    '--------------------------------------函数返回
110 Select Case DataStic
    Case responseText
        '--------------------------------直接返回字符串
112     DataS = xmlHttp.responseText
114     HttpGet = DataS
116 Case responseBody
        '--------------------------------直接返回二进制
118     DataB = xmlHttp.responseBody
120     HttpGet = DataB
122 Case responseBody + responseText
        '------------------------------二进制转字符串[直接返回字串出现乱码时尝试]
124     DataS = BytesToStr(xmlHttp.responseBody)
126     HttpGet = DataS
128 Case Else
        '--------------------------------无效的返回
130     HttpGet = ""
    End Select

    '--------------------------------------释放空间
132 Set xmlHttp = Nothing

    Exit Function

errH:
134 HttpGet = ""
136 MsgBoxEx "HttpGet失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "中联软件"
End Function

Public Function HttpPost(ByVal strUrl As String, ByVal strData As String, ByVal DataStic As DataEnum, Optional ByVal strCONTENTTYPE As String) As Variant
'    Dim XMLHTTP As Object
    Dim xmlHttp As MSXML2.ServerXMLHTTP
    Dim DataS As String
    Dim DataB() As Byte

    On Error GoTo errH:

'100 Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
    Set xmlHttp = New MSXML2.ServerXMLHTTP
102 xmlHttp.Open "POST", strUrl, True
104 xmlHttp.setRequestHeader "Content-Length", Len(HttpPost)
    If strCONTENTTYPE = "" Then
106     xmlHttp.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
    Else
        xmlHttp.setRequestHeader "CONTENT-TYPE", strCONTENTTYPE  '"application/x-www-form-urlencoded; charset=utf-8"
    End If
108 xmlHttp.send (strData)

110 Do Until xmlHttp.readyState = 4
112     DoEvents
    Loop

    '-----------------------------函数返回
114 Select Case DataStic
    Case responseText
        '--------------------------------直接返回字符串
116     DataS = xmlHttp.responseText
118     HttpPost = DataS
120 Case responseBody
        '--------------------------------直接返回二进制
122     DataB = xmlHttp.responseBody
124     HttpPost = DataS
126 Case responseBody + responseText
        '---------------------------二进制转字符串[直接返回字串出现乱码时尝试]
128     DataS = BytesToStr(xmlHttp.responseBody)
130     HttpPost = DataS
132 Case Else
        '--------------------------------无效的返回
134     HttpPost = ""
    End Select

    '------------------------------------释放空间
136     Set xmlHttp = Nothing

    Exit Function

errH:
138     HttpPost = ""
140     MsgBoxEx "HttpPost失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "中联软件"
End Function

Private Function BytesToStr(ByVal vInput As Variant) As String
    
    Dim strReturn       As String
    Dim i               As Long
    Dim intPrevCharCode As Integer
    Dim intNextCharCode As Integer

    For i = 1 To LenB(vInput)
        intPrevCharCode = AscB(MidB(vInput, i, 1))
        If intPrevCharCode < &H80 Then
            strReturn = strReturn & Chr(intPrevCharCode)
        Else
            intNextCharCode = AscB(MidB(vInput, i + 1, 1))
            strReturn = strReturn & Chr(CLng(intPrevCharCode) * &H100 + CInt(intNextCharCode))
            i = i + 1
        End If

    Next

    BytesToStr = strReturn
End Function

'==========================================================
'| 模 块 名 | XMLHTTP
'| 说    明 | 替代Inet控件，实现数据通讯
'-----------------------------------------------------------------------------《《END》》-------------------------------------------------------------------------
'==========================================================

'山东省时间戳格式返回格式化处理
'例 “Dec 30 01:12:53 2014 GMT” 处理为“2014-12-30 01:12:53”
Public Function GetTimes(ByVal GmtTime As String) As String
    Dim t1 As String
    Dim strYear As String
    Dim strTime As String

    If Len(Trim(GmtTime)) = 0 Then Exit Function
    If InStr(1, GmtTime, " GMT", vbTextCompare) = 0 Then Exit Function

    t1 = Trim(Replace(GmtTime, "GMT", "", 1, , vbTextCompare))
    strYear = Mid(t1, Len(t1) - 3, 4)
    strTime = Mid(t1, Len(t1) - 12, 8)
    t1 = Mid(t1, 1, Len(t1) - 13)
    t1 = t1 & " " & strYear
    GetTimes = Format$(t1, "yyyy-mm-dd ") & strTime
End Function

'检查证书有效性,返回证书有效期天数
Public Function CheckValidaty(ByVal endDate As Date) As Integer
    '北京CA江苏版检查证书有效性接口
    '-入参: 证书有效截止日期
    '-出参：有效天数
    Dim dblAllSp    As Double
    Dim result      As Integer
    Dim datNow As Date
    datNow = gobjComLib.zlDatabase.Currentdate
    dblAllSp = CDbl(CDate(endDate)) - CDbl(datNow)
    result = Int(dblAllSp)
    CheckValidaty = result
End Function

Public Sub WriteLog(ByVal strLogTxt As String)
    '写一行日志，如果内容中有回车,换行符，替换为<CR><LF>
    '日志保存在当前目录下的[应用程序名称]Log目录下，文件名为日期.txt,默认保存7天的日志。
    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '日志路径，文件名，配置文件名
    Dim strLogSaveDays As String '日志保留天数
    Dim dblFreeSpace As Double   '剩余空间
    Dim strDelOldFile As String  '过期文件
    Dim objFile As File
    
    If Dir(App.Path & "\电子签名调试日志*.log") = "" Then Exit Sub
    '始终保存日志
    '2、清除过期日志
    strLogSaveDays = "7"  '保留7天的日志
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\电子签名调试日志*.log")
    Do While strDelOldFile <> ""
        Set objFile = mobjFso.GetFile(strLogPath & "\" & strDelOldFile)
        If DateDiff("d", objFile.DateLastModified, Now) > Val(strLogSaveDays) Then
            mobjFso.DeleteFile strLogPath & "\" & strDelOldFile, True
        End If
        strDelOldFile = Dir
    Loop
    
    '3、空间是否足够
    dblFreeSpace = GetFreeSpace(strLogPath)
    If dblFreeSpace >= 1024 And dblFreeSpace <= 10240 Then
        '空间不足，不写日志,产生一个警告文件
        If Not mobjFso.FileExists(strLogPath & "\空间不足.txt") Then Call mobjFso.CreateTextFile(strLogPath & "\空间不足.txt", True)
        Exit Sub
    Else
        '清除警告文件
        If mobjFso.FileExists(strLogPath & "\空间不足.txt") Then Call mobjFso.DeleteFile(strLogPath & "\空间不足.txt", True)
    End If
    '4、写入日志行
    strLogFile = strLogPath & "\电子签名调试日志" & Format(Now, "yyyyMMdd") & ".log"
    Call SaveLog(strLogFile, strLogTxt)

End Sub

Private Sub SaveLog(ByVal strFileName As String, ByVal strInput As String, Optional ByVal strDate As String)
 
    Dim objStream As TextStream
    Dim strWritLing As String
    
    strWritLing = Replace$(strInput, Chr(&HD), "<CR>")
    strWritLing = Replace$(strInput, Chr(&HA), "<LF>")

    If strInput <> "" Then
        If Not mobjFso.FileExists(strFileName) Then Call mobjFso.CreateTextFile(strFileName)
        Set objStream = mobjFso.OpenTextFile(strFileName, ForAppending)
        If strDate = "" Then strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine (strDate & Chr(&H9) & strInput)
        objStream.Close
        Set objStream = Nothing
    End If
    
End Sub

Private Function GetFreeSpace(ByVal strPath As String) As Double
    '获取剩余空间
    Dim strDriv As String, Drv As Drive
    Dim strDir As String
    
    If mobjFso.FolderExists(strPath) Then
        strDriv = mobjFso.GetDriveName(mobjFso.GetAbsolutePathName(strPath))
        Set Drv = mobjFso.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function

Public Function LogWrite(ByVal strFunction As String, ByVal strLog As String)
    If Not gobjComLib Is Nothing Then
        
        On Error Resume Next
        gobjComLib.LogWrite "电子签名调试日志", "", strFunction, strLog
        If Err.Number = 438 Then        '兼容低版本没有日志接口
            WriteLog strFunction & vbCrLf & strLog
        End If
        Err.Clear: On Error GoTo 0
    Else
        Exit Function
    End If
End Function
