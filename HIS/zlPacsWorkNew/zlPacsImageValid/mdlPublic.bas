Attribute VB_Name = "mdlPublic"
Option Explicit

Public gobjFile As New FileSystemObject
Public gstrSysName As String
Public Const CON_STR_HINT_TITLE As String = "提示"


Public Enum emResult
    etUndetected = 0    '未校对或未校对出错
    etFileMiss = 1      '文件缺失
    etFileNull = 2      '文件大小为0
    etReadError = 3     '读取异常
    etRoadError = 4     '路径错误
    etSucceed = 5       '校对成功
End Enum

Public Enum TMediaType
    imgTag = 0   '图像标记
    MULFRAMETAG = 1 '多侦图
    VIDEOTAG = 2 '视频标记
    AUDIOTAG = 3 '音频标记
End Enum
 


Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'获取加密密码
Public Function getEncryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(20)
    strRandom = GetRandom(20)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    getEncryptionPassW = strBase & Join(strTemp, "") & strRandom '加密后的字串
End Function

'获取解密密码
Public Function getDecryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Integer
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim lngBase As Long
    Dim strRandom As String
    Dim strPassSouce As String

    i = 0
    
    strPassSouce = Mid(strPassW, 2, Len(strPassW) - 2)
    lngPassWLength = Len(strPassSouce)
    lngBase = Asc(Mid(strPassW, 1, 1))
    
    strRandom = Right(strPassW, 1)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
    
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassSouce, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strRandom) Xor lngBase
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop

    getDecryptionPassW = Join(strTemp, "") '解密后的字串
End Function

Private Function TranPasswd(strOld As String) As String
    Dim iBit As Integer, StrBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        StrBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(StrBit = "0", "W", StrBit = "1", "I", StrBit = "2", "N", StrBit = "3", "T", StrBit = "4", "E", StrBit = "5", "R", StrBit = "6", "P", StrBit = "7", "L", StrBit = "8", "U", StrBit = "9", "M", _
                   StrBit = "A", "H", StrBit = "B", "T", StrBit = "C", "I", StrBit = "D", "O", StrBit = "E", "K", StrBit = "F", "V", StrBit = "G", "A", StrBit = "H", "N", StrBit = "I", "F", StrBit = "J", "J", _
                   StrBit = "K", "B", StrBit = "L", "U", StrBit = "M", "Y", StrBit = "N", "G", StrBit = "O", "P", StrBit = "P", "W", StrBit = "Q", "R", StrBit = "R", "M", StrBit = "S", "E", StrBit = "T", "S", _
                   StrBit = "U", "T", StrBit = "V", "Q", StrBit = "W", "L", StrBit = "X", "Z", StrBit = "Y", "C", StrBit = "Z", "X", True, StrBit)
        Case 2
            strNew = strNew & _
                Switch(StrBit = "0", "7", StrBit = "1", "M", StrBit = "2", "3", StrBit = "3", "A", StrBit = "4", "N", StrBit = "5", "F", StrBit = "6", "O", StrBit = "7", "4", StrBit = "8", "K", StrBit = "9", "Y", _
                   StrBit = "A", "6", StrBit = "B", "J", StrBit = "C", "H", StrBit = "D", "9", StrBit = "E", "G", StrBit = "F", "E", StrBit = "G", "Q", StrBit = "H", "1", StrBit = "I", "T", StrBit = "J", "C", _
                   StrBit = "K", "U", StrBit = "L", "P", StrBit = "M", "B", StrBit = "N", "Z", StrBit = "O", "0", StrBit = "P", "V", StrBit = "Q", "I", StrBit = "R", "W", StrBit = "S", "X", StrBit = "T", "L", _
                   StrBit = "U", "5", StrBit = "V", "R", StrBit = "W", "D", StrBit = "X", "2", StrBit = "Y", "S", StrBit = "Z", "8", True, StrBit)
        Case 0
            strNew = strNew & _
                Switch(StrBit = "0", "6", StrBit = "1", "J", StrBit = "2", "H", StrBit = "3", "9", StrBit = "4", "G", StrBit = "5", "E", StrBit = "6", "Q", StrBit = "7", "1", StrBit = "8", "X", StrBit = "9", "L", _
                   StrBit = "A", "S", StrBit = "B", "8", StrBit = "C", "5", StrBit = "D", "R", StrBit = "E", "7", StrBit = "F", "M", StrBit = "G", "3", StrBit = "H", "A", StrBit = "I", "N", StrBit = "J", "F", _
                   StrBit = "K", "O", StrBit = "L", "4", StrBit = "M", "K", StrBit = "N", "Y", StrBit = "O", "D", StrBit = "P", "2", StrBit = "Q", "T", StrBit = "R", "C", StrBit = "S", "U", StrBit = "T", "P", _
                   StrBit = "U", "B", StrBit = "V", "Z", StrBit = "W", "0", StrBit = "X", "V", StrBit = "Y", "I", StrBit = "Z", "W", True, StrBit)
        End Select
    Next
    TranPasswd = strNew

End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As ADODB.Connection
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSql As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    
    On Error Resume Next
    
    strUserPwd = TranPasswd(strUserPwd)
    
    Err = 0
    DoEvents
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName & ";Persist Security Info=false;", strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If

            Set OraDataOpen = Nothing
            Exit Function
        End If
    End With

    Err = 0
    On Error GoTo errHand

    'gstrDbUser = UCase(strUserName)
    'gobjComLib.SetDbUser gstrDbUser
    
    Set OraDataOpen = cnOracle
    Exit Function

errHand:
    MsgBox strError, vbInformation, gstrSysName
    Set OraDataOpen = Nothing
    Err = 0
End Function



Public Function GetAppPath() As String
    If App.LogMode = 0 Then
        GetAppPath = "C:\Appsoft\Apply"
    Else
        GetAppPath = Replace(App.Path & "\Apply\", "\\", "")
    End If
End Function

Public Function GetAppRoot() As String
    If App.LogMode = 0 Then
        GetAppRoot = "C:\Appsoft"
    Else
        GetAppRoot = Replace(App.Path, "\\", "")
    End If
End Function

Public Function GetResourceDir() As String
'获取资源目录
    GetResourceDir = GetAppPath & "\..\附加文件\"
End Function

Public Function GetCacheDir() As String
'获取缓存目录
    GetCacheDir = GetAppPath & "\TmpImage\"
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next

    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir

    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出,根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中
'参数：vDate=时间点或时间段的开始时间

    MovedByDate = gobjComlib.zlDatabase.DateMoved(CStr(vDate), 1, 100)
    
End Function


Public Function Unicode8Decode(bTemp() As Byte) As String
'解码UNICODE UTF-8
    Dim i As Long
    Dim k As Long
    Dim strReturn As String
    Dim strTmp() As Byte
    Dim Code As Long
    Dim Code1 As Long
    Dim Code2 As Long
    Dim Code3 As Long
    Dim Code4 As Long
    Dim bNo As Long
    
    k = UBound(bTemp)
    ReDim strTmp(k * 2)
    bNo = 0
    
    For i = 0 To k
        If (bTemp(i) And 128) = 0 Then
            strTmp(bNo) = bTemp(i)
            bNo = bNo + 1
        ElseIf (bTemp(i) And 252) = 252 Then
            '11111100
            Code1 = (bTemp(i) And 1) * 64 + bTemp(i + 1) And 63
            Code2 = (bTemp(i + 2) And 63) * 4 + (bTemp(i + 3) And 48) \ 16
            Code3 = (bTemp(i + 3) And 15) * 16 + (bTemp(i + 4) And 60) \ 4
            Code4 = (bTemp(i + 4) And 3) * 64 + (bTemp(i + 5) And 63)
            Code = ((Code1 * 256 + Code2) * 256 + Code3) * 256 + Code4
            Code = CLng("&H" + hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 5
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            strTmp(bNo + 1) = Code \ 65536
            strTmp(bNo + 1) = Code \ 16777216
            bNo = bNo + 4
        ElseIf (bTemp(i) And 248) = 248 Then '11111000
            Code1 = (bTemp(i) And 3)
            Code2 = (bTemp(i + 1) And 63) * 4 + (bTemp(i + 2) And 48) \ 16
            Code3 = (bTemp(i + 2) And 15) * 16 + (bTemp(i + 3) And 60) \ 4
            Code4 = (bTemp(i + 3) And 3) * 64 + (bTemp(i + 4) And 63)
            Code = ((Code1 * 256 + Code2) * 256 + Code3) * 256 + Code4
            Code = CLng("&H" + hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 4
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            strTmp(bNo + 1) = Code \ 65536
            strTmp(bNo + 1) = Code \ 16777216
            bNo = bNo + 4
        ElseIf (bTemp(i) And 240) = 240 Then '11110000
            Code1 = (bTemp(i) And 7) * 8 + (bTemp(i + 1) And 48) \ 16
            Code2 = (bTemp(i + 1) And 15) * 16 + (bTemp(i + 2) And 60) \ 4
            Code3 = (bTemp(i + 2) And 3) * 64 + (bTemp(i + 3) And 63)
            Code = (Code1 * 256 + Code2) * 256 + Code3
            Code = CLng("&H" + hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 3
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            strTmp(bNo + 1) = Code \ 65536
            strTmp(bNo + 1) = Code \ 16777216
            bNo = bNo + 4
        ElseIf (bTemp(i) And 224) = 224 Then '11100000
            Code1 = (bTemp(i) And 15) * 16 + (bTemp(i + 1) And 60) \ 4
            Code2 = (bTemp(i + 1) And 3) * 64 + (bTemp(i + 2) And 63)
            Code = Code1 * 256 + Code2
            Code = CLng("&H" + hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 2
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            bNo = bNo + 2
        ElseIf (bTemp(i) And 192) = 192 Then '11000000
            Code1 = (bTemp(i) And 28) \ 4
            Code2 = (bTemp(i) And 3) * 64 + (bTemp(i + 1) And 63)
            Code = Code1 * 256 + Code2
            Code = CLng("&H" + hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 1
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            bNo = bNo + 2
        End If
    Next
        
    ReDim Preserve strTmp(bNo - 1)
    strReturn = StrConv(strTmp, vbUnicode)
    Unicode8Decode = strReturn
End Function


Public Function Unicode8Encode(bTemp As String) As Byte()
'编码UNICODE UTF-8
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strTotal() As Byte
    Dim strTmp As String
    Dim Code As Long
    Dim Code1 As Long
    Dim Code2 As Long
    Dim Code3 As Long
    Dim Code4 As Long
    Dim Code5 As Long
    Dim Code6 As Long  '已生成的字节数
    Dim bNo As Long
    
    k = Len(bTemp)
    bNo = 0
    
    ReDim strTotal(k * 3)
    For i = 1 To k
        Code = CLng("&H" + hex(AscW(Mid(bTemp, i, 1))))
        If Code < 128& Then
            strTotal(bNo) = Code
            bNo = bNo + 1
            If bNo > 422386 Then
                Debug.Print Code
            End If
        ElseIf Code < 2048& Then
            Code1 = ((Code And 1984&) \ 32&) + 192
            Code2 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            bNo = bNo + 2
        ElseIf Code < 65536 Then
            Code1 = ((Code And 61440) \ 4096&) + 224
            Code2 = ((Code And 4032&) \ 64&) + 128
            Code3 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            bNo = bNo + 3
        ElseIf Code < 2097152 Then
            Code1 = ((Code And 1835008) \ 262144) + 240
            Code2 = ((Code And 258048) \ 4096&) + 128
            Code3 = ((Code And 4032&) \ 64&) + 128
            Code4 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            bNo = bNo + 4
        ElseIf Code < 67108864 Then
            Code1 = ((Code And 50331648) \ 16777216) + 248
            Code2 = ((Code And 16515072) \ 262144) + 128
            Code3 = ((Code And 258048) \ 4096&) + 128
            Code4 = ((Code And 4032&) \ 64&) + 128
            Code5 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            strTotal(bNo + 4) = Code5
            bNo = bNo + 5
        Else
            Code1 = IIf(Code And 1073741824 = 1073741824, 253&, 252&)
            Code2 = ((Code And 1056964608) \ 16777216) + 128
            Code3 = ((Code And 16515072) \ 262144) + 128
            Code4 = ((Code And 258048) \ 4096&) + 128
            Code5 = ((Code And 4032&) \ 64&) + 128
            Code6 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            strTotal(bNo + 4) = Code5
            strTotal(bNo + 5) = Code6
            bNo = bNo + 6
        End If
    Next
    
    ReDim Preserve strTotal(bNo - 1)
    Unicode8Encode = strTotal
End Function


Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If Err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, "提示"
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ReadViewImage(ByVal strFile As String, Optional ByRef dcmViewer As DicomViewer = Nothing) As DicomImage
On Error GoTo errHandle
    Dim dImgs As DicomImages
        
    '如果包含_copy_vdat_，说明是临时文件
    If InStr(strFile, "_copy_vdat_") > 0 Then
        Set ReadViewImage = Nothing
        Call Kill(strFile)
        
        Exit Function
    End If
    
    If dcmViewer Is Nothing Then
        Set dImgs = New DicomImages
    Else
        Set dImgs = dcmViewer.Images
    End If
    
    Set ReadViewImage = ReadDicomFile(strFile, dImgs)
    
Exit Function
errHandle:
    Set ReadViewImage = Nothing
End Function

Public Function ReadDicomFile(ByVal strFile As String, dcmImgs As DicomImages) As DicomImage
    Dim curImage As DicomImage
    Dim blnUseUrl As Boolean
    Dim strFileTime As String
    
    On Error Resume Next
    
    blnUseUrl = IIf(InStr(strFile, " ") <= 0, True, False)
    
    If blnUseUrl Then
        'readurl不支持空格
        Set curImage = dcmImgs.ReadURL(strFile)
    Else
        Set curImage = dcmImgs.ReadFile(strFile)
    End If
    
    If Err.Number = 0 Then
    
        'dcmImgs 不能正常加载
        If curImage.Picture Is Nothing Then
            Set ReadDicomFile = Nothing
            Exit Function
        End If
    
        Set ReadDicomFile = curImage
        Exit Function
    End If
    
    '2098错误一种是文件不是dicom文件，另一种是存在共享访问错误
    If InStr(Err.Description, "sharing violation") > 0 Then
        Err.Clear
        strFileTime = Format(Now, "YYMMDD") & GetTickCount
        
        Call FileCopy(strFile, strFile & "_copy_vdat_" & strFileTime)
    
        If blnUseUrl Then
            'readurl不支持空格
            Set curImage = dcmImgs.ReadURL(strFile & "_copy_vdat_" & strFileTime)
        Else
            Set curImage = dcmImgs.ReadFile(strFile & "_copy_vdat_" & strFileTime)
        End If
    
        If Err.Number = 0 Then
            Call Kill(strFile & "_copy_vdat_" & strFileTime)
            Err.Clear
        Else
            Call Kill(strFile & "_copy_vdat_" & strFileTime)
        End If
    Else
        Err.Clear
        Set curImage = dcmImgs.AddNew
        Call curImage.FileImport(strFile, "JPG")
        
        If Err.Number <> 0 Then
            Err.Clear
            'not a JPG file
            Call curImage.FileImport(strFile, "BMP")
        End If
        
        If Err.Number <> 0 Then
            Err.Clear
            'not a BMP file
            Call curImage.FileImport(strFile, "AVI")
        End If
        
        If Err.Number <> 0 Then
            Err.Clear
            'not a AVI file
            Call curImage.FileImport(strFile, "MPG")
        End If
    End If
    
    If Err.Number = 0 Then
        Set ReadDicomFile = curImage
        Exit Function
    End If
    
    Set ReadDicomFile = Nothing
    
Err.Clear
End Function
