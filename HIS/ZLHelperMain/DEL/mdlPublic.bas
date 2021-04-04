Attribute VB_Name = "mdlPublic"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/1/21
'模块           mdlPublic
'说明
'==================================================================================================
Private mobjFSO        As New FileSystemObject
'字符串用UTF-8编码
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long

Public Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
End Function

'--------------------------------------------------------------------------------------------------
'方法           DisPlayOneValue
'功能           展示对象
'返回值         String
'入参列表:
'参数名         类型                    说明
'valValue       Variant                 传入的对象
'-------------------------------------------------------------------------------------------------
Public Function DisPlayOneValue(valValue As Variant) As String
    Dim strTmp  As String
    
    If IsArray(valValue) Then
        Dim i    As Long
        strTmp = "["
        For i = LBound(valValue) To UBound(valValue)
            strTmp = strTmp & DisPlayOneValue(valValue(i)) & ","
        Next
        If Len(strTmp) = 1 Then
            strTmp = strTmp & "]"
        Else
            strTmp = Mid(strTmp, 1, Len(strTmp) - 1) & "]"
        End If
    ElseIf IsNull(valValue) Then
        strTmp = "{NULL}"
    ElseIf IsEmpty(valValue) Then
        strTmp = "{EMPTY}"
    ElseIf IsObject(valValue) Then
        If valValue Is Nothing Then
            strTmp = "{NOTHING}"
        Else
            strTmp = "{OBJECT(" + TypeName(valValue) + ")=" & Serialize(valValue) & "}"
        End If
    Else
        If VarType(valValue) = vbString Then
            strTmp = """" & valValue & """"
        Else
            strTmp = CStr(valValue)
        End If
    End If
    DisPlayOneValue = strTmp
End Function
'--------------------------------------------------------------------------------------------------
'方法           StringToUTF8Bytes       将字符串转换为UTF-8编码的字节数组
'返回值         Byte()                  16进制字符串转换的字节组
'入参列表:
'参数名         类型                    说明
'strInput      String                  16进制字符串
'-------------------------------------------------------------------------------------------------
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '先计算需求字节数
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    '然后转换
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function

'--------------------------------------------------------------------------------------------------
'方法           UTF8BytesToString       将UTF-8编码的字节数组转换为字符串
'返回值         String                  转换后的字符串
'入参列表:
'参数名         类型                    说明
'bytInpu        Byte(）                 字节数组
'-------------------------------------------------------------------------------------------------
Public Function UTF8BytesToString(bytInpu() As Byte) As String
    Dim lngBytesRequired As Long

    '先计算需求字节数
    lngBytesRequired = MultiByteToWideChar(CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal 0, 0)
     
    '然后转换
    UTF8BytesToString = String(lngBytesRequired, 0)
    MultiByteToWideChar CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal StrPtr(UTF8BytesToString), lngBytesRequired
End Function

'-------------------------------------------------------------------------------------------------
'方法           EncBase64Char           将6-bit字节转换为Base64字符
'返回值         Byte                    字符数值
'入参列表:
'参数名         类型                    说明
'bytValue       Byte                    转换的字节
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Private Function EncBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue < 26 Then '26个大写英文字母
        EncBase64Char = bytValue + &H41
    ElseIf bytValue < 52 Then '26个小写英文字母
        EncBase64Char = bytValue + &H61 - 26
    ElseIf bytValue < 62 Then '10个数字
        EncBase64Char = bytValue + &H30 - 52
    ElseIf bytValue = 62 Then
        EncBase64Char = &H2B '+
    Else
        EncBase64Char = &H2F '/
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           DecBase64Char           将Base64字符转换为6 bit字节
'返回值         Byte                    字符数值
'入参列表:
'参数名         类型                    说明
'bytValue       Byte                    待解码的字节
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Private Function DecBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue >= &H41 And bytValue <= &H5A Then
        DecBase64Char = bytValue - &H41
    ElseIf bytValue >= &H61 And bytValue <= &H7A Then
        DecBase64Char = bytValue - &H61 + 26
    ElseIf bytValue >= &H30 And bytValue <= &H39 Then
        DecBase64Char = bytValue - &H30 + 52
    ElseIf bytValue = &H2B Then
        DecBase64Char = 62
    ElseIf bytValue = &H2F Then
        DecBase64Char = 63
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           EncodeBase64            进行Base64编码，返回Base64的字符串
'返回值         String                  Base64编码结果
'入参列表:
'参数名         类型                    说明
'varInput       Variant                 需要进行Base64编码的字符串或者字节数组，字符串采取UTF-8编码。Byte()类型前面的数组，元素个数传3的倍数，最后一次传递所有剩下的即可。
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Public Function EncodeBase64(varInput As Variant) As String
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, J              As Long, lngBit     As Long
    
    On Error GoTo ErrH
    
    If VarType(varInput) = vbString Then
        If Len(varInput) = 0 Then Exit Function
        '原始内容,先将原文以UTF-8的方式编码
        bytInput = StringToUTF8Bytes(CStr(varInput))
    ElseIf VarType(varInput) = vbArray + vbByte Then
        If UBound(varInput) < 0 Then Exit Function
        bytInput = varInput
    Else
        Exit Function
    End If
    lngInputLen = UBound(bytInput) + 1
 
    lngOutLen = lngInputLen + (lngInputLen - 1) \ 3 + 1
    ReDim bytOut(lngOutLen - 1)
    '将8-bit字节数组转换为6-bit字节数组
    For i = 0 To lngInputLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(J) = (bytInput(i) And &HFC) \ &H4
            J = J + 1
            bytOut(J) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两位
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &HF0) \ &H10)
            J = J + 1
            bytOut(J) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四位
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &HC0) / &H40)
            J = J + 1
            bytOut(J) = bytInput(i) And &H3F
            J = J + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) '转换为Base64字符
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") '原文剩余内容不足3个字节需要补齐
    Exit Function
ErrH:
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           DecodeBase64            将Base64的字符串解码为原文。
'返回值         Variant                 原始字符或者原始的字节组
'入参列表:
'参数名         类型                    说明
'strInput       String                  Base64编码字符串
'blnByteArray   Boolean                 True:返回Byte(),False-返回string
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Public Function DecodeBase64(strInput As String, Optional ByVal blnByteArray As Boolean) As Variant
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, J              As Long, lngBit     As Long
    Dim lngModLen       As Long
    On Error GoTo ErrH
    If Len(strInput) = 0 Then Exit Function
    lngModLen = InStr(strInput, "=")
    If lngModLen > 0 Then
        '编码后的内容
        lngModLen = Len(strInput) - lngModLen + 1
        bytInput = StrConv(strInput, vbFromUnicode)
    Else
        lngModLen = 0
        '编码后的内容
        bytInput = StrConv(strInput, vbFromUnicode)
    End If
    lngInputLen = UBound(bytInput) + 1
 
    '原始内容
    lngOutLen = lngInputLen - lngInputLen \ 4
    lngOutLen = lngOutLen - lngModLen
    ReDim bytOut(lngOutLen - 1)
 
    For J = 0 To lngInputLen - 1
        bytInput(J) = DecBase64Char(bytInput(J)) '从Base64字符转换为6-bit字节
    Next
    '将6-bit字节数组转换为8-bit字节数组
    For J = 0 To lngOutLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(J) = bytInput(i) * &H4
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &H30) \ &H10)
            lngBit = 2
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两字节
            bytOut(J) = (bytInput(i) And &HF) * &H10
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &H3C) \ &H4)
            lngBit = 4
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四字节
            bytOut(J) = (bytInput(i) And &H3) * &H40
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(J) = bytOut(J) Or bytInput(i)
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            lngBit = 0
        End If
    Next
    If blnByteArray Then
        DecodeBase64 = bytOut
    Else
        '最后将转换得到的UTF-8字符串转换为VB支持的Unicode字符串以便于显示。
        DecodeBase64 = UTF8BytesToString(bytOut)
    End If
    Exit Function
ErrH:
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'方法           EncodeBase64_file       对文件进行Base64编码，返回Base64的字符串
'返回值         String                  Base64编码结果
'入参列表:
'参数名         类型                    说明
'strFile        String                  需要进行Base64编码的文件
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Public Function EncodeBase64_File(ByVal strFile As String) As String
    Dim lngFileNum  As Long, lngFileSize    As Long, lngModSize As Long, lngBlocks As Long
    Dim lngCount    As Long, lngCurSize     As Long
    Dim strReturn   As String
    Dim aryChunk()    As Byte
    
    Const conChunkSize      As Long = 3000
    
    On Error GoTo ErrH
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    If lngFileSize <> 0 Then
        lngModSize = lngFileSize Mod conChunkSize
        lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
        For lngCount = 0 To lngBlocks
            If lngCount = lngFileSize \ conChunkSize Then
                lngCurSize = lngModSize
                ReDim aryChunk(lngCurSize - 1) As Byte
            Else
                lngCurSize = conChunkSize
                If lngCount = 0 Then '防止不停分配内存
                    ReDim aryChunk(lngCurSize - 1) As Byte
                End If
            End If
            Get lngFileNum, , aryChunk()
            strReturn = strReturn & EncodeBase64(aryChunk)
        Next
        Close lngFileNum
        EncodeBase64_File = strReturn
    End If
    Exit Function
ErrH:
    Close lngFileNum
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           DecodeBase64_File       将Base64的字符串解码为原文。
'返回值         String                  生成的文件名
'入参列表:
'参数名         类型                    说明
'strInput       String                  Base64编码字符串
'strFile        String                  指定文件名
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Public Function DecodeBase64_File(strInput As String, Optional ByVal strFile As String) As String
    Dim lngFileNum  As Long, lngFileSize    As Long
    Dim lngCount    As Long, lngCurSize     As Long
    Dim strTmp      As String
    Dim aryChunk()    As Byte
    Const conChunkSize      As Long = 4000
    
    On Error GoTo ErrH
    If strFile = "" Then
        strFile = mobjFSO.GetSpecialFolder(TemporaryFolder) & "\" & mobjFSO.GetTempName
    Else
        If mobjFSO.FileExists(strFile) Then Kill strFile
    End If
    lngFileNum = FreeFile
    Open strFile For Binary As lngFileNum
    lngCount = 0
    lngCurSize = 0
    lngFileSize = Len(strInput)
    If lngFileSize <> 0 Then
        For lngCount = 1 To lngFileSize Step conChunkSize
            strTmp = Mid(strInput, lngCount, conChunkSize)
            aryChunk = DecodeBase64(strTmp, True)
            Put lngFileNum, , aryChunk()
        Next
        Close lngFileNum
    End If
    DecodeBase64_File = strFile
    Exit Function
ErrH:
    Close lngFileNum
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'方法           Serialize               将对象或值序列化为字符串
'返回值         String                  序列化的字符串
'入参列表:
'参数名         类型                    说明
'objInfo        Variant                 对象或值
'strKeyName     String                  序列化的关键字
'-------------------------------------------------------------------------------------------------
Public Function Serialize(ByVal objInfo As Variant, Optional ByVal strKeyName As String = "K_Default") As String
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    
    On Error Resume Next

    objBag.WriteProperty strKeyName, objInfo
    If Err.Number = 330 Then
        '非法参数。  因为不支持持久性不能写对象。
        Serialize = "{NotPersistable}"
        Err.Clear
    Else
        bytData = objBag.Contents
        Serialize = EncodeBase64(bytData())
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           UnSerialize             将字符串反序列化为对象或具体的值
'返回值         Variant                 序列化字符串对应的对象或具体的值
'入参列表:
'参数名         类型                    说明
'strSource      String                  序列化字符串
'strKeyName     String                  序列化的关键字
'-------------------------------------------------------------------------------------------------
Public Function UnSerialize(ByVal strSource As String, Optional ByVal strKeyName As String = "K_Default") As Variant
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    
    On Error Resume Next
    If Len(strSource) = 0 Then Exit Function
    If strSource = "{NotPersistable}" Then
         Set UnSerialize = Nothing
    Else
        bytData = DecodeBase64(strSource, True)
        objBag.Contents = bytData
        If Not IsObject(objBag.ReadProperty(strKeyName)) Then
            UnSerialize = objBag.ReadProperty(strKeyName)
        Else
            Set UnSerialize = objBag.ReadProperty(strKeyName)
        End If
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           SerializeMulti          按顺序序列化多个信息
'返回值         String                  序列化的字符串
'入参列表:
'参数名         类型                    说明
'arrInfo        Variant                 多个序列化的对象
'[      ]       long                    按0开始索引，索引作为序列化的关键字
'-------------------------------------------------------------------------------------------------
Public Function SerializeMulti(ParamArray arrInfo() As Variant) As String
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    Dim i           As Long
    On Error Resume Next
    If UBound(arrInfo) < 0 Then Exit Function
    If UBound(arrInfo) = 0 And IsArray(arrInfo(0)) Then
        objBag.WriteProperty "KL", UBound(arrInfo(0))
        For i = LBound(arrInfo(0)) To UBound(arrInfo(0))
            If IsArray(arrInfo(0)(i)) Then
                objBag.WriteProperty "KD" & i, 1
                objBag.WriteProperty "K" & i, SerializeMulti(arrInfo(0)(i))
            Else
                objBag.WriteProperty "K" & i, arrInfo(0)(i)
            End If
            If Err.Number = 330 Then
                '非法参数。  因为不支持持久性不能写对象。
                Err.Clear
                objBag.WriteProperty "K" & i, Nothing
            End If
        Next
    Else
        objBag.WriteProperty "KL", UBound(arrInfo)
        For i = 0 To UBound(arrInfo)
            If IsArray(arrInfo(i)) Then
                objBag.WriteProperty "KD" & i, 1
                objBag.WriteProperty "K" & i, SerializeMulti(arrInfo(i))
            Else
                objBag.WriteProperty "K" & i, arrInfo(i)
            End If
            If Err.Number = 330 Then
                '非法参数。  因为不支持持久性不能写对象。
                Err.Clear
                objBag.WriteProperty "K" & i, Nothing
            End If
        Next
    End If
    bytData = objBag.Contents
    SerializeMulti = EncodeBase64(bytData())
End Function

'--------------------------------------------------------------------------------------------------
'方法           UnSerializeMulti        获取序列的对象
'返回值         Variant                 序列化的对象数组
'入参列表:
'参数名         类型                    说明
'strSource      String                  序列化字符串
'[      ]       long                    按0开始索引，索引作为序列化的关键字
'-------------------------------------------------------------------------------------------------
Public Function UnSerializeMulti(ByVal strSource As String) As Variant
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    Dim i           As Long, lngLen     As Long
    Dim arrVar()    As Variant
    
    On Error Resume Next
    If Len(strSource) = 0 Then Exit Function
    bytData = DecodeBase64(strSource, True)
    objBag.Contents = bytData
    lngLen = objBag.ReadProperty("KL")
    If lngLen > -1 Then
        ReDim Preserve arrVar(lngLen)
        For i = 0 To lngLen
            If Not IsObject(objBag.ReadProperty("K" & i)) Then
                If objBag.ReadProperty("KD" & i, 0) = 1 Then
                    arrVar(i) = UnSerializeMulti(arrVar(i))
                Else
                    arrVar(i) = objBag.ReadProperty("K" & i)
                End If
            Else
                Set arrVar(i) = objBag.ReadProperty("K" & i)
            End If
        Next
    End If
    UnSerializeMulti = arrVar()
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
'       实际数据存储长度
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'参数：strText=处理字符
'         blnCrlf=是否去掉换行符
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function

'--------------------------------------------------------------------------------------------------
'方法           InCollection
'功能           检查集合中是否存在某元素
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'cllTest        Collection              要检查的集合
'strKey         String                  要检查的Key
'-------------------------------------------------------------------------------------------------
Public Function InCollection(cllTest As Collection, strKey As String) As Boolean
    On Error GoTo ErrorH
    If VarType(cllTest.Item(strKey)) = vbObject Then
    End If
    InCollection = True
    Exit Function
ErrorH:
    InCollection = False
End Function

'--------------------------------------------------------------------------------------------------
'方法           GetTickCountDiff
'功能           计算GetTickCcout的差值。由于 GetTickCountVB会产生负值以及归零现象因此需要单独处理
'返回值         Double
'入参列表:
'参数名         类型                    说明
'lngStart       Long                    起始时间
'lngEnd         Long                    结束时间缺省不传
'blnInputEnd    Boolean                 标识是否传入了lngEnd
'-------------------------------------------------------------------------------------------------
Public Function GetTickCountDiff(ByVal lngStart As Long, Optional ByVal lngEnd As Long, Optional ByVal blnInputEnd As Boolean) As Double
    Dim lngCur          As Long
    Const M_OFFSET_4    As Double = 4294967296#         '无符号整形的最大值
    If blnInputEnd Then
        lngCur = lngEnd
    Else
        lngCur = GetTickCount
    End If
    If lngCur < lngStart Then
        GetTickCountDiff = M_OFFSET_4 - LongToUnsigned(lngStart) + LongToUnsigned(lngCur)
    Else
        GetTickCountDiff = lngCur - lngStart
    End If
End Function

Private Function LongToUnsigned(value As Long) As Double
    Const M_OFFSET_4    As Double = 4294967296#         '无符号整形的最大值
    If value < 0 Then LongToUnsigned = value + M_OFFSET_4 Else LongToUnsigned = value
End Function

Public Function To_Date(ByVal strDate As String, Optional ByVal strType As String = "YMDHMS") As String
'功能：获取ORACLE Date类型串
'参数：strDate=时间字符串
'         strType=格式字符串类型，ymd-年月日（yyyy-mm-dd)，ymdhm-（yyyy-mm-dd hh:mm),ymdhms-（yyyy-mm-dd hh:mm:ss)
'返回：ORACLE Date类型串
    If Not IsDate(strDate) Then To_Date = "Null": Exit Function
    Select Case UCase(strType)
        Case "YMD"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "YMDHM"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Case "YMDHMS"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
           To_Date = "Null"
    End Select
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'功能：返回VB最大支持的版本号形式:9999.9999.9999.9999,最小版本号0000.0000.0000.0000
'参数：strVer=当前版本号
'           blnMax=True,若果为空，则返回最大支持版本，False=若果为空，则返回最小支持版本
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '增加一段，以兼容特殊SP版本号
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

Public Function IsVerSion(ByVal strVer As String, Optional ByVal blnOnlyCheckSpecial As Boolean) As Boolean
'功能：判断字符串是否是版本号
'blnOnlyCheckSpecial=是否是特殊SP版本号
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    If blnOnlyCheckSpecial And UBound(arrVer) <> 3 Then Exit Function
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If i = 3 Then
            If Format(Val(arrVer(i)), "0000") <> Format(Trim(arrVer(i)), "0000") Then Exit Function
        Else
            If Val("1" & arrVer(i)) & "" <> Trim("1" & arrVer(i)) Then Exit Function
        End If
    Next
    
    IsVerSion = True
End Function

'密码加密程序
Public Function Cipher(ByVal strText As String) As String
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    Dim strSeed As String
    
    If strText = "" Then Exit Function
    '获取随机种子
    '随机种子的随机数为999
    Rnd (-1)
    Randomize (999)
    strSeed = "456"
    intSeedLen = Len(strSeed)
    strDeText = Chr(intSeedLen + MIN_ASC)
    For i = 1 To intSeedLen
        intChr = Asc(Mid(strSeed, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        End If
    Next
    Rnd (-1)
    Randomize (Val(strSeed))
    intlen = Len(strText)
    For i = 1 To intlen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        ElseIf intChr < 0 Then
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    Cipher = strDeText
End Function

Public Function DeCipher(ByVal strText As String) As String
'密码解密程序
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '随机种子长度
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '采用旧的随机算法
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
        '获取随机种子
        '随机种子的随机数为999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
            If intChr >= MIN_ASC And intChr <= MAX_ASC Then
                intChr = intChr - MIN_ASC
                lngOffset = Int((NUM_ASC + 1) * Rnd())
                intChr = ((intChr - lngOffset) Mod NUM_ASC)
                If intChr < 0 Then
                    intChr = intChr + NUM_ASC
                End If
                intChr = intChr + MIN_ASC
                strDeText = strDeText & Chr(intChr)
            End If
        Next
        If Not IsNumeric(strDeText) Then
            strDeText = "123"
            intStart = 1
        Else
            intStart = 2 + intSeedLen
        End If
    Else
        strDeText = "123"
        intStart = 1
    End If
        
    '内容解密的种子
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intlen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr - lngOffset) Mod NUM_ASC)
            If intChr < 0 Then
                intChr = intChr + NUM_ASC
            End If
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        Else
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    DeCipher = strDeText
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function IsEmptyArray(varAnyArray As Variant) As Boolean
    Dim lngUbound               As Long
    On Error GoTo ErrH
    
    If IsEmpty(varAnyArray) Then
        IsEmptyArray = True
    ElseIf IsArray(varAnyArray) Then
        lngUbound = UBound(varAnyArray)
        IsEmptyArray = (lngUbound - LBound(varAnyArray)) < 0
    Else
        IsEmptyArray = True
    End If
    Exit Function
ErrH:
    IsEmptyArray = True
End Function

'
'Private Function GetSafeArrayInfo(varAnyArray As Variant, Optional intGetDimension As Integer, Optional lngLowerBound As Long, Optional lngUpperBound As Long, Optional lngElements As Long, Optional intFlags As Integer, Optional lngcbElements As Long, Optional lngcLocks As Long, Optional lngAddressOfData As Long) As Long
''**************************************************************
''*  数组头地址 = GetSafeArrayInfo(数组 ,[维数],[下标],[上标],[元素个数],[属性],[元素长度],[锁定计数],[首元素地址])
''*  result:  成功返回数组头地址;  返回 0 代表这是一个未被初始化过的数组
''*  note:    除第1个参数外,均为输出型参数,要获取数组的什么信息,传入相应的变量,执行后,变量的值即为相应的数组数据
''*           要获取多维数组的非第1维的上/下标,应按照下面的方法调用
''*           GetSafeArrayInfo (数组, 指定一个维数,[下标],[上标]
''**************************************************************
'
'    Dim lngArrayHeaderAddress       As Long
'    Dim intDimCount                 As Integer     '维数计次
'
'    CopyMemory lngArrayHeaderAddress, ByVal VarPtr(varAnyArray) + 8, 4
'    CopyMemory lngArrayHeaderAddress, ByVal lngArrayHeaderAddress, 4                                            '获取数组头地址
'    If lngArrayHeaderAddress < 1 Then Exit Function
'    CopyMemory intDimCount, ByVal lngArrayHeaderAddress, 2                                                      '获取数组维数
'    If intGetDimension > intDimCount Then Exit Function                                                         '若指定的维数大于实际维数则退出
'    CopyMemory lngLowerBound, ByVal (lngArrayHeaderAddress + 16 + (intDimCount - intGetDimension) * 8) + 4, 4   '获取下标
'    CopyMemory lngElements, ByVal (lngArrayHeaderAddress + 16 + (intDimCount - intGetDimension) * 8), 4         '获取指定维数下的元素个数
'
'    lngUpperBound = lngElements + lngLowerBound - 1                             '获取指定维数下的上标
'    CopyMemory intFlags, ByVal lngArrayHeaderAddress + 2, 2                     '获取数组属性
'    CopyMemory lngcbElements, ByVal lngArrayHeaderAddress + 4, 2                '获取数组单个元素长度
'    CopyMemory lngcLocks, ByVal lngArrayHeaderAddress + 8, 2                      '获取数组锁定计数
'    CopyMemory lngAddressOfData, ByVal lngArrayHeaderAddress + 12, 2              '获取数组首元素地址
'    GetSafeArrayInfo = lngArrayHeaderAddress
'End Function
