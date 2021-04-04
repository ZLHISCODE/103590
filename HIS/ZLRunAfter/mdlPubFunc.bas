Attribute VB_Name = "mdlPubFunc"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2018/12/25
'模块           mdlPubFunc
'说明
'==================================================================================================
Private Const mstrCurModule     As String = "mdlPubFunc"           '当前模块名称
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'返回（retrieve）从操作系统启动所经过（elapsed）的毫秒数
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'字符串用UTF-8编码
Public Const CP_UTF8 = 65001
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
Public Const G_UA_PWD           As String = "FA74C8A530DE7E088B1ACA673DD6297D"
Public Const G_UA_KEY           As String = "0016FDE250354FA9A4BA45433DBCC35D"
Public Const G_INTERFACE_KEY    As String = "EBA1D9B8CCCB4FD0804672DEDB222CFB"
Public Const G_APP_KEY          As String = "FD304782E75C41FDB14CB7A92A8A0B97"
'SM4加密
'/**
' * \brief          SM4-ECB block encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param input    input block
' * \param output   output block
' */
Private Declare Function sm4_crypt_ecb Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, key As Byte, in_put As Byte, out_put As Byte) As Long
'SM4分组密码加密
'/**
' * \brief          SM4-CBC buffer encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param iv       initialization vector (updated after use)
' * \param input    buffer holding the input data
' * \param output   buffer holding the output data
' */
Private Declare Function sm4_crypt_cbc Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, iv As Byte, key As Byte, in_put As Byte, out_put As Byte) As Long
'获取字符串的哈希编码
'/**
' * \brief          Output = SM3( input buffer )
' *
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   SM3 checksum result
' */
Private Declare Sub sm3_hash Lib "zlSm4.dll" Alias "sm3" (in_put As Byte, ByVal Length As Long, out_put As Byte)
'获取文件的sm哈希编码
'/**
' * \brief          Output = SM3( file contents )
' *
' * \param path     input file name
' * \param output   SM3 checksum result
' *
' * \return         0 if successful, 1 if fopen failed,
' *                 or 2 if fread failed
' */
Private Declare Function sm3_file_hash Lib "zlSm4.dll" Alias "sm3_file" (in_path As Byte, out_put As Byte) As Long
'HMAC是密钥相关的哈希运算消息认证码，HMAC运算利用哈希算法，以一个密钥和一个消息为输入，生成一个消息摘要作为输出。
'/**
' * \brief          Output = HMAC-SM3( hmac key, input buffer )
' *
' * \param key      HMAC secret key
' * \param keylen   length of the HMAC key
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   HMAC-SM3 result
' */
Private Declare Sub sm3_hmac_hash Lib "zlSm4.dll" Alias "sm3_hmac" (key As Byte, ByVal keylen As Long, in_put As Byte, ByVal inputLen As Long, out_put As Byte)
'获取ZLSM4的修改版本
'1:只支持sm4_crypt_ecb,sm4_crypt_cbc
'2:增加支持sm3，sm3_file，sm3_hmac，sm_version
'/**
' * \brief          Output = zlSM4.DLL Version
' */
Private Declare Function get_sm_version Lib "zlSm4.dll" Alias "sm_version" () As Long
Public Const SM4_CRYPT_RANDOMIZE_KEY As Long = 999  'sm4加密算法密钥生成器的随机种子
Public Const SM4_CRYPT_RANDOMIZE_IV As Long = 666   'sm4加密算法初始向量生成器的随机种子
Private M_SM4_VERSION As Long
Private Enum CrypeMode
    CM_Encrypt = 1   '加密
    CM_Decrypt = 0   '解密
End Enum
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public gblnCurShow      As Boolean
'======================================================================================================================
'方法           Sm4EncryptEcb           SM4加密
'返回值         String                  加密后的值,格式：ZLSV+版本号+:+加密后的字符串
'入参列表:
'参数名         类型                    说明
'strInput       String                  要加密的字符串
'strKey         String(Optional)        加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4EncryptEcb(ByVal strInput As String, Optional ByVal strKey As String) As String
    Dim arrKey()    As Byte
    Dim arrInput()  As Byte
    Dim arrOutPut() As Byte
    Dim lngLength   As Long
    
    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput = "" Then
        Sm4EncryptEcb = ""
    Else
        arrKey = GetKey(strKey, 2)
        arrInput = BytePadding(strInput, M_SM4_VERSION)
        ReDim arrOutPut(UBound(arrInput))
        Call sm4_crypt_ecb(CM_Encrypt, UBound(arrInput) + 1, arrKey(0), arrInput(0), arrOutPut(0))
        Sm4EncryptEcb = "ZLSV" & M_SM4_VERSION & ":" & ByteToHexString(arrOutPut())
    End If
End Function

'======================================================================================================================
'方法           Sm4DecryptEcb           SM4解密
'返回值         String                  解密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  要解密的字符串（该字符串是Sm4EncryptEcb生成的结果）
'strKey         String(Optional)        加密密钥也就是解密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4DecryptEcb(ByVal strInput As String, Optional ByVal strKey As String) As String
    Dim arrKey()        As Byte
    Dim arrInput()      As Byte
    Dim arrOutPut()     As Byte
    Dim lngLength       As Long
    Dim lngVersion      As Long

    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput Like "ZLSV*:*" Then
        lngVersion = Val(Mid(strInput, 5, InStr(strInput, ":") - 5))
        strInput = Mid(strInput, InStr(strInput, ":") + 1)
        '当前客户端的ZLSM4不支持该版本的加密字符串解密，仍旧解密，因为一般来说都能解密出相同的字符串
'        If lngVersion > M_SM4_VERSION Then
'            Exit Function
'        End If
    Else
        Exit Function
    End If
    
    arrKey = GetKey(strKey, 2)
    arrInput = HexStringToByte(strInput)
    ReDim arrOutPut(UBound(arrInput))
    
    Call sm4_crypt_ecb(CM_Decrypt, UBound(arrInput) + 1, arrKey(0), arrInput(0), arrOutPut(0))
    If lngVersion = 1 Then
        Sm4DecryptEcb = Trim(StrConv(arrOutPut(), vbUnicode))
    Else
        Sm4DecryptEcb = TruncZero(StrConv(arrOutPut(), vbUnicode))
    End If
End Function
'======================================================================================================================
'方法           Sm4EncryptCbc           SM4分组加密
'返回值         String                  加密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  要加密的字符串
'strKey         String(Optional)        加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'strIv          String(Optional)        分组加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4EncryptCbc(ByVal strInput As String, Optional ByVal strKey As String, Optional ByVal strIv As String) As String
    Dim arrKey()        As Byte
    Dim arrInput()      As Byte
    Dim arrOutPut()     As Byte
    Dim arrIv()         As Byte
    Dim lngLength       As Long
    
    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput = "" Then
        Sm4EncryptCbc = ""
    Else
        arrKey = GetKey(strKey, 2)
        arrIv = GetKey(strIv, 1)
        
        arrInput = BytePadding(strInput, M_SM4_VERSION)
        ReDim arrOutPut(UBound(arrInput))
        
        Call sm4_crypt_cbc(CM_Encrypt, UBound(arrInput) + 1, arrIv(0), arrKey(0), arrInput(0), arrOutPut(0))
        Sm4EncryptCbc = "ZLSV" & M_SM4_VERSION & ":" & ByteToHexString(arrOutPut)
    End If
End Function

'======================================================================================================================
'方法           Sm4EncryptCbc           SM4分组加密对应的解密过程
'返回值         String                  解密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  已经加密的字符串
'strKey         String(Optional)        解密密钥也就是加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'strIv          String(Optional)        分组解密密钥也就是分组加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4DecryptCbc(ByVal strInput As String, Optional ByVal strKey As String, Optional ByVal strIv As String) As String
    Dim arrDest() As Byte
    Dim arrJiemi() As Byte
    Dim arrKey() As Byte
    Dim arrInput() As Byte
    Dim arrOutPut() As Byte
    Dim arrIv() As Byte
    Dim lngLength As Long
    Dim lngVersion      As Long

    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput Like "ZLSV*:*" Then
        lngVersion = Val(Mid(strInput, 5, InStr(strInput, ":") - 5))
        strInput = Mid(strInput, InStr(strInput, ":") + 1)
        '当前客户端的ZLSM4不支持该版本的加密字符串解密，仍旧解密，因为一般来说都能解密出相同的字符串
'        If lngVersion > M_SM4_VERSION Then
'            Exit Function
'        End If
    Else
        Exit Function
    End If
    
    arrKey = GetKey(strKey, 2)
    arrIv = GetKey(strIv, 1)
    
    arrInput = HexStringToByte(strInput)
    ReDim arrOutPut(UBound(arrInput))

    Call sm4_crypt_cbc(CM_Decrypt, UBound(arrInput) + 1, arrIv(0), arrKey(0), arrInput(0), arrOutPut(0))
    
    If lngVersion = 1 Then
        Sm4DecryptCbc = Trim(StrConv(arrOutPut(), vbUnicode))
    Else
        Sm4DecryptCbc = TruncZero(StrConv(arrOutPut(), vbUnicode))
    End If
End Function

'======================================================================================================================
'方法           Sm3                     计算字符串的哈希值（用来检测字符串的变动）
'返回值         String(32)              字符串的哈希值
'入参列表:
'参数名         类型                    说明
'strInput       String                  字符串内容
'======================================================================================================================
Public Function Sm3(ByRef strInput As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte

    '先将字符串由 Unicode 转成系统的缺省码页
    arrInput = StrConv(strInput, vbFromUnicode)
    lngLength = UBound(arrInput) + 1
    
    Call sm3_hash(arrInput(0), lngLength, arrOut(0))
    '将返回值转换为16进制字符串
    Sm3 = ByteToHexString(arrOut)
End Function
'======================================================================================================================
'方法           Sm3_File                计算文件的哈希值（用来检测 文件内容的变动）
'返回值         String(32)              文件的哈希值
'入参列表:
'参数名         类型                    说明
'strFile        String                  文件路径
'======================================================================================================================
Public Function Sm3_File(ByRef strFile As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte
    Dim lngReturn As Long

    '先将字符串由 Unicode 转成系统的缺省码页
    arrInput = StrConv(strFile, vbFromUnicode)
    '由于API没有传递长度，因此增加字符串Chr(0)
    lngLength = UBound(arrInput) + 1
    ReDim Preserve arrInput(lngLength)
    '计算结果
    lngReturn = sm3_file_hash(arrInput(0), arrOut(0))
    '判断是否成功处理
    If lngReturn = 0 Then
        '将返回值转换为16进制字符串
        Sm3_File = ByteToHexString(arrOut)
    ElseIf lngReturn = 1 Then
        Sm3_File = "ERROR:文件打开失败"
    ElseIf lngReturn = 2 Then
        Sm3_File = "ERROR:文件读取失败"
    End If
End Function
'======================================================================================================================
'方法           sm3_hmac                给定义一个密钥对传入的消息产生消息摘要
'返回值         String(32)              密钥加密消息后生成的消息摘要
'入参列表:
'参数名         类型                    说明
'strKey         String                  密钥
'strMsg         String                  消息内容
'======================================================================================================================
Public Function sm3_hmac(ByRef strKey As String, ByVal strMsg As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte
    Dim arrKey()    As Byte
    Dim lngKeyLen   As Long
    
    '先将字符串由 Unicode 转成系统的缺省码页
    arrInput = StrConv(strMsg, vbFromUnicode)
    lngLength = UBound(arrInput) + 1
    '先将字符串由 Unicode 转成系统的缺省码页
    arrKey = StrConv(strKey, vbFromUnicode)
    lngKeyLen = UBound(arrKey) + 1
    Call sm3_hmac_hash(arrKey(0), lngKeyLen, arrInput(0), lngLength, arrOut(0))
    '将返回值转换为16进制字符串
    sm3_hmac = ByteToHexString(arrOut)
End Function
'======================================================================================================================
'方法           sm_version              获取ZLSM4的版本号
'返回值         Long                    ZLSM4的版本号
'入参列表:
'======================================================================================================================
Public Function sm_version() As Long
    Dim lngVersion As Long
    On Error Resume Next
    lngVersion = get_sm_version
    If Err.Number <> 0 Then
        Err.Clear
        sm_version = 1
    Else
        sm_version = lngVersion
    End If
End Function
'======================================================================================================================
'方法           ByteToHexString         将字节组转换为16进制字符串
'返回值         String                  字节组转换的16进制字符串
'入参列表:
'参数名         类型                    说明
'bytInpu        Byte(）                 字节数组
'======================================================================================================================
Public Function ByteToHexString(bytInpu() As Byte) As String
    Dim i           As Long
    Dim strReturn   As String
    
    For i = LBound(bytInpu) To UBound(bytInpu)
        If Len("" & Hex(bytInpu(i))) = 1 Then
            strReturn = strReturn & "0" & Hex(bytInpu(i))
        Else
            strReturn = strReturn & Hex(bytInpu(i))
        End If
    Next
    
    ByteToHexString = strReturn
End Function
'======================================================================================================================
'方法           ByteToHexString         将16进制字符串转换为字节组
'返回值         Byte()                  16进制字符串转换的字节组
'入参列表:
'参数名         类型                    说明
'bstrInput      String                  16进制字符串
'lngRetBytLen   Long(Optional)          指定返回的字节组的长度,0-按原始长度返回，<>0返回指定的长度，不足补齐（补0），多了截取
'======================================================================================================================
Public Function HexStringToByte(ByVal strInput As String, Optional ByVal lngRetBytLen As Long) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    Dim lngLen      As Long
    
    lngLen = Len(strInput)
    If lngRetBytLen <> 0 Then
        lngLen = lngLen \ 2
        If lngLen > lngRetBytLen Then
            lngLen = lngRetBytLen
        End If
        ReDim arrReturn(lngRetBytLen - 1)
    Else
        lngLen = lngLen \ 2
        ReDim arrReturn(lngLen - 1)
    End If
    
    For i = 0 To lngLen - 1
        arrReturn(i) = Val("&H" & Mid(strInput, 2 * i + 1, 2))
    Next
    
    HexStringToByte = arrReturn()
End Function

'======================================================================================================================
'方法           BytePadding             将指定字符串按照16字节补齐，
'返回值         Byte()                  补齐后的字符串字节组
'入参列表:
'参数名         类型                    说明
'strInput       String                  字符串
'lngVersion     Long(Optional,2)        字符串补齐的版本（ZLSM4.DLL的版本，以及加密算法前缀中的版本），1-空格补齐，>1:Chr(0)补齐
'lngPaddingNum  Long(Optional,16)        补齐的字节数，缺省按照16进制补齐
'======================================================================================================================
Public Function BytePadding(ByVal strInput As String, Optional ByVal lngVersion As Long = 2, Optional ByVal lngPaddingNum As Long = 16) As Byte()
    Dim arrReturn()     As Byte
    Dim lngLenBef       As Long
    Dim i               As Long
    Dim lngLenAft       As Long
    
    '先将字符串由 Unicode 转成系统的缺省码页
    arrReturn = StrConv(strInput, vbFromUnicode)
    lngLenBef = UBound(arrReturn) + 1
    '判断得到的数组的长度，若不是16的整数倍，则补空格或:Chr(0)
    lngLenAft = -Int(-lngLenBef / lngPaddingNum) * lngPaddingNum
    If lngLenBef <> lngLenAft Then
        ReDim Preserve arrReturn(lngLenAft - 1)
        For i = lngLenBef To lngLenAft - 1
            If lngVersion = 1 Then
                arrReturn(i) = 32
            Else
                arrReturn(i) = 0
            End If
        Next
    End If
    BytePadding = arrReturn()
End Function


Private Function GetKey(ByVal strKey As String, ByVal intType As Integer) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    If strKey <> "" Then
        arrReturn = HexStringToByte(strKey, 16)
    Else
        ReDim arrReturn(15)
        If intType = 0 Then
            For i = 0 To 15
                arrReturn(i) = i * 15
            Next
        ElseIf intType = 1 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_IV)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        ElseIf intType = 2 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_KEY)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        End If
    End If
    GetKey = arrReturn
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

'--------------------------------------------------------------------------------------------------
'方法           IsDesinMode
'功能           确定当前模式为设计模式（源码环境）
'返回值         Boolean
'-------------------------------------------------------------------------------------------------
Public Function IsDesinMode() As Boolean
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
    Dim i           As Long, j              As Long, lngBit     As Long
    
    On Error GoTo errH
    
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
            bytOut(j) = (bytInput(i) And &HFC) \ &H4
            j = j + 1
            bytOut(j) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两位
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HF0) \ &H10)
            j = j + 1
            bytOut(j) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四位
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HC0) / &H40)
            j = j + 1
            bytOut(j) = bytInput(i) And &H3F
            j = j + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) '转换为Base64字符
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") '原文剩余内容不足3个字节需要补齐
    Exit Function
errH:
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
    Dim i           As Long, j              As Long, lngBit     As Long
    Dim lngModLen       As Long
    On Error GoTo errH
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
 
    For j = 0 To lngInputLen - 1
        bytInput(j) = DecBase64Char(bytInput(j)) '从Base64字符转换为6-bit字节
    Next
    '将6-bit字节数组转换为8-bit字节数组
    For j = 0 To lngOutLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(j) = bytInput(i) * &H4
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H30) \ &H10)
            lngBit = 2
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两字节
            bytOut(j) = (bytInput(i) And &HF) * &H10
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H3C) \ &H4)
            lngBit = 4
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四字节
            bytOut(j) = (bytInput(i) And &H3) * &H40
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or bytInput(i)
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
errH:
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
    
    On Error GoTo errH
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
errH:
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
    
    On Error GoTo errH
    If strFile = "" Then
        strFile = gobjFSO.GetSpecialFolder(TemporaryFolder) & "\" & gobjFSO.GetTempName
    Else
        If gobjFSO.FileExists(strFile) Then Kill strFile
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
errH:
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
'    If IsObject(objInfo) Then
''        If objInfo Is Nothing Then Exit Function
'    End If
    objBag.WriteProperty strKeyName, objInfo
    bytData = objBag.Contents
    Serialize = EncodeBase64(bytData())
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
    bytData = DecodeBase64(strSource, True)
    objBag.Contents = bytData
    If IsObject(objBag.ReadProperty(strKeyName)) Then
        Set UnSerialize = objBag.ReadProperty(strKeyName)
    Else
        UnSerialize = objBag.ReadProperty(strKeyName)
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
    objBag.WriteProperty "KL", UBound(arrInfo)
    For i = 0 To UBound(arrInfo)
        objBag.WriteProperty "K" & i, arrInfo(i)
    Next
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
            If IsObject(objBag.ReadProperty("K" & i)) Then
                Set arrVar(i) = objBag.ReadProperty("K" & i)
            Else
                arrVar(i) = objBag.ReadProperty("K" & i)
            End If
        Next
    End If
    UnSerializeMulti = arrVar()
End Function

Public Function FullDate(ByVal strText As String) As String
'功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd[ HH:mm])
    Dim curDate As Date, strTmp As String
    
    If strText = "" Or Len(strText) <> 14 Then Exit Function
    strTmp = strText
    '当作输入yyyyMMddHHmm
    strTmp = Format(strTmp, "00000000000000")
    strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Right(strTmp, 2)
    FullDate = strTmp
End Function

Public Function CopyNewRec(Optional ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'功能：复制记录集或者构造一个自定义记录集
'参数：strFields=需要复制的记录集的字段的列顺序或字段名组成的字符串
'          如：1 别名1,3 别名2,7 别名3...表示复制记录集的第1,3,7..字段组成记录集并返回
'              ID 别名1,姓名 别名2,....表示复制记录集的ID,姓名...字段组成记录集返回
'              别名*为新的记录集的列名
'              两中类型混搭容易出现列名相同的问题，请注意
'              *,在表示复制原记录集的所有字段的占位符，可能需要将原来的字段全部复制，同时增加别名列来判断改变
'           arrAppFields=追加的字段信息：列名,类型,长度,默认值,没有默认值传Empty,没有指定长度传Empty
'      blnOnlyStructure=是否只复制结构（rsSource传递时才生效）
'备注：1）在程序中，经常会涉及到相互传递记录集，而使用ADO的Clone复制产生的记录集，当其中一个记录集的数据发生变化的时候，所有副本都将发生相同的变化（通常指修改或删除），而我们往往希望这些记录集相互间保持独立
'      2)有时我们需要一种表类型的数据结构来存储数据，该函数可以产生一个自定义记录集来实现
'应用场景：
'             1）CopyNewRec(rsSource），全部复制结构以及数据
'             2）CopyNewRec(rsSource,True），只产生结构不复制数据
'             3）CopyNewRec(rsSource,,"ID 别名1,姓名")复制原纪录集的ID与性名列的数据，产生的新记录集列为别名1，姓名。若要只复制结构，blnOnlyStructure传True
'             4)CopyNewRec(rsSource,,"*,标志 新标志")复制原纪录集的所有字段，并增加新列“新标志”该列数据来源“标志列”，该中类型用来判断部分数据变化
'             5)CopyNewRec(rsSource,,,Array("是否改变", adInteger, 1, 0)），全部复制结构以及数据，新增一个空列是否改变
'             5）CopyNewRec(Nothing, , , Array("系统编号", adInteger, 5, Empty, "所有者", adVarChar, 100, Empty)) 产生一个自定义记录集
    Dim rsClone As ADODB.Recordset
    Dim rsTarget As ADODB.Recordset
    Dim intFields As Integer
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant, arrFieldsTmp As Variant
    Dim i As Long
    
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '产生记录集结构
        If strFields = "" Then
            strFields = "*"
        End If
        arrFieldsTmp = Split(strFields, ",")
        arrFieldsName = Array()
        For intFields = LBound(arrFieldsTmp) To UBound(arrFieldsTmp)
            If Trim(arrFieldsTmp(intFields)) = "*" Then '标识此处将增加原记录集的所有列
                If Not rsClone Is Nothing Then
                    For i = 0 To rsClone.Fields.Count - 1
                        ReDim Preserve arrFieldsName(UBound(arrFieldsName) + 1)
                        arrFieldsName(UBound(arrFieldsName)) = rsClone.Fields(i).Name & ""
                        .Fields.Append rsClone.Fields(i).Name, IIf(rsClone.Fields(i).Type = adNumeric, adDouble, rsClone.Fields(i).Type), rsClone.Fields(i).DefinedSize, adFldIsNullable    '0:表示新增
                    Next
                End If
            Else
                ReDim Preserve arrFieldsName(UBound(arrFieldsName) + 1)
                '列包含别名
                arrTmp = Split(arrFieldsTmp(intFields) & " ", " ")
                strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).Name & ""
                '获取字段原名，存入数组
                arrFieldsName(UBound(arrFieldsName)) = strFieldName
                '添加字段,若果存在别名，则新增列的列名为别名
                .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:表示新增
            End If
        Next
        
        '追加字段添加
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '复制数据
        If Not blnOnlyStructure Then
            If rsClone Is Nothing Then Set CopyNewRec = rsTarget: Exit Function
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '新记录集的列按顺序添加，因此可以这样
                    .Fields(intFields).Value = rsClone.Fields(arrFieldsName(intFields)).Value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Function RecDelete(ByRef rsInput As ADODB.Recordset, Optional ByVal strFilter As String) As Boolean
'功能：删除指定条件的记录集的记录
'参数：rsInput=记录集
'      strFilter=条件
'返回：是否成功
'      rsInput=经过删除后的记录集
    If Not rsInput Is Nothing Then
        rsInput.Filter = strFilter
        If rsInput.RecordCount > 0 Then
            rsInput.MoveFirst
            Do While Not rsInput.EOF
                Call rsInput.Delete
                rsInput.MoveNext
            Loop
            Call rsInput.UpdateBatch
        End If
    End If
    RecDelete = True
End Function

Public Function RecUpdate(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
'功能：更新指定条件的记录集的记录
'参数：rsInput=记录集
'      strFilter=条件
'      arrInput=输入的字段名以及值，格式：字段名1,值1, 字段名2,值2,....
'返回：是否成功
'      rsInput=经过更新后的记录集
'说明：arrInput的字段值可以用记录集中的其他字段来更新该字段，此时格式为：!字段名 处理函数(暂时支持Val)
    Dim strFiledName As String, strFileValue As String, strFun As String, strFindFiled As String
    Dim blnFiled As Boolean, i As Long
    Dim arrTmp As Variant
    
    If rsInput Is Nothing Then Exit Function
    On Error GoTo errH
    With rsInput
        .Filter = strFilter
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            For i = LBound(arrInput) To UBound(arrInput) Step 2
                strFiledName = arrInput(i)
                If arrInput(i + 1) & "" = "" Then
                    rsInput(strFiledName).Value = Null
                Else
                    strFun = ""
                    strFindFiled = arrInput(i + 1)
                    If arrInput(i + 1) Like "!?*" Then
                        blnFiled = True
                        On Error Resume Next
                        strFindFiled = Mid(arrInput(i + 1), 2)
                        arrTmp = Split(strFindFiled & " ", " ")
                        strFindFiled = Trim(arrTmp(0))
                        strFun = Trim(arrTmp(1))
                        strFileValue = rsInput(strFindFiled).Value & ""
                        If Err.Number <> 0 Then Err.Clear: blnFiled = False
                        On Error GoTo errH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).Value = arrInput(i + 1)
                    Else
                        If strFun = "" Then
                            rsInput(strFiledName).Value = rsInput(strFindFiled).Value
                        ElseIf strFun = "Val" Then
                            rsInput(strFiledName).Value = Val(rsInput(strFindFiled).Value & "")
                        ElseIf strFun = "Trim" Then
                            rsInput(strFiledName).Value = Trim(rsInput(strFindFiled).Value & "")
                            If rsInput(strFiledName).Value & "" = "" Then
                                rsInput(strFiledName).Value = Null
                            End If
                        Else
                            rsInput(strFiledName).Value = rsInput(strFindFiled).Value
                        End If
                    End If
                End If
                blnFiled = False
            Next
            .MoveNext
        Loop
        Call rsInput.UpdateBatch
    End With
    RecUpdate = True
    Exit Function
errH:
    MsgBox Err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function RecDataAppend(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, Optional ByVal lngAppendRows As Long = -1, Optional ByVal strSourceFields As String, Optional ByVal strAppendFileds As String, Optional ByVal blnKeepBookMark As Boolean, Optional ByVal arrOtherFieldValue As Variant) As Boolean
'功能：将指定记录集的数据添加到另一个记录集上
'参数：rsSource=数据源记录集
'      rsAppend=追加的数据记录集
'      lngAppendRows=添加的行数，-1,表示全部添加，>=0表示至多添加N行
'      strSourceFields，strAppendFileds=段对应规则，该参数不传时，默认两记录集结构相同，格式：[记录集1].字段1,字段2...；[记录集2].字段1,字段2...,当为"-字段1,字段2"为整体字排除这些字段后剩余的一一对应
'      blnKeepBookMark:数据添加后是否将记录回归原位置
'      arrOtherFieldValue:部分无对应字段的值处理，格式：“字段名”,值。这部分字段不能在strSourceFields，strAppendFileds的对应规则出现（可以以-字段方式出现）
'返回：是否成功
'      rsSource=添加数据后的记录集
    Dim arrSource   As Variant, arrAppend As Variant
    Dim i           As Long, arrValues() As Variant, lngIdx As Long, arrTmp As Variant
    Dim lngCount    As Long, lngCurRows         As Long
    Dim varAppendBK      As Variant
    
    If rsAppend Is Nothing Then RecDataAppend = True: Exit Function
    If rsAppend.RecordCount = 0 Then RecDataAppend = True: Exit Function
    If rsSource Is Nothing Then Set rsSource = rsAppend: RecDataAppend = True: Exit Function
    If blnKeepBookMark Then
'        If Not rsSource.EOF Then varSourceBK = rsSource.Bookmark
        If Not rsAppend.EOF Then varAppendBK = rsAppend.Bookmark
    End If
    On Error GoTo errH
    If strSourceFields = "" Or strSourceFields Like "-*" Then
        arrTmp = Split(strSourceFields, ",")
        strSourceFields = "," & Trim(Mid(strSourceFields, 2)) & ","
        arrSource = Array()
        ReDim Preserve arrSource(rsSource.Fields.Count - 1 - (UBound(arrTmp) + 1))
        Erase arrTmp
        lngIdx = 0
        For i = 0 To rsSource.Fields.Count - 1
            If InStr(strSourceFields, "," & rsSource.Fields(i).Name & ",") = 0 Then
                arrSource(lngIdx) = rsSource.Fields(i).Name & ""
                lngIdx = lngIdx + 1
            End If
        Next
    Else
        arrSource = Split(strSourceFields, ",")
    End If

    If strAppendFileds = "" Or strAppendFileds Like "-*" Then
        strAppendFileds = "," & Trim(Mid(strAppendFileds, 2)) & ","
        arrAppend = Array()
        lngIdx = 0
        ReDim Preserve arrAppend((UBound(arrSource)))
        For i = 0 To rsAppend.Fields.Count - 1
            If InStr(strAppendFileds, "," & rsAppend.Fields(i).Name & ",") = 0 Then
                ReDim Preserve arrAppend(lngIdx)
                arrAppend(lngIdx) = rsAppend.Fields(i).Name & ""
                lngIdx = lngIdx + 1
            End If
        Next
    Else
        arrAppend = Split(strAppendFileds, ",")
    End If
    
    '多余的列不对应
    lngCount = UBound(arrSource)
    If lngCount > UBound(arrAppend) Then
        lngCount = UBound(arrAppend)
    End If
    '部分自定义字段的值处理
    If TypeName(arrOtherFieldValue) = "Variant()" Then
        ReDim arrValues(lngCount + (UBound(arrOtherFieldValue) + 1) / 2)
        ReDim Preserve arrSource(UBound(arrValues))
        For lngIdx = LBound(arrOtherFieldValue) To UBound(arrOtherFieldValue) Step 2
            arrSource(lngCount + 1 + lngIdx \ 2) = arrOtherFieldValue(lngIdx)
            arrValues(lngCount + 1 + lngIdx \ 2) = arrOtherFieldValue(lngIdx + 1)
        Next
    Else
        ReDim arrValues(lngCount)
    End If
    
    
    If lngAppendRows = -1 Then
        lngAppendRows = rsAppend.RecordCount
    End If
    
    Do While Not rsAppend.EOF
        lngCurRows = lngCurRows + 1
        If lngCurRows > lngAppendRows Then Exit Do
        For i = 0 To lngCount
            arrValues(i) = rsAppend(arrAppend(i)).Value
        Next
        rsSource.AddNew arrSource, arrValues
        rsAppend.MoveNext
    Loop
    If blnKeepBookMark Then
'        If Not IsEmpty(varSourceBK) Then rsSource.Bookmark = CDbl(varSourceBK)
        If Not IsEmpty(varAppendBK) Then rsAppend.Bookmark = varAppendBK
    End If
    RecDataAppend = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox Err.Description, vbInformation, gstrSysName
    Err.Clear
End Function

Public Function RecDistinct(ByVal rsSource As ADODB.Recordset, Optional ByVal strDisFieldsName As String, Optional ByVal strFieldsName As String) As ADODB.Recordset
'功能：记录集去重复
'参数：rsSource=要去重复的记录集
'strDisFieldsName=去重复的字段,为空，则对所有字段去重
'strFieldsName=返回结果集字段，为空，则返回去重复的字段
'返回：操作后的记录集
    Dim rsReturn As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String, strOldRow As String

    '读取默认字段名
    If strDisFieldsName = "" Then
        For i = 0 To rsSource.Fields.Count - 1
            strTmp = strTmp & "," & rsSource.Fields(i).Name
        Next
        strTmp = Mid(strTmp, 2)
        If strDisFieldsName = "" Then strDisFieldsName = strTmp
    End If
    If strFieldsName = "" Then strFieldsName = strDisFieldsName
    
    Set rsReturn = CopyNewRec(rsSource, , strFieldsName)
    If rsSource.RecordCount = 0 Then Set RecDistinct = rsReturn: Exit Function
    
    rsReturn.Sort = strDisFieldsName '排序，自动将光标移动到开头
    Do While Not rsReturn.EOF
        strTmp = rsReturn.GetString(, 1, "[ColumnSpliter]", , "[NULLEXP]") '自动移动光标
        rsReturn.MovePrevious
        If strTmp = strOldRow Then  '删除重复行
            Call rsReturn.Delete: Call rsReturn.Update
        Else
            strOldRow = strTmp
        End If
        rsReturn.MoveNext
    Loop
    rsReturn.Sort = strDisFieldsName
    Set RecDistinct = rsReturn
End Function

Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Or TypeName(objTxt) = "ComboBox" Then
        If Trim(objTxt.Text) = "" Then Exit Sub
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function CheckIsDBA(ByRef connThis As ADODB.Connection) As Boolean
'功能：判断当前用户是否为DBA角色
    Dim rsTemp As ADODB.Recordset
    Dim strSQL      As String
    
    On Error GoTo errH
    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = gobjRegister.OpenSQLRecord(connThis, strSQL, "判断当前连接用户是否具有DBA角色")
    CheckIsDBA = rsTemp.RecordCount > 0
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single = -1, Optional strSQL As String, Optional strServer As String, Optional blnPer As Boolean)
'功能：显示或隐藏等待或进度窗体(strInfo)
'参数:strInfo=等待或进度提示信息
'     sngPer=进度
    
    If gblnSilence Then Exit Sub
    If Not gblnShow Then
        If gblnCurShow Then
            ShowWindow frmFlash.hWnd, 0
            gblnCurShow = False
        End If
        Exit Sub
    End If
    
    If glngSec > 0 Then
        frmFlash.lblTip.Caption = (glngSec \ 10) & "秒后会自动隐藏到任务栏，若要查看详情，请点击任务栏图标。"
    Else
        frmFlash.lblTip.Visible = False
    End If
    
    If sngPer > 1 Then sngPer = 1

    If strInfo = "" Then
'        frmFlash.avi.Close
        ShowWindow frmFlash.hWnd, 0
        gblnCurShow = False
    Else
        gblnShow = True
        frmFlash.lblServer = "服务器：" & strServer
        frmFlash.txtSQL = strSQL
        frmFlash.lbl.Caption = strInfo
        If Not gblnCurShow Then
            On Error Resume Next
            If sngPer = -1 Then
                '显示等待
'                frmFlash.avi.Open GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "") & "\" & "Findfile.avi"
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                frmFlash.Height = 1700
                SetWindowPos frmFlash.hWnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                ShowWindow frmFlash.hWnd, 5
                
'                frmFlash.avi.Play
                frmFlash.Refresh
            Else
                frmFlash.Height = 3060
                frmFlash.picDo.Visible = True
                frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
                frmFlash.lblPer.Top = frmFlash.lbl.Top
                frmFlash.lblTip.Top = frmFlash.lbl.Top
                frmFlash.lblDo.Caption = String(50 * sngPer, frmFlash.lblDo.Tag)
                If blnPer Then
                    If sngPer > 0 Then
                        frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                    Else
                        frmFlash.lblPer.Caption = ""
                    End If
                    frmFlash.lblPer.Visible = True
                End If
                
                SetWindowPos frmFlash.hWnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                ShowWindow frmFlash.hWnd, 5
                
                frmFlash.Refresh
            End If
            gblnCurShow = True
        Else
            If sngPer >= 0 Then
                frmFlash.Height = 3060
                frmFlash.lblDo.Caption = String(50 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            Else
                frmFlash.Height = 1700
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '功能：在任务栏上增加一个图标
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '功能：从任务栏上删除图标
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '事件发生的载体
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
    
End Sub

