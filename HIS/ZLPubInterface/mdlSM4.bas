Attribute VB_Name = "mdlSM4"
Option Explicit
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

Private Enum CrypeMode
    CM_Encrypt = 1   '加密
    CM_Decrypt = 0   '解密
End Enum
Private M_SM4_VERSION As Long
Public Const SM4_CRYPT_RANDOMIZE_KEY As Long = 999  'sm4加密算法密钥生成器的随机种子
Public Const SM4_CRYPT_RANDOMIZE_IV As Long = 666   'sm4加密算法初始向量生成器的随机种子
Public Const G_PASSWORD_KEY             As String = "3357F1F2CA0341A5B75DBA7F35666715"
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

Public Function GetGeneralAccountKey(ByRef strKey As String) As String
    Dim arrTmp()    As Byte
    Dim i           As Long
    arrTmp = HexStringToByte(strKey, 16)
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i Mod 2 = 0 Then
            arrTmp(i) = 255 - arrTmp(i)
        ElseIf i Mod 3 = 0 Then
            arrTmp(i) = (arrTmp(i) + i) Mod 256
        End If
    Next
    GetGeneralAccountKey = ByteToHexString(arrTmp)
End Function
