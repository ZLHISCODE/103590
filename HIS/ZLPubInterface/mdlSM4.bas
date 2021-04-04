Attribute VB_Name = "mdlSM4"
Option Explicit
'SM4����
'/**
' * \brief          SM4-ECB block encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param input    input block
' * \param output   output block
' */
Private Declare Function sm4_crypt_ecb Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, key As Byte, in_put As Byte, out_put As Byte) As Long
'SM4�����������
'/**
' * \brief          SM4-CBC buffer encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param iv       initialization vector (updated after use)
' * \param input    buffer holding the input data
' * \param output   buffer holding the output data
' */
Private Declare Function sm4_crypt_cbc Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, iv As Byte, key As Byte, in_put As Byte, out_put As Byte) As Long
'��ȡ�ַ����Ĺ�ϣ����
'/**
' * \brief          Output = SM3( input buffer )
' *
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   SM3 checksum result
' */
Private Declare Sub sm3_hash Lib "zlSm4.dll" Alias "sm3" (in_put As Byte, ByVal Length As Long, out_put As Byte)
'��ȡ�ļ���sm��ϣ����
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
'HMAC����Կ��صĹ�ϣ������Ϣ��֤�룬HMAC�������ù�ϣ�㷨����һ����Կ��һ����ϢΪ���룬����һ����ϢժҪ��Ϊ�����
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
'��ȡZLSM4���޸İ汾
'1:ֻ֧��sm4_crypt_ecb,sm4_crypt_cbc
'2:����֧��sm3��sm3_file��sm3_hmac��sm_version
'/**
' * \brief          Output = zlSM4.DLL Version
' */
Private Declare Function get_sm_version Lib "zlSm4.dll" Alias "sm_version" () As Long

Private Enum CrypeMode
    CM_Encrypt = 1   '����
    CM_Decrypt = 0   '����
End Enum
Private M_SM4_VERSION As Long
Public Const SM4_CRYPT_RANDOMIZE_KEY As Long = 999  'sm4�����㷨��Կ���������������
Public Const SM4_CRYPT_RANDOMIZE_IV As Long = 666   'sm4�����㷨��ʼ�������������������
Public Const G_PASSWORD_KEY             As String = "3357F1F2CA0341A5B75DBA7F35666715"
'======================================================================================================================
'����           Sm4EncryptEcb           SM4����
'����ֵ         String                  ���ܺ��ֵ,��ʽ��ZLSV+�汾��+:+���ܺ���ַ���
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ���
'strKey         String(Optional)        ������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
'����           Sm4DecryptEcb           SM4����
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ��������ַ�����Sm4EncryptEcb���ɵĽ����
'strKey         String(Optional)        ������ԿҲ���ǽ�����Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
        '��ǰ�ͻ��˵�ZLSM4��֧�ָð汾�ļ����ַ������ܣ��Ծɽ��ܣ���Ϊһ����˵���ܽ��ܳ���ͬ���ַ���
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
'����           Sm4EncryptCbc           SM4�������
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ���
'strKey         String(Optional)        ������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
'strIv          String(Optional)        ���������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
'����           Sm4EncryptCbc           SM4������ܶ�Ӧ�Ľ��ܹ���
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  �Ѿ����ܵ��ַ���
'strKey         String(Optional)        ������ԿҲ���Ǽ�����Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
'strIv          String(Optional)        ���������ԿҲ���Ƿ��������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
        '��ǰ�ͻ��˵�ZLSM4��֧�ָð汾�ļ����ַ������ܣ��Ծɽ��ܣ���Ϊһ����˵���ܽ��ܳ���ͬ���ַ���
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
'����           Sm3                     �����ַ����Ĺ�ϣֵ����������ַ����ı䶯��
'����ֵ         String(32)              �ַ����Ĺ�ϣֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  �ַ�������
'======================================================================================================================
Public Function Sm3(ByRef strInput As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte

    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrInput = StrConv(strInput, vbFromUnicode)
    lngLength = UBound(arrInput) + 1
    
    Call sm3_hash(arrInput(0), lngLength, arrOut(0))
    '������ֵת��Ϊ16�����ַ���
    Sm3 = ByteToHexString(arrOut)
End Function
'======================================================================================================================
'����           Sm3_File                �����ļ��Ĺ�ϣֵ��������� �ļ����ݵı䶯��
'����ֵ         String(32)              �ļ��Ĺ�ϣֵ
'����б�:
'������         ����                    ˵��
'strFile        String                  �ļ�·��
'======================================================================================================================
Public Function Sm3_File(ByRef strFile As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte
    Dim lngReturn As Long

    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrInput = StrConv(strFile, vbFromUnicode)
    '����APIû�д��ݳ��ȣ���������ַ���Chr(0)
    lngLength = UBound(arrInput) + 1
    ReDim Preserve arrInput(lngLength)
    '������
    lngReturn = sm3_file_hash(arrInput(0), arrOut(0))
    '�ж��Ƿ�ɹ�����
    If lngReturn = 0 Then
        '������ֵת��Ϊ16�����ַ���
        Sm3_File = ByteToHexString(arrOut)
    ElseIf lngReturn = 1 Then
        Sm3_File = "ERROR:�ļ���ʧ��"
    ElseIf lngReturn = 2 Then
        Sm3_File = "ERROR:�ļ���ȡʧ��"
    End If
End Function
'======================================================================================================================
'����           sm3_hmac                ������һ����Կ�Դ������Ϣ������ϢժҪ
'����ֵ         String(32)              ��Կ������Ϣ�����ɵ���ϢժҪ
'����б�:
'������         ����                    ˵��
'strKey         String                  ��Կ
'strMsg         String                  ��Ϣ����
'======================================================================================================================
Public Function sm3_hmac(ByRef strKey As String, ByVal strMsg As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte
    Dim arrKey()    As Byte
    Dim lngKeyLen   As Long
    
    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrInput = StrConv(strMsg, vbFromUnicode)
    lngLength = UBound(arrInput) + 1
    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrKey = StrConv(strKey, vbFromUnicode)
    lngKeyLen = UBound(arrKey) + 1
    Call sm3_hmac_hash(arrKey(0), lngKeyLen, arrInput(0), lngLength, arrOut(0))
    '������ֵת��Ϊ16�����ַ���
    sm3_hmac = ByteToHexString(arrOut)
End Function
'======================================================================================================================
'����           sm_version              ��ȡZLSM4�İ汾��
'����ֵ         Long                    ZLSM4�İ汾��
'����б�:
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
'����           ByteToHexString         ���ֽ���ת��Ϊ16�����ַ���
'����ֵ         String                  �ֽ���ת����16�����ַ���
'����б�:
'������         ����                    ˵��
'bytInpu        Byte(��                 �ֽ�����
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
'����           ByteToHexString         ��16�����ַ���ת��Ϊ�ֽ���
'����ֵ         Byte()                  16�����ַ���ת�����ֽ���
'����б�:
'������         ����                    ˵��
'bstrInput      String                  16�����ַ���
'lngRetBytLen   Long(Optional)          ָ�����ص��ֽ���ĳ���,0-��ԭʼ���ȷ��أ�<>0����ָ���ĳ��ȣ����㲹�루��0�������˽�ȡ
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
'����           BytePadding             ��ָ���ַ�������16�ֽڲ��룬
'����ֵ         Byte()                  �������ַ����ֽ���
'����б�:
'������         ����                    ˵��
'strInput       String                  �ַ���
'lngVersion     Long(Optional,2)        �ַ�������İ汾��ZLSM4.DLL�İ汾���Լ������㷨ǰ׺�еİ汾����1-�ո��룬>1:Chr(0)����
'lngPaddingNum  Long(Optional,16)        ������ֽ�����ȱʡ����16���Ʋ���
'======================================================================================================================
Public Function BytePadding(ByVal strInput As String, Optional ByVal lngVersion As Long = 2, Optional ByVal lngPaddingNum As Long = 16) As Byte()
    Dim arrReturn()     As Byte
    Dim lngLenBef       As Long
    Dim i               As Long
    Dim lngLenAft       As Long
    
    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrReturn = StrConv(strInput, vbFromUnicode)
    lngLenBef = UBound(arrReturn) + 1
    '�жϵõ�������ĳ��ȣ�������16�����������򲹿ո��:Chr(0)
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
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
