Attribute VB_Name = "mdlPubFunc"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2018/12/25
'ģ��           mdlPubFunc
'˵��
'==================================================================================================
Private Const mstrCurModule     As String = "mdlPubFunc"           '��ǰģ������
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'���أ�retrieve���Ӳ���ϵͳ������������elapsed���ĺ�����
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'�ַ�����UTF-8����
Public Const CP_UTF8 = 65001
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
Public Const G_UA_PWD           As String = "FA74C8A530DE7E088B1ACA673DD6297D"
Public Const G_UA_KEY           As String = "0016FDE250354FA9A4BA45433DBCC35D"
Public Const G_INTERFACE_KEY    As String = "EBA1D9B8CCCB4FD0804672DEDB222CFB"
Public Const G_APP_KEY          As String = "FD304782E75C41FDB14CB7A92A8A0B97"
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
Public Const SM4_CRYPT_RANDOMIZE_KEY As Long = 999  'sm4�����㷨��Կ���������������
Public Const SM4_CRYPT_RANDOMIZE_IV As Long = 666   'sm4�����㷨��ʼ�������������������
Private M_SM4_VERSION As Long
Private Enum CrypeMode
    CM_Encrypt = 1   '����
    CM_Decrypt = 0   '����
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

'--------------------------------------------------------------------------------------------------
'����           IsDesinMode
'����           ȷ����ǰģʽΪ���ģʽ��Դ�뻷����
'����ֵ         Boolean
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
'����           InCollection
'����           ��鼯�����Ƿ����ĳԪ��
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'cllTest        Collection              Ҫ���ļ���
'strKey         String                  Ҫ����Key
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
'����           DisPlayOneValue
'����           չʾ����
'����ֵ         String
'����б�:
'������         ����                    ˵��
'valValue       Variant                 ����Ķ���
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
'����           StringToUTF8Bytes       ���ַ���ת��ΪUTF-8������ֽ�����
'����ֵ         Byte()                  16�����ַ���ת�����ֽ���
'����б�:
'������         ����                    ˵��
'strInput      String                  16�����ַ���
'-------------------------------------------------------------------------------------------------
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '�ȼ��������ֽ���
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    'Ȼ��ת��
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function

'--------------------------------------------------------------------------------------------------
'����           UTF8BytesToString       ��UTF-8������ֽ�����ת��Ϊ�ַ���
'����ֵ         String                  ת������ַ���
'����б�:
'������         ����                    ˵��
'bytInpu        Byte(��                 �ֽ�����
'-------------------------------------------------------------------------------------------------
Public Function UTF8BytesToString(bytInpu() As Byte) As String
    Dim lngBytesRequired As Long

    '�ȼ��������ֽ���
    lngBytesRequired = MultiByteToWideChar(CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal 0, 0)
     
    'Ȼ��ת��
    UTF8BytesToString = String(lngBytesRequired, 0)
    MultiByteToWideChar CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal StrPtr(UTF8BytesToString), lngBytesRequired
End Function

'-------------------------------------------------------------------------------------------------
'����           EncBase64Char           ��6-bit�ֽ�ת��ΪBase64�ַ�
'����ֵ         Byte                    �ַ���ֵ
'����б�:
'������         ����                    ˵��
'bytValue       Byte                    ת�����ֽ�
'����˵����Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
'-------------------------------------------------------------------------------------------------
Private Function EncBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue < 26 Then '26����дӢ����ĸ
        EncBase64Char = bytValue + &H41
    ElseIf bytValue < 52 Then '26��СдӢ����ĸ
        EncBase64Char = bytValue + &H61 - 26
    ElseIf bytValue < 62 Then '10������
        EncBase64Char = bytValue + &H30 - 52
    ElseIf bytValue = 62 Then
        EncBase64Char = &H2B '+
    Else
        EncBase64Char = &H2F '/
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           DecBase64Char           ��Base64�ַ�ת��Ϊ6 bit�ֽ�
'����ֵ         Byte                    �ַ���ֵ
'����б�:
'������         ����                    ˵��
'bytValue       Byte                    ��������ֽ�
'����˵����Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
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
'����           EncodeBase64            ����Base64���룬����Base64���ַ���
'����ֵ         String                  Base64������
'����б�:
'������         ����                    ˵��
'varInput       Variant                 ��Ҫ����Base64������ַ��������ֽ����飬�ַ�����ȡUTF-8���롣Byte()����ǰ������飬Ԫ�ظ�����3�ı��������һ�δ�������ʣ�µļ��ɡ�
'����˵����Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
'-------------------------------------------------------------------------------------------------
Public Function EncodeBase64(varInput As Variant) As String
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    
    On Error GoTo errH
    
    If VarType(varInput) = vbString Then
        If Len(varInput) = 0 Then Exit Function
        'ԭʼ����,�Ƚ�ԭ����UTF-8�ķ�ʽ����
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
    '��8-bit�ֽ�����ת��Ϊ6-bit�ֽ�����
    For i = 0 To lngInputLen - 1
        If lngBit = 0 Then 'bytOut(J)δ��д��
            bytOut(j) = (bytInput(i) And &HFC) \ &H4
            j = j + 1
            bytOut(j) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)�ѱ�д����λ
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HF0) \ &H10)
            j = j + 1
            bytOut(j) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)�ѱ�д����λ
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HC0) / &H40)
            j = j + 1
            bytOut(j) = bytInput(i) And &H3F
            j = j + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) 'ת��ΪBase64�ַ�
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") 'ԭ��ʣ�����ݲ���3���ֽ���Ҫ����
    Exit Function
errH:
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'����           DecodeBase64            ��Base64���ַ�������Ϊԭ�ġ�
'����ֵ         Variant                 ԭʼ�ַ�����ԭʼ���ֽ���
'����б�:
'������         ����                    ˵��
'strInput       String                  Base64�����ַ���
'blnByteArray   Boolean                 True:����Byte(),False-����string
'����˵����Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
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
        '����������
        lngModLen = Len(strInput) - lngModLen + 1
        bytInput = StrConv(strInput, vbFromUnicode)
    Else
        lngModLen = 0
        '����������
        bytInput = StrConv(strInput, vbFromUnicode)
    End If
    lngInputLen = UBound(bytInput) + 1
 
    'ԭʼ����
    lngOutLen = lngInputLen - lngInputLen \ 4
    lngOutLen = lngOutLen - lngModLen
    ReDim bytOut(lngOutLen - 1)
 
    For j = 0 To lngInputLen - 1
        bytInput(j) = DecBase64Char(bytInput(j)) '��Base64�ַ�ת��Ϊ6-bit�ֽ�
    Next
    '��6-bit�ֽ�����ת��Ϊ8-bit�ֽ�����
    For j = 0 To lngOutLen - 1
        If lngBit = 0 Then 'bytOut(J)δ��д��
            bytOut(j) = bytInput(i) * &H4
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H30) \ &H10)
            lngBit = 2
        ElseIf lngBit = 2 Then 'bytOut(J)�ѱ�д�����ֽ�
            bytOut(j) = (bytInput(i) And &HF) * &H10
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H3C) \ &H4)
            lngBit = 4
        ElseIf lngBit = 4 Then 'bytOut(J)�ѱ�д�����ֽ�
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
        '���ת���õ���UTF-8�ַ���ת��ΪVB֧�ֵ�Unicode�ַ����Ա�����ʾ��
        DecodeBase64 = UTF8BytesToString(bytOut)
    End If
    Exit Function
errH:
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'����           EncodeBase64_file       ���ļ�����Base64���룬����Base64���ַ���
'����ֵ         String                  Base64������
'����б�:
'������         ����                    ˵��
'strFile        String                  ��Ҫ����Base64������ļ�
'����˵����Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
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
                If lngCount = 0 Then '��ֹ��ͣ�����ڴ�
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
'����           DecodeBase64_File       ��Base64���ַ�������Ϊԭ�ġ�
'����ֵ         String                  ���ɵ��ļ���
'����б�:
'������         ����                    ˵��
'strInput       String                  Base64�����ַ���
'strFile        String                  ָ���ļ���
'����˵����Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
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
'����           Serialize               �������ֵ���л�Ϊ�ַ���
'����ֵ         String                  ���л����ַ���
'����б�:
'������         ����                    ˵��
'objInfo        Variant                 �����ֵ
'strKeyName     String                  ���л��Ĺؼ���
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
'����           UnSerialize             ���ַ��������л�Ϊ���������ֵ
'����ֵ         Variant                 ���л��ַ�����Ӧ�Ķ��������ֵ
'����б�:
'������         ����                    ˵��
'strSource      String                  ���л��ַ���
'strKeyName     String                  ���л��Ĺؼ���
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
'����           SerializeMulti          ��˳�����л������Ϣ
'����ֵ         String                  ���л����ַ���
'����б�:
'������         ����                    ˵��
'arrInfo        Variant                 ������л��Ķ���
'[      ]       long                    ��0��ʼ������������Ϊ���л��Ĺؼ���
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
'����           UnSerializeMulti        ��ȡ���еĶ���
'����ֵ         Variant                 ���л��Ķ�������
'����б�:
'������         ����                    ˵��
'strSource      String                  ���л��ַ���
'[      ]       long                    ��0��ʼ������������Ϊ���л��Ĺؼ���
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
'���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd[ HH:mm])
    Dim curDate As Date, strTmp As String
    
    If strText = "" Or Len(strText) <> 14 Then Exit Function
    strTmp = strText
    '��������yyyyMMddHHmm
    strTmp = Format(strTmp, "00000000000000")
    strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Right(strTmp, 2)
    FullDate = strTmp
End Function

Public Function CopyNewRec(Optional ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'���ܣ����Ƽ�¼�����߹���һ���Զ����¼��
'������strFields=��Ҫ���Ƶļ�¼�����ֶε���˳����ֶ�����ɵ��ַ���
'          �磺1 ����1,3 ����2,7 ����3...��ʾ���Ƽ�¼���ĵ�1,3,7..�ֶ���ɼ�¼��������
'              ID ����1,���� ����2,....��ʾ���Ƽ�¼����ID,����...�ֶ���ɼ�¼������
'              ����*Ϊ�µļ�¼��������
'              �������ͻ�����׳���������ͬ�����⣬��ע��
'              *,�ڱ�ʾ����ԭ��¼���������ֶε�ռλ����������Ҫ��ԭ�����ֶ�ȫ�����ƣ�ͬʱ���ӱ��������жϸı�
'           arrAppFields=׷�ӵ��ֶ���Ϣ������,����,����,Ĭ��ֵ,û��Ĭ��ֵ��Empty,û��ָ�����ȴ�Empty
'      blnOnlyStructure=�Ƿ�ֻ���ƽṹ��rsSource����ʱ����Ч��
'��ע��1���ڳ����У��������漰���໥���ݼ�¼������ʹ��ADO��Clone���Ʋ����ļ�¼����������һ����¼�������ݷ����仯��ʱ�����и�������������ͬ�ı仯��ͨ��ָ�޸Ļ�ɾ����������������ϣ����Щ��¼���໥�䱣�ֶ���
'      2)��ʱ������Ҫһ�ֱ����͵����ݽṹ���洢���ݣ��ú������Բ���һ���Զ����¼����ʵ��
'Ӧ�ó�����
'             1��CopyNewRec(rsSource����ȫ�����ƽṹ�Լ�����
'             2��CopyNewRec(rsSource,True����ֻ�����ṹ����������
'             3��CopyNewRec(rsSource,,"ID ����1,����")����ԭ��¼����ID�������е����ݣ��������¼�¼����Ϊ����1����������Ҫֻ���ƽṹ��blnOnlyStructure��True
'             4)CopyNewRec(rsSource,,"*,��־ �±�־")����ԭ��¼���������ֶΣ����������С��±�־������������Դ����־�С����������������жϲ������ݱ仯
'             5)CopyNewRec(rsSource,,,Array("�Ƿ�ı�", adInteger, 1, 0)����ȫ�����ƽṹ�Լ����ݣ�����һ�������Ƿ�ı�
'             5��CopyNewRec(Nothing, , , Array("ϵͳ���", adInteger, 5, Empty, "������", adVarChar, 100, Empty)) ����һ���Զ����¼��
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
        '������¼���ṹ
        If strFields = "" Then
            strFields = "*"
        End If
        arrFieldsTmp = Split(strFields, ",")
        arrFieldsName = Array()
        For intFields = LBound(arrFieldsTmp) To UBound(arrFieldsTmp)
            If Trim(arrFieldsTmp(intFields)) = "*" Then '��ʶ�˴�������ԭ��¼����������
                If Not rsClone Is Nothing Then
                    For i = 0 To rsClone.Fields.Count - 1
                        ReDim Preserve arrFieldsName(UBound(arrFieldsName) + 1)
                        arrFieldsName(UBound(arrFieldsName)) = rsClone.Fields(i).Name & ""
                        .Fields.Append rsClone.Fields(i).Name, IIf(rsClone.Fields(i).Type = adNumeric, adDouble, rsClone.Fields(i).Type), rsClone.Fields(i).DefinedSize, adFldIsNullable    '0:��ʾ����
                    Next
                End If
            Else
                ReDim Preserve arrFieldsName(UBound(arrFieldsName) + 1)
                '�а�������
                arrTmp = Split(arrFieldsTmp(intFields) & " ", " ")
                strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).Name & ""
                '��ȡ�ֶ�ԭ������������
                arrFieldsName(UBound(arrFieldsName)) = strFieldName
                '����ֶ�,�������ڱ������������е�����Ϊ����
                .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:��ʾ����
            End If
        Next
        
        '׷���ֶ����
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
        '��������
        If Not blnOnlyStructure Then
            If rsClone Is Nothing Then Set CopyNewRec = rsTarget: Exit Function
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '�¼�¼�����а�˳����ӣ���˿�������
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
'���ܣ�ɾ��ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'���أ��Ƿ�ɹ�
'      rsInput=����ɾ����ļ�¼��
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
'���ܣ�����ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'      arrInput=������ֶ����Լ�ֵ����ʽ���ֶ���1,ֵ1, �ֶ���2,ֵ2,....
'���أ��Ƿ�ɹ�
'      rsInput=�������º�ļ�¼��
'˵����arrInput���ֶ�ֵ�����ü�¼���е������ֶ������¸��ֶΣ���ʱ��ʽΪ��!�ֶ��� ������(��ʱ֧��Val)
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
'���ܣ���ָ����¼����������ӵ���һ����¼����
'������rsSource=����Դ��¼��
'      rsAppend=׷�ӵ����ݼ�¼��
'      lngAppendRows=��ӵ�������-1,��ʾȫ����ӣ�>=0��ʾ�������N��
'      strSourceFields��strAppendFileds=�ζ�Ӧ���򣬸ò�������ʱ��Ĭ������¼���ṹ��ͬ����ʽ��[��¼��1].�ֶ�1,�ֶ�2...��[��¼��2].�ֶ�1,�ֶ�2...,��Ϊ"-�ֶ�1,�ֶ�2"Ϊ�������ų���Щ�ֶκ�ʣ���һһ��Ӧ
'      blnKeepBookMark:������Ӻ��Ƿ񽫼�¼�ع�ԭλ��
'      arrOtherFieldValue:�����޶�Ӧ�ֶε�ֵ������ʽ�����ֶ�����,ֵ���ⲿ���ֶβ�����strSourceFields��strAppendFileds�Ķ�Ӧ������֣�������-�ֶη�ʽ���֣�
'���أ��Ƿ�ɹ�
'      rsSource=������ݺ�ļ�¼��
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
    
    '������в���Ӧ
    lngCount = UBound(arrSource)
    If lngCount > UBound(arrAppend) Then
        lngCount = UBound(arrAppend)
    End If
    '�����Զ����ֶε�ֵ����
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
'���ܣ���¼��ȥ�ظ�
'������rsSource=Ҫȥ�ظ��ļ�¼��
'strDisFieldsName=ȥ�ظ����ֶ�,Ϊ�գ���������ֶ�ȥ��
'strFieldsName=���ؽ�����ֶΣ�Ϊ�գ��򷵻�ȥ�ظ����ֶ�
'���أ�������ļ�¼��
    Dim rsReturn As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String, strOldRow As String

    '��ȡĬ���ֶ���
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
    
    rsReturn.Sort = strDisFieldsName '�����Զ�������ƶ�����ͷ
    Do While Not rsReturn.EOF
        strTmp = rsReturn.GetString(, 1, "[ColumnSpliter]", , "[NULLEXP]") '�Զ��ƶ����
        rsReturn.MovePrevious
        If strTmp = strOldRow Then  'ɾ���ظ���
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
'���ܣ����ı���ĵ��ı�ѡ��
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
'���ܣ��жϵ�ǰ�û��Ƿ�ΪDBA��ɫ
    Dim rsTemp As ADODB.Recordset
    Dim strSQL      As String
    
    On Error GoTo errH
    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = gobjRegister.OpenSQLRecord(connThis, strSQL, "�жϵ�ǰ�����û��Ƿ����DBA��ɫ")
    CheckIsDBA = rsTemp.RecordCount > 0
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single = -1, Optional strSQL As String, Optional strServer As String, Optional blnPer As Boolean)
'���ܣ���ʾ�����صȴ�����ȴ���(strInfo)
'����:strInfo=�ȴ��������ʾ��Ϣ
'     sngPer=����
    
    If gblnSilence Then Exit Sub
    If Not gblnShow Then
        If gblnCurShow Then
            ShowWindow frmFlash.hWnd, 0
            gblnCurShow = False
        End If
        Exit Sub
    End If
    
    If glngSec > 0 Then
        frmFlash.lblTip.Caption = (glngSec \ 10) & "�����Զ����ص�����������Ҫ�鿴���飬����������ͼ�ꡣ"
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
        frmFlash.lblServer = "��������" & strServer
        frmFlash.txtSQL = strSQL
        frmFlash.lbl.Caption = strInfo
        If Not gblnCurShow Then
            On Error Resume Next
            If sngPer = -1 Then
                '��ʾ�ȴ�
'                frmFlash.avi.Open GetSetting("ZLSOFT", "ע����Ϣ", "gstrAviPath", "") & "\" & "Findfile.avi"
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
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '���ܣ�����������ɾ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '�¼�����������
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
    
End Sub

