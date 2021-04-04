Attribute VB_Name = "mdlPublic"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2019/1/21
'ģ��           mdlPublic
'˵��
'==================================================================================================
Private mobjFSO        As New FileSystemObject
'�ַ�����UTF-8����
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long

Public Function IsDesinMode() As Boolean
'���ܣ� ȷ����ǰģʽΪ���ģʽ
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
    Dim i           As Long, J              As Long, lngBit     As Long
    
    On Error GoTo ErrH
    
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
            bytOut(J) = (bytInput(i) And &HFC) \ &H4
            J = J + 1
            bytOut(J) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)�ѱ�д����λ
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &HF0) \ &H10)
            J = J + 1
            bytOut(J) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)�ѱ�д����λ
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &HC0) / &H40)
            J = J + 1
            bytOut(J) = bytInput(i) And &H3F
            J = J + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) 'ת��ΪBase64�ַ�
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") 'ԭ��ʣ�����ݲ���3���ֽ���Ҫ����
    Exit Function
ErrH:
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
    Dim i           As Long, J              As Long, lngBit     As Long
    Dim lngModLen       As Long
    On Error GoTo ErrH
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
 
    For J = 0 To lngInputLen - 1
        bytInput(J) = DecBase64Char(bytInput(J)) '��Base64�ַ�ת��Ϊ6-bit�ֽ�
    Next
    '��6-bit�ֽ�����ת��Ϊ8-bit�ֽ�����
    For J = 0 To lngOutLen - 1
        If lngBit = 0 Then 'bytOut(J)δ��д��
            bytOut(J) = bytInput(i) * &H4
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &H30) \ &H10)
            lngBit = 2
        ElseIf lngBit = 2 Then 'bytOut(J)�ѱ�д�����ֽ�
            bytOut(J) = (bytInput(i) And &HF) * &H10
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &H3C) \ &H4)
            lngBit = 4
        ElseIf lngBit = 4 Then 'bytOut(J)�ѱ�д�����ֽ�
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
        '���ת���õ���UTF-8�ַ���ת��ΪVB֧�ֵ�Unicode�ַ����Ա�����ʾ��
        DecodeBase64 = UTF8BytesToString(bytOut)
    End If
    Exit Function
ErrH:
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
ErrH:
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

    objBag.WriteProperty strKeyName, objInfo
    If Err.Number = 330 Then
        '�Ƿ�������  ��Ϊ��֧�ֳ־��Բ���д����
        Serialize = "{NotPersistable}"
        Err.Clear
    Else
        bytData = objBag.Contents
        Serialize = EncodeBase64(bytData())
    End If
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
                '�Ƿ�������  ��Ϊ��֧�ֳ־��Բ���д����
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
                '�Ƿ�������  ��Ϊ��֧�ֳ־��Բ���д����
                Err.Clear
                objBag.WriteProperty "K" & i, Nothing
            End If
        Next
    End If
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
'       ʵ�����ݴ洢����
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'������strText=�����ַ�
'         blnCrlf=�Ƿ�ȥ�����з�
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
'����           GetTickCountDiff
'����           ����GetTickCcout�Ĳ�ֵ������ GetTickCountVB�������ֵ�Լ��������������Ҫ��������
'����ֵ         Double
'����б�:
'������         ����                    ˵��
'lngStart       Long                    ��ʼʱ��
'lngEnd         Long                    ����ʱ��ȱʡ����
'blnInputEnd    Boolean                 ��ʶ�Ƿ�����lngEnd
'-------------------------------------------------------------------------------------------------
Public Function GetTickCountDiff(ByVal lngStart As Long, Optional ByVal lngEnd As Long, Optional ByVal blnInputEnd As Boolean) As Double
    Dim lngCur          As Long
    Const M_OFFSET_4    As Double = 4294967296#         '�޷������ε����ֵ
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
    Const M_OFFSET_4    As Double = 4294967296#         '�޷������ε����ֵ
    If value < 0 Then LongToUnsigned = value + M_OFFSET_4 Else LongToUnsigned = value
End Function

Public Function To_Date(ByVal strDate As String, Optional ByVal strType As String = "YMDHMS") As String
'���ܣ���ȡORACLE Date���ʹ�
'������strDate=ʱ���ַ���
'         strType=��ʽ�ַ������ͣ�ymd-�����գ�yyyy-mm-dd)��ymdhm-��yyyy-mm-dd hh:mm),ymdhms-��yyyy-mm-dd hh:mm:ss)
'���أ�ORACLE Date���ʹ�
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
'���ܣ�����VB���֧�ֵİ汾����ʽ:9999.9999.9999.9999,��С�汾��0000.0000.0000.0000
'������strVer=��ǰ�汾��
'           blnMax=True,����Ϊ�գ��򷵻����֧�ְ汾��False=����Ϊ�գ��򷵻���С֧�ְ汾
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '����һ�Σ��Լ�������SP�汾��
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

Public Function IsVerSion(ByVal strVer As String, Optional ByVal blnOnlyCheckSpecial As Boolean) As Boolean
'���ܣ��ж��ַ����Ƿ��ǰ汾��
'blnOnlyCheckSpecial=�Ƿ�������SP�汾��
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

'������ܳ���
Public Function Cipher(ByVal strText As String) As String
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    Dim strSeed As String
    
    If strText = "" Then Exit Function
    '��ȡ�������
    '������ӵ������Ϊ999
    Rnd (-1)
    Randomize (999)
    strSeed = "456"
    intSeedLen = Len(strSeed)
    strDeText = Chr(intSeedLen + MIN_ASC)
    For i = 1 To intSeedLen
        intChr = Asc(Mid(strSeed, i, 1)) 'ȡ��ĸת���ASCII��
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
        intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
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
'������ܳ���
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '������ӳ���
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '���þɵ�����㷨
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
        '��ȡ�������
        '������ӵ������Ϊ999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
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
        
    '���ݽ��ܵ�����
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intlen
        intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
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
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
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
''*  ����ͷ��ַ = GetSafeArrayInfo(���� ,[ά��],[�±�],[�ϱ�],[Ԫ�ظ���],[����],[Ԫ�س���],[��������],[��Ԫ�ص�ַ])
''*  result:  �ɹ���������ͷ��ַ;  ���� 0 ��������һ��δ����ʼ����������
''*  note:    ����1��������,��Ϊ����Ͳ���,Ҫ��ȡ�����ʲô��Ϣ,������Ӧ�ı���,ִ�к�,������ֵ��Ϊ��Ӧ����������
''*           Ҫ��ȡ��ά����ķǵ�1ά����/�±�,Ӧ��������ķ�������
''*           GetSafeArrayInfo (����, ָ��һ��ά��,[�±�],[�ϱ�]
''**************************************************************
'
'    Dim lngArrayHeaderAddress       As Long
'    Dim intDimCount                 As Integer     'ά���ƴ�
'
'    CopyMemory lngArrayHeaderAddress, ByVal VarPtr(varAnyArray) + 8, 4
'    CopyMemory lngArrayHeaderAddress, ByVal lngArrayHeaderAddress, 4                                            '��ȡ����ͷ��ַ
'    If lngArrayHeaderAddress < 1 Then Exit Function
'    CopyMemory intDimCount, ByVal lngArrayHeaderAddress, 2                                                      '��ȡ����ά��
'    If intGetDimension > intDimCount Then Exit Function                                                         '��ָ����ά������ʵ��ά�����˳�
'    CopyMemory lngLowerBound, ByVal (lngArrayHeaderAddress + 16 + (intDimCount - intGetDimension) * 8) + 4, 4   '��ȡ�±�
'    CopyMemory lngElements, ByVal (lngArrayHeaderAddress + 16 + (intDimCount - intGetDimension) * 8), 4         '��ȡָ��ά���µ�Ԫ�ظ���
'
'    lngUpperBound = lngElements + lngLowerBound - 1                             '��ȡָ��ά���µ��ϱ�
'    CopyMemory intFlags, ByVal lngArrayHeaderAddress + 2, 2                     '��ȡ��������
'    CopyMemory lngcbElements, ByVal lngArrayHeaderAddress + 4, 2                '��ȡ���鵥��Ԫ�س���
'    CopyMemory lngcLocks, ByVal lngArrayHeaderAddress + 8, 2                      '��ȡ������������
'    CopyMemory lngAddressOfData, ByVal lngArrayHeaderAddress + 12, 2              '��ȡ������Ԫ�ص�ַ
'    GetSafeArrayInfo = lngArrayHeaderAddress
'End Function
