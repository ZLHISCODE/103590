Attribute VB_Name = "mdlPublic"
Option Explicit

Public gobjFile As New FileSystemObject
Public gstrSysName As String
Public Const CON_STR_HINT_TITLE As String = "��ʾ"


Public Enum emResult
    etUndetected = 0    'δУ�Ի�δУ�Գ���
    etFileMiss = 1      '�ļ�ȱʧ
    etFileNull = 2      '�ļ���СΪ0
    etReadError = 3     '��ȡ�쳣
    etRoadError = 4     '·������
    etSucceed = 5       'У�Գɹ�
End Enum

Public Enum TMediaType
    imgTag = 0   'ͼ����
    MULFRAMETAG = 1 '����ͼ
    VIDEOTAG = 2 '��Ƶ���
    AUDIOTAG = 3 '��Ƶ���
End Enum
 


Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'��ȡ��������
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
    
    getEncryptionPassW = strBase & Join(strTemp, "") & strRandom '���ܺ���ִ�
End Function

'��ȡ��������
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

    getDecryptionPassW = Join(strTemp, "") '���ܺ���ִ�
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
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
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
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
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
'��ȡ��ԴĿ¼
    GetResourceDir = GetAppPath & "\..\�����ļ�\"
End Function

Public Function GetCacheDir() As String
'��ȡ����Ŀ¼
    GetCacheDir = GetAppPath & "\TmpImage\"
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next

    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir

    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��,����ָ�������ڡ���š�ϵͳ�ж�ָ�����ڵ������Ƿ���ת���������ݱ���
'������vDate=ʱ����ʱ��εĿ�ʼʱ��

    MovedByDate = gobjComlib.zlDatabase.DateMoved(CStr(vDate), 1, 100)
    
End Function


Public Function Unicode8Decode(bTemp() As Byte) As String
'����UNICODE UTF-8
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
'����UNICODE UTF-8
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
    Dim Code6 As Long  '�����ɵ��ֽ���
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
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If Err <> 0 Then
        MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, "��ʾ"
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ReadViewImage(ByVal strFile As String, Optional ByRef dcmViewer As DicomViewer = Nothing) As DicomImage
On Error GoTo errHandle
    Dim dImgs As DicomImages
        
    '�������_copy_vdat_��˵������ʱ�ļ�
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
        'readurl��֧�ֿո�
        Set curImage = dcmImgs.ReadURL(strFile)
    Else
        Set curImage = dcmImgs.ReadFile(strFile)
    End If
    
    If Err.Number = 0 Then
    
        'dcmImgs ������������
        If curImage.Picture Is Nothing Then
            Set ReadDicomFile = Nothing
            Exit Function
        End If
    
        Set ReadDicomFile = curImage
        Exit Function
    End If
    
    '2098����һ�����ļ�����dicom�ļ�����һ���Ǵ��ڹ�����ʴ���
    If InStr(Err.Description, "sharing violation") > 0 Then
        Err.Clear
        strFileTime = Format(Now, "YYMMDD") & GetTickCount
        
        Call FileCopy(strFile, strFile & "_copy_vdat_" & strFileTime)
    
        If blnUseUrl Then
            'readurl��֧�ֿո�
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
