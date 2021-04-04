Attribute VB_Name = "mdlPublic"
Option Explicit
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63)     As String

Public Enum DataEnum
    responseText = 1
    responseBody = 2
End Enum

'������־ģ��
Private mobjFso As New FileSystemObject '�ļ�����
'�������÷���
'*****************************************************************************************************************************
'��14λʱ���ַ���ת��Ϊ���ڣ�YYYY-MM-DD HH:mm:ss
Public Function String14ToDate(ByVal strData As String, Optional ByRef strErr As String = "0") As String
    '��ȡʱ���
    Dim strTimeStamp As String
    If strData = "" Then
        If strErr = "0" Then
            MsgBoxEx "��Чʱ�䲻��Ϊ�գ�", vbExclamation, gstrSysName
        Else
            strErr = "��Чʱ�䲻��Ϊ�գ�"
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
                    MsgBoxEx "��ȡ��ʱ�䲻��һ�����ڣ�" & strTimeStamp, vbExclamation, gstrSysName
                Else
                    strErr = "��ȡ��ʱ�䲻��һ�����ڣ�" & strTimeStamp
                End If
                String14ToDate = ""
                Exit Function
            End If
    End If
    String14ToDate = strTimeStamp
End Function

'==========================================================
'| ģ �� �� | [BASE64]
'| ˵    �� | BASE64���뼰���볣�ýӿ�
'---------------------------------------------------------------------------����Begin����---------------------------------------------------------------------------------------
'==========================================================
Private Sub InitBase()
'����:��ʼ�� BASE64����
     Dim iPtr     As Integer
     For iPtr = 0 To 63
         psBase64Chr(iPtr) = Mid$(BASE64CHR, iPtr + 1, 1)
     Next
End Sub

Public Function SaveBase64ToFile(ByVal strType As String, ByVal strSN As String, ByVal str2Decode As String) As String
'����:����Base64ΪͼƬ�ļ�
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
'����:���ַ�������Base64���벢�����ַ���
     Dim sValue()             As Byte
     sValue = StrConv(str2Encode, vbFromUnicode)
     EncodeBase64String = EncodeBase64Byte(sValue)
End Function

Public Function EncodeBase64Byte(sValue() As Byte) As String
'����:��һ���ֽ��������Base64���룬�������ַ���
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

         '//ת���ַ���Ϊ���飬Ȼ��ת��Ϊ4��6λ(0-63)
         Bits6(1) = (Bits8(1) And &HFC) \ 4
         Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
         Bits6(3) = (Bits8(2) And &HF) * 4 + (Bits8(3) And &HC0) \ &H40
         Bits6(4) = Bits8(3) And &H3F

         '//���4�����ַ�
         For lPtr = 1 To lLen + 1
             sEncoded = sEncoded & psBase64Chr(Bits6(lPtr))
         Next
     Next

     '//����4λ����=���
     Select Case lLen + 1
         Case 2: sEncoded = sEncoded & "=="
         Case 3: sEncoded = sEncoded & "="
         Case 4:
     End Select

     EncodeBase64Byte = sEncoded
End Function

Public Function EncodFileToBase64String(strFileSource As String)
'���ܣ����ļ�����Base64���벢���ر�����Base64�ַ���
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
'���ܣ����ļ�����Base64���룬��������������ֱ��д��һ���ı��ļ���
     Dim fso As New FileSystemObject, _
         ts As TextStream
    
     Set ts = fso.CreateTextFile(strFileBase64Desti, True)
     ts.Write (EncodFileToBase64String(strFileSource))
     ts.Close
     Set ts = Nothing
     Set fso = Nothing
End Sub

Public Function DecodeBase64Byte(str2Decode As String) As Byte()
'���ܣ���һ������Base64���ַ����н��뵽Դ�ֽ�����
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
    
     '//��ȥ�س�
     str = Replace(str2Decode, vbCrLf, "")

     '//ÿ4���ַ�һ�飨4���ַ���ʾ3���֣�
     For lPtr = 1 To Len(str) Step 4
         iLen = 4
         For iCtr = 0 To 3
             '//�����ַ���BASE64�ַ����е�λ��
             iValue = InStr(1, BASE64CHR, Mid$(str, lPtr + iCtr, 1), vbBinaryCompare)
             Select Case iValue   'A~Za~z0~9+/
                 Case 1 To 64:
                     bits(iCtr + 1) = iValue - 1
                 Case 65          '=
                     iLen = iCtr

 Exit For
                     '//û�з���
                 Case 0: Exit Function
             End Select
         Next

         '//ת��4��6��������Ϊ3��8������
         bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
         bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
         bits(3) = (bits(3) And &H3) * &H40 + bits(4)

         '//�����������ʼλ��
         lFrom = lTo
         lTo = lTo + (iLen - 1) - 1
                
         '//���¶����������
         ReDim Preserve Output(0 To lTo)
        
         For iIndex = lFrom To lTo
             Output(iIndex) = bits(iIndex - lFrom + 1)
         Next

         lTo = lTo + 1
        
     Next
     DecodeBase64Byte = Output
End Function

Public Function DecodeBase64String(str2Decode As String) As String
'���ܣ���һ������Base64���ַ����н��뵽Դ�ַ���
     DecodeBase64String = StrConv(DecodeBase64Byte(str2Decode), vbUnicode)
End Function

Public Sub DecodeBase64StringToFile(strBase64 As String, strFilePath As String)
'����:��һ��Base64�ַ������룬��д��������ļ�
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
'����:��һ��Base64�����ļ����룬��д��������ļ�
     Dim fso As New Scripting.FileSystemObject
     Dim ts As TextStream

     If Not fso.FileExists(strBase64FilePath) Then Exit Sub

     Set ts = fso.OpenTextFile(strBase64FilePath)
     
     DecodeBase64StringToFile ts.ReadAll, strFilePath
End Sub
'==========================================================
'| ģ �� �� | [BASE64]
'| ˵    �� | BASE64���뼰���볣�ýӿ�
'---------------------------------------------------------------------------����End����---------------------------------------------------------------------------------------
'==========================================================

'==========================================================
'| ģ �� �� | XMLHTTP
'| ˵    �� | ���Inet�ؼ���ʵ������ͨѶ
'---------------------------------------------------------------------------����Begin����---------------------------------------------------------------------------------------
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

    '--------------------------------------��������
110 Select Case DataStic
    Case responseText
        '--------------------------------ֱ�ӷ����ַ���
112     DataS = xmlHttp.responseText
114     HttpGet = DataS
116 Case responseBody
        '--------------------------------ֱ�ӷ��ض�����
118     DataB = xmlHttp.responseBody
120     HttpGet = DataB
122 Case responseBody + responseText
        '------------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
124     DataS = BytesToStr(xmlHttp.responseBody)
126     HttpGet = DataS
128 Case Else
        '--------------------------------��Ч�ķ���
130     HttpGet = ""
    End Select

    '--------------------------------------�ͷſռ�
132 Set xmlHttp = Nothing

    Exit Function

errH:
134 HttpGet = ""
136 MsgBoxEx "HttpGetʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "�������"
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

    '-----------------------------��������
114 Select Case DataStic
    Case responseText
        '--------------------------------ֱ�ӷ����ַ���
116     DataS = xmlHttp.responseText
118     HttpPost = DataS
120 Case responseBody
        '--------------------------------ֱ�ӷ��ض�����
122     DataB = xmlHttp.responseBody
124     HttpPost = DataS
126 Case responseBody + responseText
        '---------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
128     DataS = BytesToStr(xmlHttp.responseBody)
130     HttpPost = DataS
132 Case Else
        '--------------------------------��Ч�ķ���
134     HttpPost = ""
    End Select

    '------------------------------------�ͷſռ�
136     Set xmlHttp = Nothing

    Exit Function

errH:
138     HttpPost = ""
140     MsgBoxEx "HttpPostʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "�������"
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
'| ģ �� �� | XMLHTTP
'| ˵    �� | ���Inet�ؼ���ʵ������ͨѶ
'-----------------------------------------------------------------------------����END����-------------------------------------------------------------------------
'==========================================================

'ɽ��ʡʱ�����ʽ���ظ�ʽ������
'�� ��Dec 30 01:12:53 2014 GMT�� ����Ϊ��2014-12-30 01:12:53��
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

'���֤����Ч��,����֤����Ч������
Public Function CheckValidaty(ByVal endDate As Date) As Integer
    '����CA���հ���֤����Ч�Խӿ�
    '-���: ֤����Ч��ֹ����
    '-���Σ���Ч����
    Dim dblAllSp    As Double
    Dim result      As Integer
    Dim datNow As Date
    datNow = gobjComLib.zlDatabase.Currentdate
    dblAllSp = CDbl(CDate(endDate)) - CDbl(datNow)
    result = Int(dblAllSp)
    CheckValidaty = result
End Function

Public Sub WriteLog(ByVal strLogTxt As String)
    'дһ����־������������лس�,���з����滻Ϊ<CR><LF>
    '��־�����ڵ�ǰĿ¼�µ�[Ӧ�ó�������]LogĿ¼�£��ļ���Ϊ����.txt,Ĭ�ϱ���7�����־��
    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '��־·�����ļ����������ļ���
    Dim strLogSaveDays As String '��־��������
    Dim dblFreeSpace As Double   'ʣ��ռ�
    Dim strDelOldFile As String  '�����ļ�
    Dim objFile As File
    
    If Dir(App.Path & "\����ǩ��������־*.log") = "" Then Exit Sub
    'ʼ�ձ�����־
    '2�����������־
    strLogSaveDays = "7"  '����7�����־
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\����ǩ��������־*.log")
    Do While strDelOldFile <> ""
        Set objFile = mobjFso.GetFile(strLogPath & "\" & strDelOldFile)
        If DateDiff("d", objFile.DateLastModified, Now) > Val(strLogSaveDays) Then
            mobjFso.DeleteFile strLogPath & "\" & strDelOldFile, True
        End If
        strDelOldFile = Dir
    Loop
    
    '3���ռ��Ƿ��㹻
    dblFreeSpace = GetFreeSpace(strLogPath)
    If dblFreeSpace >= 1024 And dblFreeSpace <= 10240 Then
        '�ռ䲻�㣬��д��־,����һ�������ļ�
        If Not mobjFso.FileExists(strLogPath & "\�ռ䲻��.txt") Then Call mobjFso.CreateTextFile(strLogPath & "\�ռ䲻��.txt", True)
        Exit Sub
    Else
        '��������ļ�
        If mobjFso.FileExists(strLogPath & "\�ռ䲻��.txt") Then Call mobjFso.DeleteFile(strLogPath & "\�ռ䲻��.txt", True)
    End If
    '4��д����־��
    strLogFile = strLogPath & "\����ǩ��������־" & Format(Now, "yyyyMMdd") & ".log"
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
    '��ȡʣ��ռ�
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
        gobjComLib.LogWrite "����ǩ��������־", "", strFunction, strLog
        If Err.Number = 438 Then        '���ݵͰ汾û����־�ӿ�
            WriteLog strFunction & vbCrLf & strLog
        End If
        Err.Clear: On Error GoTo 0
    Else
        Exit Function
    End If
End Function
