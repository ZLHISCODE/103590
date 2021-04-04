Attribute VB_Name = "mdlLog"
Option Explicit

'��
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'д
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
'����ֵ:�����ʾ�ɹ������ʾʧ�ܡ�������GetLastError

Private Const CON_SPLIT As String = ";"
Private mobjFso As New FileSystemObject         '�ļ�����

Public Function ReadIni(ByVal strNodeName As String, ByVal strKeyName As String, strFilePath As String) As String
    Dim strBuff As String
    Dim strReadStr As String
    Dim lngPos As Long
    
    On Error GoTo ErrH

    strBuff = VBA.String(255, 0)
    GetPrivateProfileString strNodeName, strKeyName, "", strBuff, 256, strFilePath
    strReadStr = VBA.Replace(strBuff, VBA.Chr(0), "")
    
    lngPos = InStr(1, strReadStr, CON_SPLIT, vbTextCompare)     '�ҵ� ;��λ��(������־)
    If lngPos >= 1 Then
        ReadIni = Trim(Left(strReadStr, lngPos - 1))
    Else
       '���û���ҵ� ��ע�͵ı�־
       ReadIni = strReadStr
    End If
    
    Exit Function
ErrH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(ByVal strNodeName As String, ByVal strKeyName As String, ByVal strValue As String, ByVal strFilePath As String) As Long
    Dim strBuff As String
    Dim strComment As String
    Dim strReadStr As String
    
    Dim lngRet As Long
    Dim lngPos As Long
    On Error GoTo ErrH
   strBuff = String(255, 0)
   lngRet = GetPrivateProfileString(strNodeName, strKeyName, "", strBuff, 256, strFilePath)
   strReadStr = VBA.Replace(strBuff, VBA.Chr(0), "")
   lngPos = InStr(1, strReadStr, CON_SPLIT, vbTextCompare)    '�ҵ� ;��λ��(������־)
   '�����;ȡ������ע��
   If lngPos >= 1 Then
      strComment = Trim(Right(strReadStr, lngRet - lngPos))
      strValue = strValue & strComment
   End If
   
    WriteIni = WritePrivateProfileString(strNodeName, strKeyName, strValue, strFilePath)
        
    Exit Function
ErrH:
    Err.Clear
End Function

Public Sub WriteLog(ByVal strLogTxt As String)
    'дһ����־������������лس�,���з����滻Ϊ<CR><LF>
    '��־�����ڵ�ǰĿ¼�µ�[Ӧ�ó�������]LogĿ¼�£��ļ���Ϊ����.txt,Ĭ�ϱ���7�����־��

    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '��־·�����ļ����������ļ���
    Dim strLogSaveDays As String '��־��������
    Dim dblFreeSpace As Double   'ʣ��ռ�
    Dim strDelOldFile As String  '�����ļ�
    Dim objFile As File
    
    '�Ƿ�����־
    If Not gblnLog Then Exit Sub
     
    'ʼ�ձ�����־
    '2�����������־
    strLogSaveDays = "7"  '����7�����־
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\" & App.EXEName & "*.log")
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
    strLogFile = strLogPath & "\" & App.EXEName & Format(Now, "yyyyMMdd") & ".log"
    Call SaveLog(strLogFile, strLogTxt)

End Sub

Private Sub SaveLog(ByVal strFilename As String, ByVal strInput As String, Optional ByVal strDate As String)
 
    Dim objStream As TextStream
    Dim strWritLing As String
    
    strWritLing = Replace$(strInput, Chr(&HD), "<CR>")
    strWritLing = Replace$(strInput, Chr(&HA), "<LF>")

    If strInput <> "" Then
        If Not mobjFso.FileExists(strFilename) Then Call mobjFso.CreateTextFile(strFilename)
        Set objStream = mobjFso.OpenTextFile(strFilename, ForAppending)
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
