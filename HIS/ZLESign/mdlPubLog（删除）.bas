Attribute VB_Name = "mdlPubLog"
Option Explicit
'������־ģ��
Private mobjFso As New FileSystemObject '�ļ�����

Public Sub WriteLog(ByVal strLogTxt As String)
    'дһ����־������������лس�,���з����滻Ϊ<CR><LF>
    '��־�����ڵ�ǰĿ¼�µ�[Ӧ�ó�������]LogĿ¼�£��ļ���Ϊ����.txt,Ĭ�ϱ���7�����־��

    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '��־·�����ļ����������ļ���
    Dim strLogSaveDays As String '��־��������
    Dim dblFreeSpace As Double   'ʣ��ռ�
    Dim strDelOldFile As String  '�����ļ�
    Dim objFile As File
    
    'If Dir(App.Path & "\����*.log") = "" Then Exit Sub
    'ʼ�ձ�����־
    '2�����������־
    strLogSaveDays = "7"  '����7�����־
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\ҽ������*.log")
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
    strLogFile = strLogPath & "\ҽ������" & Format(Now, "yyyyMMdd") & ".log"
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

