Attribute VB_Name = "mdlCMDControl"
Option Explicit

'**************************
'����:�ܵ�ִ��CMD�����ȡ���
'��д����:ף��
'**************************

Function GetCmdTxt(Command As String, Optional opo As Boolean)
    Dim Proc As PROCESS_INFORMATION '������Ϣ
    Dim Start As STARTUPINFO '������Ϣ
    Dim SecAttr As SECURITY_ATTRIBUTES '��ȫ����
    Dim hReadPipe As Long '��ȡ�ܵ����
    Dim hWritePipe As Long 'д��ܵ����
    Dim lngBytesRead As Long '�������ݵ��ֽ���
    Dim strBuffer As String * 256 '��ȡ�ܵ����ַ���buffer
    Dim i As Integer
    On Error Resume Next
    Dim ret As Long 'API��������ֵ
    Dim retPro As Long
    Dim lpOutputs As String '���������ս��
    
    '���ð�ȫ����
    With SecAttr
        .nLength = LenB(SecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With

    '�����ܵ�
    ret = CreatePipe(hReadPipe, hWritePipe, SecAttr, 0)
    If ret = 0 Then
        MsgBox "�޷������ܵ�", vbExclamation, "����"
        Exit Function
    End If

    '���ý�������ǰ����Ϣ
    With Start
        .cb = LenB(Start)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .hStdOutput = hWritePipe '��������ܵ�
        .hStdError = hWritePipe '���ô���ܵ�
    End With

    '��������
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS������ipconfig.exeΪ��
    'Command = "Rasdial adsl CD02887573165 87573165"
    
    retPro = CreateProcess(vbNullString, Command, SecAttr, SecAttr, True, NORMAL_PRIORITY_CLASS, ByVal 0, vbNullString, Start, Proc)
    If ret = 0 Then
        MsgBox "�޷������½���", vbExclamation, "����"
        ret = CloseHandle(hWritePipe)
        ret = CloseHandle(hReadPipe)
        Exit Function
    End If
    
    '��Ϊ����д�����ݣ������ȹر�д��ܵ��������������رմ˹ܵ��������޷���ȡ����
    ret = CloseHandle(hWritePipe)

    '������ܵ���ȡ���ݣ�ÿ������ȡ256�ֽ�
    Do
        ret = ReadFile(hReadPipe, strBuffer, 256, lngBytesRead, ByVal 0)
        lpOutputs = lpOutputs & Left(strBuffer, lngBytesRead)
        DoEvents
    Loop While (ret <> 0) '��ret=0ʱ˵��ReadFileִ��ʧ�ܣ��Ѿ�û�����ݿɶ���

    '��ȡ������ɣ��رո����
    ret = CloseHandle(retPro)
    ret = CloseHandle(Proc.hProcess)
    ret = CloseHandle(Proc.hThread)
    ret = CloseHandle(hReadPipe)
End Function
