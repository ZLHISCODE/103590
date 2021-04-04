Attribute VB_Name = "mdlCMDControl"
Option Explicit

'**************************
'����:�ܵ�ִ��CMD�����ȡ���
'��д����:ף��
'**************************

Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const STARTF_USESTDHANDLES = &H100&
Public Const STARTF_USESHOWWINDOW = &H1


Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

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
