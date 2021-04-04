Attribute VB_Name = "mdlReceiveSend"
Option Explicit

Public gstrBuffFile As String  '�������ݵĴ���ļ���
Public gstrIniFile As String   '�����ļ���
Public gstrLockFile As String  '�����ļ���
Public gstrRAWDIR As String    'ԭʼ����Ŀ¼
Public gstrResultDIR As String '����������Ŀ¼
Public gstrSendDir As String   '������ָ����Ŀ¼
Public gstrGamDir As String    'ͼ�������Ŀ¼

Public gFileObject As New FileSystemObject  '�����ļ�ϵͳ���������ļ�Ŀ¼��ز���
Public gobjLisDev As Object                 '�����ͨѶ����

Public Type T��������
    
    ����       As Integer  '0-COM�ڷ�ʽ 1-IP��ʽ
    'Com
    COM�˿�       As Integer
    ������     As Long
    ����λ     As String
    У��λ     As String
    ֹͣλ     As String
    ����       As String
    �����С   As Long
    
    'TCP/TP
    IP�˿�     As Long
    IP         As String
    ����       As Long
    
    '����
    �ַ�ģʽ   As String
    �Զ�Ӧ��   As String   '�Զ�Ӧ��������λ�룬Ϊ<=0ʱ�����á�
    ͨѶ����   As String
    ͨѶ����   As String
End Type
Public g�������� As T��������   '��������ͨѶ����

'��дini �ļ���API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Sub Main()
    If Init = True Then
        frmMain.Show
    End If
End Sub

Public Sub ReadSet()
    '��ȡINI�ļ������ͨѶ����
    g��������.���� = CInt(Val(ReadIni("RECEIVE_SET", "����", gstrIniFile)))
    g��������.COM�˿� = CInt(Val(ReadIni("RECEIVE_SET", "COM�˿�", gstrIniFile)))
    g��������.IP = ReadIni("RECEIVE_SET", "IP", gstrIniFile)
    g��������.IP�˿� = CLng(Val(ReadIni("RECEIVE_SET", "IP�˿�", gstrIniFile)))
    g��������.������ = CLng(Val(ReadIni("RECEIVE_SET", "������", gstrIniFile)))
    g��������.�����С = CLng(Val(ReadIni("RECEIVE_SET", "�����С", gstrIniFile)))
    g��������.����λ = ReadIni("RECEIVE_SET", "����λ", gstrIniFile)
    g��������.ֹͣλ = ReadIni("RECEIVE_SET", "ֹͣλ", gstrIniFile)
    g��������.���� = ReadIni("RECEIVE_SET", "����", gstrIniFile)
    g��������.У��λ = ReadIni("RECEIVE_SET", "У��λ", gstrIniFile)
    g��������.���� = CLng(Val(ReadIni("RECEIVE_SET", "����", gstrIniFile)))
    g��������.�Զ�Ӧ�� = ReadIni("RECEIVE_SET", "�Զ�Ӧ��", gstrIniFile)  'Ӧ���ַ��ӽӿ���ȡ
    g��������.�ַ�ģʽ = ReadIni("RECEIVE_SET", "�ַ�ģʽ", gstrIniFile)
    g��������.ͨѶ���� = ReadIni("RECEIVE_SET", "ͨѶ����", gstrIniFile)
    
    g��������.ͨѶ���� = Val(ReadIni("RECEIVE_SET", "ͨѶ����", gstrIniFile))
    If Not (Val(g��������.ͨѶ����) > 0.1 And Val(g��������.ͨѶ����) < 600) Then g��������.ͨѶ���� = 0.5

End Sub

Private Function Init() As Boolean
    Dim strPath As String
    
    On Error GoTo errH
    
    gstrIniFile = App.Path & "\ReceiveSend.ini"
    If Not gFileObject.FileExists(gstrIniFile) Then
        MsgBox "��ͨѶ�����ļ���" & gstrIniFile & "�������������У�", vbQuestion, "ͨѶ����"
        Exit Function
    Else
        '����ͨѶ�����ļ��������ظ�����
        Dim TsTmp As TextStream
        
        gstrLockFile = App.Path & "\Lock.txt"

        If gFileObject.FileExists(gstrLockFile) And App.PrevInstance = True Then
            If Dir(gstrSendDir & "\CloseEnd.txt") = "" Then
                MsgBox "�������ظ����У�", vbQuestion, "ͨѶ����"
                Exit Function
            End If
        Else
            Set TsTmp = gFileObject.CreateTextFile(gstrLockFile, True)
            TsTmp.WriteLine "������" & Format(Now, "yyyy-MM-dd HH:mm:ss")
            TsTmp.Close
            Set TsTmp = Nothing
        End If
        If gFileObject.FileExists(gstrSendDir & "\CloseEnd.txt") Then gFileObject.DeleteFile gstrSendDir & "\CloseEnd.txt"
    End If
    
    '�������Ŀ¼
    '    RAW-ԭʼ����,Result-������,Gam-ͼ����,Send-��������
    gstrRAWDIR = App.Path & "\Raw"
    If Not gFileObject.FolderExists(gstrRAWDIR) Then Call gFileObject.CreateFolder(gstrRAWDIR)
    
    gstrResultDIR = App.Path & "\Result"
    If Not gFileObject.FolderExists(gstrResultDIR) Then Call gFileObject.CreateFolder(gstrResultDIR)
    
    gstrGamDir = App.Path & "\Gam"
    If Not gFileObject.FolderExists(gstrGamDir) Then Call gFileObject.CreateFolder(gstrGamDir)
    
    gstrSendDir = App.Path & "\Send"
    If Not gFileObject.FolderExists(gstrSendDir) Then Call gFileObject.CreateFolder(gstrSendDir)
    
    Init = True
    Exit Function
errH:
    MsgBox "��ʼ������ʱ���ִ���" & vbNewLine & Err.Description, vbQuestion, "ͨѶ����"
    
End Function

Public Function ReadIni(strItem As String, strKey As String, strPath As String) As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = Replace(GetStr, Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
    On Error GoTo errH
    WriteIni = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    WriteIni = False
End Function



Public Sub WriteErrLog(ByVal strFunc As String, ByVal StrInput As String, ByVal strOutput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------
    
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFilename As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
'    If Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv", "��ս�����־", 1)) = 1 Then
'        If Dir(App.Path & "\����.TXT") = "" Then Exit Sub
'    End If
    strFilename = App.Path & "\������־_" & Format(Date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strFilename) Then Call objFileSystem.CreateTextFile(strFilename)
    Set objStream = objFileSystem.OpenTextFile(strFilename, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (String(50, "��"))
    objStream.WriteLine ("ִ��ʱ��:" & strDate & "�汾:" & App.Major & "." & App.Minor & "." & App.Revision)
    objStream.WriteLine ("����:" & strFunc)
    objStream.WriteLine ("  :" & StrInput)
    objStream.WriteLine ("  :" & strOutput)
    'objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
End Sub

