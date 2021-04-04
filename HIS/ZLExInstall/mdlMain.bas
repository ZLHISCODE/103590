Attribute VB_Name = "mdlMain"
Option Explicit

'����ע������
Public Enum RegFileType
    RFT_NotReg = 0                  '��ע��Ķ���
    RFT_NormalReg = 1               '����ע�ᣬ�Զ�ʶ��.NET������.NET����ͨ��Regasmע�ᣬ����ͨ������DLLRegServerע��
    RFT_NETGAC = 2                  'NET����ע�ᣬͨ��gacutilע�ᵽȫ�ֳ��򼯻���
    RFT_NETServer = 3               'NET����ע�ᣬͨ��installUtil���а�װж�ء�
    RFT_NETComReg = 4               '.NET Com����ע�ᣬͨ������Regasm���
    RFT_VBComReg = 5                'ͨ����дע���ע��
    RFT_DelphiComReg = 6            'DelphiComע�ᣬͨ��DLLRegServerע��
    RFT_PBComReg = 7                'PBComע�ᣬͨ��DLLRegServerע��
End Enum

Public gobjFSO              As New FileSystemObject     '�ļ���������
Public gobjTrace            As New clsTrace             '��־���ٶ���
Public gstrGACPath          As String                   'GACUTIL.EXE·��
Public gstr7ZPath           As String                   '7z.exe�ļ�·��
Public gblnIs64Bits         As Boolean                  '�Ƿ���64λϵͳ
Public gclsRegCom           As New clsRegCom            '����ע�����
Public gstrAPPPath          As String
Public gstSysPath           As String
Public gobj7z               As New cls7zZip

Sub Main()
    Dim strErr As String
    Call InitInstall
    If Not InstallOO4O(strErr) Then
        MsgBox "OO4O�����װʧ�ܡ���Ϣ��" & strErr, vbInformation, "�������"
    Else
        MsgBox "OO4O�����װ�ɹ���", vbInformation, "�������"
    End If
End Sub

Private Function InitInstall() As Boolean
    '��װ���Ƿ����
    If IsDesinMode Then
        gstrAPPPath = "C:\APPSOFT"
    Else
        gstrAPPPath = gobjFSO.GetParentFolderName(App.Path)
    End If
    gstrGACPath = gstrAPPPath & "\Public\gacutil.exe"
    gblnIs64Bits = Is64bit
    gstSysPath = gobjFSO.GetSpecialFolder(SystemFolder)
    If gblnIs64Bits Then
        gstSysPath = gobjFSO.GetParentFolderName(gstSysPath) & "\SysWOW64"
    End If
    gstr7ZPath = gstSysPath & "7z.exe"
    gobj7z.Init7zZip (gstr7ZPath)
    Call gobjTrace.OpenTace("OO4O", gstrAPPPath)
End Function

