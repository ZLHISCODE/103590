Attribute VB_Name = "mdlMain"
Option Explicit

Public Const gstrRegPath As String = "����ģ��\zlXWInterface\"   'ע���洢·��
Public Const gstrSysName As String = "Ӱ����Ϣϵͳ�ӿ�"

Public gcnOracle As ADODB.Connection            '�������ݿ�����
Public gzlComLib As Object                      '�������ݿ⴦��ģ��zlComLib

Public glngSys As Long                          'ϵͳ��
Public glngModule As Long                       'ģ���
Public gstrDBUser As String                     '��ǰ���ݿ��û�
Public gblnBefore3510 As Boolean                '����10.35.10ǰ��汾��True=10.35.10֮ǰ�汾,��ʹ��zlRegister����ʼ��comlibʱ��ҪSetDbUser��RegCheck

Public gstrLogPath As String
Public gstrBackupPath As String
Public gblnUseInterface As Boolean

Public mfrmShowHisForms As frmShowHisForms      '����������壬��������Ϣ

Public mclsArchive As clsArchive                '���Ӳ���������
Public mobjLisInsideComm As Object              'LIS�ӿڲ���

Public Const HIS_CAPTION = "������ʾHIS����"

Public Sub Main()
'------------------------------------------------
'���ܣ������򣬸����������Ӳ����鿴����
'������
'���أ���
'-----------------------------------------------
    Dim strRegPath As String
    Dim strMsgs As String
    Dim blnLis As Boolean
    Dim strTag As String
    Dim strOld As String
On Error GoTo ErrorHand

    
    strMsgs = Command
    
    strOld = strMsgs
    C_LOG = 1
    writeTestLog "��δ���" & strOld
    C_LOG = 0
    
    If Trim(strMsgs) = "" Then Exit Sub
    '����������ȷʵ�Ƿ�����־���ؼ��ֶ�::LOG=1::
    strTag = "::LOG=1::"
    If InStr(strMsgs, strTag) > 0 Then
        C_LOG = 1
        strMsgs = Replace(strMsgs, strTag, "")
    Else
        C_LOG = 0
    End If
    
    
    
    '��ҳ��ʱȥ�����ַ�
    If InStr(strMsgs, "://") > 0 Then
        strMsgs = Split(strMsgs, "://")(1)
    End If
    If InStr(strMsgs, "/") > 0 Then
        strMsgs = Split(strMsgs, "/")(0)
    End If
    If Trim(strMsgs) = "" Then Exit Sub
    
    '����������Ѿ�������һ�Σ�����������ֱ��ˢ�½������ݺ��˳�
    If App.PrevInstance Then
        If SendMsg(strMsgs) Then
            Exit Sub
        End If
    Else

    End If
    
    '���յ�QUIT��ֱ���˳�
    If UCase(Trim(strMsgs)) = "QUIT" Then Exit Sub
    
    '���ݴ���Ĳ����жϴ����ĸ���������Ϣ��ʽ���������:���ݿ��û���:����ID:����ID:ִ�в���ID:ҽ��ID��
    '����ID �� ����Ϊ�Һ�ID ���˹Һż�¼.ID��סԺΪ��ҳID
    
    '��ʼ��comlib�����ݿ�����
    If UBound(Split(strMsgs, MSG_SPLIT)) = 5 Then
        gstrZLHIS�����ַ��� = Split(strMsgs, MSG_SPLIT)(0)
        gstr�û��� = Split(strMsgs, MSG_SPLIT)(1)
        gstr���� = Split(strMsgs, MSG_SPLIT)(2)
        gbln�Ƿ�ת������ = Val(Split(strMsgs, MSG_SPLIT)(3)) = 1
    ElseIf UBound(Split(strMsgs, MSG_SPLIT)) = 6 Then
        '����LIS����
        blnLis = Val(Split(strMsgs, MSG_SPLIT)(0)) = 25
        gstrZLHIS�����ַ��� = Split(strMsgs, MSG_SPLIT)(1)
        gstr�û��� = Split(strMsgs, MSG_SPLIT)(2)
        gstr���� = Split(strMsgs, MSG_SPLIT)(3)
        gbln�Ƿ�ת������ = Val(Split(strMsgs, MSG_SPLIT)(4)) = 1
    Else
        Exit Sub
    End If
    Call InitInterface(Split(strMsgs, MSG_SPLIT)(1))
    
    '��ʼ��ϵͳ����
    If Not blnLis Then Call InitSysParameter
    
    '������Ϣ�����壬������Ϣhook��Ȼ�����ش���
    If mfrmShowHisForms Is Nothing Then Set mfrmShowHisForms = New frmShowHisForms
    Call mfrmShowHisForms.ShowMe(True)
    mfrmShowHisForms.Hide
    
    '������Ϣ
    Call ProcessMessage(strMsgs)
    
    Exit Sub
ErrorHand:
    If errHandle("exe Main", "��ʾ�������Ĵ��ڳ��ִ���") = 1 Then Resume
End Sub

'����Ϣ���͸���Ϣѭ��������
Private Function SendMsg(ByVal strmsg As String) As Boolean
    Dim lngWinHandle As Long        '��Ҫ������Ϣ�ġ�zlSoftShowHisForms.exe������Ĵ��ھ��
    Dim wParam As Long
    Dim lResult As Long
    Dim strTemp As String
    Dim buf(1 To 1024) As Byte
    
    '��Ϣ���壺wParam = 223��dss��dwData = 33 ������Ϣ��dwData = 32 �˳�
    wParam = 223
   
    Call CopyMemory(buf(1), ByVal strmsg, LenB(StrConv(strmsg, vbFromUnicode)))
    
    'dss.dwData�����Ϣ���ã�ֻ��˫�������һ����Ƕ���
    If UCase(Trim(strmsg)) = "QUIT" Then
        dss.dwData = 32 '���Ϊ�ر����д���
    Else
        dss.dwData = 33 '���Ϊˢ�´��ڻ��ߴ��´���
    End If
    
    dss.cbData = LenB(StrConv(strmsg, vbFromUnicode)) + 1
    
    'ʹ��buf���ͣ����Կ�����Ϣ��1024֮��
    dss.lpData = VarPtr(buf(1))
    
    '������Ϣѭ��������
    lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
    

    If lngWinHandle <> 0 Then
        lResult = SendMessage(lngWinHandle, WM_COPYDATA, wParam, dss)
        SendMsg = True
    End If
End Function


