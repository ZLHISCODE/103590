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
Public mclsOrder As clsOrder                  'ҽ��������
Public mclsFee  As clsFee                     '�շѴ�����

Public mobjLisInsideComm As Object              'LIS�ӿڲ���

Public gclsReport As Object
Public gobjRegister As Object

Public Const HIS_CAPTION = "������ʾHIS����NEW"

Public Sub Main()
'------------------------------------------------
'���ܣ������򣬸����������Ӳ����鿴����
'������
'       Command��(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=127.0.0.1)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=testbase))):ZLHIS:AQA:1:2:1:1
'       Command�������壺Oracle�����ַ���:�û���:����:�Ƿ�����ת��(0��1):���ù��ܺ�(-1 ���ܳ�ʼ��,0-��������,1-�����鱨��,2-ҽ������;3-ִ�ж˸�������;4-ִ�ж˸���;5-��ӡ����;99-���Զ��屨��,999-���ܳ�ʼ��):...
'              ���ܺŲ�ͬ�����������ĸ�ʽ�뺬��Ҳ��ͬ
'              ����=0,1,2ʱ:���ܺź����������ID,��ҳID
'              ����=3,999ʱ�����ܺ��޲���
'              ����=4ʱ,���ܺ�Ϊ:����ID:ҽ����Ϣ:NOs                               ����ҽ����Ϣ��NOs���δ�һ������,ҽ����Ϣ��ִ�п���|ҽ��IDs(����ö��ŷָ�);NOs: ����ö��ŷָ�
'              ����=5ʱ,���ܺ�Ϊ����ӡ���(0=����ӡ��Ԥ��,1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel,4-�����PDF,99-��ӡ����):(��ʽ��������,���ݺ�(par)������,���ݺ�)���ܺ�Ϊ:  ������,���ݺ�(par)������,���ݺ�(par)������,���ݺ�
'              ����=99ʱ,���ܺ�Ϊ��ϵͳ��:������:��ӡ���(0=����ӡ��Ԥ��,1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel,4-�����PDF,99-��ӡ����):�������(��Ϊ�� ʾ����ʽ������id=1<par>PDF=C:\1.PDF<par>ExcelFile=C:\1.xls)
    
'���أ���
'-----------------------------------------------
    Dim strMsgs As String
    Dim blnLis As Boolean
    

    Dim varTmp As Variant
    
On Error GoTo ErrorHand

    
    strMsgs = Command
    
        '���յ�QUIT��ֱ���˳�
    If UCase(Trim(strMsgs)) = "QUIT" Then
            '����������Ѿ�������һ�Σ�����������ֱ��ˢ�½������ݺ��˳�
        If App.PrevInstance Then
            If SendMsg(strMsgs) Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    
    varTmp = Split(strMsgs, ":")
    
    glng���ܺ� = Val(varTmp(4)) '0-��������,1-LIS����,2-ҽ������;3-ִ�ж˸�������;4-ִ�ж˸���;5-���ݴ�ӡ;99-���Զ��屨��
    glngFunID = IIf(glng���ܺ� = 2, 3001, 0)

    If Trim(strMsgs) = "" Then Exit Sub
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
    End If
    
    
    '���ݴ���Ĳ����жϴ����ĸ���������Ϣ��ʽ���������:���ݿ��û���:����ID:����ID:ִ�в���ID:ҽ��ID��
    '����ID �� ����Ϊ�Һ�ID ���˹Һż�¼.ID��סԺΪ��ҳID
    
    '��ʼ��comlib�����ݿ�����
    If UBound(Split(strMsgs, MSG_SPLIT)) < 4 Then Exit Sub
    gstrZLHIS�����ַ��� = Split(strMsgs, MSG_SPLIT)(0)
    gstr�û��� = Split(strMsgs, MSG_SPLIT)(1)
    gstr���� = Split(strMsgs, MSG_SPLIT)(2)
    gbln�Ƿ�ת������ = Val(Split(strMsgs, MSG_SPLIT)(3)) = 1
    blnLis = glng���ܺ� = 2
    
    
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
    Dim lngWinHandle As Long        '��Ҫ������Ϣ�ġ�zlSoftCISInterface.exe������Ĵ��ھ��
    Dim wParam As Long
    Dim lResult As Long
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


