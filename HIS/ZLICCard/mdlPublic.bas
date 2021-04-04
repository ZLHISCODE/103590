Attribute VB_Name = "mdlPublic"
Option Explicit

Public gLastErr As String '�������һ�δ�����Ϣ
Public gbln�Զ���ȡ As Boolean '��ǰ�Ƿ�Ϊ��Ƶ��
Public gDebug As Boolean '���Կ���

'��������ֵ�ľֲ�����
Public gCol As Collection

Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object
Public gcnConnect As ADODB.Connection

Public Property Get Cards() As Collection
    Set Cards = gCol
End Property

Private Function Add(objCard As clsCard, Optional sKey As String) As clsCard
    '�����¶���
    Dim objNewMember As clsCard
    Set objNewMember = New clsCard

    '���ô��뷽��������
    Set objNewMember = objCard
    If Len(sKey) = 0 Then
        gCol.Add objNewMember
    Else
        gCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As clsCard
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = gCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = gCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    gCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = gCol.[_NewEnum]
End Property


Public Sub initCards()
    '��ʼ��IC������,�̻��������У�����IC���ӿ�ʱ����ĩβ���
    
    Dim objclsCard As clsCard
    '-- 1.Demo��,���������
    Set objclsCard = New clsCard
    objclsCard.���� = 1
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_Demo"
    objclsCard.���� = "����IC��(������)"
    objclsCard.�ɷ����� = 1 '��������
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
   '-- 2.�Ϻ�ҽ��IC��
    Set objclsCard = New clsCard
    objclsCard.���� = 2
    objclsCard.�ӿڳ����� = "zl9Insure.clsInsure"
    objclsCard.���� = "�Ϻ���ҽ��IC��"
    objclsCard.���� = 413
    objclsCard.�ɷ����� = 0 '��������
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
   '-- 3.����֤IC��
    Set objclsCard = New clsCard
    objclsCard.���� = 3
    objclsCard.�ӿڳ����� = "zlICCard.clsIDcard"
    objclsCard.���� = "�ڶ������֤"
    objclsCard.�ɷ����� = 0 '��������
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    '-- 4.����RDϵ��
    Set objclsCard = New clsCard
    objclsCard.���� = 4
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_MW_RD"
    objclsCard.���� = "����RDϵ��"
    objclsCard.�ɷ����� = 1 '������
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    '-- 5.���칫�ڳ���һ��ͨ
    Set objclsCard = New clsCard
    objclsCard.���� = 5
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_CQPubCard"
    objclsCard.���� = "���칫�ڳ���һ��ͨ"
    objclsCard.�ɷ����� = 0 '��������
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    
    '-- 6.���������ҽԺ��Ƶ���ӿ�
    Set objclsCard = New clsCard
    objclsCard.���� = 6
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_JCSRFID"
    objclsCard.���� = "���������ҽԺ��Ƶ��"
    objclsCard.�ɷ����� = 1 '��������
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 1))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    
    '-- 7.����һ��ͨ
    Set objclsCard = New clsCard
    objclsCard.���� = 7
    objclsCard.�ӿڳ����� = "zlICCard.clsIC_NBYKT"
    objclsCard.���� = "����һ��ͨ"
    objclsCard.�ɷ����� = 1 '��������
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����

    '-- 8.����D3��IC����д��
    '-- 2009-07-09 ZHQ �����ǻҽԺ����
    Set objclsCard = New clsCard
    objclsCard.���� = 8
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_D3IC"
    objclsCard.���� = "����D3��IC��"
    objclsCard.�ɷ����� = 1 '������
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    
    '-- 9.����������Ƶ��
    Set objclsCard = New clsCard
    objclsCard.���� = 9
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_URF_35H"
    objclsCard.���� = "����URF-35H��Ƶ��"
    objclsCard.�ɷ����� = 1 '��������
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    
    '-- 10.�Ű�һ��ͨ
    Set objclsCard = New clsCard
    objclsCard.���� = 10
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_SLE4428"
    objclsCard.���� = "�Ű�һ��ͨ"
    objclsCard.�ɷ����� = 1
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    
    '-- 11.����֤ͨ�𿨶�д�� ZT606
    Set objclsCard = New clsCard
    objclsCard.���� = 11
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_ZT606"
    objclsCard.���� = "����֤ͨ�𿨶�д��(ZT606)"
    objclsCard.�ɷ����� = 1
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    
    '-- 12.�������� MHCX�ſ���д��  MHCX-715K(�����������ſƼ����޹�˾)
    Set objclsCard = New clsCard
    objclsCard.���� = 12
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_MHCX_715K"
    objclsCard.���� = "��������MHCX�ſ���д��(MHCX_715K)"
    objclsCard.�ɷ����� = 1
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    
    '-- 13.��˼�ĺ�һ������  SS728MQ1(ɽ����˼���Ӽ����ɷ����޹�˾)
    Set objclsCard = New clsCard
    objclsCard.���� = 13
    objclsCard.�ӿڳ����� = "zlICCard.clsICCardDev_SS728MQ1"
    objclsCard.���� = "��˼�ĺ�һ������(SS728MQ1)"
    objclsCard.�ɷ����� = 1
    objclsCard.�Ƿ��Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & objclsCard.����, "�Զ���ȡ", 0))
    objclsCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objclsCard.����, 1) = 1
    Add objclsCard, "A" & objclsCard.����
    
End Sub

Public Sub WritLog(ByVal strDev As String, strInput As String, strOutput As String)
'    Call LogWrite("һ��ͨ�ӿڵ�����־", 1151, "�����ӿڷ���", "������:" & strDev & ";����:" & strInput & ";���:" & strOutput)
End Sub

Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnMessage As Boolean = True) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
    If err <> 0 Then
        If blnMessage = True Then
            '���������Ϣ
            strError = err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, "IC���ӿ�"
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, "IC���ӿ�"
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, "IC���ӿ�"
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, "IC���ӿ�"
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, "IC���ӿ�"
            End If
        End If
        
        err.Clear
        OraDataOpen = False
        Exit Function
    End If
    OraDataOpen = True
End Function
