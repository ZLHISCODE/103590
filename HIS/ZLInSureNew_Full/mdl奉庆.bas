Attribute VB_Name = "mdl����"
Option Explicit
Private mblnInit As Boolean
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
                         '2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

Public Enum ҵ������_����
    ����_ҽ����ʼ�� = 0
    ����_��òα���Ա��Ϣ
    ����_�α��ʸ����
    ����_���˵Ǽ�
    ����_���ͨ����������Ϣ
    ����_���¾�����Ϣ
    ����_ȷ����ϸ��Ŀ����
    ����_¼�봦����ϸ
    ����_�����˷�
    ����_��Ŀ���������ѯ
    ����_Ԥ����
    ����_����
    ����_��������
    ����_�շ�Ŀ¼����Ԥ����
    ����_�շ�Ŀ¼���ش���
    ����_��ȡ��������
    ����_�������_��չ��Ϣ
End Enum
Private gInitCard As Boolean                '��ʼ���˿���
Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    
    ������������ As Boolean
End Type
Public InitInfor_���� As InitbaseInfor

Private Type �������
        ����        As String
        ҽ��֤��    As String       '��ҽ����
        ����     As String
        �Ա�     As String
        ���֤�� As String
        ��������  As String
        ����        As Integer
        ������    As String   '��Ա������
        �������    As String   '��Ա�������
        ��Ա״̬    As String
        ��λ����    As String
        ��λ����    As String
        ҽ�����    As String 'ҽ����Ա����������
        ���ֱ���    As String
        ��������    As String
        ͳ������    As String
        ����ID      As Long
        �������    As String
        �ʻ����    As Double
        �����      As String
        סԺ��      As String
        
        �������    As String
        ��Ŀ����    As String
        ��Ŀ����    As String
        
        �����ܶ�    As Double
        ��̬��Ϣ    As String   '����16λ��̬��Ϣ
        ��չ��Ϣ    As String
        ����ID      As Long
        
        ��Ժ��ϱ��� As String
        ��Ժ�������    As String
        
        ȷ����ϱ���    As String
        ȷ���������    As String
        ��;����    As Boolean
        �¸����ʻ� As Boolean
        
        סԺ���� As String
End Type

Public g�������_���� As �������
Public gcnOracle_���� As ADODB.Connection     '�м������
Private gcnOracle_���� As ADODB.Connection     '����ҽ���������Ӵ�

Private Type ��������
        �����ܶ�    As Double
        ͳ��֧��    As Double
        �˻�֧��    As Double
        �ֽ�֧��    As Double
        �󲡵渶    As Double
        
        ��̬��Ϣ    As String
        ������ˮ��  As String
End Type
Private ����������� As ��������

'1 �ӿڳ�ʼ��
Private Declare Function Init Lib "SiInterface" Alias "INIT" () As Long
'2 ҵ��������ִ��ҽ��ҵ������Ҫ�Ĵ���
Private Declare Function OperationAsk Lib "SiInterface" Alias "BUSINESS_HANDLE" _
    (ByVal StrInput As String, ByVal strOutput As String) As Long

'3 ҵ���ѯ������ִ��ҽ��ҵ������Ҫ�Ĵ���
Private Declare Function OperationQuery Lib "SiInterface" Alias "QUERY_HANDLE" _
    (ByVal StrInput As String, ByVal strOutput As String) As Long
    
'4.��д�������������
Public gobj���� As Object
Public gobj����Err As Object
Private Const STR_����ά������ = "1"

Public Function ҽ����ʼ��_����() As Boolean
    Dim strReg As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    Dim bln������ As Boolean
    
    If mblnInit Then
        ҽ����ʼ��_���� = True
        Exit Function
    End If
    GetRegInFor g����ȫ��, "ҽ��", "������", strReg
    bln������ = Val(strReg) = 1

    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_����.ģ������ = True
    Else
        InitInfor_����.ģ������ = False
    End If
    
    Call GetRegInFor(g����ģ��, "����", "������������", strReg)
    If Val(strReg) = 1 Then
        InitInfor_����.������������ = True
    Else
        InitInfor_����.������������ = False
    End If
    
    InitInfor_����.������������ = InitInfor_����.������������ Or InitInfor_����.ģ������
    
    
    '��������ҽ������
    If gInitCard = True And bln������ Then
        Call sCard_CloseCardWithoutSave
    End If
    Set gobj���� = Nothing
    
    Err = 0
    On Error Resume Next
    Set gobj���� = CreateObject("SiCard.SiCardOperator")
    If Err <> 0 Then
        If InitInfor_����.ģ������ Then
        Else
            ShowMsgbox "���������д����ʧ��!"
            Exit Function
        End If
    End If
    Set gobj����Err = CreateObject("SiCommTool.SiErrorCtl")
    If Err <> 0 Then
        If InitInfor_����.ģ������ Then
        Else
            ShowMsgbox "���������д����ʧ��!"
            Exit Function
        End If
    End If
     
    '��ʼ����д�����
    If bln������ Then
        If sCard_InitCard = False Then
                If Not InitInfor_����.ģ������ Then
                    Exit Function
                End If
        End If
    End If
    gInitCard = True
    
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_�ٲ׷���)
    InitInfor_����.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
    
    If Open�м�� = False Then Exit Function
    
    
    If gInitCard Then
        '��ʼ��ҽ���ӿ�
        If ҵ������_����(����_ҽ����ʼ��, "", strOutput) = False Then
            Exit Function
        End If
    End If
    mblnInit = True
    ҽ����ʼ��_���� = True
End Function
Private Function Open�м��() As Boolean
    '�����м��
    '�м������
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strServer As String, strPass As String
    
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�山ҽ��", TYPE_�ٲ׷���)
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_���� = New ADODB.Connection

    If OraDataOpen(gcnOracle_����, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If


    '�м������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like '����%' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�山ҽ��", TYPE_�ٲ׷���)
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "�����û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "���ķ�����"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "�����û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_���� = New ADODB.Connection

    If OraDataOpen(gcnOracle_����, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���������ݿ⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    Open�м�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function ������_����() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������������
    '--�����:strCardData-��������
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strOutput As String
    Dim StrInput As String
    Dim strArr As Variant
    Err = 0
    On Error GoTo errHand:
    
    If InitInfor_����.������������ Then
        Readģ������ ��ȡ��������, StrInput, strOutput
        If strOutput = "" Then
            ������_���� = False
            Exit Function
        End If
        strArr = Split(strOutput, "|")
        With g�������_����
            .ҽ��֤�� = strArr(1)
            .���֤�� = strArr(2)
            .��λ���� = strArr(3)
            .���� = strArr(4)
            .���� = strArr(5)
            .�Ա� = strArr(6)
            .�������� = strArr(7)
            .������ = strArr(8)
            .��Ա״̬ = strArr(10)
            .�ʻ���� = Val(strArr(12))
        End With
        
    Else
        '��ȡ��
        If sCard_ReadCard = False Then Exit Function
        '��ȡ����Ϣ
        'bytType:1-SiCardBaseInfo�籣��������Ϣ
        '        2-SiCardDynaInfo�籣����̬��Ϣ
        '        3-SiCardAcctInfo�籣���ʻ���Ϣ
        '        4-SiCardExtInfo�籣����չ��Ϣ
        If sCard_����ֵ(1, strOutput) = False Then Exit Function
        '���˱��|���֤��|��λ����|�籣����|����|�Ա�|��������|��Ա���|�α�����|��Ա״̬|�������
        strArr = Split("0|" & strOutput, "|")
        With g�������_����
            .ҽ��֤�� = strArr(1)
            .���֤�� = strArr(2)
            .��λ���� = strArr(3)
            .���� = strArr(4)
            .���� = strArr(5)
            .�Ա� = strArr(6)
            .�������� = strArr(7)
            .������ = strArr(8)
            .��Ա״̬ = strArr(9)
        End With
        If sCard_����ֵ(3, strOutput) = False Then Exit Function
        '��ȡ�����ʻ����
        '�ʻ����
        strArr = Split(strOutput, "|")
        With g�������_����
            .�ʻ���� = Val(strArr(0))
        End With
        
        '��ȡ��̬��Ϣ
        If sCard_����ֵ(2, strOutput) = False Then Exit Function
        With g�������_����
            .��̬��Ϣ = strOutput
        End With
        '��ȡ��չ��Ϣ
        If sCard_����ֵ(4, strOutput) = False Then Exit Function
        With g�������_����
            .��չ��Ϣ = strOutput
        End With
    End If
    ������_���� = True
    Exit Function
errHand:
    ������_���� = False
    ShowMsgbox "IC������,����ʶ��!"
End Function
Private Function sCard_InitCard() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    '�ɹ������� ʧ�ܷ��ش�����Ĵ����
    Err = 0
    On Error GoTo errHand:
    lngReturn = gobj����.InitCard()
    If lngReturn <> 0 Then
        '
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    sCard_InitCard = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function sCard_����ֵ(ByVal bytType As Long, strPropertyValue As String, Optional blnWrite As Boolean = False) As Boolean
    Dim lngReturn As Long
    Dim strReturn As String
    Dim strErrInfor As String
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    
    '�ɹ������� ʧ�ܷ��ش�����Ĵ����
    sCard_����ֵ = False
    
    If InitInfor_����.ģ������ Then
        sCard_����ֵ = True
        Exit Function
    End If
    Err = 0
    On Error GoTo errHand:
    Select Case bytType
        Case 1  ' SiCardBaseInfo�籣��������Ϣ
            '���˱��|���֤��|��λ����|�籣����|����|�Ա�|��������|��Ա���|�α�����|��Ա״̬|�������
            If blnWrite Then
                gobj����.SiCardBaseInfo = strPropertyValue
                DebugTool "д��������Ϣ��" & strPropertyValue
            Else
                strReturn = gobj����.SiCardBaseInfo
            End If
            
            
        Case 2  ' SiCardDynaInfo�籣����̬��Ϣ
            If blnWrite Then
                gobj����.SiCardDynaInfo = strPropertyValue
                DebugTool "д����̬��Ϣ��" & strPropertyValue
            Else
                strReturn = gobj����.SiCardDynaInfo
            End If
        Case 3  ' SiCardAcctInfo�籣���ʻ���Ϣ
            '���ʻ����
            If blnWrite Then
                gobj����.SiCardAcctInfo = strPropertyValue
                DebugTool "д�ʻ���" & strPropertyValue
            Else
                strReturn = gobj����.SiCardAcctInfo
            End If
        Case Else 'SiCardExtInfo�籣����չ��Ϣ
            '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
            If blnWrite Then
                gobj����.SiCardExtInfo = strPropertyValue
                DebugTool "д��չ��Ϣ��" & strPropertyValue
            Else
                strReturn = gobj����.SiCardExtInfo
            End If
    End Select
    If blnWrite Then
    Else
        strPropertyValue = strReturn
    End If
    sCard_����ֵ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function sCard_CloseCardWithoutSave() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    '��ȡд�����ʼ��.
    '��ʼ��ҽ����
    '�ɹ������� ʧ�ܷ��ش�����Ĵ����
    Err = 0
    On Error GoTo errHand:
    If gobj���� Is Nothing Then Exit Function
    lngReturn = gobj����.CloseCardWithoutSave()
    If lngReturn <> 0 Then
        '
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    sCard_CloseCardWithoutSave = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function sCard_ReadCard() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    '��ȡд�����ʼ��.
    '��ʼ��ҽ����
    '�ɹ������� ʧ�ܷ��ش�����Ĵ����
    Err = 0
    On Error GoTo errHand:
    
    lngReturn = gobj����.ReadCard()
    If lngReturn <> 0 Then
        '
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    sCard_ReadCard = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function sCard_SaveCard() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    'IDL���� HRESULT SaveCard([in] BSTR prm_ctlStr, [out, retval] int *ret_appCode)
    '�������� int SaveCard (BSTR prm_ctlStr)
    '�� �� ֵ �ɹ������� ʧ�ܷ��ش�����Ĵ����
    '�� �� prm_ctlStr���� д�����ƴ����ÿ��ƴ�������������Ժ��ַ�ʽд����ȫд���߲���д�������硰Athene��������¶�̬��Ϣ��������������Ϣ����Apollo����ʾ���¶�̬����չ��Ϣ��������������Ϣ�����ڷ��籣���Ŀͻ�������ԣ������Ķ�̬��Ϣ���ʻ���Ϣ����չ��Ϣ�����Դ���������ʹ�á�Apollo��������ʵ�
    '˵ �� д�������Ǹ��ݶ���ʱ����SiCard�������Ϣ���û������SiCardDynaInfo�����Խ������ú���ĵĻ�����Ϣд�뿨�ϡ�������ִ��д������֮ǰ������ȷ��ϣ��д������Ϣ�Ѿ���ȷ�����ø��˶�Ӧ���������
    '�� �� �� ȫ��֧�֣���������Memory��֧���ض������������д��
    
    Err = 0
    On Error GoTo errHand:
        
    If InitInfor_����.ģ������ Then
        sCard_SaveCard = True
        Exit Function
    End If
    lngReturn = gobj����.SaveCard("Apollo")
    If lngReturn <> 0 Then
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    DebugTool "д���ɹ�"
    sCard_SaveCard = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function sCard_SetupCardOption_����() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    '��ȡд�����ʼ��.
    '��ʼ��ҽ����
    '�ɹ������� ʧ�ܷ��ش�����Ĵ����
    Err = 0
    On Error GoTo errHand:
    '�ȳ�ʼ����
    If gobj���� Is Nothing Then
         Set gobj���� = CreateObject("SiCard.SiCardOperator")
    End If
    lngReturn = gobj����.SetupCardOption()
    If lngReturn <> 0 Then
        '
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    sCard_SetupCardOption_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function sCard_ErrInfor(lngReturn As Long) As String
    '��ȡ��д������Ĵ�������
    Dim strReturn As String
    
    '��ʼ��ҽ����
    '�ɹ������� ʧ�ܷ��ش�����Ĵ����
    Err = 0
    On Error GoTo errHand:
    If gobj����Err Is Nothing Then
        Set gobj����Err = CreateObject("SiCommTool.SiErrorCtl")
    End If
    strReturn = gobj����Err.Describe(lngReturn)
    sCard_ErrInfor = strReturn
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��ȡסԺ״̬_����(ByRef lng״̬ As Long) As Boolean
    '����lng״̬,(0-����,1��Ժ)
    Dim strOutput As String
    Dim strArr As Variant
    
    ��ȡסԺ״̬_���� = False
    Err = 0
    On Error GoTo errHand:
    '??lng״̬ = GetHospstatus()
    '        1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    If InitInfor_����.ģ������ Then
        '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
        Call Readģ������(�������_��չ��Ϣ, "", strOutput)
        If strOutput = "" Then Exit Function
    Else
        If sCard_����ֵ(4, strOutput) = False Then Exit Function
        '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
    End If
    strArr = Split(strOutput, "|")
    lng״̬ = IIf(Val(strArr(7)) = 1, 1, 0)
    ��ȡסԺ״̬_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ҽ����ֹ_����() As Boolean
    '������д�����
    Dim bln������ As Boolean
    Dim strReg As String
    
    mblnInit = False
    GetRegInFor g����ȫ��, "ҽ��", "������", strReg
    bln������ = Val(strReg) = 1
    
    If bln������ Then
        If sCard_CloseCardWithoutSave = False Then
            If Not InitInfor_����.ģ������ Then
                Exit Function
            End If
        End If
    End If
    gInitCard = False
    
    Err = 0
    On Error Resume Next
    
    Set gobj���� = Nothing
    Set gobj����Err = Nothing
    If gcnOracle_����.State = 1 Then
        gcnOracle_����.Close
    End If
    If Not gcnOracle_���� Is Nothing Then
        gcnOracle_����.Close
    End If
    ҽ����ֹ_���� = True
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Dim bln������ As Boolean
    Dim strReg As String
    
    GetRegInFor g����ȫ��, "ҽ��", "������", strReg
    bln������ = Val(strReg) = 1
    If bln������ = False Then
        ShowMsgbox "û�ж����������ܽ��������֤,���ڱ������������"
        Exit Function
    End If
    Err = 0
    On Error GoTo errHand:
    ��ݱ�ʶ_���� = frmIdentify����.GetPatient(bytType, lng����ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_���� = ""
End Function


Public Function �������_����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_�ٲ׷���)
    
    If rsTemp.EOF Then
        �������_���� = 0
    Else
        �������_���� = rsTemp("�ʻ����")
    End If
End Function
Public Function �α��ʸ����_����() As Boolean
        '���ܣ���֤��ǰҽ����Ա���ʸ����
        '���أ�����true,���򷵻�False
        
        Dim StrInput As String
        Dim strOutput As String
        �α��ʸ����_���� = False
        Dim strArr
        
        With g�������_����
            StrInput = .ҽ��֤�� & "|"
            StrInput = StrInput & .���� & "|"
            StrInput = StrInput & .ͳ������
        End With
        
        '���: ҽ��֤����|IC����|ͳ������
        '����: ����ԭ��������
     
        Err = 0
        On Error GoTo errHand:
        If ҵ������_����(����_�α��ʸ����, StrInput, strOutput) = False Then
            Exit Function
        End If
        �α��ʸ����_���� = True
        Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function

Private Function ������ϸд��(ByVal rs��ϸ As ADODB.Recordset, Optional ByVal bln���� As Boolean = True) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr
    Dim strDate As Date
    
    Dim str������� As String
    ������ϸд�� = False
    g�������_����.�����ܶ� = 0
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Err = 0
    On Error GoTo errHand:
    'Ȼ����봦����ϸ
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        
        Do Until rs��ϸ.EOF
            gstrSQL = "select A.����,A.����,A.���,A.���,A.���㵥λ,B.��Ŀ����,B.��ע,B.�Ƿ�ҽ��,A.���㵥λ,E.���,G.���� ���� " & _
                      "from �շ�ϸĿ A,(Select ��Ŀ����, ��ע,�Ƿ�ҽ��,�շ�ϸĿID From ����֧����Ŀ where ����=" & TYPE_�ٲ׷��� & ") B,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
                      "where A.ID=[1] and A.ID=B.�շ�ϸĿID(+) " & _
                     "        AND A.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) "
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", CLng(rs��ϸ("�շ�ϸĿID")))
            If rsTemp.EOF = True Then
                MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            
            gstrSQL = "" & _
                  "   Select �շѵȼ�,�շ���� From ҽ���շ�Ŀ¼ " & _
                  "   Where ���=[1] and ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҽ�������Ϣ", CStr(Nvl(rsTemp!��ע)), CStr(Nvl(rsTemp!��Ŀ����)))
            
              
            If Val(Nvl(rs��ϸ("ʵ�ս��"), 0)) <> 0 Then
                
                '��ȡ������������Ŀ���
                If Nvl(rsTemp!�Ƿ�ҽ��, 1) = 0 Then
                    '���˱��|��Ŀ���|��������
                    StrInput = g�������_����.ҽ��֤�� & "|"
                    
                    If Val(Nvl(rsTemp!��ע)) = 1 Then
                        '��˵:ҩƷֻ�ܴ�ҽԺ����,������ֻ�ܵ�ֻ�ܴ�ҽ������
                        StrInput = StrInput & Nvl(rsTemp!����, "9000900099") & "|"
                    Else
                        StrInput = StrInput & Nvl(rsTemp!��Ŀ����, "9000900099") & "|"
                    End If
                    
                    StrInput = StrInput & strDate
                    
                    If ҵ������_����(����_��Ŀ���������ѯ, StrInput, strOutput) = False Then
                        strOutput = "|"
                    End If
                    strArr = Split(strOutput, "|")
                    str������� = strArr(1)
                Else
                    str������� = ""
                End If
                'סԺ(����)��|������|���������|�������|ҽԺ����|ҽ������|��Ŀ����|���õȼ�|
                '�������|����|����|���|��λ|���|����|��������|��������|����ҽ��|¼���־
                StrInput = g�������_����.����� & "|"
                StrInput = StrInput & g�������_����.����� & "|"
                If Not bln���� Then
                    StrInput = StrInput & Nvl(rs��ϸ!ID, 0) & "|"
                Else
                    StrInput = StrInput & Nvl(rs��ϸ.AbsolutePosition, 0) & "|"
                End If
                StrInput = StrInput & str������� & "|"
                StrInput = StrInput & Nvl(rsTemp!����) & "|"
                
                StrInput = StrInput & Nvl(rsTemp!��Ŀ����, "9000900099") & "|"
                StrInput = StrInput & Nvl(rsTemp!����) & "|"
                If rsTmp.EOF Then
                    StrInput = StrInput & "3" & "|"
                    
                    StrInput = StrInput & Split(Get�������(Nvl(!�շ����)), "-")(0) & "|"
                Else
                    StrInput = StrInput & Nvl(rsTmp!�շѵȼ�) & "|"
                        If IsNull(rsTmp!�շ����) Then
                            StrInput = StrInput & Split(Get�������(Nvl(!�շ����)), "-")(0) & "|"
                        Else
                            StrInput = StrInput & Nvl(rsTmp!�շ����) & "|"
                        End If
                End If
                
                StrInput = StrInput & Format(rs��ϸ("����"), "0.0000") & "|"
                StrInput = StrInput & Format(rs��ϸ("����"), "0.00") & "|"
                StrInput = StrInput & Format(rs��ϸ("ʵ�ս��"), "#####0.0000") & "|"         '���
                
                StrInput = StrInput & ToVarchar(rsTemp("���㵥λ"), 20) & "|"      '��λ
                StrInput = StrInput & ToVarchar(rsTemp("���"), 14) & "|"
                StrInput = StrInput & ToVarchar(rsTemp("����"), 20) & "|"
                StrInput = StrInput & strDate & "|"
                DebugTool "�����������"
                If bln���� Then
                    StrInput = StrInput & UserInfo.���� & "|"
                Else
                    gstrSQL = "Select ���� From ���ű� where id=" & Nvl(!��������ID, 0)
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��������"
                    StrInput = StrInput & Nvl(rsTemp!����) & "|"
                End If
                StrInput = StrInput & Nvl(rs��ϸ!������) & "|"
                
                 '0 ��ʾ��ʼѭ����2 ��ʾ����ѭ�����ڽ���ѭ������ύ
'                If rs��ϸ.AbsolutePosition = 1 Then
'                    If rs��ϸ.AbsolutePosition = rs��ϸ.RecordCount Then
                        StrInput = StrInput & "1"
'                    Else
'                        strInPut = strInPut & 0
'                    End If
'                ElseIf rs��ϸ.AbsolutePosition = rs��ϸ.RecordCount Then
'                    strInPut = strInPut & 2
'                Else
'                    strInPut = strInPut & "1"
'                End If
                
                If ҵ������_����(����_¼�봦����ϸ, StrInput, strOutput) = False Then
                    Exit Function
                End If
                
                If Not bln���� Then
                    '����������㣬��ȷ����ص��ϴ���־ֵ
                    '�������
                    'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                    'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                    'ժҪֵ:������ˮ��|�������|סԺ(����)��|������|ʵ�ʽ��׵���|ʵ�ʵȼ�
                    strArr = Split(strOutput, "|")  '--ʵ�ʵ���|ʵ�ʵȼ�|������ˮ��
                    
                    strOutput = strArr(3) & "|" & str������� & "|" & g�������_����.����� & "|" & g�������_����.����� & "|" & Val(strArr(1)) & "|" & strArr(2)
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strOutput & "')"
                    zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
                End If
            End If
            g�������_����.�����ܶ� = g�������_����.�����ܶ� + Nvl(rs��ϸ!ʵ�ս��, 0)
            rs��ϸ.MoveNext
        Loop
    End With
    ������ϸд�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function �������ά������(ByVal lngϸĿID As Long, ByVal dbl���� As Double) As Boolean
    '�������ά���еĵ����Ƿ���HIS�����һ��
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    �������ά������ = False
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "" & _
        "   Select ά����־,����,�շ����,��� From ҽ���շ�Ŀ¼ " & _
        "   Where (���,����) in (select ��ע,��Ŀ���� From ����֧����Ŀ where ����=[1] and �շ�ϸĿid=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������ά������", TYPE_�ٲ׷���, lngϸĿID)
    If rsTemp.EOF Then
        �������ά������ = True
        'MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
        Exit Function
    End If
    '������ѯ:��ҽ��ά���ļ۸�Ϊ1��NULL(������ʱ���Ѿ����¸ñ�־Ϊ1),������ά���ļ۸�Ϊ:0
    If Nvl(rsTemp!ά����־) <> STR_����ά������ Then
        'ȷ����ϸ��Ŀ����
        '   ҽ�����|�������|ҽԺ����
        StrInput = Nvl(rsTemp!����) & "|"
        StrInput = StrInput & Nvl(rsTemp!�շ����) & "|"
        StrInput = StrInput & Format(dbl����, "0.0000")
        If ҵ������_����(����_ȷ����ϸ��Ŀ����, StrInput, strOutput) = False Then
            Exit Function
        End If
        strOutput = Split(strOutput & "|", "|")(1)
        If Val(strOutput) <> dbl���� Then
            gstrSQL = "Select ���� From �շ�ϸĿ where ID=" & lngϸĿID
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡϸĿ����"
            ShowMsgbox "������Ŀ��" & rsTemp!���� & "���ĵ��۲�һ����" & vbCrLf & " ��ҽԺ:" & Format(dbl����, "0.0000") & vbCrLf & "������:" & Format(Val(strOutput), "0.0000")
            Exit Function
        End If
    End If
    �������ά������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Static str�ϴ������ As String
    Dim strҽ���� As String, StrInput As String, strOutput  As String
    Dim dbl�����ʻ� As Double, strMessage As String
    Dim lng����ID As Long, str��� As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    Dim strArr
    Dim strDate As String
    
    On Error GoTo errHandle
    
    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_���� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    
    If g�������_����.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    '�����˵���ǰ����������δ��ķ��ã��������ִ��Ԥ����
    If str�ϴ������ = g�������_����.����� Then
        '�Ѿ���ֵ��˵���ò��˽��й�Ԥ��
        StrInput = str�ϴ������ & "|" & str�ϴ������
        If ҵ������_����(����_�����˷�, StrInput, strOutput) = False Then
            'Exit Function
        End If
    End If
    
    Dim rsTmp As New ADODB.Recordset
    Dim dbl���� As Long
    
    '���ȼ�鵥�۷������
    With rs��ϸ
        Do While Not .EOF
            If �������ά������(Nvl(!�շ�ϸĿID, 0), Nvl(!����, 0)) = False Then Exit Function
            .MoveNext
        Loop
    End With
    
'    str�ϴ������ = g�������_����.�����
    'Ȼ����봦����ϸ
    If ������ϸд��(rs��ϸ, True) = False Then Exit Function
    
    '����Ԥ����
    '    �����ض��������ݣ�    סԺ�������
    '    �����ض��������:   �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�֧��|�󲡵渶
    '                        2.41|0|0|2.41|0
    StrInput = g�������_����.�����
    'strInput = strInput & "|" & IIf(g�������_����.�¸����ʻ�, "1", "0")
    
    If ���¾�����Ϣ_����(0, strOutput) = False Then Exit Function
    If ҵ������_����(����_Ԥ����, StrInput, strOutput) = False Then
        '����,���ԭ���������˷�
        '�Ѿ���ֵ��˵���ò��˽��й�Ԥ��
        StrInput = g�������_����.����� & "|" & g�������_����.�����
        If ҵ������_����(����_�����˷�, StrInput, strOutput) = False Then
            Exit Function
        End If
    End If
    '�����ֵ
    str�ϴ������ = g�������_����.�����
    
    strArr = Split(strOutput, "|")
    
    str���㷽ʽ = "�����ʻ�;" & Val(strArr(3)) & ";0"  '�����޸ĸ����ʻ�����Ϊ����ʱ�Ѿ����ٴ���ǰ�û���
    
    If Val(strArr(2)) > 0 Then
        str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & Val(strArr(2)) & ";0"
    End If
    If Val(strArr(5)) > 0 Then
        str���㷽ʽ = str���㷽ʽ & "|�󲡵渶;" & Val(strArr(5)) & ";0"
    End If
        
    �����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function Get��չ��Ϣ(ByVal lngסԺ״̬ As String, Optional str��Ժ���� As String) As String
    '��ȡ��չ��Ϣ
    Dim strTemp As String
    Dim strArr
    Dim i As Integer
    If g�������_����.��չ��Ϣ = "" Then Exit Function
    
    strArr = Split(g�������_����.��չ��Ϣ, "|")
    '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
    If str��Ժ���� = "" Then
    Else
        strArr(5) = str��Ժ����
    End If
    strArr(6) = lngסԺ״̬
    'strArr(7) = InitInfor_����.ҽԺ����
    'strArr(8) = g�������_����.ҽ�����
    
    For i = 0 To UBound(strArr)
        strTemp = strTemp & "|" & strArr(i)
    Next
    Get��չ��Ϣ = Mid(strTemp, 2)
End Function


Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim StrInput As String, strOutput As String
    Dim lng����ID  As Long
    Dim dbl�����ܶ� As Double
    Dim str����Ա As String, strArr
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim str��̬��Ϣ As String
    Dim str������ˮ�� As String
    Static str����ʱ�� As String
    Static oldlng����ID As Long
    Dim lng������� As Long
        
    
    Dim datCurr As Date
    On Error GoTo errHandle
    Call DebugTool("�����������")
    gstrSQL = "Select a.*, a.����*a.���� as ����,a.ʵ�ս��/(nvl(a.����,1)*nvl(a.����,1)) as ���� From ������ü�¼ a Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ϸ��¼", lng����ID)
    If rs��ϸ.EOF = True Then
        Err.Raise 9000, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If

    lng����ID = rs��ϸ("����ID")
    str����Ա = ToVarchar(IIf(IsNull(rs��ϸ("����Ա����")), UserInfo.����, rs��ϸ("����Ա����")), 20)
    
    If g�������_����.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣"
        Exit Function
    End If
        
    If lng����ID = oldlng����ID And str����ʱ�� = Format(rs��ϸ!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS") Then
            '��Ҫ���·���һ���ºŸ�����
            gstrSQL = "Select nvl(�������,0)+1 as ������� From �����ʻ� where ����=" & TYPE_�ٲ׷��� & " and ����id=" & lng����ID
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�������"
            
            '���±����ʻ�
            lng������� = Nvl(rsTemp!�������, 1)
            g�������_����.����� = lng����ID & "_" & lng�������
            
            '���±����ʻ�
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ٲ׷��� & ",'�������','" & lng������� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����������")
    
            '��������Ǽ�
            '�����ض��������ݣ�סԺ�������|ҽ��֤����|IC����|��Ժ����|��Ժ��������|������
            With g�������_����
                StrInput = .����� & "|"
                StrInput = StrInput & .ҽ��֤�� & "|"
                StrInput = StrInput & .���� & "|"
                StrInput = StrInput & .ҽ����� & "|"
                StrInput = StrInput & "" & "|"
                StrInput = StrInput & "" & "|"
                StrInput = StrInput & gstrUserName
            End With
            If ҵ������_����(����_���˵Ǽ�, StrInput, strOutput) = False Then Exit Function
            If ���¾�����Ϣ_����(0, strOutput) = False Then Exit Function
    Else
        '�����������ʱ���Ѿ�����һ��,�������޷����������ϸ�еĽ�����ˮ��,�����������ϵ������ϴ�
        StrInput = g�������_����.����� & "|" & g�������_����.�����
        If ҵ������_����(����_�����˷�, StrInput, strOutput) = False Then
            Exit Function
        End If
    End If
    
    str����ʱ�� = Format(rs��ϸ!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
    oldlng����ID = Nvl(rs��ϸ!����ID, 0)

    
        
    'д����ϸ
    If ������ϸд��(rs��ϸ, False) = False Then Exit Function
    

    '���ý���
    Call DebugTool("׼�������������")
    '�����ض��������ݣ�  ��������|סԺ(����)��|���ݺ�|����Ա����
    '�����ض��������:
    '                  �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�֧��|�󲡵渶|����17�̬��Ϣ|������ˮ��
    '�������Ͷ������£�
    '   1�������� (��Ժ����)
    '   0סԺ��;����
    '   -1������
    '   -2IC����ʧ���Ժ���㣬���ν��㣨ֻ���סԺ�������з���תΪ�ֽ�֧��������ҽ�����ı�����
    '   �˻����ѱ�־ 0 �����˻����ѣ�����־Ϊ0�����¸����ʻ����� 1  ʹ��ϵͳ��������ֵ������־Ϊ1���¸����ʻ���
    
    StrInput = "1|"
    StrInput = StrInput & g�������_����.����� & "|"
    StrInput = StrInput & g�������_����.����� & "|"
    StrInput = StrInput & str����Ա
    'strInput = strInput & IIf(g�������_����.�¸����ʻ�, "1", "0")
    If ҵ������_����(����_����, StrInput, strOutput) = False Then
        Exit Function
    End If
    Call DebugTool("���ս����¼")
    
    '��������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
            
    Dim curͳ��֧�� As Double
    Dim cur����Ա���� As Double
    Dim cur�󲡵渶 As Double
    strArr = Split(strOutput, "|")
    
    dbl�����ܶ� = Round(g�������_����.�����ܶ�, 2)
    
    curͳ��֧�� = Val(strArr(2))
    cur�����ʻ� = Val(strArr(3))
    cur�󲡵渶 = Val(strArr(5))
    Dim i As Integer
    str��̬��Ϣ = ""
    
    '��ȡ��̬��Ϣ
    For i = 6 To UBound(strArr) - 1
        str��̬��Ϣ = str��̬��Ϣ & "|" & strArr(i)
    Next
    str��̬��Ϣ = Mid(str��̬��Ϣ, 2)
    
    '����д��
    '���ö�̬����
    If sCard_����ֵ(2, str��̬��Ϣ, True) = False Then
        GoTo Err������:
    End If
    
    'д��չ��Ϣ
    '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    
    
    If sCard_����ֵ(4, Get��չ��Ϣ("2"), True) = False Then
        GoTo Err������:
    End If
    str������ˮ�� = strArr(UBound(strArr))
    
    If sCard_SaveCard = False Then GoTo Err������:
     
    
    '�ʻ������Ϣ
    datCurr = zlDatabase.Currentdate
    
    Call Get�ʻ���Ϣ(TYPE_�ٲ׷���, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�ٲ׷��� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ������Ϣ")
    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(�ʻ������ۼ�),�ʻ��ۼ�֧��_IN(�ʻ��ۼ�֧��),�ۼƽ���ͳ��_IN(�ۼƽ���ͳ��_IN),�ۼ�ͳ�ﱨ��_IN(�ۼ�ͳ�ﱨ��),סԺ����_IN(סԺ�����ۼ�),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(����ʱ������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
    ' �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�֧��|�󲡵渶|����17�̬��Ϣ|������ˮ��
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�ٲ׷��� & "," & lng����ID & "," & Year(datCurr) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0," & IIf(g�������_����.�¸����ʻ�, 1, 0) & "," & _
            dbl�����ܶ� & ",0,0," & _
            curͳ��֧�� & "," & curͳ��֧�� & ",0,0," & cur�����ʻ� & ",'" & _
            str������ˮ�� & "',NULL,NULL,'" & g�������_����.����� & "|" & str��̬��Ϣ & "')"
            
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
    �������_���� = True
    Exit Function
Err������:
    '����ʧ��:���з�����
    '���з�����
    '�����ض��������ݣ�  ��������|סԺ(����)��|���ݺ�|����Ա����
    '�����ض��������:  �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�x֧��|�󲡵渶|����16�̬��Ϣ|������ˮ��
    '�������Ͷ������£�
    '   1�������� (��Ժ����)
    '   0סԺ��;����
    '   -1������
    '   -2IC����ʧ���Ժ���㣬���ν��㣨ֻ���סԺ�������з���תΪ�ֽ�֧��������ҽ�����ı�����

    StrInput = "-1|"
    StrInput = StrInput & g�������_����.����� & "|"
    StrInput = StrInput & g�������_����.����� & "|"
    StrInput = StrInput & str����Ա
  '  strInput = strInput & IIf(g�������_����.�¸����ʻ�, "1", "0")
    Call ҵ������_����(����_����, StrInput, strOutput)
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function ���¾�����Ϣ_����(ByVal bytType As Byte, strOutPutstring As String, Optional ByVal str��Ժ���� As String = "", Optional ByVal str��Ժ���� As String = "", Optional bln������ As Boolean = False) As Boolean
        'bytType:0-����,1-��Ժ�Ǽ�,2-סԺ����
        Dim StrInput As String, i As Integer, strTemp As String
        Dim strArr
        Dim strDate As String
         If InitInfor_����.ģ������ Then
            ���¾�����Ϣ_���� = True
            strOutPutstring = "01|2"
            Exit Function
         End If
        '���¾�����Ϣ
        ���¾�����Ϣ_���� = False
        Err = 0
        On Error GoTo errHand:
        Select Case bytType
            Case 0 '����
                'ֻ����ҽ�����,
                'סԺ��|���±�־|ҽ�����|��Ժ����|��Ժ��������|��Ժ����|ȷ�Ｒ������|������|�ʻ����|סԺ����|��̬��Ϣ
                '����:δ��
                
                strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
                StrInput = g�������_����.����� & "|"
                If bln������ Then
                    'ֻ���¶�̬��Ϣ�Ͳ���Ա
                    StrInput = StrInput & "0000011011" & "|"
                ElseIf g�������_����.���ֱ��� = "" Or g�������_����.���ֱ��� = "000000" Then
                    '���������
                    StrInput = StrInput & "1101011101" & "|"
                    StrInput = StrInput & g�������_����.ҽ����� & "|"
                    StrInput = StrInput & strDate & "|"
                    StrInput = StrInput & strDate & "|"
                Else
                    StrInput = StrInput & "1111111011" & "|"
                    StrInput = StrInput & g�������_����.ҽ����� & "|"
                    StrInput = StrInput & strDate & "|"
                    StrInput = StrInput & g�������_����.�������� & "|"
                    StrInput = StrInput & strDate & "|"
                    StrInput = StrInput & g�������_����.���ֱ��� & "|"
                End If
                StrInput = StrInput & gstrUserName & "|"
                StrInput = StrInput & g�������_����.�ʻ���� & "|"
                StrInput = StrInput & g�������_����.��̬��Ϣ
                DebugTool "���¾�����Ϣ:�ʻ����:" & g�������_����.�ʻ����
                 
            Case 1  '��Ժ�Ǽ�
                'סԺ��|���±�־|ҽ�����|��Ժ����|��Ժ��������|��Ժ����|ȷ�Ｒ������|������|�ʻ����|סԺ����|��̬��Ϣ
                '����:δ��
                
                '3�� ȷ�Ｒ������ : ����ָ���ı��룬������Բ��ã�����סԺ�����ṩ��Ч�ı��롣
                '4. ��Ժ�Ǽ�ʱ����'���¾�����Ϣ'ʱ��������'��Ժ����'����Ժ����ʱ��������'��Ժ����''��������'
                
                strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
                StrInput = g�������_����.סԺ�� & "|"
                
                If g�������_����.���ֱ��� = "" Or g�������_����.���ֱ��� = "000000" Then
                    '���������
                    StrInput = StrInput & "1100011111" & "|"
                    StrInput = StrInput & g�������_����.ҽ����� & "|"
                    StrInput = StrInput & str��Ժ���� & "|"
                Else
                    StrInput = StrInput & "1110111111" & "|"
                    StrInput = StrInput & g�������_����.ҽ����� & "|"
                    StrInput = StrInput & strDate & "|"
                    StrInput = StrInput & g�������_����.�������� & "|"
                    StrInput = StrInput & g�������_����.���ֱ��� & "|"
                End If
                StrInput = StrInput & gstrUserName & "|"
                StrInput = StrInput & g�������_����.�ʻ���� & "|"
                StrInput = StrInput & g�������_����.סԺ���� & "|"
                StrInput = StrInput & g�������_����.��̬��Ϣ
                
        Case 2      'סԺ����
    
            '���¾�����Ϣ(Ŀǰֻ�ĳ�Ժ����,ȷ�����,������)
                'סԺ��|���±�־|ҽ�����|��Ժ����|��Ժ��������|��Ժ����|ȷ�Ｒ������|������|�ʻ����|סԺ����|��̬��Ϣ
                '����:δ��
            StrInput = g�������_����.סԺ�� & "|"
            
            If bln������ Then
                    'ֻ���¶�̬��Ϣ�Ͳ���Ա
                    StrInput = StrInput & "0000011011" & "|"
            ElseIf g�������_����.���ֱ��� = "" Or g�������_����.���ֱ��� = "000000" Then
                StrInput = StrInput & "0001011011" & "|"
                StrInput = StrInput & str��Ժ���� & "|"
            Else
                StrInput = StrInput & "0001111011" & "|"
                StrInput = StrInput & str��Ժ���� & "|"
                StrInput = StrInput & g�������_����.���ֱ��� & "|"
            End If
            StrInput = StrInput & gstrUserName & "|"
            StrInput = StrInput & g�������_����.�ʻ���� & "|"
            StrInput = StrInput & g�������_����.��̬��Ϣ
        End Select
        
        If ҵ������_����(����_���¾�����Ϣ, StrInput, strOutPutstring) = False Then Exit Function

        '7�� ���¶�̬��Ϣ��־������ñ�־��ʾ��Ҫ���¿���̬��Ϣ������Ҫ����ʹ�ö�д������Կ��ϵĶ�̬��Ϣ���и��¡�
        '   1 ��Ҫ���¿���̬��Ϣ��ʹ�ý������Ķ�̬��Ϣֵ���¿��ϵĶ�̬��Ϣ
        '   0 ����Ҫ���¿���̬��Ϣ
        strArr = Split(strOutPutstring, "|")
        If Val(strArr(1)) = 1 Then
            '���������¶�̬��Ϣ
            'bytType:1-SiCardBaseInfo�籣��������Ϣ
            '        2-SiCardDynaInfo�籣����̬��Ϣ
            '        3-SiCardAcctInfo�籣���ʻ���Ϣ
            '        4-SiCardExtInfo�籣����չ��Ϣ
            strTemp = ""
            For i = 2 To UBound(strArr)
                strTemp = strTemp & "|" & strArr(i)
            Next
            strTemp = Mid(strTemp, 2)
            sCard_����ֵ 2, strTemp, True
            sCard_SaveCard
        
            '���±����ʻ��е��ʻ����
            If sCard_����ֵ(3, strTemp) = False Then Exit Function
            '��ȡ�����ʻ����
            '�ʻ����
            strArr = Split(strTemp, "|")
            With g�������_����
                .�ʻ���� = Val(strArr(0))
            End With
            
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & g�������_����.����ID & "," & TYPE_�ٲ׷��� & ",'�ʻ����','''" & g�������_����.�ʻ���� & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���潻����ˮ��")
            
        End If
        ���¾�����Ϣ_���� = True
        Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function



Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput  As String, str��ˮ�� As String
    Dim lng����ID As Long
    Dim strArr
    Dim rs��ϸ As New ADODB.Recordset
    Dim i As Long
    
    Dim dbl�����ܶ� As Double, dbl�����ʻ� As Double
    Dim dbl�ʻ������ۼ� As Currency, dbl�ʻ�֧���ۼ� As Currency
    Dim dbl����ͳ���ۼ� As Currency, dblͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim str��̬��Ϣ As String
    Dim curDate As Date
    Dim strҽ��֤�� As String
    
    On Error GoTo errHandle
    
    curDate = zlDatabase.Currentdate
    If Get������Ϣ(lng����ID) = False Then Exit Function
    strҽ��֤�� = g�������_����.ҽ��֤��
    If ��ȡ�α���Ա��Ϣ_����() = False Then Exit Function
    If strҽ��֤�� <> g�������_����.ҽ��֤�� Then
        ShowMsgbox "���������!"
        Exit Function
    End If
    
    If ���¾�����Ϣ_����(0, strOutput, , , True) = False Then
        Err.Raise 9000, gstrSysName, "���¾�����Ϣʧ��!"
        Exit Function
    End If
    
    gstrSQL = "Select * From ������ü�¼  " & _
        " Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������¼", lng����ID)
    
    Do Until rs��ϸ.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        dbl�����ܶ� = dbl�����ܶ� + Nvl(rs��ϸ("���ʽ��"), 0)
        rs��ϸ.MoveNext
    Loop
    dbl�����ܶ� = Round(dbl�����ܶ�, 2)
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rsTemp("����ID")
    
    
    
    gstrSQL = "Select * From ������ü�¼ " & _
        " Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������¼", lng����ID)
    Do While Not rsTemp.EOF
        rs��ϸ.Filter = 0
        rs��ϸ.Filter = "NO='" & Nvl(rsTemp!NO) & "' and ��¼����=" & Nvl(rsTemp!��¼����) & " and ���=" & Nvl(rsTemp!���)
        If rs��ϸ.EOF Then
            ShowMsgbox "������һ�����ϵĳ�����ϸδ�ҵ�!"
            Exit Function
        End If
        
        '�����ϴ���־
        gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(rsTemp!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rs��ϸ!ժҪ) & "')"
        zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭ���Ľ����¼", TYPE_�ٲ׷���, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    
    str��ˮ�� = rsTemp("֧��˳���")
    
    '���з�����
    '�����ض��������ݣ�  ��������|סԺ(����)��|���ݺ�|����Ա����
    '�����ض��������:  �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�x֧��|�󲡵渶|����16�̬��Ϣ|������ˮ��
    '�������Ͷ������£�
    '   1�������� (��Ժ����)
    '   0סԺ��;����
    '   -1������
    '   -2IC����ʧ���Ժ���㣬���ν��㣨ֻ���סԺ�������з���תΪ�ֽ�֧��������ҽ�����ı�����
    g�������_����.����� = Split(Nvl(rsTemp!��ע) & "|", "|")(0)
    StrInput = "-1|"
    StrInput = StrInput & g�������_����.����� & "|"
    StrInput = StrInput & g�������_����.����� & "|"
    StrInput = StrInput & gstrUserName & "|" & IIf(Nvl(rsTemp!ʵ������, 0) = 1, 1, 0)
    
    If ҵ������_����(����_����, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    strArr = Split(strOutput, "|")
    Dim dblͳ��֧�� As Double
    Dim dbl�󲡵渶 As Double
    
    dblͳ��֧�� = Val(strArr(2))
    dbl�����ʻ� = Val(strArr(3))
    dbl�󲡵渶 = Val(strArr(5))
    If Abs(dbl�����ʻ�) <> Abs(Nvl(rsTemp!�����ʻ�֧��, 0)) Then
        Err.Raise 9000, gstrSysName, "�������ĳ��ʵĸ����ʻ�֧���������ϴν���ĸ����ʻ�֧��!"
        Exit Function
    End If
    
    str��̬��Ϣ = ""
    '��ȡ��̬��Ϣ
    For i = 6 To UBound(strArr) - 1
        str��̬��Ϣ = str��̬��Ϣ & "|" & strArr(i)
    Next
    
    str��̬��Ϣ = Mid(str��̬��Ϣ, 2)
    
    '����д��
    '���ö�̬����
    If sCard_����ֵ(2, str��̬��Ϣ, True) = False Then
        GoTo Err����:
    End If
    
    'д��չ��Ϣ
    '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    
    '����
    
    If sCard_����ֵ(4, Get��չ��Ϣ("2"), True) = False Then
        GoTo Err����:
    End If
    str��ˮ�� = strArr(UBound(strArr))
    If sCard_SaveCard = False Then GoTo Err����:
    
    
    
    '�˴�����
    StrInput = g�������_����.����� & "|" & g�������_����.�����
    If ҵ������_����(����_�����˷�, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�ٲ׷���, lng����ID, Year(curDate), intסԺ�����ۼ�, dbl�ʻ������ۼ�, dbl�ʻ�֧���ۼ�, dbl����ͳ���ۼ�, dblͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�ٲ׷��� & "," & Year(curDate) & "," & _
        dbl�ʻ������ۼ� & "," & dbl�ʻ�֧���ۼ� - dbl�����ʻ� & "," & dbl����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        dblͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ������Ϣ")
    '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(�ʻ������ۼ�),�ʻ��ۼ�֧��_IN(�ʻ��ۼ�֧��),�ۼƽ���ͳ��_IN(�ۼƽ���ͳ��_IN),�ۼ�ͳ�ﱨ��_IN(�ۼ�ͳ�ﱨ��),סԺ����_IN(סԺ�����ۼ�),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(����ʱ������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�ٲ׷��� & "," & lng����ID & "," & Year(curDate) & "," & _
        dbl�ʻ������ۼ� & "," & dbl�ʻ�֧���ۼ� - dbl�����ʻ� & "," & dbl����ͳ���ۼ� & "," & dblͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0," & Nvl(rsTemp!ʵ������, 0) & "," & _
        dbl�����ܶ� * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0,0," & dbl�����ʻ� * -1 & ",'" & _
       str��ˮ�� & "',NULL,0,'" & g�������_����.����� & "|" & str��̬��Ϣ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���±��ս�����Ϣ")
    ����������_���� = True
    Exit Function
Err����:
    Call ҵ������_����(����_��������, str��ˮ�� & "|10|" & gstrUserName, strOutput)
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function ������Ժ�ǼǴ���(lng����ID As Long, lng��ҳID As Long) As Boolean
    '��������Ǽ�
    Dim StrInput As String, strOutput As String
    Dim str������ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    '�����ض��������ݣ�סԺ�������|ҽ��֤����|IC����|��Ժ����|��Ժ��������|������
    '�����ض��������:       ִ�гɹ�ʱ���ؽ�����ˮ�� , ִ��ʧ��ʱΪʧ��ԭ������
    gstrSQL = "Select C.סԺ��,C.��ǰ����,to_char(A.ȷ������,'yyyy-MM-dd') as ȷ������,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') ��Ժ����ʱ��," & _
        " to_char(A.��Ժ����,'yyyy-mm-dd') ��Ժ����  ,to_char(A.��Ժ����,'yyyy-mm-dd hh24:mi:ss') ��Ժʱ��,D.��Ժ��ϱ���,D.��Ժ�������,G.ȷ����ϱ���,g.ȷ��������� " & _
        " From ������ҳ A,���ű� B,������Ϣ C, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ��ϱ���,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ������� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by  ����id,��ҳid)   D," & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,2,b.����,'')) AS ȷ����ϱ���,max(DECODE(a.��ϴ���,2,b.����,'')) AS ȷ��������� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by  ����id,��ҳid)   g" & _
        " Where A.����id=C.����id and C.����id=" & lng����ID & _
        "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
        "       and A.��ҳid=g.��ҳid(+) and a.����id=g.����id(+) " & _
        ""
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ��Ϣ"
    
    With g�������_����
        .��Ժ��ϱ��� = Nvl(rsTemp!��Ժ��ϱ���)
        .��Ժ������� = Nvl(rsTemp!��Ժ�������)
        .ȷ����ϱ��� = Nvl(rsTemp!ȷ����ϱ���)
        .ȷ��������� = Nvl(rsTemp!ȷ���������)
        .סԺ���� = Nvl(rsTemp!��Ժ����)
        'סԺ�������|���˱��|IC����|ҽ�����|��Ժ����|��Ժ��������|����|������
        
        StrInput = .סԺ�� & "|"
        StrInput = StrInput & .ҽ��֤�� & "|"
        StrInput = StrInput & .���� & "|"
        StrInput = StrInput & .ҽ����� & "|"
        
        StrInput = StrInput & Nvl(rsTemp!��Ժ����) & "|"
        StrInput = StrInput & .�������� & "|"
        StrInput = StrInput & .סԺ���� & "|"
        'strInput = strInput & Nvl(rsTemp!��Ժ�������) & "|"
        StrInput = StrInput & Nvl(rsTemp!������, gstrUserName)
    End With
    
    Err = 0
    On Error GoTo errHand:
    
    If ҵ������_����(����_���˵Ǽ�, StrInput, strOutput) = False Then Exit Function
    str������ˮ�� = Split(strOutput, "|")(2)
      
    '���潫������ˮ��
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ٲ׷��� & ",'˳���','''" & str������ˮ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���潻����ˮ��")
    
    
    If ���¾�����Ϣ_����(1, strOutput, Nvl(rsTemp!��Ժ����)) = False Then
        GoTo Err����:
    End If
  
  
      
    'д��չ��Ϣ
    '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
  '  str��Ժ���� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If sCard_����ֵ(4, Get��չ��Ϣ("1", ""), True) = False Then
        GoTo Err����:
        Exit Function
    End If
    
    '��ı�סԺ״̬
    If sCard_SaveCard = False Then
        GoTo Err����:
        Exit Function
    End If
    
    ������Ժ�ǼǴ��� = True
    Exit Function
Err����:
        '������������ˮ��|�������������ʹ���|����Ա����
        StrInput = str������ˮ�� & "|"
        StrInput = StrInput & "01" & "|"
        StrInput = StrInput & gstrUserName
        If ҵ������_����(����_��������, StrInput, strOutput) = False Then
        End If
        Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
    Dim lng������� As Long
    Dim str������ˮ�� As String
    '��ȡסԺ��
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select nvl(�������,0)+1 as ������� From �����ʻ� where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ�Ǽ�"
    lng������� = Nvl(rsTemp!�������, 1)
    g�������_����.סԺ�� = lng����ID & "_" & lng�������
    
    '���±����ʻ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ٲ׷��� & ",'�������','" & lng������� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����������")
    
    '�Ƚ��еǼǴ���
    If ������Ժ�ǼǴ���(lng����ID, lng��ҳID) = False Then
        Exit Function
    End If
    
    '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ٲ׷��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function
Private Function Get���״���(ByVal intType As ҵ������_����, Optional bln������ As Boolean = False) As String
    Select Case intType
        Case ����_���˵Ǽ�
            Get���״��� = IIf(bln������, "���˵Ǽ�", "01")
        Case ����_��Ŀ���������ѯ
            Get���״��� = IIf(bln������, "��Ŀ���������ѯ", "02")
        Case ����_���ͨ����������Ϣ
            Get���״��� = IIf(bln������, "���ͨ����������Ϣ", "03")
        Case ����_�α��ʸ����
            Get���״��� = IIf(bln������, "�α��ʸ����", "04")
        Case ����_���¾�����Ϣ
            Get���״��� = IIf(bln������, "���¾�����Ϣ", "05")
        Case ����_¼�봦����ϸ
            Get���״��� = IIf(bln������, "¼�봦����ϸ", "06")
        Case ����_ȷ����ϸ��Ŀ����
            Get���״��� = IIf(bln������, "ȷ����ϸ��Ŀ����", "07")
        Case ����_�����˷�
            Get���״��� = IIf(bln������, "�����˷�", "08")
        Case ����_Ԥ����
            Get���״��� = IIf(bln������, "Ԥ����", "09")
        Case ����_����
            Get���״��� = IIf(bln������, "����", "10")
        Case ����_��òα���Ա��Ϣ
            Get���״��� = IIf(bln������, "��òα���Ա��Ϣ", "13")
        Case ����_��������
            Get���״��� = IIf(bln������, "��������", "99")
        Case ����_�շ�Ŀ¼���ش���
            Get���״��� = IIf(bln������, "�շ�Ŀ¼���ش���", "02")
        Case ����_�շ�Ŀ¼����Ԥ����
            Get���״��� = IIf(bln������, "�շ�Ŀ¼����Ԥ����", "01")
        Case ����_��ȡ��������
            Get���״��� = IIf(bln������, "��ȡ��������", "-1")
        Case ����_�������_��չ��Ϣ
            Get���״��� = IIf(bln������, "�������_��չ��Ϣ", "-1")
        Case Else
            Get���״��� = IIf(bln������, "����Ľ��״���", "-1")
    End Select
End Function
Public Function ҵ������_����(ByVal intType As ҵ������_����, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strOutput As String, strReturn As String
    Dim str���״��� As String
    Dim str�������� As String
    Dim i As Integer
    Dim strArr
    
    str���״��� = Get���״���(intType)
    StrInput = str���״��� & "|" & strInputString
    str�������� = str���״��� & "��" & Get���״���(intType, True)
    
    DebugTool "����ҵ��������(ҵ������Ϊ:" & str�������� & ") " & vbCrLf & ".....�������Ϊ" & Trim(StrInput)
    
    ҵ������_���� = False
    If InitInfor_����.ģ������ Then
        '��ȡģ������
        Readģ������ intType, StrInput, strOutPutstring
         ҵ������_���� = True
        Exit Function
    End If
    strOutput = Space(5000)
    Err = 0
    On Error GoTo errHand:
    
    Select Case intType
        Case ����_ҽ����ʼ��
            lngReturn = Init()
            If lngReturn <> 0 Then
                MsgBox "������ȷ���ó�ʼ��ҽ���ӿڡ�", vbInformation, gstrSysName
                Exit Function
            End If
        Case ����_�շ�Ŀ¼���ش���, ����_�շ�Ŀ¼����Ԥ����
            lngReturn = OperationQuery(StrInput, strOutput)
            '4�� ����0��ʾִ�гɹ�������-1��ʾִ��ʧ�ܣ�����100��ʾ������Ŀ�����ڡ�
            If lngReturn = -1 Then
                ShowMsgbox "����ʧ��!"
                Exit Function
            End If
            If lngReturn = 100 Then
                ShowMsgbox "������Ŀ������!"
                Exit Function
            End If
        Case Else
            '
            '��òα���Ա��Ϣ, ���˵Ǽ�, �α��ʸ����
            lngReturn = OperationAsk(StrInput, strOutput)
            strOutput = Trim(TruncZero(strOutput))
            strArr = Split(strOutput, "|")
            '���������ǰ6λ��ҵ��ִ�д��롣���ҵ��ɹ���ִ�д���Ϊ'     0'����һ��Ԫ���ǽ�����ˮ�ţ����ҵ��ʧ�ܣ�ҵ��ִ�д�������һ��Ԫ���ǳ�����Ϣ��
            If lngReturn <> 0 Then
                'ҵ�����ʧ��
                strReturn = "ҽ���ӿڳ��־��棺" & vbCrLf & strArr(0)
                ShowMsgbox strReturn
                Exit Function
            End If
    End Select
    strOutPutstring = "0|" & Trim(strOutput)
    ҵ������_���� = True
    DebugTool ".....�������Ϊ(�ɹ�):" & Trim(strOutPutstring)
     Exit Function
errHand:
    DebugTool ".....�������Ϊ(ʧ��):" & Trim(strOutPutstring)
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    '���˺�:20040923���ӵ�
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strҽ����  As String
    Dim str��Ժ���� As String
    
    Err = 0
    On Error GoTo errHand
    
    DebugTool "������Ժ�ǳ����ӿ�"
    
    ��Ժ�Ǽǳ���_���� = False
    
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ����ã����ܳ�����Ժ�Ǽ�"
        Exit Function
    End If
    
    If ��ȡ�α���Ա��Ϣ_���� = False Then Exit Function
    
    
    'д��չ��Ϣ
    '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    str��Ժ���� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If sCard_����ֵ(4, Get��չ��Ϣ("2", ""), True) = False Then
        Exit Function
    End If
    
    '��ı�סԺ״̬
    If sCard_SaveCard = False Then Exit Function
    
    
    
    '���ó�������
    gstrSQL = "Select ˳��� From �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ�Ǽ�ʱ�Ľ�����ˮ��"
    If rsTemp.EOF Then
        ShowMsgbox "��ҽ�����޴˲���!"
        Exit Function
    End If
    '�����ض��������ݣ�  ������������ˮ��|�������������ʹ���|����Ա����
    '�����ض��������:   δ��
    
    StrInput = Nvl(rsTemp!˳���) & "|"
    StrInput = StrInput & "01" & "|"
    StrInput = StrInput & gstrUserName
    If ҵ������_����(����_��������, StrInput, strOutput) = False Then Exit Function
    
    
    '����ҽ���ʻ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ٲ׷��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    
    DebugTool "ȡ���ɹ�"
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
  
    '�ı䵱ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ٲ׷��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function
Public Function ��Ժ�Ǽǳ���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '��Ժ�Ǽǳ���
     '�ı䲡��״̬
     If Not ����δ�����(lng����ID, lng��ҳID) Then
            ShowMsgbox "�ò����Ѿ���Ժ������,��Խ�����з�����!"
            Exit Function
     End If
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ٲ׷��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long, Optional strAdvance As String = "") As Boolean
  '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String, str����Ա As String, str��̬��Ϣ As String
    Dim lng��ҳID As Long, intMouse As Integer, intסԺ�����ۼ� As Integer, i As Integer
    Dim blnOld As Boolean
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim datCurr As Date
    Dim strArr

    If g�������_����.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣", vbInformation, gstrSysName
        Exit Function
    End If
        
    Err = 0: On Error GoTo errHand:
    Call DebugTool("����סԺ����")
    
    
    With g��������
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
        If IsNull(rsTemp("��ҳID")) = True Then
            Err.Raise 9000, gstrSysName, "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        lng��ҳID = rsTemp("��ҳID")
    End With
    
    
    
    
    '�ٴ�Ԥ�ᣬ��Ϊ���ܴ��ڼ����ϴ�����ϸ����
    '�����ض��������ݣ�   סԺ�������
    '�����ض��������:   �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�֧��|�󲡵渶
    
    Dim str���㷽ʽ  As String
    StrInput = g�������_����.סԺ�� & "|" & IIf(g�������_����.�¸����ʻ�, "1", "0")
    If ҵ������_����(����_Ԥ����, StrInput, strOutput) = False Then
        Exit Function
    End If
    strArr = Split(strOutput, "|")
    
    '�����������Ƿ�һ��
    With �����������
        If Round(.�����ܶ�, 2) <> Round(Val(strArr(1)), 2) Or Round(.ͳ��֧��, 2) <> Round(Val(strArr(2)), 2) Or _
            Round(.�˻�֧��, 2) <> Round(Val(strArr(3)), 2) Or Round(.�ֽ�֧��, 2) <> Round(Val(strArr(4)), 2) Or _
            Round(.�󲡵渶, 2) <> Round(Val(strArr(5)), 2) Then
            ShowMsgbox "���ν���ʱ��������㲻��,����..." & vbCrLf & _
                    "   �����ܶ�:" & Format(.�����ܶ�, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(1)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   ͳ��֧��:" & Format(.ͳ��֧��, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(2)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   �˻�֧��:" & Format(.�˻�֧��, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(3)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   �ֽ�֧��:" & Format(.�ֽ�֧��, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(4)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   �󲡵渶:" & Format(.�󲡵渶, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(5)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    ""
            Exit Function
        End If
    End With
    
    '��ʽ����
    '�����ض��������ݣ�  ��������|סԺ(����)��|���ݺ�|����Ա����|�ʻ����ѱ�־
    '�����ض��������:  �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�֧��|�󲡵渶|����16�̬��Ϣ|������ˮ��
    '�������Ͷ������£�
    '   1�������� (��Ժ����)
    '   0סԺ��;����
    '   -1������
    '   -2IC����ʧ���Ժ���㣬���ν��㣨ֻ���סԺ�������з���תΪ�ֽ�֧��������ҽ�����ı�����
    '   �˻����ѱ�־ 0 �����˻����ѣ�����־Ϊ0�����¸����ʻ����� 1  ʹ��ϵͳ��������ֵ������־Ϊ1���¸����ʻ���
    
    If g�������_����.��;���� = True Then
        StrInput = 0 & "|"
    Else
        StrInput = 1 & "|"
    End If
    StrInput = StrInput & g�������_����.סԺ�� & "|"
    StrInput = StrInput & lng����ID & "|"
    StrInput = StrInput & gstrUserName & "|"
    StrInput = StrInput & IIf(g�������_����.�¸����ʻ�, "1", "0")
    
    If ҵ������_����(����_����, StrInput, strOutput) = False Then
        Exit Function
    End If
        
    strArr = Split(strOutput, "|")
    str��̬��Ϣ = ""
    '��ȡ��̬��Ϣ
    For i = 6 To UBound(strArr) - 1
        str��̬��Ϣ = str��̬��Ϣ & "|" & strArr(i)
    Next
    
    str��̬��Ϣ = Mid(str��̬��Ϣ, 2)
    
    Dim objData As ��������
    With objData
        .�����ܶ� = Val(strArr(1))
        .ͳ��֧�� = Val(strArr(2))
        .�˻�֧�� = Val(strArr(3))
        .�ֽ�֧�� = Val(strArr(4))
        .�󲡵渶 = Val(strArr(5))
        .��̬��Ϣ = str��̬��Ϣ
        .������ˮ�� = strArr(UBound(strArr))
    End With
    
    '�����������Ƿ�һ��
    With �����������
        If Round(.�����ܶ�, 2) <> Round(objData.�����ܶ�, 2) Or Round(.ͳ��֧��, 2) <> Round(objData.ͳ��֧��, 2) Or _
            Round(.�˻�֧��, 2) <> Round(objData.�˻�֧��, 2) Or _
            Round(.�󲡵渶, 2) <> Round(objData.�󲡵渶, 2) Then
            
            str���㷽ʽ = "�����ʻ�|" & objData.�˻�֧��
            If objData.ͳ��֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & objData.ͳ��֧��
            If objData.�󲡵渶 <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�󲡵渶|" & objData.�󲡵渶
             
            '������صĽ�����Ϣ
            #If gverControl < 2 Then
                blnOld = True
                gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
            #Else
                strAdvance = str���㷽ʽ
                gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
            #End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
            
            intMouse = Screen.MousePointer
            Screen.MousePointer = 1
            '��ʾ�������
            If blnOld Then
                Call frm������Ϣ.ShowME(lng����ID)
            End If
            Screen.MousePointer = intMouse
        End If
    End With
    
    
    
    'д��.
    '   ��������˻���Ϊ�㣬���ݽ��㺯�����ص�"�˻�֧��"���ÿ�����SaleTrans�������ʲ�����
    '   ���ҵ��ÿ�����UpdateJrtclj, UpdateQfxzflj�Ϳ�����UpdateDynInfo�Կ���̬��Ϣ���и��²�����
    '   ע��һ��Ҫ���ÿ�����UpdateHospStatus�����˵�סԺ״̬����Ϊ2����Ժ״̬����
    '   �����ÿ�����AddHospTimes�����˵�סԺ������һ
    
    '??�����д������
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    
    '����д��
    '���ö�̬����
    If sCard_����ֵ(2, str��̬��Ϣ, True) = False Then
        GoTo Err����:
    End If
    
    'д��չ��Ϣ
    '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    If g�������_����.��;���� Then
    Else
        If sCard_����ֵ(4, Get��չ��Ϣ("2"), True) = False Then
            GoTo Err����:
        End If
    End If
    If sCard_SaveCard = False Then GoTo Err����:


    '��д�����
    Call DebugTool("��д�����¼")
    datCurr = zlDatabase.Currentdate
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�ٲ׷���, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�ٲ׷��� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & _
        cur����ͳ���ۼ� + objData.ͳ��֧�� & "," & _
        curͳ�ﱨ���ۼ� + objData.ͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ������Ϣ")
    
    
 
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(�ʻ������ۼ�),�ʻ��ۼ�֧��_IN(�ʻ��ۼ�֧��),�ۼƽ���ͳ��_IN(�ۼƽ���ͳ��_IN),�ۼ�ͳ�ﱨ��_IN(�ۼ�ͳ�ﱨ��),סԺ����_IN(סԺ�����ۼ�),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(����ʱ������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�ٲ׷��� & "," & lng����ID & "," & Year(datCurr) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0," & IIf(g�������_����.�¸����ʻ�, "1", "0") & "," & _
            g�������_����.�����ܶ� & ",0,0," & _
            objData.ͳ��֧�� & "," & objData.ͳ��֧�� & ",0,0," & objData.�˻�֧�� & ",'" & _
            objData.������ˮ�� & "'," & lng��ҳID & "," & IIf(g�������_����.��;����, 1, 0) & ",'" & g�������_����.סԺ�� & "'" & IIf(blnOld, "", ",1") & ")"
            
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
    
    סԺ����_���� = True
    Exit Function
Err����:
    Call ҵ������_����(����_��������, objData.������ˮ�� & "|10|" & gstrUserName, strOutput)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function סԺ�������_����(lng����ID As Long) As Boolean
     '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput  As String
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, i As Integer
    Dim curDate As Date
    Dim strArr
    Dim str��̬��Ϣ  As String, strҽ��֤�� As String
    Err = 0: On Error GoTo errHand:
    
    curDate = zlDatabase.Currentdate
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & TYPE_�ٲ׷��� & " and ��¼ID=" & lng����ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    
   
    '�жϲ��˵�סԺ���������Ƿ��������ϡ��жϱ�׼�Ǽ�鲡�����µ�סԺ��¼������У��Ͳ��ܽ�����
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    If Get������Ϣ(rsTemp("����id").Value) = False Then Exit Function
    
    strҽ��֤�� = g�������_����.ҽ��֤��
    If ��ȡ�α���Ա��Ϣ_����() = False Then Exit Function
    If strҽ��֤�� <> g�������_����.ҽ��֤�� Then
        Err.Raise 9000, gstrSysName, "���Ǹ�ҽ�����˵�ҽ��֤"
        Exit Function
    End If
     
    If ���¾�����Ϣ_����(2, strOutput, , , True) = False Then
        Err.Raise 9000, gstrSysName, "���¾�����Ϣʧ��!"
        Exit Function
    End If
    
    str��ˮ�� = rsTemp("֧��˳���")
    g�������_����.סԺ�� = Split(Nvl(rsTemp!��ע, "|"), "|")(0)
    g�������_����.�¸����ʻ� = IIf(Nvl(rsTemp!ʵ������, 0) = 1, True, False)
    g�������_����.��;���� = Nvl(rsTemp!��;����, 0) = 1
    
    '���з�����
    '�����ض��������ݣ�  ��������|סԺ(����)��|���ݺ�|����Ա����|�˻����ѱ�־
    '�����ض��������:  �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�֧��|�󲡵渶|����16�̬��Ϣ|������ˮ��
    '�������Ͷ������£�
    '   1�������� (��Ժ����)
    '   0סԺ��;����
    '   -1������
    '   -2IC����ʧ���Ժ���㣬���ν��㣨ֻ���סԺ�������з���תΪ�ֽ�֧��������ҽ�����ı�����
    '�˻����ѱ�־ 0 �����˻����ѣ�����־Ϊ0�����¸����ʻ����� 1  ʹ��ϵͳ��������ֵ������־Ϊ1���¸����ʻ�����
    StrInput = "-1|"
    StrInput = StrInput & g�������_����.סԺ�� & "|"
    StrInput = StrInput & lng����ID & "|"
    StrInput = StrInput & gstrUserName & "|"
    StrInput = StrInput & IIf(g�������_����.�¸����ʻ�, 1, 0)
    If ҵ������_����(����_����, StrInput, strOutput) = False Then
        Exit Function
    End If
    strArr = Split(strOutput, "|")
    str��̬��Ϣ = ""
    '��ȡ��̬��Ϣ
    For i = 6 To UBound(strArr) - 1
        str��̬��Ϣ = str��̬��Ϣ & "|" & strArr(i)
    Next
    
    str��̬��Ϣ = Mid(str��̬��Ϣ, 2)
    
    
    Dim objData As ��������
    With objData
        .�����ܶ� = Val(strArr(1))
        .ͳ��֧�� = Val(strArr(2))
        .�˻�֧�� = Val(strArr(3))
        .�ֽ�֧�� = Val(strArr(4))
        .�󲡵渶 = Val(strArr(5))
        .��̬��Ϣ = str��̬��Ϣ
        .������ˮ�� = strArr(UBound(strArr))
    End With
            
    'ȷ���Ƿ���Ҫд��
    'д��.
    '   ��������˻���Ϊ�㣬���ݽ��㺯�����ص�"�˻�֧��"���ÿ�����SaleTrans�������ʲ�����
    '   ���ҵ��ÿ�����UpdateJrtclj, UpdateQfxzflj�Ϳ�����UpdateDynInfo�Կ���̬��Ϣ���и��²�����
    '   ע��һ��Ҫ���ÿ�����UpdateHospStatus�����˵�סԺ״̬����Ϊ2����Ժ״̬����
    '   �����ÿ�����AddHospTimes�����˵�סԺ������һ
    
    '??�����д������
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    
    '����д��
    '���ö�̬����
    If sCard_����ֵ(2, str��̬��Ϣ, True) = False Then
        GoTo Err����:
    End If
    
    'д��չ��Ϣ
    '����ҽԺ1|����ҽԺ2|����ҽԺ3|����ҽԺ4|����ҽԺ5|��Ժ����|סԺ״̬��1��סԺ��2����Ժ��|����ҽԺ|ҽ�����
    'bytType:1-SiCardBaseInfo�籣��������Ϣ
    '        2-SiCardDynaInfo�籣����̬��Ϣ
    '        3-SiCardAcctInfo�籣���ʻ���Ϣ
    '        4-SiCardExtInfo�籣����չ��Ϣ
    If g�������_����.��;���� Then
    Else
        If sCard_����ֵ(4, Get��չ��Ϣ("1"), True) = False Then
            GoTo Err����:
        End If
    End If
    If sCard_SaveCard = False Then GoTo Err����:
    
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�ٲ׷���, rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & TYPE_�ٲ׷��� & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    

   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(�ʻ������ۼ�),�ʻ��ۼ�֧��_IN(�ʻ��ۼ�֧��),�ۼƽ���ͳ��_IN(�ۼƽ���ͳ��_IN),�ۼ�ͳ�ﱨ��_IN(�ۼ�ͳ�ﱨ��),סԺ����_IN(סԺ�����ۼ�),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(����ʱ������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�ٲ׷��� & "," & rsTemp("����ID") & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� & "," & curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0," & Nvl(rsTemp!ʵ������, 0) & "," & _
        Nvl(rsTemp("�������ý��"), 0) * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0,0," & _
        Nvl(rsTemp("�����ʻ�֧��"), 0) * -1 & ",'" & objData.������ˮ�� & "'," & rsTemp("��ҳID") & "," & Nvl(rsTemp("��;����"), 0) & ",'" & Nvl(rsTemp!��ע) & "'" & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")

    סԺ�������_���� = True
    Exit Function
Err����:
    Call ҵ������_����(����_��������, objData.������ˮ�� & "|10|" & gstrUserName, strOutput)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function Get��ˮ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng�շ�ϸĿID As Long, ByVal dbl���� As Double, ByVal dbl���� As Double, lng����ID As Long) As Variant
    '   ���ȡһ��������¼����ˮ�ţ����ڸ������ʣ�,���������������ۼ�����һ��.
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strArr  '0-������ˮ��,1-�������,2-������,3-ʵ�ʽ��׵���,4-ʵ�ʵȼ�....
    
    gstrSQL = " Select id, ժҪ From סԺ���ü�¼" & _
              " Where �շ�ϸĿID=[1] And ����ID=[2] And ��ҳID=[3]" & _
              " And ��¼״̬=1 And Nvl(�Ƿ��ϴ�,0)=1 and A.����*nvl(A.����,1)=" & dbl���� & " and Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4))=" & dbl���� & " And Nvl(ʵ�ս��,0)>0 And Rownum<2"
    
    DebugTool "�����ȡ��ˮ�ź���:GET��ˮ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ˮ��", lng�շ�ϸĿID, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        strTemp = "||||||"
        lng����ID = 0
    Else
        strTemp = Nvl(rsTemp!ժҪ, "|") & "||||||"
        lng����ID = rsTemp!ID
    End If
    strArr = Split(strTemp, "|")
    Get��ˮ�� = strArr
    DebugTool "������ȡ��ˮ�ź���:GET��ˮ�� ����ֵΪ:" & strTemp
End Function

Private Function �����ϴ�(ByVal lng��¼���� As Long, lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '������ϸ�ϴ�
    '����:�ϴ��²����ļ�����ϸ��ҽ������
    '����:  str���ݺ�   NO
    '       int����     ��¼����
    '       lng����ID  Ĭ��Ϊ0����ʾ�������ŵ��ݣ�����Ϊ������ָ�����˵ġ�����Ҫ����Ϊҽ���ڱ�����ʵ�ʱ���Ƿֲ������ύ���ݶ�����һ���ύ��
    '����:
    Dim rsTemp As New ADODB.Recordset, rs��ϸ As New ADODB.Recordset
    Dim StrInput As String, strOutput As String, strArr As Variant
    Dim str������ As String, str������ϸ��ˮ�� As String, str������� As String
    Dim lng����ID As Long
    Dim bln���� As Boolean
    
    �����ϴ� = False
    Err = 0
    On Error GoTo errHandle
    
   '�������ŵ��ݵķ�����ϸ
    gstrSQL = "Select A.ID,A.NO,A.���,A.����ID,A.��ҳID,to_char(A.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,C.��ע,J.��� as �շ����,C.�Ƿ�ҽ��,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,y.���� as ��������,A.����Ա����,B.���㵥λ,E.���,G.���� ����,M.ҽ����,M.������� " & _
              "  From סԺ���ü�¼ A,�շ���� J,�շ�ϸĿ B,�����ʻ� M,(Select  ��Ŀ����, ��ע,�Ƿ�ҽ��,�շ�ϸĿID From ����֧����Ŀ where ����=[3]) C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G,���ű� Y " & _
              "  where a.����id=M.����id  and M.����=[3] and A.NO=[1] and A.��¼����=[2] and A.��¼״̬=1 And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
              "        and A.�շ����=J.����(+)  and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=[3]" & _
              "        and A.�շ�ϸĿID=B.ID and a.��������id=y.id(+) and A.�շ�ϸĿID=C.�շ�ϸĿID(+) " & _
              "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
              "  Order by A.����ID,A.����ʱ��"
              
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ�ϴ�", str���ݺ�, lng��¼����, TYPE_�ٲ׷���)
    Dim lng����ID As Long
    
    '�ȼ���Ƿ�����˵������������ڣ����з��Ӧ�ļ�¼��.
    With rs��ϸ
        '�ϴ���ϸ
        
        bln���� = False
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 1 Then
            If InStr(1, "7�в�ҩ", Nvl(!�շ����)) <> 0 Then
                '���������Է�
                bln���� = True
            End If
        End If
        Do While Not .EOF
            '���ۼ��
            g�������_����.ҽ��֤�� = Nvl(!ҽ����)
            If �������ά������(Nvl(!�շ�ϸĿID, 0), Nvl(!�۸�, 0)) = False Then Exit Function
            
'            If Val(!����) < 0 Or Val(!�۸�) < 0 Then
'                '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
'                str������ϸ��ˮ�� = Get��ˮ��(!����ID, !��ҳID, !�շ�ϸĿID, Val(!����), Val(!�۸�), lng����ID)
'                If Trim(str������ϸ��ˮ��) = "" Or lng����ID = 0 Then
'                    MsgBox "û���ҵ����Գ����ļ�¼��[" & !���� & "]" & !����, vbInformation, gstrSysName
'                    Exit Function
'                End If
'            Else
'                str������ϸ��ˮ�� = ""
'            End If
            .MoveNext
        Loop
    End With
    
    If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
    Dim strArrժҪ
    
    '���з��ô���
    With rs��ϸ
        Do Until .EOF
            gstrSQL = "Select * from ҽ���շ�Ŀ¼ where ���=[1] and ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ϴ�", CLng(Val(Nvl(!��ע, 0))), CStr(Nvl(!��Ŀ����)))
'            If rsTemp.EOF Then
'                MsgBox "����Ŀδ����ҽ�����룬�����ϴ���ϸ!", vbInformation, gstrSysName
'                Exit Function
'            End If
            
'            If Val(!����) < 0 Or Val(!�۸�) < 0 Then
'                '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
'                'strArrժҪ :0������ˮ��,1�������,2סԺ(����)��,3������,4ʵ�ʽ��׵���,5ʵ�ʵȼ�,6-�Ƿ񱻳���"
'                strArrժҪ = Get��ˮ��(!����ID, !��ҳID, !�շ�ϸĿID, Val(!����), Val(!�۸�), lng����ID)
'                '�������Ӧ����ˮ������
'                '�����ض��������ݣ�  ������������ˮ��|�������������ʹ���|����Ա����
'                '�����ض��������:   δ��
'                strInput = strArrժҪ(0) & "|"
'                strInput = strInput & Get���״���(¼�봦����ϸ) & "|"
'                strInput = strInput & ToVarchar(Nvl(!����Ա����, gstrUserName), 20)
'                If ҵ������_����(����_��������, strInput, strOutput) = False Then Exit Function
'                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strArrժҪ(0) & "|" & strArrժҪ(1) & "|" & strArrժҪ(2) & "|" & strArrժҪ(3) & "|" & strArrժҪ(4) & "|" & strArrժҪ(5) & "|" & 0 & "')"
'                zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
'
'                '�����ü�¼�Ѿ�������
'                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & lng����ID & ",NULL,NULL,NULL,NULL,1,'" & strArrժҪ(0) & "|" & strArrժҪ(1) & "|" & strArrժҪ(2) & "|" & strArrժҪ(3) & "|" & strArrժҪ(4) & "|" & strArrժҪ(5) & "|" & "1" & "')"
'                zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
'            Else
                '�ϴ���ϸ��¼
                'ȷ���������
                If Nvl(!�Ƿ�ҽ��, 1) = 0 Then
                  '���˱��|��Ŀ���|��������
                    StrInput = Nvl(!ҽ����) & "|"
                    If Val(Nvl(!��ע)) = 1 Then
                        '��˵:ҩƷֻ�ܴ�ҽԺ����,������ֻ�ܵ�ֻ�ܴ�ҽ������
                        StrInput = StrInput & Nvl(!����, "9000900099") & "|"
                    Else
                        StrInput = StrInput & Nvl(!��Ŀ����, "9000900099") & "|"
                    End If
                        
                    StrInput = StrInput & Nvl(rs��ϸ!�Ǽ�ʱ��)
                    
                    If ҵ������_����(����_��Ŀ���������ѯ, StrInput, strOutput) = False Then
                        Exit Function
                    End If
                    '���ۼ��
                    strArr = Split(strOutput, "|")
                    str������� = strArr(1)
                Else
                    str������� = ""
                End If
                
                lng����ID = Nvl(!����ID, 0)
                g�������_����.סԺ�� = lng����ID & "_" & Nvl(!�������, 0)
                '�ϴ���ϸ
                '�����ض��������ݣ�סԺ(����)��|������|���������|�������|ҽԺ����|ҽ������|��Ŀ����|
                '                  ���õȼ�|�������|����|����|���|��λ|���|����|��������|��������|����ҽ��|¼���־
                
                '�����ض��������:   �ô�����Ӧ��Ŀ��ʵ�ʵ���|ʵ�ʵȼ�|������ˮ�š�
                
                str������ = rs��ϸ!NO & "_" & lng��¼����
                StrInput = Nvl(!����ID) & "_" & Nvl(!�������, 0) & "|"
                StrInput = StrInput & str������ & "|"
                StrInput = StrInput & Nvl(!���) & "|"
                StrInput = StrInput & str������� & "|"
                StrInput = StrInput & Nvl(!����) & "|"
                StrInput = StrInput & IIf(bln����, "9000900099", Nvl(!��Ŀ����, "9000900099")) & "|"
                StrInput = StrInput & Nvl(!����) & "|"
                If rsTemp.EOF Or bln���� Then
                    StrInput = StrInput & "3" & "|"
                    StrInput = StrInput & Split(Get�������(Nvl(!�շ����)), "-")(0) & "|"
                Else
                    StrInput = StrInput & Nvl(rsTemp!�շѵȼ�) & "|"
                    If IsNull(rsTemp!�շ����) Then
                        StrInput = StrInput & Split(Get�������(Nvl(!�շ����)), "-")(0) & "|"
                    Else
                        StrInput = StrInput & Nvl(rsTemp!�շ����) & "|"
                    End If
                End If
                
                StrInput = StrInput & Format(rs��ϸ("�۸�"), "0.0000") & "|"
                StrInput = StrInput & Format(rs��ϸ("����"), "0.00") & "|"
                StrInput = StrInput & Format(rs��ϸ("ʵ�ս��"), "#####0.00") & "|"         '���
                
                StrInput = StrInput & ToVarchar(rs��ϸ("���㵥λ"), 20) & "|"      '��λ
                StrInput = StrInput & ToVarchar(rs��ϸ("���"), 14) & "|"
                StrInput = StrInput & ToVarchar(rs��ϸ("����"), 20) & "|"
                StrInput = StrInput & Nvl(rs��ϸ!�Ǽ�ʱ��) & "|"
                
                StrInput = StrInput & Nvl(rs��ϸ!��������) & "|"
                StrInput = StrInput & Nvl(rs��ϸ!ҽ��) & "|"
            
                 '0 ��ʾ��ʼѭ����2 ��ʾ����ѭ�����ڽ���ѭ������ύ
                'If rs��ϸ.AbsolutePosition = 1 Then
                '    If rs��ϸ.AbsolutePosition = rs��ϸ.RecordCount Then
                        'ֻ��һ����¼ʱ
                        StrInput = StrInput & "1"
                '    Else
                '        StrInput = StrInput & 0
                '    End If
                'ElseIf rs��ϸ.AbsolutePosition = rs��ϸ.RecordCount Then
                '    StrInput = StrInput & 2
                'Else
                '    StrInput = StrInput & "1"
                'End If
                
                If ҵ������_����(����_¼�봦����ϸ, StrInput, strOutput) = False Then
                    Exit Function
                End If
                
                'ʵ�ʵ���|ʵ�ʵȼ�|������ˮ��
                strArr = Split(strOutput, "|")
                
                'ժҪ:������ˮ��|�������||סԺ(����)��|������(����:����ţ�סԺ:���ݺ�+��¼����)|ʵ�ʽ��׵���|ʵ�ʵȼ�
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strArr(3) & "|" & str������� & "|" & g�������_����.סԺ�� & "|" & str������ & "|" & strArr(1) & "|" & strArr(2) & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
'           End If
            .MoveNext
        Loop
    End With
        
    �����ϴ� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function
Public Function �����Ǽ�_����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str���ִ��� As String
    Dim dbl���� As Double, dbl��� As Double
    Dim StrInput As String, strOutput As String
    Dim str�Ƿ�ҩƷ  As String
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    
    
    �����Ǽ�_���� = False
    If lng��¼״̬ = 1 Then
        '��������
        If �����ϴ�(lng��¼����, lng��¼״̬, str���ݺ�) = False Then Exit Function
    Else
        '��������
        If ��������(lng��¼����, lng��¼״̬, str���ݺ�) = False Then Exit Function
    End If
        
    �����Ǽ�_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ��������(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ʴ�������,����¼״̬=2�ļ�¼
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsԭ��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim str������ As String, StrInput As String, strOutput As String, str������ˮ�� As String
    Dim strArr
    Dim lng����ID As Long
    Dim str�ѱ��������� As String
    �������� = False
   
    Err = 0: On Error GoTo errHand:
          
    '���õ��ݵ�ԭʼ�����Ƿ���ڸ���
    
    gstrSQL = " Select ժҪ,A.ID,a.�շ�ϸĿid,A.���,A.����*nvl(A.����,1) as ����,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4) as ���� " & _
              " From סԺ���ü�¼ A,�����ʻ� B " & _
              " where a.����id=b.����id and A.NO=[1] and A.��¼����=[2] and A.��¼״̬=3 and   Nvl(���ӱ�־,0)<>9  order by A.����id"
    Set rsԭ��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ�ϴ�", str���ݺ�, lng��¼����)
    If rsԭ��ϸ.EOF Then
        ShowMsgbox "�õ���û����Ӧ����ϸ��¼,��������!"
        Exit Function
    End If
    
    gstrSQL = " Select * " & _
              " From סԺ���ü�¼ A,�����ʻ� b" & _
              " where a.����id=b.����id and A.NO=[1] and A.��¼����=[2] and A.��¼״̬=2 and Nvl(���ӱ�־,0)<>9 AND nvl(a.�Ƿ��ϴ�,0)=0 "
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ�ϴ�", str���ݺ�, lng��¼����)
    
    lng����ID = 0
    '����ԭ���ݵ�ֵ
    With rs��ϸ
        'ժҪ��ֵ����Ϊ:"������ˮ��|�������|סԺ(����)��|������(����:����ţ�סԺ:���ݺ�+��¼����)|ʵ�ʽ��׵���|ʵ�ʵȼ�"
        Do While Not .EOF
            rsԭ��ϸ.Filter = "���=" & Nvl(!���, 0) & "  and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
            If rsԭ��ϸ.EOF Then
                ShowMsgbox "����ʱδ�ҵ���Ӧ�ļ�¼,����ʧ��!"
                Exit Function
            End If
            strArr = Split(Nvl(rsԭ��ϸ!ժҪ) & "|||||", "|")
            str������ˮ�� = strArr(0)
            If str������ˮ�� = "" Then
                ShowMsgbox "��ԭ���в����ڽ�����ˮ��,���ܼ�����"
                Exit Function
            End If
            
            '�����ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rsԭ��ϸ!ժҪ) & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        str�ѱ��������� = ""
        
        Do While Not .EOF
            rsԭ��ϸ.Filter = "���=" & Nvl(!���, 0) & "  and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
            '��Ҫ�Ǽ��ʱ�ʱ����ȷ�������˵��˷�
            '"������ˮ��|�������|סԺ(����)��|������(����:����ţ�סԺ:���ݺ�+��¼����)|ʵ�ʽ��׵���|ʵ�ʵȼ�"
            If lng����ID <> Nvl(!����ID, 0) Then
                '��Ҫ�Ǽ��ʱ�ʱ����ȷ�������˵��˷�
                '"������ˮ��|�������|סԺ(����)��|������(����:����ţ�סԺ:���ݺ�+��¼����)|ʵ�ʽ��׵���|ʵ�ʵȼ�"
                strArr = Split(Nvl(rsԭ��ϸ!ժҪ) & "|||||", "|")
                g�������_����.סԺ�� = strArr(2)
                str������ = strArr(3)
                StrInput = g�������_����.סԺ�� & "|"
                StrInput = StrInput & str������
                If ҵ������_����(����_�����˷�, StrInput, strOutput) = False Then Exit Function
                lng����ID = Nvl(!����ID, 0)
            End If
            .MoveNext
        Loop
    End With
    
 '   rsԭ��ϸ.Filter = "����<0 or �۸�<0 "
 '   If rsԭ��ϸ.EOF Then
        '�˴�����,ȫ����
        '   �����ض��������ݣ�   סԺ(����)��|������
        '   �����ض��������:               �˵��Ĵ�����ϸ�ļ�¼����
'        strInput = g�������_����.סԺ�� & "|"��
'        strInput = strInput & str������
'        If ҵ������_����(����_�����˷�, strInput, strOutput) = False Then Exit Function
'    Else
        'ԭ���ݴ��ڸ�������,����ó���
'        With rsԭ��ϸ
'            .Filter = 0
'            Do While Not .EOF
'                '�����ض��������ݣ� ������������ˮ��|�������������ʹ���|����Ա����
'                '�����ض��������:  δ��
'                'ժҪ:"������ˮ��|�������|סԺ(����)��|������(����:����ţ�סԺ:���ݺ�+��¼����)|ʵ�ʽ��׵���|ʵ�ʵȼ�|������־"
'
'                '����Ҹñ����������Ƿ��Ѿ�����������Ĵ�����(����)����.
'
'                strArr = Split(Nvl(!ժҪ, "|||||"), "|")
'                If Val(strArr(6)) = 1 Then
'                    '�����ü�¼�Ѿ�����������ĸ���¼���������ٳ�
'                Else
'                    strInput = strArr(0) & "|"
'                    strInput = Get���״���(¼�봦����ϸ)
'                    If ҵ������_����(����_��������, strInput, strOutput) = False Then Exit Function
'                End If
'                .MoveNext
'            Loop
'        End With
'    End If
    �������� = True
    Exit Function
errHand:
   If ErrCenter = 1 Then
        Resume
   End If
End Function
Private Function Readģ������(ByVal intҵ������ As ҵ������_����, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ���ù��ܶ�ȡģ������,�Ա����
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim STRNAME As String
    
    If intҵ������ = ��ȡ�������� Then
        strFile = App.Path & "\������.txt"
    Else
        strFile = App.Path & "\ģ���ύ��.txt"
    End If
    
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    STRNAME = Get���״���(intҵ������, True)
    
    Dim blnStart As Boolean
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If intҵ������ = ��ȡ�������� Then
                    strArr = Split(strText, vbTab)
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab
                            End If
                            strArr = Split(strText, vbTab)
                            
                            If Val(strArr(0)) = 1 Then
                                str = strArr(1)
                                Exit Do
                            End If
                        Else
                             If "<" & STRNAME & ">" = strText Then
                                 blnStart = True
                             End If
                        End If
                        If "</" & STRNAME & ">" = strText Then
                            Exit Do
                        End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function

Private Function Get������Ϣ(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strArr
    
    Get������Ϣ = False
    'COMMENT ON COLUMN �����ʻ�.����ID   is '����ID';
    'COMMENT ON COLUMN �����ʻ�.����     is '96';
    'COMMENT ON COLUMN �����ʻ�.����     is '0';
    'COMMENT ON COLUMN �����ʻ�.����     is '����';
    'COMMENT ON COLUMN �����ʻ�.ҽ����   is 'ҽ��֤���';
    'COMMENT ON COLUMN �����ʻ�.����     is 'ҽ�����';
    'COMMENT ON COLUMN �����ʻ�.��Ա��� is 'Ŀǰδ����';
    'COMMENT ON COLUMN �����ʻ�.��λ���� is '��λ����';
    'COMMENT ON COLUMN �����ʻ�.˳���   is '�Ǽ�ʱ�Ľ�����ˮ��';
    'COMMENT ON COLUMN �����ʻ�.����֤�� is '���ͨ����:�������|��Ŀ���|��Ŀ����';
    'COMMENT ON COLUMN �����ʻ�.�ʻ���� is '��ǰ�����ʻ����';
    'COMMENT ON COLUMN �����ʻ�.��ǰ״̬ is '0-����,1-��Ժ';
    'COMMENT ON COLUMN �����ʻ�.����ID   is '����ID���뱣�ղ��ֵ�ID����';
    'COMMENT ON COLUMN �����ʻ�.��ְ     is 'Ŀǰ�����ֵ��1�����ô�';
    'COMMENT ON COLUMN �����ʻ�.�����   is 'Ŀǰ�������ҽ����������';
    'COMMENT ON COLUMN �����ʻ�.�Ҷȼ�   is '�������';
    'COMMENT ON COLUMN �����ʻ�.����ʱ�� is '��ǰ�����ʱ��';
    'COMMENT ON COLUMN �����ʻ�.������� is '���˵�ǰ����Ĵ���,����ID-�����������Ŀǰ�������';
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "select a.*,b.����,b.�Ա�, b.����, b.��������, b.���֤��,b.������λ " & _
             " from �����ʻ� a,������Ϣ b " & _
             " WHERE a.����id=" & lng����ID & " AND a.����id=b.����id and a.����=" & TYPE_�ٲ׷���
 
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
    
    With g�������_����
        .���� = Nvl(rsTemp!����)
        .ҽ��֤�� = Nvl(rsTemp!ҽ����)
        
        .���� = Nvl(rsTemp!����)
        .�Ա� = Nvl(rsTemp!�Ա�)
        .���� = Nvl(rsTemp!�����, 0)
        .�������� = Format(rsTemp!��������, "yyyy-mm-dd")
        .��λ���� = Nvl(rsTemp!��λ����)
      
        strTemp = Nvl(rsTemp!������λ)
        If InStr(1, strTemp, "(") <> 0 Then
            .��λ���� = Mid(strTemp, 1, InStr(1, strTemp, "(") - 1)
        Else
            .��λ���� = strTemp
        End If
        
        .ҽ����� = Nvl(rsTemp!����)
        strArr = Split(Nvl(rsTemp!����֤��) & "||", "|")
        
        .������� = strArr(0)
        .��Ŀ���� = strArr(1)
        .��Ŀ���� = strArr(2)
        .�ʻ���� = Val(Nvl(rsTemp!�ʻ����))
        .������� = Nvl(rsTemp!�Ҷȼ�)
        .סԺ�� = lng����ID & "_" & Nvl(rsTemp!�������, 0)
        .����� = lng����ID & "_" & Nvl(rsTemp!�������, 0)
        .���֤�� = Nvl(rsTemp!���֤��)
        .����ID = Nvl(rsTemp!����ID, 0)
        
        If .����ID <> 0 Then
           gstrSQL = "Select ����,���� From ҽ������Ŀ¼ where id=" & .����ID
           OpenRecordset_���� rsTemp, "��ȡ����"
           
           If rsTemp.EOF Then
                .���ֱ��� = ""
                .�������� = ""
           Else
                .���ֱ��� = Nvl(rsTemp!����)
                .�������� = Nvl(rsTemp!����)
           End If
        Else
            .���ֱ��� = ""
            .�������� = ""
        End If
    End With
    Get������Ϣ = True
Exit Function
errHand:
        DebugTool "��ȡ������Ϣʧ��" & vbCrLf & " �����:" & Err.Number & vbCrLf & " ������Ϣ:" & Err.Description
End Function

Private Sub OpenRecordset_����(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional cnOracle As ADODB.Connection)
    '���ܣ��򿪼�¼��
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    If cnOracle Is Nothing Then
        rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_����, adOpenStatic, adLockReadOnly
    Else
        If cnOracle.State <> 1 Then
            rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_����, adOpenStatic, adLockReadOnly
        Else
            rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), cnOracle, adOpenStatic, adLockReadOnly
        End If
    End If
    Call SQLTest
End Sub


Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    
    'rsExse:�ַ���
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim lng��ҳID As Long
    Dim StrInput As String, strOutput   As String
    Dim strArr As Variant
    Dim cur�����ʻ� As Double, curͳ��֧�� As Double, cur���ͳ�� As Double, cur����Ա���� As Double, cur�������� As Double
    Dim str�ܽ��ҽԺ As String, str�ܽ��ҽ�� As String, str������ϸ��ˮ�� As String
    Dim strҽ�� As String, datCurr As Date, intMsg As Integer
    Dim str��Ժ���� As String, str��Ժ���� As String
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo errHand:
    
    g�������_����.����ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    If Get������Ϣ(lng����ID) = False Then Exit Function
    
    If bln���ʴ� Then
        Screen.MousePointer = 1
        If ��ݱ�ʶ_����(4, lng����ID) = "" Then
            Screen.MousePointer = intMouse
            סԺ�������_���� = ""
            Exit Function
        End If
        Screen.MousePointer = intMouse
    Else
        g�������_����.�¸����ʻ� = True
    End If
    
    rsExse.MoveFirst
    datCurr = zlDatabase.Currentdate
    With g��������
        .����ID = rsExse("����ID")
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", CLng(rsExse("����ID")))
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        lng��ҳID = rsTemp("��ҳID")
    End With

    Screen.MousePointer = vbHourglass
    
    '1.2 �������˵���Ժʱ��
    gstrSQL = "" & _
        "   Select ��Ժ����,��Ժ���� " & _
        "   From ������ҳ where ����ID=[1] and ��ҳID=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, lng��ҳID)
    If IsNull(rsTemp("��Ժ����")) Then
        g�������_����.��;���� = 1
        str��Ժ���� = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    Else
        '��ʾ�ò����Ѿ���Ժ
        g�������_����.��;���� = 0
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-mm-dd")
    End If
    '���
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-mm-dd")
    
    g�������_����.�����ܶ� = 0
    Do While Not rsExse.EOF
        g�������_����.�����ܶ� = g�������_����.�����ܶ� + rsExse("���")
        rsExse.MoveNext
    Loop
    g�������_����.�����ܶ� = Round(g�������_����.�����ܶ�, 2)
    
    '������ϸ
    If ����סԺ��ϸ��¼(lng����ID, lng��ҳID) = False Then Exit Function
    
    '���¾�����Ϣ
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1 and a.����id=" & lng����ID & " and a.��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    
    g�������_����.ȷ����ϱ��� = Nvl(rsTemp!ȷ�����)
    
    If ���¾�����Ϣ_����(2, strOutput, "", str��Ժ����) = False Then Exit Function
    
    
    
    '���ز���:���¶�̬��Ϣ��־|����16λ��̬��Ϣ��Ŀǰ������.��
    
    'Ԥ�ᴦ��
    '�����ض��������ݣ�   סԺ�������|�ʻ����ѱ�־
    '�����ض��������:   �����ܶ�|ͳ��֧��|�˻�֧��|�ֽ�֧��|�󲡵渶
    '3�� �˻����ѱ�־ 0 �����˻����ѣ�����־Ϊ0�����¸����ʻ����� 1  ʹ��ϵͳ��������ֵ������־Ϊ1���¸����ʻ���
    
    Dim str���㷽ʽ  As String
    StrInput = g�������_����.סԺ�� & "|"
    StrInput = StrInput & IIf(g�������_����.�¸����ʻ�, "1", "0")
    If ҵ������_����(����_Ԥ����, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    strArr = Split(strOutput, "|")
    
    With �����������
        .�����ܶ� = Val(strArr(1))
        .ͳ��֧�� = Val(strArr(2))
        .�˻�֧�� = Val(strArr(3))
        .�ֽ�֧�� = Val(strArr(4))
        .�󲡵渶 = Val(strArr(5))
        str���㷽ʽ = "�����ʻ�;" & .�˻�֧�� & ";0"   '�����޸ĸ����ʻ�����Ϊ����ʱ�Ѿ����ٴ���ǰ�û���
        If .ͳ��֧�� > 0 Then
            str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & .ͳ��֧�� & ";0"
        End If
        If .�󲡵渶 > 0 Then
            str���㷽ʽ = str���㷽ʽ & "|�󲡵渶;" & .�󲡵渶 & ";0"
        End If
        If .�����ܶ� <> g�������_����.�����ܶ� Then
            ShowMsgbox "���ķ��ý����ܶ�(" & .�����ܶ� & " ) ������ҽԺʵ�ʷ��������ܶ�(" & g�������_����.�����ܶ� & ")"
            Exit Function
        End If
    End With
    
    סԺ�������_���� = str���㷽ʽ
    g�������_����.����ID = lng����ID   '��ʾ�ò����Ѿ��������������
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get�������(ByVal str�շ���� As String) As String
    '��ȡ�������
    Dim rsTemp As New ADODB.Recordset
    Dim str�շ� As String
     str�շ� = str�շ����
    If zlCommFun.ActualLen(str�շ����) = 1 Then
            gstrSQL = "Select * From �շ���� where ����='" & str�շ���� & "'"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�շ����"
            
            If rsTemp.EOF Then
            Else
                str�շ� = Nvl(rsTemp!���)
            End If
    End If
    gstrSQL = "Select * From ���ղ��� where ������='" & str�շ� & "' and ����=" & TYPE_�ٲ׷���
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���"
    If rsTemp.EOF Then
        Get������� = ""
    Else
        Get������� = Nvl(rsTemp!����ֵ)
        
    End If
    If Get������� = "" Then
        Get������� = "-"
    End If
End Function
Private Function ����סԺ��ϸ��¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '���������ϸ��¼
    Dim cnTemp As New ADODB.Connection
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput  As String, strOutput As String
    Dim strArr, strArrժҪ
    Dim lng����ID As Long
    Dim str������ϸ��ˮ�� As String, str������� As String, str������ As String
    Err = 0
    On Error GoTo errHand:
      
      
    ����סԺ��ϸ��¼ = False
    
    '����δ�ϴ���ϸ�������Ա����ϴ�����ϸ�����ϴ�����ϸ��
    gstrSQL = "Select A.ID,A.NO,A.��¼����,J.��� as �շ����,A.��¼״̬,A.���,A.����ID,A.��ҳID,A.����ʱ�� as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս��" & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,C.��ע,C.�Ƿ�ҽ��,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,y.���� as ��������,A.����Ա����,B.���㵥λ,E.���,G.���� ���� " & _
              "  From סԺ���ü�¼ A,�շ���� J,�շ�ϸĿ B, (Select ��Ŀ����, ��ע,�Ƿ�ҽ��,�շ�ϸĿID From ����֧����Ŀ where ����=" & TYPE_�ٲ׷��� & ") C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G ,���ű� Y" & _
              "  where A.����ID=[1] and A.��ҳID=[2] and A.���ʷ���=1 and nvl(A.ʵ�ս��,0)<>0 and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 " & _
              "         and a.��������id=y.id(+) and A.�շ����=J.����(+) and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=[1]" & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID(+) " & _
              "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
              "  Order by A.����ʱ��,A.��¼����,Decode(A.��¼״̬,2,2,1)"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "�������", TYPE_�ٲ׷���, lng����ID, lng��ҳID)
    
    Call DebugTool("��������")
    Set cnTemp = GetNewConnection
    Call DebugTool("�����ӳɹ�����ʼ�����ϸ���ݵĺϷ��ԡ�")
    
    '�ȼ���Ƿ�����˵������������ڣ����з��Ӧ�ļ�¼��.
    With rs��ϸ
        '�ϴ���ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '���ۼ��
            If �������ά������(Nvl(!�շ�ϸĿID, 0), Nvl(!�۸�, 0)) = False Then Exit Function
            
'            If (Val(!����) < 0 Or Val(!�۸�)) < 0 And rs��ϸ!��¼״̬ = 1 Then
'                '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
'                str������ϸ��ˮ�� = Get��ˮ��(!����ID, !��ҳID, !�շ�ϸĿID, Val(!����), Val(!�۸�), lng����ID)
'                If Trim(str������ϸ��ˮ��) = "" Or lng����ID = 0 Then
'                    MsgBox "û���ҵ����Գ����ļ�¼��[" & !���� & "]" & !����, vbInformation, gstrSysName
'                    Exit Function
'                End If
'            Else
'                str������ϸ��ˮ�� = ""
'            End If
            .MoveNext
        Loop
    End With
    
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
           gstrSQL = "Select * from ҽ���շ�Ŀ¼ where ���=[1] and ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ϴ�", CLng(Val(Nvl(!��ע, 0))), CStr(Nvl(!��Ŀ����)))
            If rsTemp.EOF Then
'                MsgBox "����Ŀδ����ҽ�����룬�����ϴ���ϸ!", vbInformation, gstrSysName
'                Exit Function
            End If
                
            If rs��ϸ!��¼״̬ = 1 Then
'                If Val(rs��ϸ!����) < 0 Or Val(rs��ϸ!�۸�) < 0 Then
'                    '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
'                    'strArrժҪ :0������ˮ��,1�������,2סԺ(����)��,3������,4ʵ�ʽ��׵���,5ʵ�ʵȼ�,6-�Ƿ񱻳���"
'                    strArrժҪ = Get��ˮ��(!����ID, !��ҳID, !�շ�ϸĿID, Val(!����), Val(!�۸�), lng����ID)
'                    '�������Ӧ����ˮ������
'                    '�����ض��������ݣ�  ������������ˮ��|�������������ʹ���|����Ա����
'                    '�����ض��������:   δ��
'                    strInput = strArrժҪ(0) & "|"
'                    strInput = strInput & Get���״���(¼�봦����ϸ) & "|"
'                    strInput = strInput & ToVarchar(Nvl(!����Ա����, gstrUserName), 20)
'
'                    If ҵ������_����(����_��������, strInput, strOutput) = False Then Exit Function
'                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strArrժҪ(0) & "|" & strArrժҪ(1) & "|" & strArrժҪ(2) & "|" & strArrժҪ(3) & "|" & strArrժҪ(4) & "|" & strArrժҪ(5) & "|" & 0 & "')"
'                    cnTemp.Execute gstrSQL, , adCmdStoredProc
'
'                    '�����ü�¼�Ѿ�������
'                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & lng����ID & ",NULL,NULL,NULL,NULL,1,'" & strArrժҪ(0) & "|" & strArrժҪ(1) & "|" & strArrժҪ(2) & "|" & strArrժҪ(3) & "|" & strArrժҪ(4) & "|" & strArrժҪ(5) & "|" & "1" & "')"
'                    cnTemp.Execute gstrSQL, , adCmdStoredProc
'                Else
                    
                    '�����������ϴ�
                    'ȷ���������
                    If Nvl(!�Ƿ�ҽ��, 1) = 0 Then
                        StrInput = g�������_����.ҽ��֤�� & "|"
                        If Val(Nvl(!��ע)) = 1 Then
                            '��˵:ҩƷֻ�ܴ�ҽԺ����,������ֻ�ܵ�ֻ�ܴ�ҽ������
                            StrInput = StrInput & Nvl(!����, "9000900099") & "|"
                        Else
                            StrInput = StrInput & Nvl(!��Ŀ����, "9000900099") & "|"
                        End If
                        StrInput = StrInput & Nvl(rs��ϸ!�Ǽ�ʱ��)
                        
                        If ҵ������_����(����_��Ŀ���������ѯ, StrInput, strOutput) = False Then
                            Exit Function
                        End If
                        strArr = Split(strOutput, "|")
                        str������� = strArr(1)
                    Else
                        str������� = ""
                    End If
                    '�ϴ���ϸ
                    '�����ض��������ݣ�סԺ(����)��|������|���������|�������|ҽԺ����|ҽ������|��Ŀ����|
                    '                  ���õȼ�|�������|����|����|���|��λ|���|����|��������|��������|����ҽ��|¼���־
                    
                    '�����ض��������:   �ô�����Ӧ��Ŀ��ʵ�ʵ���|ʵ�ʵȼ�|������ˮ�š�
                    
                    str������ = rs��ϸ!NO & "_" & Nvl(!��¼����)
                
                    StrInput = g�������_����.סԺ�� & "|"
                    StrInput = StrInput & str������ & "|"
                    StrInput = StrInput & Nvl(!���) & "|"
                    StrInput = StrInput & str������� & "|"
                    StrInput = StrInput & Nvl(!����) & "|"
                    StrInput = StrInput & Nvl(!��Ŀ����, "9000900099") & "|"
                    StrInput = StrInput & Nvl(!����) & "|"
                    If rsTemp.EOF Then
                        StrInput = StrInput & "3" & "|"
                        StrInput = StrInput & Split(Get�������(Nvl(!�շ����)), "-")(0) & "|"
                    Else
                        StrInput = StrInput & Nvl(rsTemp!�շѵȼ�) & "|"
                            If IsNull(rsTemp!�շ����) Then
                                StrInput = StrInput & Split(Get�������(Nvl(!�շ����)), "-")(0) & "|"
                            Else
                                StrInput = StrInput & Nvl(rsTemp!�շ����) & "|"
                            End If
                    End If
                    StrInput = StrInput & Format(rs��ϸ("�۸�"), "0.0000") & "|"
                    StrInput = StrInput & Format(rs��ϸ("����"), "0.00") & "|"
                    StrInput = StrInput & Format(rs��ϸ("ʵ�ս��"), "#####0.0000") & "|"         '���
                
                    StrInput = StrInput & ToVarchar(rs��ϸ("���㵥λ"), 20) & "|"      '��λ
                    StrInput = StrInput & ToVarchar(rs��ϸ("���"), 14) & "|"
                    StrInput = StrInput & ToVarchar(rs��ϸ("����"), 20) & "|"
                    StrInput = StrInput & Nvl(rs��ϸ!�Ǽ�ʱ��) & "|"
                    StrInput = StrInput & Nvl(rs��ϸ!��������) & "|"
                    StrInput = StrInput & Nvl(rs��ϸ!ҽ��) & "|"
                    StrInput = StrInput & 1
                    
                    If ҵ������_����(����_¼�봦����ϸ, StrInput, strOutput) = False Then
                        Exit Function
                    End If
                    'ʵ�ʵ���|ʵ�ʵȼ�|������ˮ��
                    strArr = Split(strOutput, "|")
                    'ժҪ:������ˮ��|�������||סԺ(����)��|������(����:����ţ�סԺ:���ݺ�+��¼����)|ʵ�ʽ��׵���|ʵ�ʵȼ�
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strArr(3) & "|" & str������� & "|" & g�������_����.סԺ�� & "|" & str������ & "|" & strArr(1) & "|" & strArr(2) & "')"
                    cnTemp.Execute gstrSQL, , adCmdStoredProc
'                End If
            Else
                '���ϴ��������ļ�¼,�����
                'strArrժҪ :0������ˮ��,1�������,2סԺ(����)��,3������,4ʵ�ʽ��׵���,5ʵ�ʵȼ�,6-�Ƿ񱻳���"
                strArrժҪ = Getԭ����ժҪ(Nvl(!NO), Nvl(!���), Nvl(!��¼����, 0))
                '�������Ӧ����ˮ������
                '�����ض��������ݣ�  ������������ˮ��|�������������ʹ���|����Ա����
                '�����ض��������:   δ��
                If Val(strArr(6)) = 1 Then
                        '�����ü�¼�Ѿ�����������ĸ���¼���������ٳ�
                Else
                    StrInput = strArrժҪ(0) & "|"
                    StrInput = StrInput & Get���״���(¼�봦����ϸ) & "|"
                    StrInput = StrInput & ToVarchar(Nvl(!����Ա����, gstrUserName), 20)
                    If ҵ������_����(����_��������, StrInput, strOutput) = False Then Exit Function
                End If
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strArrժҪ(0) & "|" & strArrժҪ(1) & "|" & strArrժҪ(2) & "|" & strArrժҪ(3) & "|" & strArrժҪ(4) & "|" & strArrժҪ(5) & "|" & 0 & "')"
                cnTemp.Execute gstrSQL, , adCmdStoredProc
            End If
            .MoveNext
        Loop
    End With
    
    ����סԺ��ϸ��¼ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Getԭ����ժҪ(ByVal strNO As String, ByVal int��� As Integer, ByVal int���� As Integer) As Variant
    '����ָ����ֵ����ȡժҪ�������Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
        
    
    gstrSQL = " Select ժҪ From סԺ���ü�¼" & _
              " Where NO=[1] And ���=[2] And ��¼����=[3] And ��¼״̬=3"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡԭʼ������ϸ����ˮ��", strNO, int���, int����)
    
    If Not rsTemp.EOF Then
        strTemp = Nvl(rsTemp!ժҪ) & "|||||||"
    Else
        strTemp = "|||||||"
    End If
    Getԭ����ժҪ = Split(strTemp, "|")
End Function

'----200410���˺����
Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.��������
    
End Function
'
'Public Function ���ط�����ĿĿ¼_����(ByVal bytType As Byte, ByVal objProgss As Object) As Boolean
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    '����:���ط�����ĿĿ¼
'    '����:bytType-1-ҩƷ,2-����,3-����,4-�������,5-����Ŀ¼
'    '����:���سɹ�,����true,���򷵻�False
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSql As String
'    Dim rsTemp As New ADODB.Recordset
'    Dim strDate As String, strInput As String, strOutput As String
'    Dim lngCount As Long
'    Dim i As Long
'    Dim strArr
'    Dim strTitle As String
'
'    ���ط�����ĿĿ¼_���� = False
'    strTitle = Switch(bytType = 1, "ҩƷ", bytType = 2, "������Ŀ", bytType = 3, "������ʩ", bytType = 4, "�������", True, "����Ŀ¼ȡ")
'
'    Err = 0
'    On Error GoTo ErrHand:
'    strSql = "" & _
'        "   Select to_char(Max(���ʱ��),'yyyy-mm-dd hh24:mi:ss')  as ���ʱ�� " & _
'        "   From ҽ���շ�Ŀ¼ " & _
'        "   where ���=" & bytType
'    zlDatabase.OpenRecordset rsTemp, strSql, "��ȡ�����ʱ��"
'
'    strDate = Nvl(rsTemp!���ʱ��)
'    strDate = IIf(strDate = "", "1977-01-01 00:00:00", strDate)
'
'    If Not objProgss Is Nothing Then
'    Else
'        zlCommFun.ShowFlash "�������ء�" & strTitle & "������,��ȴ�..."
'    End If
'    'Ԥ����
'    strInput = bytType & "|" & strDate
'    If ҵ������_����(����_�շ�Ŀ¼����Ԥ����, strInput, strOutput) = False Then Exit Function
'    strArr = Split(strOutput, "|")
'    lngCount = Val(strArr(1))
'
'    If Not objProgss Is Nothing Then
'        objProgss.Max = IIf(lngCount = 0, 1, lngCount)
'        objProgss.Min = 1
'        objProgss.Value = 1
'    End If
'
'   For i = 1 To lngCount
'        '��������
'        If ҵ������_����(����_�շ�Ŀ¼���ش���, strInput, strOutput) = False Then Exit Function
'        strArr = Split(strOutput, "|")
'        '�����շ�Ŀ¼
'
'        '����:���,����,����,Ӣ������,�շ����,�շѵȼ�,���õȼ�,ƴ����,��λ,����,���,��ע,���ʱ��,��ά����־,֧����׼
'        strSql = "ZL_ҽ���շ�Ŀ¼_UPDATE("
'        strSql = strSql & bytType & ",'"
'        strSql = strSql & strArr(1) & "','" '����
'        strSql = strSql & strArr(2) & "','" '����
'        Select Case bytType
'        Case 1
'            strSql = strSql & strArr(3) & "','" 'Ӣ������
'            strSql = strSql & strArr(4) & "','" '�շ����
'            strSql = strSql & strArr(5) & "','" '���õȼ�
'            strSql = strSql & strArr(6) & "','" 'ƴ����
'            strSql = strSql & strArr(7) & "','" '��λ
'            strSql = strSql & strArr(8) & "','" '����
'            strSql = strSql & strArr(9) & "','" '����
'            strSql = strSql & strArr(10) & "','" '���
'            strSql = strSql & strArr(11) & "',to_date('" '��ע
'            strSql = strSql & strArr(12) & "','yyyy-mm-dd hh24:mi:ss'),'"  '���ʱ��
'            strSql = strSql & strArr(13) & "','"     '��ά����־
'            strSql = strSql & "" & "')" '֧����׼
'        Case 2
'            strSql = strSql & "" & "','" 'Ӣ������
'            strSql = strSql & strArr(3) & "','" '�շ����
'            strSql = strSql & "" & "','" '���õȼ�
'            strSql = strSql & strArr(4) & "','" 'ƴ����
'            strSql = strSql & strArr(5) & "','" '��λ
'            strSql = strSql & strArr(6) & "','" '����
'            strSql = strSql & "" & "','" '����
'            strSql = strSql & "" & "','" '���
'            strSql = strSql & strArr(7) & "',to_date('" '��ע
'            strSql = strSql & strArr(8) & "','yyyy-mm-dd hh24:mi:ss'),'"  '���ʱ��
'            strSql = strSql & strArr(9) & "','"     '��ά����־
'            strSql = strSql & "" & "')" '֧����׼
'        Case 3
'            strSql = strSql & "" & "','" 'Ӣ������
'            strSql = strSql & strArr(3) & "','" '�շ����
'            strSql = strSql & "" & "','" '���õȼ�
'            strSql = strSql & strArr(6) & "','" 'ƴ����
'            strSql = strSql & "" & "','" '��λ
'            strSql = strSql & strArr(4) & "','" '����
'            strSql = strSql & "" & "','" '����
'            strSql = strSql & "" & "','" '���
'            strSql = strSql & "" & "',to_date('" '��ע
'            strSql = strSql & strArr(7) & "','yyyy-mm-dd hh24:mi:ss'),'"  '���ʱ��
'            strSql = strSql & "" & "',"     '��ά����־
'            strSql = strSql & strArr(5) & "')" '֧����׼
'        Case 4
'            ' ����������|�����������
'            strSql = "ZL_ҽ���շ����_UPDATE("
'
'            strSql = strSql & strArr(1) & "','" '����
'            strSql = strSql & strArr(2) & "')" '����
'        Case Else
'            '���ֱ���|��������|ƴ����|�������
'            strSql = "ZL_ҽ������Ŀ¼_UPDATE("
'            strSql = strSql & strArr(1) & "','" '����
'            strSql = strSql & strArr(2) & "','" '����
'            strSql = strSql & strArr(3) & "',to_date('" '������
'            strSql = strSql & strArr(4) & "','yyyy-mm-dd hh24:mi:ss')" '���ʱ��
'        End Select
'        gcnOracle_����.Execute strSql, , adCmdStoredProc
'        If Not objProgss Is Nothing Then
'            objProgss.Value = i
'        Else
'            zlCommFun.ShowFlash "�������ء�" & strTitle & "������,������" & i & "/" & lngCount & ""
'        End If
'   Next
'   ���ط�����ĿĿ¼_���� = True
'   Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'End Function






Public Function ���ط�����ĿĿ¼_����(ByVal bytType As Byte, ByVal objProgss As Object) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�����ĿĿ¼
    '����:bytType-1-ҩƷ,2-����,3-����,4-�������,5-����Ŀ¼
    '����:���سɹ�,����true,���򷵻�False
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Dim strDate As String, StrInput As String, strOutput As String
    Dim lngCount As Long
    Dim i As Long
    Dim strArr
    Dim strTitle As String
    
    ���ط�����ĿĿ¼_���� = False
    

    If gcnOracle_���� Is Nothing Then
        If Open�м�� = False Then Exit Function
    End If
    If gcnOracle_����.State <> 1 Then
        If Open�м�� = False Then Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand:
    
    strSQL = "" & _
        "   Select to_char(Max(���ʱ��),'yyyy-mm-dd hh24:mi:ss')  as ���ʱ�� " & _
        "   From ҽ���շ�Ŀ¼ " & _
        "   where ���=" & bytType
    zlDatabase.OpenRecordset rsTemp, strSQL, "��ȡ�����ʱ��"
    
    strDate = Nvl(rsTemp!���ʱ��)
    strDate = IIf(strDate = "", "1477-01-01 00:00:00", strDate)
       
    If Not objProgss Is Nothing Then
    Else
        zlCommFun.ShowFlash "�������ء�" & strTitle & "������,��ȴ�..."
    End If
    
    Select Case bytType
    Case 1      'ҩƷ
        strTitle = "ҩƷ"
        gstrSQL = "Select * From medicine_info where AAE035>to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')"
    Case 2      '������Ŀ
        strTitle = "������Ŀ"
        gstrSQL = "Select * From examine_info where AAE035>to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')"
    Case 3      '������ʩ
        strTitle = "������ʩ"
        gstrSQL = "Select * From equipment_info where AAE035>to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')"
    
    Case 4      '�������
        strTitle = "�������"
        gstrSQL = "Select * From CHARGETYPE_INFO "
    Case 5      '����Ŀ¼
        strTitle = "����Ŀ¼"
        gstrSQL = "Select * From illness_info where AAE035>to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')"
    End Select
    
    
    OpenRecordset_���� rsData, "��ȡ" & strTitle & "����", , gcnOracle_����
            
    If Not objProgss Is Nothing Then
        objProgss.Max = IIf(rsData.RecordCount = 0, 1, rsData.RecordCount) + 1
        objProgss.Min = 1
        objProgss.Value = 1
    End If
    i = 1
    With rsData
        Do While Not .EOF
            
            strSQL = "ZL_ҽ���շ�Ŀ¼_UPDATE("
            '    ���_IN IN ҽ���շ�Ŀ¼.���%TYPE,
            strSQL = strSQL & bytType & ","
            Select Case bytType
            Case 1
                    '    ����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                    '    AKA060  VARCHAR2(60)    N   ҩƷ����    ʡĿ¼���
                    
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA060)) & "',"    'AKA060  VARCHAR2(20)                   ҩƷ����
                    '    ����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                    '    AKA061  VARCHAR2(100)   Y   ��������    ͨ����
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA061)) & "',"   'AKA061 VARCHAR2(50)  Y                ��������
                    '    AKA062  VARCHAR2(50)    Y   Ӣ������
                    '    Ӣ������_IN IN ҽ���շ�Ŀ¼.Ӣ������%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA062)) & "',"  'AKA062 VARCHAR2(50)  Y                Ӣ������
                    '    AKA063  VARCHAR2(3) Y   �շ����    11��ҩ�� 12��ҩ 13��ҩ
                    '    �շ����_IN IN ҽ���շ�Ŀ¼.�շ����%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA063)) & "'," 'AKA063 VARCHAR2(3)   Y                �շ����
                    '    AKA065  VARCHAR2(3) Y   �շ���Ŀ�ȼ�    1���� 2���� 3�Է�
                    '    �շѵȼ�_IN IN ҽ���շ�Ŀ¼.�շѵȼ�%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA065)) & "'," 'AKA065 VARCHAR2(3)   Y                �շ���Ŀ�ȼ�
                    '    AKA066  VARCHAR2(30)    Y   ������
                    '    ������_IN IN ҽ���շ�Ŀ¼.������%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA066)) & "',"  'AKA066 VARCHAR2(14)  Y                ������
                    '    AKA067  VARCHAR2(20)    Y   ��λ
                    '    ��λ_IN IN ҽ���շ�Ŀ¼.��λ%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA067)) & "',"  'AKA067 VARCHAR2(20)  Y                ��λ
                    '    AKA068  NUMBER(8,2) Y   ��׼�۸�
                    '    ��׼�۸�_IN IN ҽ���շ�Ŀ¼.��׼�۸�%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!AKA068)) & ","  'AKA068 NUMBER(8,2)   Y                ��׼�۸�
                    '    AKA070  VARCHAR2(50)    Y   ����
                    '    ����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA070)) & "',"   'AKA070 VARCHAR2(50)  Y                ����
                    '    AKA074  VARCHAR2(50)    Y   ���
                    '    ���_IN IN ҽ���շ�Ŀ¼.���%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA074)) & "',"    'AKA074 VARCHAR2(50)  Y                ���
                    '    AAE013  VARCHAR2(100)   Y   ��ע
                    '    ��ע_IN IN ҽ���շ�Ŀ¼.��ע%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AAE013)) & "',"      'AAE013 VARCHAR2(100) Y                ��ע
                    '    AAE035  DATE    Y   �������
                    '    ���ʱ��_IN IN ҽ���շ�Ŀ¼.���ʱ��%TYPE,
                strSQL = strSQL & "to_date('" & Format(!AAE035, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                    '    AAA104  VARCHAR2(3) Y   �����ά����־
                    '    ά����־_IN IN ҽ���շ�Ŀ¼.ά����־%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AAA104, "1")) & "',"    'AAA104 VARCHAR2(3)   Y                �����ά����־
                    '    ֧����׼_IN IN ҽ���շ�Ŀ¼.֧����׼%TYPE,
                strSQL = strSQL & "NULL,"  'AAA104 VARCHAR2(3)   Y                ֧����׼
                    '    ҩƷ����_IN IN ҽ���շ�Ŀ¼.ҩƷ����%TYPE,
                    '    AKA305  VARCHAR2(3) Y   ҩƷ����
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA305)) & "',"      'AKA305 VARCHAR2(3)   Y                ҩƷ����
                    '    AKA064  VARCHAR2(3) Y   ����ҩ��־
                    '    ����ҩ��־_IN IN ҽ���շ�Ŀ¼.����ҩ��־%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA064)) & "',"      'AKA064 VARCHAR2(3)   Y                ����ҩ��־
                    '    AKA069  NUMBER(5,4) Y   �Ը�����
                    '    �Ը�����_IN IN ҽ���շ�Ŀ¼.�Ը�����%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!AKA069)) & ","       'AKA069 NUMBER(5,4)   Y                �Ը�����
                    
                    '    AKA071  NUMBER(5,2) Y   ÿ������
                    '    ÿ������_IN IN ҽ���շ�Ŀ¼.ÿ������%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!AKA071)) & ","      'AKA071 NUMBER(5,2)   Y                ÿ������
                    '    AKA072  VARCHAR2(20)    Y   ʹ��Ƶ��
                    '    ʹ��Ƶ��_IN IN ҽ���շ�Ŀ¼.ʹ��Ƶ��%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA072)) & "',"      'AKA072 VARCHAR2(20)  Y                ʹ��Ƶ��
                    '    AKA073  VARCHAR2(50)    Y   �÷�
                    '    �÷�_IN IN ҽ���շ�Ŀ¼.�÷�%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA073)) & "',"      'AKA073 VARCHAR2(50)  Y                �÷�
                    '    AKA030  VARCHAR2(3) Y   ��ǰʹ�ñ�־
                    '    ��ǰʹ�ñ�־_IN IN ҽ���շ�Ŀ¼.��ǰʹ�ñ�־%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA030)) & "',"      'AKA030 VARCHAR2(3)   Y                ��ǰʹ�ñ�־
                    '    �籣�������_IN IN ҽ���շ�Ŀ¼.�籣�������%TYPE,
                strSQL = strSQL & "NULL,"      'AAB034 VARCHAR2(14)  Y                �籣�������
                    '    ҽԺ�ȼ�_IN IN ҽ���շ�Ŀ¼.ҽԺ�ȼ�%TYPE,
                strSQL = strSQL & "NUll,"
                '   סԺ�Էѱ���_IN IN ҽ���շ�Ŀ¼.סԺ�Էѱ���%TYPE,
                strSQL = strSQL & "NUll,"
                '    �α���Ա_IN IN ҽ���շ�Ŀ¼.�α���Ա%TYPE,
                strSQL = strSQL & "NUll,"
                
                '    AKA075  VARCHAR2(40)    Y   ͨ��������
                    '    ͨ��������_IN   ҽ���շ�Ŀ¼.ͨ��������%type:=NULL,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA075)) & "',"
                '    AKA076  VARCHAR2(100)   Y   ��Ʒ��
                    '    ��Ʒ��_IN   ҽ���շ�Ŀ¼.��Ʒ��%type:=NULL,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA076)) & "',"
                '    AKA078  VARCHAR2(20)    Y   С��λ
                '    С��λ_IN   ҽ���շ�Ŀ¼.С��λ%type:=NULL,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA078)) & "',"
                '    AKA079  NUMBER(8,2) Y   ��С��λ���ۼ�
                    '    ��С��λ�ۼ�_IN ҽ���շ�Ŀ¼.��С��λ�ۼ�%type:=NULL,
                strSQL = strSQL & "" & Val(Nvl(!AKA079)) & ","
                '    AKA031  VARCHAR2(100)   Y   ����
                    '    ����_IN     ҽ���շ�Ŀ¼.����%type:=NULL,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA031)) & "',"
                '    AKA032  VARCHAR2(1) Y   �¾�Ŀ¼��־    0ԭĿ¼��1ʡĿ¼
                '    �¾�Ŀ¼��־_IN ҽ���շ�Ŀ¼.�¾�Ŀ¼��־%type:=NULL,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA032)) & "',"
                
                '    AKA033  VARCHAR2(3) Y   ������ҩ��־    0Ϊ������ҩ��1Ϊ������ҩ
                '    ������ҩ��־_IN ҽ���շ�Ŀ¼.������ҩ��־%type:=NULL,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA033)) & "')"
                '    �������_IN ҽ���շ�Ŀ¼.�������%type:=NULL,
                '    ��������޼�_IN ҽ���շ�Ŀ¼.��������޼�%type:=NULL,
                '    ��������޼�_IN ҽ���շ�Ŀ¼.��������޼�%type:=NULL,
                '    ��׼��λ_IN ҽ���շ�Ŀ¼.��׼��λ%type:=NULL,
                '    �ؼ��־_IN ҽ���շ�Ŀ¼.�ؼ��־%type:=NULL,
                '    �¾ɱ�־_IN ҽ���շ�Ŀ¼.�¾ɱ�־%type:=NULL
            Case 2
            
                '    AKA090  VARCHAR2(20)        ��Ŀ����    �ӳ�
                '����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AKA090)) & "',"
                '    AKA091  VARCHAR2(200)   Y   ��Ŀ����    �ӳ�
                '����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & Replace(Nvl(!AKA091), "'", "��") & "',"
                '    Ӣ������_IN IN ҽ���շ�Ŀ¼.Ӣ������%TYPE,
                strSQL = strSQL & "'" & "" & "',"
                '    AKA063  VARCHAR2(3) Y   �շ����    ҽ���������chargetype_info�ж�Ӧֵ��
                '    �շ����_IN IN ҽ���շ�Ŀ¼.�շ����%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA063) & "',"
                '    AKA065  VARCHAR2(3) Y   �շ���Ŀ�ȼ�
                '    �շѵȼ�_IN IN ҽ���շ�Ŀ¼.�շѵȼ�%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA065) & "',"
                '    AKA066  VARCHAR2(14)    Y   ������
                '    ������_IN IN ҽ���շ�Ŀ¼.������%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA066) & "',"
                '    AKA067  VARCHAR2(20)    Y   ��λ
                '    ��λ_IN IN ҽ���շ�Ŀ¼.��λ%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA067) & "',"
                '    AKA068  NUMBER(8,2) Y   һ������޼�    һ��ҽԺ����޼ۣ�����ҽԺ��
                '    ��׼�۸�_IN IN ҽ���շ�Ŀ¼.��׼�۸�%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!AKA068)) & ","
                '    ����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "NUll,"
                '    ���_IN IN ҽ���շ�Ŀ¼.���%TYPE,
                strSQL = strSQL & "NUll,"
                '    AAE013  VARCHAR2(1000)  Y   ��ע
                '    ��ע_IN IN ҽ���շ�Ŀ¼.��ע%TYPE,
                strSQL = strSQL & "'" & Nvl(!AAE013) & "',"
                '    AAE035  DATE    Y   �������
                '    ���ʱ��_IN IN ҽ���շ�Ŀ¼.���ʱ��%TYPE,
                strSQL = strSQL & "to_date('" & Format(!AAE035, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                'NULL��1��ҽԺά����0 ������ά����?,��NULL�������ת����1
                '    ά����־_IN IN ҽ���շ�Ŀ¼.ά����־%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AAA104, "1")) & "',"
                '    ֧����׼_IN IN ҽ���շ�Ŀ¼.֧����׼%TYPE,
                strSQL = strSQL & "NUll,"
                '    ҩƷ����_IN IN ҽ���շ�Ŀ¼.ҩƷ����%TYPE,
                strSQL = strSQL & "NUll,"
                '    ����ҩ��־_IN IN ҽ���շ�Ŀ¼.����ҩ��־%TYPE,
                strSQL = strSQL & "NUll,"
                '    AKA069  NUMBER(5,4) Y   ��ͨ��Ա�Ը�����(1-֧������1)
                '    �Ը�����_IN IN ҽ���շ�Ŀ¼.�Ը�����%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!AKA069)) & ","
                '    ÿ������_IN IN ҽ���շ�Ŀ¼.ÿ������%TYPE,
                strSQL = strSQL & "NUll,"
                '    ʹ��Ƶ��_IN IN ҽ���շ�Ŀ¼.ʹ��Ƶ��%TYPE,
                strSQL = strSQL & "NUll,"
                '    �÷�_IN IN ҽ���շ�Ŀ¼.�÷�%TYPE,
                strSQL = strSQL & "NUll,"
                
                
                '    AKA030  VARCHAR2(1) Y   ͣ�ñ�־    1����,2ͣ��
                '    ��ǰʹ�ñ�־_IN IN ҽ���շ�Ŀ¼.��ǰʹ�ñ�־%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA030) & "',"
                '    �籣�������_IN IN ҽ���շ�Ŀ¼.�籣�������%TYPE
                strSQL = strSQL & "NULL,"
                '    AKA101  VARCHAR2(3)     ҽԺ�ȼ�
                '    ҽԺ�ȼ�_IN IN ҽ���շ�Ŀ¼.ҽԺ�ȼ�%TYPE)
                strSQL = strSQL & "'" & Nvl(!AKA101) & "',"
                '    ZKA001  NUMBER(5,4) Y   ҽ���չ���Ա�Ը�����(1-֧������2)
                '   סԺ�Էѱ���_IN IN ҽ���շ�Ŀ¼.סԺ�Էѱ���%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!ZKA001)) & ","
                '   �α���Ա_IN IN ҽ���շ�Ŀ¼.�α���Ա%TYPE
                strSQL = strSQL & "NULL,"
                '    ͨ��������_IN   ҽ���շ�Ŀ¼.ͨ��������%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    ��Ʒ��_IN   ҽ���շ�Ŀ¼.��Ʒ��%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    С��λ_IN   ҽ���շ�Ŀ¼.С��λ%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    ��С��λ�ۼ�_IN ҽ���շ�Ŀ¼.��С��λ�ۼ�%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    ����_IN     ҽ���շ�Ŀ¼.����%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    �¾�Ŀ¼��־_IN ҽ���շ�Ŀ¼.�¾�Ŀ¼��־%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    ������ҩ��־_IN ҽ���շ�Ŀ¼.������ҩ��־%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    AKA064  VARCHAR2(4) Y   �������    ����������Ĳ������
                '    �������_IN ҽ���շ�Ŀ¼.�������%type:=NULL,
                strSQL = strSQL & "'" & Nvl(!AKA064) & "',"
                '    AKA070  NUMBER(8,2) Y   ��������޼�    ����ҽԺ����޼ۣ�����ҽԺ��
                '    ��������޼�_IN ҽ���շ�Ŀ¼.��������޼�%type:=NULL,
                strSQL = strSQL & "" & Val(Nvl(!AKA070)) & ","
                '    AKA071  NUMBER(8,2) Y   ��������޼�    ����ҽԺ����޼ۣ�һ��ҽԺ��
                '    ��������޼�_IN ҽ���շ�Ŀ¼.��������޼�%type:=NULL,
                strSQL = strSQL & "" & Val(Nvl(!AKA071)) & ","
                '    AKA072  VARCHAR2(200)   Y   ��׼��λ
                '    ��׼��λ_IN ҽ���շ�Ŀ¼.��׼��λ%type:=NULL,
                strSQL = strSQL & "'" & Nvl(!AKA072) & "',"
                '    AKA031  VARCHAR2(1) Y   �ؼ��־    0��ͨ,1����
                '    �ؼ��־_IN ҽ���շ�Ŀ¼.�ؼ��־%type:=NULL,
                strSQL = strSQL & "'" & Nvl(!AKA031) & "',"
                '    AKA032  VARCHAR2(1) Y   �¾ɱ�־    0��,1��
                '    �¾ɱ�־_IN ҽ���շ�Ŀ¼.�¾ɱ�־%type:=NULL
                strSQL = strSQL & "'" & Nvl(!AKA032) & "')"
                
            Case 3
                '    AKA100  VARCHAR2(20)        ҽ�Ʒ�����ʩ����
                '����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA100) & "',"
                '    AKA102  VARCHAR2(200)   Y   ������ʩ����
                '����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA102) & "',"
                'Ӣ������_IN IN ҽ���շ�Ŀ¼.Ӣ������%TYPE,
                strSQL = strSQL & "NULL,"
                '    AKA063  VARCHAR2(3) Y   �շ����
                '�շ����_IN IN ҽ���շ�Ŀ¼.�շ����%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA063) & "',"
                '    AKA103  VARCHAR2(3) Y   �����ȼ�
                '�շѵȼ�_IN IN ҽ���շ�Ŀ¼.�շѵȼ�%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA103) & "',"
                '    AKA066  VARCHAR2(14)    Y   ������
                '������_IN IN ҽ���շ�Ŀ¼.������%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA066) & "',"
                '��λ_IN IN ҽ���շ�Ŀ¼.��λ%TYPE,
                strSQL = strSQL & "NULL,"
                '    AKA068  NUMBER(8,2) Y   һ������޼�    һ��ҽԺ����޼ۣ�����ҽԺ��
                '��׼�۸�_IN IN ҽ���շ�Ŀ¼.��׼�۸�%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!AKA068)) & ","
                '����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "NULL,"
                '���_IN IN ҽ���շ�Ŀ¼.���%TYPE,
                strSQL = strSQL & "NULL,"
                '    AAE013  VARCHAR2(100)   Y   ��ע
                '��ע_IN IN ҽ���շ�Ŀ¼.��ע%TYPE,
                strSQL = strSQL & "'" & Nvl(!AAE013) & "',"
                '    AAE035  DATE    Y   �������
                '���ʱ��_IN IN ҽ���շ�Ŀ¼.���ʱ��%TYPE,
                strSQL = strSQL & "to_date('" & Format(!AAE035, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                'ά����־_IN IN ҽ���շ�Ŀ¼.ά����־%TYPE,
                strSQL = strSQL & "'" & strProcess(Nvl(!AAA104, "1")) & "',"
                '    AKA104  NUMBER(8,2) Y   ����֧����׼
                '֧����׼_IN IN ҽ���շ�Ŀ¼.֧����׼%TYPE,
                '
                strSQL = strSQL & "" & Val(Nvl(!AKA104)) & ","
                'ҩƷ����_IN IN ҽ���շ�Ŀ¼.ҩƷ����%TYPE,
                strSQL = strSQL & "NULL,"
                '����ҩ��־_IN IN ҽ���շ�Ŀ¼.����ҩ��־%TYPE,
                strSQL = strSQL & "NULL,"
                '�Ը�����_IN IN ҽ���շ�Ŀ¼.�Ը�����%TYPE,
                strSQL = strSQL & "NULL,"
                'ÿ������_IN IN ҽ���շ�Ŀ¼.ÿ������%TYPE,
                strSQL = strSQL & "NULL,"
                'ʹ��Ƶ��_IN IN ҽ���շ�Ŀ¼.ʹ��Ƶ��%TYPE,
                strSQL = strSQL & "NULL,"
                '�÷�_IN IN ҽ���շ�Ŀ¼.�÷�%TYPE,
                strSQL = strSQL & "NULL,"
                
                '    AKA030  VARCHAR2(3) Y   ��ǰʹ�ñ�־    1����,2ͣ��
                '��ǰʹ�ñ�־_IN IN ҽ���շ�Ŀ¼.��ǰʹ�ñ�־%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA030) & "',"
                '�籣�������_IN IN ҽ���շ�Ŀ¼.�籣�������%TYPE,
                strSQL = strSQL & "NULL,"
                '    AKA101  VARCHAR2(3)     ҽԺ�ȼ�
                'ҽԺ�ȼ�_IN IN ҽ���շ�Ŀ¼.ҽԺ�ȼ�%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA101) & "',"
                'סԺ�Էѱ���_IN IN ҽ���շ�Ŀ¼.סԺ�Էѱ���%TYPE,
                strSQL = strSQL & "NULL,"
                '�α���Ա_IN IN ҽ���շ�Ŀ¼.�α���Ա%TYPE
                strSQL = strSQL & "NULL,"
                
                '    ͨ��������_IN   ҽ���շ�Ŀ¼.ͨ��������%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    ��Ʒ��_IN   ҽ���շ�Ŀ¼.��Ʒ��%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    С��λ_IN   ҽ���շ�Ŀ¼.С��λ%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    ��С��λ�ۼ�_IN ҽ���շ�Ŀ¼.��С��λ�ۼ�%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    ����_IN     ҽ���շ�Ŀ¼.����%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    �¾�Ŀ¼��־_IN ҽ���շ�Ŀ¼.�¾�Ŀ¼��־%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    ������ҩ��־_IN ҽ���շ�Ŀ¼.������ҩ��־%type:=NULL,
                strSQL = strSQL & "NULL,"
                
                '    AKA064  VARCHAR2(4) Y   �������    ����������Ĳ������
                '    �������_IN ҽ���շ�Ŀ¼.�������%type:=NULL,
                strSQL = strSQL & "'" & Nvl(!AKA064) & "',"
                '    AKA070  NUMBER(8,2) Y   ��������޼�    ����ҽԺ����޼ۣ�����ҽԺ��
                '    ��������޼�_IN ҽ���շ�Ŀ¼.��������޼�%type:=NULL,
                strSQL = strSQL & "" & Val(Nvl(!AKA070)) & ","
                '    AKA071  NUMBER(8,2) Y   ��������޼�    ����ҽԺ����޼ۣ�һ��ҽԺ��
                '    ��������޼�_IN ҽ���շ�Ŀ¼.��������޼�%type:=NULL,
                strSQL = strSQL & "" & Val(Nvl(!AKA071)) & ","
                '    AKA072  VARCHAR2(200)   Y   ��׼��λ
                '    ��׼��λ_IN ҽ���շ�Ŀ¼.��׼��λ%type:=NULL,
                strSQL = strSQL & "'" & Nvl(!AKA072) & "',"
                '    �ؼ��־_IN ҽ���շ�Ŀ¼.�ؼ��־%type:=NULL,
                strSQL = strSQL & "NULL,"
                '    AKA032  VARCHAR2(1) Y   �¾ɱ�־    0��,1��
                '    �¾ɱ�־_IN ҽ���շ�Ŀ¼.�¾ɱ�־%type:=NULL
                strSQL = strSQL & "'" & Nvl(!AKA032) & "')"
            
            Case 4
                ' ����������|�����������
                strSQL = "ZL_ҽ���շ����_UPDATE("
                '����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA063) & "',"
                '����_IN IN ҽ���շ�Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA110) & "',"
                '    �������_IN IN ҽ���շ����.�������%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA111) & "',"
                '    ��������_IN IN ҽ���շ����.��������%TYPE
                strSQL = strSQL & "'" & Nvl(!AKA112) & "')"
            Case 88
            '   �����medicine_info_modified
            '   ������"ҩƷ�����Ϣ��"���˱�Ϊ����ά������ҩƷĿ¼�����ѯ��

            '    AKA060  VARCHAR2(60)    N   ҩƷ����    ʡĿ¼���
            '    AKA061  VARCHAR2(100)   Y   ��������    ͨ����
            '    AKA063  VARCHAR2(3) Y   �շ����    11��ҩ�� 12��ҩ 13��ҩ
            '    AKA065  VARCHAR2(3) Y   �շ���Ŀ�ȼ�    1���� 2���� 3�Է�
            '    AKA066  VARCHAR2(30)    Y   ������
            '    AKA067  VARCHAR2(20)    Y   ��λ
            '    AKA068  NUMBER(8,2) Y   ��׼�۸�
            '    AKA070  VARCHAR2(50)    Y   ����
            '    AKA074  VARCHAR2(50)    Y   ���
            '    AAE035  DATE    Y   �������
            '    AKA030  VARCHAR2(3) Y   ��ǰʹ�ñ�־
            '    AKA075  VARCHAR2(40)    Y   ͨ��������
            '    AKA076  VARCHAR2(100)   Y   ��Ʒ��
            '    AKA078  VARCHAR2(20)    Y   С��λ
            '    AKA079  NUMBER(8,2) Y   ��С��λ���ۼ�
            '    AKA031  VARCHAR2(100)   Y   ����
            '    AKA032  VARCHAR2(1) Y   �¾�Ŀ¼��־    0ԭĿ¼��1ʡĿ¼
            '    AKA033  VARCHAR2(3) Y   ������ҩ��־    0Ϊ������ҩ��1Ϊ������ҩ
            '    AKA030  VARCHAR2(3) Y   ��ǰʹ�ñ�־    0���á�1ͣ��
            '    FILENAME    VARCHAR2(100)       �����ļ�����
            '    CHANGETYPE  VARCHAR2(3)     �������
            '    SIGN    VARCHAR2(3)     �¾ɱ�־
            '    AAE036  DATE        ��������
            
            Case Else
                strSQL = "ZL_ҽ������Ŀ¼_UPDATE("
                'AKA120 VARCHAR2(20)                  ���ֱ���
                '    ����_IN IN ҽ������Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA120) & "',"
                'AKA121 VARCHAR2(50) Y                ��������
                '    ����_IN IN ҽ������Ŀ¼.����%TYPE,
                strSQL = strSQL & "'" & Replace(Nvl(!AKA121), "'", "��") & "',"
                'AKA066 VARCHAR2(14) Y                ������
                '    ������_IN IN ҽ������Ŀ¼.������%TYPE,
                strSQL = strSQL & "'" & Replace(Nvl(!AKA066), "'", "��") & "',"
                'AKA122 VARCHAR2(3)  Y                ���ַ���
                '    ���ַ���_IN IN ҽ������Ŀ¼.���ַ���%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA122) & "',"
                'AKA124 NUMBER(6,4)  Y                �����Ը�����
                '    �Ը�����_IN IN ҽ������Ŀ¼.�Ը�����%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!AKA124)) & ","
                'AKA030 VARCHAR2(3)  Y                ��ǰʹ�ñ�־
                '    ��ǰʹ�ñ�־_IN IN ҽ������Ŀ¼.��ǰʹ�ñ�־%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA030) & "',"
                'AKA125 NUMBER(8,2)  Y
                '    AKA125_IN IN ҽ������Ŀ¼.AKA125%TYPE,
                strSQL = strSQL & "" & Val(Nvl(!AKA125)) & ","
                'AKA126 VARCHAR2(3)  Y                ͳ����
                '    ͳ����_IN IN ҽ������Ŀ¼.ͳ����%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA126) & "',"
                'AKA128 VARCHAR2(14) Y                ������ˮ��
                '    ������ˮ��_IN IN ҽ������Ŀ¼.������ˮ��%TYPE,
                strSQL = strSQL & "'" & Nvl(!AKA128) & "',"
                'AAE035 DATE         Y                �������
                '    ���ʱ��_IN IN ҽ������Ŀ¼.���ʱ��%TYPE
                strSQL = strSQL & "to_date('" & Format(!AAE035, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"
            End Select
            gcnOracle_����.Execute strSQL, , adCmdStoredProc
            If Not objProgss Is Nothing Then
                objProgss.Value = i
            Else
                zlCommFun.ShowFlash "�������ء�" & strTitle & "������,������" & i & "/" & lngCount & ""
            End If
            i = i + 1
            .MoveNext
        Loop
    End With
    
   ���ط�����ĿĿ¼_���� = True
   Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function strProcess(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    strProcess = IIf(IsNull(varValue), DefaultValue, varValue)
    strProcess = Replace(strProcess, "'", "��")
End Function


Public Function GetItemInfo_����(ByVal lngPatiID As Long, ByVal lngItemID As Long, Optional ByVal strժҪ As String, Optional intType As Integer = 0) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�������˵������ʾ��Ϣ
    '--�����:
    '--������:
    '--��  ��:��ʾ��
    '-----------------------------------------------------------------------------------------------------------
    Dim strMsgInfor As String
    Dim strԭժҪ As String, StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim str��Ŀ���� As String
    Dim strҽԺ���� As String
    Dim blnҩƷ  As Boolean
    strԭժҪ = strժҪ
    
    gstrSQL = "Select a.��Ŀ����,a.��ע,b.���� from ����֧����Ŀ a,�շ�ϸĿ b where a.�շ�ϸĿid=b.id and  nvl(a.�Ƿ�ҽ��,0)=0 and ����=[1] and a.�շ�ϸĿiD=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ƿ��������ƣ�������͹���ҩƷ!", TYPE_�ٲ׷���, lngItemID)
    
    If rsTemp.EOF Then
        Exit Function
    End If
    
    str��Ŀ���� = Nvl(rsTemp!��Ŀ����, "9000900099")
    strҽԺ���� = Nvl(rsTemp!����, "9000900099")
    blnҩƷ = Val(Nvl(rsTemp!��ע)) = 1
     
    gstrSQL = "Select ҽ���� From �����ʻ� where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ��֤��", lngPatiID)
    If rsTemp.EOF Then
        ShowMsgbox "�ò��˲���ҽ������!"
        Exit Function
    End If
    
    '����Ƿ�����
    '��ȡ������������Ŀ���
    StrInput = Nvl(rsTemp!ҽ����, 0) & "|"
   If blnҩƷ Then
    '��˵:ҩƷֻ�ܴ�ҽԺ����,������ֻ�ܵ�ֻ�ܴ�ҽ������
        StrInput = StrInput & strҽԺ���� & "|"
    Else
        StrInput = StrInput & str��Ŀ���� & "|"
    End If
    StrInput = StrInput & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If ҵ������_����(����_��Ŀ���������ѯ, StrInput, strOutput) = False Then
        ShowMsgbox "����Ŀ�Ƿ��ؼ����λ����ҩƷ����δ����������"
        Exit Function
    End If
    
    '  ִ�гɹ�ʱΪ�������(=0ʱ),ִ��ʧ��ʱ��=100���Ϊ�գ�Ϊû�в鵽��Ӧ������Ϣ����ʧ��ԭ��������< 0ʱ����ѯʧ�ܣ�����ʧ��ԭ��������
    If Split(strOutput, "|")(1) = "" Then
        ShowMsgbox "ע��:" & vbCrLf & "   ���շ�ϸĿδ�ܹ�����,����ȫ�ԷѴ���"
    End If
End Function
Public Function ��ȡ�α���Ա��Ϣ_����() As Boolean
    '��ȡ�α���Ա��Ϣ
    Dim StrInput As String
    Dim strOutput As String
    Dim strArr
    ��ȡ�α���Ա��Ϣ_���� = False
    
    Err = 0
    On Error GoTo errHand:
    
    '���ӿ��ϻ�ȡҽ��֤��
    If ������_����() = False Then Exit Function
    
    '������ҽ��֤�Ż�ȡ�α���Ϣ
    StrInput = g�������_����.ҽ��֤��
    If ҵ������_����(����_��òα���Ա��Ϣ, StrInput, strOutput) = False Then
        
        Exit Function
    End If
    If strOutput = "" Then
        ShowMsgbox "�ڻ�ȡ�α���Ա��Ϣʱ���ӿڷ����˿�ֵ!"
        Exit Function
    End If
    strArr = Split(strOutput, "|")
    '����: ����|�Ա�|���֤|��������|��Ա������|��Ա�������|��λ����|��λ����|ͳ������
    '   ͳ�����ţ��Ƕ���˾���ǳ�Ҫ��Ҫ�ں����Ĳ���
    With g�������_����
        
        .���� = strArr(1)
        .�Ա� = strArr(2)
        .���֤�� = strArr(3)
        .�������� = strArr(4)
        .������ = strArr(5)
        .������� = strArr(6)
        .��λ���� = strArr(7)
        .��λ���� = strArr(8)
        .ͳ������ = strArr(9)
        .���� = Get����(.��������)
    End With
    ��ȡ�α���Ա��Ϣ_���� = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function Get����(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as ���� from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����"
    If Not rsTemp.EOF Then
        Get���� = Int(Nvl(rsTemp!����, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function




