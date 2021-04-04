Attribute VB_Name = "mdlPublicExpense"
Option Explicit
Public gstrSysName As String                'ϵͳ����
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String    '��Ʒ����
Public gstrSQL As String
Public glngSys As Long
Public glngMainModule As Long '�����ߵ�ģ���
Public gstrMainPrivs As String '�����ߵ����Ȩ��
Public gblnOK As Boolean
Public gclsInsure As New clsInsure 'ҽ������
Public gstrDBUser As String '������
Public gcnOracle As ADODB.Connection
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gobjSquare As Object '�����㲿��
Public gobjPlugIn As Object '��ҹ���

'�Һ��ò���
Public gstrRooms As String
Public glngModul As Long
Public gbytState As Byte
Public gstrDocs As String
Public gstrDeptIDs As String
Public gstrPrivs As String
Public gblnBill�Һ� As Boolean
Public gbytRegistMode As Byte
Public gdatRegistTime As Date

Public grsҽ�Ƹ��ʽ As ADODB.Recordset
Public grsOneCard As ADODB.Recordset

Private Type TY_Decimal_Precision 'С������
    byt_Bit As Byte 'С��λ��:��ʾ���㵽С�����ڶ���λ��
    strFormt_VB As String   'VB��ʽ��:0.0000;...
    strFormt_ORA As String  'Oracle��ʽ��:999990.00000...
End Type

Private Type ty_SysPara
    bln�����������۷���  As Boolean
    bytƱ�ݷ������ As Byte   'Ʊ�ݷ������:0-����ʵ�ʴ�ӡ����Ʊ��;1-����ϵͳԤ���������;2-�����û��Զ���������
    Money_Decimal As TY_Decimal_Precision  '���ý��С����ʽ
    Price_Decimal As TY_Decimal_Precision  '���õ���С����ʽ
    bln��������ۿ�  As Boolean
    bytҩƷ������ʾ As Byte '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
    byt����ҩƷ��ʾ As Byte '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��

    byt������˷�ʽ As Byte '������˷�ʽ:0-δ��˲�������ʣ�ȱʡΪ0;1-���ʱ����������ú�ҽ��������ҽ�������ͷ��õ�����
    blnδ��ƽ�ֹ����  As Boolean
    bln����ִ�з��� As Boolean 'ִ��֮�������Զ�����
    blnִ�к���� As Boolean
    blnִ��ǰ�Ƚ��� As Boolean 'һ��ִͨ��ǰ���շѻ�������
    bln�������������� As Boolean '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
    intҽ������ As Integer '�Ƿ��סԺҽ�����˵���Ŀ����������м��:0-�����,1-��鲢����,2-��鲢��ֹ
    dblMaxMoney As Double   '�������
    bytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    dblԤ��������鿨 As Double 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
    bln����ƥ�䷽ʽ�л� As Boolean '�����ڴ��ڽ���Ĺ������л�����ƥ�䷽ʽ�л�

    intסԺ�Զ����� As Integer   'סԺ������ɺ��Ƿ��Զ�����:0-���Զ����ϣ�1-�Զ����ϣ�2-�����ҿ���ʱ�Զ�����
    bln�����Զ����� As Boolean '���������ɺ��Ƿ��Զ�����
    bln�շѺ��Զ���ҩ As Boolean '
    bln���뷢ҩ As Boolean
    strҽ���������� As String 'ҽ����������ķ�������
    str���ѷ������� As String '���Ѳ�������ķ�������
    strLike As String
    bytCode As Byte
    bln�շ���� As Boolean '�Ƿ������������
    blnFeeKindCode As Boolean '�������ʱ,��λ�����շ�������
    strMatchMode As String '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
    blnStock As Boolean 'ָ��ҩ��ʱ�Ƿ��޶�����ҩƷ�Ŀ��
    bln�������ۼ���  As Boolean
    blnסԺ���ۼ��� As Boolean
    bln��Һ�ģʽ As Boolean '�Ƿ����ģʽ,���̣�ֱ���ڷ���̨ȡ�ţ�Ȼ���ڽ���ʱ���������۵�
    byt��������ʶ����� As Byte   '�Ƿ������ʶ��::1-����ͨ��ɨ��¼���¼������;0-�����ƣ�����ͨ������Ȳ���
    bln����ʾ�޿������ As Boolean
End Type

Public gSysPara As ty_SysPara
Public Enum gEm_BulidIng_SQLType
    EM_Bulid_�ַ� = 0
    EM_Bulid_���� = 1
End Enum
Public Const gstrCompentsName = ""
Public Enum Enum_Inside_Program
    pסԺ���ʲ��� = 1150
    pҽ�����ѹ��� = 1257
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    
End Enum
Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
    רҵ����ְ�� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum ����Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum

'----------------------------------------------------
'����������
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gstrNodeNo As String 'վ����
Public glngInstanceCount As Long '��������
Public glngMax��ͥ��ַ As Long       '��ͥ��ַ�������¼�볤��
Public glngMax���ڵ�ַ As Long       '���ڵ�ַ�������¼�볤��
Public glngMax�����ص� As Long       '�����ص��������¼�볤��
Public glngMax��ϵ�˵�ַ As Long    '��ϵ�˵�ַ�������¼�볤��

Public Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�Ȩ�ֱ���
    '����:���˺�
    '����:2014-03-20 16:07:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, varTmp As Variant
    Dim strValue As String
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = gobjComlib.GetUnitName
    gbytRegistMode = Val(Split(gobjDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(0))
    If Split(gobjDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(1) <> "" Then
        gdatRegistTime = CDate(Format(Split(gobjDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(1), "yyyy-mm-dd hh:mm:ss"))
    End If
    
    With gSysPara
        .bln�����������۷��� = gobjDatabase.GetPara(98, glngSys) = "1"
        With .Money_Decimal '���ý��С��λ��
            .byt_Bit = Val(gobjDatabase.GetPara(9, glngSys, , 2))
            .strFormt_VB = "0." & String(.byt_Bit, "0")
            .strFormt_ORA = "FM" & String(14, "9") & "0." & String(.byt_Bit, "9")
        End With
        With .Price_Decimal  '���õ���С��λ��
            .byt_Bit = Val(gobjDatabase.GetPara(157, glngSys, , 5))
            .strFormt_VB = "0." & String(.byt_Bit, "0")
            .strFormt_ORA = "FM" & String(14, "9") & "0." & String(.byt_Bit, "9")
        End With
        '���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
        strTmp = Trim(gobjDatabase.GetPara("Ʊ�ݷ������", glngSys, 1121, "0||0;0;0;0;0"))
        varTmp = Split(strTmp & "||", "||")
        .bytƱ�ݷ������ = Val(varTmp(0))
        .bln��������ۿ� = Val(gobjDatabase.GetPara(93, glngSys)) <> 0
        .bytҩƷ������ʾ = Val(gobjDatabase.GetPara("ҩƷ������ʾ", , , "2"))
        .byt����ҩƷ��ʾ = gobjDatabase.GetPara("����ҩƷ��ʾ", , , 0)
        .byt������˷�ʽ = Val(gobjDatabase.GetPara(185, glngSys))    '49501
        .blnδ��ƽ�ֹ���� = Val(gobjDatabase.GetPara(215, glngSys)) = 1 '51612
        .bln�������ۼ��� = gobjDatabase.GetPara("�������۲��˼���", glngSys, 1150) = "1"
        .blnסԺ���ۼ��� = gobjDatabase.GetPara("סԺ���۲��˼���", glngSys, 1150) = "1"
        
        '33:����ԭBUG��Ϊ14403����ϪҽԺ�����ٴ���ҽ��ִ�еǼ�ʱ��
        ' �ü�顢���顢������Ŀ�Ѿ���ɣ����������Ѿ�ʹ���ˣ�
        ' ���ԣ�Ӧ�ö��������õ����ҵģ����ƵľͲ������߷��ϵ����̣�
        ' ���ԣ�������û�б�Ҫ���ڣ�Ӧ�ö�����Ϊִ�к��Զ��Ը������õ����ķ��ϣ�
        ' ȡ��ִ��ʱ�Զ�����
        
        .bln����ִ�з��� = True ' Val(gobjDatabase.GetPara(33, glngSys)) <> 0
        ' 81����:�ò���������10.03��ǰ�ʹ��ڣ�δ�ҵ�BUG�š���˻��۵���Ŀ����ȷ�Ϸ��ã�ִ��֮�������ȷ�Ϸ��ã��ͻ���Ҫ�˹�����ȥ��˻��۵�����ҵ��������˵��������û�б�Ҫ���ڣ�Ӧ�ö�����Ϊִ�к��Զ���˻��۵���������ؿ��ư����ϴ˲������д���
        .blnִ�к���� = True  ' Val(gobjDatabase.GetPara(81, glngSys)) <> 0
        '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
        .blnִ��ǰ�Ƚ��� = Val(gobjDatabase.GetPara(163, glngSys)) <> 0
        '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
        .bln�������������� = Val(gobjDatabase.GetPara(232, glngSys)) <> 0
        'ҽ��������
        .intҽ������ = Val(gobjDatabase.GetPara(59, glngSys, , 1))
        '���ʷ���������ѽ��
        .dblMaxMoney = Val(gobjDatabase.GetPara(60, glngSys))
    
        '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
        .bytBillOpt = Val(gobjDatabase.GetPara(23, glngSys))
        'һ��ͨ������֤
        strValue = gobjDatabase.GetPara(28, glngSys, , "1|0")
        If InStr(strValue, "|") = 0 Then strValue = "1|0"
        .dblԤ��������鿨 = Val(Split(strValue, "|")(0))
        .bln����ƥ�䷽ʽ�л� = Val(gobjDatabase.GetPara("����ƥ�䷽ʽ�л�", , , "1")) = 1
        '�����Զ�����
        .bln�����Զ����� = Val(gobjDatabase.GetPara(92, glngSys)) <> 0
        'סԺ�Զ�����
        .intסԺ�Զ����� = Val(gobjDatabase.GetPara(63, glngSys))
        '�Զ���ҩ��ҩ
        .bln�շѺ��Զ���ҩ = gobjDatabase.GetPara(45, glngSys) = "1"
        '�����շ��뷢ҩ����
        .bln���뷢ҩ = gobjDatabase.GetPara(15, glngSys) = "1"
        'ҽ����������
        .strҽ���������� = "'" & Replace(gobjDatabase.GetPara(41, glngSys), "|", "','") & "'"
    
        '���ѷ�������
        .str���ѷ������� = "'" & Replace(gobjDatabase.GetPara(42, glngSys), "|", "','") & "'"
            
        '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
        .strMatchMode = gobjDatabase.GetPara(44, glngSys, , "00")
        
        .strLike = IIf(gobjDatabase.GetPara("����ƥ��") = "0", "%", "")
        .bytCode = Val(gobjDatabase.GetPara("���뷽ʽ"))
        '�Ƿ�Ҫ�������������
        .bln�շ���� = Val(gobjDatabase.GetPara(72, glngSys, , 1)) <> 0
        '���������ʱ,���������Ŀʱ,��λ����������
        .blnFeeKindCode = Val(gobjDatabase.GetPara(144, glngSys)) <> 0 And Not .bln�շ����
        'ָ��ҩ��ʱ���ƿ��
        .blnStock = Val(gobjDatabase.GetPara(18, glngSys)) <> 0
        .bln��Һ�ģʽ = Val(gobjDatabase.GetPara("��Һ�ģʽ", glngSys)) = 1
        .byt��������ʶ����� = Val(gobjDatabase.GetPara(320, glngSys, , "0"))      '1-����ͨ��ɨ��¼���¼������;0-�����ƣ�����ͨ������Ȳ���
        .bln����ʾ�޿������ = Val(gobjDatabase.GetPara(316, glngSys)) = 1
    End With
    Call InitAddressLength
End Sub

Public Function zlGetFeeFields(Optional strTableName As String = "������ü�¼", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ�����ֵ
    '��Σ�strTableName:��:������ü�¼;סԺ���ü�¼;....
    '      blnReadDatabase-�����ݿ��ж�ȡ
    '���Σ�
    '���أ��ֶμ�
    '���ƣ����˺�
    '���ڣ�2010-03-10 10:41:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, strFileds As String
    
    Err = 0: On Error GoTo Errhand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "������ü�¼"
        zlGetFeeFields = "" & _
        "Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, " & _
        "����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, " & _
        "�Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, " & _
        "����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, " & _
        "���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
        Exit Function
    Case "סԺ���ü�¼"
        zlGetFeeFields = "" & _
         " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, " & _
         " �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, " & _
         " ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, " & _
         " ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, " & _
         " ����id , ���ʽ��, ���մ���ID, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
         Exit Function
    Case "���˽��ʼ�¼"
        zlGetFeeFields = "Id, No, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������, ��ע"
        Exit Function
    Case "����Ԥ����¼"
        zlGetFeeFields = "" & _
        " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, " & _
        " ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�,Ԥ�����,�����ID,���㿨���,����,������ˮ��,����˵��,������λ,�������,У�Ա�־"
        Exit Function
    Case "��Ա��"
        zlGetFeeFields = "" & _
        "Id, ���, ����, ����, ���֤��, ��������, �Ա�, ����, ��������, �칫�ҵ绰, �����ʼ�, ִҵ���, ִҵ��Χ, " & _
        "����ְ��, רҵ����ְ��, Ƹ�μ���ְ��, ѧ��, ��ѧרҵ, ��ѧʱ��, ��ѧ����, ������ѵ, ���п���, ���˼��, ����ʱ��, " & _
        "����ʱ��, ����ԭ��, ����, վ��"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo Errhand:
    strSql = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ����Ϣ", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & Nvl(!Column_Name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
Errhand:
    zlGetFeeFields = "*"
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = gobjComlib.Nvl(varValue, DefaultValue)
End Function

Public Function GetPatiMoney(ByVal bytType As Byte, ByVal lng����ID As Long, ByRef objPatiFee As clsPatiFeeinfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˵���ط�����Ϣ
    '���:bytType-0-����;1-סԺ
    '     lng����ID-����ID
     '����:
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 16:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Set objPatiFee = New clsPatiFeeinfor
    On Error GoTo errHandle
    If bytType = 0 Then
        strSql = "" & _
        "   Select Nvl(Ԥ�����,0) Ԥ�����,Nvl(�������,0) �������,0 as Ԥ�����,0 as ������ " & _
        "   From ������� " & _
        "   Where ����=1 And ����=1 And ����ID=[1]" & _
        "   "
    Else
        strSql = "" & _
        "   Select Nvl(Ԥ�����,0) Ԥ�����,Nvl(�������,0) �������,0 as Ԥ����� ,0 as ������" & _
        "   From ������� " & _
        "   Where ����=1 And ����=2 And ����ID=[1]" & _
        "   Union ALL " & _
        "   Select 0 as Ԥ�����,0 as �������,Sum(B.���) as Ԥ����� ,0 as ������" & _
        "   From ������Ϣ A,����ģ����� B" & _
        "   Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[1]"
    End If
    strSql = strSql & "" & _
    "   Union ALL " & _
    "   Select 0 as Ԥ�����,0 as �������,0 as Ԥ�����,������" & _
    "   From ������Ϣ B " & _
    "   Where ����ID=[1]"
    
    strSql = "" & _
    "   Select Nvl(Sum(Ԥ�����),0) as Ԥ�����,Nvl(Sum(�������),0) as �������,Nvl(Sum(Ԥ�����),0) as Ԥ�����,Nvl(Sum(������),0) as ������  " & _
    "   From (" & strSql & ")"
    
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ���˵���ط��ý��", lng����ID)
    If rsTemp.EOF Then GetPatiMoney = True: Exit Function
    With objPatiFee
        .Ԥ����� = FormatEx(Val(Nvl(rsTemp!Ԥ�����)), 6)
        .δ����� = FormatEx(Val(Nvl(rsTemp!�������)), 6)
        .Ԥ����� = FormatEx(Val(Nvl(rsTemp!Ԥ�����)), 6)
        .������ = FormatEx(Val(Nvl(rsTemp!������)), 6)
        .ʣ��� = FormatEx(.Ԥ����� + .Ԥ����� - .δ�����, 6)
    End With
    GetPatiMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function FromIDsBulidIngSQL(ByVal bytBulidType As gEm_BulidIng_SQLType, _
    ByVal strValues As String, _
    ByRef varPara As Variant, ByRef strBulitSQL As String, _
    ByVal strAliaName As String, Optional intStartPara As Integer = 1 _
    ) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����IDs����ȡ��ص�SQL,��:select ... From str2List Union ALL Selelct ..
    '���:strValues-ֵ,����ö��ŷ���
    '     strAliaName-����
    '     bytType-0-�ַ���;1-������;
    '     intStartPara-�����Ĳ���
    '����:varPara-���صĲ���ֵ������
    '     strBulitSQL-���صĹ�����SQL��
    '����:�����ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-25 17:04:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, strTemp As String
    Dim i As Long, j As Long, strSql As String
    Dim strPara() As Variant, strTable As String, strColumnName As String
    
    On Error GoTo errHandle
    
    strColumnName = " Column_Value "
    If strAliaName <> "" Then strColumnName = strColumnName & " As " & strAliaName
    
    If bytBulidType = EM_Bulid_�ַ� Then
        strTable = "Table(f_str2list([0]))"
    Else
        strTable = "Table(f_Num2list([0]))"
    End If
    
    j = intStartPara
    ReDim Preserve strPara(0 To j - 1) As Variant
    
    
    varData = Split(strValues, ",")
    strTemp = ""
    For i = 0 To UBound(varData)
        If gobjCommFun.ActualLen(strTemp & "," & varData(i)) > 4000 Then
            strSql = strSql & " Union ALL  Select " & strColumnName & " From " & Replace(strTable, "[0]", "[" & j & "]")
            ReDim Preserve strPara(0 To j - 1) As Variant
            strPara(j - 1) = Mid(strTemp, 2)
            j = j + 1
            strTemp = ""
        End If
        strTemp = strTemp & "," & varData(i)
    Next
    If strTemp <> "" Then
        strSql = strSql & " Union ALL  Select " & strColumnName & " From " & Replace(strTable, "[0]", "[" & j & "]")
        ReDim Preserve strPara(0 To j - 1) As Variant
        strPara(j - 1) = Mid(strTemp, 2)
    End If
    
    varPara = strPara
    If strSql <> "" Then strSql = Mid(strSql, 11)
    strBulitSQL = strSql
    FromIDsBulidIngSQL = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetFeeMoneyFromAdviceIDs(ByVal strҽ��IDs As String, _
    ByRef dblOutӦ�ս�� As Double, ByRef dblOutʵ�ս�� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ��IDs����ȡӦ�պ�ʵ�ս��
    '���:strҽ��IDs-ҽ��ID,����ö��ŷ���
    '����:dblOutӦ�ս��-Ӧ�ս��
    '     dblOutʵ�ս��-ʵ�ս��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-25 16:11:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim varPara As Variant
    dblOutӦ�ս�� = 0: dblOutʵ�ս�� = 0
    If strҽ��IDs = "" Then Exit Function
    
    '���ܴ���4000
    If gobjCommFun.ActualLen(strҽ��IDs) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_����, strҽ��IDs, varPara, strSql, "ҽ��ID") = False Then Exit Function
        strSql = "" & _
        " Select /*+ RULE */ " & _
        "   Nvl(Sum(Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(ʵ�ս��), 0) As ʵ�ս�� " & _
        " From (With ҽ������ As (" & strSql & ") " & _
        "        Select Nvl(Sum(a.Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս�� " & _
        "        From ������ü�¼ A, ҽ������ B " & _
        "        Where a.ҽ����� = b.ҽ��id " & _
        "        Union All " & _
        "        Select Nvl(Sum(a.Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս�� " & _
        "        From סԺ���ü�¼ A, ҽ������ B " & _
        "        Where a.ҽ����� = b.ҽ��id)"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID��ȡ��صķ��ý��", varPara)
    Else
        strSql = "" & _
        " Select /*+ RULE */ " & _
        "   Nvl(Sum(Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(ʵ�ս��), 0) As ʵ�ս�� " & _
        " From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Num2list([1]))) " & _
        "        Select Nvl(Sum(a.Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս�� " & _
        "        From ������ü�¼ A, ҽ������ B " & _
        "        Where a.ҽ����� = b.ҽ��id " & _
        "        Union All " & _
        "        Select Nvl(Sum(a.Ӧ�ս��), 0) As Ӧ�ս��, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս�� " & _
        "        From סԺ���ü�¼ A, ҽ������ B " & _
        "        Where a.ҽ����� = b.ҽ��id)"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID��ȡ��صķ��ý��", strҽ��IDs)
    End If
    
    On Error GoTo errHandle
    dblOutӦ�ս�� = FormatEx(Val(Nvl(rsTemp!Ӧ�ս��)), 6)
    dblOutʵ�ս�� = FormatEx(Val(Nvl(rsTemp!ʵ�ս��)), 6)
    
    GetFeeMoneyFromAdviceIDs = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         Set gobjSquare.objSquareCard = Nothing
     End If
     Set gobjSquare = Nothing
     If Err <> 0 Then Err.Clear: Err = 0
End Sub

Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub


Public Function AdviceIsCharged(ByVal strҽ��IDs As String, _
    ByVal strNos As String, ByRef bytOutChargeStatus As Byte, Optional ByRef strOutδ��ҽ��IDs As String, _
    Optional ByRef bytOutBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ҽ���Ƿ��Ѿ��շ�
    '���:strҽ��IDs-ҽ��ID,����ö��ŷ���
    '����:bytOutChargeStatus-�շ�״̬(0-δ�շ�,1-��ȫ�շ�;2-�����շ�)
    '     strOutδ��ҽ��IDs-����δ�շѻ�δ����˵�ҽ��ID
    '     bytOutBillType:���ص�ǰ�ĵ�������(0-�������κε���;1-�շѵ�;2-���ʵ�;3-�շѺͼ��ʶ���)
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 09:48:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim varPara As Variant
    Dim bytStatus As Byte
    strOutδ��ҽ��IDs = "": bytOutBillType = 0: bytOutChargeStatus = 0
    If strNos = "" And strҽ��IDs = "" Then Exit Function
    
    If strҽ��IDs <> "" Then
        '���ܴ���4000
        If gobjCommFun.ActualLen(strҽ��IDs) > 4000 Then
            If FromIDsBulidIngSQL(EM_Bulid_����, strҽ��IDs, varPara, strSql, "ҽ��ID") = False Then Exit Function
            strSql = "" & _
            " Select /*+ RULE */ distinct  ��¼����, ��¼״̬,ҽ�����" & _
            " From (With ҽ������ As (" & strSql & ") " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From ������ü�¼ A,ҽ������ B " & _
            "        Where a.ҽ����� = b.ҽ��id And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From סԺ���ü�¼ A, ҽ������ B " & _
            "        Where a.ҽ����� = b.ҽ��id And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID��ȡ��صķ��ý��", varPara)
        Else
            strSql = "" & _
            " Select /*+ RULE */ distinct  ��¼����, ��¼״̬,ҽ�����" & _
            " From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Num2list([1]))) " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From ������ü�¼ A,ҽ������ B " & _
            "        Where a.ҽ����� = b.ҽ��id And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From סԺ���ü�¼ A, ҽ������ B " & _
            "        Where a.ҽ����� = b.ҽ��id And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID��ȡ��صķ��ý��", strҽ��IDs)
        End If
    Else
        '�����ݺŴ���
        '���ܴ���4000
        If gobjCommFun.ActualLen(strNos) > 4000 Then
            If FromIDsBulidIngSQL(EM_Bulid_�ַ�, strNos, varPara, strSql, "NO") = False Then Exit Function
            strSql = "" & _
            " Select /*+ RULE */ distinct  ��¼����, ��¼״̬,ҽ�����" & _
            " From (With ҽ������ As (" & strSql & ") " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From ������ü�¼ A,ҽ������ B " & _
            "        Where a.NO = b.NO And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From סԺ���ü�¼ A, ҽ������ B " & _
            "        Where a.NO = b.NO And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID��ȡ��صķ��ý��", varPara)
        Else
            strSql = "" & _
            " Select /*+ RULE */ distinct  ��¼����, ��¼״̬,ҽ�����" & _
            " From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Str2list([1]))) " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From ������ü�¼ A,ҽ������ B " & _
            "        Where a.NO = b.NO And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.��¼����,A.��¼״̬,A.ҽ����� " & _
            "        From סԺ���ü�¼ A, ҽ������ B " & _
            "        Where a.NO = b.NO And A.��¼���� in (1,2,3) And A.��¼״̬ IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID��ȡ��صķ��ý��", strҽ��IDs)
        End If
        
    End If
    On Error GoTo errHandle
    With rsTemp
        bytStatus = -1
        Do While Not .EOF
             If Val(Nvl(!��¼״̬)) = 0 Then  'δ�շ�
                If Val(Nvl(!ҽ�����)) <> 0 Then
                    strOutδ��ҽ��IDs = strOutδ��ҽ��IDs & "," & Nvl(rsTemp!ҽ�����)
                End If
             End If
             If bytStatus = -1 Then
                If Val(Nvl(!��¼״̬)) = 0 Then
                    bytStatus = IIf(Val(Nvl(!��¼״̬)) = 0, 0, 1)
                End If
             ElseIf bytStatus = 0 And (Val(Nvl(!��¼״̬)) = 1 Or Val(Nvl(!��¼״̬)) = 3) Then
                bytStatus = 2   '�����շ�
             ElseIf bytStatus = 1 And Val(Nvl(!��¼״̬)) = 0 Then
                bytStatus = 2 '�����շ�
             End If
             
             If bytOutBillType = 0 Then
                bytOutBillType = Val(Nvl(!��¼����))
             ElseIf bytOutBillType <> Val(Nvl(!��¼����)) Then
                '��������
                bytOutBillType = 3
             End If
            .MoveNext
        Loop
    End With
    bytOutChargeStatus = bytStatus
    AdviceIsCharged = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function BillExistNotBalance(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��շѵ����Ƿ����δ�շѵ�
    '���:strNOs:ָ���ĵ��ݺ�,������,�ö��ŷ���
    '����:
    '����:�����д���δ�շѵ�,����true,���򷵻�False
    '����:Ƚ����
    '����:2016-08-25 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSql As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    '���ܴ���4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_�ַ�, strNos, varPara, strSql, "NO") = False Then Exit Function
        strSql = "Select /*+cardinality(b,10)*/ 1" & vbNewLine & _
                " From ������ü�¼ A,(" & strSql & ") B" & vbNewLine & _
                " Where Mod(a.��¼����, 10) = 1 And a.NO = b.NO And a.��¼״̬ = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݵ��ݺ����ж��Ƿ����δ�շѵ�", varPara)
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSql = "Select /*+cardinality(b,10)*/ 1" & vbNewLine & _
                " From ������ü�¼ A,(Select Column_Value As NO From Table(f_str2list([1]))) B" & vbNewLine & _
                " Where Mod(a.��¼����, 10) = 1 And a.NO = b.NO And a.��¼״̬ = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݵ��ݺ����ж��Ƿ����δ�շѵ�", strNos)
    Else
        strSql = "Select 1" & vbNewLine & _
                " From ������ü�¼" & vbNewLine & _
                " Where Mod(��¼����, 10) = 1 And NO = [1] And ��¼״̬ = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݵ��ݺ����ж��Ƿ����δ�շѵ�", strNos)
    End If
    
    If rsTemp.EOF Then
        BillExistNotBalance = False '��ȫ���շ�
    Else
        BillExistNotBalance = True '����δ�շ�
    End If
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBillChargeStatus(ByVal strNos As String, ByRef bytOutStatus As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѵ��ݵļƷ�״̬
    '���:strNOs:ָ���ĵ��ݺ�,������,�ö��ŷ���
    '����:bytOutStatus:0-δ�շ�;1-�����շ�/�˷�;2-ȫ���շ�;3-ȫ���˷�
    '����:��ȡ�ɹ�,����true,���򷵻�False(��δ�ҵ����ݲ���)
    '����:���˺�
    '����:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    '���ܴ���4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_�ַ�, strNos, varPara, strSql, "NO") = False Then Exit Function
        strSql = "Select /*+cardinality(b,10)*/ Sum(a.���� * Nvl(a.����, 1)) As ʣ������," & vbNewLine & _
                "        Sum(Decode(a.��¼����, 1, 1, 0) * Decode(a.��¼״̬, 2, 0, 1) * a.���� * Nvl(a.����, 1)) As ԭʼ����," & vbNewLine & _
                "        Sum(Decode(a.��¼״̬, 0, 1, 0) * a.���� * Nvl(a.����, 1)) As δ������" & vbNewLine & _
                " From ������ü�¼ A,(" & strSql & ") B " & _
                " Where Mod(a.��¼����, 10) = 1 And a.�۸񸸺� Is Null And a.NO = b.NO"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݵ��ݺ����ж��Ƿ��Ѿ��շ�", varPara)
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSql = "Select /*+cardinality(b,10)*/ Sum(a.���� * Nvl(a.����, 1)) As ʣ������," & vbNewLine & _
                "        Sum(Decode(a.��¼����, 1, 1, 0) * Decode(a.��¼״̬, 2, 0, 1) * a.���� * Nvl(a.����, 1)) As ԭʼ����," & vbNewLine & _
                "        Sum(Decode(a.��¼״̬, 0, 1, 0) * a.���� * Nvl(a.����, 1)) As δ������" & vbNewLine & _
                " From ������ü�¼ A,(Select Column_Value As NO From Table(f_str2list([1]))) B " & _
                " Where Mod(a.��¼����, 10) = 1 And a.�۸񸸺� Is Null And a.NO = b.NO"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݵ��ݺ����ж��Ƿ��Ѿ��շ�", strNos)
    Else
        strSql = "Select Sum(���� * Nvl(����, 1)) As ʣ������," & vbNewLine & _
                "        Sum(Decode(��¼����, 1, 1, 0) * Decode(��¼״̬, 2, 0, 1) * ���� * Nvl(����, 1)) As ԭʼ����," & vbNewLine & _
                "        Sum(Decode(��¼״̬, 0, 1, 0) * ���� * Nvl(����, 1)) As δ������" & vbNewLine & _
                " From ������ü�¼" & vbNewLine & _
                " Where Mod(��¼����, 10) = 1 And �۸񸸺� Is Null And NO = [1]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݵ��ݺ����ж��Ƿ��Ѿ��շ�", strNos)
    End If
    
    If Val(Nvl(rsTemp!ԭʼ����)) = 0 Then Exit Function
    If Val(Nvl(rsTemp!ԭʼ����)) = Val(Nvl(rsTemp!δ������)) Then
        bytOutStatus = 0 'δ�շ�
    ElseIf Val(Nvl(rsTemp!ԭʼ����)) = Val(Nvl(rsTemp!ʣ������)) And Val(Nvl(rsTemp!δ������)) = 0 Then
        bytOutStatus = 2 'ȫ���շ�
    ElseIf Val(Nvl(rsTemp!ʣ������)) = 0 Then
        bytOutStatus = 3 'ȫ���˷�
    Else
        bytOutStatus = 1 '�����շ�/�˷�
    End If
    GetBillChargeStatus = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceStatus(ByVal strNos As String, ByRef bytOutStatus As Byte, _
    Optional bln���� As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϼ��ʵ��Ƿ��Ѿ�����(ֻ����ʵ�)
    '���:strNOs:ָ���ĵ��ݺ�,������,�ö��ŷ���
    '     bln����-������ʵ�
    '����:bytOutStatus:0-δ����;1-���ֽ���;2-ȫ������
    '����:��ȡ�ɹ�,����true,���򷵻�False(��δ�ҵ����ݲ���)
    '����:���˺�
    '����:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSql As String, rsTemp As ADODB.Recordset
    Dim strTable As String
    
    bytOutStatus = 0
    On Error GoTo errHandle
    strTable = IIf(bln����, "������ü�¼", "סԺ���ü�¼")
    '���ܴ���4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_�ַ�, strNos, varPara, strSql, "NO") = False Then Exit Function
        
        strSql = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When δ���� <> 0 And ���ʽ�� <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As ���ʱ�־" & vbNewLine & _
        "From (Select /*+Cardinality(B,10)*/" & vbNewLine & _
        "        a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.����Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From ������ü�¼ A, (" & strSql & ") B" & vbNewLine & _
        "       Where a.No = b.No And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.�۸񸸺�, a.���))"

        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݵ��ݺ����ж��Ƿ��Ѿ��շ�", varPara)
        
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSql = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When δ���� <> 0 And ���ʽ�� <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As ���ʱ�־" & vbNewLine & _
        "From (Select /*+Cardinality(B,10)*/" & vbNewLine & _
        "        a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.����Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From ������ü�¼ A, Table(f_Str2list([1])) B" & vbNewLine & _
        "       Where a.No = b.Column_Value And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.�۸񸸺�, a.���))"
        
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID��ȡ��صķ��ý��", strNos)
    Else
        strSql = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When δ���� <> 0 And ���ʽ�� <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As ���ʱ�־" & vbNewLine & _
        "From (Select " & vbNewLine & _
        "        a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.����Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From ������ü�¼ A " & vbNewLine & _
        "       Where a.No = [1] And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.�۸񸸺�, a.���))"

        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݵ��ݺŻ�ȡ���ʵ��Ƿ��Ѿ�����", strNos)
    End If
    bytOutStatus = Val(Nvl(rsTemp!���ʱ�־))
    GetBalanceStatus = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceExpenseDetails(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng����ID As Long, ByRef rsOutDetails As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ʵķ�����ϸ����
    '���:frmMain -����������
    '    lngModule -ģ���
    '    lng����id -����ID
    '����:rsOutDetails-��������(���õ��ţ��շ�����շ����ơ��շ����������ʽ��շѵ��ۡ����㵥λ��ִ�п��ң�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle
    Dim blnNOMoved As Boolean
    
    Set rsOutDetails = Nothing
    blnNOMoved = gobjDatabase.NOMoved("���˽��ʼ�¼", "", "ID", lng����ID, gstrCompentsName & ":�������Ƿ�ת������ʷ��ռ�")
    
   strSql = "" & _
    "   Select A.����ʱ��, A.NO,nvl(�۸񸸺�,���) as ���,A.�շ����,A.�շ�ϸĿID," & _
    "           Avg(Nvl(����,1)) *Avg(����) as ����,A.���㵥λ,sum(A.���ʽ��) as ���ʽ��,sum(a.��׼���� ) as �շѵ���, " & _
    "           a.ִ�в���ID" & _
    "   From " & IIf(blnNOMoved, "H", "") & "������ü�¼ A" & _
    "   Where A.����ID=[1]" & _
    "   Group by A.����ʱ��, A.NO,nvl(�۸񸸺�,���),A.�շ����,A.�շ�ϸĿID,A.���㵥λ,a.ִ�в���ID" & _
    "   Union ALL " & _
    "   Select A.����ʱ��, A.NO,nvl(�۸񸸺�,���) as ���,A.�շ����,A.�շ�ϸĿID," & _
    "           Avg(Nvl(����,1)) *Avg(����) as ����,A.���㵥λ,sum(A.���ʽ��) as ���ʽ��,sum(a.��׼���� ) as �շѵ���, " & _
    "           a.ִ�в���ID" & _
    "   From " & IIf(blnNOMoved, "H", "") & "סԺ���ü�¼ A" & _
    "   Where A.����ID=[1] " & _
    "   Group by A.����ʱ��, A.NO,nvl(�۸񸸺�,���),A.�շ����,A.�շ�ϸĿID,A.���㵥λ,a.ִ�в���ID" & _
    "   "
    strSql = _
    "  Select    A.NO as ���õ���,A.���,A.�շ����,Nvl(E.����,D.����) as �շ�����,A.���� as �շ�����, " & _
    "             a.���ʽ��,a.�շѵ��� ,A.���㵥λ,Nvl(B.����,'δ֪') as ִ�п��� " & _
    " From (" & strSql & ") A,���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
    " Where A.ִ�в���ID=B.ID(+) And A.�շ�ϸĿID=D.ID" & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=3" & _
    " Order by ����ʱ�� Desc,���õ��� Desc,���"
    Set rsOutDetails = gobjDatabase.OpenSQLRecord(strSql, gstrCompentsName & ":���ݽ���ID��ȡ��������", lng����ID)
    GetBalanceExpenseDetails = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function GetBalanceInfor(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng����ID As Long, ByRef rsOutBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����������
    '���:frmMain -����������
    '    lngModule -ģ���
    '    lng����id -����ID
    '����:rsOutDetails-��������( ���㷽ʽ��������������,ҽ�ƿ����ID,���ѿ�,������ˮ��,����˵��,ˢ�����ţ�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle
    Dim blnNOMoved As Boolean
    
    Set rsOutBalance = Nothing
    blnNOMoved = gobjDatabase.NOMoved("���˽��ʼ�¼", "", "ID", lng����ID, gstrCompentsName & ":�������Ƿ�ת������ʷ��ռ�")
    
   strSql = "" & _
    "   Select decode(mod(A.��¼����,10),1,'[��Ԥ��]', A.���㷽ʽ) as ���㷽ʽ,  " & _
    "       ��Ԥ�� as ������,A.�������, " & _
    "       A.�����ID,A.���㿨���,decode(nvl(A.���㿨���,0),0,0,1) as ���ѿ�, " & _
    "       A.������ˮ��,A.����˵��,A.���� as ˢ������ " & _
    "   From " & IIf(blnNOMoved, "H", "") & "����Ԥ����¼ A" & _
    "   Where A.����ID=[1]"
    Set rsOutBalance = gobjDatabase.OpenSQLRecord(strSql, gstrCompentsName & ":���ݽ���ID��ȡ��������", lng����ID)
    GetBalanceInfor = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    IncStr = gobjComlib.zlStr.Increase(strVal)
End Function
Public Function GetInsidePrivs(ByVal lngProg As Long, Optional ByVal blnLoad As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ڲ�ģ���������е�Ȩ��
    '���:lngProg-�����
    '   blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    '����:
    '����:����Ȩ�޴�
    '����:���˺�
    '����:2014-04-09 11:58:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = gobjComlib.GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = gobjDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.����ID = Nvl(rsTmp!����ID, 0)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.���� = Get��Ա����
            UserInfo.רҵ����ְ�� = Getרҵ����ְ��(UserInfo.ID)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.�û���
End Function

Public Function Getרҵ����ְ��(ByVal lng��Աid As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��¼��Ա��רҵ����ְ��
    '����:����ָд��Ա��רҵ����ְ��
    '����:���˺�
    '����:2014-04-09 13:45:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    
 
    strSql = "Select רҵ����ְ�� From ��Ա�� Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��Աרҵְ��", lng��Աid)
    
    Getרҵ����ְ�� = "" & rsTmp!רҵ����ְ��
  
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    '����:������Ա����,����ö��ŷ���
    '����:���˺�
    '����:2014-04-09 13:46:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    If str���� <> "" Then
        strSql = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��Ա����", str����)
    Else
        strSql = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��Ա����", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function GetRoom(str�ű� As String) As String
'���ܣ����ݺű�ķ��﷽ʽ��ȡ�ű������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
            
    strSql = "Select ID,Nvl(���﷽ʽ,0) as ���� From �ҺŰ��� Where ����=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlPublicExpense", str�ű�)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!���� = 0 Then Exit Function '������
    
    '�������
    If rsTmp!���� = 1 Then
        'ָ������
        strSql = "Select �������� From �ҺŰ������� Where �ű�ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlPublicExpense", CLng(rsTmp!ID))
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 2 Then
        '��̬����ø��ű���Һ�δ�������ٵ�����   //todoδ����ԤԼ�Һ�
        strSql = _
            " Select ��������,Sum(NUM) as NUM From (" & _
                " Select ��������,0 as NUM From �ҺŰ������� Where �ű�ID=[1]" & _
                " Union ALL" & _
                " Select ����,Count(����) as NUM From ���˹Һż�¼" & _
                " Where Nvl(ִ��״̬,0)=0 And ��¼����=1 and ��¼״̬=1 and  ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű�=[2]" & _
                " And ���� IN(Select �������� From �ҺŰ������� Where �ű�ID=[1])" & _
                " Group by ����)" & _
            " Group by �������� Order by Num"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlPublicExpense", CLng(rsTmp!ID), str�ű�)
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        strSql = "Select �ű�ID,��������,��ǰ���� From �ҺŰ������� Where �ű�ID=" & rsTmp!ID
        Set rsTmp = New ADODB.Recordset
        Call gobjDatabase.OpenRecordset(rsTmp, strSql, "mdlPublicExpense", adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!��ǰ����), 0, rsTmp!��ǰ����) = 1 Then
                    GetRoom = rsTmp!��������
                    rsTmp!��ǰ���� = 0
                    
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!��ǰ���� = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '�����һ��ƽ������
            If GetRoom = "" Then
                rsTmp.MoveFirst
                GetRoom = rsTmp!��������
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!��ǰ���� = 1
                rsTmp.Update
            End If
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function ReadRegistPrice(ByVal lng��Ŀid As Long, ByVal bln���� As Boolean, ByVal bln���￨ As Boolean, _
    Optional str�ѱ� As String, Optional rsItems As ADODB.Recordset, Optional rsIncomes As ADODB.Recordset, _
    Optional lng����ID As Long, Optional int���� As Integer, Optional str�ű� As String, Optional bytMode As Integer, _
    Optional lng�Һſ���ID As Long = 0, Optional ByVal strPriceGrade As String, Optional strDate As String) As Long
'���ܣ���ȡָ���Һ���Ŀ��Ӧ�ķ�����Ϣ����¼����
'������lng��ĿID=��ʾ�Ƿ��ȡ�Һŷ���(Ҫ���ĹҺ���ĿID)
'      bln����=��ʾ�Ƿ��ȡ����������(���ܽ���ȡ������)
'      bln���￨=��ʾ�Ƿ��ȡ���￨����(��Һŷѻ�����һ����ȡ)
'      str�ѱ�=�Һŷѱ�
'      rsItems(Out)=�����Һ���Ŀ��������Ŀ,������New��ʽ����
'      rsInComes(Out)=����������Ŀ���������,������New��ʽ����
'���أ���ȡ����Ŀ����,ͬʱrsItems,rsInCome=Nothing
'˵������������Ϊ1,����趨���δ���,��Ϊ�̶�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lngԭ��ID As Long
    Dim rsFeeTmp As ADODB.Recordset
    Dim strFee As String
    Dim str������ĿID As String
    Dim strWherePriceGrade As String
    Dim strDateCondition As String
    
    Set rsItems = Nothing
    Set rsIncomes = Nothing
    
    If strDate <> "" Then
        strDateCondition = " [4] "
    Else
        strDateCondition = " Sysdate "
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.�۸�ȼ� = [3]" & vbNewLine & _
            "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [3]" & vbNewLine & _
            "                                   And " & strDateCondition & " Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null "
    End If
    
    '��ȡ�Һ���Ŀ��������Ŀ�ķ���
    If lng��Ŀid <> 0 Then
        strSql = _
            "Select 1 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=[1]" & _
            " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
        strSql = strSql & " Union ALL " & _
            "Select 2 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " D.�������� as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շѴ�����Ŀ D" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=[1]" & _
            " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    
    '��ȡ���������Ѷ�Ӧ�ķ���
    If bln���� Then
        strSql = strSql & IIf(strSql <> "", " Union ALL ", "") & _
            "Select 3 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,A.ִ�п��� as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.�շ�ϸĿID And D.�ض���Ŀ='������'" & _
            " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    
    If bytMode <> 1 And bytMode <> 10 Then
        strFee = "Select zl_Fun_CustomRegExpenses([1],[2],[3]) As ���ӷ� From Dual"
        Set rsFeeTmp = gobjDatabase.OpenSQLRecord(strFee, "zl_Fun_CustomRegExpenses", lng����ID, int����, str�ű�)
        If Not rsFeeTmp.EOF Then
            str������ĿID = Nvl(rsFeeTmp!���ӷ�)
        End If
        
        If str������ĿID <> "" Then
            If strSql = "" Then
                strSql = " " & _
                    "Select /*+cardinality(D,10)*/ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
                    " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,Table(f_str2list([2])) D " & _
                    " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.Column_Value " & _
                    "       And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                            strWherePriceGrade
            Else
                strSql = strSql & " Union ALL " & _
                    "Select /*+cardinality(D,10)*/ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
                    " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,Table(f_str2list([2])) D " & _
                    " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.Column_Value " & _
                    "       And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                            strWherePriceGrade
            End If
            strSql = strSql & " Union ALL " & _
                "Select /*+cardinality(E,10)*/ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
                " D.�������� as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շѴ�����Ŀ D,Table(f_str2list([2])) E" & _
                " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=E.Column_Value " & _
                "       And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                        strWherePriceGrade
        End If
    End If
    
    If strSql = "" Then Exit Function
    
    '������,����,����˳������
    strSql = "Select * From (" & strSql & ") Order by ����,��Ŀ����,�������"
    
    On Error GoTo errH
    If strDate <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng��Ŀid, str������ĿID, strPriceGrade, CDate(strDate))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng��Ŀid, str������ĿID, strPriceGrade)
    End If
    If Not rsTmp.EOF Then
        '�ȴ�����¼��
        Set rsItems = New ADODB.Recordset
        rsItems.Fields.Append "����", adSmallInt '1-����,2-����,3-������,4-���￨��
        rsItems.Fields.Append "ִ�п���ID", adBigInt
        rsItems.Fields.Append "���", adVarChar, 1
        rsItems.Fields.Append "��ĿID", adBigInt
        rsItems.Fields.Append "��Ŀ����", adVarChar, 80
        rsItems.Fields.Append "���㵥λ", adVarChar, 20, adFldIsNullable
        rsItems.Fields.Append "����", adSingle
        rsItems.Fields.Append "������Ŀ��", adSmallInt, , adFldIsNullable
        rsItems.Fields.Append "���մ���ID", adBigInt, , adFldIsNullable
        rsItems.Fields.Append "���ձ���", adVarChar, 80
        
        rsItems.CursorLocation = adUseClient
        rsItems.LockType = adLockOptimistic
        rsItems.CursorType = adOpenStatic
        rsItems.Open
        
        Set rsIncomes = New ADODB.Recordset
        rsIncomes.Fields.Append "��ĿID", adBigInt
        rsIncomes.Fields.Append "������ĿID", adBigInt
        rsIncomes.Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
        rsIncomes.Fields.Append "����", adSingle
        rsIncomes.Fields.Append "Ӧ��", adCurrency
        rsIncomes.Fields.Append "ʵ��", adCurrency
        rsIncomes.Fields.Append "ͳ����", adCurrency, , adFldIsNullable
        rsIncomes.CursorLocation = adUseClient
        rsIncomes.LockType = adLockOptimistic
        rsIncomes.CursorType = adOpenStatic
        rsIncomes.Open
        
        For i = 1 To rsTmp.RecordCount
            '�Һ���Ŀ����
            If lngԭ��ID <> rsTmp!��ĿID Then
                rsItems.AddNew
                rsItems!���� = rsTmp!����
                 '0-����ȷ����,1-�������ڿ���,2-�������ڲ���,3-���������ڿ���,4-ָ������
                If rsTmp!ִ�п������� = -1 Then
                    rsItems!ִ�п���ID = lng�Һſ���ID      '0-��ʾ�Һſ���
                Else
                    rsItems!ִ�п���ID = Get�Һ�ִ�п���ID(rsTmp!��ĿID, rsTmp!ִ�п�������)
                    If rsItems!ִ�п���ID = 0 Then rsItems!ִ�п���ID = lng�Һſ���ID
                End If
                
                rsItems!��� = rsTmp!���
                rsItems!��ĿID = rsTmp!��ĿID
                rsItems!��Ŀ���� = rsTmp!��Ŀ����
                rsItems!���㵥λ = rsTmp!���㵥λ
                rsItems!���� = Format(Nvl(rsTmp!����, 0), "0.000")
                rsItems.Update
            End If
            lngԭ��ID = rsTmp!��ĿID
            
            '������Ŀ����
            rsIncomes.AddNew
            rsIncomes!��ĿID = rsTmp!��ĿID
            rsIncomes!������ĿID = rsTmp!������ĿID
            rsIncomes!�վݷ�Ŀ = rsTmp!�վݷ�Ŀ
            rsIncomes!���� = Format(Nvl(rsTmp!����, 0), "0.00")
            rsIncomes!Ӧ�� = Format(rsItems!���� * rsIncomes!����, "0.00")
            If Nvl(rsTmp!���ηѱ�, 0) = 1 Then
                rsIncomes!ʵ�� = rsIncomes!Ӧ��
            Else
                rsIncomes!ʵ�� = Format(GetActualMoney(str�ѱ�, rsTmp!������ĿID, rsIncomes!Ӧ��, rsTmp!��ĿID), "0.00")
            End If
            rsIncomes.Update
            rsTmp.MoveNext
        Next
        ReadRegistPrice = rsItems.RecordCount
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
    Set rsItems = Nothing
    Set rsIncomes = Nothing
End Function

Public Function Get�Һ�ִ�п���ID(ByVal lng��Ŀid As Long, ByVal intִ�п������� As Integer) As Long
'���ܣ���ȡ�ҺŸ�����Ŀ(������,���￨��)���շ���Ŀ��ִ�п���
'������
'���أ����������,��ʾ�Һſ���(ҽ�����ڿ���)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Get�Һ�ִ�п���ID = UserInfo.����ID
    
    Select Case intִ�п�������
        Case 0 '0-����ȷ����
        Case 1 '1-�������ڿ���
            Get�Һ�ִ�п���ID = 0
        Case 2 '2-�������ڲ���
            Get�Һ�ִ�п���ID = 0
        Case 3 '3-����Ա����
        Case 4 '4-ָ������
            strSql = "Select ִ�п���ID From �շ�ִ�п��� Where �շ�ϸĿID=[1] And Nvl(������Դ,1)=1 "
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng��Ŀid)
            
            If Not rsTmp.EOF Then Get�Һ�ִ�п���ID = rsTmp!ִ�п���ID
        Case 5 'Ժ��ִ��(Ԥ��,������δ��)
        Case 6 '�����˿���
    End Select
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function ActualMoney(str�ѱ� As String, ByVal lng������ĿID As Long, ByVal curӦ�ս�� As Currency, _
    Optional ByVal lng�շ�ϸĿID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal dbl���� As Double, Optional ByVal dbl�Ӱ�Ӽ��� As Double) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�ϸĿID��������ĿID(ǰ������),Ӧ�ս��,���ѱ����õķֶα������۹������ʵ�ս�
    '     ���ҩƷ���ɱ����ձ����������ʵ�ս��
    '���:str�ѱ�=���˷ѱ�����ǰ���̬�ѱ�,�����ʽΪ"���˷ѱ�,��̬�ѱ�1,��̬�ѱ�2,..."
    '      lng�ⷿID,dbl����,��ҩƷ����Ŀ���ɱ��ۼ��մ���ʱ����Ҫ����
    '      dbl����=�����������ڵ��ۼ�����
    '      dbl�Ӱ�Ӽ���=С������,�����Ӧ�ս���Ѱ��Ӱ�Ӽۼ���ʱ��Ҫ�����ڻ�ԭ������
    '����:
    '����:���أ������۹���ͱ��������ʵ�ս��,����Ƕ�̬�ѱ�,��"str�ѱ�"�������Żݷѱ�(ע�����δ���ۼ���,����ԭ������,Ҳ���ܷ��ص�һ��)
    '����:���˺�
    '����:2014-04-09 13:54:17
    '˵��:
    '   ���ɱ��ۼ��ձ������۵����ּ��㷽��(ʵ����һ��)��
    '       1.���۽�� = �ɱ���� * (1 + ���ձ���)
    '       2.���۽�� = �ɱ��� * (1 + ���ձ���) * ��������
    '   ��صļ��㹫ʽ��
    '      �ɱ��� = ҩƷ�ۼ� * (1 - �����)
    '      �ɱ���� = �ۼ۽�� * (1 - �����) = �ɱ��� * ��������
    '      �п����ʱ:����� = ����� / �����,����:����� = ָ�������
    '      ���ڷ���ҩƷ��Ӧÿ���������ηֱ����ɱ��ۺͳɱ����
    '      ����ʱ�۷�����"ҩƷ�ۼ�=Nvl(���ۼ�,ʵ�ʽ��/ʵ������)"��������ʱ��ҩƷ��治��ʱ��������ۼ��㡣
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    
    strSql = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, str�ѱ�, lng�շ�ϸĿID, lng������ĿID, curӦ�ս�� / (1 + dbl�Ӱ�Ӽ���), dbl����, lng�ⷿID)
        
    str�ѱ� = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl�Ӱ�Ӽ���), gSysPara.Money_Decimal.strFormt_VB)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function


Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = True) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
    '���:vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    '����:
    '����:���ظ�ʽ���Ĵ�
    '����:���˺�
    '����:2014-04-09 14:05:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    FormatEx = gobjComlib.FormatEx(vNumber, intBit, blnShowZero)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ģ��Oracle��Decode����
    '����:��������������ֵ
    '����:���˺�
    '����:2014-04-09 14:04:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "") As Long
'���ܣ���ȡ�������ò���ָ��Ʊ��������÷�Χ�ڵ�����ID
'������bytKind      =   Ʊ��
'      intNum       =   Ҫ��ӡ��Ʊ������
'      lngLastUseID =   �ϴ�ʹ�õ�����ID
'      lngShareUseID=   ���ز���ָ���Ĺ���ID
'      strBill      =   ��ǰƱ�ݺţ����ڼ���������ε�Ʊ�ݷ�Χ
'      strUseType-ʹ�����
'���أ�
'      >0   =   �ɹ������õ�����ID
'      =0   =   ʧ��
'      -1   =   û������(����򲻹�����δ����),δ���ù���
'      -2   =   û������(����򲻹�����δ����),���õĹ���������򲻹�
'      -3   =   ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
'      -4   =   ָ�����ε�Ʊ�ݲ�����
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo errH
    '1.�ϴε����������Ƿ���ò�����
    If lngLastUseID > 0 Then
        strSql = "" & _
        "   Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
        "   From Ʊ�����ü�¼ " & _
        "   Where Ʊ��=[1] And ʣ������>=[2] And ID=[3]  " & _
        "           And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL) "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, intNum, lngLastUseID, IIf(Trim(strUseType) = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then    'Ŀǰ��Ʊ�ݺſ��ܺ��ϴβ�ͬ��������Ҫ��鷶Χ
                If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '����û�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
                
            ElseIf intNum > 1 Then  '����ȷ���������ε���ʱ,��ǰƱ�ݺ��������β�����
                GetInvoiceGroupID = -4: Exit Function
            End If
        End With
    End If
    
    '2.�ϴε��������β����û򲻿���ʱ,ȡ������Ĳ������õ�
    '  �ж��������ʹ�õ�����,�ٵ�����,��������
    strSql = "" & _
    "   Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����" & vbNewLine & _
    "   From Ʊ�����ü�¼" & vbNewLine & _
    "   Where Ʊ�� = [1] And ʣ������ >= [2] And ������ = [3]  " & _
    "           And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL ) " & _
    "           And ʹ�÷�ʽ = 1" & vbNewLine & _
    "   Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� desc, ��ʼ����"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, intNum, UserInfo.����, IIf(strUseType = "", "LXH", strUseType))
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then GetInvoiceGroupID = !ID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
            blnTmp = False
            strPre = "" & !ǰ׺�ı�
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = !ID: Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.û�����õ�,ʹ�ñ��ز���ָ���Ĺ�������
    If lngShareUseID > 0 Then
        strSql = "" & _
        "   Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
        "   From Ʊ�����ü�¼  " & _
        "   Where Ʊ��=[1] And ʣ������>=[2] And ID=[3] " & _
        "   And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL) "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, intNum, lngShareUseID, IIf(strUseType = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    GetInvoiceGroupID = lngReturn   '����δ�ҵ���ԭ�����
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function CheckUsedBill(bytKind As Byte, ByVal lng����ID As Long, _
    Optional ByVal strBill As String, _
     Optional ByVal strUseType As String = "") As Long
    '���ܣ���鵱ǰ����Ա�Ƿ��п���Ʊ������(���û���),�����ؿ��õ�����ID
    '������bytKind=Ʊ��
    '      lng����ID=��һ�μ��ʱΪ�������õĹ�������ID,�Ժ�Ϊ�ϴ�ʹ�õ�����ID
    '      strBill=Ҫ��鷶Χ��Ʊ�ݺ�
    '˵����
    '    1.�ڼ�鷶Χʱ,��������ж�������Ʊ��,��ֻҪ������һ��֮�о�����
    '    2.�ڼ�鷶Χʱ,����Ҳ�ڼ�鷶Χ֮�ڡ�
    '    3.���ж�������ʱ,ȱʡ���ٵ�����,��������,"���ʹ�õ�����"ԭ��
    '���أ�
    '      ������Ʊ������ID>0
    '      0=ʧ��
    '      -1:û������(�����δ����)��Ҳû�й���(δ����)
    '      -2:���õĹ���������
    '      -3:ָ��Ʊ�ݺŲ��ڵ�ǰ���÷�Χ��(������������Ʊ�ݵ����)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSql As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    
    '����Ա��ʣ�������Ʊ�ݼ�
    strSql = _
        "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
        "From Ʊ�����ü�¼" & vbNewLine & _
        "Where Ʊ�� = [1] And ʹ�÷�ʽ = 1 And ʣ������ > 0 And ������ = [2] And (Nvl(ʹ�����,'LXH')=[3] or  ʹ����� is NULL)" & vbNewLine & _
        "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� Desc, ��ʼ����"
    Set rsSelf = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, UserInfo.����, IIf(strUseType = "", "LXH", strUseType))
    If lng����ID = 0 Then
        '�����е�һ�μ��,��û�����ñ��ع���
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function 'Ҳû������Ʊ��
        '������Ʊ��,������ԭ�򷵻�
        lngReturn = rsSelf!ID
    Else
        '�ϴ�ʹ�õ�����ID���һ�μ��Ĺ���ID,���ж�����
        strSql = "Select ID,ʹ�÷�ʽ,ʣ������,ǰ׺�ı�,��ʼ����,��ֹ���� From Ʊ�����ü�¼ Where Ʊ��=[1]  And (Nvl(ʹ�����,'LXH')=[3] or  ʹ����� is NULL) And ID=[2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, lng����ID, IIf(strUseType = "", "LXH", strUseType))
        '����26352 by ���ջ� 2009-11-20
        If rsTmp.EOF Then CheckUsedBill = -2: Exit Function
        
        If rsTmp!ʹ�÷�ʽ = 2 Then '����,Ҫ�ȿ���û������
            If Not rsSelf.EOF Then
                '�����õģ�����
                lngReturn = rsSelf!ID
            Else
                'û������ȡ����
                If rsTmp!ʣ������ = 0 Then CheckUsedBill = -2: Exit Function '�����Ѿ�����
                lngReturn = rsTmp!ID
                blnTmp = True
            End If
        Else
            '����Ʊ��
            If rsTmp!ʣ������ > 0 Then
                '��ʣ��
                lngReturn = rsTmp!ID
            Else
                '������ʣ�������
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '��������Ҳû��ʣ��
                lngReturn = rsSelf!ID
            End If
        End If
    End If
    
    '���Ʊ�ŷ�Χ�Ƿ���ȷ
    If strBill <> "" Then
        If blnTmp Then
            '�ڹ��÷�Χ�ڷ�Χ�ж�
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!��ʼ����) And UCase(strBill) <= UCase(rsTmp!��ֹ����) And Len(strBill) = Len(rsTmp!��ʼ����)) Then
                lngReturn = -3
            End If
        Else
            '�ڿ������÷�Χ���ж�
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '����������,�������������м��
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
    CheckUsedBill = 0
End Function

Public Function GetNextBill(lng����ID As Long) As String
'���ܣ�������������ID,��ȡ��һ��ʵ��Ʊ�ݺ�
'˵����1.��ȡ������Χ�ڵ���ЧƱ��ʱ,���ؿ����û�����
'      2.�ſ��ѱ���ĺ���
    Dim rsMain As ADODB.Recordset
    Dim rsDelete As ADODB.Recordset
    Dim strSql As String, strBill As String
    
    On Error GoTo errH
    
    strSql = "Select ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ����" & _
        " From Ʊ�����ü�¼ Where ʣ������>0 And ID=[1]"
    Set rsMain = gobjDatabase.OpenSQLRecord(strSql, "ȡһ��Ʊ�ݺ�", lng����ID)
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!��ǰ����) Then
        strBill = UCase(rsMain!��ʼ����)
    Else
        strBill = UCase(gobjCommFun.IncStr(rsMain!��ǰ����))
    End If
    
     '�����:25448
     '���˺�:ȡ����;����=1 And ԭ��=5 And ���:ԭ���ǿ��ܴ����Ѿ�ʹ���˵�Ʊ��,ʹ���˵�,���ų�
     'Ʊ��: 1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
     '����:1-����(ԭ����1��3��5��������)��2-�ջ�(ԭ����2��4��������)
     'ԭ��:1-��������Ʊ�ݣ�2-�����ջط�Ʊ��3-�ش򷢳�Ʊ�ݣ�4-�ش��ջ�Ʊ�ݣ�5-��������Ʊ��
     
    strSql = "Select Upper(����) as ���� From Ʊ��ʹ����ϸ" & _
        " Where ����||''>=[1] And ����ID=[2]" & _
        " Order by ����"
        
    Set rsDelete = gobjDatabase.OpenSQLRecord(strSql, "ȡһ��Ʊ�ݺ�", strBill, lng����ID)
    Do While True
        '��鷶Χ
        If Left(strBill, Len("" & rsMain!ǰ׺�ı�)) <> UCase("" & rsMain!ǰ׺�ı�) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!��ʼ����) And strBill <= UCase(rsMain!��ֹ����)) Then
            Exit Function
        End If
                
        '�ſ������
        rsDelete.Filter = "����='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = gobjCommFun.IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������ڼ�,�������������ڴ�(yyyy-MM-dd[ HH:mm])
    '���:strText-�����ı�
    '     blnTime=�Ƿ���ʱ�䲿��
    '����:
    '����:�������������ڴ�(yyyy-MM-dd[ HH:mm])
    '����:���˺�
    '����:2014-04-09 14:03:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = gobjDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '���봮�а������ڷָ���
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                'ֻ���������ڲ���
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                'ֻ������ʱ�䲿��
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '����Ƿ�����,����ԭ����
            strTmp = strText
        End If
    Else
        '���������ڷָ���
        If Len(strTmp) <= 2 Then
            '��������dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '��������MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '��������yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '��������MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '��������yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '��������yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) And Not blnTime Then
        strTmp = Format(strTmp, "yyyy-MM-dd")
    End If
    GetFullDate = strTmp
End Function
Public Function NeedName(strList As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ж��Իس����ָ�
    '���:strList:1-strList��()��[]�ָ����������ʱ��������[����]��(����)��ͷ,�������Ϊ���ֻ���ĸ
    '     2-�ָ��������ȼ����س���(Chr(13)��> - > [] > ()
    '����:
    '����: ��ȡ����
    '����:���˺�
    '����:2014-04-09 14:03:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    NeedName = gobjComlib.zlStr.NeedName(strList)
    
End Function
Public Function BillExistBalance(ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ�����շѻ��۵��Ƿ�����Ѿ��շѵ�����
    '���:strNO-���ݺ�
    '����:
    '����:���շѷ���true,���򷵻�False
    '����:���˺�
    '����:2014-04-09 14:12:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    
    strSql = "Select ID From ������ü�¼ Where ��¼����=1 And ��¼״̬ IN(1,3) And NO=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "BillExistBalance", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function


Public Function ExistIOClass(bytBill As Byte) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ����ָ�������������͵�������
    '����:����������ID
    '����:���˺�
    '����:2014-04-09 14:17:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ���ID From ҩƷ�������� Where ����=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!���ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function GetBillMax���(ByVal strNO As String, ByVal int��¼���� As Integer, str�Ǽ�ʱ�� As String, int������Դ As Integer) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ݵ�ǰ��������+1
    '���:str�Ǽ�ʱ��=���ҽ��ֻ�����˲���������ʱ����Ҫ�����ɵ��շѻ��۵�(NO��ͬ)��ʱ���������ɵ�һ�¡�
    '     int������Դ:1-���2-סԺ
    '����:
    '����:���ص�ǰ������+1
    '����:���˺�
    '����:2014-04-09 14:18:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strTab As String
    
    strTab = IIf(int��¼���� = 1 Or (int��¼���� = 2 And int������Դ = 1), "������ü�¼", "סԺ���ü�¼")
    On Error GoTo errHandle
    
    str�Ǽ�ʱ�� = ""
    strSql = "Select Max(���) as ���,Max(�Ǽ�ʱ��) as ʱ�� From " & strTab & " Where NO=[1] And ��¼����=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", strNO, int��¼����)
    If Not rsTmp.EOF Then
        GetBillMax��� = Nvl(rsTmp!���, 0) + 1
        If Not IsNull(rsTmp!ʱ��) Then
            str�Ǽ�ʱ�� = Format(rsTmp!ʱ��, "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        GetBillMax��� = 1
    End If
    Exit Function
    
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function
Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��0��ת��Ϊ"NULL"��,������SQL���ʱ��
    '���:blnForceNum=��ΪNullʱ���Ƿ�ǿ�Ʊ�ʾΪ������
    '����:
    '����:����������SQL���
    '����:���˺�
    '����:2014-04-09 14:23:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZVal = gobjComlib.ZVal(varValue, blnForceNum)
End Function


Public Function AnalyseComputer() As String
    AnalyseComputer = gobjComlib.OS.ComputerName
End Function

Public Function GetPatiDayMoney(lng����ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵��췢���ķ����ܶ�
    '����:���ز��˵ĵ��շ����ܶ�
    '����:���˺�
    '����:2014-04-09 14:59:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    On Error GoTo errH
    strSql = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function zlPatiCardCheck(ByVal byt���ó��� As Byte, lng����ID As Long, str���� As String, bytˢ����ʽ As Byte) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鲡��ˢ����ʽ
    '��Σ�byt���ó���: 1-�Һ�;2-�շ�
    '         lng����ID:����ID(δ������,������)
    '         str����;δˢ��ʱ,Ϊ��
    '         bytˢ����ʽ: 1-����ˢ��;2-ҽ��ˢ��
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-04-27 16:09:08
    '˵����һ�����ŵ����ݲ��ˣ�ʹ�õ�ҽ����ͬʱҲ�Ǿ��￨��ҽԺҪ�������ҽ����ʽ����
    '          �����֤�Һš��շѣ����������Էѷ�ʽֱ��ˢ�����У����Ҫ���ڹҺš��շ�ʱ�����ݲ���ˢ�������������ҽ�������֤��ʽˢ�Ŀ���
    '          ����ֱ��ˢ�Ŀ�������ʾ�������������
    '����:29283
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    strSql = " Select Zl_Paticardcheck([1],[2],[3],[4]) as ��ʾ��Ϣ From Dual "
    ' Zl_Paticardcheck
    '  ���ó���_IN NUMBER ,
    '  ����id_In Number,
    '  ����_In   Varchar2,
    '  ˢ����ʽ_In Number:=1
    On Error GoTo errHandle
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��鲡��ˢ����ʽ�Ƿ�Ϸ�", byt���ó���, lng����ID, str����, bytˢ����ʽ)
    strSql = Nvl(rsTemp!��ʾ��Ϣ)
    If strSql <> "" Then
        MsgBox strSql, vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    zlPatiCardCheck = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str���� As String, ByVal curʣ���� As Currency, _
    ByVal cur���ս�� As Currency, ByVal cur���ʽ�� As Currency, ByVal cur������� As Currency, _
    ByVal str�շ���� As String, ByVal str������� As String, str�ѱ���� As String, _
    intWarn As Integer, Optional ByVal bln���� As Boolean, _
    Optional blnNotCheck��� As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Բ��˼��ʽ��б�����ʾ
    '���:rsWarn=���������������õļ�¼��(�ò��˲���,�����ֺ���ҽ��)
    '     str�շ����=��ǰҪ�������,���ڷ��౨��
    '     str�������=�������,������ʾ
    '     bln����=���ɻ��۷���ʱ�ı��������ƾ���Ƿ��ǿ�Ƽ���Ȩ��ʱ�Ĵ���
    '     intWarn=�Ƿ���ʾѯ���Ե���ʾ,-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��
    '     blnNotCheck���:���������м��(��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
    '����:
    '����:intWarn=����ѯ������ʾ�е�ѡ����,0=Ϊ��,1-Ϊ��
    '     0;û�б���,����
    '     1:������ʾ���û�ѡ�����
    '     2:������ʾ���û�ѡ���ж�
    '     3:������ʾ�����ж�
    '     4:ǿ�Ƽ��ʱ���,����
    '����:���˺�
    '����:2014-04-09 15:00:33
    '˵��:str�ѱ����="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim bln�ѱ��� As Boolean, byt��־ As Byte
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str���� As String, i As Long
    
    BillingWarn = 0
    
    '�����������:NULL��û������,0�������˵�
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str�շ����) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־1 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str�շ����) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־2 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str�շ����) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־3 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    'ʾ����"-" �� ",ABC,567,DEF"
    '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
    bln�ѱ��� = InStr(str�ѱ����, str�շ����) > 0 Or str�ѱ���� Like "-*"
    
    If bln�ѱ��� Then '��intWarn = -1ʱ,Ҳ��ǿ���ٱ���
        If byt��־ = 2 Then
            If str�ѱ���� Like "-*" Then
                byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
            Else
                arrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str�շ����) > 0 Then
                        byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                        'Exit For 'ȡ��˵����סԺ����ģ��
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str������� <> "" Then str������� = """" & str������� & """����"
    str���� = IIf(cur������� = 0, "", "(��������:" & Format(cur�������, "0.00") & ")")
    curʣ���� = curʣ���� + cur������� - cur���ʽ��
    cur���ս�� = cur���ս�� + cur���ʽ��
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If curʣ���� < 0 Then
                        byt��ʽ = 2
                        If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�,����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    ElseIf curʣ���� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If curʣ���� < 0 Then
                            byt��ʽ = 2
                            If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�,����ò��˼�����", frmParent)
                                    If vMsg = vbNo Or vMsg = vbCancel Then
                                        If vMsg = vbCancel Then intWarn = 0
                                        BillingWarn = 2
                                    ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                        If vMsg = vbIgnore Then intWarn = 1
                                        BillingWarn = 4
                                    End If
                                Else
                                    If intWarn = 0 Then
                                        BillingWarn = 2
                                    ElseIf intWarn = 1 Then
                                        BillingWarn = 4
                                    End If
                                End If
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gSysPara.Money_Decimal.strFormt_VB) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gSysPara.Money_Decimal.strFormt_VB) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gSysPara.Money_Decimal.strFormt_VB) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gSysPara.Money_Decimal.strFormt_VB) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־3
            End If
        End If
    End If
End Function


Public Function zlIsCheckMedicinePayMode(ByVal strҽ�Ƹ������� As String, _
    Optional ByRef blnҽ�� As Boolean, Optional ByRef bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ�Ƹ��ʽ�Ƿ񹫷ѻ�ҽ��
    '���:strҽ�Ƹ�������-ҽ�Ƹ�������
    '����:blnҽ��-true,��ʾҽ��
    '        bln����-true,��ʾ�ǹ���
    '����:��ҽ���򹫷�ҽ��,����true,���򷵻�False
    '����:���˺�
    '����:2012-01-17 16:25:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "": blnҽ�� = False: bln���� = False
    If grsҽ�Ƹ��ʽ Is Nothing Then
        strSql = "Select ����,����,����,ȱʡ��־,�Ƿ�ҽ��,�Ƿ񹫷� From ҽ�Ƹ��ʽ"
    ElseIf grsҽ�Ƹ��ʽ.State <> 1 Then
        strSql = "Select ����,����,����,ȱʡ��־,�Ƿ�ҽ��,�Ƿ񹫷� From ҽ�Ƹ��ʽ"
    End If
    If strSql <> "" Then
        Set grsҽ�Ƹ��ʽ = gobjDatabase.OpenSQLRecord(strSql, "��ȡҽ�Ƹ��ʽ")
    End If
    grsҽ�Ƹ��ʽ.Find "����='" & strҽ�Ƹ������� & "'", , adSearchForward, 1
    If grsҽ�Ƹ��ʽ.EOF Then Exit Function
    blnҽ�� = Val(Nvl(grsҽ�Ƹ��ʽ!�Ƿ�ҽ��)) = 1
    bln���� = Val(Nvl(grsҽ�Ƹ��ʽ!�Ƿ񹫷�)) = 1
    zlIsCheckMedicinePayMode = blnҽ�� Or bln����
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function
Public Function ShowHelp(ByVal ChmName As String, SHwnd As Long, ByVal htmName As String, Optional Sys As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��������
    '����:ChmName:CHM��ʽ�ļ�(Ŀǰ�������:App.ProductName)
    '     SHwnd:���봰�ھ��(��Ϊ��������)
    '     htmName:��ӳ��CHM�е�htm�ļ�����
    '����:���˺�
    '����:2014-05-15 15:49:52
    '-----------------------------------------------------------------------------------------------------------------------------
    ShowHelp = gobjComlib.ShowHelp(ChmName, SHwnd, htmName, Sys)
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ������״̬�����󶥱߽糬��ʱ�����Զ�����Ϊ0
    '���:objForm:Ҫ�ָ��Ĵ���
    '     strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
    '     strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-15 15:53:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
   RestoreWinState = gobjComlib.RestoreWinState(objForm, strProjectName, strUserDef)
End Function

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���洰�弰���и��ֿؼ���״̬
    '���: objForm:Ҫ����Ĵ���
    '      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
    '      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    '����:���˺�
    '����:2014-05-15 15:55:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
   SaveWinState = gobjComlib.SaveWinState(objForm, strProjectName, strUserDef)
End Function
Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������ض���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-15 15:34:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function
 



Public Function zlGetDefaultWindow(ByVal str��� As String, ByVal lngҩ��ID As Long, _
    ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ��ҩ����������
    '���:str���-�շ����
    '     lngҩ��ID-ҩ��ID
    '     lngModule-ģ���
    '����:
    '����:����ȱʡ�ķ�ҩ����
    '����:���˺�
    '����:2014-07-23 18:38:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    Dim str���� As String, lng��ҩ�� As Long
    Dim str�ɴ� As String, lng��ҩ�� As Long
    Dim str�д� As String, lng��ҩ�� As Long
    Select Case str���
        Case "5"
            str���� = gobjDatabase.GetPara("��ҩ������", glngSys, lngModule)
            If lngModule = 1252 Then
                lng��ҩ�� = Val(gobjDatabase.GetPara("����ȱʡ��ҩ��", glngSys, lngModule))
            Else
                lng��ҩ�� = Val(gobjDatabase.GetPara("ȱʡ��ҩ��", glngSys, lngModule))
            End If
            If InStr(str����, ":") > 0 Then '������û�д�ҩ��ID
                 strTmp = str����
            ElseIf lng��ҩ�� > 0 And str���� <> "" Then
                strTmp = lng��ҩ�� & ":" & str����
            End If
        Case "6"
            str�ɴ� = gobjDatabase.GetPara("��ҩ������", glngSys, lngModule)
            If lngModule = 1252 Then
                lng��ҩ�� = Val(gobjDatabase.GetPara("����ȱʡ��ҩ��", glngSys, lngModule))
            Else
                lng��ҩ�� = Val(gobjDatabase.GetPara("ȱʡ��ҩ��", glngSys, lngModule))
            End If
            If InStr(str�ɴ�, ":") > 0 Then
                 strTmp = str�ɴ�
            ElseIf lng��ҩ�� > 0 And str�ɴ� <> "" Then
                 strTmp = lng��ҩ�� & ":" & str�ɴ�
            End If
        Case "7"
            str�д� = gobjDatabase.GetPara("��ҩ������", glngSys, lngModule)
            If lngModule = 1252 Then
                lng��ҩ�� = Val(gobjDatabase.GetPara("����ȱʡ��ҩ��", glngSys, lngModule))
            Else
                lng��ҩ�� = Val(gobjDatabase.GetPara("ȱʡ��ҩ��", glngSys, lngModule))
            End If
            If InStr(str�д�, ":") > 0 Then
                 strTmp = str�д�
            ElseIf lng��ҩ�� > 0 And str�д� <> "" Then
                 strTmp = lng��ҩ�� & ":" & str�д�
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str���
                Case "5"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    zlGetDefaultWindow = strTmp
End Function

Public Function zlGet��ҩ����(ByVal lngModule As Long, ByVal curDate As Date, ByVal lngҩ��ID As Long, ByVal str��� As String, _
    str���� As String, str�ɴ� As String, str�д� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ��Ӧ�ķ�ҩ����
    '���:lngҩ��ID=ִ�в���ID
    '     curDate=��ǰʱ��
    '����:����ҩƷ��Ӧ�ķ�ҩ����
    '����:���˺�
    '����:2014-07-23 18:40:35
    '˵��:��ͬһ������ҩ���ķ�ҩ������ƽ������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng��ҩ�� As Long, lng��ҩ�� As Long, lng��ҩ�� As Long
    
    On Error GoTo errH
    
    'ָ��ʱ�̶�����(ָ����ָû�ж�Ӧҩ���ϰ�ʱָ��)
    Select Case str���
        Case "5"
            lng��ҩ�� = Val(gobjDatabase.GetPara(18, glngSys, lngModule))

            If str���� <> "" Then
                zlGet��ҩ���� = str����
            ElseIf lng��ҩ�� > 0 Then
                zlGet��ҩ���� = zlGetDefaultWindow(str���, lngҩ��ID, lngModule)
                str���� = zlGet��ҩ����
            End If
        Case "6"
            lng��ҩ�� = Val(gobjDatabase.GetPara(19, glngSys, lngModule))
            If str�ɴ� <> "" Then
                zlGet��ҩ���� = str�ɴ�
            ElseIf lng��ҩ�� > 0 Then
                zlGet��ҩ���� = zlGetDefaultWindow(str���, lngҩ��ID, lngModule)
                str�ɴ� = zlGet��ҩ����
            End If
        Case "7"
            lng��ҩ�� = Val(gobjDatabase.GetPara(20, glngSys, lngModule))
            If str�д� <> "" Then
                zlGet��ҩ���� = str�д�
            ElseIf lng��ҩ�� > 0 Then
                zlGet��ҩ���� = zlGetDefaultWindow(str���, lngҩ��ID, lngModule)
                str�д� = zlGet��ҩ����
            End If
    End Select
    
    
    If zlGet��ҩ���� <> "" Then
        strSql = "Select ���� From ��ҩ���� Where �ϰ��=1 And ҩ��ID=[1] And ����=[2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlOutExse", lngҩ��ID, zlGet��ҩ����)
        If rsTmp.EOF Then zlGet��ҩ���� = ""
        Exit Function
    End If
    
    '��̬�����ϰ�ķ�ר�Ҵ���,98876
    strSql = "Select Zl_Get��ҩ����([1],[2],[3]) As ���� From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��ҩ����", lngҩ��ID, Val(gobjDatabase.GetPara(19, glngSys, , 0)), curDate)
    If Not rsTmp.EOF Then
        zlGet��ҩ���� = Nvl(rsTmp!����)
    End If
    
    If zlGet��ҩ���� <> "" Then
        Select Case str���
            Case "5"
                str���� = zlGet��ҩ����
            Case "6"
                str�ɴ� = zlGet��ҩ����
            Case "7"
                str�д� = zlGet��ҩ����
        End Select
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function GetActualMoney(ByVal str�ѱ� As String, ByVal lng����ID As Long, ByVal curӦ�� As Currency, ByVal lng�շ�ϸĿID As Long) As Currency
'���ܣ�����ָ���ķѱ��������Ŀ���շ���Ŀ,����ָ������ʵ���տ���
'������
'   str�ѱ�   ���ѱ�
'   lng����ID  ��������ĿID
'   curӦ�գ�Ӧ�ս��ֵ
'���أ�ʵ��Ӧ�յĽ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
        
    strSql = "Select ʵ�ձ���" & vbNewLine & _
            "From �ѱ���ϸ" & vbNewLine & _
            "Where �ѱ� = [1] And �շ�ϸĿid = [3] And Abs([4]) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select ʵ�ձ���" & vbNewLine & _
            "From �ѱ���ϸ A" & vbNewLine & _
            "Where �ѱ� = [1] And ������Ŀid = [2] And Abs([4]) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And Not Exists" & vbNewLine & _
            " (Select 1 From �ѱ���ϸ C Where C.�ѱ� = A.�ѱ� And C.�շ�ϸĿid = [3])"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, str�ѱ�, lng����ID, lng�շ�ϸĿID, curӦ��)
    If rsTmp.EOF Then
        GetActualMoney = curӦ��
    Else
        GetActualMoney = curӦ�� * rsTmp!ʵ�ձ��� / 100
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function Getδ��ҩƷ��ҩ����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As String
    '-------------------------------------------------------------------------
    '���ܣ��жϵ�ǰ�����Ƿ������ִͬ�в��ŵ�δ��ҩƷ���������򷵻�δ��ҩƷ�ķ�ҩ����
    '���أ���������ִͬ�в��ŵ�δ��ҩƷ���򷵻�δ��ҩƷ�ķ�ҩ���ڣ����򷵻ؿ�
    '���ƣ�Ƚ����
    '���ڣ�2014-04-09
    '���⣺71902
    '˵����
    '   ͬһ���˲��˲�ͬʱ��ζ��ŵ����շѣ�����ͬһ����ҩ���ڣ����㲡��ȡҩ
    '-------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    strSql = "Select ��ҩ����" & vbNewLine & _
            "From δ��ҩƷ��¼" & vbNewLine & _
            "Where ���� = 8 And ��ҩ���� Is Not Null And ����id = [1] And �ⷿid = [2]" & vbNewLine & _
            "Order By ���շ� Desc, �������� Desc"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ����δ��ҩƷ��ҩ����", lng����ID, lngִ�в���ID)
    
    If Not rsTemp.EOF Then
        Getδ��ҩƷ��ҩ���� = Nvl(rsTemp!��ҩ����)
    End If
    rsTemp.Close: Set rsTemp = Nothing
    
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function zlGetDrugWindow(ByVal lngModule As Long, ByVal lngҩ��ID As Long, ByVal str��� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ�ķ�ҩ����,�������ָ����ȱʡ,����ָ��Ϊ׼,����,����ǻ��۵�,���Ե�һҩƷ�еĴ���Ϊ׼,��������������ͬҩƷ�Ĵ���Ϊ׼
    '����:���ط�ҩ����
    '����:���˺�
    '����:2014-07-23 18:49:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��ҩ���� As String
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim p As Integer, i As Integer, varData As Variant, varTemp As Variant
    Err = 0: On Error GoTo errH:
    str��ҩ���� = zlGetDefaultWindow(str���, lngҩ��ID, lngModule)
    If str��ҩ���� = "" Then Exit Function
    strSql = "Select ���� From ��ҩ���� Where �ϰ��=1 And ҩ��ID=[1] And ����=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡȱʡ��ҩ����", lngҩ��ID, str��ҩ����)
    If rsTmp.EOF Then Exit Function
    zlGetDrugWindow = str��ҩ����
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function
Public Function zlAddUpdateSwapSQL(ByVal blnԤ�� As Boolean, ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    str���� As String, str������ˮ�� As String, str����˵�� As String, _
    ByRef cllPro As Collection, Optional intУ�Ա�־ As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������ˮ�ź���ˮ˵��
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    '����:cllPro-����SQL��
    '����:���˺�
    '����:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    strSql = "Zl_�����ӿڸ���_Update("
    '  �����id_In   ����Ԥ����¼.�����id%Type,
    strSql = strSql & "" & lng�����ID & ","
    '  ���ѿ�_In     Number,
    strSql = strSql & "" & IIf(bln���ѿ�, 1, 0) & ","
    '  ����_In       ����Ԥ����¼.����%Type,
    strSql = strSql & "'" & str���� & "',"
    '  ����ids_In    Varchar2,
    strSql = strSql & "'" & strIDs & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    strSql = strSql & "'" & str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type
    strSql = strSql & "'" & str����˵�� & "',"
    'Ԥ����ɿ�_In Number := 0
    strSql = strSql & "" & IIf(blnԤ��, 1, 0) & ","
    '�˷ѱ�־ :1-�˷�;0-����
    strSql = strSql & "0,"
    'У�Ա�־
    strSql = strSql & "" & IIf(intУ�Ա�־ = 0, "NULL", intУ�Ա�־) & ")"
    zlAddArray cllPro, strSql
End Function
'
'Public Function zlAddThreeSwapSQLToCollection(ByVal blnԤ���� As Boolean, _
'    ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
'    ByVal str���� As String, strExpend As String, ByRef cllPro As Collection) As Boolean
'
'
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:����������������
'    '���: blnԤ����-�Ƿ�Ԥ����
'    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
'    ' ����:cllPro-����SQL��
'    '����:�ɹ�,����true,���򷵻�False
'    '����:���˺�
'    '����:2011-07-19 10:23:30
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
'    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long
'
'    Err = 0: On Error GoTo Errhand:
'    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
'    'strExpend:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
'    varData = Split(strExpend, "||")
'    Dim str������Ϣ As String, strTemp As String
'    For i = 0 To UBound(varData)
'        If Trim(varData(i)) <> "" Then
'            varTemp = Split(varData(i) & "|", "|")
'            If varTemp(0) <> "" Then
'                strTemp = varTemp(0) & "|" & varTemp(1)
'                If gobjCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
'                    str������Ϣ = Mid(str������Ϣ, 3)
'                    'Zl_�������㽻��_Insert
'                    strSQL = "Zl_�������㽻��_Insert("
'                    '�����id_In ����Ԥ����¼.�����id%Type,
'                    strSQL = strSQL & "" & lng�����ID & ","
'                    '���ѿ�_In   Number,
'                    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
'                    '����_In     ����Ԥ����¼.����%Type,
'                    strSQL = strSQL & "'" & str���� & "',"
'                    '����ids_In  Varchar2,
'                    strSQL = strSQL & "'" & strIDs & "',"
'                    '������Ϣ_In Varchar2:������Ŀ|��������||...
'                    strSQL = strSQL & "'" & str������Ϣ & "',"
'                    'Ԥ����ɿ�_In Number := 0
'                    strSQL = strSQL & IIf(blnԤ����, "1", "0") & ")"
'                    zlAddArray cllPro, strSQL
'                    str������Ϣ = ""
'                End If
'                str������Ϣ = str������Ϣ & "||" & strTemp
'            End If
'        End If
'    Next
'    If str������Ϣ <> "" Then
'        str������Ϣ = Mid(str������Ϣ, 3)
'        'Zl_�������㽻��_Insert
'        strSQL = "Zl_�������㽻��_Insert("
'        '�����id_In ����Ԥ����¼.�����id%Type,
'        strSQL = strSQL & "" & lng�����ID & ","
'        '���ѿ�_In   Number,
'        strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
'        '����_In     ����Ԥ����¼.����%Type,
'        strSQL = strSQL & "'" & str���� & "',"
'        '����ids_In  Varchar2,
'        strSQL = strSQL & "'" & strIDs & "',"
'        '������Ϣ_In Varchar2:������Ŀ|��������||...
'        strSQL = strSQL & "'" & str������Ϣ & "',"
'        'Ԥ����ɿ�_In Number := 0
'        strSQL = strSQL & IIf(blnԤ����, "1", "0") & ")"
'        zlAddArray cllPro, strSQL
'    End If
'    zlAddThreeSwapSQLToCollection = True
'    Exit Function
'Errhand:
'    If gobjComlib.ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Public Function zlFormatNum(ByVal strMoney As String) As String
    strMoney = Replace(strMoney, Chr(44), "")
    zlFormatNum = strMoney
End Function

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str�������� As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�������,���ò�ͬ�������͵���ʾ��ɫ
    '���:objPatiControl-���˿ؼ�(�ı���,��ǩ)
    '    str��������-��������
    '    lngDefaultColor-ȱʡ���˵���ʾ��ɫ
    '����:True-������ɫ�ɹ���False-ʧ��
    '����:���ϴ�
    '����:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str�������� <> "" Then
        lngColor = gobjDatabase.GetPatiColor(str��������)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function GetMoneyInfoRegist(lng����ID As Long, Optional dblModiMoney As Double, _
    Optional blnInsure As Boolean, _
    Optional int���� As Integer = -1, _
    Optional bln������ͳ�� As Boolean = False, _
    Optional bytModiMoneyType As Byte = 0, _
    Optional ByVal blnFamilyMoney As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵�ʣ���
    '���:blnInsure=�Ƿ��ſ�ҽ�����˵�Ԥ�����
    '       curModiMoney=�޸�ʱ,ԭ���ݵĵ�ǰ���˵ķ��úϼ�
    '       int����:����(0-�����סԺ����;1-����;2-סԺ),-1��ʾ����
    '       bytModiMoneyType-�޸ķ��õ����(�ڰ����ͳ��ʱ��Ч)
    '       blnFamilyMoney-�Ƿ��ȡ�������
    '����:
    '����:����ʣ���
    '����:���˺�
    '����:2011-07-21 15:33:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, blnҽ�� As Boolean, lng��ҳId As Long
    Dim strSql As String
    On Error GoTo errH
    If blnInsure Then
        strSql = "Select A.����,A.��ҳID From ������ҳ A,������Ϣ B" & _
                " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
                " And B.����ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID)
        If Not rsTmp.EOF Then
            blnҽ�� = Not IsNull(rsTmp!����)
            lng��ҳId = rsTmp!��ҳID
        End If
    End If
    strSql = "Select " & IIf(bln������ͳ��, "����,", "") & IIf(blnFamilyMoney, "0 As ����,", "") & _
            "       Nvl(�������,0) As �������,Nvl(Ԥ�����,0) As Ԥ�����" & _
            " From �������" & _
            " Where ����=1 And ����ID=[1] " & IIf(int���� = -1, "", " And ����=[4]")
    '79868,��ȡ���˼������
    If blnFamilyMoney Then
        strSql = strSql & " Union All " & _
                " Select " & IIf(bln������ͳ��, "a.����,", "") & IIf(blnFamilyMoney, "1 As ����,", "") & _
                "       Nvl(a.�������, 0) As �������, Nvl(a.Ԥ�����, 0) As Ԥ�����" & _
                " From ������� A, ���˼��� B" & _
                " Where a.����id = b.����id And b.����id = [1] And a.���� = 1 " & _
                "       And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) " & _
                IIf(int���� = -1, "", " And ����=[4]")
    End If
  
    If dblModiMoney <> 0 Then   '����Ҫ��Union��ʽ,���ֱ��ȥ��,�ڲ�������޼�¼ʱ,���᷵�ؼ�¼
        strSql = strSql & " Union All " & _
                " Select " & IIf(bln������ͳ��, "[4] as ����,", "") & IIf(blnFamilyMoney, "0 As ����,", "") & _
                "       -1*[3] as �������,0 as Ԥ����� From Dual"
    End If
    
    '���Ϊҽ��סԺ���ˣ����ڷ���������ſ�Ԥ���еķ���(���ڱ���)
    If blnInsure And blnҽ�� Then
        strSql = strSql & " Union All " & _
        " Select  " & IIf(bln������ͳ��, "Decode(��ҳID,NULL,1,0,1,2) as ����,", "") & IIf(blnFamilyMoney, "0 As ����,", "") & _
        "       -1*Nvl(���,0) as �������,0 as Ԥ�����" & _
        " From ����ģ�����" & _
        " Where ����ID=[1] And ��ҳID=[2] "
    End If
    strSql = "Select " & IIf(bln������ͳ��, "����,", "") & IIf(blnFamilyMoney, "����,", "") & _
            "       nvl(Sum(�������),0) as �������,nvl(Sum(Ԥ�����),0) as Ԥ����� " & _
            " From (" & strSql & ")" & vbCrLf & _
            IIf(bln������ͳ�� And blnFamilyMoney, " Group by ����,����", _
                IIf(bln������ͳ��, " Group by ����", IIf(blnFamilyMoney, " Group by ����", "")))
    
    Set GetMoneyInfoRegist = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, lng��ҳId, dblModiMoney, int����)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function ReCalcOld(ByVal DateBir As Date, Optional ByRef cbo���䵥λ As ComboBox, Optional ByVal lng����ID As Long, Optional ByVal blnSetControl As Boolean = True) As String
'����:���ݳ����������¼��㲡�˵�����,�������䵥λ
'����:blnSetControl�Ƿ��������䵥λ�ؼ�
'����:����,���䵥λ
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
 
    strSql = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, DateBir)
    If blnSetControl = False Then
        ReCalcOld = Trim(Nvl(rsTmp!old))
        Exit Function
    End If
    
    If Not IsNull(rsTmp!old) Then
        If rsTmp!old Like "*��" Or rsTmp!old Like "*��" Or rsTmp!old Like "*��" Then
            strTmp = Mid(rsTmp!old, 1, Len(rsTmp!old) - 1)
            If IsNumeric(strTmp) Then
                Call gobjControl.Cbo.Locate(cbo���䵥λ, Mid(rsTmp!old, Len(rsTmp!old), 1))
            Else
                strTmp = rsTmp!old
                cbo���䵥λ.ListIndex = -1
            End If
        Else
            strTmp = rsTmp!old
            If IsNumeric(strTmp) Then
                cbo���䵥λ.ListIndex = 0
            Else
                cbo���䵥λ.ListIndex = -1
            End If
        End If
    End If
    If cbo���䵥λ.ListIndex = -1 Then
        cbo���䵥λ.Visible = False
    Else
        If cbo���䵥λ.Visible = False Then cbo���䵥λ.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function CheckChargeItemByPlugIn(objPlugIn As Object, _
    lngSys As Long, ByVal lngModule As Long, _
    ByVal intType As Integer, ByVal intMode As Integer, _
    ByRef rsDetail As ADODB.Recordset, Optional strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ҳ������շ���Ŀ��Ч�Խ��м��
    '���:lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '     intType:0-����;1-סԺ
    '     intMode:0-¼����ϸʱ�ĳ�����;1-���浥��ǰ�Ļ��ܼ��
    '     rsDetail-����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������,
    '                  ִ�п���ID���������ʣ�1-�շѵ�,2-���ʵ�)���Ƿ񻮼�(1-����;0-�������շѼ����ʵ�)
    '     strExpend-���Ժ���չ��������
    '����:strExpend-���Ժ���չ��������
    '����:���ݺϷ�����true,���򷵻�False
    '����:Ƚ����
    '����:2017-04-19 10:09:26
    '�����:105189
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '1.û����Ҳ���ʱ����Ϊ���ͨ��
    '2.��Ҳ�������CheckChargeItem�ӿڣ�Ҳ��Ϊ���ͨ��
    If objPlugIn Is Nothing Then CheckChargeItemByPlugIn = True: Exit Function
    
    On Error Resume Next
    If objPlugIn.CheckChargeItem(lngSys, lngModule, intType, intMode, rsDetail, strExpend) = False Then
        'ע�⣬�ӿڲ�����ʱҲ�����
        If Err <> 0 Then
            If Err.Number = 438 Then '�ӿڲ����ڣ���Ϊ���ͨ��
                CheckChargeItemByPlugIn = True
                Exit Function
            End If
            Call zlPlugInErrH(Err, "CheckChargeItem")
        End If
        Exit Function
    End If
    CheckChargeItemByPlugIn = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function CheckStructAddr(ByVal objCtl As PatiAddress, ByVal lngLen As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ṹ����ַ�ؼ��е���Ϣ¼���Ƿ���ȷ
    '���:objCtl-�ṹ����ַ�ؼ���lngLen-���Ƴ���
    '����:True-������Ϣ�Ϸ�
    '����:���ϴ�
    '����:2015-12-7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjCommFun.ActualLen(objCtl.Value) > lngLen Then
        MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "���ֻ������" & lngLen \ 2 & "������,���顣", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    If objCtl.CheckNullValue(, True, False) <> "" Then
        MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "��" & objCtl.CheckNullValue & "��δ����,���顣", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    CheckStructAddr = True
End Function

Public Function zlReadAddrInfo(ByVal objCtrl As PatiAddress, ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
                               ByVal intType As Integer, Optional ByVal strAddress As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���Ĳ��˵�ַ��Ϣ���ؼ���
    '���:objCtrl-�ṹ����ַ�ؼ�,intType -��ַ����1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ
    '����:
    '����:���ϴ�
    '����:2015/12/3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    
    strSql = "Select ʡ,��,��,����,���� From ���˵�ַ��Ϣ Where ����ID=[1] and Nvl(��ҳID,0)=[2] and ��ַ���=[3]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ѯ�ṹ����ַ", lng����ID, lng��ҳId, intType)
    If rsTmp.RecordCount > 0 Then
        Call objCtrl.LoadStructAdress(Nvl(rsTmp!ʡ), Nvl(rsTmp!��), Nvl(rsTmp!��), Nvl(rsTmp!����), Nvl(rsTmp!����))
    Else
        objCtrl.Value = strAddress
    End If
    zlReadAddrInfo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlPatiIsReturnVisit(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�����Ƿ��ﲡ��
    '���:lng����ID-����ID
    '    lngִ�в���ID-�Һſ���ID
    '����:
    '����:true-����,false-���ﲡ��
    '����:���˺�
    '����:2017-10-27 15:29:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    strSql = "Select Zl1_Fun_GetReturnVisit([1],[2]) As �����־ From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡȱʡ�����־", lng����ID, lngִ�в���ID)
    
    zlPatiIsReturnVisit = Val(Nvl(rsTmp!�����־)) = 1
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ShowMsgBox_Custom(ByVal frmMain As Object, ByVal strInfo As String, Optional ByVal blnNoAsk As Boolean, Optional ByVal intType As Integer) As VbMsgBoxResult
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ϣ��
    '���:frmMain-���õ�������
    '     strInfo=��ʾ��Ϣ,��Ҫ���Ѵ�����,����"^"��ʾ�س�,">"��ʾ����
    '     intType=��Ϣ������=0(ȱʡ)=MsgBox����,1-Ƥ������
    '     blnNoAsk="intType=0"ʱ��Ч����ʾ�Ƿ�ֻ��ʾһ��ȷ����ť,����ѯ�ʷ�ʽ��ʾ�Ǻͷ�
    '����:
    '    intType=0��vbIgnore=���Ҳ�����ʾ,vbCancel=���Ҳ�����ʾ,vbYes=��,vbNo=��
    '    intType=1��vbYes=����,vbNo=����,vbCancel=ȡ��
    '����:���˺�
    '����:2017-11-08 11:17:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmNewMsg As frmMsgBox
    
    Set frmNewMsg = New frmMsgBox
    ShowMsgBox_Custom = frmNewMsg.ShowMsgBox(strInfo, frmMain, blnNoAsk, intType)
    If Not frmNewMsg Is Nothing Then Unload frmNewMsg: Set frmNewMsg = Nothing
End Function

Public Function SelectWholeItems(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
     ByRef rsOutSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ŀѡ����(ѡ�������)
    '���:lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:rsOutSel-�ɹ�ʱ,����ѡ��ĳ�����Ŀ(���ֶ�:ϸĿID,����,����,���,��������,ִ�п���....)
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-08 16:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmNew As frmWholeSelect
    On Error GoTo errHandle
    Set frmNew = New frmWholeSelect
    SelectWholeItems = frmNew.ShowSelect(frmMain, lngModule, strPrivs, rsOutSel)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function zlChargeSaveValied_Plugin(ByVal lngModule As Long, ByVal int��¼���� As Integer, ByVal bln���� As Boolean, _
    ByVal bln���۵� As Boolean, ByVal strNos As String, ByVal rsSaveItems As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ң���鱣�����ݵĺϷ���
    '���:lngModule-ģ���
    '     int��¼����-1-�շѵ�;2-���ʵ�
    '     bln���۵�-�Ƿ�ǰ�Ǳ���Ļ��۵�
    '     strNOs-�����շ�ʱ������Ļ��۵��ţ��Ա����շѵĻ��۵���)
    '     rsSaveItems=��ǰ�������Ŀ�����ֶ�(�ֶ� :����ID����ҳID,�������, ���,�۸񸸺�,�շ�ϸĿID��������Ŀid������ �����Σ���׼���ۣ�Ӧ�ս�� ��
    '                                            ʵ�ս�����ʱ�䣬��Ŀ���룬��Ŀ���ƣ��������,��������ID,������,ִ�в���ID)
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2017-12-13 17:55:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1.û����Ҳ���ʱ����Ϊ���ͨ��
    '2.��Ҳ�������CheckChargeItem�ӿڣ�Ҳ��Ϊ���ͨ��
    If gobjPlugIn Is Nothing Then zlChargeSaveValied_Plugin = True: Exit Function
    
    On Error Resume Next
    If gobjPlugIn.ChargeSaveValied(glngSys, lngModule, int��¼����, bln����, bln���۵�, strNos, rsSaveItems) = False Then
        'ע�⣬�ӿڲ�����ʱҲ�����
        If Err <> 0 Then
            If Err.Number = 438 Then '�ӿڲ����ڣ���Ϊ���ͨ��
                zlChargeSaveValied_Plugin = True
                Exit Function
            End If
            Call zlPlugInErrH(Err, "ChargeSaveValied")
            Err = 0: On Error GoTo 0
        End If
        Exit Function
    End If
    zlChargeSaveValied_Plugin = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Sub zlChargeSaveAfter_Plugin(ByVal lngModule As Long, ByVal lng����ID, ByVal lng��ҳId As Long, ByVal bln���� As Boolean, _
                                    ByVal int��¼���� As Integer, ByVal strNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ң���鱣�����ݵĺϷ���
    '���:     lngSys , lngModual = ��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '   lng����ID�����ʱ�ʱ������0)
    '   lng��ҳID�����ʱ�ʱ������0)
    '   bln���� -�Ƿ��������
    '   int��¼����-1-�շ�;2-����
    '   strNOs-���ݺ�,����ö��ŷָ�
    '����:���˺�
    '����:2017-12-13 17:55:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1.û����Ҳ���ʱ����Ϊ���ͨ��
    '2.��Ҳ�������CheckChargeItem�ӿڣ�Ҳ��Ϊ���ͨ��
    If gobjPlugIn Is Nothing Then Exit Sub
    
    On Error Resume Next
    Call gobjPlugIn.ChargeSaveAfter(glngSys, lngModule, lng����ID, lng��ҳId, bln����, int��¼����, strNos)
    If Err = 0 Then Exit Sub
    
    'ע�⣬�ӿڲ�����ʱҲ�����
    If Err.Number = 438 Then Exit Sub  '�ӿڲ����ڣ���Ϊ���ͨ��
    Call zlPlugInErrH(Err, "ChargeSaveAfter")
    Err = 0: On Error GoTo 0
    
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Sub


Public Function zlGetSaveDataItems_Plugin(ByVal objBills As ExpenseBill, ByRef rsItems As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��Ҫ�����������ϸ(�˹��̣���Ҫ��Ӧ������ҽӿ�,���û����Һţ���ֱ�ӷ���True,��¼������Nothing)
    '���:objBills-���ݶ���
    '����:str����Nos-���ص�ǰ�շ����漰�Ļ��۵�
    '     rsItems-���ص�ǰ��Ҫ��������ݼ�(�ֶ� :����ID����ҳID,�������, ���,�۸񸸺�,�շ�ϸĿID��������Ŀid������ �����Σ���׼���ۣ�Ӧ�ս�� ��
    '                                            ʵ�ս�����ʱ�䣬��Ŀ���룬��Ŀ���ƣ��������,��������ID,������,ִ�в���ID)
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2017-12-14 11:41:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objBillDetail As BillDetail  '���ݵ��շ�ϸĿ����
    Dim objBillIncome As BillInCome
    Dim int�۸񸸺� As Integer
    Dim int��� As Integer
    
    On Error GoTo errHandle
    
    Set rsItems = Nothing
    
    If gobjPlugIn Is Nothing Then zlGetSaveDataItems_Plugin = True: Exit Function
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "�������", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "���", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "�۸񸸺�", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "�շ���ĿID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "����", adDouble, , adFldIsNullable
    rsItems.Fields.Append "����", adDouble, , adFldIsNullable
    rsItems.Fields.Append "��׼����", adDouble, , adFldIsNullable
    rsItems.Fields.Append "Ӧ�ս��", adDouble, , adFldIsNullable
    rsItems.Fields.Append "ʵ�ս��", adDouble, , adFldIsNullable
    rsItems.Fields.Append "����ʱ��", adVarChar, 20, adFldIsNullable
    rsItems.Fields.Append "��Ŀ����", adVarChar, 30, adFldIsNullable
    rsItems.Fields.Append "��Ŀ����", adVarChar, 200, adFldIsNullable
    rsItems.Fields.Append "�������", adVarChar, 2, adFldIsNullable
    rsItems.Fields.Append "��������ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "������", adVarChar, 20, adFldIsNullable
    rsItems.Fields.Append "ִ�в���ID", adBigInt, , adFldIsNullable
    
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
    
     '��ÿ�ŵ��ݶ���ִ�б���

    int��� = 0
    For Each objBillDetail In objBills.Details
        If objBillDetail.���� <> 0 Then
            int�۸񸸺� = 0
            For Each objBillIncome In objBillDetail.InComes
              int��� = int��� + 1 '��ǰ��¼���
               rsItems.AddNew
               rsItems!����ID = objBills.����ID
               rsItems!��ҳID = objBills.��ҳID
               rsItems!������� = 1
               rsItems!��� = int���
               rsItems!�۸񸸺� = IIf(int�۸񸸺� = 0, Null, int���)
               rsItems!�շ���ĿID = objBillDetail.�շ�ϸĿID
               rsItems!������ĿID = objBillIncome.������ĿID
               rsItems!���� = objBillDetail.����
               rsItems!���� = objBillDetail.����
               rsItems!��׼���� = objBillIncome.��׼����
               rsItems!Ӧ�ս�� = objBillIncome.Ӧ�ս��
               rsItems!ʵ�ս�� = objBillIncome.ʵ�ս��
               rsItems!����ʱ�� = Format(objBills.����ʱ��, "yyyy-mm-dd HH:MM:SS")
               rsItems!��Ŀ���� = objBillDetail.Detail.����
               rsItems!��Ŀ���� = objBillDetail.Detail.����
               rsItems!������� = objBillDetail.�շ����
               rsItems!ִ�в���ID = objBillDetail.ִ�в���ID
               rsItems!��������ID = objBills.��������ID
               rsItems!������ = objBills.������
               rsItems.Update
              If int�۸񸸺� = 0 Then int�۸񸸺� = int���
            Next     'ÿһ���շ���Ŀ
        End If
    Next
    
    zlGetSaveDataItems_Plugin = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetOneCard() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ��ͨ���ü�¼��
    '����:����һ��ͨ���ü�¼��
    '����:���˺�
    '����:2014-07-04 10:17:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errH
    
    If Not grsOneCard Is Nothing Then
        If grsOneCard.State = 1 Then
            Set GetOneCard = grsOneCard
            Exit Function
        End If
    End If
    strSql = "Select ���,����,ҽԺ����,���㷽ʽ From һ��ͨĿ¼ Where ����=1"
    Set grsOneCard = gobjDatabase.OpenSQLRecord(strSql, App.ProductName)
    Set GetOneCard = grsOneCard
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function
Public Function zlOldOneCardIsStart(ByVal str���㷽ʽ As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ��ͨ�Ƿ�����
    '���:str���㷽ʽ-���㷽ʽ
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-01 11:45:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsOne As ADODB.Recordset
    Dim blnSart As Boolean
    On Error GoTo errHandle
    Set rsOne = GetOneCard
    If rsOne Is Nothing Then Exit Function
    
    rsOne.Filter = "���㷽ʽ='" & str���㷽ʽ & "'"
    blnSart = Not rsOne.EOF
    rsOne.Filter = 0
    zlOldOneCardIsStart = blnSart
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function



Public Function zlInterfacePrayMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lng�ҺŽ���ID As Long, ByRef cllTheeSwap As Collection, _
    ByRef cllTheeSwapOther As Collection, dblMoney As Double, ByVal strCardNO As String, lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����:cllTheeSwap-�޸�������������
    '        cll��������-����������������
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If lngҽ�ƿ����ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(frmMain, lngModule, lngҽ�ƿ����ID, bln���ѿ�, strCardNO, lng�ҺŽ���ID, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '����������������
     If lng�ҺŽ���ID <> 0 Then
        '����:58322
        'mbytMode As Integer '0-�Һ�,1-ԤԼ,2-����,3-ȡ��ԤԼ ,4-�˺� ԤԼ������ģʽ:0-�Һ�,��ʱԤԼҪ�շ�,1-ԤԼ,���շ�
        If Not bln���ѿ� Then
            '���ѿ��Ѿ��ڲ���Һż�¼ʱ,�Ѿ��ۿ�
            Call zlAddUpdateSwapSQL(False, lng�ҺŽ���ID, lngҽ�ƿ����ID, bln���ѿ�, strCardNO, strSwapGlideNO, strSwapMemo, cllTheeSwap)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lng�ҺŽ���ID, lngҽ�ƿ����ID, bln���ѿ�, strCardNO, strSwapExtendInfor, cllTheeSwapOther)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlAddThreeSwapSQLToCollection(ByVal blnԤ���� As Boolean, _
    ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal str���� As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    ' ����:cllPro-����SQL��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSql As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
    'strExpend:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
    varData = Split(strExpend, "||")
    Dim str������Ϣ As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If gobjCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                    str������Ϣ = Mid(str������Ϣ, 3)
                    'Zl_�������㽻��_Insert
                    strSql = "Zl_�������㽻��_Insert("
                    '�����id_In ����Ԥ����¼.�����id%Type,
                    strSql = strSql & "" & lng�����ID & ","
                    '���ѿ�_In   Number,
                    strSql = strSql & "" & IIf(bln���ѿ�, 1, 0) & ","
                    '����_In     ����Ԥ����¼.����%Type,
                    strSql = strSql & "'" & str���� & "',"
                    '����ids_In  Varchar2,
                    strSql = strSql & "'" & strIDs & "',"
                    '������Ϣ_In Varchar2:������Ŀ|��������||...
                    strSql = strSql & "'" & str������Ϣ & "',"
                    'Ԥ����ɿ�_In Number := 0
                    strSql = strSql & IIf(blnԤ����, "1", "0") & ")"
                    zlAddArray cllPro, strSql
                    str������Ϣ = ""
                End If
                str������Ϣ = str������Ϣ & "||" & strTemp
            End If
        End If
    Next
    If str������Ϣ <> "" Then
        str������Ϣ = Mid(str������Ϣ, 3)
        'Zl_�������㽻��_Insert
        strSql = "Zl_�������㽻��_Insert("
        '�����id_In ����Ԥ����¼.�����id%Type,
        strSql = strSql & "" & lng�����ID & ","
        '���ѿ�_In   Number,
        strSql = strSql & "" & IIf(bln���ѿ�, 1, 0) & ","
        '����_In     ����Ԥ����¼.����%Type,
        strSql = strSql & "'" & str���� & "',"
        '����ids_In  Varchar2,
        strSql = strSql & "'" & strIDs & "',"
        '������Ϣ_In Varchar2:������Ŀ|��������||...
        strSql = strSql & "'" & str������Ϣ & "',"
        'Ԥ����ɿ�_In Number := 0
        strSql = strSql & IIf(blnԤ����, "1", "0") & ")"
        zlAddArray cllPro, strSql
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Sub zlCloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub
Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = Forms.Count = 0
End Function
Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ͷ���Դ
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '������Ϊ0ʱ���ŷ���Դ
    If glngInstanceCount > 0 Then Exit Function
    
    Call zlCloseSquareCardObject  '�ͷ�CardSquare����
    
    Call zlCloseWindows   '�رմ���
    
    Err = 0: On Error Resume Next
    If Not gcolPrivs Is Nothing Then Set gcolPrivs = Nothing
    If Not gclsInsure Is Nothing Then Set gclsInsure = Nothing
    If Not gobjPlugIn Is Nothing Then Set gobjPlugIn = Nothing
    If Not gobjComlib Is Nothing Then Set gobjComlib = Nothing
    If Not gobjCommFun Is Nothing Then Set gobjCommFun = Nothing
    If Not gobjControl Is Nothing Then Set gobjControl = Nothing
    If Not gobjInExse Is Nothing Then Set gobjInExse = Nothing
    If Not grsҽ�Ƹ��ʽ Is Nothing Then Set grsҽ�Ƹ��ʽ = Nothing
    If Not grsOneCard Is Nothing Then Set grsOneCard = Nothing
    If Not grs������Ŀ Is Nothing Then Set grs������Ŀ = Nothing
    zlReleaseResources = True
End Function

Public Function PatiIdentify(ByVal lngModlue As Long, ByVal frmMain As Object, ByVal lng����ID As Long, ByVal curMoney As Currency, _
    Optional ByVal bln�˷� As Boolean = False, Optional ByVal bytDepositShowMode As Byte = 0, Optional ByVal lngDefaultCardTypeID As Long = 0, _
    Optional ByVal blnFamilyMoney As Boolean, Optional ByVal blnOlnyFamilyIDs As Boolean, Optional strFamilyPatiIDs_Out As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ����֤
    '���:lngModlue-ģ���
    '     dblMoney-���
    '     lng����ID-����ID
    '     bln�˷�-��ǰ�Ƿ��˷Ѳ���
    '     bytDepositShowMode- Ԥ����ʾ��ʽ(0-��������ʾ;1-ֻ��ʾ�������;2-ֻ��ʾסԺ���)
    '     lngDefaultCardTypeID-ȱʡ��ˢ�����
    '     blnFamilyMoney-�Ƿ��ȡ����Ԥ�����
    '     blnOlnyFamilyIDs-true:���鿨��ֻ��ȡ����IDs;False-��Ҫ��ȡ���鿨
    '����:strFamilyPatiIDs-���˼���ID,����ö��ŷָ���79868
    '����:�����֤�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-10-24 14:55:59
    '˵��:
    '   һ�������鿨�����������bln�˷�=falseʱ):
    '       1.������ˢ����֤,ֱ�ӷ���True
    '       2.��������ʱ����Ҫ����ˢ����֤��ͬʱ��Ҫ�������루������ʱ,���Ҫ���������)
    '       3.��������ʱ��������ģ������ˢ���鿨���������룬������ʱ,����Ҫ�鿨��������
    '       4.��ʾ����������NԪ�ڱ���ˢ��,�����������뼴��֧��;���������������(������ʱ�����Ҫ���������)
    '  �����˷��鿨��bln�˷�=trueʱ):
    '       1.������ˢ�����ƣ�ֱ�ӷ���true
    '       2.���������˷�ʱ��Ҫˢ����֤,ͬʱ��Ҫ�������루������ʱ,���Ҫ���������)
    '       3.���������˷�ʱ��������ģ������ˢ����֤,������ʱ,����Ҫ�鿨��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue  As String, dblBrushCardMoney As Double
    Dim byt�����鿨 As Byte, byt�˷��鿨 As Byte, blnPassWord As Boolean
    Dim varPara As Variant
    
    On Error GoTo errHandle
    'һ��ͨ������֤
    strValue = gobjDatabase.GetPara(28, glngSys, , "1|0")
    varPara = Split(strValue & "|||", "|")
    byt�˷��鿨 = Val(varPara(1)) '���ѿ��˷�ʱ�Ƿ�ˢ����֤
    
    dblBrushCardMoney = Val(varPara(0))
    
    If dblBrushCardMoney < 0 Then
        byt�����鿨 = 3 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤  3-��ʾ����������NԪ�ڱ���ˢ��,�����������뼴��֧��;���������������
        dblBrushCardMoney = -1 * dblBrushCardMoney  'ˢ��ʱ����֧�����("gbytԤ��������鿨"Ϊ3ʱ��Ч)
    Else
        byt�����鿨 = Decode(dblBrushCardMoney, 1, 1, 2, 2, 0)
    End If
    
    If bln�˷� Then
        '  byt�����鿨 'Ԥ����˷�ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
        If byt�˷��鿨 = 0 Then PatiIdentify = True: Exit Function '������ˢ����֤,ֱ�ӷ���True
        
        If gobjDatabase.PatiIdentify(frmMain, glngSys, lng����ID, curMoney, lngModlue, bytDepositShowMode, lngDefaultCardTypeID, , _
                                 blnFamilyMoney, strFamilyPatiIDs_Out, Not blnOlnyFamilyIDs, (byt�˷��鿨 = 2)) Then Exit Function
        
        PatiIdentify = True: Exit Function
    End If
    If byt�����鿨 = 0 Then PatiIdentify = True: Exit Function '������ˢ����֤,ֱ�ӷ���True
    
    
    If byt�����鿨 <> 3 Then
        blnPassWord = True
    ElseIf dblBrushCardMoney = 0 Then
        blnPassWord = True
    ElseIf curMoney > dblBrushCardMoney Then
        blnPassWord = True
    ElseIf curMoney = 0 Then '�޽��ʱ��������֤����
        blnPassWord = False
    Else
        blnPassWord = False
    End If
    
    If gobjDatabase.PatiIdentify(frmMain, glngSys, lng����ID, curMoney, lngModlue, bytDepositShowMode, lngDefaultCardTypeID, blnPassWord, _
                                   blnFamilyMoney, strFamilyPatiIDs_Out, Not blnOlnyFamilyIDs, (byt�����鿨 = 2)) = False Then Exit Function
    PatiIdentify = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ReserveRegNo(ByVal str���� As String, ByVal bln�ϸ���� As Boolean, ByVal bln��ʱ�� As Boolean, _
                            ByVal strTime As String, ByRef lng��� As Long, _
                            Optional ByVal str��ע As String, Optional ByVal lng��¼ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �Ե�ǰ�������
    ' ��� :str����-�ű�
    '       blnԤԼ-�Ƿ�ԤԼ����
    '       bln�ϸ����-�Ƿ��ϸ����
    '       bln��ʱ��-�Ƿ��ʱ��
    '       lng�����Ҫ���ŵ����
    '       lng��¼ID - �����Ű�ģʽ��Ҫ�����¼id
    '       str��ע - ����������������
    ' ���� :lng���:���lng��ű���������ȡ�µ�������ţ��������µ����
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/2/18 15:34
    '---------------------------------------------------------------------------------------
    Dim lngRegLimit As Long, lngLastNo As Long, lngCurrentNo As Long, intTimes As Integer
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    If Not bln�ϸ���� Then ReserveRegNo = True: Exit Function
    If Not strTime Like "To_Date*" And strTime <> "" Then strTime = "To_Date('" & Format(strTime, "yyyy-MM-dd hh:mm:00") & "','YYYY-MM-DD HH24:MI:SS')"
    '138960:���ϴ�,2019/3/26,����ʱ�ε�ֻȡ����
    If Not bln��ʱ�� Then
        strTime = Mid(strTime, InStr(strTime, "'") + 1)
        strTime = Trim(Left(strTime, InStr(strTime, "'") - 1))
        strTime = "To_Date('" & Format(strTime, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    End If
    On Error GoTo errH:
    If bln�ϸ���� And Not bln��ʱ�� And lng��� = 0 Then
Retry:
        If lng��¼ID <> 0 Then
            strSql = "Select A.�޺���,B.���,Nvl(B.�Һ�״̬,0) as ״̬,Nvl(B.�Ƿ�ͣ��,0) as ͣ��  From �ٴ������¼ A,�ٴ�������ſ��� B Where A.ID=B.��¼ID And A.ID= [1] order by B.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�Һ���Ų�ѯ", lng��¼ID)
            If rsTmp.RecordCount = 0 Then ReserveRegNo = True: Exit Function
            lngRegLimit = Val(Nvl(rsTmp!�޺���))
            Do While Not rsTmp.EOF
                If Val(Nvl(rsTmp!״̬)) = 0 And Val(rsTmp!ͣ��) <> 1 Then
                    lngCurrentNo = Val(Nvl(rsTmp!���))
                    Exit Do
                End If
                lngLastNo = Val(Nvl(rsTmp!���))
                rsTmp.MoveNext
            Loop
            If lngCurrentNo = 0 Then lngCurrentNo = lngLastNo + 1
            If lngCurrentNo > lngRegLimit Then
                ReserveRegNo = True
                Exit Function
            End If
            lng��� = lngCurrentNo
        Else
            strSql = "Select A.�޺���" & vbNewLine & _
                     "From �ҺŰ������� A, �ҺŰ��� B" & vbNewLine & _
                     "Where A.������Ŀ =" & vbNewLine & _
                     "      Decode(To_Char(Sysdate, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) And" & vbNewLine & _
                     "      A.����id = B.ID And B.���� = [1]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "�Һ��޺���Լ", str����)
            If Not rsTmp.EOF Then
                lngRegLimit = Val(Nvl(rsTmp!�޺���))
            End If
            strSql = "Select ���,״̬" & vbNewLine & _
                     "From �Һ����״̬" & vbNewLine & _
                     "Where ���� = [1] And ���� Between Trunc(Sysdate) And Trunc(Sysdate + 1) - 1 / 24 / 60 / 60" & vbNewLine & _
                     "Order By ��� Asc"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "�Һ���Ų�ѯ", str����)
            Do While Not rsTmp.EOF
                If lngLastNo = 0 Then
                    lngLastNo = Val(Nvl(rsTmp!���))
                Else
                    If Val(Nvl(rsTmp!���)) - lngLastNo > 1 Then
                        lngCurrentNo = lngLastNo + 1
                    Else
                        lngLastNo = Val(Nvl(rsTmp!���))
                    End If
                End If
                If Val(Nvl(rsTmp!״̬)) = 4 Then lngRegLimit = lngRegLimit + 1
                rsTmp.MoveNext
            Loop
            If lngCurrentNo = 0 Then lngCurrentNo = lngLastNo + 1
            If lngCurrentNo > lngRegLimit Then '˵���ǼӺţ����������������
                ReserveRegNo = True
                Exit Function
            End If
            lng��� = lngCurrentNo
        End If
    End If
    On Error GoTo errTry
    If lng��� <> 0 Then
        If strTime <> "" Then
            strSql = "Zl_�Һ����״̬_Lock(1,'" & UserInfo.���� & "','" & str���� & _
                      "'," & strTime & "," & lng��� & "," & ZVal(lng��¼ID) & ",'" & str��ע & "')"
        Else
            strSql = "Zl_�Һ����״̬_Lock(1,'" & UserInfo.���� & "','" & str���� & _
                      "',To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd") & "','YYYY-MM-DD')," & lng��� & _
                      "," & ZVal(lng��¼ID) & ",'" & str��ע & "')"
        End If
        Call gobjDatabase.ExecuteProcedure(strSql, "ReserveRegNo")
    End If
    ReserveRegNo = True
    Exit Function
errTry:
    intTimes = intTimes + 1
    If bln�ϸ���� And Not bln��ʱ�� And intTimes < 4 Then GoTo Retry
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Sub CancelRegNo(Optional ByVal lng��¼ID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:ȡ���Һ�ʱɾ�������Һ����
    '����:���ϴ�
    '����:2019/2/18 15:34
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo Errhand
    
    strSql = "Zl_�Һ����״̬_Lock(2,'" & UserInfo.���� & "',Null,Null,Null," & ZVal(lng��¼ID) & ")"
    Call gobjDatabase.ExecuteProcedure(strSql, "CancelRegNo")
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Sub

Public Sub InitAddressLength()
    Dim strSql As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSql = "Select ��ͥ��ַ, ���ڵ�ַ, �����ص�, ��ϵ�˵�ַ From ������Ϣ Where Rownum < 2"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��ַ����")
    If Not rsTmp.EOF Then
        glngMax��ͥ��ַ = rsTmp.Fields("��ͥ��ַ").DefinedSize
        glngMax���ڵ�ַ = rsTmp.Fields("���ڵ�ַ").DefinedSize
        glngMax�����ص� = rsTmp.Fields("�����ص�").DefinedSize
        glngMax��ϵ�˵�ַ = rsTmp.Fields("��ϵ�˵�ַ").DefinedSize
    End If
    If glngMax��ͥ��ַ = 0 Then glngMax��ͥ��ַ = 100: If glngMax���ڵ�ַ = 0 Then glngMax���ڵ�ַ = 100
    If glngMax�����ص� = 0 Then glngMax�����ص� = 100: If glngMax��ϵ�˵�ַ = 0 Then glngMax��ϵ�˵�ַ = 100
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Sub

'========================================================================================================
'zlPlugIn��ҽӿ�
'========================================================================================================
Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ҳ���������
    '���:objErr ������� strFunName �ӿڷ�������
    '����:
    '����:���˺�
    '����:2014-04-09 13:27:19
    '˵��:�����������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & _
            objErr.Number & vbCrLf & _
            objErr.Description, vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

Public Function CreatePlugIn(ByVal lngModule As Long, _
    Optional ByVal int���� As Integer) As Boolean
    '���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugIn = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    If gobjPlugIn Is Nothing Then
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    End If
    If gobjPlugIn Is Nothing Then Exit Function
    
    Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngModule, int����)
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "Initialize")
        Set gobjPlugIn = Nothing
        Exit Function
    End If
    
    CreatePlugIn = True
End Function

Public Function zlSaveRgstAfterByPlugIn(ByVal lngModule As Long, ByVal strNO As String, ByVal blnԤԼ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �Һ���ɺ����
    ' ��� : lngModual=��ǰ���ýӿڵ�������ģ���
    '        strNo-�Һŵ���
    '        blnԤԼ-ԤԼ�Һ�
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/22 10:03
    '---------------------------------------------------------------------------------------
    If CreatePlugIn(lngModule, -1) = False Then Exit Function
    
    On Error Resume Next
    If gobjPlugIn.SaveRegisterAfter(glngSys, lngModule, strNO, blnԤԼ) = False Then Exit Function
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "SaveRegisterAfter")
        Set gobjPlugIn = Nothing
        Exit Function
    End If
    
    zlSaveRgstAfterByPlugIn = True
End Function

