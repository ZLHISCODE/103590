Attribute VB_Name = "mdlCISKernel"
Option Explicit
Public gfrmMain As Object                   '����̨����
Public gclsInsure As New clsInsure          'ҽ������
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gobjCISBase As Object                '���ƻ�������
Public gobjInExse As Object                 'סԺ���ò���
Public gobjPath As Object                   '�ٴ�·������(ͨ��ҽ��վ��ʿվ����)
Public gobjPathOut As Object                '�����ٴ�·������(ͨ��ҽ��վ��ʿվ����)
Public gobjLIS As Object                    'LIS���벿��
Public gobjExchange As Object               'HL7���ݽ�������
Public gobjSquareCard As Object             'һ��ͨ���ײ���(���������ʼ���ӿڣ�������ҽ��վ��ҽ��վ���룬��������������)
Public gobjEmrInterface As Object           '�°没�����븽���ȡ����
Public gobjRecipeAudit As Object            '�������ϵͳ����
Public gobjPublicExpense As Object          '���ù�������
Public gobjPublicDrug As Object             'ҩƷ��������
Public gobjPublicBlood As Object            'Ѫ�⹫������
Public gobjPublicPatient As Object          '������Ϣ��������

Public gstrTsPrivsMZ As String              '����ҽ������ҽ��Ȩ���ַ���
Public gstrTsPrivsZY As String              'סԺҽ������ҽ��Ȩ���ַ���

Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gMainPrivs As String                 '���������������е�Ȩ��,ע����ڲ�ģ��Ȩ��
Public gstrSysName As String                'ϵͳ����
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String            'OEM��Ʒ����
Public glngSys As Long

Public gstrSQL As String
Public gbln�Ӱ�Ӽ� As Boolean              '״̬�ж���ʱ����
Public grsDuty As ADODB.Recordset           '���ҽ��ְ��
Public gstr��̬�ѱ� As String               '������ﵱǰ���ҿ��ö�̬�ѱ�,�ڹ���������ʹ��,ʹ��ʱ�Ÿ�ֵ:CalcDrugPrice,CalcPrice
Public gblnKSSStrict As Boolean             '�Ƿ����ÿ���ҩ���ϸ����
Public gbln����ҩ��ʹ���Ա�ҩ As Boolean

Public grsSkinTest As ADODB.Recordset       '���Ƥ����Ŀ�������Զ���
Public grsTube As ADODB.Recordset           '����Թ������Ϣ
Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gbytCode As Byte '�������뷽ʽ
Public grsҽ�Ƹ��ʽ As ADODB.Recordset
Public gbyt������˷�ʽ As Byte '49501:������˷�ʽ:0-δ��˲�������ʣ�ȱʡΪ0;1-���ʱ����������ú�ҽ��������ҽ�������ͷ��õ�����
Public gblnδ��ƽ�ֹ����  As Boolean

Public gbln�����ּ����� As Boolean  '�Ƿ����������ּ�����
Public gbln������Ȩ���� As Boolean  '�Ƿ�����������ҽʦ��Ȩ����
Public gbln�����ȼ����� As Boolean  '�Ƿ����ò���������ҽʦ�ﵽ�����ȼ��������
Public gbln��Ѫ�ּ����� As Boolean  '�Ƿ�������Ѫ�ּ�����
Public gbln��Ѫ�����м����� As Boolean  '��Ѫ����ֻ�����м�������ҽʦ���
Public gbln��Ѫ����������� As Boolean  '��Ѫ�����������
Public gintҽ��ִ����Ч���� As Integer '�����޸�n���ڵǼǵ�ҽ��ִ�м�¼
Public gbytת��ʱδ������ʵ��ݼ�� As Byte  'ת��ʱ�Ƿ��鲡�˴���δ��˵����ʵ���:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gblnѪ��ϵͳ As Boolean  '�Ƿ�װѪ��ϵͳ
Public gbln��ʾѪҺ��� As Boolean '��Ѫ�����Ƿ���ʾѪҺ���
Public gbln�´���Ѫ����ȷ��ѪҺ��Ϣ As Boolean '���뵥�ķ�ʽ�´���Ѫ���룬�Ƿ�Ҫ��ѡ�����ѪҺ��������
Public gbyt����ԭ�� As Byte 'ҽ������ʱ��������ԭ�� 0-������дԭ��1��������д
Public gstr��¼�������� As String
Public gbln����ҩƷ�ֿ����� As Boolean '����ҩƷ�ֿ�����
Public gblnҽ����ֹԭ�� As Boolean 'ϵͳ������ͣ��ʱ¼��ԭ��
Public gstr�ɲ���ͣ��ԭ����� As String 'ϵͳ�������ɲ���ͣ��ԭ�����
Public gobjDrugExplain As Object 'ҩƷ˵���鲿��
'--------------------
'���뵥���û��� �������������뵥�����ʹ�����뵥�´�ҽ�� ����
Public gstrInUseApp As String 'סԺ
Public gstrOutUseApp As String '����
Public gblnIn���� As Boolean 'סԺ
Public gblnOut���� As Boolean '����
'-------------------
Public gbln�������� As Boolean 'ҩƷ¼��ʱ���������뵥��������ҩ����
Public gbln����Ӱ����Ϣϵͳ�ӿ� As Boolean
Public gbln����Ӱ����ϢϵͳԤԼ As Boolean
Public gbln����ҩ�����հ������������� As Boolean
Public gbln��������´�ҽ������ As Boolean 'ϵͳ��������������´�ҽ���ɻ���������Ҵ���
Public gbln��ϵͳ As Boolean '����������ҩ�󷽹������ÿ��أ�ͨ����ȡ�����������ж�  ������������Ŀ¼

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    pסԺ���ʲ��� = 1150
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    p�ٴ�·��Ӧ�� = 1256
    P����·��Ӧ�� = 1248
    p�����¼���� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    p��Ѫ��˹��� = 1268
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    pXWPACS��Ƭ = 1288
    p��Ƭ���߹��� = 1289
    p��Һ�������� = 1345
    p�°����ﲡ�� = 2251
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
    �����ȼ� As String
End Type
Public UserInfo As TYPE_USER_INFO

'����ǩ��
Public gintCA As Integer '����ǩ����֤����
Public gstrESign As String '����ǩ�����Ƴ���
Public gobjESign As Object '����ǩ���ӿڲ���
Public grsSign As Recordset  '����ǩ�����ò��ţ����棩

'��ҹ���
Public gobjPlugIn As Object

'RIS�ӿڲ���
Public gobjRis As Object

'CISϵͳ����
Public gblnҩƷ�������ҽ�� As Boolean
Public gint�����Ǽ���Ч���� As Integer
Public gbln����ҽ��������Ч As Boolean
Public gstrסԺ���ͻ��۵� As String
Public gstr���﷢�ͻ��۵� As String
Public gstr��Һ�������� As String
Public glng��¼ʱ�� As Long
Public gblnֻ����¼���� As Boolean
Public gblnShowOrigin As Boolean '�Ƿ���ʾ����

Public gblnִ��ǰ�Ƚ��� As Boolean 'һ��ִͨ��ǰ���շѻ�������

Public gintRXCount As Integer
Public gint��¼��� As Integer '�Զ�ʶ��Ϊ��¼ҽ����ʱ����(����)
Public gbln�������������� As Boolean '�Ƿ��ڼ���ҽ������ʱ����������
Public gdblԤ��������鿨 As Double 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
                                                      'Ϊ����(-N)ʱ��ʾ,NԪ������֧��,��ʾ����������NԪ�ڱ���ˢ��,�����������뼴��֧��;���������������
Public gblnָ��ҽ������������ִ�� As Boolean 'ָ��ҽ������������ִ��
Public gstrҽ���˶� As String    '��ѪƤ��ҽ����Ҫ�˶� ��λ��ȡ11����һλΪ ��Ѫҽ�����ڶ�λΪ Ƥ��ҽ��
Public gint�����¿�ҽ�����  As Integer '�����¿�ҽ�����
Public gbln�����������շѻ������� As Boolean '��Ŀ�����������շѻ�������

'HISϵͳ����
Public gbytMediOutMode As Byte '����ҩƷ���ⷽʽ��0-�������Ƚ��ȳ���1-��Ч������ȳ�,Ч����ͬ�����ٰ������Ƚ��ȳ�
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '�����С��λ������ĸ�ʽ����,��"0.0000"
Public gbytDecPrice As Byte '���õ��۵�С����λ��
Public gstrDecPrice As String '�۸�С��λ������ĸ�ʽ����,��"0.0000"

Public gint��ͨ�Һ����� As Integer '��ͨ�Һŵ���Ч����
Public gint����Һ����� As Integer '����Һŵ���Ч����
Public gint���Ʊ��� As Integer '0-˳����,1-����+�����+˳����
Public gbytҩƷ������ʾ As Byte '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
Public gbyt����ҩƷ��ʾ As Byte '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
Public gbln����ƥ�䷽ʽ�л� As Boolean '�����ڴ��ڽ���Ĺ������л�����ƥ�䷽ʽ�л�

Public gbln��������ۿ� As Boolean '������Ŀ���ܼ����ۿ�
Public gint�����Դ As Integer '1-��ҽ��ѡ��������Դ,2-������ϱ�׼����,3-���ռ�����������
Public gstr������� As String '1����/2סԺ��1-������������,2-�����ݿ���ȡ����,3-��ҽ�����˴����ݿ�����
Public gstrMatchMode As String '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
Public gbyt��Ժ���δִ�� As Byte '��Ժʱ�Ƿ�����δִ����Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gbytת�Ƽ��δִ�� As Byte 'ת��ʱ�Ƿ�����δִ����Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gbyt��Ժ���δ��ҩ As Byte '��Ժʱ�Ƿ�����δ��ҩƷ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gbytת�Ƽ��δ��ҩ As Byte 'ת��ʱ�Ƿ�����δ��ҩƷ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ

Public gintҽ������ As Integer '�Ƿ��סԺҽ�����˵���Ŀ����������м��:0-�����,1-��鲢����,2-��鲢��ֹ
Public gcurMaxMoney As Currency '���ʷ���������ѽ��

'ҽ������վϵͳ���ò���
Public gblnStock As Boolean 'ָ��ҩ��ʱ�Ƿ��޶�����ҩƷ�Ŀ��
Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gbln�շ���� As Boolean '�Ƿ������������
Public gblnFeeKindCode As Boolean '�������ʱ,��λ�����շ�������

Public gbytסԺ�Զ����� As Byte  'סԺ������ɺ��Ƿ��Զ����� 0-���Զ����ϣ�1-�Զ����ϣ�2-�����ҿ���ʱ�Զ�����
Public gbln�����Զ����� As Boolean '���������ɺ��Ƿ��Զ�����
Public gbln�շѺ��Զ���ҩ As Boolean '

Public gbln�����������۷��� As Boolean '���ʱ����������۷���
Public gstrҽ���������� As String 'ҽ����������ķ�������
Public gstr���ѷ������� As String '���Ѳ�������ķ�������
Private mlng���ű���ƽ������ As Long


Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29 '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    support�������� = 35            '�Ƿ����������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
    
    supportҽ��ȷ���������� = 48
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    supportʵʱ��� = 60
    
    support�ϴ����ﵵ�� = 70                    '������ҽ������ʱ���Ƿ����TranElecDossier����������ﲡ�˵��Ӿ���/���ӵ������ϴ�
    support�ϴ�סԺ���� = 70                    '��סԺҽ������ʱ���Ƿ����TranElecDossier ����˵��������סԺʹ��ͬһ��ҵ�����
End Enum

'Pass
Public gobjPass As Object  'PASS �ӿ�

'�ӿ�����ö��
Public Enum G_PASS_TYPE
    UNPASS = 0          'δ����
    MK = 1              '����
    DT = 2              '��ͨ
    TYT = 3             '̫Ԫͨ
    YWS = 4             '����ҩ��ʿ
    HZYY = 5            '��������
End Enum
'����ģ����
Public Enum PASS_MODEL
    PM_����༭ = 0
    PM_סԺ�༭ = 1
    PM_סԺҽ���嵥 = 2
    PM_��ʿУ�� = 3
    PM_����ҽ���嵥 = 4
End Enum


Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.����ID = NVL(rsTmp!����ID, 0)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.���� = Get��Ա����
            UserInfo.רҵ����ְ�� = NVL(rsTmp!רҵ����ְ��)
            UserInfo.�����ȼ� = NVL(rsTmp!�����ȼ�)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.�û���
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
'���ܣ���ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    If str���� <> "" Then
        strSQL = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str����)
    Else
        strSQL = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAuditName(ByVal strName As String) As String
'���ܣ���"���ҽ��/ʵϰҽ��"��ȡ���ҽ����
    GetAuditName = Mid(strName, 1, IIF(InStr(strName, "/") > 0, InStr(strName, "/") - 1, Len(strName)))
End Function

Public Function HaveAuditPriv(Optional ByVal str���� As String) As Boolean
'���ܣ��жϵ�ǰ��Ա�Ƿ����"ִҵҽʦ"���ʸ�
'������str����=�ж�ָ����Ա��������ʱ�жϵ�ǰ��Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strHave As String, strNone As String
    Dim strSQL As String
    
    If str���� = "" Then str���� = UserInfo.����
    
    If InStr(strHave & "|", "|" & str���� & "|") > 0 Then
        HaveAuditPriv = True: Exit Function
    ElseIf InStr(strNone & "|", "|" & str���� & "|") > 0 Then
        HaveAuditPriv = False: Exit Function
    End If
    
    On Error GoTo errH
    strSQL = "Select B.���� From ��Ա�� A,ִҵ��� B Where A.����=[1] And A.ִҵ���=B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "HaveAuditPriv", str����)
    If Not rsTmp.EOF Then
        If rsTmp!���� = "ִҵҽʦ" Or rsTmp!���� = "ִҵ����ҽʦ" Then HaveAuditPriv = True
    End If
    If HaveAuditPriv Then
        strHave = strHave & "|" & str����
    Else
        strNone = strNone & "|" & str����
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��������(ByVal lng����ID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select �������� From ��������˵�� Where ����ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get��������", lng����ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!��������
        rsTmp.MoveNext
    Loop
    Get�������� = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean, Optional ByVal lngSys As Long) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    If lngSys = 0 Then lngSys = glngSys
    On Error Resume Next
    strPrivs = gcolPrivs(lngSys & "_" & lngProg)
    If err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove lngSys & "_" & lngProg
        End If
    Else
        err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(lngSys, lngProg)
        gcolPrivs.Add strPrivs, lngSys & "_" & lngProg
    End If
    GetInsidePrivs = IIF(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function GetPatiUnitID(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Long
'���ܣ����ݲ��˻�ȡ��Ӧ�Ĳ���ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��ǰ����ID as ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    GetPatiUnitID = NVL(rsTmp!����ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiRsByUnit(ByVal lngUnitID As Long, ByVal lng����ID As Long, _
    ByVal bln���ò��� As Boolean, ByVal blnʣ��� As Boolean, ByVal bln������Ժ���� As Boolean, _
    Optional ByVal blnIsPreOut As Boolean, Optional ByVal lngҽ������Χ As Long) As ADODB.Recordset
'���ܣ���ȡ��ǰ��������Ժ�����б��Լ���ǰ�����Զ������б�ת����Ԥ��Ժ����Ժ��
'������bln������Ժ����=��ʿվ�����ջص���ʱ����ʾ�����Ժ�Ĳ���
'     lngҽ������Χ=-1���в��˰���Ӥ����0���ˣ�1Ӥ��
    Dim strSQL As String, intBedLen As Integer, strIF As String
    Dim curDate As Date, intDay As Integer, dtOutEnd As Date, dtOutBegin As Date
    Dim strPreOut As String
    Dim strʣ��� As String
    Dim str���ò��� As String
    Dim str��Χ���� As String '-1
    Dim str��ΧӤ�� As String '1
    Dim strOther As String '1 or -1
    
    On Error GoTo errH
    intBedLen = GetMaxBedLen(lngUnitID, False)
    strPreOut = " And Nvl(b.״̬,0)<>3 "
    'Union��Ϊ����������(�Զ�ȥ���ظ���)������In����0+[2]����Ϊ�˱������:���ʽ����������Ӧ���ʽ��ͬ����������
    strIF = "Select a.����ID" & vbNewLine & _
            " From ������Ϣ A, ������ҳ B,��Ժ���� R" & vbNewLine & _
            " Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����ID=R.����ID And A.��ǰ����ID=R.����ID And (R.����id = [1] or B.Ӥ������ID = [1]) " & IIF(blnIsPreOut, "", strPreOut) & vbNewLine & _
            " Union" & vbNewLine & _
            " Select 0+[2] as ����ID From Dual"

    
    If bln������Ժ���� Then
         '��Ժ����ʱ�䷶Χ
        curDate = zlDatabase.Currentdate
        intDay = Val(zlDatabase.GetPara("��Ժ���˽������", glngSys, pסԺ��ʿվ, 0))
        dtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
        intDay = Val(zlDatabase.GetPara("��Ժ���˿�ʼ���", glngSys, pסԺ��ʿվ, 1))
        dtOutBegin = Format(curDate - intDay, "yyyy-MM-dd 00:00:00")
    
        strIF = strIF & " Union" & vbNewLine & _
                " Select a.����id" & vbNewLine & _
                " From ������Ϣ A, ������ҳ B" & vbNewLine & _
                " Where a.����id = b.����id And a.��ҳid = b.��ҳid And (b.��ǰ����id + 0 = [1] Or b.Ӥ������id + 0 = [1]) And B.��Ժ���� Between [3] And [4]"
    End If
    
    If blnʣ��� Then
        strʣ��� = "Nvl(E.Ԥ�����,0)-Nvl(E.�������,0)+Decode(B.����,Null,0,(Select Nvl(Sum(���),0) From ����ģ����� F" & _
            " Where B.����ID=F.����ID And B.��ҳID=F.��ҳID))"
    Else
        strʣ��� = "Null"
    End If
    
    str���ò��� = IIF(bln���ò���, "zl_PatiWarnScheme(A.����ID,B.��ҳID)", "Null")
    
    If lngҽ������Χ = 1 Then str��ΧӤ�� = " And exists(select 1 from ������������¼ Z Where z.����id=b.����ID And z.��ҳid=b.��ҳID)"
    
    If lngҽ������Χ = -1 Then
        str��Χ���� = " Select ����id, ��ҳid, ����, סԺ��, ����, ������, ʣ���,  ���ò���, ����, סԺҽʦ, �ѱ�, ����ȼ�, ����,����id," & _
            " ��Ժ����, ��Ժ����, ��������, �Ա�, ��˱�־,Ӥ������ID,Ӥ������ID,Null as Ӥ������,Null as Ӥ�����,����״̬,���ۺ� From Pati Union All "
    End If
    
    If lngҽ������Χ = 1 Or lngҽ������Χ = -1 Then
        strOther = "select a.����id,a.��ҳid,a.����,a.סԺ��,a.����,a.������,a.ʣ���, a.���ò���,a.����,a.סԺҽʦ,a.�ѱ�,a.����ȼ�,a.����,a.����id," & _
            " a.��Ժ����, a.��Ժ����, a.��������, a.�Ա�, a.��˱�־,a.Ӥ������ID,a.Ӥ������ID,b.Ӥ������,B.��� AS Ӥ�����,����״̬,a.���ۺ� from Pati A,������������¼ B" & _
            " where A.����id=b.����id and A.��ҳID=b.��ҳid "
    Else
        strOther = "Select ����id, ��ҳid, ����, סԺ��, ����, ������, ʣ���,  ���ò���, ����, סԺҽʦ, �ѱ�, ����ȼ�, ����, ����id," & _
            " ��Ժ����, ��Ժ����, ��������, �Ա�, ��˱�־,Ӥ������ID,Ӥ������ID,Null as Ӥ������,Null as Ӥ�����,����״̬,���ۺ� From Pati "
    End If
    
    strSQL = "Select * from (" & _
        " With Pati as (Select A.����ID,B.��ҳID,A.����,B.סԺ��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,A.������," & _
        strʣ��� & " as ʣ���," & str���ò��� & " as ���ò���,B.����,B.סԺҽʦ,B.�ѱ�,D.���� as ����ȼ�,C.���� as ����," & _
        " c.id as ����id,B.��Ժ����,B.��Ժ����,B.��������,A.�Ա�,b.��˱�־,B.Ӥ������ID,B.Ӥ������ID,Decode(B.״̬,3,1,Decode(B.��Ժ����,Null,0,2)) as ����״̬,B.���ۺ�" & _
        " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ D,������� E,(" & strIF & ") F" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.����ȼ�ID=D.ID(+)" & _
        " And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
        " And A.����ID=E.����ID(+) And E.����(+)=1 And E.����(+) = 2 And A.����ID=F.����ID" & _
        str��ΧӤ�� & " Order by ����) " & _
        str��Χ���� & _
        strOther & _
        ") order by ����,NVL(Ӥ�����,0)"
        
    Set GetPatiRsByUnit = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngUnitID, lng����ID, dtOutBegin, dtOutEnd)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPatiTurnLimit(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lng����id As Long, _
    ByRef datTurn As Date, ByVal intPState As TYPE_PATI_State) As Boolean
'���ܣ�����Ժ��Ԥ��Ժ��ת�ơ�ת�����Ĳ��˵Ĳ�¼ʱ�ޣ���������򷵻ظò�����ʱ��
    Dim strSQL As String, strMsg As String
    Dim rsTmp As ADODB.Recordset
    
    If intPState = ps��Ժ Then
        strSQL = "Select ��Ժ���� as ��ֹʱ�� From ������ҳ Where ����id = [1] and ��ҳID=[2]"
        strMsg = "��Ժ"
    Else
        If intPState = ps���ת�� Then
            strSQL = " And (��ֹԭ�� =3 and ����ID=[4] or ��ֹԭ�� =15 and ����ID=[3])"
            strMsg = "ת�ƻ�ת����"
        ElseIf intPState = psԤ�� Then
            strSQL = " And ��ֹԭ�� = 10"
            strMsg = "����Ԥ��Ժ"
        End If
        strSQL = "Select Max(��ֹʱ��) as ��ֹʱ�� From ���˱䶯��¼" & _
                " Where ����id = [1] and ��ҳID=[2]   And ��ֹʱ�� is Not Null" & strSQL
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��¼��ʱ�޼��", lng����ID, lng��ҳID, lng����ID, lng����id)
    If IsNull(rsTmp!��ֹʱ��) Then
        MsgBox "�������쳣��û���ҵ��ò��˵ı䶯��Ϣ���޷�����ҽ����¼������", vbInformation, gstrSysName
        Exit Function
    End If
    
    datTurn = CDate(Format(rsTmp!��ֹʱ��, "yyyy-mm-dd HH:MM:SS"))
    If glng��¼ʱ�� = 0 Then
        ShowMsgBox "ע��:" & vbCrLf & "    �ò�����" & strMsg & ",ϵͳ����Ϊ���������ҽ����¼������"
        Exit Function
    Else
        If datTurn + 1 / 24 * glng��¼ʱ�� < zlDatabase.Currentdate Then
            ShowMsgBox "ע��:" & vbCrLf & "    �ò���" & strMsg & "�Ѿ�������" & glng��¼ʱ�� & "Сʱ,���������ҽ����¼������"
            Exit Function
        Else
            '�������Ƿ��ǲ���Ա�Ŀ��ÿ���
            If Get��������ID(UserInfo.ID, 0, lng����id) <> lng����id Then
                ShowMsgBox "ע��:�����ǵ�ǰ���ҵ�ҽ��,���������ҽ����¼������"
                Exit Function
            End If
        End If
    End If
        
    CheckPatiTurnLimit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim strTmp As String
    gstrLike = IIF(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    gbytCode = Val(zlDatabase.GetPara("���뷽ʽ"))
    gbytҩƷ������ʾ = zlDatabase.GetPara("ҩƷ������ʾ", , , 2)
    gbyt����ҩƷ��ʾ = zlDatabase.GetPara("����ҩƷ��ʾ", , , 0)
    gbln����ƥ�䷽ʽ�л� = zlDatabase.GetPara("����ƥ�䷽ʽ�л�", , , 1)

    '��¼ҽ��ʶ��ʱ����(����)
    gint��¼��� = Val(zlDatabase.GetPara(5, glngSys, , 30))
    
    '�����¿�ҽ�����
    gint�����¿�ҽ����� = Val(zlDatabase.GetPara(223, glngSys, , 1))
    
    '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    gbytDecPrice = Val(zlDatabase.GetPara(157, glngSys, , 5))
    gstrDecPrice = "0." & String(gbytDecPrice, "0")
        
    'ָ��ҩ��ʱ���ƿ��
    gblnStock = Val(zlDatabase.GetPara(18, glngSys)) <> 0
        
    '�Һ���Ч����
    strTmp = zlDatabase.GetPara(21, glngSys)
    If Len(strTmp) = 1 Then strTmp = strTmp & strTmp
    gint��ͨ�Һ����� = Val(Mid(strTmp, 1, 1))
    gint����Һ����� = Val(Mid(strTmp, 2, 1))
    
    '���δִ����Ŀ
    gbyt��Ժ���δִ�� = Val(zlDatabase.GetPara(22, glngSys))
    gbytת�Ƽ��δִ�� = Val(zlDatabase.GetPara(32, glngSys))
    
    gbyt��Ժ���δ��ҩ = Val(zlDatabase.GetPara(154, glngSys))
    gbytת�Ƽ��δ��ҩ = Val(zlDatabase.GetPara(155, glngSys))
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    '����ǩ����֤����
    gintCA = Val(zlDatabase.GetPara(25, glngSys))
    
    '����ǩ�����Ƴ���
    gstrESign = zlDatabase.GetPara(26, glngSys)
    
    '��ȡ������������
    Set grsSign = New ADODB.Recordset
    grsSign.Fields.Append "����ID", adBigInt
    grsSign.Fields.Append "����", adBigInt
    grsSign.Fields.Append "�Ƿ�����", adBigInt
    grsSign.CursorLocation = adUseClient
    grsSign.LockType = adLockOptimistic
    grsSign.CursorType = adOpenStatic
    grsSign.Open
    
    'һ��ͨ������֤
    strTmp = zlDatabase.GetPara(28, glngSys) & "|"
    gdblԤ��������鿨 = Val(Split(strTmp, "|")(0))
   
    'ָ��ҽ������������ִ��
    gblnָ��ҽ������������ִ�� = Val(zlDatabase.GetPara(34, glngSys)) <> 0
    
    'ҽ����������
    gstrҽ���������� = "'" & Replace(zlDatabase.GetPara(41, glngSys), "|", "','") & "'"

    '���ѷ�������
    gstr���ѷ������� = "'" & Replace(zlDatabase.GetPara(42, glngSys), "|", "','") & "'"
    
    '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    
    '���������Դ
    gint�����Դ = Val(zlDatabase.GetPara(55, glngSys, , 1))
    
    '���ﴦ����������
    gintRXCount = Val(zlDatabase.GetPara(56, glngSys))
    
    'ҽ��������
    gintҽ������ = Val(zlDatabase.GetPara(59, glngSys, , 1))
    
    '���ʷ���������ѽ��
    gcurMaxMoney = Val(zlDatabase.GetPara(60, glngSys))
    
    '���Ʊ������ģʽ
    gint���Ʊ��� = Val(zlDatabase.GetPara(61, glngSys))
    
    'סԺ�Զ�����
    gbytסԺ�Զ����� = Val(zlDatabase.GetPara(63, glngSys))
    
    '������뷽ʽ
    gstr������� = zlDatabase.GetPara(65, glngSys, , "11")
    
    'ҩƷ�������ҽ��
    gblnҩƷ�������ҽ�� = Val(zlDatabase.GetPara(69, glngSys)) = 1
    
    'Ƥ�Խ����Чʱ��
    gint�����Ǽ���Ч���� = Val(zlDatabase.GetPara(70, glngSys))
    
    '����ҽ��������Ч
    gbln����ҽ��������Ч = Val(zlDatabase.GetPara(71, glngSys)) = 1
    
    '�Ƿ�Ҫ�������������
    gbln�շ���� = Val(zlDatabase.GetPara(72, glngSys, , 1)) <> 0
    
    
    'ҽ���������ɻ��۵������
    gstrסԺ���ͻ��۵� = zlDatabase.GetPara(80, glngSys)
    gstr���﷢�ͻ��۵� = zlDatabase.GetPara(86, glngSys)

    '�����Զ�����
    gbln�����Զ����� = Val(zlDatabase.GetPara(92, glngSys)) <> 0
    
    '�Զ���ҩ��ҩ
    gbln�շѺ��Զ���ҩ = zlDatabase.GetPara(45, glngSys) = "1"
    
    '������Ŀ���ܼ����ۿ�
    gbln��������ۿ� = Val(zlDatabase.GetPara(93, glngSys)) <> 0
    
    '���ʱ����������۷���
    gbln�����������۷��� = Val(zlDatabase.GetPara(98, glngSys)) <> 0
    
    '����ҽ������ʱ����������
    gbln�������������� = Val(zlDatabase.GetPara(143, glngSys)) <> 0
    
    '���������ʱ,���������Ŀʱ,��λ����������
    gblnFeeKindCode = Val(zlDatabase.GetPara(144, glngSys)) <> 0 And Not gbln�շ����
    
    '����ҩƷ���ⷽʽ
    gbytMediOutMode = Val(zlDatabase.GetPara(150, glngSys))
    
    '��Һ��������(����Ϊ���������ġ���ҩ��)
    gstr��Һ�������� = Get��Һ��������
    gbyt������˷�ʽ = Val(zlDatabase.GetPara(185, glngSys))    '49501
    gblnδ��ƽ�ֹ���� = Val(zlDatabase.GetPara(215, glngSys)) = 1 '51612

    'ҽ����¼ʱ�Ƿ�ֻ����¼����
    gblnֻ����¼���� = Val(zlDatabase.GetPara(191, glngSys)) <> 0
    'ת�Ʋ��˵Ĳ�¼ʱ��
    glng��¼ʱ�� = Val(zlDatabase.GetPara(158, glngSys, , "24"))
    
    '�´�ҽ��ʱ��ʾ����
    gblnShowOrigin = Val(zlDatabase.GetPara(162, glngSys, , "1")) <> 0
    
    '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
    gblnִ��ǰ�Ƚ��� = Val(zlDatabase.GetPara(163, glngSys)) <> 0
    
    
    '��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�
    gstrҽ���˶� = zlDatabase.GetPara(186, glngSys)
        
    '����ҩ��ּ�����
    gblnKSSStrict = Val(zlDatabase.GetPara(187, glngSys)) <> 0
    gbln����ҩ��ʹ���Ա�ҩ = Val(zlDatabase.GetPara(188, glngSys)) <> 0
    
    '�Ƿ����������ּ�����
    gbln�����ּ����� = Val(zlDatabase.GetPara(209, glngSys)) <> 0
    
    '�Ƿ�����������ҽʦ��Ȩ����
    gbln������Ȩ���� = Val(zlDatabase.GetPara(217, glngSys)) <> 0
    
    '�Ƿ����ò���������ҽʦ�ﵽ�����ȼ��������
    gbln�����ȼ����� = Val(zlDatabase.GetPara(254, glngSys)) <> 0
    
    '�Ƿ�������Ѫ�ּ�����
    gbln��Ѫ�ּ����� = Val(zlDatabase.GetPara(216, glngSys)) <> 0
    '��Ѫ�����������
    gbln��Ѫ����������� = Val(zlDatabase.GetPara(218, glngSys)) <> 0
    '��Ѫ����ֻ�����м�������ҽʦ���
    gbln��Ѫ�����м����� = Val(zlDatabase.GetPara(219, glngSys)) <> 0
    
    '�Ƿ�װѪ��ϵͳ
    gblnѪ��ϵͳ = (IsSysSetUp(2200) And Val(zlDatabase.GetPara(236, glngSys)) <> 0)
    gbln��ʾѪҺ��� = Val(zlDatabase.GetPara(286, glngSys)) = 0 And gblnѪ��ϵͳ = True
    gbln�´���Ѫ����ȷ��ѪҺ��Ϣ = Val(zlDatabase.GetPara(293, glngSys)) <> 0 And gblnѪ��ϵͳ = True
    'ҽ������ʱ��������ԭ��
    gbyt����ԭ�� = Val(zlDatabase.GetPara(230, glngSys, , 1))
    
    '�����˵�ǰ���һ��߹Һſ�������֮�ڵ���ҽ��ʱ���Բ�¼�볬��˵������ʽΪ����id�����߷ָ�
    gstr��¼�������� = zlDatabase.GetPara(233, glngSys, , "")
    gstr��¼�������� = "," & Replace(gstr��¼��������, "|", ",") & ","
    
    '�����˵�ǰ��������֮�ڵ�ͣ��ʱ���Բ�¼��ͣ��˵������ʽΪ����id�����߷ָ�
    gstr�ɲ���ͣ��ԭ����� = zlDatabase.GetPara(285, glngSys, , "")
    gstr�ɲ���ͣ��ԭ����� = "," & Replace(gstr�ɲ���ͣ��ԭ�����, "|", ",") & ","
    
    '�����޸�n���ڵǼǵ�ҽ��ִ�м�¼
    gintҽ��ִ����Ч���� = Val(zlDatabase.GetPara(220, glngSys))
    
    'ת��ʱδ������ʵ��ݼ��
    gbytת��ʱδ������ʵ��ݼ�� = Val(zlDatabase.GetPara(227, glngSys))
    
    '��Ŀ�����������շѻ�������
    gbln�����������շѻ������� = Val(zlDatabase.GetPara(232, glngSys)) = 1
    
    '���뵥���û��ڣ��������뵥�����ʹ�����뵥�´�ҽ��
    Call Get���뵥��ز���
    
    gbln�������� = Val(zlDatabase.GetPara(240, glngSys)) = 1
    gbln����Ӱ����Ϣϵͳ�ӿ� = Val(zlDatabase.GetPara(255, glngSys)) = 1
    gbln����ҩƷ�ֿ����� = Val(zlDatabase.GetPara(262, glngSys)) <> 0
    gblnҽ����ֹԭ�� = Val(zlDatabase.GetPara(271, glngSys)) = 1
    gbln����ҩ�����հ������������� = Val(zlDatabase.GetPara(274, glngSys)) = 1
    gbln��������´�ҽ������ = Val(zlDatabase.GetPara(302, glngSys)) = 1
    gbln��ϵͳ = "" <> GetParaURL("ҩʦ�������", "�������ѯ")
    
    On Error Resume Next
    If gobjDrugExplain Is Nothing Then Set gobjDrugExplain = CreateObject("zlKnowledgeConvert.CallView")
    err.Clear: On Error GoTo errH
    
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiYear(lng����ID As Long) As Integer
'���ܣ���ȡ���˵�׼ȷ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intYear As Integer
    
    On Error GoTo errH
    
    strSQL = "Select Sysdate as ��ǰ,��������,���� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!��������) Then
            intYear = Year(rsTmp!��ǰ) - Year(rsTmp!��������)
            If Format(rsTmp!��ǰ, "MMdd") < Format(rsTmp!��������, "MMdd") Then
                intYear = intYear - 1
            End If
            If intYear < 0 Then intYear = 0
        Else
            intYear = Val(NVL(rsTmp!����))
        End If
    End If
    GetPatiYear = intYear
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getȱʡ�÷�ID(ByVal int���� As Integer, ByVal int��Դ As Integer, Optional ByVal strFilter As String = "", Optional ByVal lng�������� As Long) As Long
'���ܣ�����ȱʡ�ĸ�ҩ;������ҩ�巨
'������int����=2-��ҩ;��,3-��ҩ�巨,4-��ҩ�÷�,6-�ɼ�����(����)
'      int��Դ=1-����,2-סԺ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str������� As String
    
    If lng�������� = 1 Then
        str������� = ",1,2,3,"
    Else
        str������� = "," & int��Դ & ",3,"
    End If
    
    strSQL = "Select ID From ������ĿĿ¼" & _
        " Where ���='E' And ��������=[1] " & strFilter & _
        " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
        " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
        " And Instr([2],','||�������||',')>0 And Rownum<100" & _
        " Order by ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", CStr(int����), str�������)
    If Not rsTmp.EOF Then
        Getȱʡ�÷�ID = rsTmp!ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Check�����÷�(ByVal lng�÷�ID As Long, ByVal lng��Ŀid As Long, ByVal int��Դ As Integer, Optional ByVal lng�շ�ϸĿID As Long, Optional ByVal lng�������� As Long) As Boolean
'���ܣ����ָ�����÷��Ƿ�������ָ������Ŀ
'������int��Դ=1-����,2-סԺ
'      lng��������   0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnCheckUseM As Boolean
    Dim bln���� As Boolean
    Dim str������� As String
    
    On Error GoTo errH
    
    If lng�������� = 1 Then
        str������� = ",1,2,3,"
    Else
        str������� = "," & int��Դ & ",3,"
    End If
    
     '������ҩ��ȡ�÷��Ƿ��ϸ����
    blnCheckUseM = CheckDrugUseM(lng�շ�ϸĿID, lng��Ŀid)

    '�����Ŀ�÷�����
    strSQL = "Select Count(A.�÷�ID) as ����,Max(Decode(A.�÷�ID,[2],1,0)) as ָ��" & _
        " From �����÷����� A,������ĿĿ¼ B" & _
        " Where A.�÷�ID=B.ID And Instr([3],','||B.�������||',')>0 And A.��ĿID=[1] And A.����>0" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid, lng�÷�ID, str�������)
    
    If Not blnCheckUseM Then
        If NVL(rsTmp!����, 0) <= 1 Then
            Check�����÷� = True
    
        ElseIf NVL(rsTmp!ָ��, 0) = 1 Then
            Check�����÷� = True
        End If
    Else
        If NVL(rsTmp!����, 0) = 1 And (Not blnCheckUseM) Then
            Check�����÷� = True
            Exit Function
        ElseIf NVL(rsTmp!����, 0) = 0 Then
            If (Not blnCheckUseM) And lng�շ�ϸĿID = 0 Then
                Check�����÷� = True
                Exit Function
            Else
                bln���� = True
            End If
        ElseIf NVL(rsTmp!ָ��, 0) = 1 Then
            Check�����÷� = True
            Exit Function
        End If
    
        '���ҩƷ�÷�����
        If blnCheckUseM And lng�շ�ϸĿID <> 0 Then
            strSQL = "Select Count(A.�÷�ID) as ����,Max(Decode(A.�÷�ID,[2],1,0)) as ָ��" & _
                " From ҩƷ�÷����� A,������ĿĿ¼ B" & _
                " Where A.�÷�ID=B.ID And Instr([3],','||B.�������||',')>0 And A.ҩƷID=[1] And A.����>0" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng�շ�ϸĿID, lng�÷�ID, str�������)
            If NVL(rsTmp!����, 0) = 0 And bln���� Then
                Check�����÷� = True
                Exit Function
            ElseIf NVL(rsTmp!ָ��, 0) = 1 Then
                Check�����÷� = True
                Exit Function
            End If
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDrugUseM(ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long) As Boolean
    '�ж�ҩƷ�Ƿ������ϸ�����÷�����
    On Error GoTo errH
    
    'ʹ��ҩƷID��ѯ
    If lngҩƷID <> 0 Then
        If Val(Sys.RowValue("ҩƷ���", lngҩƷID, "�ϸ�����÷�����", "ҩƷID") & "") = 1 Then
            CheckDrugUseM = True
            Exit Function
        Else
            '��ȡҩ��ID
            If lngҩ��ID = 0 Then lngҩ��ID = Val(Sys.RowValue("ҩƷ���", lngҩƷID, "ҩ��ID", "ҩƷID") & "")
        End If
    End If
    
    'ʹ��ҩ��ID��ѯ
    If lngҩ��ID <> 0 Then
        If Val(Sys.RowValue("ҩƷ����", lngҩ��ID, "�ϸ�����÷�����", "ҩ��ID") & "") = 1 Then
            CheckDrugUseM = True
            Exit Function
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function Check�ϰల��(ByVal blnҩ�� As Boolean) As Boolean
'���ܣ����ҽԺ�Ŀ����Ƿ�ʹ�����ϰల��
'������blnҩ��=�Ǽ��ҩ���ϰ໹����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Static blnҩ��Load As Boolean
    Static blnҩ��Last As Boolean
    Static bln��ҩLoad As Boolean
    Static bln��ҩLast As Boolean
    
    If blnҩ�� Then '�Ƿ��а���ֻ���ȡһ��
        If blnҩ��Load Then Check�ϰల�� = blnҩ��Last: Exit Function
    Else
        If bln��ҩLoad Then Check�ϰల�� = bln��ҩLast: Exit Function
    End If
    
    On Error GoTo errH
    
    If blnҩ�� Then
        strSQL = "Select 1 From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� IN('��ҩ��','��ҩ��','��ҩ��') And Rownum<2"
    Else
        strSQL = "Select 1 From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� Not IN('��ҩ��','��ҩ��','��ҩ��') And Rownum<2"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "Check�ϰల��")
    Check�ϰల�� = rsTmp.RecordCount > 0
    
    If blnҩ�� Then
        blnҩ��Load = True: blnҩ��Last = Check�ϰల��
    Else
        bln��ҩLoad = True: bln��ҩLast = Check�ϰల��
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����Ա����ID(ByVal int������� As Integer, Optional ByVal lngĬ�ϲ��� As Long) As Long
'���ܣ�ȡ����Ա���������ָ������Ĳ��ţ�ȱʡ��������
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    
    If blnNew Then
        strSQL = "Select Distinct B.����ID,Nvl(B.ȱʡ,0) as ȱʡ,C.������� From ������Ա B,��������˵�� C" & _
            " Where B.��ԱID = [1] And B.����ID=C.����ID" & _
            " Order by ȱʡ Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    
    '74794,Ƚ����,2014-7-18,��ʿ�ڼ���ʱ����ó��׷���ʱδʹ�ó��׷����ڵ�ִ�п���
    If lngĬ�ϲ��� <> 0 Then
        rsTmp.Filter = "(������� = 3 and ����ID = " & lngĬ�ϲ��� & ") " & _
                    "or (������� = " & int������� & " and ����ID = " & lngĬ�ϲ��� & ")"
        If Not rsTmp.EOF Then Get����Ա����ID = rsTmp!����ID: Exit Function
    End If
    
    rsTmp.Filter = "������� = 3 or ������� = " & int�������
    
    If Not rsTmp.EOF Then
        Get����Ա����ID = rsTmp!����ID
    Else
        Get����Ա����ID = UserInfo.����ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�շ�ִ�п���ID(ByVal lng����ID As Long, lng��ҳID As Long, _
    ByVal str��� As String, ByVal lng��Ŀid As Long, ByVal intִ�п��� As Integer, _
    ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal lngִ�п���ID As Long, _
    Optional ByVal bytMode As Byte, Optional ByVal bytCallBy As Byte, _
    Optional ByVal int���ó��� As Integer = 1, _
    Optional lng����ȱʡִ�п��� As Long = 0) As Long
'���ܣ���ȡ��ҩ�շ���Ŀ��ִ�п���
'������int��Χ=1.����,2-סԺ
'      lngִ�п���ID=ָ����ȱʡִ�п���ID(����ҩƷ������)
'      bytMode=1-Ҫ����ȱʡֵ,0-����
'      bytCallBy=0-ҽ���������,1-���ѳ������
'      int���ó���=1-����,2-סԺ
'      lng����ȱʡִ�п���-ȱʡִ�п���ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strҩ�� As String, lngҩ�� As Long
    Dim lng���˲���ID As Long, bytDay As Byte
    
    On Error GoTo errH
    
    If str��� = "4" Then
        lngҩ�� = Val(zlDatabase.GetPara(IIF(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ���ϲ���", glngSys, _
            IIF(bytCallBy = 1, pҽ�����ѹ���, IIF(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�)), , , , , lng���˿���ID))
        
        If lng��ҳID > 0 Then
            lng���˲���ID = GetPatiUnitID(lng����ID, lng��ҳID)
        End If
        '��ִ�п�������ʱ
        strSQL = _
            " Select Distinct" & _
            "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            " And B.������� IN([1],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And (A.������Դ is NULL Or A.������Դ=[1])" & _
            " And (A.��������ID is NULL Or instr([2],','||a.��������id||',')>0)" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int��Χ, "," & lng���˿���ID & "," & lng���˲���ID & ",", lng��Ŀid)
        If Not rsTmp.EOF Then
            If bytMode = 1 Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID  '�����û�У��򷵻ص�һ�����õ�ִ�п���
            
            '1:ȱʡΪָ����(ҽ����)ִ�п���,�����Ƿ�����ڲ��˿���
            rsTmp.Filter = "ִ�п���ID=" & lngִ�п���ID
            
            '2.ȱʡΪ����ָ����ȱʡ����
            If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lngҩ��
            
            '3:�����ɷ����ڲ��˿��ҵ�ִ�п���
            If rsTmp.EOF Then
                '2.0 ��������д���ȱʡ��ִ�п���,��ȱʡΪ����ָ����ȱʡ����
                If lng����ȱʡִ�п��� <> 0 Then
                    rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                    If Not rsTmp.EOF Then
                            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                    End If
                End If
                '2.1:����ȱʡΪ���˿���
                If lngִ�п���ID <> lng���˿���ID And lngҩ�� <> lng���˿���ID Then
                    rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˿���ID
                End If
                '3.2:����ȱʡΪ���˲���
                If rsTmp.EOF And lng��ҳID <> 0 Then
                    If lng���˲���ID <> 0 And lng���˲���ID <> lng���˿���ID And lng���˲���ID <> lngִ�п���ID And lng���˲���ID <> lngҩ�� Then
                        rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˲���ID
                    End If
                End If
            End If
            '3.3:�ɷ����ڲ��˿��ҵ�һ��ִ�п���
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            
            '3.4�ɷ��������п��ҵĵ�ǰ���˿���ִ��
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0 And ִ�п���ID=" & lng���˿���ID
            
            '4:�����û�У��򷵻�0���ڼ��
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zlDatabase.GetPara(IIF(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIF(bytCallBy = 1, pҽ�����ѹ���, IIF(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�)), , , , , lng���˿���ID))
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zlDatabase.GetPara(IIF(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIF(bytCallBy = 1, pҽ�����ѹ���, IIF(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�)), , , , , lng���˿���ID))
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zlDatabase.GetPara(IIF(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIF(bytCallBy = 1, pҽ�����ѹ���, IIF(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�)), , , , , lng���˿���ID))
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not Check�ϰల��(True) Then
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                " And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                " And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strҩ��, int��Χ, lng���˿���ID, lng��Ŀid, bytDay)
        If Not rsTmp.EOF Then
            If lng����ȱʡִ�п��� <> 0 Then
                rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                If Not rsTmp.EOF Then
                        Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                End If
            End If
            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
            rsTmp.Filter = "ִ�п���ID=" & lngִ�п���ID
            If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lngҩ��
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    Else
        Select Case intִ�п���
            Case 0 '0-����ȷ����
                If lng����ȱʡִ�п��� <> 0 Then
                    Get�շ�ִ�п���ID = lng����ȱʡִ�п���: Exit Function
                End If
                Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ)
            Case 1 '1-�������ڿ���
                Get�շ�ִ�п���ID = lng���˿���ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get�շ�ִ�п���ID = lng���˿���ID
                Else
                    Get�շ�ִ�п���ID = GetPatiUnitID(lng����ID, lng��ҳID)
                End If
            Case 3 '3-����Ա���ڿ���
                Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ, lng����ȱʡִ�п���)
            Case 4 '4-ָ������
                strSQL = "Select Distinct Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID,Decode(A.������Դ,Null,2,1) as ����" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.�շ�ϸĿID=[1] And A.ִ�п���ID=B.����ID" & _
                    " And B.������� IN([2],3) And (A.������Դ is NULL Or A.������Դ=[2])" & _
                    " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                    " And A.ִ�п���ID=C.ID And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " Order by ����" 'Ĭ�Ͽ�������
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid, int��Χ, lng���˿���ID)
                If Not rsTmp.EOF Then
                    If lng����ȱʡִ�п��� <> 0 Then
                         rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                         If Not rsTmp.EOF Then
                                 Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                         End If
                     End If
                    Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                    rsTmp.Filter = "��������ID=" & lng���˿���ID
                    If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                End If
            Case 6 '6-���������ڿ���
                Get�շ�ִ�п���ID = lng��������ID
        End Select
        If Get�շ�ִ�п���ID = 0 Then Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ִ�п���ID(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal str��� As String, ByVal lng��Ŀid As Long, ByVal lngҩƷID As Long, ByVal intִ�п��� As Integer, _
    ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, ByVal int��Ч As Integer, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal blnByȱʡ As Boolean, Optional ByVal lngPreID As Long, Optional ByVal int���ó��� As Integer = 1) As Long
'���ܣ�����������Ŀִ�п�����Ϣ����ȱʡ��ִ�п���ID
'������lngҩƷID=ҩƷID,ȷ�������ʱҪ��
'      intִ�п���=��Ŀִ�п��ұ�־
'      lng���˿���ID=���˿���ID
'      lng��ҩ��,lng��ҩ��,lng��ҩ��=ҩƷȱʡҩ��,ҩƷ��ʱ��Ҫ
'      int��Ч=0-����,1-����,��������Ҫ�ж��ϰ�ʱ��
'      int��Χ=1-����,2-סԺ(ȱʡ)
'      blnByȱʡ=��ȡȱʡҩ��ʱ�����������ָ�����Ƿ񰴱���ȱʡָ����ҩ������û���򲻷���(�������һ�����Է���)
'      lngPreID=ǰ��ҽ����ִ�п��ң�������˸�ֵ�����������ȱʡ�����ʣ�Ŀǰֻ����ҩƷ
'      int���ó���=1-����,2-סԺ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strҩ�� As String, lngҩ�� As Long, strҩ��IDs As String
    Dim bytDay As Byte, bln�ϰల�� As Boolean
    Dim bln��� As Boolean, strStock As String
    
    On Error GoTo errH
    
    '���û��ض����ֱ�ӷ���
    strSQL = "Select zl_ClinicExeDept([1]," & IIF(lng��ҳID = 0, "Null", "[2]") & ",[3],[4]," & IIF(lngҩƷID = 0, "Null", "[5]") & ",[6]," & IIF(lng��������ID = 0, "Null", "[7]") & ") as ִ�п���ID From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get����ִ�п���ID", lng����ID, lng��ҳID, str���, lng��Ŀid, lngҩƷID, lng���˿���ID, lng��������ID)
    If Not rsTmp.EOF Then
        If NVL(rsTmp!ִ�п���ID, 0) <> 0 Then
            Get����ִ�п���ID = rsTmp!ִ�п���ID
            Exit Function
        End If
    End If
    
    'δ���Ʒ��ص�����¸��ݳ��������
    If InStr(",5,6,7,", str���) > 0 Then
        bln��� = ((int��Ч = 1 Or gblnҩƷ�������ҽ��) And lngҩƷID <> 0) Or lngҩƷID <> 0
        
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zlDatabase.GetPara(IIF(int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, IIF(int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID))
            strҩ��IDs = zlDatabase.GetPara(IIF(int���ó��� = 2, "סԺ", "����") & "������ҩ��", glngSys, IIF(int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zlDatabase.GetPara(IIF(int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, IIF(int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID))
            strҩ��IDs = zlDatabase.GetPara(IIF(int���ó��� = 2, "סԺ", "����") & "���ó�ҩ��", glngSys, IIF(int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zlDatabase.GetPara(IIF(int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, IIF(int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID))
            strҩ��IDs = zlDatabase.GetPara(IIF(int���ó��� = 2, "סԺ", "����") & "������ҩ��", glngSys, IIF(int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
        End If
        
        'ҩƷ�������
        If bln��� And strҩ��IDs <> "" Then
            If gblnStock Then
                strStock = " And Exists(" & _
                    " Select 1 From ҩƷ���" & _
                    " Where (Nvl(����,0)=0 Or Ч�� Is Null Or Ч��>Trunc(Sysdate))" & _
                    " And ����=1 And ҩƷID=[4] And �ⷿID=A.ִ�п���ID" & _
                    " And ��������>0 And Instr('," & strҩ��IDs & ",',','||�ⷿID||',')>0)"
            Else
                strStock = " And Instr('," & strҩ��IDs & ",',','||A.ִ�п���ID||',')>0"
            End If
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If int��Χ = 1 Then bln�ϰల�� = Check�ϰల��(True) 'סԺҽ������ҩ���ϰల��
        If Not bln�ϰల�� Then
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                 IIF(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & strStock & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[6]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                IIF(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & strStock & _
                " Order by B.�������,C.����"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strҩ��, int��Χ, lng���˿���ID, lngҩƷID, lng��Ŀid, bytDay)
        If Not rsTmp.EOF Then
            If blnByȱʡ And (lngҩ�� <> 0 Or lngPreID <> 0) Then
                If rsTmp.RecordCount > 1 Then
                    rsTmp.Filter = "ִ�п���ID=" & lngPreID
                    If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lngҩ��
                End If
            Else
                Get����ִ�п���ID = rsTmp!ִ�п���ID
                rsTmp.Filter = "ִ�п���ID=" & lngPreID
                If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lngҩ��
                If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            End If
            If Not rsTmp.EOF Then Get����ִ�п���ID = rsTmp!ִ�п���ID
        End If
    Else
        Select Case intִ�п���
            Case 0 '0-��ִ�еĶ���
                Exit Function
            Case 1 '1-�������ڿ���
                Get����ִ�п���ID = lng���˿���ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get����ִ�п���ID = lng���˿���ID
                Else
                    Get����ִ�п���ID = GetPatiUnitID(lng����ID, lng��ҳID)
                End If
            Case 3 '3-����Ա���ڿ���
                Get����ִ�п���ID = Get����Ա����ID(int��Χ)
            Case 4 '4-ָ������
                If lng��Ŀid = 0 Then
                    If int��Χ = 1 Then
                        Get����ִ�п���ID = lng���˿���ID
                    ElseIf int��Χ = 2 Then
                        Get����ִ�п���ID = GetPatiUnitID(lng����ID, lng��ҳID)
                    End If
                Else
                    If int��Ч = 1 Then bln�ϰల�� = Check�ϰల��(False)
                    If Not bln�ϰల�� Then
                        strSQL = "Select Distinct Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID,Decode(A.������Դ,Null,2,1) as ����" & _
                            " From ����ִ�п��� A,��������˵�� B,���ű� C" & _
                            " Where A.ִ�п���ID=B.����ID And A.������ĿID=[1]" & _
                            " And B.������� IN([2],3) And (A.������Դ is NULL Or A.������Դ=[2])" & _
                            " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                            " And A.ִ�п���ID=C.ID And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " Order by ����" 'Ĭ�Ͽ�������
                    Else
                        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
                        strSQL = _
                            " Select Distinct Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID,Decode(A.������Դ,Null,2,1) as ����" & _
                            " From ����ִ�п��� A,���Ű��� B,��������˵�� C,���ű� D" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.����=[4]" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.��ʼʱ��,'HH24:MI:SS') and To_Char(B.��ֹʱ��,'HH24:MI:SS') " & _
                            " And A.ִ�п���ID=C.����ID And C.������� IN([2],3) And (A.������Դ is NULL Or A.������Դ=[2])" & _
                            " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                            " And A.ִ�п���ID=D.ID And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & _
                            " And (D.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� is NULL)" & _
                            " And A.������ĿID=[1]" & _
                            " Order by ����"
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid, int��Χ, lng���˿���ID, bytDay)
                    If Not rsTmp.EOF Then
                        Get����ִ�п���ID = rsTmp!ִ�п���ID
                        rsTmp.Filter = "��������ID=" & lng���˿���ID
                        If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lng���˿���ID
                        If rsTmp.EOF And int��Χ = 2 Then rsTmp.Filter = "ִ�п���ID=" & GetPatiUnitID(lng����ID, lng��ҳID)
                        If Not rsTmp.EOF Then Get����ִ�п���ID = rsTmp!ִ�п���ID
                    ElseIf gblnָ��ҽ������������ִ�� Then
                        If int��Χ = 1 Then
                            Get����ִ�п���ID = lng���˿���ID
                        ElseIf int��Χ = 2 Then
                            Get����ִ�п���ID = GetPatiUnitID(lng����ID, lng��ҳID)
                        End If
                    End If
                End If
            Case 5 '5-Ժ��ִ��
                Exit Function
            Case 6 '6-���������ڿ���
                Get����ִ�п���ID = lng��������ID
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckExecDeptValidate(ByVal lngִ�п���ID As Long, ByVal lng���˿���ID As Long, ByVal int��Χ As Integer, ByVal lng������ĿID As Long) As Boolean
'���ܣ����ָ����ִ�п����Ƿ���Ч
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select 1" & vbNewLine & _
            "From ���ű� A, ��������˵�� C" & vbNewLine & _
            "Where a.Id = [1] And a.Id = c.����id And c.������� In ([3], 3) And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)" & vbNewLine & _
            "  And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And (Exists" & vbNewLine & _
            " (Select 1 From ����ִ�п��� B Where b.������ĿID =[4] And b.ִ�п���ID = a.Id And (b.������Դ Is Null Or b.������Դ = [3]) And (b.��������ID Is Null Or b.��������ID = [2]))" & vbNewLine & _
            ") And Rownum<2"
            'Or��һ������ţ�41496��Ϊ�˳�����������ִ�п��ң����ó���ʱ��ʱ��֤������ҿ��á�
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckExecDeptValidate", lngִ�п���ID, lng���˿���ID, int��Χ, lng������ĿID)
    CheckExecDeptValidate = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ҩ��IDs(ByVal str��� As String, ByVal lng��Ŀid As Long, _
    ByVal lngҩƷID As Long, ByVal lng����id As Long, Optional ByVal int��Χ As Integer = 2) As String
'���ܣ���ȡҩƷ����Ч����ִ�п���ID��,�����ж�ȱʡִ�п���
'������lng����ID=���˿���ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strҩ�� As String
    Dim bytDay As Byte, bln�ϰల�� As Boolean
    Dim strҩ��IDs As String, str����ҩ�� As String
    
    'ϵͳ����ָ��ҩƷִ�п���,������ȡ���п�ѡ�Ĺ���ѡ��
    If str��� = "5" Then
        strҩ�� = "��ҩ��"
        str����ҩ�� = zlDatabase.GetPara(Decode(int��Χ, 1, "����", 2, "סԺ", "") & "������ҩ��", glngSys, Decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , lng����id)
    ElseIf str��� = "6" Then
        strҩ�� = "��ҩ��"
        str����ҩ�� = zlDatabase.GetPara(Decode(int��Χ, 1, "����", 2, "סԺ", "") & "���ó�ҩ��", glngSys, Decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , lng����id)
    ElseIf str��� = "7" Then
        strҩ�� = "��ҩ��"
        str����ҩ�� = zlDatabase.GetPara(Decode(int��Χ, 1, "����", 2, "סԺ", "") & "������ҩ��", glngSys, Decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , lng����id)
    End If
            
    'ҩƷ��ϵͳָ���Ĵ���ҩ������
    If int��Χ = 1 Then bln�ϰల�� = Check�ϰల��(True) 'סԺҽ������ҩ���ϰల��
    If Not bln�ϰల�� Then
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIF(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
            " And B.������� IN([2],3) And B.����ID=C.ID" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            IIF(int��Χ <> 3, " And (A.������Դ is NULL Or A.������Դ=[2])", "") & _
            IIF(lng����id <> 0, " And (A.��������ID is NULL Or A.��������ID=[3])", "") & _
            IIF(lngҩƷID <> 0, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]")
    Else
        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIF(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
            " And B.������� IN([2],3) And B.����ID=C.ID" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And D.����ID=C.ID And D.����=[6]" & _
            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
            IIF(int��Χ <> 3, " And (A.������Դ is NULL Or A.������Դ=[2])", "") & _
            IIF(lng����id <> 0, " And (A.��������ID is NULL Or A.��������ID=[3])", "") & _
            IIF(lngҩƷID <> 0, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]")
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strҩ��, int��Χ, lng����id, lngҩƷID, lng��Ŀid, bytDay)
    Do While Not rsTmp.EOF
        If str����ҩ�� = "" Then
            strҩ��IDs = strҩ��IDs & "," & rsTmp!ID
        ElseIf InStr("," & str����ҩ�� & ",", "," & rsTmp!ID & ",") > 0 Then
            strҩ��IDs = strҩ��IDs & "," & rsTmp!ID
        End If
        rsTmp.MoveNext
    Loop
    Get����ҩ��IDs = Mid(strҩ��IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���÷��ϲ���IDs(ByVal lng����ID As Long, ByVal lng����id As Long, Optional ByVal int��Χ As Integer = 2, Optional ByVal lng������ĿID As Long) As String
'���ܣ���ȡ���ĵ���Ч����ִ�п���ID��,�����ж�ȱʡִ�п���
'������lng����ID=���˿���ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str���ϲ���IDs As String
    
    strSQL = _
        " Select Distinct C.ID" & _
        " From " & IIF(lng����ID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
        " And B.������� IN([1],3) And B.����ID=C.ID " & IIF(lng����ID <> 0, " And A.�շ�ϸĿID=[3]", " And A.������ĿID=[4]") & _
        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
        IIF(int��Χ <> 3, " And (A.������Դ is NULL Or A.������Դ=[1])", "") & _
        IIF(lng����id <> 0, " And (A.��������ID is NULL Or A.��������ID=[2])", "")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int��Χ, lng����id, lng����ID, lng������ĿID)
    Do While Not rsTmp.EOF
        str���ϲ���IDs = str���ϲ���IDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    Get���÷��ϲ���IDs = Mid(str���ϲ���IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ִ�п���(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    objCbo As Object, ByVal str��� As String, ByVal lng��Ŀid As Long, ByVal lngҩƷID As Long, _
    ByVal intִ�п��� As Integer, ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, _
    ByVal lng��ǰִ��ID As Long, ByVal int��Ч As Integer, Optional ByVal int��Χ As Integer = 2, _
    Optional ByVal bln��Һ As Boolean, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal bln������Һ���� As Boolean, Optional ByVal lng�÷�ID As Long, Optional ByVal bln����Ӫ�� As Boolean, Optional ByVal lng����ID As Long, Optional ByRef strִ�п���ids As String, Optional ByVal lng�������� As Long) As Boolean
'���ܣ�����������Ŀִ�п�����Ϣ���ؿ��õ�ִ�п�����ָ����������
'������intִ�п���=��Ŀִ�п��ұ�־
'      lng���˿���ID=���˿���ID
'      lng��ǰִ��ID=ҽ����ǰ��ִ�п���ID
'      int��Ч=0-����,1-����,��������Ҫ�ж��ϰ�ʱ��
'      int��Χ=1-����,2-סԺ(ȱʡ)
'      bln��Һ=��ǰҩƷ�Ƿ�������Һ��ģ���ҩ;��Ϊ��Һ��
'      blnEditable=false��ʾ��ǰҽ�����ɱ༭
'      bln������Һ����=���һ����Һ����ҽ����ִ�п��Ҳ�����Һ�������ģ������ҽ��Ҳ��������Ϊ��Һ��������
'      lng�÷�ID=�����ҩƷ������ҩƷ��Ӧ�ĸ�ҩ;��ID,�������Ѫ���飬����ɼ�����Ѫ;��
'      bln����Ӫ��=����Ӫ����ҩƷ��ֻ������ҩ������
'      strִ�п���ids ���η���ǰ�����п�ѡ����IDs�������ŷָ�,��δ����
'      lng�������� 0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��� (��ҩ;������ҩ�巨����ҩ����������ɼ�����Ѫ;������Ѫ�ɼ�) �������۲��˲����ַ�����󣬿�������
'˵�����Է�ҩҽ��,��ǰ��ִ�п��ҿ�����ǿ��ѡ�������,��Ҫ��ʾ��ѡ�����;��ѡ���������һ��������ѡ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strҩ�� As String, strҩ��IDs As String
    Dim bytDay As Byte, bln�ϰల�� As Boolean
    Dim bln��� As Boolean, i As Long
    Dim strStock As String, lng����Ա����ID As Long
    Dim str��Һ���ÿ��� As String
    Dim int��Һ����ҽ����Ч As Integer
    Dim bln��Ĭ���������� As Boolean
    Dim blnָ����Һ�������� As Boolean
    Dim str��Һ��ҩ;�� As String
    Dim strFilter As String
    Dim bln������Һҩ As Boolean '�Ա�����ȡ����Ժ��
    Dim strTmp As String
    Dim strĬ��ҩ�� As String
    Dim strSQLOther As String
    
    If str��� = "4" Then
        strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.�������" & _
            " From " & IIF(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            " And B.������� IN([2],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And (A.������Դ is NULL Or A.������Դ=[2])" & _
            " And (A.��������ID is NULL Or A.��������ID=[3] or exists (Select 1 From �������Ҷ�Ӧ x Where x.����id=[3] and x.����id=A.��������ID))" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            IIF(lngҩƷID <> 0, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
            " Order by B.�������,C.����"
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        bln��� = ((int��Ч = 1 Or gblnҩƷ�������ҽ��) And lngҩƷID <> 0) Or lngҩƷID <> 0
        
        'ϵͳ����ָ��ҩƷִ�п���,������ȡ���п�ѡ�Ĺ���ѡ��
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            strҩ��IDs = zlDatabase.GetPara(IIF(int��Χ = 2, "סԺ", "����") & "������ҩ��", glngSys, IIF(int��Χ = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
            strĬ��ҩ�� = zlDatabase.GetPara(IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, IIF(int��Χ = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            strҩ��IDs = zlDatabase.GetPara(IIF(int��Χ = 2, "סԺ", "����") & "���ó�ҩ��", glngSys, IIF(int��Χ = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
            strĬ��ҩ�� = zlDatabase.GetPara(IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, IIF(int��Χ = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            strҩ��IDs = zlDatabase.GetPara(IIF(int��Χ = 2, "סԺ", "����") & "������ҩ��", glngSys, IIF(int��Χ = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
            strĬ��ҩ�� = zlDatabase.GetPara(IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, IIF(int��Χ = 2, pסԺҽ���´�, p����ҽ���´�), , , , , lng���˿���ID)
        End If
            
        'ҩƷ�������
        If bln��� And strҩ��IDs <> "" And blnEditable Then
            If gblnStock Then
                strStock = " And Exists(" & _
                    " Select 1 From ҩƷ���" & _
                    " Where (Nvl(����,0)=0 Or Ч�� Is Null Or Ч��>Trunc(Sysdate))" & _
                    " And ����=1 And ҩƷID=[4] And �ⷿID=A.ִ�п���ID" & _
                    " And ��������>0 And Instr('," & strҩ��IDs & ",',','||�ⷿID||',')>0)"
            Else
                strStock = " And Instr('," & strҩ��IDs & ",',','||A.ִ�п���ID||',')>0"
            End If
        End If
        
        '�Ƿ�������Һ��������
        'bln������Һҩ  �Ա�����ȡ����Ժ�� ��3������������ʱΪ������
        strTmp = ""
        bln������Һҩ = Val(zlDatabase.GetPara("�Ա�ҩ��������������", glngSys, p��Һ��������, "0")) = 1
        strTmp = strTmp & IIF(bln������Һҩ, "1", "0")
        bln������Һҩ = Val(zlDatabase.GetPara("��ȡҩ��������������", glngSys, p��Һ��������, "0")) = 1
        strTmp = strTmp & IIF(bln������Һҩ, "1", "0")
        bln������Һҩ = Val(zlDatabase.GetPara("��Ժ��ҩ��������������", glngSys, p��Һ��������, "0")) = 1
        strTmp = strTmp & IIF(bln������Һҩ, "1", "0")
         
        bln������Һҩ = strTmp = "000"
       
        If bln������Һ���� And bln������Һҩ Then
            '71101��ѡ���Ա�ҩ������ִ������ʱ��Ҳ����ѡ����Һ�������ģ���Ϊ�п�����Һ��������ͬʱ������ͨҩ���Ĺ�����
            'strStock = strStock & " And A.ִ�п���ID <> [12]  "
        Else
            If gstr��Һ�������� <> "" And blnEditable Then
                If bln��Һ And (int��Χ = 2 Or lng�������� = 1) Then
                    str��Һ���ÿ��� = zlDatabase.GetPara("��Դ����", glngSys, p��Һ��������, "")
                    int��Һ����ҽ����Ч = Val(zlDatabase.GetPara("ҽ������", glngSys, p��Һ��������, "1")) - 1
                    str��Һ��ҩ;�� = zlDatabase.GetPara("��Һ��ҩ;��", glngSys, p��Һ��������)
                    If (str��Һ���ÿ��� = "" Or InStr("," & str��Һ���ÿ��� & ",", "," & lng����ID & ",") > 0) And (int��Ч = int��Һ����ҽ����Ч Or int��Һ����ҽ����Ч = -1) And _
                            (InStr("," & str��Һ��ҩ;�� & ",", "," & lng�÷�ID & ",") > 0 Or str��Һ��ҩ;�� = "") Or bln����Ӫ�� Then
                        blnָ����Һ�������� = True
                    Else
                        'bln��Ĭ����������=��������Һ�������ģ�������Ч\����\��ҩ;��������ƥ�䣬���������ķ����
                        bln��Ĭ���������� = True
                    End If
                Else
                    '������Һ��Ļ�Χ����סԺ�ģ�����Һ�������ķ����
                    bln��Ĭ���������� = True
                End If
            End If
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If int��Χ = 1 Then bln�ϰల�� = Check�ϰల��(True) 'סԺҽ������ҩ���ϰల��
        If Not bln�ϰల�� Then
            strSQL = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.�������,Decode(c.Id,[13],0,1) as Ĭ�ϲ���" & _
                " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                IIF(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & strStock & _
                IIF(bln��Ĭ����������, " Order By Ĭ�ϲ���,Decode(instr('," & gstr��Һ�������� & ",',',' || C.ID || ','),0,B.�������,9),C.����", _
                " Order by Ĭ�ϲ���,B.�������,C.����")
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.�������,Decode(c.Id,[13],0,1) as Ĭ�ϲ���" & _
                " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[8]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                IIF(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & strStock & _
                IIF(bln��Ĭ����������, " Order By Ĭ�ϲ���,Decode(instr('," & gstr��Һ�������� & ",',',' || C.ID || ','),0,B.�������,9),C.����", _
                " Order by Ĭ�ϲ���,B.�������,C.����")
        End If
    Else
        Select Case intִ�п���
            Case 0, 5 '0-��ִ�еĶ���,5-Ժ��ִ��
                Get����ִ�п��� = True: Exit Function
            Case 1 '1-�������ڿ���
                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([3],[6]) Order by ����"
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([3],[6]) Order by ����"
                Else
                    strSQL = _
                        " Select A.ID,A.����,A.����,A.����" & _
                        " From ���ű� A,������ҳ B" & _
                        " Where A.ID=B.��ǰ����ID And B.����ID=[9] And B.��ҳID=[10]" & _
                        " Union " & _
                        " Select ID,����,����,���� From ���ű� Where ID=[6]" & _
                        " Order by ����"
                End If
            Case 3 '3-����Ա���ڿ���
                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([7],[6]) Order by ����"
                lng����Ա����ID = Get����Ա����ID(int��Χ)
            Case 4 '4-ָ������
                If int��Ч = 1 Then bln�ϰల�� = Check�ϰల��(False)
                If Not bln�ϰల�� Then
                    strSQL = _
                        " Select Distinct A.ID,A.����,A.����,A.����" & _
                        " From ���ű� A,����ִ�п��� B,��������˵�� C" & _
                        " Where A.ID=B.ִ�п���ID And B.������ĿID=[5] And A.ID=C.����ID" & _
                        " And C.������� IN([2],3) And (B.������Դ is NULL Or B.������Դ=[2])" & _
                        " And (B.��������ID is NULL Or B.��������ID=[3])" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                        " Union Select ID,����,����,���� From ���ű� Where ID=[6]" & _
                        " Order by ����"
                Else
                    bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����" & _
                        " From ����ִ�п��� A,���Ű��� B,���ű� C,��������˵�� D" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.����ID=C.ID And B.����=[8]" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.��ʼʱ��,'HH24:MI:SS') and To_Char(B.��ֹʱ��,'HH24:MI:SS') " & _
                        " And C.ID=D.����ID And D.������� IN([2],3) And (A.������Դ is NULL Or A.������Դ=[2])" & _
                        " And (A.��������ID is NULL Or A.��������ID=[3]) And A.������ĿID=[5]" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " Union Select ID,����,����,���� From ���ű� Where ID=[6]" & _
                        " Order by ����"
                End If
                '�����������۲��˲����ָ����������ϰల�ţ�������󣬿������ң�������Դ
                If lng�������� = 1 Then
                    Set rsTmp = Get������Ŀ��¼(lng��Ŀid)
                    If "E" = rsTmp!��� & "" And InStr(",2,3,4,6,8,9,", "," & rsTmp!�������� & ",") > 0 Then
                        strSQL = " Select Distinct C.ID,C.����,C.����,C.����" & _
                            " From ����ִ�п��� A,���ű� C" & _
                            " Where A.ִ�п���ID+0=C.ID And A.������ĿID=[5]" & _
                            " Union Select ID,����,����,���� From ���ű� Where ID=[6]" & _
                            " Order by ����"
                    End If
                End If
            Case 6 '6-���������ڿ���
                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([11],[6]) Order by ����"
        End Select
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strҩ��, int��Χ, lng���˿���ID, lngҩƷID, lng��Ŀid, _
         lng��ǰִ��ID, lng����Ա����ID, bytDay, lng����ID, lng��ҳID, lng��������ID, gstr��Һ��������, Val(strĬ��ҩ��))
         
    '��ҺҩƷ���͵���Һ��������
    If blnָ����Һ�������� Then
        For i = 0 To UBound(Split(gstr��Һ��������, ","))
            If strҩ��IDs = "" Or InStr("," & strҩ��IDs & ",", "," & Split(gstr��Һ��������, ",")(i) & ",") > 0 Then
                strFilter = strFilter & " Or ID=" & Split(gstr��Һ��������, ",")(i)
            End If
        Next
        rsTmp.Filter = Mid(strFilter, 5)
        If rsTmp.RecordCount = 0 Then
            '�����Һ�������Ĳ��Ǵ洢�ⷿ�������ѡ�������洢�ⷿΪ��ҩҩ��(65111)
            rsTmp.Filter = 0
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
    End If
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        'ʹ��API���ټ���,��Ȼ�����е���
        AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, rsTmp!���� & "-" & rsTmp!����
        SetComboData objCbo.hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
        If InStr(",5,6,7,", str���) > 0 And Val(strĬ��ҩ��) <> 0 And Val(strĬ��ҩ��) = Val(rsTmp!ID & "") And objCbo.Text = "" Then
            Call Cbo.SetIndex(objCbo.hwnd, i - 1)
        End If
        If lng��ǰִ��ID = rsTmp!ID Then
            Call Cbo.SetIndex(objCbo.hwnd, i - 1)
        End If
        rsTmp.MoveNext
    Next
    
    '����ҩ��������ҽ������ѡ��
    If InStr(",4,5,6,7,", str���) = 0 And gblnָ��ҽ������������ִ�� And IIF(gblnѪ��ϵͳ = True And str��� = "K", 0, 1) = 1 Then
        AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, "[����...]"
        SetComboData objCbo.hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetƵ����Ϣ_����(ByVal str���� As String, strƵ�� As String, _
    intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String) As Boolean
'���ܣ�����Ƶ�ʵ������Ϣ
'������str����=Ƶ�ʱ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strƵ�� = ""
    intƵ�ʴ��� = 0
    intƵ�ʼ�� = 0
    str�����λ = ""
    
    strSQL = "Select ����,����,����,Ӣ������,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,���÷�Χ From ����Ƶ����Ŀ Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str����)
    If Not rsTmp.EOF Then
        strƵ�� = NVL(rsTmp!����)
        intƵ�ʴ��� = NVL(rsTmp!Ƶ�ʴ���, 0)
        intƵ�ʼ�� = NVL(rsTmp!Ƶ�ʼ��, 0)
        str�����λ = NVL(rsTmp!�����λ)
    End If
    GetƵ����Ϣ_���� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetƵ����Ϣ_����(ByVal strƵ�� As String, intƵ�ʴ��� As Integer, _
    intƵ�ʼ�� As Integer, str�����λ As String, str��Χ As String, Optional strƵ�ʱ��� As String) As Boolean
'���ܣ�����Ƶ�ʵ������Ϣ
'������strƵ��=Ƶ������
'      str��Χ=1-��ҽ,2-��ҽ,-1-һ����,-2-������
'���أ���������ȡ��ʱ������True�����򷵻�False
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
    
    intƵ�ʴ��� = 0
    intƵ�ʼ�� = 0
    str�����λ = ""
    
    strSQL = "Select Ƶ�ʴ���,Ƶ�ʼ��,�����λ,���� From ����Ƶ����Ŀ Where ����=[1] And Instr([2],','||���÷�Χ||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strƵ��, "," & str��Χ & ",")
    If Not rsTmp.EOF Then
        intƵ�ʴ��� = NVL(rsTmp!Ƶ�ʴ���, 0)
        intƵ�ʼ�� = NVL(rsTmp!Ƶ�ʼ��, 0)
        str�����λ = NVL(rsTmp!�����λ)
        strƵ�ʱ��� = "" & rsTmp!����
        GetƵ����Ϣ_���� = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetȱʡƵ��(ByVal lng��Ŀid As Long, ByVal int��Χ As Integer, strƵ�� As String, _
    intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String) As Boolean
'���ܣ�����������Ƶ����Ŀ��ȡһ����ΪȱʡƵ��
'������int��Χ=1-��ҽ,2-��ҽ,-1-һ����,-2-������
'���أ�ȱʡƵ����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnLoad As Boolean
    
    On Error GoTo errH
    
    strƵ�� = ""
    intƵ�ʴ��� = 0
    intƵ�ʼ�� = 0
    str�����λ = ""
    
    '��ȡ���Ƴ���Ƶ��
    blnLoad = True
    If lng��Ŀid <> 0 And int��Χ = 1 Then
        strSQL = "Select B.����,B.Ƶ�ʴ���,B.Ƶ�ʼ��,B.�����λ From �����÷����� A,����Ƶ����Ŀ B" & _
            " Where A.��ĿID=[1] And A.�÷�ID Is Null And A.Ƶ��=B.���� And B.���÷�Χ=[2]" & _
            " Order by A.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetȱʡƵ��", lng��Ŀid, int��Χ)
        If Not rsTmp.EOF Then blnLoad = False
    End If
    If blnLoad Then
        strSQL = "Select ����,Ƶ�ʴ���,Ƶ�ʼ��,�����λ From ����Ƶ����Ŀ Where ���÷�Χ=[1] Order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetȱʡƵ��", int��Χ)
    End If
    If Not rsTmp.EOF Then
        strƵ�� = NVL(rsTmp!����)
        intƵ�ʴ��� = NVL(rsTmp!Ƶ�ʴ���, 0)
        intƵ�ʼ�� = NVL(rsTmp!Ƶ�ʼ��, 0)
        str�����λ = NVL(rsTmp!�����λ)
    End If
    GetȱʡƵ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckƵ�ʿ���(ByVal lng��Ŀid As Long, ByVal int��Χ As Integer, ByVal strƵ�� As String) As Boolean
'���ܣ����ָ��Ƶ���Ƿ���������Ŀ�ĳ���Ƶ��
'������int��Χ=1-��ҽ,2-��ҽ,-1-һ����,-2-������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng��Ŀid = 0 Then CheckƵ�ʿ��� = True: Exit Function
    
    strSQL = "Select B.���� From �����÷����� A,����Ƶ����Ŀ B" & _
        " Where A.��ĿID=[1] And A.�÷�ID Is Null And A.Ƶ��=B.���� And B.���÷�Χ=[2]" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckƵ�ʿ���", lng��Ŀid, int��Χ)
    If rsTmp.EOF Then
        'û�����ó���Ƶ�ʣ���û����
        CheckƵ�ʿ��� = True
    ElseIf rsTmp.RecordCount = 1 Then
        'ֻ������һ������Ƶ�ʣ�ֻ��ȱʡ��Ҳû����
        CheckƵ�ʿ��� = True
    ElseIf strƵ�� <> "" Then
        rsTmp.Filter = "����='" & strƵ�� & "'"
        If Not rsTmp.EOF Then CheckƵ�ʿ��� = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getȱʡʱ��(int��Χ As Integer, strƵ�� As String, Optional lng��ҩ;��ID As Long) As String
'���ܣ���ȡָ��ִ��Ƶ��ȱʡ��ִ��ʱ�䷽��
'������int��Χ=1-��ҽ;2-��ҽ;-1-һ����;-2-������
'      lng��ҩ;��ID=�Ƿ���԰�ָ����ҩ;������ȡ,����ȡ��ȷ����ҩ;���ķ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(A.��ҩ;��ID,0) as �÷�,A.ʱ�䷽��" & _
        " From ����Ƶ��ʱ�� A,����Ƶ����Ŀ B" & _
        " Where A.ִ��Ƶ��=B.���� And B.���÷�Χ=[1]" & _
        " And (A.��ҩ;��ID is NULL Or A.��ҩ;��ID=[2]) And B.����=[3]" & _
        " Order by �÷� Desc,A.�������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int��Χ, lng��ҩ;��ID, strƵ��)
    If Not rsTmp.EOF Then Getȱʡʱ�� = NVL(rsTmp!ʱ�䷽��)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getʱ�䷽��(objCbo As Object, int��Χ As Integer, strƵ�� As String, Optional lng��ҩ;��ID As Long) As Boolean
'���ܣ���ȡָ��Ƶ�ʿ��õ�����Ƶ��ʱ�䷽����ָ����������,������ȱʡ��(�򱣳�ԭ��ֵ)
'������int��Χ=1-��ҽ;2-��ҽ;-1-һ����;-2-������;-3-��Ҫʱ;-5-��Ҫʱ
'      strƵ��=����Ƶ����Ŀ����
'      lng��ҩ;��ID=�Ƿ�ֻ��ȡָ����ҩ;����ʱ�䷽��,����Ϊ��ҩƷ��ִ��ʱ�䷽��
'      strִ��ʱ��=ȱʡִ��ʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    '����ͬ����(�����Ƿ�ָ����ҩ;��)��ִ��ʱ��Ӧ�ò���ͬ,������ظ�����
    strSQL = "Select A.�������,A.ʱ�䷽��" & _
        " From ����Ƶ��ʱ�� A,����Ƶ����Ŀ B" & _
        " Where A.ִ��Ƶ��=B.���� And B.����=[1]" & _
        " And (A.��ҩ;��ID is NULL Or A.��ҩ;��ID=[2])" & _
        " And B.���÷�Χ=[3]" & _
        " Order by A.�������"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strƵ��, lng��ҩ;��ID, int��Χ)
    strSQL = objCbo.Text: objCbo.Clear 'Clear�ᵼ��Text���
    For i = 1 To rsTmp.RecordCount
        objCbo.AddItem rsTmp!ʱ�䷽��
        rsTmp.MoveNext
    Next
    objCbo.Text = strSQL: objCbo.Tag = ""
    Getʱ�䷽�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ҽ��(ByVal lng���˿���ID As Long, ByVal bln��ʿվ As Boolean, strȱʡҽ�� As String, lngҽ��ID As Long, _
    Optional objCbo As Object, Optional ByVal int��Χ As Integer = 2, Optional ByVal blnOnlyDefault As Boolean) As Boolean
'���ܣ���ȡ���õĿ���ҽ����ָ������������
'������lng���˿���ID=�������ڿ���ID
'      bln��ʿվ=�Ƿ��ɻ�ʿ��ҽ����ҽ��
'      objCbo=Ҫ����ҽ���嵥��������
'      strȱʡҽ��=ȱʡ��λ��ҽ��,�������objCbo,�������ȶ�λ,�ٷ���ȱʡҽ����ҽ��ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
'      blnOnlyDefault=��ָ����ȱʡҽ��ʱ���Ƿ�ֻ��ȡ��ҽ������Ϣ����ʱӦ����"strȱʡҽ��"��"bln��ʿվ=True"��
'                     ���ͬʱ������objCbo�����򽫵�ǰȱʡҽ��׷�ӵ����б�ؼ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    If bln��ʿվ Then
        If blnOnlyDefault And strȱʡҽ�� <> "" Then
            strSQL = "Select ID,���,����,���� From ��Ա�� Where ����=[4]"
        Else
            '�������ڿ��ҵ�ҽ��
            strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIF(objCbo Is Nothing, ",B.����ID", "") & _
                " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
                " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
                " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And B.����ID=[1]" & _
                " Order by A.����"
            'ȫԺסԺ���ҵ�ҽ��
            strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN([2],3)"
            strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIF(objCbo Is Nothing, ",B.����ID", "") & _
                " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
                " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
                " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And B.����ID IN(" & strSQL & ")" & _
                " Order by A.����"
        End If
    Else 'ҽ����ҽ��ʱ,����Ϊֻ��Ϊҽ������
        strSQL = "Select ID,���,����,���� From ��Ա�� Where ID=[3]"
        'ҽ���´�ҽ��ʱ����ѡ������´��ҽ�������޸�ʱ������ҽ��Ӧ�ü��أ������ڱ༭������ѡ������ҽ���´��ҽ��ʱ���·�ѡ��еĿ���ҽ��Ϊ�ա�
        If strȱʡҽ�� <> "" And strȱʡҽ�� <> UserInfo.���� Then
            strSQL = strSQL & " union all Select ID,���,����,���� From ��Ա�� Where ����=[4]"
        End If
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get����ҽ��", lng���˿���ID, int��Χ, UserInfo.ID, strȱʡҽ��)
    If objCbo Is Nothing Then
        If Not rsTmp.EOF Then
            If Not bln��ʿվ Then
                lngҽ��ID = rsTmp!ID
                strȱʡҽ�� = rsTmp!����
            ElseIf bln��ʿվ Then
                If strȱʡҽ�� <> "" Then
                    'ȱʡҽ��(סԺҽʦ)����
                    rsTmp.Filter = "����='" & strȱʡҽ�� & "'"
                Else
                    '���˿��ҵ�ҽ������
                    rsTmp.Filter = "����ID=" & lng���˿���ID
                End If
                If rsTmp.EOF Then rsTmp.Filter = 0
                lngҽ��ID = rsTmp!ID
                strȱʡҽ�� = rsTmp!����
            End If
        End If
    Else
        If blnOnlyDefault Then
            '��ɾ��"����"
            i = Cbo.FindIndex(objCbo, -1)
            If i <> -1 Then objCbo.RemoveItem objCbo.ListCount - 1
            
            '��λ���������ѡ��
            If Not rsTmp.EOF Then
                i = Cbo.FindIndex(objCbo, rsTmp!ID)
                If i = -1 Then
                    objCbo.AddItem NVL(rsTmp!����) & "-" & Chr(13) & rsTmp!����
                    objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
                    Call Cbo.SetIndex(objCbo.hwnd, objCbo.NewIndex)
                Else
                    Call Cbo.SetIndex(objCbo.hwnd, i)
                End If
            End If
            
            '����"����"��ѡ��
            AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, "[����...]"
            SetComboData objCbo.hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
        Else
            'ȫ���¼���
            objCbo.Clear
            For i = 1 To rsTmp.RecordCount
                AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, NVL(rsTmp!����) & "-" & Chr(13) & rsTmp!����
                SetComboData objCbo.hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
                If rsTmp!���� = strȱʡҽ�� Then
                    Call Cbo.SetIndex(objCbo.hwnd, i - 1)
                End If
                rsTmp.MoveNext
            Next
        End If
    End If
    Get����ҽ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check����ִ��(ByVal lngִ�п���ID As Long) As Boolean
'���ܣ�ȷ��ָ����ִ�п����Ƿ񱾿�(ҽ������)
'������lngִ�п���ID=ҽ����ִ�п���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr����ID(1 To 4) As Long
    
    strSQL = "Select ����ID From ������Ա Where ��ԱID=[1] And ����ID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID, lngִ�п���ID)
    Check����ִ�� = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStock(ByVal lngҩƷID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal int��Χ As Integer = 2, _
        Optional ByVal strDepartments As String, Optional ByVal lng���� As Double) As Double
'���ܣ���ȡָ���ָⷿ��ҩƷ���������(�������סԺ��λ)
'������int��Χ=1-����,2-סԺ(ȱʡ),0-��ʾ���ۼ�
'      strDepartments����ִ�п����ַ���������������ѯ���
'      lng���� ���lng������Ϊ�գ����ѯ�Ƿ��п������������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    '��ȡҩƷ���(�����������ҩƷ),ҩ��������ҩƷ����Ч��
    If int��Χ = 0 Or int��Χ = 3 Then
        strSQL = _
            " Select Nvl(Sum(A.��������),0) as ���" & _
            " From ҩƷ��� A" & _
            " Where A.����=1" & _
            " And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
            " And A.ҩƷID=[1] And Instr([2],',' || a.�ⷿid || ',')>0 Group By A.�ⷿID"
    Else
        strTmp = IIF(int��Χ = 1, "����", "סԺ")
        strSQL = _
            " Select Nvl(Sum(A.��������),0)/Nvl(B." & strTmp & "��װ,1) as ���" & _
            " From ҩƷ��� A,ҩƷ��� B" & _
            " Where A.ҩƷID=B.ҩƷID(+) And A.����=1" & _
            " And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
            " And A.ҩƷID=[1] And Instr([2],',' || a.�ⷿid || ',')>0" & _
            " Group by Nvl(B." & strTmp & "��װ,1),A.�ⷿID"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҩƷID, IIF(strDepartments = "", "," & lng�ⷿID & ",", "," & strDepartments & ","))
    
    Do While Not rsTmp.EOF
    
        If strDepartments = "" Then
            GetStock = Format(rsTmp!���, "0.00000")
            Exit Function
        Else
            If Val(rsTmp!���) & "" > lng���� Then
                GetStock = Format(rsTmp!���, "0.00000")
                Exit Function
            End If
        End If
        rsTmp.MoveNext
    
    Loop
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetGroupCount(ByVal lng���ID As Long, ByVal int��Դ As Integer, Optional bln��Ч As Boolean = True, Optional ByVal lng�������� As Long) As Long
'���ܣ���ȡ�����Ŀ�е���Ŀ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str������� As String
    Dim strWhere As String
     
    On Error GoTo errH
    
    str������� = "," & int��Դ & ",3,"
    
    If lng�������� = 1 Then
        strWhere = "  And (Instr([2],','||B.�������||',')>0  or instr(',E2,E3,E4,E6,E8,E9,',','||b.���||b.��������||',')>0)"
    Else
        strWhere = "  And Instr([2],','||B.�������||',')>0 "
    End If
    
    strSQL = "Select Count(*) as NUM" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,�շ���ĿĿ¼ C" & _
        " Where A.������ĿID=B.ID(+) And A.�շ�ϸĿID=C.ID(+) And A.�������ID=[1]" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is NULL) " & strWhere & _
        " And (A.�շ�ϸĿID is NULL Or (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL) And Instr([2],','||c.�������||',')>0 )" & _
        IIF(bln��Ч And int��Դ <> 2, " And A.��Ч=1", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng���ID, str�������)
    If Not rsTmp.EOF Then GetGroupCount = NVL(rsTmp!Num, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetGroupNone(ByVal lng�䷽ID As Long, ByVal int��Դ As Integer) As String
'���ܣ���ȡָ���䷽����Ч�������ҩ��ʾ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    strSQL = "Select B.����" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,ҩƷ��� C,�շ���ĿĿ¼ D" & _
        " Where A.������ĿID=B.ID And B.ID=C.ҩ��ID And C.ҩƷID=D.ID And A.�������ID=[1]" & _
        " And (Not (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is NULL)" & _
        " Or Nvl(B.�������,0) Not IN([2],3))" & _
        " And (Not (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� is NULL)" & _
        " Or Nvl(D.�������,0) Not IN([2],3))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng�䷽ID, int��Դ)
    Do While Not rsTmp.EOF
        strMsg = strMsg & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    GetGroupNone = Mid(strMsg, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcDrugPrice(ByVal lngҩƷID As Long, lngҩ��ID As Long, ByVal dbl���� As Double, _
    Optional ByVal str�ѱ� As String, Optional ByVal blnNone�Ӱ�Ӽ� As Boolean, Optional ByVal int���� As Integer, Optional ByVal strҩƷ�۸�ȼ� As String, Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ��Ŀ�۸�ȼ� As String) As Double
'���ܣ�����ҩƷʵ��(��ȻҪ����ʵ��,ҩƷ��϶�Ϊ���)������ѱ�ʱ�������ʵ�ս��
'������dbl����=�ۼ�����,���ѱ����ʱ�������ʵ�ս��
'      str�ѱ�=�Ƿ񰴷ѱ������۵ļ۸�,��Ҫ��ֱ�Ӽ���ҩƷ�Ľ�������ʾ����ʱ��
'      gbln�Ӱ�Ӽ�=����ʱ���������,�����ط���ΪFalse
'      blnNone�Ӱ�Ӽ�=Ϊ��ʱ,"gbln�Ӱ�Ӽ�"��Ч
'      int����  0��סԺ��1������

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim dblʱ�� As Double
    
    If dbl���� = 0 Then Exit Function
    '����������סԺ��ͳһ����Zl_Fun_Getprice �����ӿ�
    On Error GoTo errH
    strSQL = "select Zl_Fun_Getprice([1],[2],[3],0,null) as ��� from dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CalcDrugPrice", lngҩƷID, lngҩ��ID, dbl����)
    strSQL = rsTmp!��� & ""
    If InStr(strSQL, "|") > 0 Then dblʱ�� = Val(Split(strSQL, "|")(0))
    '���зѱ����ʱ���ǽ�������������ʵ�ս��
    If str�ѱ� <> "" And dblʱ�� <> 0 Then
        dblʱ�� = Format(dblʱ�� * dbl����, gstrDec)
        strSQL = _
            " Select A.���ηѱ�,B.������ĿID" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
            " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CalcDrugPrice", lngҩƷID)
        If rsTmp.EOF Then Exit Function
        
        '���ݷѱ����¼���ʵ�ս��
        If Not (NVL(rsTmp!���ηѱ�, 0) = 1) Then
            dblʱ�� = ActualMoney(str�ѱ� & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsTmp!������ĿID, dblʱ��, lngҩƷID, lngҩ��ID, dbl����)
        End If
    End If
    CalcDrugPrice = dblʱ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcPrice(ByVal lng��Ŀid As Long, Optional ByVal str�ѱ� As String, _
    Optional ByVal dbl���� As Double, Optional ByVal blnNone�Ӱ�Ӽ� As Boolean, _
    Optional ByVal lngִ�п���ID As Long, Optional ByVal lng����ҽ��ID As Long, Optional ByVal strҩƷ�۸�ȼ� As String, Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ��Ŀ�۸�ȼ� As String) As Double
'���ܣ���ȡ�շ�ϸĿ�ĵ�ǰ�ۼۼ۸���,��۷���ȱʡ�۸�
'������str�ѱ�=�Ƿ񰴷ѱ������۵�ʵ�ս��
'      dbl����=���ѱ����ʱ,����Ҫ��������(���ۼ۵�λ),��ʱ�������ʵ�ս��
'      lngִ�п���ID=�������˷ѱ�ʱ��Ҫ,���ܰ��ɱ��Ӵ��ۼ���
'      gbln�Ӱ�Ӽ�=����ʱ���������,�����ط���ΪFalse
'      blnNone�Ӱ�Ӽ�=Ϊ��ʱ,"gbln�Ӱ�Ӽ�"��Ч
'      lng����ҽ��ID=�������ò�������ʾ�Ǹ������õ�ʱ������ҽ���ļ۸��ҽ���Ƽ��ж�ȡ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl���� As Double, dbl��� As Double
    
    On Error GoTo errH
    
    If lng����ҽ��ID <> 0 Then
        strSQL = "Select ���� From ����ҽ���Ƽ� Where ҽ��ID=[1] And �շ�ϸĿID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CalcPrice", lng����ҽ��ID, lng��Ŀid)
        If Not rsTmp.EOF Then dbl���� = NVL(rsTmp!����, 0)
    End If
    
    If str�ѱ� = "" Then
        strSQL = _
            " Select Sum(Decode(Nvl(A.�Ƿ���,0),1,Decode([2],0,B.ȱʡ�۸�,[2]),B.�ּ�)" & _
                IIF(gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ�, "*Decode(A.�Ӱ�Ӽ�,1,1+Nvl(B.�Ӱ�Ӽ���,0)/100,1)", "") & ") as ���" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
            " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
            GetPriceGradeSQL(strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ��Ŀ�۸�ȼ�, "A", "B", "3", "4", "5") & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��Ŀid, dbl����, strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ��Ŀ�۸�ȼ�)
        If Not rsTmp.EOF Then dbl��� = NVL(rsTmp!���, 0)
    Else
        '�������Խ�ActualMoney������SQLһ��д��������ѱ���ܱ�ɾ�����󲻳�����
        strSQL = _
            " Select A.���ηѱ�,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.������ĿID,Decode(Nvl(A.�Ƿ���,0),1,Decode([2],0,B.ȱʡ�۸�,[2]),B.�ּ�)" & _
                IIF(gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ�, "*Decode(A.�Ӱ�Ӽ�,1,1+Nvl(B.�Ӱ�Ӽ���,0)/100,1)", "") & " as ���" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
            " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
            GetPriceGradeSQL(strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ��Ŀ�۸�ȼ�, "A", "B", "3", "4", "5") & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��Ŀid, dbl����, strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ��Ŀ�۸�ȼ�)
        For i = 1 To rsTmp.RecordCount
            If NVL(rsTmp!���ηѱ�, 0) = 1 Then
                dbl��� = dbl��� + Format(dbl���� * Format(NVL(rsTmp!���, 0), "0.00000"), gstrDec)
            Else
                dbl��� = dbl��� + ActualMoney(str�ѱ� & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsTmp!������ĿID, Format(dbl���� * Format(NVL(rsTmp!���, 0), "0.00000"), gstrDec), _
                    lng��Ŀid, lngִ�п���ID, dbl����, IIF(gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0))
            End If
            rsTmp.MoveNext
        Next
    End If
    CalcPrice = dbl���
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function Check����ȼ��䶯����(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strDate As String) As Boolean
'���ܣ���鲡�������ǰ����Ч����ȼ�ҽ����������ֹͣ�Ļ���ȼ�ҽ������ֹͣʱ���ڵ�ǰ��ʼʱ��֮������ʾ��ֹ��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    Check����ȼ��䶯���� = False
    If Not IsDate(strDate) Then Exit Function
    
    strSQL = "Select 1" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
            "Where a.����id = [1] And a.��ҳid = [2] And a.������� = 'H' And a.������Ŀid = b.Id And b.�������� = '1' And a.ҽ��״̬ In (8, 9) And" & vbNewLine & _
            "      a.ִ����ֹʱ�� > [3] And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID, CDate(strDate))
    If rsTmp.RecordCount > 0 Then
        '�����ǰ������Ч�Ļ���ȼ�����ҽ�����ͺ���Զ�ֹͣ�ɵ�ҽ�����䶯��¼�Ŀ�ʼ�ͽ���ʱ��Ͳ��ύ��
        strSQL = "Select 1" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
            "Where a.����id = [1] And a.��ҳid = [2] And a.������� = 'H' And a.������Ŀid = b.Id And b.�������� = '1' And a.ҽ��״̬ In (3, 5, 7)"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
        If rsTmp.RecordCount = 0 Then
            Check����ȼ��䶯���� = True
            MsgBox "����ȼ�ҽ���Ŀ�ʼʱ�䲻��������ֹͣ�Ļ���ȼ�ҽ��֮ǰ��", vbInformation, gstrSysName
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ActualMoney(str�ѱ� As String, ByVal lng������ĿID As Long, ByVal curӦ�ս�� As Currency, _
    Optional ByVal lng�շ�ϸĿID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal dbl���� As Double, Optional ByVal dbl�Ӱ�Ӽ��� As Double) As Currency
'���ܣ������շ�ϸĿID��������ĿID(ǰ������),Ӧ�ս��,���ѱ����õķֶα������۹������ʵ�ս�
'       ���ҩƷ���ɱ����ձ����������ʵ�ս��
'������str�ѱ�=���˷ѱ�����ǰ���̬�ѱ�,�����ʽΪ"���˷ѱ�,��̬�ѱ�1,��̬�ѱ�2,..."
'      lng�ⷿID,dbl����,��ҩƷ����Ŀ���ɱ��ۼ��մ���ʱ����Ҫ����
'      dbl����=�����������ڵ��ۼ�����
'      dbl�Ӱ�Ӽ���=С������,�����Ӧ�ս���Ѱ��Ӱ�Ӽۼ���ʱ��Ҫ�����ڻ�ԭ������
'���أ������۹���ͱ��������ʵ�ս��,����Ƕ�̬�ѱ�,��"str�ѱ�"�������Żݷѱ�(ע�����δ���ۼ���,����ԭ������,Ҳ���ܷ��ص�һ��)
'˵����
'���ɱ��ۼ��ձ������۵����ּ��㷽��(ʵ����һ��)��
'1.���۽�� = �ɱ���� * (1 + ���ձ���)
'2.���۽�� = �ɱ��� * (1 + ���ձ���) * ��������
'��صļ��㹫ʽ��
'      �ɱ��� = ҩƷ�ۼ� * (1 - �����)
'      �ɱ���� = �ۼ۽�� * (1 - �����) = �ɱ��� * ��������
'      �п����ʱ:����� = ����� / �����,����:����� = ָ�������
'      ���ڷ���ҩƷ��Ӧÿ���������ηֱ����ɱ��ۺͳɱ����
'      ����ʱ�۷�����"ҩƷ�ۼ�=Nvl(���ۼ�,ʵ�ʽ��/ʵ������)"��������ʱ��ҩƷ��治��ʱ��������ۼ��㡣
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str�ѱ�, lng�շ�ϸĿID, lng������ĿID, curӦ�ս�� / (1 + dbl�Ӱ�Ӽ���), dbl����, lng�ⷿID)
        
    str�ѱ� = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl�Ӱ�Ӽ���), gstrDec)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Load��̬�ѱ�(lng����id As Long) As String
'���ܣ�Ȩ��ָ�����Ҷ�ȡ��ǰ��Ч�Ķ�̬�ѱ�(Ŀǰֻ��������)
'���أ��ѱ�="���˽�,��һ��"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select ����,����,���� From �ѱ�" & _
        " Where Nvl(����,1)=2 And Nvl(���ÿ���,1)=1 And Nvl(�������,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Union ALL" & _
        " Select Distinct A.����,A.����,A.����" & _
        " From �ѱ� A,�ѱ����ÿ��� B" & _
        " Where A.����=B.�ѱ� And B.����ID=[1]" & _
        " And Nvl(A.����,1)=2 And Nvl(A.���ÿ���,1)=2 And Nvl(A.�������,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(A.��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(A.��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Load��̬�ѱ�", lng����id)
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    Load��̬�ѱ� = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckUserEmpower(ByVal lng������ĿID As Long) As Boolean
'���ܣ�������Ա�Ƿ����������Ŀ�Ŀ���Ȩ
    Dim strSQL As String, rsTmp As Recordset
    
    strSQL = "Select Count(*) as Ȩ�� From ��Ա����Ȩ�� Where ��Աid = [1] And ������Ŀid = [2] And ��¼���� = 1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", UserInfo.ID, lng������ĿID)
    CheckUserEmpower = rsTmp!Ȩ�� > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDocEmpower(ByVal lng������ĿID As Long, ByVal strAppend As String) As Boolean
'���ܣ�������Ա�Ƿ����������Ŀ��ִ��Ȩ
'������strAppend=��ǰ���븽�����д�����,��ʽΪ"��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
    Dim strSQL As String, rsTmp As Recordset
    Dim arrItem As Variant, arrSub As Variant
    Dim strItem As String, i As Integer
    Dim lngID As Long
    Dim strDoc As String
    
    On Error GoTo errH
    strSQL = "select A.ID from ����������Ŀ A,������������ B where a.����id=b.id and b.����='06' and A.������='����ҽ��'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDocEmpower")
    If rsTmp.RecordCount > 0 Then
        lngID = rsTmp!ID
        arrItem = Split(strAppend, "<Split1>")
        For i = 0 To UBound(arrItem)
            arrSub = Split(arrItem(i), "<Split2>")
            If Val(arrSub(2)) = lngID Then
                If Trim(arrSub(3)) <> "" Then
                    strDoc = Trim(arrSub(3))
                End If
                Exit For
            End If
        Next
    End If
    If strDoc = "" Then strDoc = UserInfo.����
    strSQL = "Select Count(*) as Ȩ�� From ��Ա����Ȩ�� A,��Ա�� B Where A.��Աid = B.ID And B.����=[1] And A.������Ŀid = [2] And A.��¼���� = 2"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", strDoc, lng������ĿID)
    CheckDocEmpower = Val(rsTmp!Ȩ�� & "") > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�ٴ�����(ByVal int��Χ As Integer, Optional ByVal lng���˿���ID As Long, _
    Optional lngȱʡ����ID As Long, Optional objCbo As Object, Optional ByVal blnBed As Boolean, Optional ByVal blnNode As Boolean = True, _
    Optional bln������Ա���� As Boolean, Optional ByVal int��� As Integer) As Boolean
'���ܣ������ٴ������嵥��ȱʡ�ٴ�����
'������int��Χ=1-����,2-סԺ,3-�����סԺ
'      lng���˿���ID=���˵�ǰ�Ŀ���,����Ҫ�ſ��ÿ���
'      objCbo=Ҫ��������嵥��������,����ʱ,����ȱʡ����
'      lngȱʡ����ID=��objCboʱ,Ϊȱʡ��λ�Ŀ��ң�����ΪҪ���ص�ȱʡ����
'      blnBed=�Ƿ�ֻȡ�д�λ�Ŀ���
'      blnNode=�Ƿ�����Ϊ��ǰվ��Ŀ��ң�ת��ҽ������ʱ������
'      bln������Ա����=��Ժ������ҽ����ȱʡ����Ϊҽ��������סԺ����
'      int��� ҽ�������1����ʾ��ǰҽ��Ϊ����תסԺҽ��Z2��Ŀ����1Ϊ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String, blnHave As Boolean
        
    On Error GoTo errH
    
    If int��Χ = 1 Then
        strTmp = "1,3"
    ElseIf int��Χ = 2 Then
        strTmp = "2,3"
    ElseIf int��Χ = 3 Then
        strTmp = "1,2,3"
    End If
    
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,A.���� " & _
        " From ���ű� A,��������˵�� B " & IIF(bln������Ա����, ",������Ա C ", "") & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        IIF(blnNode, " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
        " And A.ID=B.����ID And Instr([1],','||B.�������||',')>0 And B.��������='�ٴ�'" & _
        IIF(lng���˿���ID <> 0, " And A.ID<>[2]", "") & _
        IIF(bln������Ա����, " And A.id = C.����id And c.��Աid =[3]", "") & _
        IIF(blnBed, " And (Exists(Select ����ID From ��λ״����¼ Where ����ID=A.ID) Or Exists(Select ����ID From �������Ҷ�Ӧ Where ����ID=A.ID))", "") & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", "," & strTmp & ",", lng���˿���ID, UserInfo.ID)
    
    If Not objCbo Is Nothing Then
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            objCbo.AddItem rsTmp!���� & "-" & rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
            If rsTmp!ID = lngȱʡ����ID Then
                Call Cbo.SetIndex(objCbo.hwnd, objCbo.NewIndex)
                blnHave = True
            End If
            rsTmp.MoveNext
        Next
        
        If int��� = 1 Then
            If lngȱʡ����ID <> 0 And Not blnHave Then
                strSQL = "Select A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,��������˵�� B Where A.ID=B.����ID And B.������� IN(2,3) and a.id=[1]" & _
                    IIF(blnNode, " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
                    " And B.��������='�ٴ�' And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngȱʡ����ID)
                If Not rsTmp.EOF Then
                    objCbo.AddItem rsTmp!���� & "-" & rsTmp!����
                    objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
                    Call Cbo.SetIndex(objCbo.hwnd, objCbo.NewIndex)
                End If
            End If
        End If
        
        If bln������Ա���� Then
            AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, "[����...]"
            SetComboData objCbo.hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
        End If
    ElseIf Not rsTmp.EOF Then
        lngȱʡ����ID = rsTmp!ID
    End If
    Get�ٴ����� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get������ĿID(ByVal lngҽ��ID As Long, bln��ҩ���� As Boolean) As Long
'���ܣ���ȡָ��ҽ����������ĿID
'������lngҽ��ID=���IDΪNULL��ҽ��ID(��ҩ,�����Ŀ,��Ҫ����,��ҩ�÷�,������ҽ��)
'      bln��ҩ����=ҽ��ID�Ƿ�Ϊ��ҩ�÷���ɼ�������ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bln��ҩ���� Then
        strSQL = "Select ���,������ĿID From ����ҽ����¼ Where ������� IN('7','C') And ���ID=[1]"
    Else
        strSQL = "Select ���,������ĿID From ����ҽ����¼ Where ID=[1]"
    End If
    strSQL = strSQL & " Union ALL " & Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    strSQL = strSQL & " Order by ���"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
    If Not rsTmp.EOF Then Get������ĿID = NVL(rsTmp!������ĿID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get������Ŀ��¼(ByVal lngID As Long, Optional ByVal strIDs As String) As ADODB.Recordset
'���ܣ���ȡָ��������ĿID�ļ�¼
'������
    Dim strSQL As String
    
    strSQL = "Select /*+ rule*/ �������,վ��,���,����ID,ID,����,����,�걾��λ,���㵥λ,���㷽ʽ,ִ��Ƶ��,�����Ա�,����Ӧ��,�����Ŀ,��������,ִ�а���,ִ�п���,�������,�Ƽ�����,�ο�Ŀ¼ID,��ԱID,����ʱ��,����ʱ��,¼������,�Թܱ���,ִ�з���,ִ�б��" & _
            " From ������ĿĿ¼ Where ID"
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = strSQL & " IN (Select Column_Value From Table(f_Num2list([1])))"
        Set Get������Ŀ��¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strIDs)
    Else
        strSQL = strSQL & " = [1]"
        Set Get������Ŀ��¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�շ���Ŀ��¼(ByVal lngID As Long, Optional ByVal strIDs As String) As ADODB.Recordset
'���ܣ���ȡָ���շ���ĿID�ļ�¼
'������
    Dim strSQL As String
    
    strSQL = "Select /*+ rule*/ ���,����ID,ID,����,����,���,����,���㵥λ,˵��,��Ŀ����,��������,�������,���ηѱ�,�Ƿ���,�Ӱ�Ӽ�,����ժҪ,����ȷ��,ִ�п���,��ʶ����,��ʶ����,��ѡ��,����޼�,����޼�,����ʱ��,����ʱ��,¼������,���㷽ʽ,վ��,����ԭ��,ͣ��ԭ��,������Ŀ" & _
            " From �շ���ĿĿ¼ Where ID"
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = strSQL & " IN(Select Column_Value From Table(f_Num2list([1])))"
        Set Get�շ���Ŀ��¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strIDs)
    Else
        strSQL = strSQL & " = [1]"
        Set Get�շ���Ŀ��¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDrugInfo(lngҩ��ID As Long, lngҩƷID As Long, lngҩ��ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal blnͣ�� As Boolean = True) As ADODB.Recordset
'���ܣ���ȡָ��ҩƷ�����Ϣ
'������int��Χ=1-����,2-סԺ(ȱʡ)
'      blnͣ��=�Ƿ��ſ���ͣ��ҩƷ,���ڳ���ҩƷ���ʹ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strTmp = IIF(int��Χ = 1, "����", "סԺ")
    
    strSQL = _
        " Select ҩƷID,Sum(Nvl(��������,0)) as ��� From ҩƷ���" & _
        " Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))" & _
        " And ���� = 1 And �ⷿID=[1]" & IIF(lngҩƷID <> 0, " And ҩƷID=[2]", "") & _
        " Group by ҩƷID Having Sum(Nvl(��������,0))<>0"
    strSQL = "Select A.ҩ��ID,A.ҩƷID,A.����ϵ��,A." & strTmp & "��װ,A." & strTmp & "��λ,A." & strTmp & "�ɷ���� As �ɷ����,A.��̬����," & _
        " A.ҩ������,B.�Ƿ���,C.���/A." & strTmp & "��װ as ���,B.����,Nvl(D.����,B.����) as ����,B.���,B.����,B.����ʱ��,B.�������,a.�Ƿ��ҩ" & _
        " From ҩƷ��� A,�շ���ĿĿ¼ B,(" & strSQL & ") C,�շ���Ŀ���� D" & _
        " Where A.ҩƷID=B.ID And A.ҩƷID=C.ҩƷID(+)" & _
        " And B.ID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=[5]" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
        IIF(blnͣ��, " And B.������� IN([3],3) And (B.����ʱ�� is NULL Or B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))", "") & _
        " And A.ҩ��ID=[4]" & IIF(lngҩƷID <> 0, " And A.ҩƷID=[2]", "") & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҩ��ID, lngҩƷID, int��Χ, lngҩ��ID, IIF(gbytҩƷ������ʾ = 0, 1, 3))
    Set GetDrugInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'���ܣ��ж��Ƿ����ָ�������������͵�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���ID From ҩƷ�������� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = NVL(rsTmp!���ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetClinicBillID(ByVal lng��Ŀid As Long, ByVal int���� As Integer) As Long
'���ܣ���ȡ������Ŀ��Ӧ�����Ƶ���(���ܸ���,�������ɷ���NO)
'������int����=1-����,2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select �����ļ�ID From ��������Ӧ�� Where ������ĿID=[1] And Ӧ�ó���=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid, int����)
    If Not rsTmp.EOF Then GetClinicBillID = NVL(rsTmp!�����ļ�ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptIsWoman(ByVal lng����id As Long, Optional ByVal str����IDs As String) As Boolean
'���ܣ��ж�ָ�������Ƿ����
'������str����IDs-����������ID���ж��Ƿ������в���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If str����IDs = "" Then
        strSQL = "Select ��������,����ID,������� From ��������˵�� Where ��������='����' And ����ID=[1]"
    Else
        strSQL = "Select /*+ Rule*/ ��������,����ID,������� From ��������˵�� Where ��������='����' And ����ID In (Select Column_Value From Table(Cast(f_Str2List([2]) As zlTools.t_StrList)))"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����id, str����IDs)
    DeptIsWoman = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����IDs(Optional ByVal lng����ID As Long) As String
'���ܣ����ݲ���ID��ò�����Ӧ�Ŀ���ID
'�������Ƿ�ȡ���������µĿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long

    strSQL = _
            " Select B.����ID From �������Ҷ�Ӧ B" & _
            " Where B.����ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get����IDs", lng����ID)
    Get����IDs = lng����ID
    For i = 1 To rsTmp.RecordCount
        If InStr("," & Get����IDs & ",", "," & rsTmp!����ID & ",") = 0 Then
            Get����IDs = Get����IDs & "," & rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer) As String
'���ܣ���鲡����ҽ�������Ƿ���δִ�����(δִ�л�����ִ��)����Ŀ
'���أ�ҽ����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(2,[1],[2],[3]) as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitExe", lng����ID, lng��ҳID, intӤ��)
    
    If Not rsTmp.EOF Then
        ExistWaitExe = NVL(rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitDrug(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer) As String
'���ܣ���鲡����ҩ���Ƿ���δ��ҩ��ҩƷ������
'���أ�ҩ���ͷ��ϲ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(1,[1],[2],[3]) as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitDrug", lng����ID, lng��ҳID, intӤ��)
    
    If Not rsTmp.EOF Then
        ExistWaitDrug = NVL(rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicePause(ByVal lngҽ��ID As Long, Optional ByVal lng��ID As Long) As String
'���ܣ���ȡָ��ҽ������ͣʱ��μ�¼
'���أ�"��ͣʱ��,��ʼʱ��;...."
'ע�⣺�����������˾�̬����л���ʹ��ʱע������һ�λ��淽ʽ Call GetAdvicePause(0)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    Static strLastPause As String
    Static lng���ID As Long
    
    On Error GoTo errH
    
    If lng���ID = lng��ID And lng��ID <> 0 Then GetAdvicePause = strLastPause: Exit Function
    If lngҽ��ID <> 0 Then
        strSQL = "Select ��������,����ʱ�� From ����ҽ��״̬" & _
            " Where �������� IN(6,7) And ҽ��ID=[1]" & _
            " Order by ����ʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
        For i = 1 To rsTmp.RecordCount
            If rsTmp!�������� = 6 Then
                strTmp = strTmp & ";" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & ","
            ElseIf rsTmp!�������� = 7 Then
                '���õ���һ�벻����ͣ�ķ�Χ֮��
                strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
            End If
            rsTmp.MoveNext
        Next
    End If
    lng���ID = lng��ID
    strLastPause = Mid(strTmp, 2)
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlClinicCodeRepeat(str���� As String, Optional lng��Ŀid As Long) As Boolean
'���ܣ����������Ŀ������Ƿ������б����ظ����ظ��������ʾ
'��Σ�str����-����ı��룻lng��ĿID-�Լ���ID�ţ����޸�ʱ����Ҫ��������������ж�
'���Σ��ظ�����True��������Flase
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select K.����||' ['||I.����||']'||I.���� as ����" & _
        " From ������ĿĿ¼ I,������Ŀ��� K" & _
        " Where I.���=K.���� And I.����=[1] And I.ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str����, lng��Ŀid)
    If Not rsTmp.EOF Then
        MsgBox "����Ŀ�����롰" & rsTmp!���� & "���ı����ظ���", vbInformation, gstrSysName
        zlClinicCodeRepeat = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillMax���(ByVal strNO As String, ByVal int��¼���� As Integer, str�Ǽ�ʱ�� As String, int������Դ As Integer) As Integer
'���ܣ���ȡָ�����ݵ�ǰ��������+1
'������str�Ǽ�ʱ��=���ҽ��ֻ�����˲���������ʱ����Ҫ�����ɵ��շѻ��۵�(NO��ͬ)��ʱ���������ɵ�һ�¡�
'      int������Դ:1-���2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIF(int��¼���� = 1 Or (int��¼���� = 2 And int������Դ = 1), "������ü�¼", "סԺ���ü�¼")
    On Error GoTo errH
    str�Ǽ�ʱ�� = ""
    strSQL = "Select Max(���) as ���,Max(�Ǽ�ʱ��) as ʱ�� From " & strTab & " Where NO=[1] And ��¼����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strNO, int��¼����)
    If Not rsTmp.EOF Then
        GetBillMax��� = NVL(rsTmp!���, 0) + 1
        If Not IsNull(rsTmp!ʱ��) Then
            str�Ǽ�ʱ�� = Format(rsTmp!ʱ��, "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        GetBillMax��� = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptExist(ByVal str�������� As String, ByVal int������� As Integer) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ID From ��������˵�� Where ��������=[1] And ������� IN([2],3) And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str��������, int�������)
    DeptExist = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedBySend(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, byt��Դ As Byte) As Boolean
'���ܣ����ĳ�η��͵�ҽ���еķ����Ƿ��Ѿ�ִ��������ת��
'������lng���ͺ�=��Ϊ����ҽ��ֻ��һ�η���,���Բ�����,byt��Դ:1-���2-סԺ
'˵����1.��ҽ��δת��������£�ִ�л��˻����ϲ���ʱ�����������ת���ķ��ã����ֹ
'      2.����סԺ�����ж�η��͵������ֻ�жϵ�ǰҪ���˵����ҽ�����ͷ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    strTab = IIF(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
    
    strSQL = "Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1]"
    strSQL = "Select B.NO From ����ҽ������ A,H" & strTab & " B" & _
        " Where A.��¼����=B.��¼���� And A.NO=B.NO" & _
        IIF(lng���ͺ� <> 0, " And A.���ͺ�+0=[2]", "") & _
        " And A.ҽ��ID IN(" & strSQL & ")" & _
        " Group by B.NO"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID, lng���ͺ�)
    If Not rsTmp.EOF Then MovedBySend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getҽ�Ƹ�����(ByVal str���� As String) As String
'���ܣ�����ҽ�Ƹ��ʽ���ƻ�ȡҽ�Ƹ�������  1-ҽ��  2-����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If str���� = "" Then Exit Function
    
    strSQL = "Select Decode(�Ƿ�ҽ��,1,1,Decode(�Ƿ񹫷�,1,2,0)) ���� From ҽ�Ƹ��ʽ Where ����=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str����)
    If Not rsTmp.EOF Then Getҽ�Ƹ����� = NVL(rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPatiDataMoved(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��ж�ָ�����˵������Ƿ���ת��
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select ����ת�� From ������ҳ Where ����ID = [1] And ��ҳID = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ת��", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        CheckPatiDataMoved = Val("" & rsTmp!����ת��) = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckOneDuty(ByVal strҽ�� As String, ByVal strְ�� As String, ByVal strҽ�� As String, ByVal blnҽ�� As Boolean) As String
'���ܣ���鵱ǰָ��ҩƷ����ְ���Ƿ����
'������strҽ��=ҩƷҽ����ʾ����
'      strְ��=ҩƷ����ְ��
'      strҽ��=����ҽ��
'      blnҽ��=�Ƿ񹫷ѻ�ҽ������
'      grsDuty=��¼ҽ��ְ�񻺴�
'���أ�ְ���������ʾ��Ϣ����������򷵻ؿա�
    Const STR_ְ�� = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim intְ��A As Integer, intְ��B As Integer
    
    If Len(strְ��) <> 2 Or strҽ�� = "" Then Exit Function
    
    'ȡҩƷ����ְ��
    If blnҽ�� Then
        intְ��B = Val(Right(strְ��, 1))
    Else
        intְ��B = Val(Left(strְ��, 1))
    End If
    If intְ��B = 0 Then Exit Function '������
    
    'ȡҽ��ְ��
    If grsDuty Is Nothing Then
        Set grsDuty = New ADODB.Recordset
        grsDuty.Fields.Append "ҽ��", adVarChar, 50
        grsDuty.Fields.Append "ְ��", adInteger
        grsDuty.CursorLocation = adUseClient
        grsDuty.LockType = adLockOptimistic
        grsDuty.CursorType = adOpenStatic
        grsDuty.Open
    End If
    grsDuty.Filter = "ҽ��='" & strҽ�� & "'"
    If grsDuty.EOF Then
        On Error GoTo errH
        strSQL = "Select ����,Nvl(Ƹ�μ���ְ��,0) as ְ�� From ��Ա�� Where ����=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strҽ��)
        On Error GoTo 0
        If Not rsTmp.EOF Then
            grsDuty.AddNew
            grsDuty!ҽ�� = rsTmp!����
            grsDuty!ְ�� = rsTmp!ְ��
            grsDuty.Update
        End If
    End If
    If Not grsDuty.EOF Then
        intְ��A = grsDuty!ְ��
    End If
        
    '���ְ��Ҫ��
    If intְ��A = 0 Then
        'ҽ��δ����ְ������
        strMsg = """" & strҽ�� & """Ҫ���" & IIF(blnҽ��, "ҽ��", "����") & "ְ�����㣺" & vbCrLf & vbCrLf & IIF(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """δ����ְ��"
    ElseIf intְ��B < intְ��A Then
        '��ֵԽСְ��Խ��
        strMsg = """" & strҽ�� & """Ҫ���" & IIF(blnҽ��, "ҽ��", "����") & "ְ�����㣺" & vbCrLf & vbCrLf & IIF(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """��ְ��Ϊ""" & Split(STR_ְ��, ",")(intְ��A - 1) & """��"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistsDiagNoses(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str���� As String) As Boolean
'���ܣ���鲡��ָ��������Ƿ����
'������lng����ID=���ﲡ��Ϊ�Һ�ID,סԺ����Ϊ��ҳID
'      str����=�������,��"1,11"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��¼��Դ,����ID,���ID,�������,�Ƿ����� From ������ϼ�¼" & _
        " Where ����ID=[1] And Nvl(��ҳID,0)=[2] And Instr([3],','||�������||',')>0 And ȡ��ʱ�� Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng����ID, "," & str���� & ",")
    ExistsDiagNoses = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get������ϼ�¼(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str���� As String) As ADODB.Recordset
'���ܣ���ȡ������ϼ�¼
'������lng����ID�����ﲡ�˴��Һ�ID��סԺ���˴���ҳID
'       �������-1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
'        11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���
'       ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����(����ҽ��վ,���ժҪ);
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select a.����id, a.���id, a.�������, a.��ϴ���, Nvl(b.����, c.����) As ����, Nvl(b.����, c.����) ����" & vbNewLine & _
             "From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & vbNewLine & _
             "Where a.����id = [1] And a.��ҳid = [2] And ȡ��ʱ�� Is Null And ��¼��Դ IN (1, 3)  And NVL(A.�������,1) = 1 And Instr(',' ||[3]|| ',', ',' || ������� || ',') > 0 And a.����id = b.Id(+) And" & vbNewLine & _
             "      a.���id = c.Id(+)" & vbNewLine & _
             "Order By ��¼��Դ, �������, ��ϴ���"
    Set Get������ϼ�¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng����ID, str����)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���˹�����¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˹�����¼
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng��ҳID = 0 Then
        strSQL = "Select Distinct ҩ��ID,ҩ����,����Դ���� From ���˹�����¼ Where ����ID=[1] And ���=1 And Nvl(����ʱ��,��¼ʱ��)>Trunc(Sysdate-[3])"
    Else
        strSQL = "Select Distinct ҩ��ID,ҩ����,����Դ���� From ���˹�����¼ Where ����ID=[1] And ��ҳID=[2] And ���=1"
    End If
    Set Get���˹�����¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID, gint�����Ǽ���Ч����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadAdviceSignSource(ByVal int�������� As Integer, _
    ByVal lng����ID As Long, ByVal varTime As Variant, strIDs As String, _
    ByVal lngǩ��id As Long, ByVal blnMoved As Boolean, strSource As String, _
    Optional ByVal strǰ��IDs As String, Optional ByVal colSomeTime As Collection, _
    Optional ByRef ColIDs As Collection, Optional ByRef ColSource As Collection) As Integer
'���ܣ���ȡ�������ڵ���ǩ��/��֤��ҽ��Դ������
'������
'  int��������=Ҫǩ��/��֤ǩ����ҽ��״̬
'  ǩ��ʱ���룺
'    lng����ID
'    varTime=���˹Һŵ��Ż���ҳID
'    strIDs=ָ��Ҫǩ����ҽ��ID����(��ID)
'    strǰ��IDs=�¿�ҽ��Ҫǩ����ҽ����Դ(�Ƿ�ҽ��)
'    colSomeTime=ĳҽ����ʱ�����ݣ���ֹͣҽ��ǩ��ʱ���������ҽ��ִ����ֹʱ������ݣ�У��ʱ����У��ʱ������
'  ��֤ǩ��ʱ��
'    lngǩ��ID=ǩ����¼��ID
'    blnMoved=�Ƿ�ҽ��������ת��
'���أ�ǩ��/��֤ǩ����Դ�����ɹ���
'      strIDs=ǩ��/��֤ǩ����ҽ��ID����(ÿ����ϸID)
'      strSource=ǩ��/��֤ǩ����ҽ��Դ��
'      ColIDs=���ÿ��ҽ��ǩ��һ�Σ��򷵻ذ�ÿ��ҽ������ҽ��ID����
'      ColstrSource=���ÿ��ҽ��ǩ��һ�Σ��򷵻ذ�ÿ��ҽ������ҽ��Դ�ļ���
    Dim rsTmp As New ADODB.Recordset
    Dim str��IDs As String, strSQL As String, i As Long
    Dim arrField As Variant, strField As String
    Dim strLine As String, intRule As Integer
    Dim blnÿ��ҽ������ǩ�� As Boolean
    Dim strID As String, str���ID As String
    Dim strSourceTmp As String, strIDsTmp As String
    Dim strWhere As String
    
    On Error GoTo errH
    
    str��IDs = strIDs
    strSource = "": strIDs = ""
    intRule = 1 '�������µ�ҽ��ǩ��Դ�����ɹ�����
    Set ColIDs = New Collection
    Set ColSource = New Collection
    blnÿ��ҽ������ǩ�� = Val(zlDatabase.GetPara(239, glngSys) & "") <> 0
    
    If lngǩ��id = 0 Then
        'ǩ��ʱ
        If int�������� = 1 Then
            If gblnѪ��ϵͳ Then
                strWhere = " And a.������Ŀid = c.Id(+) And" & vbNewLine & _
                    "      (a.������Ŀid Is Null Or" & vbNewLine & _
                    "      Not (Nvl(a.���״̬, 0) <> 2 And (a.������� = 'K' And Exists (Select 1" & vbNewLine & _
                    "                                                              From ����ҽ����¼ X, ������ĿĿ¼ Y" & vbNewLine & _
                    "                                                              Where x.���id = a.Id And x.������Ŀid = y.Id And x.������� = 'E' And" & vbNewLine & _
                    "                                                                    y.�������� = '8' And Nvl(y.ִ�з���, 0) = 0) Or" & vbNewLine & _
                    "        a.������� = 'E' And c.�������� = '8' And Nvl(c.ִ�з���, 0) = 0)))"
            End If
            '���¿���ҽ������ǩ�������ξ���/סԺ��ǰҽ�����´��δǩ��ҽ���������Ѫҽ���¿�ǩ��ʱ���������Ѫ��ϵͳ��ֻ�ܶ����ͨ���¿���Ѫҽ��ǩ�������״̬��2����
            strSQL = _
                " Select A.* From ����ҽ����¼ A,����ҽ��״̬ B" & IIF(gblnѪ��ϵͳ, ",������ĿĿ¼ C", "") & " Where A.ID=B.ҽ��ID And B.ǩ��ID is Null And B.��������=1" & _
                " And A.ҽ��״̬=1 And Nvl(A.ǰ��ID,0) in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([5]) As zlTools.t_Numlist)) X)" & _
                " And Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))=[3]" & _
                " And Exists(Select M.���� From ��Ա�� M,ִҵ��� N" & _
                " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
                " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ'))" & strWhere & _
                IIF(TypeName(varTime) = "String", " And A.����ID+0=[1] And A.�Һŵ�=[2]", " And A.����ID=[1] And A.��ҳID=[2]") & _
                IIF(str��IDs <> "", " And Nvl(A.���ID,A.ID) IN (Select /*+cardinality(x,10)*/ x.Column_Value From Table(f_Num2list([4])) X)", "") & _
                " Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, varTime, UserInfo.����, str��IDs, IIF("" = strǰ��IDs, "0", strǰ��IDs))
        Else
            '��Ҫ���ϡ�ֹͣ��У�Ե�ҽ������ǩ�����¿�ʱǩ������ָ��ҽ������һ���ǵ�ǰҽ���´�
            strSQL = _
                " Select A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID is Not Null And B.��������=1" & _
                IIF(TypeName(varTime) = "String", " And A.����ID+0=[1] And A.�Һŵ�=[2]", " And A.����ID=[1] And A.��ҳID=[2]") & _
                IIF(str��IDs <> "", " And Nvl(A.���ID,A.ID) IN(Select /*+cardinality(x,10)*/ x.Column_Value From Table(f_Num2list([3])) X)", "") & _
                " Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, varTime, str��IDs)
        End If
    Else
        '��֤ǩ��ʱ:�ȶ�ȡǩ��ʱ��Դ�����ɹ���
        strSQL = "Select ǩ������ From ҽ��ǩ����¼ Where ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "ҽ��ǩ����¼", "Hҽ��ǩ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngǩ��id)
        If Not rsTmp.EOF Then intRule = NVL(rsTmp!ǩ������, 1)
        '--
        strSQL = _
            " Select A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID=[1] Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
        If blnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngǩ��id)
    End If
    
    'ҽ��Դ�ĵĲ�ͬ���ɹ���
    If intRule = 1 Then
        If int�������� = 3 Then
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,У��ʱ��,ִ������,������־,����ҽ��,����ʱ��"
        ElseIf int�������� = 8 Then
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,ִ����ֹʱ��,ִ������,������־,����ҽ��,����ʱ��"
        Else
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,ִ������,������־,����ҽ��,����ʱ��"
        End If
    End If
    arrField = Split(strField, ",")
        
    '����ҽ��ǩ��Դ��
    Do While Not rsTmp.EOF
        strLine = ""
        For i = 0 To UBound(arrField)
            If lngǩ��id = 0 And int�������� = 3 And arrField(i) = "У��ʱ��" Then
                'У��ҽ��ǩ��ʱ,��У��ʱ�����⴦����������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                strLine = strLine & vbTab & colSomeTime("_" & NVL(rsTmp!���ID, rsTmp!ID))
            ElseIf lngǩ��id = 0 And int�������� = 8 And arrField(i) = "ִ����ֹʱ��" Then
                'ֹͣҽ��ǩ��ʱ,����ֹʱ�����⴦����������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                strLine = strLine & vbTab & colSomeTime("_" & NVL(rsTmp!���ID, rsTmp!ID))
            Else
                If IsNull(rsTmp.Fields(arrField(i)).value) Then
                    strLine = strLine & vbTab & ""
                Else
                    If Rec.IsType(rsTmp.Fields(arrField(i)).Type, adDBTimeStamp) Then
                        strLine = strLine & vbTab & Format(rsTmp.Fields(arrField(i)).value, "yyyy-MM-dd HH:mm:ss")
                    Else
                        strLine = strLine & vbTab & rsTmp.Fields(arrField(i)).value
                    End If
                End If
            End If
        Next
        strSource = strSource & vbCrLf & Mid(strLine, 2)
        strIDs = strIDs & "," & rsTmp!ID
        strSourceTmp = strSourceTmp & vbCrLf & Mid(strLine, 2)
        strIDsTmp = strIDsTmp & "," & rsTmp!ID
        strID = rsTmp!ID: str���ID = rsTmp!���ID & ""
        rsTmp.MoveNext
        If blnÿ��ҽ������ǩ�� Then
            'ÿ��ҽ������ǩ���򷵻ؼ���
            If rsTmp.EOF = False Then
                If rsTmp!ID & "" <> str���ID And (rsTmp!���ID & "" <> str���ID Or (str���ID = "" And rsTmp!���ID & "" = "")) And rsTmp!���ID & "" <> strID Then
                    ColIDs.Add Mid(strIDsTmp, 2)
                    ColSource.Add Mid(strSourceTmp, 3)
                    strIDsTmp = "": strSourceTmp = ""
                End If
            ElseIf strSourceTmp <> "" Then
                ColIDs.Add Mid(strIDsTmp, 2)
                ColSource.Add Mid(strSourceTmp, 3)
                strIDsTmp = "": strSourceTmp = ""
            End If
        End If
    Loop
    
    strSource = Mid(strSource, 3)
    strIDs = Mid(strIDs, 2)
    If ColIDs.Count = 0 Then
        ColIDs.Add strIDs
        ColSource.Add strSource
    End If
    
    ReadAdviceSignSource = intRule
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceSign(ByVal lngҽ��ID As Long, ByVal int���� As Integer, ByVal str��Ա As String, ByVal datʱ�� As Date) As Long
'���ܣ���ȡָ��ҽ��������ǩ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ǩ��ID From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=[2] And ������Ա=[3] And ����ʱ��=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID, int����, str��Ա, datʱ��)
    If Not rsTmp.EOF Then
        GetAdviceSign = NVL(rsTmp!ǩ��ID, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceSigns(ByVal lngҽ��ID As Long, ByVal int���� As Integer, ByVal str��Ա As String, ByVal datʱ�� As Date) As String
'���ܣ���ȡָ��ҽ��������ǩ��ID�ַ���(�ಡ��ǩ�������--ȷ��ֹͣ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strAdciseSigns As String
    
    On Error GoTo errH
    
    strSQL = "Select distinct ǩ��ID" & vbNewLine & _
            "From ����ҽ��״̬ A" & vbNewLine & _
            "Where a.�������� = [2] And a.������Ա = [3] And a.����ʱ�� = [4] And Exists" & vbNewLine & _
            "  (Select 1 From ����ҽ��״̬ B" & vbNewLine & _
            "        Where a.�������� = b.�������� And a.������Ա = b.������Ա And a.����ʱ�� = b.����ʱ�� And b.ҽ��id = [1])"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID, int����, str��Ա, datʱ��)
    Do While Not rsTmp.EOF
        If NVL(rsTmp!ǩ��ID, 0) <> 0 Then
            strAdciseSigns = strAdciseSigns & "," & rsTmp!ǩ��ID
        End If
        rsTmp.MoveNext
    Loop
    GetAdviceSigns = Mid(strAdciseSigns, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicesSameSign(ByVal lngǩ��id As Long) As String
'���ܣ���ȡ��ͬǩ��ID�Ķ���ҽ��IDs(�ಡ��ǩ��һ��ǩ��ʱ������ĳ�����˵Ķ���ҽ����ȷ��ֹͣ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strAdciseSigns As String
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(b.���ID,b.ID) as ҽ��ID" & vbNewLine & _
            " From ����ҽ��״̬ A,����ҽ����¼ B" & vbNewLine & _
            " Where a.ҽ��id=b.id" & vbNewLine & _
            " and a.ǩ��id=[1]"


    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngǩ��id)
    Do While Not rsTmp.EOF
        If NVL(rsTmp!ҽ��ID, 0) <> 0 Then
            strAdciseSigns = strAdciseSigns & "," & rsTmp!ҽ��ID
        End If
        rsTmp.MoveNext
    Loop
    GetAdvicesSameSign = Mid(strAdciseSigns, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetAdviceTime(ByVal lngҽ��ID As Long, ByVal int���� As Integer) As Date
'���ܣ���ȡҽ��ָ��������ʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select Max(����ʱ��) as ʱ�� From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID, int����)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!ʱ��) Then
            GetAdviceTime = rsTmp!ʱ��
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetMergeIDs(ByRef vsAdvice As VSFlexGrid, ByVal lngRow As Long, ByVal COL_���ID As Long, ByVal COL_ID As Long) As String
'���ܣ���ȡָ��һ����ҩ��ҽ��ID��(��һ����ҩ���ص�ǰҽ��ID)
'������lngRow=һ����ҩ�Ŀ�ʼҩƷ��
    Dim lng���ID As Long, i As Long
    Dim strҽ��ID As String
    
    With vsAdvice
        lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
        For i = lngRow To .Rows - 1
            If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                strҽ��ID = strҽ��ID & "," & Val(.TextMatrix(i, COL_ID))
            Else
                Exit For
            End If
        Next
    End With
    
    GetMergeIDs = Mid(strҽ��ID, 2)
End Function

Public Function GetRXKey(ByRef rsRXKey As ADODB.Recordset, ByVal strKey As String, ByVal strҽ��ID As String) As String
'���ܣ�����ҩƷ�����������ƹؼ���,���ڴ���NO����
'������strKey=��ǰ����NO��Key,������������������Key����
'      strҽ��ID=��ǰҩƷ��ҽ��ID����һ����ҩ�������ID��"ID1,ID2,..."
'                һ����ҩ��ʼ�л����ҩƷ�вŴ���,һ����ҩ�м��д����
    Dim intNextCount As Integer
    Dim strNextID As String
    
    rsRXKey.Filter = "Key='" & strKey & "'"
    If rsRXKey.EOF Then
        strNextID = zlStr.ListMinus(strҽ��ID, "")
        intNextCount = UBound(Split(strNextID, ",")) + 1
        
        rsRXKey.AddNew
        rsRXKey!Key = strKey
        rsRXKey!ҽ��ID = strNextID
        rsRXKey!���� = intNextCount
        rsRXKey!���� = 1
        rsRXKey.Update
    ElseIf strҽ��ID <> "" Then
        strNextID = zlStr.ListMinus(strҽ��ID, rsRXKey!ҽ��ID)
        intNextCount = UBound(Split(strNextID, ",")) + 1
        
        rsRXKey!ҽ��ID = rsRXKey!ҽ��ID & "," & strNextID
        rsRXKey!���� = rsRXKey!���� + intNextCount
        rsRXKey.Update
    
        If rsRXKey!���� > gintRXCount Then
            strNextID = zlStr.ListMinus(strҽ��ID, "")
            intNextCount = UBound(Split(strNextID, ",")) + 1
            
            rsRXKey!���� = rsRXKey!���� + 1
            rsRXKey!ҽ��ID = strNextID
            rsRXKey!���� = intNextCount
            rsRXKey.Update
        End If
    ElseIf strҽ��ID = "" Then
        'һ����ҩ�м���,���ֵ�һ�еĹؼ���
    End If

    GetRXKey = rsRXKey!����
End Function

Public Function GetMergeCount(ByRef vsAdvice As VSFlexGrid, ByVal lngRow As Long, ByVal COL_���ID As Long, ByVal COL_�շ�ϸĿID As Long) As Long
'���ܣ���ȡָ��һ����ҩ��ҩƷ��������(��һ����ҩ����1��)
'������lngRow=һ����ҩ�Ŀ�ʼҩƷ��
    Dim lng���ID As Long, i As Long
    Dim strҩƷID As String
    
    With vsAdvice
        lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
        For i = lngRow To .Rows - 1
            If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                If InStr(strҩƷID & ",", "," & Val(.TextMatrix(i, COL_�շ�ϸĿID)) & ",") = 0 Then
                    strҩƷID = strҩƷID & "," & Val(.TextMatrix(i, COL_�շ�ϸĿID))
                End If
            Else
                Exit For
            End If
        Next
    End With
    
    GetMergeCount = UBound(Split(Mid(strҩƷID, 2), ",")) + 1
End Function

Public Function GetAdviceState(ByVal lngҽ��ID As Long, ByVal vDate As Date) As Integer
'���ܣ���ȡҽ����ָ��ʱ���ҽ��״̬(��Ҫ������ͣ����,��Ϊ��ͣ���õĲ���ʱ�����Ƿ���ʱ��)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select �������� From ����ҽ��״̬ Where ��������<>10 And ҽ��ID=[1] And ����ʱ��<=[2] Order by ����ʱ�� Desc"
    strSQL = "Select �������� From (" & strSQL & ") Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, vDate)
    If Not rsTmp.EOF Then
        GetAdviceState = rsTmp!��������
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistsSpecAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ��жϲ��˱���סԺ�Ƿ��Ѿ��´���ȷ�ϵ�����ҽ��(��Ժ��תԺ������)
'���أ�������ڣ�����ҽ����ʾ��Ϣ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    strSQL = "Select A.����,A.Ӥ��,A.ҽ������ From ����ҽ����¼ A,������ĿĿ¼ B" & _
        " Where A.������ĿID=B.ID And A.�������='Z' And B.�������� IN('5','6','11')" & _
        " And A.ҽ��״̬ Not IN(1,2,4) And A.����ID=[1] And A.��ҳID=[2]" & _
        " Order by A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    Do While Not rsTmp.EOF
        strMsg = strMsg & vbCrLf & "������" & rsTmp!ҽ������ & IIF(NVL(rsTmp!Ӥ��, 0) <> 0, "(Ӥ��" & NVL(rsTmp!Ӥ��, 0) & ")", "")
        rsTmp.MoveNext
    Loop
    If strMsg <> "" Then
        rsTmp.MoveFirst
        strMsg = "������������""" & rsTmp!���� & """�Ѿ�ȷ���´���������ҽ����" & vbCrLf & strMsg & vbCrLf & ""
    End If
    ExistsSpecAdvice = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistBalance(ByVal strNO As String) As Boolean
'���ܣ��ж�ָ�����շѻ��۵��Ƿ�����Ѿ��շѵ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From ������ü�¼ Where Mod(��¼����,10)=1 And ��¼״̬ IN(1,3) And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "BillExistBalance", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ҽ������(ByVal lngҽ��ID As Long) As String
'���ܣ�����ָ��ҽ���ĸ���������
'������lngҽ��ID=�ɼ��е�ҽ��ID(��ҩƷ�⣬�����IDΪ�յ�ҽ��ID)
'���أ���ʽΪ"��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��Ŀ,����,Ҫ��ID,���� From ����ҽ������ Where ҽ��ID=[1] Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get����ҽ������", lngҽ��ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "<Split1>" & rsTmp!��Ŀ & "<Split2>" & NVL(rsTmp!����, 0) & "<Split2>" & NVL(rsTmp!Ҫ��ID) & "<Split2>" & NVL(rsTmp!����)
        rsTmp.MoveNext
    Loop
    Get����ҽ������ = Mid(strSQL, Len("<Split1>") + 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getҽ����Ŀ����(ByVal lng��Ŀid As Long, Optional ByVal int���� As Integer = 2) As String
'���ܣ�����ָ��������ĿҪ��ĸ���������
'������lng��ĿID=������ĿID,int����=1���=2סԺ��=4���
'���أ���ʽΪ"��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select C.��Ŀ,C.����,C.Ҫ��ID,C.����" & _
        " From ��������Ӧ�� A,�����ļ��б� B,�������ݸ��� C" & _
        " Where A.������ĿID=[1] And A.Ӧ�ó���=[2]" & _
        " And A.�����ļ�ID=B.ID And B.����=7 And B.ID=C.�ļ�ID" & _
        " Order by C.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Getҽ����Ŀ����", lng��Ŀid, int����)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        '����һ��<Split2>��������ֱ�Ӷ�ȡ����Ŀ���������������Ǳ༭�����ģ������޸ĸ���ʱʶ�𣬼�frmAdviceEditEx���޸Ĵ���
        strSQL = strSQL & "<Split1>" & rsTmp!��Ŀ & "<Split2>" & NVL(rsTmp!����, 0) & "<Split2>" & NVL(rsTmp!Ҫ��ID) & "<Split2>" & NVL(rsTmp!����) & "<Split2>1"
        rsTmp.MoveNext
    Loop
    Getҽ����Ŀ���� = Mid(strSQL, Len("<Split1>") + 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckMedicineSended(ByVal lngAdviceID As Long, ByVal DateLast As Date) As Boolean
'���ܣ����ָ��ҽ�������һ�η����Ƿ��ѷ�ҩ
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1" & vbNewLine & _
            "From ����ҽ������ A, סԺ���ü�¼ B, ҩƷ�շ���¼ C" & vbNewLine & _
            "Where a.ҽ��id = [1] And a.ĩ��ʱ�� = [2] And a.No = b.No And" & vbNewLine & _
            "      a.��¼���� = b.��¼���� And a.ҽ��id = b.ҽ�����" & _
            " And b.Id = c.����id And c.���� In (9, 10) And Mod(c.��¼״̬, 3) = 1 And c.����� Is Null And Rownum = 1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckMedicineSended", lngAdviceID, DateLast)
    CheckMedicineSended = rsTmp.RecordCount = 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastSendMediCineID(ByVal lngAdviceID As Long, ByVal DateLast As Date, ByVal lng�������� As Long) As Long
'���ܣ�����ҩƷҽ��ID��ȡ���һ�η��͵ķ��ö�Ӧ���շ���ĿID(ҩƷ���ID)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select B.�շ�ϸĿID" & _
            " From ����ҽ������ A," & IIF(lng�������� = 1, "����", "סԺ") & "���ü�¼ B" & _
            " Where A.NO=B.NO And A.��¼����=B.��¼����" & _
            " And B.��¼״̬ IN(0,1,3) And B.ҽ�����=A.ҽ��ID And A.ҽ��ID=[1] And A.ĩ��ʱ��=[2]"
    '���ܸ�ҩƷδ�Ʒ�(���Ա�ҩ)
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetLastSendMediCineID", lngAdviceID, DateLast)
    If rsTmp.RecordCount > 0 Then
        GetLastSendMediCineID = Val("" & rsTmp!�շ�ϸĿID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowRollNotify(ByVal strPatis As String)
'���ܣ����һ�鲡���Ƿ������ֹͣ����Ҫ�����ջص�ҽ������������ʾ
'������strRollNotify=����ID:��ҳID,...
    Dim rsTmp As ADODB.Recordset, rsDrug As ADODB.Recordset
    Dim strSQL As String, strMsg As String, strSQLPati As String, strTemp As String
    Dim strThis As String, p As Long, n As Long, strUnRoll As String, blnDo As Boolean, lngҩƷID As Long
    Dim varPar(0 To 10) As String

    On Error GoTo errH
    strUnRoll = zlDatabase.GetPara("��ҩ���ջ�", glngSys, pסԺҽ������)
    strTemp = "Select C1 As ����ID,C2 As ��ҳID From Table(f_Num2list2([1]))"
    n = 0
    Do While True
        If Len(strPatis) < 4000 Then
            p = Len(strPatis) + 1
        Else
            p = InStrRev(Mid(strPatis, 1, 4000), ",")
        End If
        strThis = Mid(strPatis, 1, p - 1)
        
        If n > 10 Then
            strSQLPati = strSQLPati & vbNewLine & " Union All " & Replace(strTemp, "[1]", "'" & strThis & "'")
        Else
            varPar(n) = strThis
            strSQLPati = IIF(strSQLPati = "", "", strSQLPati & vbNewLine & " Union All ") & Replace(strTemp, "[1]", "[" & (n + 1) & "]")
        End If
        
        n = n + 1
        strPatis = Mid(strPatis, p + 1)
        If strPatis = "" Then Exit Do
    Loop
    
    '�����볬���ջ���һ�£���ֻ������ǰ״̬Ϊ(�Զ�)ֹͣ�ġ�
    strSQL = "(A.ִ��ʱ�䷽�� is NULL And (Nvl(A.Ƶ�ʴ���,0)=0 Or Nvl(A.Ƶ�ʼ��,0)=0 Or A.Ƶ�ʼ�� is NULL))"
    strSQL = _
        " Select /*+ Rule*/ A.����,A.ҽ������,A.ID,A.�������,A.�ϴ�ִ��ʱ��,A.�շ�ϸĿID,b.��������" & _
        " From ����ҽ����¼ A,������ҳ B,������ĿĿ¼ E,(" & strSQLPati & ") F" & _
        " Where A.������ĿID=E.ID And a.����id=b.����id and a.��ҳid=b.��ҳid And A.����ID = F.����ID And A.��ҳID = F.��ҳID" & _
        " And Not(A.�������='H' And E.��������='1') And Not(A.�������='Z' And E.�������� In('4','14'))" & _
        " And Nvl(A.ִ������,0)<>0 And A.�ܸ����� is NULL And Nvl(A.ҽ����Ч,0)=0" & _
        " And ((Not " & strSQL & " And A.ִ����ֹʱ��<A.�ϴ�ִ��ʱ��)" & _
        " Or (" & strSQL & " And Trunc(A.ִ����ֹʱ��)<Trunc(A.�ϴ�ִ��ʱ��)+1))" & _
        " And A.ҽ��״̬=8 And (A.���ID is Null Or A.������� IN('5','6'))" & _
        " And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3  And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ' And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ'" & _
        " And Not Exists(Select 1 From ����ҽ����¼ X Where ������� IN('5','6') And X.���ID=A.ID)" & _
        " Order by A.����ID,A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ջؼ��", varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
    Do While Not rsTmp.EOF
        blnDo = True
        
        If InStr(",5,6,", rsTmp!�������) > 0 And strUnRoll <> "" Then
            If Not IsNull(rsTmp!�շ�ϸĿID) Then
                lngҩƷID = rsTmp!�շ�ϸĿID
            Else
                lngҩƷID = GetLastSendMediCineID(Val(rsTmp!ID), CDate(rsTmp!�ϴ�ִ��ʱ��), Val(rsTmp!�������� & ""))
            End If
            If lngҩƷID <> 0 Then
                strSQL = "Select ��ҩ���� From ҩƷ��� Where ҩƷID = [1] And ��ҩ���� is Not Null"
                Set rsDrug = zlDatabase.OpenSQLRecord(strSQL, "�����ջؼ��", lngҩƷID)
                If rsDrug.RecordCount > 0 Then
                    If InStr("," & strUnRoll & ",", "," & rsDrug!��ҩ���� & ",") > 0 Then
                        If CheckMedicineSended(Val(rsTmp!ID), CDate(rsTmp!�ϴ�ִ��ʱ��)) Then
                            blnDo = False
                        End If
                    End If
                End If
            Else '�����ջأ�ҽ��δ�Ƿѣ����Ա�ҩ������ط��ñ�ɾ���ˣ��绮�۵���ɾ����
                blnDo = False
            End If
        End If
        If blnDo Then
            strMsg = strMsg & vbCrLf & "�񡡲��ˣ�" & rsTmp!���� & "��ҽ����" & rsTmp!ҽ������
        End If
        rsTmp.MoveNext
    Loop
    If strMsg <> "" Then
        MsgBox "������ֹͣ��ҽ�������ڷ��ͣ�" & vbCrLf & strMsg & vbCrLf & vbCrLf & "����ҽ������ʹ��""���ڷ����ջ�""���д���", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CheckPathNotEvaluete(ByVal lng����ID As Long, lng��ҳID As Long, Optional ByRef blnIsSend As Boolean, Optional ByRef str���� As String) As Boolean
'���ܣ����·�����˵�ǰ�Ƿ�δ����
'������blnIsSend  �����Ƿ��Ѿ�����
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select b.����,sysdate As ��ǰ����" & vbNewLine & _
            "From �����ٴ�·�� A, ����·��ִ�� B" & vbNewLine & _
            "Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.·����¼id And a.��ǰ�׶�id = b.�׶�id And a.��ǰ���� = b.���� And Rownum = 1 And" & vbNewLine & _
            "      Not Exists" & vbNewLine & _
            "(Select 1 From ����·������ C Where c.·����¼id = a.Id And c.�׶�id = a.��ǰ�׶�id And c.���� = a.��ǰ����)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ·������", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        str���� = Format(rsTmp!���� & "", "yyyy-MM-dd") & ""
        If Format(rsTmp!���� & "", "yyyy-MM-dd") = Format(rsTmp!��ǰ���� & "", "yyyy-MM-dd") Then
            blnIsSend = True '
        Else
            blnIsSend = False
        End If
        CheckPathNotEvaluete = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathNotEvalueteOut(ByVal lng�Һ�ID As Long, Optional ByRef blnIsSend As Boolean, Optional ByRef str���� As String) As Boolean
'���ܣ����·�����˵�ǰ�Ƿ�δ����
'������blnIsSend  �����Ƿ��Ѿ�����
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select b.����,sysdate As ��ǰ����" & vbNewLine & _
            "From ��������·�� A, ��������·��ִ�� B, ��������·����¼ C" & vbNewLine & _
            "Where C.�Һ�ID=[1] And C.·����¼ID=A.ID And a.Id = b.·����¼id And a.��ǰ�׶�id = b.�׶�id And a.��ǰ���� = b.���� And Rownum = 1 And" & vbNewLine & _
            "      Not Exists" & vbNewLine & _
            "(Select 1 From ��������·������ d Where d.·����¼id = a.Id And d.�׶�id = a.��ǰ�׶�id And d.���� = a.��ǰ����)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ·������", lng�Һ�ID)
    If rsTmp.RecordCount > 0 Then
        str���� = Format(rsTmp!���� & "", "yyyy-MM-dd") & ""
        If Format(rsTmp!���� & "", "yyyy-MM-dd") = Format(rsTmp!��ǰ���� & "", "yyyy-MM-dd") Then
            blnIsSend = True '
        Else
            blnIsSend = False
        End If
        CheckPathNotEvalueteOut = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathItemIsMust(ByVal bytִ�з�ʽ As Byte, ByVal int���� As Integer, ByVal lng·����¼Id As Long, _
                                    ByVal lng�׶�Id As Long, ByVal lng��Ŀid As Long, Optional ByVal intType As Integer) As Boolean
'����:���·����Ŀ�Ƿ��Ǳ������ɵ���Ŀ
    Dim blnMust As Boolean
    Dim strSQL As String
    Dim rsStep As ADODB.Recordset
    
    On Error GoTo errH:
    If bytִ�з�ʽ = 1 Then
        blnMust = True
    ElseIf bytִ�з�ʽ = 2 Or bytִ�з�ʽ = 4 Then  '����һ�λ����һ��
        strSQL = "Select ��ʼ����,�������� From " & IIF(intType = 1, "����·���׶�", "�ٴ�·���׶�") & " Where ID = [1]"
        Set rsStep = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng�׶�Id)
        If Not IsNull(rsStep!��ʼ����) Then
            If Not IsNull(rsStep!��������) Then
                blnMust = (int���� = Val("" & rsStep!��������))
                If blnMust Then   '�Ƿ����һ��
                    '�жϸ���Ŀ֮ǰ��û��ִ�й�(·������Ŀ����)
                    strSQL = "Select 1 From " & IIF(intType = 1, "��������·��ִ��", "����·��ִ��") & " Where ·����¼ID = [1] And �׶�ID = [2] And ��ĿID = [3] And ����<[4] And rownum<2"
                    Set rsStep = zlDatabase.OpenSQLRecord(strSQL, "���·��ҽ��", lng·����¼Id, lng�׶�Id, lng��Ŀid, int����)
                    If rsStep.RecordCount > 0 Then blnMust = False
                End If
            Else
                blnMust = True  '����
            End If
        End If
    End If
    
    CheckPathItemIsMust = blnMust
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CheckAdviceAppend(ByVal strAppend As String) As String
'���ܣ���ָ��ҽ�������븽����д������м��
'������strAppend=��ǰ���븽�����д�����,��ʽΪ"��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
'���أ�����Ҫ��д�����븽�����ݣ���"��Ŀ1,��Ŀ2..."
    Dim arrItem As Variant, arrSub As Variant
    Dim strItem As String, i As Integer
    
    arrItem = Split(strAppend, "<Split1>")
    For i = 0 To UBound(arrItem)
        arrSub = Split(arrItem(i), "<Split2>")
        If Val(arrSub(1)) = 1 And Trim(arrSub(3)) = "" Then
            strItem = strItem & "," & arrSub(0)
        End If
    Next
    
    CheckAdviceAppend = Mid(strItem, 2)
End Function

Public Function ReplaceAppend(ByVal strTarget As String, ByVal strSource As String) As String
'���ܣ�������������븽�����ݣ���ָ��ҽ������ͬ���հ׸������ȱʡ�滻
'������strTarget=Ҫ���滻��ҽ������������,��ʽΪ"��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
'      strSource=�������ҽ������������,��ʽ��ͬ��
'���أ��ѱ��滻��ҽ����������������ʽ��ͬ������ֵ������strTarget��ͬ��
    Dim arrTarget As Variant, arrSub1 As Variant
    Dim arrSource As Variant, arrSub2 As Variant
    Dim i As Integer, j As Integer
    Dim strReturn As String, blnReplace As Boolean
    
    If strTarget = "" Or strSource = "" Then
        ReplaceAppend = strTarget: Exit Function
    End If
    
    arrTarget = Split(strTarget, "<Split1>")
    arrSource = Split(strSource, "<Split1>")
    
    For i = 0 To UBound(arrTarget)
        arrSub1 = Split(arrTarget(i), "<Split2>")
        
        blnReplace = False
        For j = 0 To UBound(arrSource)
            arrSub2 = Split(arrSource(j), "<Split2>")
            If arrSub1(0) = arrSub2(0) Then
                If arrSub1(3) = "" And arrSub2(3) <> "" Then
                    arrSub1(3) = arrSub2(3): blnReplace = True
                End If
                Exit For
            End If
        Next
        
        strReturn = strReturn & "<Split1>" & arrSub1(0) & "<Split2>" & arrSub1(1) & "<Split2>" & arrSub1(2) & "<Split2>" & arrSub1(3)
        If UBound(arrSub1) >= 4 Then
            '���Զ��滻�˵ģ��൱��ȡ��ȱʡֵ���޸�ʱ����ʶ�𣬼�"Getҽ����Ŀ����"����
            If Not blnReplace Then strReturn = strReturn & "<Split2>" & arrSub1(4)
        End If
    Next
    
    ReplaceAppend = Mid(strReturn, Len("<Split1>") + 1)
End Function

Public Function GetMaxDate(lng����ID As Long, lng��ҳID As Long, Optional intԭ�� As Integer) As Date
'���ܣ���ȡת�Ʋ��������ϴα䶯ʱ��
'������intԭ��=�����ϴα䶯��ԭ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    GetMaxDate = #1/1/1900#
    intԭ�� = 0
    
    strSQL = "Select ��ʼʱ��,��ʼԭ�� From ���˱䶯��¼" & _
        " Where ��ʼʱ�� is Not NULL And ��ֹʱ�� is NULL" & _
        " And ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        GetMaxDate = IIF(IsNull(rsTmp!��ʼʱ��), GetMaxDate, rsTmp!��ʼʱ��)
        intԭ�� = NVL(rsTmp!��ʼԭ��, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBabyRegList(ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByRef rsBaby As ADODB.Recordset) As String
'���ܣ���ȡ���˵�Ӥ�������б�
'������lng����ID=סԺ����Ϊ"��ҳID",���ﲡ��Ϊ"�Һ�ID"
'���أ�"����1<Split>����2<Split>����3..."
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select a.���,a.Ӥ������,a.Ӥ���Ա� as �Ա� From ������������¼ a Where a.����ID=[1] And a.��ҳID=[2] Order by a.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetBabyRegList", lng����ID, lng����ID)
    Set rsBaby = zlDatabase.CopyNewRec(rsTmp)
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "<Split>" & NVL(rsTmp!Ӥ������)
        rsTmp.MoveNext
    Loop
    GetBabyRegList = Mid(strSQL, Len("<Split>") + 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicePrintPage(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, _
    ByVal int��Ч As Integer, ByVal lngҳ�� As Long) As Long
'���ܣ�����ָ����ҽ����ӡҳ�ţ���ȡ���ҳһ���ҳ��ӡ��ǰ��ҳ��
'���أ��뵱ǰҳһ���ҳ��ӡ����ʼҳ�ţ�������ǰ��ҳ�����������������Ĵ�ӡҳ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "(Select Nvl(Max(ҳ��),0) From ����ҽ����ӡ Where ����ID=[1] And ��ҳID=[2] And Nvl(Ӥ��, 0)=[3] And Nvl(��Ч,0)=[4] And ҽ��ID Is Null)"
    strSQL = "Select D.ҳ�� From ����ҽ����ӡ A,����ҽ����¼ B,����ҽ����¼ C,����ҽ����ӡ D" & _
        " Where A.ҽ��ID=B.ID And B.���ID=C.���ID And C.ID=D.ҽ��ID And D.ҳ��=[5]-1 And D.ҳ��>=" & strSQL & _
        " And A.����ID=[1] And A.��ҳID=[2] And Nvl(A.Ӥ��,0)=[3] And Nvl(A.��Ч,0)=[4] And A.ҳ��=[5] And A.�к�=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdvicePrintPage", lng����ID, lng��ҳID, intӤ��, int��Ч, lngҳ��)
    If Not rsTmp.EOF Then
        GetAdvicePrintPage = GetAdvicePrintPage(lng����ID, lng��ҳID, intӤ��, int��Ч, rsTmp!ҳ��)
    Else
        GetAdvicePrintPage = lngҳ��
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Open_LIS_Report(ByVal frmParent As Object, ByVal lngҽ��ID As Long, ByVal lng����ID As Long, ByVal blnCurrMoved As Boolean, ByVal blnPrint As Boolean, Optional ByVal bln��Ԥ����ӡ As Boolean) As Boolean
'����LiwWork��ӡ��ͼ�ε�LIS����
'������bln��Ԥ����ӡ �ڵ��ñ���Ԥ��ʱ�Ƿ���ʾԤ������Ĵ�ӡ��ť��false Ҫ��ʾ��true ����ʾ
    Dim strChart(0 To 8) As String
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim intLoop As Integer
    Dim objLisWork As Object
    Dim lng���ͺ� As Long, lng�걾id As Long
                    
    On Error GoTo errHandle
    Set objLisWork = CreateObject("zl9LisWork.clsLISImg")
    
    strSQL = "select ���ͺ� from ����ҽ������ a,����ҽ����¼ b where b.id = a.ҽ��id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Open_LIS_Report", lngҽ��ID)
    If Not rsTmp.EOF Then
        lng���ͺ� = NVL(rsTmp!���ͺ�, 0)
    End If
    strSQL = "select max(�걾ID) as ID from ������Ŀ�ֲ� where ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Open_LIS_Report", lngҽ��ID)
    If Not rsTmp.EOF Then
        lng�걾id = NVL(rsTmp!ID, 0)
    End If
    If lng�걾id = 0 Then
        strSQL = "select ID from ����걾��¼ b where b.ҽ��id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Open_LIS_Report", lngҽ��ID)
        If Not rsTmp.EOF Then
            lng�걾id = NVL(rsTmp!ID, 0)
        End If
    End If
    If lng���ͺ� = 0 Or lng�걾id = 0 Then Exit Function
    
    strSQL = "select id from ����ͼ���� where �걾id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lng�걾id)
    intLoop = 0
    Do Until rsTmp.EOF
        If Not objLisWork Is Nothing Then
            If objLisWork.Get_Chart2d_File(App.Path, rsTmp("ID")) Then
                strChart(intLoop) = App.Path & "\" & rsTmp("ID") & ".cht"
            End If
        End If
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    If Not objLisWork Is Nothing Then
        If objLisWork.Get_ReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
            If Not blnPrint And bln��Ԥ����ӡ Then
                strTmp = "DisabledPrint=1"
            Else
                strTmp = "DisabledPrint=0"
            End If
            Call ReportOpen(gcnOracle, glngSys, strReportCode, frmParent, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngҽ��ID, _
                            "����ID=" & lng����ID, "�걾ID=" & lng�걾id, _
                            "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), "ͼ��4=" & strChart(3), _
                            "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                            "ͼ��9=" & strChart(8), strTmp, IIF(blnPrint, 2, 1))
        End If
    End If
    'ɾ��ͼ���ļ�
    For intLoop = 0 To 8
        If strChart(intLoop) <> "" Then
            If Dir(strChart(intLoop)) <> "" Then Kill strChart(intLoop)
        End If
    Next
    
    Open_LIS_Report = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCuvetteNumber(rsNumber As ADODB.Recordset, ByVal str���� As String, ByVal lngҽ��ID As Long, _
    ByVal lng���ID As Long, ByVal str��� As String, ByVal int�������� As Integer, ByVal lngִ�п���ID As Long, _
    ByVal intӤ�� As Integer, ByVal lng������ĿID As Long, ByVal int���� As Integer, ByVal str�걾 As String, ByVal lng�ɼ�����ID As Long) As String
    '���ܣ��Լ���ҽ��������������
    '      1.һ���ɼ���ͬһ����ҽ��ʹ����ͬ����������
    '      2.��ͬ����ļ���ʹ����ͬ����������
    '      3.У���������:12λ��"����+ҽ��ID"
    '������rsNumber=��̬��¼��������"���롢���ID����������"���ֶ�
    Dim strTmp���� As String, strTmp���� As String
    
    If str��� = "C" And str���� <> "" Then '������Ŀ���й���
        rsNumber.Filter = "���ID=" & lng���ID
        If rsNumber.EOF Then
            rsNumber.Filter = "������Ŀid=" & lng������ĿID
            If rsNumber.EOF Then
                rsNumber.Filter = "����='" & str���� & "' And ִ�п���ID=" & lngִ�п���ID & " And Ӥ��=" & intӤ�� & _
                    " And ������־=" & int���� & " And �걾='" & str�걾 & "' And �ɼ�����ID=" & lng�ɼ�����ID
                If rsNumber.EOF Then
                    '�����µ�����
                    rsNumber.AddNew
                    rsNumber!���� = str����
                    rsNumber!���ID = lng���ID
'                    rsNumber!�������� = str���� & Format(lngҽ��ID, Replace(Space(12 - Len(str����)), " ", "0"))
                    rsNumber!�������� = zlDatabase.GetNextNo(125, lngҽ��ID)
                    rsNumber!������ĿID = lng������ĿID
                    rsNumber!ִ�п���ID = lngִ�п���ID
                    rsNumber!Ӥ�� = intӤ��
                    rsNumber!������־ = int����
                    rsNumber!�걾 = str�걾
                    rsNumber!�ɼ�����ID = lng�ɼ�����ID
                    rsNumber.Update
                    
                    strTmp���� = rsNumber!��������
                Else
                    '��ͬ���롢ִ�п��ҡ�Ӥ���ļ���ʹ����ͬ����������
                    strTmp���� = NVL(rsNumber!����)
                    strTmp���� = NVL(rsNumber!��������)
                    
                    rsNumber.AddNew
                    rsNumber!���� = strTmp����
                    rsNumber!���ID = lng���ID
                    rsNumber!�������� = strTmp����
                    rsNumber!������ĿID = lng������ĿID
                    rsNumber!ִ�п���ID = lngִ�п���ID
                    rsNumber!Ӥ�� = intӤ��
                    rsNumber!������־ = int����
                    rsNumber!�걾 = str�걾
                    rsNumber!�ɼ�����ID = lng�ɼ�����ID
                    rsNumber.Update
                End If
            Else
                '�����µ����룺��ͬ�����ҽ��ʹ��"��ͬ��"����
                rsNumber.AddNew
                rsNumber!���� = str����
                rsNumber!���ID = lng���ID
'                rsNumber!�������� = str���� & Format(lngҽ��ID, Replace(Space(12 - Len(str����)), " ", "0"))
                rsNumber!�������� = zlDatabase.GetNextNo(125, lngҽ��ID)
                rsNumber!������ĿID = lng������ĿID
                rsNumber!ִ�п���ID = lngִ�п���ID
                rsNumber!Ӥ�� = intӤ��
                rsNumber!������־ = int����
                rsNumber!�걾 = str�걾
                rsNumber!�ɼ�����ID = lng�ɼ�����ID
                rsNumber.Update
                
                strTmp���� = rsNumber!��������
            End If
        Else
            'һ���ɼ��ļ�����Ŀʹ����ͬ������
            strTmp���� = NVL(rsNumber!����)
            strTmp���� = NVL(rsNumber!��������)
            
            rsNumber.AddNew
            rsNumber!���� = strTmp����
            rsNumber!���ID = lng���ID
            rsNumber!�������� = strTmp����
            rsNumber!������ĿID = lng������ĿID
            rsNumber!ִ�п���ID = lngִ�п���ID
            rsNumber!Ӥ�� = intӤ��
            rsNumber!������־ = int����
            rsNumber!�걾 = str�걾
            rsNumber!�ɼ�����ID = lng�ɼ�����ID
            rsNumber.Update
        End If
    ElseIf str��� = "E" And int�������� = 6 Then
        '�ɼ���ʽʹ����ҽ����ͬ(���)������
        If Not rsNumber.EOF Then
            If NVL(rsNumber!���ID, 0) = lngҽ��ID Then
                strTmp���� = NVL(rsNumber!��������)
            End If
        End If
    End If
    
    GetCuvetteNumber = strTmp����
End Function

Public Function GetExecDays(ByVal str�ֽ�ʱ�� As String) As ADODB.Recordset
'���ܣ����ݵ�ǰҽ����ִ��ʱ�䴮���ز��ظ���ִ��������¼��
    Dim rsTmp As ADODB.Recordset
    Dim arrTmp As Variant, i As Long, strTmp As String
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "�շ�ʱ��", adVarChar, 10
    rsTmp.Fields.Append "����", adInteger '���ھ����Ƿ�����Ѵ��ڵ��б�
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    arrTmp = Split(str�ֽ�ʱ��, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = Format(arrTmp(i), "yyyy-MM-dd")
        rsTmp.Filter = "�շ�ʱ��='" & strTmp & "'"
        If rsTmp.EOF Then
            rsTmp.AddNew
            rsTmp!�շ�ʱ�� = strTmp
            rsTmp!���� = 0
            rsTmp.Update
        End If
    Next
    rsTmp.Filter = ""
    Set GetExecDays = rsTmp
End Function


Public Function AdviceMoneyMake(ByVal lng����ID As Long, ByVal lng��ҳID As Long, rsMoneyNow As Recordset, rsMoneyDay As ADODB.Recordset, _
    ByVal lngҽ��ID As Long, ByVal lng������ĿID, ByVal lng�շ���ĿID As Long, ByVal lngִ�в���id As Long, ByVal str�Թܱ��� As String, _
    ByVal str�շ���� As String, ByVal int�շѷ�ʽ As Integer, ByVal str�ֽ�ʱ�� As String, ByVal byt��Դ As Byte, ByRef lng���ô��� As Long, ByVal dbl���� As Double, _
    Optional ByVal lng��ǰҽ��ID As Long, Optional ByVal lng���ͺ� As Long, Optional ByVal dbl�Ƽ����� As Double, Optional rsExec As Recordset, _
    Optional ByVal lng���㷽ʽ As Long, Optional ByVal strƵ�� As String, Optional ByVal dbl���� As Double, Optional ByVal int��Ч As Integer = 1, _
    Optional ByVal int�������� As Integer, Optional ByVal str������� As String, Optional ByVal str�������� As String, Optional ByVal str��λ���� As String, Optional ByRef dbl�շ����� As Double, Optional ByVal strMinDate As String) As Boolean
'���ܣ��ж�ָ����ҽ�������Ƿ�Ӧ�ò���
'������lng��ҳID=סԺ���˲�ʹ�ã����ﲡ�˴���0���־���Һ�
'      rsMoneyNow=��ǰ���˱���Ҫ���͵ķ���,��̬��¼��(�շѷ�ʽ=-1,��ʾ�״β���ʱ��һ��ֻ��һ�ε���Ŀ�ļ�¼)
'      rsMoneyDay=��ǰ���˵����ѷ��͵ķ���,��̬��¼��
'      lngҽ��ID=һ��ҽ����ID
'      str�ֽ�ʱ��=���η��͵�ִ��ʱ�䴮���Զ��ŷָ��������ų�����ͣ��ʱ���
'      byt��Դ:1-���2-סԺ
'      dbl�Ƽ�����=�շ���Ŀ�ļƼ�����
'      ����=��ǰ�з���ҽ����������Ϣ
'      lng��ǰҽ��ID=��ǰ��ҽ��id
'      str��������=����ҽ��������������
'      str��λ���� �����Ŀҽ����ҽ���ļ�鲿λ�ͷ��� �̶���ʽ����鲿λ<sTab>��鷽�����磺"ͷ��<sTab>ƽɨ"
'      dbl�շ�����  �շ����Σ���Ҿ��������Ҳ�ָ����Ϊ0
'      strMinDate   ��ѯ�ѷ��͵�ҽ�����շ����ʱ����Сʱ��
'�����Ǽ�����ʱҽ����������֯����
'1��������ѡƵ�ʡ������ԡ���Ҫʱ�Ͳ���ʱ�Ե�����Ϊ���Ρ�
'2������һ���Ժ���ҪʱƵ�ʵ�ҽ��ȡ������Ϊ���Ρ�
'3��������ѡƵ��ȡ������Ϊ���Σ����һ��ȡ�������Ե���ȡĩ��Ϊ���Σ���������������ƣ�����80������25��ÿ��4�Σ���ôִ�еǼ�ʱ����ִ���ĴΣ�ǰ���α�������Ϊ25�����Ĵ�Ϊ80����25ȡģ=5��
'4������ִ�еǼ�ҳ��ҽ���嵥�����������У��������Σ�������ʾ�������Ρ�
'5��ҽ���༭ʱ������¼���״�������
'���أ�
'      lng���ô���=һ��ֻ��һ��ʱ��3,4,5,6,7�������ر��η���Ҫ��ȡ�Ĵ���
'      dbl����=�ܵķ��ʹ���������
'      rsExec=ҽ��ִ�мƼ۵�����
    Dim lng����ID As Long, blnMakeMoney As Boolean
    Dim rsDays As ADODB.Recordset, i As Long
    Dim arrTmp As Variant
    Dim dbl���� As Double
    Dim strDate As String
    Dim dbl����Tmp As Double
    Dim strSQL As String, rsTmp As Recordset, strTmp As String
    Dim str��λ As String, str���� As String
    Dim blnTmp As Boolean
    Dim dblTmp�շ����� As Double
    Dim str��С���� As String
    
    blnMakeMoney = True
    lng���ô��� = 1
    dbl�շ����� = 0
    
    If str��λ���� <> "" Then
        str��λ = Split(str��λ����, "<sTab>")(0)
        str���� = Split(str��λ����, "<sTab>")(1)
    End If
    
    If int�շѷ�ʽ = 9 Then
        '�Զ���
        '������ҽӿڣ��ӿڳ�����������ȡ
        strTmp = lngҽ��ID & "<sTab>" & lng��ǰҽ��ID & "<sTab>" & lng������ĿID & "<sTab>" & lng�շ���ĿID & "<sTab>" & int�շѷ�ʽ & "<sTab>" & str��λ & "<sTab>" & str���� & "<sTab>" & (dbl���� * dbl�Ƽ�����)
        dblTmp�շ����� = -1
        On Error Resume Next
        Call CreatePlugInOK(IIF(byt��Դ = 1, p����ҽ���´�, pסԺҽ������), -1)
        If Not gobjPlugIn Is Nothing Then
            blnTmp = gobjPlugIn.AdviceMakeFee(glngSys, IIF(byt��Դ = 1, p����ҽ���´�, pסԺҽ������), strTmp, rsMoneyNow, dblTmp�շ�����)
            Call zlPlugInErrH(err, "AdviceMakeFee")
            If err.Number <> 0 Then
                '�ӿڳ�����
                dblTmp�շ����� = -1
            Else
                If blnTmp Then
                    If dblTmp�շ����� = 0 Then
                        '���������Ϊ0������Ϊ���β����շ���
                        blnMakeMoney = False
                    End If
                    If dblTmp�շ����� > 0 Then
                        dbl�շ����� = dblTmp�շ�����
                    End If
                End If
            End If
        End If
        err.Clear: On Error GoTo 0
        
        On Error GoTo errH
        
        If dblTmp�շ����� = -1 Then
            strSQL = "Select zl_fun_CustomExpenses([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17]) as ���ؽ�� From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "AdviceMoneyMake", lng����ID, lng��ҳID, byt��Դ, lng��ǰҽ��ID, lngҽ��ID, int��Ч, strƵ��, lng������ĿID, lng�շ���ĿID, _
                                                lngִ�в���id, str�������, str�շ����, dbl����, dbl����, dbl�Ƽ�����, int��������, lng���㷽ʽ)
            If rsTmp.RecordCount > 0 Then
                strTmp = rsTmp!���ؽ�� & ""
                If Val(Split(strTmp, ":")(0)) = 0 Then
                    '����ȡ
                    blnMakeMoney = False
                Else
                    'Ҫ��ȡ
                    If InStr(strTmp, ":") > 0 Then
                        If Val(Split(strTmp, ":")(1)) > 0 Then lng���ô��� = Val(Split(strTmp, ":")(1)): dbl�շ����� = lng���ô���
                    End If
                End If
            End If
        End If
    End If
    
    If int�շѷ�ʽ = 0 Then
        '�����շѵģ�����ڱ��η����С���ҽ�����Ƿ��ų�
        rsMoneyNow.Filter = "(ҽ��ID=" & lngҽ��ID & " And ������ĿID=" & lng������ĿID & " And �շѷ�ʽ=5)" & _
            " Or (ҽ��ID=" & lngҽ��ID & " And ������ĿID=" & lng������ĿID & " And �շѷ�ʽ=6)" 'Or��ʹ��
        If Not rsMoneyNow.EOF Then blnMakeMoney = False
    ElseIf int�շѷ�ʽ = 1 Then '�����Թܷ���(һ�η���ֻ��ȡһ��)
        If str�Թܱ��� <> "" Then
            '��ͬ����(�Թ�)ֻ��ȡһ��
            rsMoneyNow.Filter = "�Թܱ���='" & str�Թܱ��� & "' And ��������='" & str�������� & "' And �շ���ĿID=" & lng�շ���ĿID & " And �շѷ�ʽ<>-1"
            If Not rsMoneyNow.EOF Then blnMakeMoney = False
            
            'ֻ��ȡ�Թܶ�Ӧ�����ķ���
            If blnMakeMoney And str�շ���� = "4" Then
                lng����ID = GetTubeMaterial(str�Թܱ���)
                If lng����ID <> 0 And lng�շ���ĿID <> lng����ID Then blnMakeMoney = False
            End If
        End If
    ElseIf int�շѷ�ʽ = 2 Then 'һ�η���ֻ��ȡһ��
        rsMoneyNow.Filter = "������ĿID=" & lng������ĿID & " And �շ���ĿID=" & lng�շ���ĿID & " And �շѷ�ʽ<>-1"
        If Not rsMoneyNow.EOF Then blnMakeMoney = False
    ElseIf InStr(",3,4,5,6,7,", int�շѷ�ʽ) > 0 Then
        '3-����ֻ��ȡһ�Σ�4-����δִ����ȡһ�Σ�5-����ֻ��ȡһ�Σ��ų�������Ŀ��6-����δִ����ȡһ�Σ��ų�������Ŀ
        
        '�����շѵģ�����ڱ��η����С���ҽ�����Ƿ��ų�
        If int�շѷ�ʽ = 7 Then
            rsMoneyNow.Filter = "(ҽ��ID=" & lngҽ��ID & " And ������ĿID=" & lng������ĿID & " And �շѷ�ʽ=5)" & _
                " Or (ҽ��ID=" & lngҽ��ID & " And ������ĿID=" & lng������ĿID & " And �շѷ�ʽ=6)" 'Or��ʹ��
            If Not rsMoneyNow.EOF Then blnMakeMoney = False
        End If
        
        If blnMakeMoney Then
            Set rsDays = GetExecDays(str�ֽ�ʱ��)
                        
            '�ȴӱ��η����е���(Ƶ��Ϊһ��һ����û���յģ��ж�ʱ��������ȡ,�Ա����������ҽ��"�״β���"ʱ������Ϊ���״�)
            For i = 1 To rsDays.RecordCount
                rsMoneyNow.Filter = "�շ�ʱ��='" & rsDays!�շ�ʱ�� & "' And ������ĿID=" & lng������ĿID & " And �շ���ĿID=" & lng�շ���ĿID & _
                    IIF(int�շѷ�ʽ = 7, "", " And �շѷ�ʽ<>-1") & _
                    IIF((int�շѷ�ʽ = 4 Or int�շѷ�ʽ = 6) And lngִ�в���id <> 0, " And ִ�в���ID=" & lngִ�в���id, "")
                If rsMoneyNow.RecordCount > 0 Then rsDays!���� = 1
                rsDays.MoveNext
            Next
            '�ٴ��ѷ����е���(���켰����ִ�е�)
            rsDays.Filter = "����=0"
            For i = 1 To rsDays.RecordCount
                If i = 1 Then
                    If rsMoneyDay Is Nothing Then
                        If strMinDate = "" Then
                            str��С���� = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd")
                        Else
                            str��С���� = Format(strMinDate, "yyyy-MM-dd")
                        End If
                        Call GetPatiDayMoneyDetail(rsMoneyDay, lng����ID, lng��ҳID, byt��Դ, CDate(str��С����))
                    End If
                End If
                rsMoneyDay.Filter = "�շ�ʱ��='" & rsDays!�շ�ʱ�� & "' And ������ĿID=" & lng������ĿID & " And �շ���ĿID=" & lng�շ���ĿID & _
                    IIF(int�շѷ�ʽ = 7, "", " And �շѷ�ʽ<>-1") & _
                    IIF((int�շѷ�ʽ = 4 Or int�շѷ�ʽ = 6) And lngִ�в���id <> 0, " And ִ�з�=0 And ִ�в���ID=" & lngִ�в���id, "")
                If rsMoneyDay.RecordCount > 0 Then rsDays!���� = 1
                rsDays.MoveNext
            Next
        End If
    End If
                            
    '��¼�����η�����ϸ��Ŀ��¼��
    If InStr(",3,4,5,6,7,", int�շѷ�ʽ) > 0 Then
        If int�շѷ�ʽ = 7 Then
            If blnMakeMoney Then
                rsDays.Filter = "����=0"    'û�չ�����Щ��(Ƶ��Ϊһ��һ�ε�δ�յĵ����չ���)���״β���
                lng���ô��� = dbl���� - rsDays.RecordCount
                blnMakeMoney = lng���ô��� > 0
            End If
        Else
            rsDays.Filter = "����=0"
            blnMakeMoney = rsDays.RecordCount > 0
            lng���ô��� = rsDays.RecordCount    'һ��һ�Σ��ж�����Ҫ�վ��ж��ٴ�
        End If
        If blnMakeMoney Or int�շѷ�ʽ = 7 And lng���ô��� = 0 Then
            For i = 1 To rsDays.RecordCount
                rsMoneyNow.AddNew
                rsMoneyNow!ҽ��ID = lngҽ��ID
                rsMoneyNow!������ĿID = lng������ĿID
                rsMoneyNow!�շ���ĿID = lng�շ���ĿID
                rsMoneyNow!�Թܱ��� = str�Թܱ���
                rsMoneyNow!�������� = str��������
                
                '�״β���ʱ�����Ƶ��Ϊһ��һ�Σ�������ķ��ô���Ϊ0,Ϊ���ñ��κ������͵�����ҽ����ȷ�������Ƿ���ȡ����Ҫ������¼�����շѷ�ʽ�����¼Ϊ-1
                rsMoneyNow!�շѷ�ʽ = IIF(int�շѷ�ʽ = 7 And lng���ô��� = 0, -1, int�շѷ�ʽ)
                rsMoneyNow!�շ�ʱ�� = rsDays!�շ�ʱ��
                rsMoneyNow!ִ�в���ID = lngִ�в���id
                rsMoneyNow.Update
            
                rsDays.MoveNext
            Next
        End If
    ElseIf blnMakeMoney Then
        rsMoneyNow.AddNew
        rsMoneyNow!ҽ��ID = lngҽ��ID
        rsMoneyNow!������ĿID = lng������ĿID
        rsMoneyNow!�շ���ĿID = lng�շ���ĿID
        rsMoneyNow!�Թܱ��� = str�Թܱ���
        rsMoneyNow!�������� = str��������
        rsMoneyNow!�շѷ�ʽ = int�շѷ�ʽ
        '�����Ŀר��������ʹ��
        rsMoneyNow!��ҽ��ID = lng��ǰҽ��ID
        rsMoneyNow!��鲿λ = str��λ
        rsMoneyNow!��鷽�� = str����
        rsMoneyNow!���� = IIF(dbl�շ����� > 0, dbl�շ�����, dbl���� * dbl�Ƽ�����)
        
        If str�ֽ�ʱ�� <> "" Then
            rsMoneyNow!�շ�ʱ�� = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd")  '��ʱ����ʱû���ô�
        Else
            rsMoneyNow!�շ�ʱ�� = ""
        End If
        rsMoneyNow!ִ�в���ID = lngִ�в���id
        rsMoneyNow.Update
    End If
    '��ȡҽ��ִ�мƼ�(��ҩƷ����ҽ����ĲŴ洢)
    If InStr(",5,6,7,", "," & str������� & ",") = 0 Then
        If str�ֽ�ʱ�� <> "" And Not rsExec Is Nothing Then
            arrTmp = Split(str�ֽ�ʱ��, ",")
            If dbl�շ����� = 0 Then
                dbl����Tmp = dbl���� * dbl�Ƽ�����
            Else
                '�������Ҵ����ˣ�����Ӧ�����µ�����������
                dbl����Tmp = dbl�շ�����
            End If
            If dbl���� = 0 Then dbl���� = 1
            For i = 0 To UBound(arrTmp)
                rsExec.AddNew
                rsExec!ҽ��ID = lng��ǰҽ��ID
                rsExec!���ͺ� = lng���ͺ�
                rsExec!Ҫ��ʱ�� = Format(arrTmp(i), "yyyy-MM-dd HH:mm:ss")
                rsExec!�շ�ϸĿID = lng�շ���ĿID
                rsExec!�������� = int��������
                If blnMakeMoney Then
                    '����Ҳ�������뵥������
                    If strƵ�� <> "" And ((lng���㷽ʽ = 0 Or lng���㷽ʽ = 3) And dbl���� > 0 Or lng���㷽ʽ = 1 Or lng���㷽ʽ = 2 Or str������� = "4") Then
                        '�����ͼ�ʱ����Ҫ��������
                        If int��Ч = 0 Then
                            '1��������ѡƵ�ʡ������ԡ���Ҫʱ�Ͳ���ʱ�Ե�����Ϊ���Ρ�
                            dbl���� = dbl�Ƽ����� * dbl����
                        Else
                            '3��������ѡƵ��ȡ������Ϊ���Σ����һ��ʣ�����������������������ƣ�����80������25��ÿ��4�Σ���ôִ�еǼ�ʱ����ִ���ĴΣ�ǰ���α�������Ϊ25�����Ĵ�Ϊ80����25ȡģ=5��
                            '�����п���û��¼��ִ��ʱ��,�ֽ�ʱ���ֻ��һ������������Ϊ����
                            If UBound(arrTmp) = 0 Then
                                If InStr(",1,2,3,4,5,6,7,9,", int�շѷ�ʽ) > 0 Then
                                    '�����շѷ�ʽ������  lng���ô��� �������ҽ�����ʹ����еķ��ü�¼��������һ��
                                    If dbl�շ����� > 0 Then
                                        dbl���� = dbl�շ�����
                                    Else
                                        dbl���� = lng���ô��� * dbl�Ƽ�����
                                    End If
                                Else
                                    dbl���� = dbl����Tmp
                                End If
                            Else
                                If i = UBound(arrTmp) Then
                                    dbl���� = dbl����Tmp
                                Else
                                    If dbl����Tmp >= dbl���� Then
                                        dbl���� = dbl�Ƽ����� * dbl����
                                    Else
                                        dbl���� = dbl����Tmp
                                    End If
                                    dbl����Tmp = dbl����Tmp - dbl����
                                End If
                            End If
                        End If
                    Else
                        dbl���� = dbl�Ƽ�����
                    End If
                    If i <> 0 Then
                        strDate = Format(arrTmp(i - 1), "yyyy-MM-dd")
                    End If
                    'һ�η�����ȡһ�Σ���ֻ�е�һ����ȡ
                    If InStr(",1,2,", int�շѷ�ʽ) > 0 Then
                        If i <> 0 Then dbl���� = 0
                    ElseIf InStr(",3,4,5,6,", int�շѷ�ʽ) > 0 Then
                        '3456����ֻ��ȡһ�εģ�����=0����ȡ��Ĭ�ϵ�һ��������
                        rsDays.Filter = "����=0 And �շ�ʱ��='" & Format(arrTmp(i), "yyyy-MM-dd") & "'"
                        If Not (rsDays.RecordCount > 0 And Format(arrTmp(i), "yyyy-MM-dd") <> strDate) Then
                            dbl���� = 0
                        End If
                    ElseIf int�շѷ�ʽ = 7 Then
                        '�����״β���ȡ�ģ�����=1����ȡ������=0��Ϊ�״�
                        rsDays.Filter = "����=1 And �շ�ʱ��='" & Format(arrTmp(i), "yyyy-MM-dd") & "'"
                        If rsDays.RecordCount = 0 And Format(arrTmp(i), "yyyy-MM-dd") <> strDate Then
                            dbl���� = 0
                        End If
                    End If
                Else
                    '�������ȡ��������Ϊ0
                    dbl���� = 0
                End If
                rsExec!���� = dbl����
                rsExec.Update
            Next
        End If
    End If
    AdviceMoneyMake = blnMakeMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetTubeMaterial(ByVal str�Թܱ��� As String) As Long
'���ܣ����ݹ����ȡ��Ӧ���Թܲ���ID
    Dim strSQL As String
    
    If grsTube Is Nothing Then
        On Error GoTo errH
        
        strSQL = "Select ����,����ID From ��Ѫ������ Where ����ID is Not NULL"
        Set grsTube = New ADODB.Recordset
        Set grsTube = zlDatabase.OpenSQLRecord(strSQL, "GetTubeMaterial")
    End If
    
    grsTube.Filter = "����='" & str�Թܱ��� & "'"
    If Not grsTube.EOF Then GetTubeMaterial = NVL(grsTube!����ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatiDayMoneyDetail(rsMoneyDay As ADODB.Recordset, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal byt��Դ As Byte, ByVal dat�շ�ʱ�� As Date) As Boolean
'���ܣ���ȡָ������ĳ��(dat�շ�ʱ��)֮��ҽ�������ķ�����Ŀ��ϸ
'������lng��ҳID=סԺ���˲�ʹ��
'      byt��Դ:1-����(��סԺ�������͵�����)��2-סԺ
'      dat�շ�ʱ�� ��Ҫ�жϵ������ʱ��
'���أ�rsMoneyDay������"������ĿID,�շ���ĿID,ִ�в���ID,ִ�з�,�շ�ʱ��"�ֶ�
'      һ�η�ĳ����ʱ���������ҽ��ִ��ʱ���ֲ��ڲ�ͬ�����ﲢ���������ڲ��������ܻ�©�㣬�����¡�
'        ����һ�η��ͺ�ִ��ʱ���ֱ�Ϊ��2014-11-1 XX:XX,2014-11-3 XX:XX,2014-11-5 XX:XX,2014-11-7 XX:XX
'        1.dat�շ�ʱ�� = 2014-11-1��ֻ�����죬2014-11-1��2014-11-7 ��ȷӦ����3�� 2014-11-1��2014-11-5��2014-11-7
'        2.dat�շ�ʱ�� = 2014-11-2��ֻ��1�죬2014-11-7             ��ȷӦ����2�� 2014-11-5��2014-11-7
    Dim strSQL As String, strҽ��IDs As String
    Dim rsTmp As ADODB.Recordset
    Dim rsִ��ʱ�� As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strToday As String, strDay As String
    Dim strTmp As String
    Dim n As Long, p As Long
    Dim strThis As String
    Dim strTabTmp As String
    Dim varPar(0 To 10) As String
        
    On Error GoTo errH
 
    'ִ���жϣ�
    '1.������ǽ�������ü�¼�е�ִ�в��ţ����Ҳ�Է��ü�¼�е�ִ�в���Ϊ׼�жϡ�
    '2.���͸��������⣬ҽ�����õ�ִ�п�����ҽ��ִ�п�����ͬ���Ժ������ͬ�ˣ��ú���Ҳ������Ӧ
    '3.ҽ��ִ��ʱ����Ӧ���õ�ִ��״̬Ҳ��ͬ����ǡ�
    '4.�״β��յ���Ŀ����û�в������ü�¼������ҽ�����ͼ�¼��,��Ҫ���������������ɵģ��Ա������״β��յ���Ŀ�ж�
    If byt��Դ = 1 Then
        strSQL = "Select A.������ĿID,C.�շ�ϸĿID as �շ���ĿID,C.ִ�в���ID,Decode(Nvl(C.ִ��״̬,0),0,0,1) as ִ�з�," & _
            " To_Char(b.�״�ʱ��,'yyyy-mm-dd') as �״�ʱ��,Trunc(b.ĩ��ʱ��) - Trunc(b.�״�ʱ��) + 1 AS ����,0 as �շѷ�ʽ," & _
            " To_Char(b.ĩ��ʱ��,'yyyy-mm-dd') as ĩ��ʱ��,A.Ƶ�ʼ��,A.�����λ,a.ҽ����Ч,a.id,nvl(a.���id,0) as ���id" & _
            " From ����ҽ����¼ A,����ҽ������ B,������ü�¼ C" & _
            " Where A.����ID=[1] And Nvl(A.��ҳID,0) = [2] And a.ҽ����Ч = 1 And A.ID=B.ҽ��ID And B.��¼����=C.��¼���� And B.NO=C.NO" & _
            " And B.ҽ��ID=C.ҽ����� And C.��¼״̬ IN(0,1) And b.ĩ��ʱ��>=[3]" & _
            " Union " & _
            " Select A.������ĿID,D.�շ�ϸĿid,D.ִ�п���ID as ִ�в���ID,0 as ִ�з�," & _
            " To_Char(B.�״�ʱ��,'yyyy-mm-dd') as �״�ʱ��,Trunc(b.ĩ��ʱ��) - Trunc(b.�״�ʱ��) + 1 AS ����,-1 as �շѷ�ʽ," & _
            " To_Char(b.ĩ��ʱ��,'yyyy-mm-dd') as ĩ��ʱ��,A.Ƶ�ʼ��,A.�����λ,a.ҽ����Ч,a.id,nvl(a.���id,0) as ���id" & _
            " From ����ҽ����¼ A,����ҽ������ B,����ҽ���Ƽ� D" & _
            " Where A.����ID=[1] And Nvl(A.��ҳID,0) = [2] And a.ҽ����Ч = 1" & _
            " And A.ID=B.ҽ��ID And b.ĩ��ʱ��>=[3] And A.ID=D.ҽ��ID And D.�շѷ�ʽ=7" & vbNewLine & _
            " And Not Exists (Select 1 From ������ü�¼ C Where c.�շ�ϸĿid=d.�շ�ϸĿid  And b.��¼���� = c.��¼���� And b.No = c.No And a.Id = c.ҽ�����)" & vbNewLine & _
            " Order by ������ĿID,�շ���ĿID"
    Else
        '����������ҽ������ͬ���ã����ܲ�ͬʱ���η���,Unionȥ�����ظ���¼
        '�״β��յ���Ŀ����û�в������ü�¼������ҽ�����ͼ�¼��,��Ҫ���������������ɵģ��Ա������״β��յ���Ŀ�ж�
        strSQL = "Select a.������Ŀid, c.�շ�ϸĿid As �շ���Ŀid, c.ִ�в���id, Decode(Nvl(c.ִ��״̬, 0), 0, 0, 1) As ִ�з�," & vbNewLine & _
            "     To_Char(b.�״�ʱ��,'yyyy-mm-dd') As �״�ʱ��, Decode(b.�״�ʱ��,null, 1,Trunc(b.ĩ��ʱ��) - Trunc(b.�״�ʱ��) + 1) As ����,0 as �շѷ�ʽ," & vbNewLine & _
            " To_Char(b.ĩ��ʱ��,'yyyy-mm-dd') as ĩ��ʱ��,A.Ƶ�ʼ��,A.�����λ,a.ҽ����Ч,a.id,nvl(a.���id,0) as ���id" & _
            " From ����ҽ����¼ A, ����ҽ������ B, סԺ���ü�¼ C" & vbNewLine & _
            " Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.ҽ��id And b.��¼���� = c.��¼���� And b.No = c.No And b.ҽ��id = c.ҽ����� And" & vbNewLine & _
            "      c.��¼״̬ In (0, 1) And b.ĩ��ʱ��>=[3]" & vbNewLine & _
            " Union " & vbNewLine & _
            " Select a.������Ŀid, D.�շ�ϸĿid, D.ִ�п���ID as ִ�в���id, 0 As ִ�з�," & vbNewLine & _
            " To_Char(b.�״�ʱ��,'yyyy-mm-dd') As �״�ʱ��,Decode(a.ҽ����Ч, 0, Trunc(b.ĩ��ʱ��) - Trunc(b.�״�ʱ��) + 1, 1) As ����,-1 as �շѷ�ʽ," & vbNewLine & _
            " To_Char(b.ĩ��ʱ��,'yyyy-mm-dd') as ĩ��ʱ��,A.Ƶ�ʼ��,A.�����λ,a.ҽ����Ч,a.id,nvl(a.���id,0) as ���id" & _
            " From ����ҽ����¼ A, ����ҽ������ B, ����ҽ���Ƽ� D" & vbNewLine & _
            " Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.ҽ��id And b.ĩ��ʱ��>=[3]" & vbNewLine & _
            " And A.ID=D.ҽ��ID And D.�շѷ�ʽ=7" & vbNewLine & _
            " And Not Exists (Select 1 From סԺ���ü�¼ C Where c.�շ�ϸĿid=d.�շ�ϸĿid  And b.��¼���� = c.��¼���� And b.No = c.No And a.Id = c.ҽ�����)" & vbNewLine & _
            " Order By ������Ŀid, �շ���Ŀid"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���켰������ҽ��", lng����ID, lng��ҳID, dat�շ�ʱ��)
    
    If byt��Դ = 2 Then
        rsTmp.Filter = "ҽ����Ч=0 and ���id=0"
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                If InStr("," & strҽ��IDs & ",", "," & rsTmp!ID & ",") = 0 Then strҽ��IDs = strҽ��IDs & "," & rsTmp!ID
                rsTmp.MoveNext
            Next
            strҽ��IDs = Mid(strҽ��IDs, 2)
            If strҽ��IDs <> "" Then
                strTmp = "Select /*+cardinality(x,10)*/ x.Column_Value as ҽ��ID From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)) X"
                n = 0
                Do While True
                    If Len(strҽ��IDs) < 4000 Then
                        p = Len(strҽ��IDs) + 1
                    Else
                        p = InStrRev(Mid(strҽ��IDs, 1, 4000), ",")
                    End If
                    strThis = Mid(strҽ��IDs, 1, p - 1)
                    
                    If n > 10 Then
                        strTabTmp = strTabTmp & vbNewLine & " Union All " & Replace(strTmp, "[2]", "'" & strThis & "'")
                    Else
                        varPar(n) = strThis
                        strTabTmp = IIF(strTabTmp = "", "", strTabTmp & vbNewLine & " Union All ") & Replace(strTmp, "[2]", "[" & (n + 2) & "]")
                    End If
                    n = n + 1
                    strҽ��IDs = Mid(strҽ��IDs, p + 1)
                    If strҽ��IDs = "" Then Exit Do
                Loop
                strSQL = "select a.ҽ��id,To_Char(Trunc(a.Ҫ��ʱ��), 'yyyy-mm-dd') As ִ��ʱ�� from ҽ��ִ��ʱ�� a" & _
                    " where a.Ҫ��ʱ��>=[1] and a.ҽ��id in (" & strTabTmp & ")  group by a.ҽ��id,Trunc(a.Ҫ��ʱ��)"
                Set rsִ��ʱ�� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���켰������ҽ��", dat�շ�ʱ��, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
            End If
        End If
        rsTmp.Filter = 0
    End If
    
    '����¼����ִ��ʱ��ֳɶ�����¼
    strToday = Format(dat�շ�ʱ��, "yyyy-MM-dd")
    Set rsMoneyDay = New ADODB.Recordset '�������Filter����
    Set rsMoneyDay = InitPatiExecDays
    
    For i = 1 To rsTmp.RecordCount
        If Val(rsTmp!ҽ����Ч & "") = 1 Then
            If Val(rsTmp!Ƶ�ʼ�� & "") = 1 And rsTmp!�����λ & "" = "��" Or rsTmp!�����λ & "" = "Сʱ" Or rsTmp!�����λ & "" = "����" Then
                For j = 1 To rsTmp!����
                    If j = 1 Then
                        strDay = Format(rsTmp!�״�ʱ��, "yyyy-MM-dd")
                    Else
                        strDay = Format(DateAdd("d", j - 1, CDate(rsTmp!�״�ʱ��)), "yyyy-MM-dd")
                    End If
                    If strDay >= strToday Then Call AddMoneyDayItem(rsTmp, rsMoneyDay, strDay)
                Next
            Else
                '��������������ܻ�©��
                strDay = Format(rsTmp!�״�ʱ��, "yyyy-MM-dd")
                If strDay >= strToday Then Call AddMoneyDayItem(rsTmp, rsMoneyDay, strDay)
                strDay = Format(rsTmp!ĩ��ʱ��, "yyyy-MM-dd")
                If strDay >= strToday Then Call AddMoneyDayItem(rsTmp, rsMoneyDay, strDay)
            End If
        Else
            '���������ͨ�� ҽ��ִ��ʱ�� ����ȷ��������һ��ִ��
            If Not rsִ��ʱ�� Is Nothing Then
                rsִ��ʱ��.Filter = "ҽ��id=" & rsTmp!ID & " or ҽ��id=" & rsTmp!���ID
                For j = 1 To rsִ��ʱ��.RecordCount
                    strDay = rsִ��ʱ��!ִ��ʱ�� & ""
                    Call AddMoneyDayItem(rsTmp, rsMoneyDay, strDay)
                    rsִ��ʱ��.MoveNext
                Next
            End If
        End If
        
        rsTmp.MoveNext
    Next
    rsMoneyDay.Filter = ""
    
    GetPatiDayMoneyDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddMoneyDayItem(ByVal rsTmp As ADODB.Recordset, ByRef rsMoneyDay As ADODB.Recordset, ByVal strDay As String)
'���ܣ��� rsMoneyDay �������
    rsMoneyDay.Filter = "������ĿID=" & Val("" & rsTmp!������ĿID) & " And �շ���ĿID=" & Val("" & rsTmp!�շ���ĿID) & _
        " And �շ�ʱ��='" & strDay & "' And ִ�з�=" & Val("" & rsTmp!ִ�з�) & " And �շѷ�ʽ=" & Val("" & rsTmp!�շѷ�ʽ)
    If rsMoneyDay.RecordCount = 0 Then
        rsMoneyDay.AddNew
        rsMoneyDay!������ĿID = Val("" & rsTmp!������ĿID)
        rsMoneyDay!�շ���ĿID = Val("" & rsTmp!�շ���ĿID)
        rsMoneyDay!ִ�в���ID = Val("" & rsTmp!ִ�в���ID)
        rsMoneyDay!ִ�з� = Val("" & rsTmp!ִ�з�)
        rsMoneyDay!�շѷ�ʽ = Val("" & rsTmp!�շѷ�ʽ)
        rsMoneyDay!�շ�ʱ�� = strDay
        rsMoneyDay.Update
    End If
End Sub

Private Function InitPatiExecDays() As ADODB.Recordset
'���ܣ���ʼ��ҽ����ط���ִ�еļ�¼��
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "������ĿID", adBigInt
    rsTmp.Fields.Append "�շ���ĿID", adBigInt
    rsTmp.Fields.Append "ִ�в���ID", adBigInt
    rsTmp.Fields.Append "�շѷ�ʽ", adInteger
    rsTmp.Fields.Append "ִ�з�", adInteger
    rsTmp.Fields.Append "�շ�ʱ��", adVarChar, 10
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set InitPatiExecDays = rsTmp
End Function

Public Function GetSkinTestResult(ByVal lng��Ŀid As Long, ByVal str��� As String) As Integer
'���ܣ�����Ƥ�Խ����ע������������
'������str���=Ƥ�Խ����ע����,��"(+)"
'���أ�-1-����,1-����,0-�޽��
    Dim arr���� As Variant, arr���� As Variant
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    If grsSkinTest Is Nothing Then
        strSQL = "Select ID,Nvl(�걾��λ,'����(+);����(-)') as ��ע From ������ĿĿ¼ Where ���='E' And ��������='1'"
        Set grsSkinTest = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(grsSkinTest, strSQL, "GetSkinTestResult")
    End If
    
    grsSkinTest.Filter = "ID=" & lng��Ŀid
    If grsSkinTest.EOF Then Exit Function
    
    arr���� = Split(Split(grsSkinTest!��ע, ";")(0), ",")
    arr���� = Split(Split(grsSkinTest!��ע, ";")(1), ",")
    
    For i = 0 To UBound(arr����)
        If Right(arr����(i), Len(str���)) = str��� Then
            GetSkinTestResult = 1: Exit Function
        End If
    Next
    For i = 0 To UBound(arr����)
        If Right(arr����(i), Len(str���)) = str��� Then
            GetSkinTestResult = -1: Exit Function
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲿����Ϣ�Ƿ���ر���
    '����:��ʾ����,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 13:11:01
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    If mlng���ű���ƽ������ = 0 Then
        strSQL = "Select Avg(length(����)) As ���� From ���ű�"
        On Error GoTo errH
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡ���ű����ƽ������")
        mlng���ű���ƽ������ = Val(NVL(rsTemp!����))
    End If
    '���ڱ��볤�ȿ��ܹ���,�޷���ʾ���ŵ�����,����Զ���ʾ�Ͳ���ʾ����,������5ʱ,����ʾ.С��5ʱ,��ʾ
   zlIsShowDeptCode = mlng���ű���ƽ������ <= 5
      
   Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���в��� As String = "", _
    Optional blnSendKeys As Boolean = True, Optional blnAddItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str���в���-���в�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim intIndex As Integer
    Dim strIDs As String, str���� As String, strLike As String
    strLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "*", "")
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = strLike & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���в��� <> "" Then
        str���� = zlCommFun.SpellCode(str���в���)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str���в���) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If NVL(!����) = strSearch Then lngDeptID = NVL(!ID): iCount = 0:  Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp): Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(NVL(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(NVL(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If NVL(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(NVL(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(NVL(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(NVL(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(NVL(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(NVL(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(NVL(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(NVL(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(NVL(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(NVL(!����)) Like strCompents Or UCase(Trim(NVL(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(NVL(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(NVL(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = NVL(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 And lngDeptID <= 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIF(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIF(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIF(blnNot���ȼ�, "", "���ȼ�,") & "����"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ," & IIF(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(NVL(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    If Cbo.Locate(cboDept, lngDeptID, True) = False Then
        If blnAddItem = True Then
            If rsTemp.RecordCount = 1 Then
                cboDept.RemoveItem cboDept.ListCount - 1
                cboDept.AddItem IIF(zlIsShowDeptCode, rsTemp!���� & "-", "") & rsTemp!����
                cboDept.ItemData(cboDept.ListCount - 1) = Val(NVL(rsTemp!ID))
                intIndex = cboDept.NewIndex
                cboDept.AddItem "�������ҡ�"
                cboDept.ItemData(cboDept.ListCount - 1) = 0
                cboDept.ListIndex = intIndex
            Else
                cboDept.RemoveItem cboDept.ListCount - 1
                cboDept.AddItem IIF(zlIsShowDeptCode, rsReturn!���� & "-", "") & rsReturn!����
                cboDept.ItemData(cboDept.ListCount - 1) = Val(NVL(rsReturn!ID))
                intIndex = cboDept.NewIndex
                cboDept.AddItem "�������ҡ�"
                cboDept.ItemData(cboDept.ListCount - 1) = 0
                cboDept.ListIndex = intIndex
            End If
            rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
            zlSelectDept = True
            Exit Function
        Else
            GoTo GoNotSel
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing

    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
    zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function

'*********************************************************************************************************************
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Function CheckSpecialAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByRef strPaits As String) As Boolean
'���ܣ���鲡���Ƿ�����ҪֻУ�ԣ������͵�����ҽ��
'       ��������ȼ�,����/Σҽ��,��ǰ����ҽ��������,��¼�����,,ת�ƣ���Ժ��תԺ������
'
'������strPaits=�Զ��ŷָ��Ĳ���ID��,���ش�����������ҽ���Ĳ���ID,��ҳID;
'      lng����ID,lng��ҳID=��������ʱ�Ŵ���
    Dim rsTmp As ADODB.Recordset, rsExists As ADODB.Recordset, strSQL As String, i As Long
    Dim strDepts As String, blnOnePati As Boolean, blnExists As Boolean
    
    strDepts = GetUser����IDs   '��ǰ������Ա���������������п���
    On Error GoTo errH
    If strPaits = "" Then
        strSQL = "Select A.����ID,A.��ҳID,a.�������,e.��������,a.ִ��Ƶ��" & vbNewLine & _
                "From ����ҽ����¼ A, ������ĿĿ¼ E" & vbNewLine & _
                "Where a.����id = [1] And a.��ҳid = [2]"
    Else
        strSQL = "Select/*+ Rule*/ Distinct A.����ID,A.��ҳID,a.�������,e.��������,a.ִ��Ƶ��" & vbNewLine & _
                "From ����ҽ����¼ A, ������ĿĿ¼ E,Table(f_Num2list([1])) B" & vbNewLine & _
                "Where a.����id = B.Column_value"
    End If
    strSQL = strSQL & " And a.������Ŀid = e.Id And a.ҽ��״̬ = 1" & vbNewLine & _
            " And (a.������� = 'H' And e.�������� = '1' And e.ִ��Ƶ�� = 2 Or a.������� = 'Z' And e.�������� In ('3', '4', '14', '5', '6', '9', '10', '11', '12') Or NVL(a.ִ��Ƶ��,'��')='��Ҫʱ' Or NVL(a.ִ��Ƶ��,'��')='��Ҫʱ')" & _
            " And Exists(" & _
            "Select M.���� From ��Ա�� M,ִҵ��� N" & _
            " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1)," & _
            "2,Substr(A.����ҽ��,1,Decode(Instr(A.����ҽ��,'/'),0,length(A.����ҽ��),Instr(A.����ҽ��,'/')-1))," & _
            "Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
            " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')" & _
            " )"
            
    If InStr(GetInsidePrivs(pסԺҽ������), "ȫԺҽ��У��") = 0 Then
        strSQL = strSQL & "  And (A.��������ID In (Select Column_Value From Table(f_Num2list([3]))) Or Exists(select 1 From ����ҽ����¼ Q,������ĿĿ¼ O where Q.����ID=A.����ID AND Q.��ҳID=A.��ҳID AND Q.������ĿID=O.ID AND Q.�������='Z'AND O.�������� ='7' AND Q.ִ�п���ID=A.��������ID And Q.ҽ��״̬=8)) "
    End If
    
    If strPaits = "" Then
        blnOnePati = True
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�������ҽ��", lng����ID, lng��ҳID, strDepts)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�������ҽ��", strPaits, 0, strDepts)
    End If
    
    strPaits = ""
    For i = 1 To rsTmp.RecordCount
        blnExists = False
        
        '5-��Ժ;6-תԺ,11-����,У��ʱֹͣ��������������ʱ�����˱��ΪԤ��Ժ
        '3-ת��;4,14-��ǰ����ҽ��
        If rsTmp!������� = "Z" And InStr(",3,4,14,5,6,11,", "," & rsTmp!�������� & ",") > 0 Or rsTmp!ִ��Ƶ�� & "" = "��Ҫʱ" Then
            blnExists = True
        Else
            strSQL = "Select 1" & vbNewLine & _
                    "From ����ҽ����¼ A, ������ĿĿ¼ E" & vbNewLine & _
                    "Where a.����id = [1] And a.��ҳid = [2] And a.������Ŀid = e.Id And a.ҽ��״̬ In (3,5,6,7)"
            '����ȼ�ҽ��
            If rsTmp!������� = "H" And rsTmp!�������� = "1" Then
                strSQL = strSQL & " And a.������� = 'H' And e.�������� = '1' And e.ִ��Ƶ�� = 2"
            ElseIf rsTmp!������� = "Z" Then
                If rsTmp!�������� = "9" Then    '����
                    strSQL = strSQL & " And a.������� = 'Z' And e.�������� = '10'"
                ElseIf rsTmp!�������� = "10" Then '��Σ
                    strSQL = strSQL & " And a.������� = 'Z' And e.�������� = '9'"
                ElseIf rsTmp!�������� = "12" Then '��¼�����
                    strSQL = strSQL & " And a.������� = 'Z' And e.�������� = '12'"
                End If
            End If
                    
            Set rsExists = zlDatabase.OpenSQLRecord(strSQL, "�������ҽ��", Val(rsTmp!����ID), Val("" & rsTmp!��ҳID))
            blnExists = rsExists.RecordCount > 0
        End If
        
        If blnExists Then
            If blnOnePati Then
                Exit For
            Else
                If InStr(strPaits & ";", ";" & rsTmp!����ID & "," & rsTmp!��ҳID & ";") = 0 Then
                    strPaits = strPaits & ";" & rsTmp!����ID & "," & rsTmp!��ҳID
                End If
            End If
        End If
        rsTmp.MoveNext
    Next
    
    If blnOnePati Then
        CheckSpecialAdvice = blnExists
    Else
        If strPaits = "" Then
            CheckSpecialAdvice = False
        Else
            strPaits = Mid(strPaits, 2)
            CheckSpecialAdvice = True
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckFeeItemAvailable(ByVal lngFeeItemID As Long, ByVal bytFlag As Byte) As Boolean
'����:����շ���Ŀ�Ƿ�δͣ��,���ҷ����ڲ���
'����:bytFlag:�������:1-����,2-סԺ
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From �շ���ĿĿ¼ Where ID = [1] And (����ʱ�� is Null Or ����ʱ�� > Sysdate) And ������� In (" & bytFlag & ",3)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItemID)
    CheckFeeItemAvailable = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    err = 0: On Error GoTo ErrHand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
ErrHand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("������������", , , True)) = 1
    err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
ErrHand:
End Function


Public Function AssembleImage(AssembleViewer As Object, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As Object

'���viewer�е���ʾ������ͼ���һ��ͼ��

    Dim Image As New DicomImage '��ͼ��
    Dim imgs As New DicomImages '��ʱ�洢��Ļ�ɼ���ͼ��
    Dim intWidth As Integer     '��ͼ��Ŀ��
    Dim intHeight As Integer    '��ͼ��ĸ߶�
    Dim Simg As New DicomImage
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '����ͼ���ռ�õ�������
    Dim intImgRectHeight As Integer '����ͼ���ռ�õ�����߶�
    Dim i As Integer
    Dim intMaxWidth As Integer      'ƴ�Ӻ�ͼ��������
    Dim intMaxHeight As Integer     'ƴ�Ӻ�ͼ������߶�
    Dim intBorder As Integer        'ͼ��֮��ı߾�
    Dim intOffsetX As Integer       'ƴ��ʱX�����λ��
    Dim intOffsetY As Integer       'ƴ��ʱY�����λ��
    Dim lngWhiteX As Long           '��ͼ���ɫ�ĳɰ�ɫ��X���
    Dim lngWhiteY As Long           '��ͼ���ɫ�ĳɰ�ɫ��Y�߶�
    
    If AssembleViewer.Count <= 0 Then
        '����һ����ͼ**************
        Exit Function
    End If

    On Error GoTo err
    '������ͼ��Ŀ�Ⱥ͸߶�

    '��ͼ��Ŀ�Ⱥ͸߶Ȳ��ܹ�����intMaxWidth��intMaxHeight����ȡ��߶ȣ�
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '������ͼ��Ŀ�Ⱥ͸߶�

    'ʹ��ԭͼ��Ŀ�Ⱥ͸߶Ⱥͣ�����Viewer�ı�����������

    '����ͼ����¿��
    For i = 1 To AssembleViewer.Count
        If intImgRectWidth < AssembleViewer(i).SizeX Then intImgRectWidth = AssembleViewer(i).SizeX
        If intImgRectHeight < AssembleViewer(i).SizeY Then intImgRectHeight = AssembleViewer(i).SizeY
    Next i
    
    '������������ͼ������
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows
    
    '����ͼ��Ŀ�ߣ����ܴ������ֵ
    '�������intMaxWidth��intMaxHeight�򣬰���ͼ���ܳ���ȣ�ʹ��С�ڵ���intMaxWidth��intMaxHeight��Ϊ�¿��,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '�ɼ�ͼ��
    '��ͼ��ɼ�����ʱͼ��
    For i = 1 To AssembleViewer.Count
        '�������ű��� hj�޸�,�����ͼ�ϲ�ʱ���Ŵ��ͼ���޷������Ŵ������
        sZoom = intImgRectHeight / AssembleViewer(i).SizeY
        If sZoom > intImgRectWidth / AssembleViewer(i).SizeX Then
            sZoom = intImgRectWidth / AssembleViewer(i).SizeX
        End If
        
        AssembleViewer(i).StretchToFit = False
        AssembleViewer(i).Zoom = sZoom
        '�ɼ�ͼ��
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '��ȷ������ͼ��Ŀ�Ⱥ͸߶�
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '������ͼ��
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT����MONOCHROME2,CR����MONOCHROME1��
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    'ƴ����ͼ��
    For i = 1 To imgs.Count
        '����ͼ����λ��
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set AssembleImage = Image
    Exit Function
err:
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'���ܣ����������ͼ��������ͼ������Ŀ�Ⱥ͸߶ȣ�������ѵ�ͼ����������������
'������ ImageCount����ͼ������
'       RegionWidth--ͼ����ʾ����Ŀ��
'       RegionHeight--ͼ����ʾ����ĸ߶�
'       Rows����[����]�������
'       Cols����[����]�������
'���أ������������Rows���������Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    Rows = iRows: Cols = iCols
    
    If ImageCount <> 0 Then
        If Rows * Cols > ImageCount Then
            iBase = 6
            blnDoLoop = True
            
            While blnDoLoop
                iBase = iBase - 1
                
                If ImageCount Mod iBase = 0 Then
                    blnDoLoop = False
                End If
            Wend
        

            If RegionWidth > RegionHeight Then
                If ImageCount / iBase > iBase Then
                    Cols = ImageCount / iBase
                    Rows = iBase
                Else
                    Rows = ImageCount / iBase
                    Cols = iBase
                End If
            Else
                If ImageCount / iBase > iBase Then
                    Cols = iBase
                    Rows = ImageCount / iBase
                Else
                    Rows = iBase
                    Cols = ImageCount / iBase
                End If
            End If
        End If
    End If
err:
End Sub

Public Function GetRPTPicture(ByVal blnMoved As Boolean, ByVal lngReportID As Long, ByVal strRPTNO As String, ByRef aryPrintPara As Variant) As String
'���ܣ���ȡ����ͼƬ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strPicPath As String
    Dim intPCount As Integer
    Dim oPicture As StdPicture
    Dim i As Integer, j As Integer, intParaCount As Integer
    Dim strPicFile As String
    Dim aryPara(19) As String, aryFlagPara(1) As String     '����ͼ�е�ͼ���¼
    Dim strFlagString As String 'ʵ�ʴ����Զ��屨�������
    Dim int��ʽ�� As Integer
    Dim intRows As Integer, intCols As Integer
    Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
    
    Const cprET_��������� = 3
    Const EPRMarkedPicture = 1
    
    Dim cTable As Object, objFile As Object
    
    
    On Error GoTo err
    
    Set cTable = CreateObject("zlRichEPR.cEPRTable")
    If cTable Is Nothing Then Exit Function
    Set objFile = CreateObject("Scripting.FileSystemObject")
        
    '��ȡͼ��
    strPicPath = App.Path & "\TmpImage\"
    If objFile.FolderExists(strPicPath) = False Then objFile.CreateFolder strPicPath
    
    '��ȡ����ͼ�񣨰������ͼ�����ɱ����ļ�
    'һ���������п������ж������ͼ
    intPCount = 0
    strSQL = "Select Id As ���Id From ���Ӳ�������" & vbNewLine & _
        "       Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By �������"
    If blnMoved = True Then strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ��", lngReportID)
    Do While Not rsTmp.EOF
        If cTable.GetTableFromDB(cprET_���������, lngReportID, Val("" & rsTmp!���ID)) Then
            For i = 1 To cTable.Pictures.Count
                strPicFile = "PACSPic" & i & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                    Set oPicture = cTable.Pictures(i).DrawFinalPic
                Else
                    Set oPicture = cTable.Pictures(i).OrigPic
                End If
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '������ͼ��ͼ���·��
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        aryFlagPara(0) = strPicFile
                    Else
                        aryPara(intPCount) = strPicFile
                        dcmImages.AddNew
                        dcmImages(dcmImages.Count).FileImport strPicFile, "BMP"
                        intPCount = intPCount + 1
                        If intPCount > UBound(aryPara) Then Exit Do
                    End If
                End If
            Next i
        End If
        rsTmp.MoveNext
    Loop
    
    '����ѡ����Զ��屨���ʽ����֯ͼ��
    '����һ�ֱ����ʽ����
    int��ʽ�� = 1
    strSQL = "Select b.����,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
    "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = [2] And b.���� not like '���%'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�Ƿ���Ҫ���ͼ��", strRPTNO, int��ʽ��)
    If rsTmp.RecordCount = 1 And intPCount >= 1 Then
        '���ͼ��
        ResizeRegion intPCount, rsTmp("W"), rsTmp("H"), intRows, intCols
        Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTmp("H"), rsTmp("W"))
        dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
    End If
    
    
    '��ȡͼ�񣬵��ñ���
    intPCount = 0       '��¼ͼ�������
    strSQL = "Select b.���� From zlReports a, zlRptItems b" & vbNewLine & _
    "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = [2]" & vbNewLine & _
    "       Order By b.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ���", strRPTNO, int��ʽ��)
    'װ��ͼ������
    intParaCount = 0
    Do While Not rsTmp.EOF
        
        '�ֱ�װ�ڱ��ͼ�ͱ���ͼ
        If InStr(rsTmp!����, "���") <> 0 Then '���ͼ
            If aryFlagPara(0) <> "" Then strFlagString = rsTmp!���� & "=" & aryFlagPara(0)
        Else    '����ͼ
            If intPCount > UBound(aryPara) Then Exit Do     'ͼ���������������е�ͼ���˳�
            If aryPara(intPCount) = "" Then Exit Do         '�����е�ͼ���ȱ����еĶ࣬�˳�
            
            aryPrintPara(intParaCount) = rsTmp!���� & "=" & aryPara(intPCount)
            intPCount = intPCount + 1
            intParaCount = intParaCount + 1
        End If
        rsTmp.MoveNext
    Loop
    
    '��������ͼ�αȱ������ٵ����
    For j = intParaCount To UBound(aryPrintPara)
        If aryPrintPara(j) Like "*=*" Then aryPrintPara(j) = ""
    Next
    
    GetRPTPicture = strFlagString
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetAdviceRs() As ADODB.Recordset
'���ܣ�����һ������ҽ����¼�����ֶεı��ؼ�¼������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select a.*,c.����ID,0 as EditState From ����ҽ����¼ A,�������ҽ�� B,������ϼ�¼ C Where A.id=b.ҽ��id and c.id=b.���id And a.ID=0 And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceRS")
    
    Set GetAdviceRs = zlDatabase.zlCopyDataStructure(rsTmp)
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckUnExecutedAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngSpecialAdviceID As Long, ByVal intBabyNum As Integer) As String
'���ܣ����δУ�Ե�ҽ������δ���͵�������
'������strSpecialAdviceIDs=����ҽ��ID(���ʱҪ�ų���ǰ����ҽ��ID)
'      intBabyNum=Ӥ����š�
'���أ���ʾ��Ϣ
'˵�����ú���ִ��ʱֻ����ĸ������ҽ��ʱ����Զ�Ӥ��ҽ���ļ�飬���ͬʱ������������
'      �洢����  Zl_���˱䶯��¼_Change �� Zl_���˱䶯��¼_Preout Ҳ�����Ƶļ�飬������Ӧ�ĵ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    'δУ��ҽ��
    strSQL = "Select 1 From ����ҽ����¼ Where ����id = [1] And ��ҳid = [2] And ҽ��״̬ = 1 And ID<>[3] And Nvl(Ӥ��, 0) = [4] And Rownum = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ҽ��", lng����ID, lng��ҳID, lngSpecialAdviceID, intBabyNum)
    If rsTmp.RecordCount > 0 Then strMsg = "δУ�Ե�ҽ��"
    
    'δ���͵�����
    strSQL = "Select 1" & vbNewLine & _
        "From ����ҽ����¼" & vbNewLine & _
        "Where ����id = [1] And ��ҳid = [2] And ҽ����Ч = 1 And ҽ��״̬ In (2, 3) And Nvl(ִ�б��, 0) <> -1 And ID <> [3] And Nvl(Ӥ��, 0) = [4] And" & vbNewLine & _
        " Nvl(ִ������,0)<>0 and Rownum = 1"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ҽ��", lng����ID, lng��ҳID, lngSpecialAdviceID, intBabyNum)
    If rsTmp.RecordCount > 0 Then
        If strMsg <> "" Then
            strMsg = strMsg & "��δ���͵�����"
        Else
            strMsg = "δ���͵�����"
        End If
    End If
        
    strSQL = "Select 1 From ����ҽ����¼ a Where a.����id = [1] And a.��ҳid = [2] And a.ҽ����Ч = 0 And a.ҽ��״̬=8 " & _
            "And Exists(Select 1 From ����ҽ����¼ b Where b.id = [3] And a.ִ����ֹʱ�� > b.��ʼִ��ʱ��) And Rownum = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ҽ��", lng����ID, lng��ҳID, lngSpecialAdviceID)
    If rsTmp.RecordCount > 0 Then
        If strMsg <> "" Then
            strMsg = strMsg & "��δ��ֹͣʱ��ĳ���"
        Else
            strMsg = "δ��ֹͣʱ��ĳ���"
        End If
    End If
    
    
    CheckUnExecutedAdvice = strMsg
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function getChargeMode(ByVal intChargeMode As Integer) As String
    getChargeMode = Decode(intChargeMode, 0, "������ȡ", 1, "�����Թܷ���", 2, "һ�η���ֻ��ȡһ��", 3, "����ֻ��ȡһ��", 4, _
        "����δִ����ȡһ��", 5, "����ֻ��ȡһ��,�ų�������Ŀ", 6, "����δִ����ȡһ��,�ų�������Ŀ", 7, "ÿ���״β���ȡ", 9, "�Զ���")
End Function

Public Sub Load��Һ����(ByRef cbo���� As ComboBox, ByRef lbl���ٵ�λ As Label, ByVal blnTurn As Boolean, Optional ByVal bln��Ѫ As Boolean = False)
'���ܣ����ص��������б�
'������blnTurn=�Ƿ�ת����Һ��λ
    Dim i As Long
    Dim arrTmp() As String
    
    If bln��Ѫ = False Then
        If blnTurn Then
            If lbl���ٵ�λ.Caption = "����/Сʱ" Then
                lbl���ٵ�λ.Caption = "��/����"
            Else
                lbl���ٵ�λ.Caption = "����/Сʱ"
            End If
        End If
        If lbl���ٵ�λ.Caption = "��/����" Then
            arrTmp = Split("20,30,40,50,60,70,80,20-40,40-60", ",")
        Else
            arrTmp = Split("60,120,180,240,300,600,120-240", ",")
        End If
    Else
        arrTmp = Split("15,30,����,��ѹ", ",")
    End If
    
    cbo����.Clear
    For i = 0 To UBound(arrTmp)
        cbo����.AddItem arrTmp(i)
    Next
    
End Sub

Public Function GetKSSAuditQuestion(ByVal lngҽ��ID As Long) As String
'���ܣ���ȡ������ҩ���δͨ���ķ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(����˵��,'��') as ����˵�� From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=12 Order by ����ʱ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
    If Not rsTmp.EOF Then
        GetKSSAuditQuestion = rsTmp!����˵��
    Else
        If gblnѪ��ϵͳ Then
            'Ѫ�⽫��Ѫҽ������Ϊ��˲�ͨ���󣬲�������Ϊ16
            strSQL = "Select Nvl(����˵��,'��') as ����˵�� From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=16 Order by ����ʱ�� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
            If Not rsTmp.EOF Then GetKSSAuditQuestion = rsTmp!����˵��
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetMaxAdviceNO(ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal lngӤ�� As Long, Optional ByVal str�Һŵ� As String) As Long
'���ܣ���ȡ��ǰ���˵����ҽ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If str�Һŵ� <> "" Then
        strSQL = "Select nvl(Max(���),1) as ��� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[4] And Nvl(Ӥ��,0)=[3]"
    Else
        If lng��ҳID = 0 Then
            strSQL = "Select Nvl(Max(���),1) as ��� From ����ҽ����¼ Where ����ID=[1] And ��ҳID Is Null"
        Else
            strSQL = "Select Nvl(Max(���),1) as ��� From ����ҽ����¼ Where ����ID=[1] And ��ҳID=[2] And Nvl(Ӥ��,0)=[3]"
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID, lngӤ��, str�Һŵ�)
    If Not rsTmp.EOF Then GetMaxAdviceNO = rsTmp!���

    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
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
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "": blnҽ�� = False: bln���� = False
    If grsҽ�Ƹ��ʽ Is Nothing Then
        strSQL = "Select ����,����,����,ȱʡ��־,�Ƿ�ҽ��,�Ƿ񹫷� From ҽ�Ƹ��ʽ"
    ElseIf grsҽ�Ƹ��ʽ.State <> 1 Then
        strSQL = "Select ����,����,����,ȱʡ��־,�Ƿ�ҽ��,�Ƿ񹫷� From ҽ�Ƹ��ʽ"
    End If
    If strSQL <> "" Then
        Set grsҽ�Ƹ��ʽ = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ�Ƹ��ʽ")
    End If
    grsҽ�Ƹ��ʽ.Find "����='" & strҽ�Ƹ������� & "'", , adSearchForward, 1
    If grsҽ�Ƹ��ʽ.EOF Then Exit Function
    blnҽ�� = Val(NVL(grsҽ�Ƹ��ʽ!�Ƿ�ҽ��)) = 1
    bln���� = Val(NVL(grsҽ�Ƹ��ʽ!�Ƿ񹫷�)) = 1
    zlIsCheckMedicinePayMode = blnҽ�� Or bln����
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check��鲿λEnable(ByVal lng��Ŀid As Long, ByVal str��λ As String, ByVal str�Ա� As String, Optional ByVal str���� As String, Optional ByRef blnExists As Boolean) As Boolean
'���ܣ���鲿λ�Ƿ�������ָ�����Ա�
'������blnExists �ж��Ƿ���������鲿λ�򷽷�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(b.�����Ա�, 0) as �����Ա�" & vbNewLine & _
            "From ������Ŀ��λ A, ���Ƽ�鲿λ B, ������ĿĿ¼ C" & vbNewLine & _
            "Where a.���� = b.���� And a.��λ = b.���� And a.���� = c.�������� And a.��Ŀid = c.Id And" & vbNewLine & _
            "       c.Id = [1] And a.��λ = [2] And Replace(A.����,chr(9),'')=[3]"
    blnExists = True
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid, str��λ, str����)
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "�����Ա�=" & IIF(str�Ա� = "��", 1, IIF(str�Ա� = "Ů", 2, 0)) & " Or �����Ա�=0"
        Check��鲿λEnable = rsTmp.RecordCount > 0
    Else
        blnExists = False
    End If
        
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckStopedUnAffirm(ByVal strPatis As String, ByRef strPatisName As String) As Boolean
'���ܣ����ָ���Ĳ����Ƿ������ֹͣ��δȷ��ֹͣ��ҽ��
'������strPatis=����ID1:��ҳID,����ID2:��ҳID,......
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    If InStr(strPatis, ",") = 0 Then
        strSQL = "Select a.����" & vbNewLine & _
            "From ����ҽ����¼ A" & vbNewLine & _
            "Where a.����id = [1] And a.��ҳid = [2] And Exists (Select 1 From ����ҽ��״̬ B Where a.Id = b.ҽ��id And b.�������� = 8 And b.ǩ��id Is Not Null)" & _
            " And a.ҽ��״̬ = 8 And a.ҽ����Ч = 0 And Nvl(a.Ӥ��, 0) = 0 And Nvl(A.ִ�б��,0)<>-1 And Rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", Val(Split(strPatis, ":")(0)), Val(Split(strPatis, ":")(1)))
    Else
        strSQL = "Select /*+ leading(b) use_nl(a)*/Distinct a.����" & vbNewLine & _
                "From ����ҽ����¼ A, Table(f_Num2list2([1])) B" & vbNewLine & _
                "Where a.����id = b.C1 And a.��ҳid = b.C2 And a.ҽ��״̬ = 8 And Nvl(A.ִ�б��,0)<>-1 " & _
                " And Exists (Select 1 From ����ҽ��״̬ C Where a.Id = C.ҽ��id And C.�������� = 8 And C.ǩ��id Is Not Null) " & _
                " And a.ҽ����Ч = 0 And Nvl(a.Ӥ��, 0) = 0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strPatis)
    End If

    CheckStopedUnAffirm = rsTmp.RecordCount > 0
    For i = 1 To rsTmp.RecordCount
        strPatisName = strPatisName & "," & rsTmp!����
        rsTmp.MoveNext
    Next
    If strPatisName <> "" Then strPatisName = Mid(strPatisName, 2)

    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetRsRedoDate(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Date
'���ܣ���ȡ����ҽ���������ʱ��
    Dim strSQL As String, rsTmp As ADODB.Recordset
 
    strSQL = "Select ҽ������ʱ�� as ʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", lng����ID, lng��ҳID)
    
    GetRsRedoDate = NVL(rsTmp!ʱ��, CDate("1900-01-01"))
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check��������(frmParent As Object, ByVal txtҽ������ As Object, ByVal lng����ID As Long, ByVal lngҩ��ID As Long, ByVal str���� As String, ByVal bln�Զ�Ƥ�� As Boolean, Optional ByRef lngƤ��ID As Long, _
    Optional ByVal lng��ҳID As Long, Optional ByVal lngҩƷID As Long, Optional ByRef bln������ҩ As Boolean, Optional ByVal str��ʼʱ�� As String, Optional ByRef bln���Խ�ʾ As Boolean) As String
'���ܣ��������ҩ���г�ҩ�Ĺ������飬�����סԺҽ���´﹫��
'������frmParent��ǰ���õĴ���
'      txtҽ������ �ؼ���ҽ���༭������ı������ؼ�����
'      lngҩ��ID=ҩƷ������ĿID
'      str����=ҩƷ����,������ʾ
'      bln������ҩ bln�ж�������ҩ =true ����Ҫ�жϲ�������� bln������ҩ ���أ�������סԺ��
'      str��ʼʱ�� ҽ���Ŀ�ʼִ��ʱ��
'      bln���Խ�ʾ ҩƷƷ�ֻ��߹���Ӧ�Ĺ���ʵ����Ŀ��Ϊ ����������ʹ�ã����أ����ڽ����жϣ�
'���أ�Ϊ�ձ�ʾͨ��
'      lngƤ��ID=Ҫ�Զ���ӵ�Ƥ����ĿID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dat��ʼʱ�� As Date
    Dim lngTmpƤ��ID As Long
    Dim strƤ������ As String
    
    Dim bln��� As Boolean
    Dim rsƤ�� As ADODB.Recordset
    Dim blnTmp���Խ�ʾ As Boolean
    
    On Error GoTo errH
    
    lngƤ��ID = 0
    bln������ҩ = False
    bln���Խ�ʾ = False
    
    '�жϵ�ǰҩƷ�ǲ��ǰ���Ƥ����Ŀ�����жϹ�����ж�Ʒ��
    If lngҩƷID <> 0 Then
        strSQL = "Select A.�÷�ID,B.����,B.ִ�б�� From ҩƷ�÷����� A,������ĿĿ¼ B Where A.�÷�ID=B.ID And A.����=0 And A.ҩƷID=[1] And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) order by b.����"
        Set rsƤ�� = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lngҩƷID)
        If Not rsƤ��.EOF Then bln��� = True
    End If
    
    If Not bln��� Then
        strSQL = "Select A.�÷�ID,B.����,B.ִ�б�� From �����÷����� A,������ĿĿ¼ B Where A.�÷�ID=B.ID And A.����=0 And A.��ĿID=[1] And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) order by b.����"
        Set rsƤ�� = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lngҩ��ID)
    End If
    
    
    '��ҪƤ��
    If Not rsƤ��.EOF Then
        '���ֻ��һ��Ƥ����Ŀ�����¼����
        If rsƤ��.RecordCount = 1 Then
            lngTmpƤ��ID = rsƤ��!�÷�ID
            strƤ������ = rsƤ��!���� & ""
        End If
        
        rsƤ��.Filter = "ִ�б��=2"
        blnTmp���Խ�ʾ = Not rsƤ��.EOF
        
        If str��ʼʱ�� <> "" Then
            dat��ʼʱ�� = CDate(str��ʼʱ��)
        Else
            dat��ʼʱ�� = zlDatabase.Currentdate
        End If
        'ȡ��Чʱ���ڵ����һ�ι�������Ǽ�
        strSQL = "Select ҩ����,���,Nvl(����ʱ��,��¼ʱ��) as ����ʱ�� From ���˹�����¼" & _
            " Where ����ID=[1] And ҩ��ID=[2] And [3]+Nvl(����ʱ��,��¼ʱ��)>=[4]" & _
            " Order by Nvl(����ʱ��,��¼ʱ��) Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lng����ID, lngҩ��ID, gint�����Ǽ���Ч����, dat��ʼʱ��)
        If Not rsTmp.EOF Then
            '�й�������ǼǼ�¼,�����Ƿ����Ծ����Ƿ���ʾ
            If NVL(rsTmp!���, 0) = 1 Then
                If blnTmp���Խ�ʾ Then
                    bln���Խ�ʾ = True
                    strMsg = "�ò�����" & Format(rsTmp!����ʱ��, "M��d��") & "�Ĺ���ʵ���ж�""" & NVL(rsTmp!ҩ����, str����) & """����(+)��" & _
                        vbCrLf & vbCrLf & "����Ŀ�������˲��ܽ�������ʹ�ã��ʽ�ֹʹ�á�"
                Else
                    strMsg = "�ò�����" & Format(rsTmp!����ʱ��, "M��d��") & "�Ĺ���ʵ���ж�""" & NVL(rsTmp!ҩ����, str����) & """����(+)��" & _
                        vbCrLf & vbCrLf & "�Ƿ���Ȼʹ�ø�ҩƷ��"
                End If
            Else
                strMsg = "" 'Ϊ����,ͨ��
            End If
        Else
            If lng��ҳID <> 0 Then
                '���������´����ж� ҩƷid,����Ʒ���ж�75874
                strSQL = "select 1 from ����ҽ����¼ a where a.����id=[1] and a.��ҳid=[2]" & _
                    IIF(lngҩƷID <> 0, " and a.�շ�ϸĿid=[3]", " and a.������Ŀid=[3]") & _
                    " And [4]+a.�ϴ�ִ��ʱ��>=[5] And Rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lng����ID, lng��ҳID, IIF(0 <> lngҩƷID, lngҩƷID, lngҩ��ID), gint�����Ǽ���Ч����, dat��ʼʱ��)
                bln������ҩ = Not rsTmp.EOF
            End If
            
            If Not bln������ҩ Then
                If bln�Զ�Ƥ�� Then
                    '���⣺31144,����Ƕ��Ƥ����Ŀ���򵯳�ѡ���������û�ѡ��һ��Ƥ�ԡ�
                    If lngTmpƤ��ID = 0 Then
                        vRect = zlControl.GetControlRect(txtҽ������.hwnd)
                        
                        If bln��� Then
                            strSQL = "Select A.�÷�ID as ID,B.���� as ��ѡ��һ��Ƥ��" & _
                                    " From ҩƷ�÷����� A,������ĿĿ¼ B" & _
                                    " Where A.�÷�ID=B.ID And A.����=0 And A.ҩƷID=[1]" & _
                                    " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) order by b.����"
                            Set rsTmp = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "Ƥ��ҽ��ѡ��", False, "", "", False, False, True, _
                                vRect.Left, vRect.Top, txtҽ������.Height, blnCancel, False, True, lngҩƷID)
                        Else
                            strSQL = "Select A.�÷�ID as ID,B.���� as ��ѡ��һ��Ƥ��" & _
                                    " From �����÷����� A,������ĿĿ¼ B" & _
                                    " Where A.�÷�ID=B.ID And A.����=0 And A.��ĿID=[1]" & _
                                    " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) order by b.����"
                            Set rsTmp = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "Ƥ��ҽ��ѡ��", False, "", "", False, False, True, _
                                vRect.Left, vRect.Top, txtҽ������.Height, blnCancel, False, True, lngҩ��ID)
                        End If
                            
                        If Not blnCancel Then
                            lngƤ��ID = rsTmp!ID
                            strMsg = "" '�Զ����,����ʾ
                        Else
                            strMsg = "�ڶԲ���ʹ��""" & str���� & """ǰ��Ҫ���Ƚ��й������ԣ�" & vbCrLf & _
                                    "�����ղ�û��ѡ���Ӧ�������ԣ��Ƿ���Ȼʹ�ø�ҩƷ��"
                        End If
                    Else
                        lngƤ��ID = lngTmpƤ��ID
                        strMsg = "" '�Զ����,����ʾ
                    End If
                Else
                    'Ҫ��Ƥ��,����ʾƤ��
                    strMsg = "�ڶԲ���ʹ��""" & str���� & """ǰ��Ҫ���Ƚ���""" & strƤ������ & """��" & vbCrLf & _
                        "��û�з�����Ч�Ĺ������������Ƿ���Ȼʹ�ø�ҩƷ��"
                End If
            End If
        End If
    End If
    Check�������� = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNoneSendID(ByVal lng����ID As Long, ByVal str��ʶ As String, ByVal bytType As Byte, Optional ByVal bln������ҺҩƷ As Boolean, Optional ByVal lng�Һ�ID As Long, Optional ByRef strAdviceDrugIDs As String) As String
'���ܣ���ȡ��ΪƤ�Թ���(����(+))����Ƥ�Խ���������͵�ҽ��ID
'������lng���� ����ID
'      str��ʶ ������������ǹҺŵ���סԺ������ҳid
'      bytType 1���2סԺ
'      bln������ҺҩƷ =��ҺҩƷ����ҳ�����
'      strAdviceDrugIDs  ������   �����Ƶ�ҩƷ�е�ҽ��IDs
    Dim rsTmp As ADODB.Recordset
    Dim rsDrug As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    Dim str�÷�IDs As String
    Dim strPatiFilter As String
    Dim strPatiFilterAnd As String
    Dim bln�Զ�Ƥ�� As Boolean
    Dim strOKҽ��IDs As String
    Dim strNOҽ��IDs As String
    Dim lng��ҳID As Long
    Dim strDrugIDs As String 'Ҫ���ſ���ҩƷҽ��ID
    
    strAdviceDrugIDs = ""
    
    bln�Զ�Ƥ�� = Val(zlDatabase.GetPara("ҽ������Ƥ������", glngSys, IIF(bytType = 1, p����ҽ���´�, pסԺҽ���´�))) <> 0
    
    If Not bln�Զ�Ƥ�� Then Exit Function
    
    If bytType = 2 And bln������ҺҩƷ = False Then
        If Val(zlDatabase.GetPara("����Ƥ�Խ������ҽ����������", glngSys, pסԺҽ���´�)) = 1 Then Exit Function
    End If
    
    strPatiFilter = IIF(bytType = 1, " And A.����ID+0=[1] And  A.�Һŵ�=[2]", " And A.����ID=[1] And  A.��ҳID=[2]")
    strPatiFilterAnd = IIF(bytType = 1, " And A.����ID+0=B.����ID And  A.�Һŵ�=B.�Һŵ�", " And A.����ID=B.����ID And  A.��ҳID=B.��ҳID")
    
    On Error GoTo errH
    
    'ȡ����ҪƤ�Ե�ҩƷҽ��
    strSQL = "Select a.���id,a.id,b.�÷�id,a.������ĿID,a.��ʼִ��ʱ��,a.��ʼִ��ʱ�� as ����ʱ�� From ����ҽ����¼ A,�����÷����� B" & _
        " Where a.������Ŀid = b.��Ŀid And b.����=0 and A.������� IN('5','6') And a.ҽ��״̬<>4 and a.Ƥ������˵�� is null and nvl(a.Ƥ�Խ��,'��')<>'������ҩ'" & strPatiFilter & _
        " union all " & _
        " Select a.���id,a.id,b.�÷�id,a.������ĿID,a.��ʼִ��ʱ��,a.��ʼִ��ʱ�� as ����ʱ�� From ����ҽ����¼ A,ҩƷ�÷����� B" & _
        " Where a.�շ�ϸĿid = b.ҩƷid And b.����=0 and A.������� IN('5','6') And a.ҽ��״̬<>4 and a.Ƥ������˵�� is null and nvl(a.Ƥ�Խ��,'��')<>'������ҩ'" & strPatiFilter
    strSQL = "Select a.���id,a.id,a.�÷�id,a.������ĿID,a.��ʼִ��ʱ��,a.����ʱ��,1 as ��� from (" & strSQL & ") a group by a.���id,a.id,a.�÷�id,a.������ĿID,a.��ʼִ��ʱ��,a.����ʱ�� order by a.���id"
    Set rsDrug = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, IIF(bytType = 1, str��ʶ, Val(str��ʶ)))
    
    If Not rsDrug.EOF Then
        If bytType = 1 Then
            lng��ҳID = lng�Һ�ID
        Else
            lng��ҳID = Val(str��ʶ)
        End If
        strSQL = "Select a.ҩ��id,[3]+a.����ʱ�� as ����ʱ��, Nvl(a.���,0) as ��� From ���˹�����¼ A " & _
            " Where a.��¼��Դ = 2 And a.����ID=[1] And a.��ҳID=[2] And a.����ʱ�� = (Select Max(x.����ʱ��) From ���˹�����¼ X  Where x.����id = a.����id And x.��ҳid = a.��ҳid And x.ҩ��id = a.ҩ��id)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID, gint�����Ǽ���Ч����)
        If Not rsTmp.EOF Then
            '�����ݽ���ɸѡ
            Set rsDrug = zlDatabase.CopyNewRec(rsDrug)
            For i = 1 To rsDrug.RecordCount
                rsTmp.Filter = "ҩ��id=" & rsDrug!������ĿID
                If Not rsTmp.EOF Then
                    rsDrug!����ʱ�� = rsTmp!����ʱ��
                    rsDrug!��� = rsTmp!���
                End If
                rsDrug.MoveNext
            Next
            rsDrug.MoveFirst
            For i = 1 To rsDrug.RecordCount
                If rsDrug!����ʱ�� < rsDrug!��ʼִ��ʱ�� Or Val(rsDrug!��� & "") = 1 Then
                    If InStr("," & strNOҽ��IDs & ",", "," & rsDrug!���ID & ",") = 0 Then
                        strNOҽ��IDs = strNOҽ��IDs & "," & rsDrug!���ID
                    End If
                    strDrugIDs = strDrugIDs & "," & rsDrug!ID
                End If
                rsDrug.MoveNext
            Next
        Else
            '�޹����������ȫ���ų���
            For i = 1 To rsDrug.RecordCount
                If InStr("," & strNOҽ��IDs & ",", "," & rsDrug!���ID & ",") = 0 Then
                    strNOҽ��IDs = strNOҽ��IDs & "," & rsDrug!���ID
                End If
                strDrugIDs = strDrugIDs & "," & rsDrug!ID
                rsDrug.MoveNext
            Next
        End If
    End If
    
    If strNOҽ��IDs <> "" Then
        'strAdviceDrugIDs    ҩƷ�е�ҽ��IDs
        strAdviceDrugIDs = Mid(strDrugIDs, 2)
        strSQL = "Select a.ID From ����ҽ����¼ a Where a.������� IN ('5','6','E') And  Instr([3],','||Nvl(a.���ID,ID)||',')>0 " & strPatiFilter
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, IIF(bytType = 1, str��ʶ, Val(str��ʶ)), "," & strNOҽ��IDs & ",")
        strSQL = ""
        Do While Not rsTmp.EOF
            strSQL = strSQL & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        GetNoneSendID = Mid(strSQL, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckApplication(ByVal lngID As Long, ByVal int���� As Integer) As Boolean
'���ܣ����������Ŀ�Ƿ��Ӧ�����븽��
'������lngID������ĿID��int���� 1-���2-סԺ
'���أ�true=��Ӧ�����븽��
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "Select 1 From ��������Ӧ�� A, �������ݸ��� B Where a.�����ļ�id = b.�ļ�id And a.������Ŀid = [1] And Ӧ�ó��� = [2] And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckApplication", lngID, int����)
    CheckApplication = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPrice(ByVal lngPatId As Long, ByVal lngPageId As Long, _
    ByVal lng������ĿID As Long, ByVal str���ַ������ As String, ByVal lngִ������ As Long, _
    ByVal lng������Դ As Long, ByVal lngִ�п���ID As Long) As Double
'���ܣ���ȡ�����Ŀ�ķ��úϼ�
'���������ַ�����ϸ�ʽ����λ;����1,����2|��λ2;����1,����3
'     ִ�����ͣ�1���棬2���ԣ�3����
'    ������Դ��1-���2סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim strSQL��λ As String
    Dim str���ַ��� As String
    Dim str��λ As String
    Dim str���� As String
    Dim objExp As Object
    Dim strGrad As String
    Dim strTmp As String
    
    If gobjPublicExpense Is Nothing Then
        Set gobjPublicExpense = GetObject("", "zlPublicExpense.clsPublicExpense")
        If objExp Is Nothing Then Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
    End If
    
    '��ȡ���õȼ�
    If gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lngPatId, lngPageId, "", strTmp, strTmp, strGrad) = False Then
        strGrad = ""
    End If
    
    strSQL��λ = " Select '' as �걾��λ, '' as ��鷽�� From Dual "
    
    If Not Trim(str���ַ������) = "" Then
        For i = 0 To UBound(Split(str���ַ������, "|"))
            str���ַ��� = Split(str���ַ������, "|")(i)
            str��λ = Split(str���ַ���, ";")(0)
            str���� = Split(str���ַ���, ";")(1)
            For j = 0 To UBound(Split(str����, ","))
                strSQL��λ = strSQL��λ & " Union All " & _
                    "Select '" & str��λ & "','" & Split(str����, ",")(j) & "' From Dual "
            Next
        Next
    End If
    
    strSQL = "Select Sum(b.�շ����� * d.�ּ�) As �ϼ�" & vbNewLine & _
            "From (Select *" & vbNewLine & _
            "       From (Select c.������Ŀid, c.�շ���Ŀid, c.��鲿λ, c.��鷽��, c.��������, c.�շ�����, c.���ж���, c.������Ŀ, c.�շѷ�ʽ, c.���ÿ���id," & vbNewLine & _
            "                     Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & vbNewLine & _
            "              From �����շѹ�ϵ C," & vbNewLine & _
            "                   (" & strSQL��λ & ") A, ������ĿĿ¼ d " & vbNewLine & _
            "              Where c.������Ŀid = D.ID And d.�Ƽ����� =0 and c.������Ŀid = [1] And (a.�걾��λ Is Null And [2] In (1, 2) And c.�������� = 1 Or" & vbNewLine & _
            "                    a.�걾��λ = c.��鲿λ And a.��鷽�� = c.��鷽�� And Nvl(c.��������, 0) = 0 Or" & vbNewLine & _
            "                    a.��鷽�� Is Null And Nvl(c.��������, 0) = 0 And c.��鲿λ Is Null And c.��鷽�� Is Null) And" & vbNewLine & _
            "                    (c.���ÿ���id Is Null Or c.���ÿ���id = [3] And c.������Դ = [4]))" & vbNewLine & _
            "       Where Nvl(���ÿ���id, 0) = Top) B, �շ���ĿĿ¼ C, �շѼ�Ŀ D" & vbNewLine & _
            "Where b.�շ���Ŀid = c.Id And b.�շ���Ŀid = d.�շ�ϸĿid And" & vbNewLine & _
           IIF(strGrad = "", _
            "       D.�۸�ȼ� Is Null ", _
            "      ((instr( ';4;5;6;7;', ';' || C.��� || ';')=0 And D.�۸�ȼ�=[5]) " & _
                    " Or (D.�۸�ȼ� Is Null And Not Exists(Select 1 From �շѼ�Ŀ " & _
                                    " Where b.�շ���ĿId=�շ�ϸĿID And Sysdate Between d.ִ������ And d.��ֹ���� And" & _
                                    " (instr( ';4;5;6;7;', ';' || C.��� || ';')=0 And �۸�ȼ�=[5]) )))") & " And " & vbNewLine & _
            "      ((Sysdate Between d.ִ������ And d.��ֹ����) Or (Sysdate >= d.ִ������ And d.��ֹ���� Is Null)) And" & vbNewLine & _
            "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.������� In ([4], 3) And" & vbNewLine & _
            "      (c.վ�� = '0' Or c.վ�� Is Null)"

    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPrice", lng������ĿID, lngִ������, lngִ�п���ID, lng������Դ, strGrad)
    If rsTmp.RecordCount > 0 Then GetPrice = Val(rsTmp!�ϼ� & "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceStopTime(ByVal lngҽ��ID As Long) As String
'���ܣ�������д��ִ�еǼǵ�ҽ����ȷ��ֹͣʱ��������һ�ε�Ҫ��ִ��ʱ��
'���أ������ֹͣʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select To_Char(Max(Ҫ��ʱ��), 'YYYY-MM-DD HH24:MI') As ִ��ʱ�� From ����ҽ��ִ�� Where ҽ��id = [1]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
    If Not rsTmp.EOF Then GetAdviceStopTime = "" & rsTmp!ִ��ʱ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CreateScript(Optional ByRef objVBA As Object, Optional ByRef objScript As clsScript) As Boolean
'���ܣ�����Script��VBA����
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    err.Clear: On Error GoTo 0
    
    If Not objVBA Is Nothing Then
        objVBA.Language = "VBScript"
        Set objScript = New clsScript
        objVBA.AddObject "clsScript", objScript, True
        CreateScript = True
    End If
End Function

Public Function GetAdviceDiag(ByVal lngҽ��ID As Long, Optional ByRef str��� As String) As String
'���ܣ����ҽ����Ӧ�������Ϣ
'������str���=������ϵ���������ַ���
'���أ�������ϵ�ID�����ŷָ�
    Dim rsTmp As Recordset, strSQL As String
    Dim strReturn As String
    
    strSQL = "Select  A.ID,a.�������" & vbNewLine & _
            "From ������ϼ�¼ A, �������ҽ�� B" & vbNewLine & _
            "Where b.���id=a.id And  b.ҽ��ID=[1]" & vbNewLine & _
            "Order By b.rowID"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��������", lngҽ��ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            str��� = str��� & "," & rsTmp!�������
            strReturn = strReturn & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        str��� = Mid(str���, 2)
        strReturn = Mid(strReturn, 2)
    End If
    GetAdviceDiag = strReturn
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CanUnExec(ByVal datExec As Date, Optional ByVal datNow As Date) As Boolean
'���ܣ�����ִ�м�¼��ִ��ʱ���ж��ܷ�ȡ��ִ�л�ȡ�����
'������datExec=ִ�м�¼��ִ��ʱ��
'      datNow =��ǰʱ��
'���أ�CanUnExec=true-����ȡ��ִ�У�false-������ȡ��ִ��

    Dim lngDatDiff As Long
    If datExec <> CDate(Format("0", "yyyy-MM-dd HH:mm")) Then
        If datNow = CDate(0) Then
            datNow = zlDatabase.Currentdate
        End If
        lngDatDiff = DateDiff("D", datExec, datNow)
        CanUnExec = lngDatDiff <= gintҽ��ִ����Ч����
    Else
        CanUnExec = True
    End If
    
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub

Public Function InitObjRecipeAudit(ByVal lngProgram As Long) As Boolean
    If gobjRecipeAudit Is Nothing Then
        On Error Resume Next
        Set gobjRecipeAudit = CreateObject("zl9RecipeAudit.clsBusiness")
        If Not gobjRecipeAudit Is Nothing Then
            Call gobjRecipeAudit.Init(gcnOracle, lngProgram = 1252)
        End If
        err.Clear: On Error GoTo 0
    End If
    InitObjRecipeAudit = Not gobjRecipeAudit Is Nothing
End Function

Public Function Check��ҩ�洢�ⷿ(ByVal lngBegin As Long, ByVal lngEnd As Long, str��ҩ���� As String, ByRef vsAdvice As VSFlexGrid, ByVal bytMode As Byte, _
ByVal lng���˿���ID As Long, ByVal COL_��� As Long, ByVal col_ҽ������ As Long, ByVal COL_�շ�ϸĿID As Long, ByVal COL_ִ�п���ID As Long) As Boolean
'���ܣ����ָ������ҩ�䷽������ҩƷ�Ƿ������˴洢�ⷿ
'������str��ҩ����=����δ���ô洢�ⷿ����ҩ���ƴ�
'      bytMode 1-���ﳡ��,2-סԺ����
    Dim lng��ҩ�� As Long, strIDs As String
    Dim i As Long, lngִ�п���ID As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim colID As New Collection
    
    Check��ҩ�洢�ⷿ = True
   
    With vsAdvice
         For i = lngBegin To lngEnd
             If .TextMatrix(i, COL_���) = "7" Then
                 colID.Add .TextMatrix(i, col_ҽ������), "C" & .TextMatrix(i, COL_�շ�ϸĿID)
                 strIDs = strIDs & "," & .TextMatrix(i, COL_�շ�ϸĿID)
                 If lngִ�п���ID = 0 Then lngִ�п���ID = Val(.TextMatrix(i, COL_ִ�п���ID))
             End If
         Next
    End With
    If strIDs <> "" Then
        strIDs = Mid(strIDs, 2)
        strSQL = "Select /*+ rule*/Column_Value as ID" & vbNewLine & _
                "From Table(f_Num2list([1])) B" & vbNewLine & _
                "Where Not Exists (Select 1 From �շ�ִ�п��� A Where a.�շ�ϸĿid = b.Column_Value" & vbNewLine & _
                " And Nvl(a.������Դ,[4]) = [4] And ִ�п���id = [2] And Nvl(��������id, [3]) = [3])"

        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strIDs, lngִ�п���ID, lng���˿���ID, bytMode)
        If rsTmp.RecordCount > 0 Then
            For i = 1 To rsTmp.RecordCount
                str��ҩ���� = str��ҩ���� & "," & colID("C" & rsTmp!ID)
                rsTmp.MoveNext
            Next
            str��ҩ���� = Mid(str��ҩ����, 2)
            Check��ҩ�洢�ⷿ = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathAdviceIsExe(ByVal lngҽ��ID As Long) As Boolean
'����:��鵱ǰ·��ҽ����Ӧ��·����Ŀ�Ƿ��Ѿ�ִ�еǼ�
'����:ҽ��ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If Val(zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, p�ٴ�·��Ӧ��, 1)) Then
        '92225 ��¼ҽ��δУ�ԣ�������Ŀ�Ƿ�ִ�еǼǶ�����ɾ����
        strSQL = "Select Count(1) as ���� From ����ҽ����¼ Where ID = [1] And ������־ = 2"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
        If rsTmp.RecordCount > 0 Then
            If NVL(rsTmp!����, 0) = 1 Then Exit Function
        End If
        
        strSQL = "Select a.ִ��ʱ��" & vbNewLine & _
                 "From ����·��ִ�� A, (Select Min(a.·��ִ��id) As ·��ִ��id From ����·��ҽ�� A Where a.����ҽ��id = [1]) B" & vbNewLine & _
                 "Where a.Id = b.·��ִ��id"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
        If rsTmp.RecordCount > 0 Then
            If Not IsNull(rsTmp!ִ��ʱ��) Then
                CheckPathAdviceIsExe = True
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathAdviceIsExeOut(ByVal lngҽ��ID As Long) As Boolean
'����:��鵱ǰ·��ҽ����Ӧ��·����Ŀ�Ƿ��Ѿ�ִ�еǼ�
'����:ҽ��ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If Val(zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, P����·��Ӧ��, 1)) Then
        strSQL = "Select a.ִ��ʱ��" & vbNewLine & _
                 "From ��������·��ִ�� A, (Select Min(a.·��ִ��id) As ·��ִ��id From ��������·��ҽ�� A Where a.����ҽ��id = [1]) B" & vbNewLine & _
                 "Where a.Id = b.·��ִ��id"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
        If rsTmp.RecordCount > 0 Then
            If Not IsNull(rsTmp!ִ��ʱ��) Then
                CheckPathAdviceIsExeOut = True
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceInsure(ByVal int���� As Integer, ByVal bln���Ѷ��� As Boolean, ByVal lng����ID As Long, ByVal lng�������� As Long, _
   ByVal strIDs1 As String, ByVal strIDs2 As String, ByVal strҽ������ As String, Optional ByVal lng���˲���ID As Long) As String
'���ܣ�ҽ�������´�ҽ��ʱ��ҽ��¼��󣬶�ҽ���漰�ļƼ���Ŀ�ı��ն���������м��
'������strIDs1:ҩƷ���ĵ��շ�ϸĿID�ַ�����һ��ҽ�����磺��ù��+�����ǣ�:�շ�ϸĿID1,�շ�ϸĿID2,������
'      strIDs2 ������������Ŀ��������ĿID��һ��ҽ�����磺��Ѫ��Ŀ+��Ѫ;����:ִ�п����ַ��� ������ĿID1:ִ�п���1,������ĿID2:ִ�п���2,������
'      lng��������=1���=2סԺ
'      strҽ�����ݣ��û���ʾʱ��ʾ��ҽ������
'      bln���Ѷ���=False ��ʾ��ǰ��������飬=True �������
'���أ���ʾ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    If gintҽ������ = 0 Or int���� = 0 Or Not bln���Ѷ��� Then Exit Function
    If gclsInsure.GetCapability(support��������ҽ����Ŀ, lng����ID, int����) Then Exit Function
    
    
    If strIDs1 = "" And strIDs2 = "" Then Exit Function
    
    If strIDs1 <> "" Then
        If Mid(strIDs1, 1, 1) = "," Then strIDs1 = Mid(strIDs1, 2)
        strSQL = "Select Column_Value as �շ���ĿID From Table(f_Num2list([1]))"
    End If
    If strIDs2 <> "" Then
        If Mid(strIDs2, 1, 1) = "," Then strIDs2 = Mid(strIDs2, 2)
        If strIDs1 <> "" Then strSQL = strSQL & " Union All "
        '����û�мӲ�λ������������Ҫ��Distinct
        strSQL = strSQL & "Select �շ���ĿID From (" & _
                "Select Distinct C.�շ���ĿID,C.���ÿ���id" & _
                " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                " From �����շѹ�ϵ C,Table(f_Num2list2([2])) D Where C.������ĿID=D.c1" & _
                "      And (C.���ÿ���ID is Null or C.���ÿ���ID = Nvl(D.c2,[4]) And C.������Դ = " & IIF(lng�������� = 1, 1, 2) & ")" & _
                " ) Where Nvl(���ÿ���id, 0) = Top"
    End If
    
    strSQL = "Select /*+ RULE */ Distinct C.����,B.�շ�ϸĿID" & _
        " From (" & strSQL & ") A,����֧����Ŀ B,�շ���ĿĿ¼ C" & _
        " Where A.�շ���ĿID=B.�շ�ϸĿID(+) And A.�շ���ĿID=C.ID" & _
        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
        " And B.����(+)=[3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckAdviceInsure", strIDs1, strIDs2, int����, lng���˲���ID)
    strSQL = "": i = 0
    Do While Not rsTmp.EOF
        If IsNull(rsTmp!�շ�ϸĿID) Then
            If i = 8 Then
                strSQL = strSQL & vbCrLf & "�� ��"
                Exit Do
            End If
            strSQL = strSQL & vbCrLf & "��" & rsTmp!����
            i = i + 1
        End If
        rsTmp.MoveNext
    Loop
    If strSQL <> "" Then
        CheckAdviceInsure = "��ǰ������ҽ�����ˣ���ҽ�������¼Ƽ���Ŀû�����ö�Ӧ�ı�����Ŀ��" & vbCrLf & vbCrLf & _
            "ҽ�����ݣ�" & vbCrLf & strҽ������ & vbCrLf & vbCrLf & "�Ƽ���Ŀ��" & strSQL
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckWaitQuittance(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ�����ת��ҽ��ʱ�� ����Ƿ����δ������ʵĵ����ݸ��ݲ�������Ӧ��ʾ
'���أ����
    
    Dim strSQL As String
    Dim strDrug As String
    Dim rsTmp As ADODB.Recordset
    
    If gbytת��ʱδ������ʵ��ݼ�� = 0 Then Exit Function
    
    On Error GoTo ErrHand
    
    strSQL = " Select Distinct a.No, d.���� ��Ŀ, c.���� As ����" & _
        " From סԺ���ü�¼ a, ���˷������� b, ���ű� c, �շ���ĿĿ¼ d" & _
        " Where a.Id = b.����id And a.�շ�ϸĿid = d.Id And b.��˲���id = c.Id(+)" & _
        " And b.���ʱ�� Is Null And a.����id = [1] And Nvl(a.��ҳid, 0) = [2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckWaitQuittance", lng����ID, lng��ҳID)
    strDrug = ""
    Do While Not rsTmp.EOF
        If strDrug = "" Then
            strDrug = "����[" & NVL(rsTmp!NO) & "]�е�" & NVL(rsTmp!��Ŀ) & "����" & NVL(rsTmp!����, "[δ֪����]") & "δ���"
        Else
            If InStr(1, vbCrLf & strDrug & vbCrLf, vbCrLf & "����[" & NVL(rsTmp!NO) & "]�е�" & NVL(rsTmp!��Ŀ) & "����" & NVL(rsTmp!����, "[δ֪����]") & "δ���" & vbCrLf) = 0 Then
                If LenB(StrConv(strDrug & vbCrLf & "����[" & NVL(rsTmp!NO) & "]�е�" & NVL(rsTmp!��Ŀ) & "����" & NVL(rsTmp!����, "[δ֪����]") & "δ���", vbFromUnicode)) <= 1000 Then
                    strDrug = strDrug & vbCrLf & "����[" & NVL(rsTmp!NO) & "]�е�" & NVL(rsTmp!��Ŀ) & "����" & NVL(rsTmp!����, "[δ֪����]") & "δ���"
                Else
                    strDrug = strDrug & vbCrLf & "... ..."
                End If
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If strDrug <> "" Then
        If gbytת��ʱδ������ʵ��ݼ�� = 1 Then
            If MsgBox("�ò��˴���δ������ʵĵ��ݣ�" & vbCrLf & vbCrLf & strDrug & vbCrLf & vbCrLf & "ȷ��Ҫ����ת��ҽ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckWaitQuittance = True
                Exit Function
            End If
        Else
            MsgBox "�ò��˴���δ������ʵĵ��ݣ�" & vbCrLf & vbCrLf & strDrug & vbCrLf & vbCrLf & "���ʷ���ת��ҽ����", vbInformation, gstrSysName
            CheckWaitQuittance = True
            Exit Function
        End If
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Getҽ����������(ByVal lngҽ��ID As Long, ByVal str������ As String) As String
'����:����ҽ��ID��Ԫ�����ơ�����ҽ���Ķ�ӦԪ�ص����븽������
'����:str������  ����������Ŀ.������

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
     
    strSQL = "Select a.���� From ����ҽ������ A, ����������Ŀ B" & _
        " Where a.Ҫ��id = b.Id And a.ҽ��id = [1] And b.������ = [2]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID, str������)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strTmp = IIF(strTmp = "", "", strTmp & ",") & rsTmp!����
            rsTmp.MoveNext
        Next
    End If
    
    Getҽ���������� = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getҽ������ҽ��IDs(ByVal lng����ID As Long, ByVal str����ID As String, ByVal lngִ�п���ID As Long, ByVal blnסԺ As Boolean, ByVal lngǰ��ID As Long) As String
'���ܣ��ڵ�ǰ����ִ�е�����ҽ��
'������str����ID �����סԺ����ҳid �������ǹҺŵ���blnסԺ=true ��ʾסԺҽ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    strSQL = "Select ID From ����ҽ����¼ A Where a.����id = [1] And a.ִ�п���id = [2]" & _
        IIF(blnסԺ, " And a.��ҳid = [3]", " And a.�Һŵ� = [3] ") & " Order By ����ʱ�� Desc"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lngִ�п���ID, IIF(blnסԺ, Val(str����ID), str����ID))
    Do While Not rsTmp.EOF
        strTmp = IIF(strTmp = "", "", strTmp & ",") & rsTmp!ID
        rsTmp.MoveNext
    Loop
    If Len(strTmp) > 4000 Then
        '�������ȡ4000����������SQL����ͱȽ��鷳�ˣ���ҵ���Ͽ���ֻ�������ξ���ʹ����ͬ�ĹҺŵ���ҽ��վ��������ʱ�������ô��ҽ��ҽ��������Ѫ͸�ȣ�����ֻ��ȡ���4000���ȵ�ID���ɣ���ǰ��ҽ��û��ʲô����
        strTmp = Mid(strTmp, 1, 3980)
        strTmp = Mid(strTmp, 1, InStrRev(strTmp, ",") - 1)
    End If
    
    If strTmp <> "" Then
        If InStr("," & strTmp & ",", "," & lngǰ��ID & ",") = 0 Then
            strTmp = strTmp & "," & lngǰ��ID
        End If
    Else
        strTmp = lngǰ��ID
    End If
    
    Getҽ������ҽ��IDs = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncTraReaction(ByVal lngҽ��ID As Long, ByVal lngMoudle As Long, ByVal blnMoved As Boolean) As Boolean
'���ܣ���Ѫ��Ӧ
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, lng��ҳID As Long, lng������Դ As Long
    

    If InitObjBlood(True) = False Then Exit Function
    
    On Error GoTo errH
    strSQL = "Select B.����ID,B.��ҳID,B.������Դ,A.ID ����ID From ���˹Һż�¼ A,����ҽ����¼ B where B.�Һŵ�=A.NO(+) And  B.id=[1]"
    If blnMoved = True Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����Ӧ", lngҽ��ID)
    lng����ID = Val("" & rsTemp!����ID)
    If IsNull(rsTemp!��ҳID) Then
        lng��ҳID = Val("" & rsTemp!����ID)
    Else
        lng��ҳID = Val("" & rsTemp!��ҳID)
    End If
    lng������Դ = Val("" & rsTemp!������Դ)
    Call gobjPublicBlood.zlShowBloodReaction(Nothing, glngSys, lngMoudle, 1, lng����ID, lng��ҳID, lng������Դ)
    
    FuncTraReaction = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Read��ҩ����(lngID As Long) As ADODB.Recordset
'���ܣ���ȡָ��ҩ���ķ�ҩ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ��ҩ���� Where ҩ��ID=[1] Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngID)
    Set Read��ҩ���� = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���˴�ӡ��¼DelSQL(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal intBaby As Integer, Optional ByVal int��Ч As Integer, _
    Optional ByVal lngҽ��ID As Long, Optional ByVal strҽ��IDs As String, Optional ByVal blnHaveBaby As Boolean, Optional ByRef strMsg As String) As String
'���ܣ���ȡ���˽�Ҫɾ��  ����ҽ����ӡ  �������ݵĹ���SQL������ɾ��Ԥ��ӡ�ļ�¼ ����ҽ����ӡ.��ӡʱ�� is null
'      �磺"Zl_����ҽ����ӡ_Delete(718,1,0,0,1)|Zl_����ҽ����ӡ_Delete(718,1,0,1,1)|Zl_����ҽ����ӡ_Delete(718,1,1,0,1)";
'������lng����ID��lng��ҳID��intBaby��int��Ч��lngҽ��ID��strҽ��IDs
'      blnHaveBaby ��ǰ�����Ƿ���Ӥ��
'intType ����ʱ����
'      intType=2ͨ�ý�������ҽ��ʱ�����봫�Ĳ����� intType��lng����ID��lng��ҳID��strҽ��IDs��blnHaveBaby
'      intType=3ҽ��վ����ҽ��ʱ(ֻ�ܵ�������)�����봫�Ĳ����� intType��lng����ID��lng��ҳID��intBaby��int��Ч��lngҽ��ID��blnHaveBaby
'      intType=4����վ������ɾ��ҽ��ʱ�����봫�Ĳ����� intType��lng����ID��lng��ҳID��intBaby��int��Ч��strҽ��IDs��blnHaveBaby
'          intType=5���δ�ӡʱ�����봫�Ĳ����� intType��lng����ID��lng��ҳID��intBaby��int��Ч��lngҽ��ID��blnHaveBaby
'          intType=5�������ʱûʹ�ã�����ǰ�ĳ������Ѿ���ȫ�����ˡ�
'      strMsg ���ز��������������ĳ���ҽ������ӡ��¼���ӵ�3ҳ���������
'˵����intType in (3,4,5)ʱ���صĹ���SQLֻ��һ��������ɲ�������ֱ���á�

    Dim rsTmp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim strSQL As String, strMsgTmp As String
    Dim strӤ������ As String
    Dim i As Integer
    Dim intNoP As Integer '1�����δ��ӡ�������ɵļ�¼��0������Ѿ�����ļ�¼��֮���
    
    On Error GoTo errH
    
    If blnHaveBaby Then
        If intType = 1 Or intType = 2 Then
            strSQL = "Select ���,Ӥ������ as ���� From ������������¼ Where ����ID=[1] And ��ҳID=[2]"
            Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "Get����Ԥ��ӡ��¼DelSQL", lng����ID, lng��ҳID)
        Else
            strSQL = "Select Ӥ������ as ���� From ������������¼ Where ����ID=[1] And ��ҳID=[2] and ���=[3]"
            Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "Get����Ԥ��ӡ��¼DelSQL", lng����ID, lng��ҳID, intBaby)
            If Not rsBaby.EOF Then strӤ������ = rsBaby!���� & ""
        End If
    End If
    
    'SQL��ȡ����  λ��  ��һ��ҽ���е���Сλ�ã�����һ��ҽ��������������õ���������ɾ��ҽ��Ԥ��ӡ��¼ʱһ������һ��ҽ��Ϊ��λ��
    '��ɾ���Ѿ���ӡ�ļ�¼ʱ����ҳɾ����ҩƷҽ�����ܻ������������ҩƷҽ��ռ�����У�������Ӱ����ȷ�ԡ�
    Select Case intType
    Case 1, 2
        strSQL = "Select a.Ӥ��,a.��Ч, Min(LPad(ҳ��,4,'0')||LPad(�к�,3,'0')) As λ��,Min(a.ҳ��) As ҳ��,min(a.��ӡʱ��) as ��ӡʱ��" & _
            " From ����ҽ����ӡ A, ����ҽ����¼ B" & vbNewLine & _
            " Where a.����id = b.����id and b.id=a.ҽ��id And a.��ҳid = b.��ҳid and a.����id=[1] and a.��ҳid=[2]" & _
            " And b.id In (Select Column_Value From Table(Cast(f_Num2list([6]) As zlTools.t_Numlist)))" & _
            " Group By a.Ӥ��,a.��Ч having Min(a.ҳ��)>0"
    Case 3, 5
        strSQL = "Select Min(LPad(ҳ��,4,'0')||LPad(�к�,3,'0')) As λ��,Min(a.ҳ��) As ҳ��,min(a.��ӡʱ��) as ��ӡʱ��" & _
            " From ����ҽ����ӡ A, ����ҽ����¼ B" & _
            " Where a.����id = b.����id And a.��ҳid = b.��ҳid and b.id=a.ҽ��id and a.����id=[1] and a.��ҳid=[2]" & _
            " And a.Ӥ��=[3] and a.��Ч=[4] and (b.id =[5] or b.���id=[5]) having Min(a.ҳ��)>0"
    Case 4
        strSQL = "Select Min(LPad(ҳ��,4,'0')||LPad(�к�,3,'0')) As λ��,Min(a.ҳ��) As ҳ��,min(a.��ӡʱ��) as ��ӡʱ��" & _
            " From ����ҽ����ӡ A, ����ҽ����¼ B" & _
            " Where a.����id = b.����id And a.��ҳid = b.��ҳid and b.id=a.ҽ��id and a.����id=[1] and a.��ҳid=[2]" & _
            " And a.Ӥ��=[3] and a.��Ч=[4] and (b.id =(Select Column_Value From Table(Cast(f_Num2list([6]) As zlTools.t_Numlist)))" & _
            " or b.���id=(Select Column_Value From Table(Cast(f_Num2list([6]) As zlTools.t_Numlist)))) having Min(a.ҳ��)>0"
    End Select
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get����Ԥ��ӡ��¼DelSQL", lng����ID, lng��ҳID, intBaby, int��Ч, lngҽ��ID, strҽ��IDs)
    
    ' IsNull(rsTmp!��ӡʱ��) ��ֻ��Ҫ�����¼���ڽ����ϲ�����ʾ
    strSQL = ""
    If Not rsTmp.EOF Then
        Select Case intType
        Case 1, 2
            For i = 1 To rsTmp.RecordCount
                If Val(rsTmp!Ӥ�� & "") <> 0 Then
                    rsBaby.Filter = "���=" & Val(rsTmp!Ӥ�� & "")
                    If Not rsBaby.EOF Then strӤ������ = rsBaby!���� & ""
                End If
                intNoP = 1
                If Not IsNull(rsTmp!��ӡʱ��) Then
                    intNoP = 0
                    strMsgTmp = IIF(strMsgTmp = "", "", strMsgTmp & vbCrLf) & _
                        "�ò���" & IIF(strӤ������ = "", "", "Ӥ��-" & strӤ������) & "��" & IIF(Val(rsTmp!��Ч & "") = 0, "����", "��ʱ") & "ҽ�����Ĵ�ӡ��¼���ӵ�" & _
                        Val(rsTmp!ҳ�� & "") & "ҳ��ʼ�������"
                    strӤ������ = ""
                End If
                strSQL = strSQL & "|" & "Zl_����ҽ����ӡ_Delete(" & lng����ID & "," & lng��ҳID & "," & Val(rsTmp!Ӥ�� & "") & "," & Val(rsTmp!��Ч & "") & "," & Val(rsTmp!ҳ�� & "") & ",'" & rsTmp!λ�� & "')"
                rsTmp.MoveNext
            Next
            strSQL = Mid(strSQL, 2)
        Case 3, 4, 5
            intNoP = 1
            If Not IsNull(rsTmp!��ӡʱ��) Then
                intNoP = 0
                strMsgTmp = "�ò���" & IIF(strӤ������ = "", "", "Ӥ��-" & strӤ������) & "��" & IIF(int��Ч = 0, "����", "��ʱ") & "ҽ�����Ĵ�ӡ��¼���ӵ�" & _
                        Val(rsTmp!ҳ�� & "") & "ҳ��ʼ�������"
            End If
            strSQL = "Zl_����ҽ����ӡ_Delete(" & lng����ID & "," & lng��ҳID & "," & intBaby & "," & int��Ч & "," & Val(rsTmp!ҳ�� & "") & ",'" & rsTmp!λ�� & "')"
        End Select
    End If
    
    strMsg = ""
    
    strMsg = strMsgTmp
    
    Get���˴�ӡ��¼DelSQL = strSQL
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'���ܣ���ȡҩƷ�����ĳ�����ļ���
'������bytType:0-ҩƷ��1-����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '�������
    
    strSQL = _
        " Select Distinct A.ID,C.��鷽ʽ" & _
        " From ���ű� A,��������˵�� B," & IIF(bytType = 0, "ҩƷ������", "���ϳ�����") & " C" & _
        " Where B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� " & IIF(bytType = 0, "IN('��ҩ��','��ҩ��','��ҩ��')", "='���ϲ���'") & _
        " And C.�ⷿID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetStockCheck")
    For i = 1 To rsTmp.RecordCount
        colStock.Add NVL(rsTmp!��鷽ʽ, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function ItemExistInsure(ByVal lng����ID As Long, ByVal lng�շ�ϸĿID As Long, ByVal int���� As Integer) As Boolean
'���ܣ��ж��շ���Ŀ�Ƿ������˱���֧����Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    If gclsInsure.GetCapability(support��������ҽ����Ŀ, lng����ID, int����) Then
        ItemExistInsure = True: Exit Function
    End If
    
    strSQL = "Select 1 From ����֧����Ŀ Where �շ�ϸĿID=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng�շ�ϸĿID, int����)
    ItemExistInsure = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(lng����ID As Long, ByVal byt��Դ As Byte) As Currency
'����:��ȡָ�����˵ļ��ʻ��۵����ϼ�
'����:byt��Դ:1-���2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    strTab = IIF(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ� From " & strTab & " Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1]"
    
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!���۷��úϼ�
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAuditRecord(lng����ID As Long, lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡָ�����˵ķ���������Ŀ
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ��ĿId,ʹ������,��������,ʹ������-�������� �������� From ����������Ŀ Where ����ID=[1] And ��ҳID=[2]"
    Set GetAuditRecord = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBalanceRevoke(ByVal lngҽ��ID As Long) As Boolean
'���ܣ�(����)��Ҫ���ϵ�ҽ����Ӧ�ķ��õĽ���������м��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng���ͺ� As Long
    
    If gbytBillOpt = 0 Then
        CheckAdviceBalanceRevoke = True
        Exit Function
    End If
    
    On Error GoTo errH
    
    'ҽ��IDΪ����ֵ������ҽ����һ�������˵�,�����޷��͡�
    strSQL = "Select Distinct ���ͺ� From ����ҽ������" & _
        " Where ҽ��ID IN(Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngҽ��ID)
    If rsTmp.EOF Then Exit Function
    lng���ͺ� = rsTmp!���ͺ�
    
    '����������"ZL_����ҽ����¼_����"
    strSQL = "Select A.NO,Nvl(A.�۸񸸺�,A.���) as ���,Sum(Nvl(A.���ʽ��,0)) as ���ʽ��" & _
        " From ������ü�¼ A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ I" & _
        " Where A.NO=B.NO And A.��¼���� IN(2,12) And A.��¼״̬=1 And A.ҽ�����=B.ҽ��ID And B.ҽ��ID=C.ID" & _
        " And B.��¼����=2 And C.������ĿID=I.ID And B.���ͺ�=[1] And (C.ID=[2] Or C.���ID=[2])" & _
        " And (" & _
            " A.�շ���� Not In ('5','6','7','E')" & _
            " Or A.�շ����='E' And I.�������� Not In ('2','3','4')" & _
            " Or A.�շ���� In ('5','6','7') And Nvl(A.ִ��״̬,0)=0" & _
            " Or Exists(Select 1 From zlParameters Where ϵͳ=[3] And ģ�� is NULL And Nvl(˽��,0)=0 And ������=68 And Nvl(����ֵ,'0')='0')" & _
            " )" & _
        " Group by A.NO,Nvl(A.�۸񸸺�,A.���) Having Sum(Nvl(A.���ʽ��,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng���ͺ�, lngҽ��ID, glngSys)
    If Not rsTmp.EOF Then
        If gbytBillOpt = 1 Then
            If MsgBox("Ҫ����ҽ���Ķ�Ӧ�����д����ѽ��ʵķ��ã�ȷʵҪ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        ElseIf gbytBillOpt = 2 Then
            MsgBox "Ҫ����ҽ���Ķ�Ӧ�����д����ѽ��ʵķ��ã��������ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckAdviceBalanceRevoke = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBillingRevoke(ByVal lngҽ��ID As Long) As Boolean
'���ܣ�(����)��Ҫ���ϵ�ҽ����Ӧ�ļ��ʷ��õ����������м��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng���ͺ� As Long
    
    On Error GoTo errH
    
    'ҽ��IDΪ����ֵ������ҽ����һ�������˵�,�����޷��͡�
    strSQL = "Select Distinct ���ͺ� From ����ҽ������" & _
        " Where ҽ��ID IN(Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngҽ��ID)
    If rsTmp.EOF Then Exit Function
    lng���ͺ� = rsTmp!���ͺ�
    
    '����������"ZL_����ҽ����¼_����"
    strSQL = "Select A.NO,A.���" & _
        " From ������ü�¼ A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ I" & _
        " Where A.NO=B.NO And A.��¼���� IN(2,12) And A.��¼״̬=1" & _
        " And A.������ Is Not NULL And A.������<>A.����Ա����" & _
        " And A.ҽ�����=B.ҽ��ID And B.ҽ��ID=C.ID And B.��¼����=2" & _
        " And C.������ĿID=I.ID And B.���ͺ�=[1] And (C.ID=[2] Or C.���ID=[2])" & _
        " And (" & _
            " A.�շ���� Not In ('5','6','7','E')" & _
            " Or A.�շ����='E' And I.�������� Not In ('2','3','4')" & _
            " Or A.�շ���� In ('5','6','7') And Nvl(A.ִ��״̬,0)=0" & _
            " Or Exists(Select 1 From zlParameters Where ϵͳ=[3] And ģ�� is NULL And Nvl(˽��,0)=0 And ������=68 And Nvl(����ֵ,'0')='0')" & _
            " )"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng���ͺ�, lngҽ��ID, glngSys)
    If Not rsTmp.EOF Then Exit Function
    
    CheckAdviceBillingRevoke = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMoneyInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional curModiMoney As Currency) As ADODB.Recordset
'���ܣ���ȡָ�����˵�ʣ���
'������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(�������,0) as �������,Nvl(Ԥ�����,0) as Ԥ�����" & _
            " From ������� Where ����=1 And ���� = " & IIF(lng��ҳID = 0, 1, 2) & " And ����ID= [1] "
    
    If curModiMoney <> 0 Then   '����Ҫ��Union��ʽ,���ֱ��ȥ��,�ڲ�������޼�¼ʱ,���᷵�ؼ�¼
        strSQL = strSQL & " Union All  Select -1* " & curModiMoney & " as �������,0 as Ԥ����� From Dual"
        strSQL = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ����� From (" & strSQL & ")"
    End If
            
    '���Ϊҽ��סԺ���ˣ����ڷ���������ſ�Ԥ���еķ���(���ڱ���)
    If lng��ҳID <> 0 Then
        strSQL = strSQL & " Union All " & _
            " Select -1*Nvl(Sum(���),0) as �������,0 as Ԥ�����" & _
            " From ����ģ����� Where ����ID=[1] And ��ҳID=[2]"
        strSQL = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ����� From (" & strSQL & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBalanceRoll(ByVal lng���ͺ� As Long, ByVal lngҽ��ID As Long, Optional ByVal blnBat As Boolean) As Boolean
'���ܣ�(סԺ)��Ҫ���˵�ҽ����Ӧ�ķ��õĽ���������м��(һ������һ��סԺ��)
'������blnBat=�Ƿ�Ҫ������������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intInsure As Integer
    
    On Error GoTo errH
        
    'ȡҪ���˵ļ���NO
    If blnBat Then
        strSQL = "Select Distinct ҽ��ID,NO From ����ҽ������ Where ��¼����=2 And ���ͺ�=[1]"
    Else
        strSQL = "Select Distinct A.ҽ��ID,A.NO From ����ҽ������ A,����ҽ����¼ B" & _
            " Where A.ҽ��ID=B.ID And A.��¼����=2 And A.���ͺ�=[1] And (B.ID=[2] Or B.���ID=[2])"
    End If
    'ȡ��ЩNO�Ľ������(�ǻ���δ����)
    strSQL = "Select A.NO,Nvl(A.�۸񸸺�,A.���) as ���,Sum(Nvl(A.���ʽ��,0)) as ���ʽ��" & _
        " From סԺ���ü�¼ A,(" & strSQL & ") B Where A.NO=B.NO And A.ҽ�����=B.ҽ��ID And A.��¼���� IN(2,12) " & _
        " Group by A.NO,Nvl(A.�۸񸸺�,A.���) Having Sum(Nvl(A.���ʽ��,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng���ͺ�, lngҽ��ID)
    If Not rsTmp.EOF Then
        strSQL = "Select A.����ID,A.���� From ������ҳ A,����ҽ����¼ B" & _
            " Where Rownum=1 And A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngҽ��ID)
        If Not rsTmp.EOF Then intInsure = NVL(rsTmp!����, 0)
        If intInsure <> 0 Then '�ȶ�ҽ�������ƽ��м��
            If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, rsTmp!����ID, intInsure) Then
                MsgBox "�ò���Ϊҽ�����ˣ�Ҫ����ҽ���ķ��ͷ����д����ѽ��ʵķ��ã����ܻ��ˡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If gbytBillOpt <> 0 Then
            If gbytBillOpt = 1 Then
                If MsgBox("Ҫ����ҽ���ķ��ͷ����д����ѽ��ʵķ��ã�ȷʵҪ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf gbytBillOpt = 2 Then
                MsgBox "Ҫ����ҽ���ķ��ͷ����д����ѽ��ʵķ��ã����ܻ��ˡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    CheckAdviceBalanceRoll = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceDrugSurplus(ByVal lng���ͺ� As Long, Optional ByVal lngҽ��ID As Long) As String
'���ܣ���������ҩƷҽ���������Ƿ���ڵ�ǰ���������
'������lng���ͺ�=Ҫ���˵ķ��ͺ�
'      lngҽ��ID=Ҫ���˵�һ��ҩƷҽ����ID�������ָ���ɱ�ʾ�������˶���ҽ��
'���أ���ʾ��Ϣ
'˵������ʿ���ܻ���ҽ���Ĳ���������ֻ�漰סԺ���ü�¼(ҽ���ſ��ܷ�������Ϊ�������)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select C.ҽ������ as ҩƷ,A.�շ�ϸĿID as ҩƷID,A.���˲���ID as ����ID,A.ִ�в���ID as �ⷿID,Sum(A.����) as ��������" & _
        " From סԺ���ü�¼ A,����ҽ������ B,����ҽ����¼ C" & _
        " Where A.ҽ�����=B.ҽ��ID And A.NO=B.NO And A.��¼����=B.��¼����" & _
        " And B.ҽ��ID=C.ID And A.�շ���� In('5','6') And A.�۸񸸺� Is Null" & _
        " And B.���ͺ�=[1] And C.������� IN('5','6') And (C.���ID=[2] Or [2]=0)" & _
        " Group by C.ҽ������,A.�շ�ϸĿID,A.���˲���ID,A.ִ�в���ID"
    strSQL = _
        " Select A.ҩƷ,D.���� as �ⷿ,C.סԺ��װ,C.סԺ��λ,A.��������,B.��������" & _
        " From (" & strSQL & ") A,ҩƷ����ƻ� B,ҩƷ��� C,���ű� D" & _
        " Where A.�ⷿID=D.ID And A.ҩƷID=C.ҩƷID" & _
        " And A.����ID=B.����ID(+) And A.�ⷿID=B.�ⷿID(+) And A.ҩƷID=B.ҩƷID(+) And B.״̬(+)=0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckAdviceDrugSurplus", lng���ͺ�, lngҽ��ID)
    Do While Not rsTmp.EOF
        If NVL(rsTmp!��������, 0) > NVL(rsTmp!��������, 0) And NVL(rsTmp!��������, 0) <> 0 Then
            strMsg = strMsg & vbCrLf & "��[" & rsTmp!ҩƷ & "]��""" & rsTmp!�ⷿ & """�Ļ������� " & _
                FormatEx(NVL(rsTmp!��������, 0) / NVL(rsTmp!סԺ��װ, 1), 5) & rsTmp!סԺ��λ & "����ǰ�������� " & _
                FormatEx(NVL(rsTmp!��������, 0) / NVL(rsTmp!סԺ��װ, 1), 5) & rsTmp!סԺ��λ
        End If
        rsTmp.MoveNext
    Loop
    
    If strMsg <> "" Then strMsg = "����ҩƷ�Ļ���������������������" & vbCrLf & strMsg & vbCrLf & vbCrLf & "Ҫ������"
    CheckAdviceDrugSurplus = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CanEditBloodAdvice(ByVal lngID As Long, ByVal int���״̬ As Integer, ByVal bln�� As Boolean, Optional ByVal bln��Ѫ As Boolean = False, Optional ByVal blnMsg As Boolean = True) As Boolean
'���ܣ���Ѫҽ���ɷ�༭ int���״̬ ȡֵ�У�0��1��2��4��5��bln�� �Ƿ��ǽ���ҽ������ֻ��鱸Ѫҽ����������Ҫ����ǰ�ϱ�Ѫ���̵ļ�飩
    Dim strMsg As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    CanEditBloodAdvice = True
    
    If int���״̬ = 0 Or int���״̬ = 1 Then Exit Function
    On Error GoTo ErrHand
    strSQL = "Select ҽ��ID from ����ҽ��״̬ where ҽ��ID=[1] and ��������=[2]"
    If gblnѪ��ϵͳ Then
        If bln��Ѫ = True Then Exit Function '��Ѫҽ�������޸�
        If int���״̬ = 5 Or int���״̬ = 2 Then
            strMsg = "����Ѫ�����Ѿ���Ѫ�����" & IIF(int���״̬ = 5, "������Ѫ", "�����������Ѫ") & "���������޸ģ���Ҫ�޸�������Ѫ����ϵ��"
        ElseIf int���״̬ = 6 Then
            strMsg = "����Ѫ�����Ѿ���Ѫ����ղ�����ֹͣ��Ѫ���������޸ģ���Ҫ�޸�������Ѫ����ϵ��"
        ElseIf int���״̬ = 4 Or int���״̬ = 7 Then
            If Not bln�� And gbln��Ѫ�ּ����� Then
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��Ѫ����", lngID, IIF(int���״̬ = 4, 11, 18))
                If Not rsTmp.EOF Then
                    If int���״̬ = 4 Then
                        strMsg = "����Ѫ�����Ѿ������ˣ��������޸ģ���Ҫ�޸�������Ѫ��˹����л�����˲�����"
                    Else
                        strMsg = "����Ѫ�����Ѿ���ʼ��ˣ��������޸ģ���Ҫ�޸�������Ѫ��˹����л�����˲�����"
                    End If
                End If
            End If
        End If
    Else
        If int���״̬ = 2 Then
            strMsg = "����Ѫ�����Ѿ���ˣ��������޸ģ���Ҫ�޸�������Ѫ��˹����л�����˲�����"
        End If
    End If
    
    If strMsg <> "" Then
        If blnMsg = True Then
            MsgBox strMsg, vbInformation, "��Ѫ����"
        End If
        CanEditBloodAdvice = False
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckCHLimited(ByVal lngRow As Long, ByVal int���� As Integer, ByRef blnOutOfRange As Boolean, ByRef vsAdvice As VSFlexGrid, ByVal COL_���ID As Long, ByVal COL_������ĿID As Long, ByVal COL_��� As Long, ByVal COL_���� As Long) As Boolean
'���ܣ������ҩ�䷽ÿζҩ�Ĵ�������
'������blnMsg �Ƿ񵯳���Ϣ��ʾ��
'      blnOutOfRange �������������ó�����Ը��ݸò�������һ���������罫 .TextMatrix(i, COL_�Ƿ���)="1"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lng��ҩ���� As Long
    Dim colAmount As New Collection, strIDs As String

    CheckCHLimited = True
    
    On Error GoTo errH
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_���) = "7" Then
                
                    '��ҩ���԰�����´�󣬼�������ʱҪ���⴦��
                    If InStr("," & strIDs & ",", "," & .TextMatrix(i, COL_������ĿID) & ",") = 0 Then
                        strIDs = strIDs & "," & .TextMatrix(i, COL_������ĿID)
                        Call colAmount.Add(Val(.TextMatrix(i, COL_����)), "_" & .TextMatrix(i, COL_������ĿID))
                    Else
                        lng��ҩ���� = colAmount("_" & .TextMatrix(i, COL_������ĿID))
                        lng��ҩ���� = lng��ҩ���� + Val(.TextMatrix(i, COL_����))
                        
                        Call colAmount.Remove("_" & .TextMatrix(i, COL_������ĿID))
                        Call colAmount.Add(lng��ҩ����, "_" & .TextMatrix(i, COL_������ĿID))
                    End If
                End If
            Else
                Exit For
            End If
        Next
    End With
    If strIDs = "" Then Exit Function
    strIDs = Mid(strIDs, 2)
        
    strSQL = "Select /*+ rule*/ A.ID,A.����,A.���㵥λ,B.�������� From ������ĿĿ¼ A,ҩƷ���� B Where A.ID=B.ҩ��ID And Nvl(B.��������,0)<>0" & _
            " And A.ID IN (Select Column_Value From Table(f_Num2list([1])))"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckCHLimited", strIDs)
             
    For i = 1 To rsTmp.RecordCount
        If int���� * colAmount("_" & rsTmp!ID) > rsTmp!�������� Then
            blnOutOfRange = True: Exit Function
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub InitCardRs(ByRef rsCard As ADODB.Recordset)
'���ܣ���ʼ����¼����ҽ���༭�����·���Ƭ�ؼ���Ϣ
'˵��������סԺͨ�ã��޸Ĵ˷���ʱע�⿼������סԺ�������
    Set rsCard = New ADODB.Recordset
    
    With rsCard.Fields
        .Append "�Ƿ�����", adInteger '0-����,1-�޸�
        .Append "�Ƿ񱣴�", adInteger 'Ĭ��Ϊ 0 ���뵥���汣�水ť������ 0�������ã�1������
        .Append "����ҽ��", adInteger 'Ĭ��Ϊ 0
        .Append "����Ϊ���շѼ��ʵ�", adInteger 'Ĭ��Ϊ 0
        .Append "ֹͣ���г���", adInteger 'Ĭ��Ϊ 0
        .Append "��Чʱ��", adVarChar, 20 '��ҽ����ʼִ��ʱ��
        .Append "����ʱ��", adVarChar, 20
        .Append "�������", adInteger 'Ĭ��Ϊ 0
        .Append "ҽ������", adVarChar, 500
        .Append "ִ�п���ID", adBigInt 'ҽ��ִ�п��ҵ� ����ID
        .Append "����ִ��ID", adBigInt '����ҽ��ִ�п��ҵ� ����ID����Ѫ����ʱΪ��Ѫ;��ִ�п���id
        .Append "����ִ�п���ID", adBigInt '����ҽ��ִ�п��ҵ� ����ID����Ѫ����ʱΪ��Ѫ;��ִ�п���id
        .Append "��ĿIDs", adVarChar, 500 'һ��ҽ���е���ĿID����˳��̶�����Ϊҽ���е�˳���磺3542,2532,5478,......
        .Append "�������IDs", adVarChar, 500 '�μӻ���Ŀ���id��
        .Append "����ʱ��", adVarChar, 20 '����ҽ����������ʱ�䣬��Ѫҽ��������Ѫʱ�䣬�������ҽ���Ŀ�ʼʱ��
        .Append "��ҩ����", adVarChar, 500 '��Ѫ����ʱΪ ��Ѫԭ��
        .Append "ҽ����־", adInteger '0-��ȫ����,1-ȷ������Ŀ������2-ҽ������������3-�޸��Ѿ���ҽ����4-�޸�ԭʼҽ��
        .Append "����ĿID", adBigInt   '��ҽ����������ĿID
        .Append "����", adVarChar, 10 '��Ѫ����ʱΪ ��Ѫ����
    End With
    rsCard.CursorLocation = adUseClient
    rsCard.LockType = adLockOptimistic
    rsCard.CursorType = adOpenStatic
    rsCard.Open
End Sub

Public Sub DeleteRsExec(ByRef rsExec As Recordset, ByRef lngҽ��ID As Long)
'���ܣ����ݴ����ҽ��ID��ɾ��ҽ��ִ�мƼ��еĸ�ҽ����Ӧ������
    rsExec.Filter = "ҽ��ID=" & lngҽ��ID
    Do While Not rsExec.EOF
        rsExec.Delete
        rsExec.MoveNext
    Loop
End Sub

Public Function ApplyInPacs(frmParent As Object, ByRef lng������� As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngӤ����� As Long, ByVal lng�������� As Long, ByVal lngҽ��ID As Long, _
    ByVal lngҽ������ID As Long, ByVal lng����id As Long, ByVal lng����ID As Long, ByRef objVBA As Object, ByRef objScript As Object, ByRef rsDefine As ADODB.Recordset, Optional ByRef clsMipModule As Object, Optional ByVal lng��Ŀid As Long, Optional ByVal lngǰ��ID As Long) As Long
'���ܣ����ü�����뵥
'������ lngҽ��ID=�޸����뵥ʱ��ǰ�е�ҽ��ID,lng������� =��ǰ�޸��е��������
'       lng�������� 0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
'       lngҽ������ID ҽ������ID
'       lng����ID �����ת�����ˣ���Ϊԭ����ID��
'       lng����ID �����ת�����ˣ���Ϊԭ����ID��
'       objVBA objScript rsDefine VB����ͼ�¼�����ڲ���ҽ�������ı���blnMoved �����Ƿ��Ѿ�ת����clsMipModule ��Ϣ����
'���أ��������
    Dim objPacspplication As New clsPacsApplication
    Dim objAppPages()  As clsApplicationData
    Dim rsPati As Recordset
    Dim i As Long, j As Long, k As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim lngAdviceID As Long
    Dim strMsg As String
    Dim strExtra As String
    Dim strTmp As String
    Dim str��λ As String, str���� As String
    Dim strTmp���� As String
    Dim objTmp As New clsApplicationData
    Dim str�����Ժ��� As String
    Dim bln��ҽ As Boolean
    Dim str���� As String
    Dim strժҪ As String
    Dim strItems As String
    Dim strRISDel As String
    Dim strRISAdd As String 'RIS��������ʽ��ҽ��ID:������ĿID,....
    Dim arrSQL() As String
    Dim lng��ҽ��ID As Long  '����ҽ��ID����ֵ���˷ѣ�������ύ����ʱ�������ҽ��ID
    Dim blnDo As Long
    Dim strTabAdvice As String 'ҽ����Ϣ��ʱ��
    Dim strTxtAdvice As String 'ҽ������
    Dim bln���Ѷ��� As Boolean
    Dim vMsg As VbMsgBoxResult
    Dim rsPrice As ADODB.Recordset
    Dim lngҽ����� As Long
    
    On Error GoTo errH
    
    str�����Ժ��� = zlDatabase.GetPara("Ҫ��������Ժ���", glngSys, pסԺҽ���´�)
    '��ϼ��
    If InStr(str�����Ժ���, "D") > 0 Then
        bln��ҽ = Sys.DeptHaveProperty(lng����id, "��ҽ��")
        str���� = IIF(bln��ҽ, "2,12", "2")
        If Not ExistsDiagNoses(lng����ID, lng��ҳID, str����) Then
            strMsg = "���˵���Ժ��ϻ�û�����룬�������벡�˵���Ժ������´������롣"
        End If
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSQL = "Select A.סԺ��, A.��ǰ����, A.��������, Nvl(B.����, A.����) as  ����, Nvl(B.�Ա�, A.�Ա�) as  �Ա�, Nvl(B.����, A.����) as ����, A.�����, A.������,b.����,b.�ѱ�,b.��������" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B" & vbNewLine & _
            "Where A.����id = B.����id And A.����id = [1] And B.��ҳid = [2]"
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ApplyInPacs", lng����ID, lng��ҳID)
    
    If rsPati.RecordCount = 0 Then
        MsgBox "δ����ȷ��ȡ������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    '��ʼ��������뵥����
    Call objTmp.MakePacsData(lng�������, objAppPages())
    Call objPacspplication.InitComponents(Get��������ID(UserInfo.ID, 0, lng����id, IIF(lng�������� = 1, 1, 2)), frmParent)
    If objPacspplication.ShowApplicationForm(lng����ID, IIF(lng�������� = 1, 1, 2), 0, lng��ҳID, IIF(lng������� = 0, lngҽ��ID, lng�������), objAppPages(), lngӤ�����, , lng��Ŀid) Then
        On Error GoTo errH
        If lng������� = 0 Then
            strSQL = "Select ����ҽ����¼_�������.Nextval as ������� From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ApplyInPacs")
            lng������� = Val(rsTmp!�������)
        End If
        On Error Resume Next
        err.Clear
        If UBound(objAppPages) >= 0 Then
            If err.Number = 0 Then
                On Error GoTo errH
                ReDim Preserve arrSQL(0)
                lngҽ����� = GetMaxAdviceNO(lng����ID, lng��ҳID, lngӤ�����)     '����ҽ����¼.��ţ�����
                For i = 0 To UBound(objAppPages)
                    If lngҽ��ID = 0 Or objAppPages(i).blnIsModify = True Then
                        '���� zl_AdviceCheck �������
                        strժҪ = ""
                        strժҪ = gclsInsure.GetItemInfo(Val(rsPati!���� & ""), lng����ID, 0, "", 0, "", CStr(objAppPages(i).lngProjectId))
                        objAppPages(i).strAbstract = strժҪ
                        strTmp = objAppPages(i).strPartMethod
                        If strTmp <> "" Then
                            For k = 0 To UBound(Split(strTmp, "|"))
                                str��λ = Split(Split(strTmp, "|")(k), ";")(0)
                                strTmp���� = Split(Split(strTmp, "|")(k), ";")(1)
                                For j = 0 To UBound(Split(strTmp����, ","))
                                    str���� = Split(strTmp����, ",")(j)
                                    strExtra = strExtra & "," & str��λ & ":" & str����
                                Next
                            Next
                            strExtra = "||0||0||" & Mid(strExtra, 2) & "||0"
                        End If
                        strExtra = strժҪ & strExtra
                        strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as ��� From Dual"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", 2, lng����ID, lng��ҳID, Val(rsPati!���� & ""), 1, "D", objAppPages(i).lngProjectId, _
                           objAppPages(i).lngRequestRoomId, UserInfo.����, IIF(objAppPages(i).lngExeRoomId <= 0, 0, objAppPages(i).lngExeRoomId), IIF(objAppPages(i).lngExeRoomId <= 0, "5", objAppPages(i).lngExeRoomType), 0, 0, strExtra)
                        
                        If Not rsTmp.EOF Then
                            strMsg = NVL(rsTmp!���)
                            If strMsg <> "" Then
                                Select Case Val(Split(strMsg, "|")(0))
                                Case 1 '��ʾ
                                    If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        strMsg = "": Exit Function
                                    End If
                                Case 2 '��ֹ
                                    MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                                    strMsg = "": Exit Function
                                End Select
                                strMsg = ""
                            End If
                        End If
                        
                        blnDo = GetPacsAdviceSQLData(objAppPages(i), 0, lng�������, lng����ID, lng��ҳID, lngӤ�����, "", lng����id, objVBA, objScript, rsDefine, _
                            strRISDel, strRISAdd, arrSQL, lng��ҽ��ID, strTabAdvice, strTxtAdvice, lngǰ��ID, lngҽ�����)
                        If Not blnDo Then Exit Function
                        
                        strItems = objAppPages(i).lngProjectId & ":" & objAppPages(i).lngExeRoomId
                        'ҽ��������
                        If gintҽ������ = 2 Then bln���Ѷ��� = True
                        strMsg = CheckAdviceInsure(Val(rsPati!���� & ""), bln���Ѷ���, lng����ID, Val(rsPati!�������� & ""), "", strItems, Left(strTxtAdvice, 50))
                        If strMsg <> "" Then
                            If gintҽ������ = 1 Then
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                                If vMsg = vbIgnore Then bln���Ѷ��� = False
                            ElseIf gintҽ������ = 2 Then
                                MsgBox strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, gstrSysName
                                Exit Function
                            End If
                            strMsg = ""
                        End If
                        If Val(rsPati!���� & "") <> 0 Then
                            If gclsInsure.GetCapability(supportʵʱ���, lng����ID, Val(rsPati!���� & "")) Then
                                If MakePriceRecord���뵥("22", lng����ID, lng��ҳID, strTabAdvice, strItems, rsPati!�ѱ� & "", lng����id, rsPrice) Then
                                    If Not gclsInsure.CheckItem(Val(rsPati!���� & ""), 1, 0, rsPrice) Then
                                        MsgBox "ҽ�������δͨ(ִ��Insure.CheckItem�ӿ�)�������´��PACS���뵥���ܱ��档", vbInformation, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                        strTabAdvice = ""
                    End If
                Next
                
                If Not SavePacsData(0, Mid(strRISDel, 2), Mid(strRISAdd, 2), arrSQL, lng��ҽ��ID, lngAdviceID) Then
                     Exit Function
                End If
                ApplyInPacs = lng�������
                Call ZLHIS_CIS_001(clsMipModule, lng����ID, rsPati!���� & "", rsPati!סԺ�� & "", , IIF(lng�������� = 1, 1, 2), lng��ҳID, lng����ID, , lng����id, "", , rsPati!��ǰ���� & "", _
                    lngAdviceID, 0, 1, "D", "", UserInfo.����, Format(objAppPages(0).strRequestTime, "yyyy-MM-dd HH:mm:ss"), lng����id, "", , , "")
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ApplyOutPacs(frmParent As Object, ByRef lng������� As Long, ByVal lng����ID As Long, ByVal str�Һŵ� As String, ByVal lngҽ��ID As Long, ByVal lng�Һſ���ID As Long, _
    ByRef objVBA As Object, ByRef objScript As Object, ByRef rsDefine As ADODB.Recordset, ByVal blnMoved As Boolean, Optional ByVal lng��Ŀid As Long, Optional ByVal lngǰ��ID As Long) As Long
'���ܣ����ü�����뵥
'������lngҽ��ID=�޸����뵥ʱ��ǰ�е�ҽ��ID,lng������� =��ǰ�޸��е��������
'       lng�Һſ���ID �Һ�ִ�п���ID
'       objVBA objScript rsDefine VB����ͼ�¼�����ڲ���ҽ�������ı���blnMoved �����Ƿ��Ѿ�ת����
    Dim objPacspplication As New clsPacsApplication
    Dim objAppPages()  As clsApplicationData
    Dim rsPati As Recordset
    Dim i As Long, j As Long, k As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim strMsg As String
    Dim strExtra As String
    Dim strTmp As String
    Dim str��λ As String, str���� As String
    Dim strTmp���� As String
    Dim objTmp As New clsApplicationData
    Dim strժҪ As String
    Dim strItems As String
    Dim strRISDel As String
    Dim strRISAdd As String 'RIS��������ʽ��ҽ��ID:������ĿID,....
    Dim arrSQL() As String
    Dim lng��ҽ��ID As Long  '����ҽ��ID����ֵ���˷ѣ�������ύ����ʱ�������ҽ��ID
    Dim blnDo As Long
    Dim strTabAdvice As String 'ҽ����Ϣ��ʱ��
    Dim strTxtAdvice As String 'ҽ������
    Dim bln���Ѷ��� As Boolean
    Dim vMsg As VbMsgBoxResult
    Dim rsPrice As ADODB.Recordset
    Dim lngҽ����� As Long

    On Error GoTo errH
    'ִ�в���(�ű����)�����˿���
    strSQL = "Select A.����,A.�Ա�,A.����,B.�����,B.סԺ��,B.������,a.ID as �Һ�ID," & _
        " B.����,B.��������,C.���� as ִ�в���,A.�Ǽ�ʱ��,b.�ѱ�" & _
        " From ���˹Һż�¼ A,������Ϣ B,���ű� C" & _
        " Where A.NO(+)=[2] And a.��¼����(+)=1 And a.��¼״̬(+)=1 And B.����ID=[1]" & _
        " And A.����ID(+)=B.����ID And A.ִ�в���ID=C.ID(+)"
    If blnMoved Then
        strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
    End If
    
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ApplyOutPacs", lng����ID, str�Һŵ�)

    If rsPati.RecordCount = 0 Then
        MsgBox "δ����ȷ��ȡ������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    '��ʼ��������뵥����
    Call objPacspplication.InitComponents(lng�Һſ���ID, frmParent)
    Call objTmp.MakePacsData(lng�������, objAppPages())
    If objPacspplication.ShowApplicationForm(lng����ID, 1, Val(rsPati!�Һ�ID & ""), 0, IIF(lng������� = 0, lngҽ��ID, lng�������), objAppPages(), , , lng��Ŀid) Then
        On Error GoTo errH
        If lng������� = 0 Then
            strSQL = "Select ����ҽ����¼_�������.Nextval as ������� From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ApplyOutPacs")
            lng������� = Val(rsTmp!�������)
        End If
        On Error Resume Next
        err.Clear
        If UBound(objAppPages) >= 0 Then
            If err.Number = 0 Then
                On Error GoTo errH
                ReDim Preserve arrSQL(0)
                lngҽ����� = GetMaxAdviceNO(lng����ID, 0, 0)
                For i = 0 To UBound(objAppPages)
                    If lngҽ��ID = 0 Or objAppPages(i).blnIsModify = True Then
                        '���� zl_AdviceCheck �������
                        strժҪ = ""
                        strժҪ = gclsInsure.GetItemInfo(Val(rsPati!���� & ""), lng����ID, 0, "", 0, "", CStr(objAppPages(i).lngProjectId))
                        objAppPages(i).strAbstract = strժҪ
                        strTmp = objAppPages(i).strPartMethod
                        If strTmp <> "" Then
                            For k = 0 To UBound(Split(strTmp, "|"))
                                str��λ = Split(Split(strTmp, "|")(k), ";")(0)
                                strTmp���� = Split(Split(strTmp, "|")(k), ";")(1)
                                For j = 0 To UBound(Split(strTmp����, ","))
                                    str���� = Split(strTmp����, ",")(j)
                                    strExtra = strExtra & "," & str��λ & ":" & str����
                                Next
                            Next
                            strExtra = "||0||0||" & Mid(strExtra, 2) & "||0"
                        End If
                        strExtra = strժҪ & strExtra
                        strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as ��� From Dual"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", 1, lng����ID, Val(rsPati!�Һ�ID & ""), Val(rsPati!���� & ""), 1, "D", objAppPages(i).lngProjectId, _
                           objAppPages(i).lngRequestRoomId, UserInfo.����, objAppPages(i).lngExeRoomId, IIF(objAppPages(i).lngExeRoomId <= 0, "5", objAppPages(i).lngExeRoomType), 0, 0, strExtra)
                        
                        If Not rsTmp.EOF Then
                            strMsg = NVL(rsTmp!���)
                            If strMsg <> "" Then
                                Select Case Val(Split(strMsg, "|")(0))
                                Case 1 '��ʾ
                                    If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        strMsg = "": Exit Function
                                    End If
                                Case 2 '��ֹ
                                    MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                                    strMsg = "": Exit Function
                                End Select
                                strMsg = ""
                            End If
                        End If
                        
                        blnDo = GetPacsAdviceSQLData(objAppPages(i), 1, lng�������, lng����ID, 0, 0, str�Һŵ�, lng�Һſ���ID, objVBA, objScript, rsDefine, _
                            strRISDel, strRISAdd, arrSQL, lng��ҽ��ID, strTabAdvice, strTxtAdvice, lngǰ��ID, lngҽ�����)
                        If Not blnDo Then Exit Function
                        
                        strItems = objAppPages(i).lngProjectId & ":" & objAppPages(i).lngExeRoomId
                        'ҽ��������
                        If gintҽ������ = 2 Then bln���Ѷ��� = True
                        strMsg = CheckAdviceInsure(Val(rsPati!���� & ""), bln���Ѷ���, lng����ID, 1, "", strItems, Left(strTxtAdvice, 50))
                        If strMsg <> "" Then
                            If gintҽ������ = 1 Then
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                                If vMsg = vbIgnore Then bln���Ѷ��� = False
                            ElseIf gintҽ������ = 2 Then
                                MsgBox strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, gstrSysName
                                Exit Function
                            End If
                            strMsg = ""
                        End If
                        
                        If Val(rsPati!���� & "") <> 0 Then
                            If gclsInsure.GetCapability(supportʵʱ���, lng����ID, Val(rsPati!���� & "")) Then
                                If MakePriceRecord���뵥("21", lng����ID, Val(rsPati!�Һ�ID & ""), strTabAdvice, strItems, rsPati!�ѱ� & "", lng�Һſ���ID, rsPrice) Then
                                    If Not gclsInsure.CheckItem(Val(rsPati!���� & ""), 0, 0, rsPrice) Then
                                        MsgBox "ҽ�������δͨ(ִ��Insure.CheckItem�ӿ�)�������´��PACS���뵥���ܱ��档", vbInformation, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                        strTabAdvice = ""
                    End If
                Next
                If Not SavePacsData(1, Mid(strRISDel, 2), Mid(strRISAdd, 2), arrSQL, lng��ҽ��ID) Then
                     Exit Function
                End If
                ApplyOutPacs = lng�������
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SavePacsData(ByVal intType As Integer, ByVal strRISDel As String, ByVal strRISAdd As String, ByRef arrSQL() As String, ByVal lng��ҽ��ID As Long, Optional ByRef lngҽ��ID As Long) As Boolean
'���ܣ��ύ����
'������  intType 0��סԺҽ����1������ҽ��,lngҽ��ID������һ��ҽ����ID
    Dim varTmp As Variant
    Dim strTmp As String
    Dim i As Long, j As Long
    Dim varID As Variant
    Dim blnStartTran As Boolean
    Dim lngTmp As Long
    
    On Error GoTo errH
    
    '������ʵ��ҽ��IDֵ
    For i = 1 To lng��ҽ��ID
        j = zlDatabase.GetNextID("����ҽ����¼")
        If i = 1 Then
            strTmp = j
        Else
            strTmp = strTmp & "," & j
        End If
    Next

    varID = Split(strTmp, ",")

    For i = 1 To UBound(arrSQL)
        strTmp = arrSQL(i)
        If InStr(strTmp, "<FAKEID>") > 0 Then
            j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
            strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
            If InStr(strTmp, "<FAKEID>") > 0 Then '����滻����
                j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
            End If
            arrSQL(i) = strTmp
        End If
    Next
    
    varTmp = Split(strRISDel, ",")
    If strRISDel <> "" Then
        On Error Resume Next
        For i = 0 To UBound(varTmp)
            strTmp = varTmp(i)
            If 0 <> gobjRis.HISSchedulingEx(Val(Split(strTmp, ":")(0)), Val(Split(strTmp, ":")(1))) Then
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ����β���ɾ�����޸����Ѿ�ԤԼҽ����������Ӱ����Ϣϵͳ�ӿ�(HISSchedulingEx)ȡ��ϢԤԼδ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
        Next
        err.Clear: On Error GoTo 0
    End If
    On Error GoTo errH
    Call gcnOracle.BeginTrans: blnStartTran = True
        For i = 1 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "PACS���뵥����ҽ��")
        Next
    Call gcnOracle.CommitTrans: blnStartTran = False
    
    SavePacsData = True
    
    lngҽ��ID = Val(varID(0))
    
    If strRISAdd <> "" Then
        varTmp = Split(strRISAdd, ",")
        j = IIF(intType = 1, 1, 2)
        On Error Resume Next
        For i = 0 To UBound(varTmp)
            strTmp = varTmp(i)
            
            lngTmp = Val(Split(strTmp, ":")(0))
            lngTmp = Val(varID(lngTmp - 1)) '��Ϊ��ʵ��ҽ��
            
            Call gobjRis.HISScheduling(j, lngTmp, Val(Split(strTmp, ":")(1)))
        Next
        err.Clear: On Error GoTo 0
    End If
    Exit Function
errH:
    If blnStartTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPacsAdviceSQLData(ByRef adviceInf As clsApplicationData, ByVal intType As Integer, ByVal lng������� As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lngӤ����� As Long, ByVal str�Һŵ� As String, ByVal lng����id As Long, ByRef objVBA As Object, ByRef objScript As Object, ByRef rsDefine As ADODB.Recordset, _
    ByRef strRISDel As String, ByRef strRISAdd As String, ByRef arySql() As String, ByRef lng��ҽ��ID As Long, ByRef strTabAdvice As String, ByRef strTxtAdvice As String, ByVal lngǰ��ID As Long, ByRef lngҽ����� As Long) As Boolean
'------------------------------------------------
'���ܣ���ȡ����ҽ����SQL��ע����ʵ��ҽ��ID����ִ������ʱ�Ų�����
'������ intType 0��סԺҽ����1������ҽ����arySql����SQL��strTabAdvice ҽ����Ϣ���ɵ��ٱ�SQL��ѯ
'       lng����ID סԺ���� �����ת�����ˣ���Ϊԭ����ID��������� �Һ�ִ�п�ID
'       strRISDel ��Ҫȡ��RISԤԼ����Ϣ��strRISAdd ��Ҫ����ԤԼ��RIS��Ϣ��"21341,12343,..."
'���أ�true�������false�˳�
'------------------------------------------------
    Dim i As Long, j As Long
    Dim arrSQL() As String
    Dim int�� As Integer '������־
    
    Dim strҽ������ As String
    Dim str��Ŀ���� As String
    
    Dim lngҽ��ID As Long
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim lng�Ƽ����� As Long
    Dim blnRISԤԼ As Boolean
    Dim strҽ��IDs As String
    Dim strTmp As String
    Dim varID As Variant
    Dim rsData As ADODB.Recordset
    Dim str��λ As String
    Dim strTmp����  As String
    Dim str���� As String
    Dim lngTmpID As Long
    
    Dim arrAppend As Variant
    Dim blnIsDel As Boolean
    Dim lng���� As Long
    Dim lng���� As Long
    Dim lngҪ��ID As Long
    Dim str�������� As String
    Dim strDiag As String

    On Error GoTo errH


    '��ȡҽ��������������
    If adviceInf Is Nothing Then Exit Function
    
    strSQL = "Select ����,�Ƽ����� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", adviceInf.lngProjectId)
    If Not rsTmp.EOF Then
        str��Ŀ���� = rsTmp!���� & ""
        lng�Ƽ����� = Val(rsTmp!�Ƽ����� & "")
    End If
    
    lng��ҽ��ID = lng��ҽ��ID + 1
    lngҽ��ID = lng��ҽ��ID ' zlDatabase.GetNextID("����ҽ����¼")        '��ȡҽ��ID
    
    ReDim Preserve arrSQL(1)
    
    If adviceInf.lngUpdateAdviceId <> 0 Then
        '�޸�ҽ����ɾ�������²���
        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & adviceInf.lngUpdateAdviceId & ",1)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    End If
    
    If intType = 0 Then
        '��֯��ҽ���������
        int�� = IIF(adviceInf.blnIsAdditionalRec = True, 2, IIF(adviceInf.blnIsPriority, 1, 0))
    Else
        '��֯��ҽ���������
        int�� = IIF(adviceInf.blnIsPriority, 1, 0)
    End If
    
    strҽ������ = FormatAdviceContext(str��Ŀ����, adviceInf.strPartMethod, adviceInf.lngExeType, objVBA, objScript, rsDefine)
    strTxtAdvice = strҽ������
    lngҽ����� = lngҽ����� + 1
    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(<FAKEID>" & lngҽ��ID & "</FAKEID>,NULL," & lngҽ����� & "," & IIF(intType = 0, 2, 1) & _
                    "," & lng����ID & "," & IIF(intType = 0, lng��ҳID & "," & lngӤ�����, "Null,0") & ",1,1,'D'," & adviceInf.lngProjectId & _
                    ",NULL,NULL,NULL,1,'" & strҽ������ & "',Null,Null,'һ����',NULL,NULL,NULL,NULL," & lng�Ƽ����� & "," & _
                    IIF(adviceInf.lngExeRoomId <= 0, "Null", adviceInf.lngExeRoomId) & _
                    "," & IIF(adviceInf.lngExeRoomId <= 0, "5", adviceInf.lngExeRoomType) & "," & int�� & _
                    ",to_date('" & Format(adviceInf.strStartExeTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                    ",NULL," & lng����id & "," & adviceInf.lngRequestRoomId & ",'" & UserInfo.���� & "'," & _
                    "to_date('" & Format(adviceInf.strRequestTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                    IIF(intType = 0, "null", "'" & str�Һŵ� & "'") & "," & ZVal(lngǰ��ID) & ",Null," & adviceInf.lngExeType & _
                    ",NULL," & IIF(adviceInf.strAbstract = "", "Null", "'" & adviceInf.strAbstract & "'") & ",'" & UserInfo.���� & "',Null,NULL,NULL,NULL," & lng������� & ")"
    
    strTabAdvice = "Select " & lngҽ��ID & " As ID," & lngҽ����� & " As ���,-null As ���id, 'D' As �������," & adviceInf.lngProjectId & " As ������Ŀid," & adviceInf.lngProjectId & " As ������Ŀid," & vbNewLine & _
        " 1 As ����, 0 As ����, null As �걾��λ,null As ��鷽��," & adviceInf.lngExeType & " As ִ�б��," & lng�Ƽ����� & " As �Ƽ�����,null As ��������," & _
        IIF(adviceInf.lngExeRoomId <= 0, "5", adviceInf.lngExeRoomType) & " As ִ������," & adviceInf.lngExeRoomId & " As ִ�п���id From Dual"
    
    strDiag = adviceInf.strDiagnoseId
    If intType = 1 Then
        If strDiag = "" Then
            '���ﲡ�������뵥ʱȡһ�����������Ĭ�Ϲ���
            strSQL = "Select a.ID From ������ϼ�¼ A,���˹Һż�¼ b Where a.����id=b.����id and a.��ҳid =b.id and b.no=[1] and  a.��¼��Դ = 3  order by a.�������,a.��ϴ���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str�Һŵ�)
            If Not rsTmp.EOF Then
                strDiag = rsTmp!ID & ""
            End If
        End If
    End If
    
    '��Ϲ�����Ϣ
    If strDiag <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Insert(<FAKEID>" & lngҽ��ID & "</FAKEID>,'" & strDiag & "')"
    End If
    
    '��֯��λ�������
    For i = 0 To UBound(Split(adviceInf.strPartMethod, "|")) '��λ1;����1,����2,����3|��λn;����1,����2,����3---
        str��λ = Split(Split(adviceInf.strPartMethod, "|")(i), ";")(0)
        strTmp���� = Split(Split(adviceInf.strPartMethod, "|")(i), ";")(1)
        For j = 0 To UBound(Split(strTmp����, ","))
            lngҽ����� = lngҽ����� + 1     '����ҽ����¼.��ţ�����
            str���� = Split(strTmp����, ",")(j)
            lng��ҽ��ID = lng��ҽ��ID + 1
            lngTmpID = lng��ҽ��ID
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(<FAKEID>" & lngTmpID & "</FAKEID>,<FAKEID>" & lngҽ��ID & "</FAKEID>," & lngҽ����� & _
                 IIF(intType = 0, ",2," & lng����ID & "," & lng��ҳID & "," & lngӤ����� & ",", ",1," & lng����ID & ",NULL,0,") & _
                 "1,1,'D'," & adviceInf.lngProjectId & ",NULL,NULL,NULL,1," & _
                 "'" & str��Ŀ���� & "',NULL," & _
                 "'" & str��λ & "','һ����',NULL,NULL,NULL,NULL," & lng�Ƽ����� & "," & _
                 IIF(adviceInf.lngExeRoomId <= 0, "null", adviceInf.lngExeRoomId) & "," & IIF(adviceInf.lngExeRoomId <= 0, "5", adviceInf.lngExeRoomType) & "," & int�� & ",to_date('" & Format(adviceInf.strStartExeTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),NULL," & _
                 lng����id & "," & adviceInf.lngRequestRoomId & _
                 ",'" & UserInfo.���� & "',to_date('" & Format(adviceInf.strRequestTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                 IIF(intType = 0, "NULL", "'" & str�Һŵ� & "'") & "," & ZVal(lngǰ��ID) & ",'" & str���� & "'," & adviceInf.lngExeType & ",NULL,NULL,'" & UserInfo.���� & "',NULL,NULL,NULL,NULL," & lng������� & ")"
            
            strTabAdvice = strTabAdvice & " Union All " & _
                "Select " & lngTmpID & " As ID," & lngҽ����� & " As ���," & lngҽ��ID & " As ���id, 'D' As �������," & adviceInf.lngProjectId & " As ������Ŀid," & adviceInf.lngProjectId & " As ������Ŀid," & vbNewLine & _
                " 1 As ����, 0 As ����,'" & str��λ & "' As �걾��λ,'" & str���� & "' As ��鷽��," & adviceInf.lngExeType & " As ִ�б��," & lng�Ƽ����� & " As �Ƽ�����,null As ��������," & _
                IIF(adviceInf.lngExeRoomId <= 0, "5", adviceInf.lngExeRoomType) & " As ִ������," & adviceInf.lngExeRoomId & " As ִ�п���id From Dual"
            
        Next
    Next
    lngҽ����� = lngҽ����� + 1
    
    '��֯���븽��������
    blnIsDel = False
    If adviceInf.strRequestAffix <> "" Then
        arrAppend = Split(adviceInf.strRequestAffix, "|")
        For j = 0 To UBound(arrAppend)
            strTmp = "": lng���� = 0: lng���� = 0: lngҪ��ID = 0
            If InStr(adviceInf.strRequestAffixCfg, Split(arrAppend(j), ":")(0) & ":") > 0 Then
                strTmp = Mid(adviceInf.strRequestAffixCfg, InStr(adviceInf.strRequestAffixCfg, Split(arrAppend(j), ":")(0) & ":") + Len(Split(arrAppend(j), ":")(0) & ":"))
                If InStr(strTmp, "|") > 0 Then
                    strTmp = Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                End If
                If strTmp <> "" Then
                    lng���� = Split(strTmp, ",")(0)
                    lng���� = Split(strTmp, ",")(1)
                    lngҪ��ID = Val(Split(strTmp, ",")(2))
                End If
            End If
            strTmp = arrAppend(j)
            str�������� = Replace(strTmp, Mid(strTmp, 1, InStr(strTmp, ":")), "")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(<FAKEID>" & lngҽ��ID & "</FAKEID>,'" & Split(arrAppend(j), ":")(0) & "'," & lng���� & "," & lng���� & "," & ZVal(lngҪ��ID) & ",'" & str�������� & "'," & IIF(Not blnIsDel, 1, 0) & ")"
            blnIsDel = True
        Next
    End If
    
    If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
        blnRISԤԼ = True
        strSQL = "select a.ID,b.ԤԼid from ����ҽ����¼ a,RIS���ԤԼ b where a.id=b.ҽ��id and a.id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", adviceInf.lngUpdateAdviceId)
    End If
    If blnRISԤԼ Then
        If Not rsTmp.EOF Then
            strRISDel = strRISDel & "," & Val(rsTmp!ID & "") & ":" & Val(rsTmp!ԤԼid & "")
        End If
    End If
     
    For i = 1 To UBound(arrSQL)
        ReDim Preserve arySql(UBound(arySql) + 1)
        arySql(UBound(arySql)) = arrSQL(i)
    Next
    
    If blnRISԤԼ Then
        strRISAdd = strRISAdd & "," & lngҽ��ID & ":" & adviceInf.lngProjectId
    End If
    
    GetPacsAdviceSQLData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FormatAdviceContext(ByVal strAdvicePro As String, _
    ByVal strAdvicePart As String, ByVal lngExeType As Long, ByRef objVBA As Object, ByRef objScript As Object, ByRef rsDefine As ADODB.Recordset) As String
'����ϵͳ������������ʽ��ҽ������

    Dim strReturn As String
    Dim i As Long
    Dim Arr��λ As Variant
    Dim str��λ���� As String
    Dim strTmp As String
    
    If objVBA Is Nothing Then
        On Error Resume Next
        Set objVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not objVBA Is Nothing Then
            objVBA.Language = "VBScript"
            Set objScript = New clsScript
            objVBA.AddObject "clsScript", objScript, True
        End If
    End If
    On Error GoTo errH
    
    rsDefine.Filter = "�������='D'"
    If rsDefine.RecordCount > 0 Then
        strReturn = rsDefine!ҽ������ & ""
    End If
    strTmp = ""
    '��ȡ��λ����
    'ǰ:��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����
    '��:��λ��1(������1,������2),��λ��2(������1,������2)-----
    If strAdvicePart <> "" Then
        strTmp = strAdvicePart
        Arr��λ = Split(Split(strTmp, Chr(9))(0), "|")
        strTmp = ""
        For i = 0 To UBound(Arr��λ)
            strTmp = strTmp & "," & Split(Arr��λ(i), ";")(0) & "(" & Split(Arr��λ(i), ";")(1) & ")"
        Next
        strTmp = Mid(strTmp, 2)
    End If
    str��λ���� = strTmp
    
    If strReturn = "" Then
        strReturn = strAdvicePro & "," & _
                            Decode(lngExeType, 1, ",����ִ��", 2, ",����ִ��", "") & IIF(strAdvicePart <> "", ":" & str��λ����, "")
    Else
        If InStr(strReturn, "[�����Ŀ]") > 0 Then
            strReturn = Replace(strReturn, "[�����Ŀ]", _
                                            """" & strAdvicePro & Decode(lngExeType, 1, ",����ִ��", 2, ",����ִ��", "") & _
                                            """")
        End If

        '�滻��λ����
        If InStr(strReturn, "[��鲿λ]") > 0 Then
            strReturn = Replace(strReturn, "[��鲿λ]", _
                                            """" & str��λ���� & """")
        End If
        strReturn = objVBA.Eval(strReturn)
    End If
    FormatAdviceContext = strReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBloodState(ByVal int���� As Integer, ByVal int��Ѫ As Integer) As String
'���ܣ���ȡ���´����Ѫҽ�������״̬
'������int���� 0���ǽ�����1��������int��Ѫ 0����Ѫ ҪѪ�ⷢѪ��1����Ѫ ��һ��֪ͨ����
    Dim strTmp As String
    Dim str���״̬ As String
    
    strTmp = IIF(gbln��Ѫ�ּ�����, "1", "0")
    strTmp = strTmp & IIF(gblnѪ��ϵͳ, "1", "0")
    strTmp = strTmp & int���� & int��Ѫ
    Select Case strTmp
    Case "1000", "1001", "1100"
        str���״̬ = "1"
    Case "0100", "0110", "1110"
        str���״̬ = "4"
    End Select
    
    GetBloodState = str���״̬
End Function

Public Function CanAutoExeItem(ByVal lng����id As Long, ByVal str��� As String, ByVal str�������� As String, ByVal intִ�з��� As Integer) As Boolean
'���ܣ��жϵ�ǰ��Ŀ�ǲ��ǿ����Զ����
    Dim varArr As Variant
    Dim i As Long
    Dim blnResult As Boolean
    Dim lngResult As Long
    Dim strҽ����� As String
    Dim strPar As String
    
    strPar = zlDatabase.GetPara("����ִ���Զ����ҽ�����", glngSys, pסԺҽ������, , , , , lng����id)
    
    strҽ����� = str��� & IIF("" = str��������, 0, str��������) & intִ�з���
    Select Case strҽ�����
        Case "E21"
            lngResult = 0 '��Һ
        Case "E22"
            lngResult = 1 'ע��
        Case "E24"
            lngResult = 2 '�ڷ�
        Case "E60"
            lngResult = 3 '�ɼ�
        Case "E13", "E15"
            lngResult = 4 '��������
        Case "E00"
            lngResult = 5 '��ͨ����
        Case "E50"
            lngResult = 6 '��������
        Case "E20"
            lngResult = 7 '������ҩ;��
        Case Else
            lngResult = 8 '����ҽ��
    End Select
    
    If InStr("," & strPar & ",", "," & lngResult & ",") > 0 Or strPar = "*" Then
        blnResult = True
    End If
    
    If blnResult Then
        If (strҽ����� = "E13" Or strҽ����� = "E15") And Mid(gstrҽ���˶�, 2, 1) = "1" Then
            blnResult = False
        ElseIf (str��� = "K" Or str��� = "E" And str�������� = "8") And Mid(gstrҽ���˶�, 1, 1) = "1" Then
            blnResult = False
        End If
    End If
    
    CanAutoExeItem = blnResult
End Function

Public Function RevokeOutAdvice(ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, ByVal str�Һŵ� As String, ByVal str���� As String, ByVal str����� As String, ByVal lng�Һſ���ID As Long, ByVal lngҽ��ID As Long, _
    ByVal lngҽ��״̬ As Long, ByVal str������� As String, ByVal str�������� As String, ByVal lng���״̬ As Long, ByVal str����ʱ�� As String, _
    ByVal lngǩ�� As Long, ByVal lngType As Long, ByVal strҽ������ As String, ByVal blnMoved As Boolean, ByRef clsMip As Object, ByVal int���� As Integer) As Boolean
'���ܣ�����ҽ�����ϲ���
'������lngType 1����ʾһ����ҩ��2�������У�
'      blnMoved �����Ƿ�ת��
'      clsMip ��Ϣ����

    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lng֤��ID As Long, lngǩ��id As Long
    Dim strSign As String, intRule As Integer
    Dim strSource As String, strIDs As String, blnDo As Boolean
    Dim strTimeStamp As String, blnTran As Boolean, strErr As String, strTimeStampCode As String
    Dim lngRISҽ��ID As Long
    Dim strAdvice��Ѫ As String
    Dim arrSQL As Variant
    Dim i As Long
    
    '����Ƿ��������
    If lngҽ��ID = 0 Then
        MsgBox "�ò���û��ҽ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lngҽ��״̬ <> 8 Then
        MsgBox "��ǰѡ���ҽ����δ���ͻ��Ѿ����ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If str������� = "K" And gblnѪ��ϵͳ Then
        If InitObjBlood(True) Then
            strAdvice��Ѫ = lngҽ��ID
        End If
    End If
    
    If str������� = "K" And gblnѪ��ϵͳ And lng���״̬ = 2 Then
        On Error GoTo errH
        strSQL = "Select Nvl(ִ�з���,0) as ִ�з��� from ����ҽ����¼ A, ������ĿĿ¼ B  where A.���ID  = [1] and A.������ĿID = B.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������Ŀ��ִ�з���", lngҽ��ID)
        If rsTmp.RecordCount > 0 Then
            If Val(rsTmp!ִ�з���) = 0 Then
                MsgBox "�������ϵ���Ѫҽ���Ѿ������Ѫ������ֱ������ҽ������Ҫ����������Ѫ����ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        On Error GoTo 0
    End If
    
    '���з���ת������������
    If zlDatabase.DateMoved(str����ʱ��) Then
        If MovedBySend(lngҽ��ID, 0, 1) Then
            MsgBox "��ҽ���ķ����Ѿ�ȫ���򲿷�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����ǩ��������ʾ
    If lngǩ�� = 1 Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ��������ϡ�", vbInformation, gstrSysName
            Else
                MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ���������ϡ�", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If gobjESign.CertificateStoped(UserInfo.����) = False Then strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
    End If
    
    '�������ҽ����Ӧ�ķ��ý������
    If Not CheckAdviceBalanceRevoke(lngҽ��ID) Then Exit Function
    
    '����˼��ʷ��ü��
    If InStr(GetInsidePrivs(p����ҽ���´�), "��������˼���ҽ��") = 0 Then
        If Not CheckAdviceBillingRevoke(lngҽ��ID) Then
            MsgBox "Ҫ����ҽ���Ķ�Ӧ���ʻ��۷����Ѿ���ˣ��������ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If lngType = 1 Then
        If MsgBox("����һ����ҩ��ҽ������һ�����ϣ�ȷʵҪ������" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If MsgBox("ȷʵҪ����ҽ��""" & strҽ������ & """��" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    arrSQL = Array()
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_����(" & lngҽ��ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_�������ҽ��_delete(" & lngҽ��ID & ")"   '����ҽ������ʱɾ���������ҽ�� �ж�Ӧ�ļ�¼
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_����Σ��ֵҽ��_Update(3,null," & lngҽ��ID & ")"   'ɾ��Σ��ֵ��Ӧ��ϵ
    
    '����ʱ���е���ǩ��
    If strSign <> "" Then
        If gobjESign.CertificateStoped(UserInfo.����) = False Then
            '��ȡǩ��ҽ��Դ��
            strIDs = lngҽ��ID
            intRule = ReadAdviceSignSource(4, lng����ID, str�Һŵ�, strIDs, 0, blnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "���ܶ�ȡ��Ҫ���ϵ���ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
                strSign = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSign
            Else
                Exit Function
            End If
        End If
    End If
    
    'RIS���ˣ�����ʧ�����˳�
    If InStr(",D,F,", str�������) > 0 Or str������� = "E" And lngType <> 2 Then '��顢����������
        If HaveRIS(True) Then
            On Error Resume Next
            If gobjRis.HISRollAdvice(lngҽ��ID) <> 1 Then
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISRollAdvice)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            lngRISҽ��ID = lngҽ��ID
            err.Clear: On Error GoTo 0
        End If
    End If
    
    Call CreatePlugInOK(p����ҽ���´�, int����)
    
    '��������ǰ��ҽӿ�
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        strErr = ""
        blnDo = gobjPlugIn.AdviceRevokedBefore(glngSys, p����ҽ���´�, lng����ID, lng�Һ�ID, lngҽ��ID, int����, strErr)
        Call zlPlugInErrH(err, "AdviceRevokedBefore")
        If 0 = err.Number Then '�ӿ�û�г������������жϽӿڵķ���ֵ
            If Not blnDo Then
                MsgBox strErr, vbInformation, gstrSysName
                Exit Function
            End If
        End If
        blnDo = False
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "RevokeOutAdvice"
    Next
    If strAdvice��Ѫ <> "" Then
        If gobjPublicBlood.AdviceOperation(p����ҽ��վ, lngҽ��ID, 4, False, strErr) = False Then
            gcnOracle.RollbackTrans: blnTran = False
            MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    
    If Not (clsMip Is Nothing) Then
        If clsMip.IsConnect Then
            Call ZLHIS_CIS_024(clsMip, lng����ID, str����, , str�����, 1, lng�Һ�ID, lng�Һſ���ID, "", lngҽ��ID, str�������, str��������)
        End If
    End If
    '�������Ϻ���ҽӿ�
    On Error Resume Next
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.AdviceRevoked(glngSys, p����ҽ���´�, lng����ID, lng�Һ�ID, lngҽ��ID, int����)
        Call zlPlugInErrH(err, "AdviceRevoked")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    Call InitObjLis(p����ҽ��վ)
    '����LIS�������뵥
    If Not gobjLIS Is Nothing Then
        If gobjLIS.DelLisApplicationForm(CStr(lngҽ��ID), strErr) = False Then
            MsgBox "ɾ����������ʧ�ܣ�" & strErr, vbInformation, gstrSysName
        End If
    End If
    '�������ݽ���ƽ̨����LIS,PACSȡ�����뵥
    If gobjExchange Is Nothing Then
        On Error Resume Next
        Set gobjExchange = CreateObject("zlExchange.clsExchange")
        If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
        err.Clear: On Error GoTo 0
    End If
    If Not gobjExchange Is Nothing Then
        If str������� = "D" Then
            blnDo = True
        ElseIf str������� = "E" Then
            blnDo = lngType = 2
        End If
        If blnDo Then
            Call gobjExchange.SendMsg(IIF(str������� = "D", 2, 1), "����ID::" & lng����ID & "||��ҳID::0||ҽ��ID::" & lngҽ��ID & "||��������::0")
        End If
    End If
    '����ԤԼ���ķ���
    If str������� = "Z" And str�������� = "2" Then
        Call SvrԤԼ��Ժȡ������(lng�Һ�ID)
    End If
    RevokeOutAdvice = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
        If blnTran Then
        blnTran = False
        'HIS����ع��ٵ���RIS���� lngRISҽ��ID
        If lngRISҽ��ID <> 0 And HaveRIS(True) Then
            strSQL = "Select a.����id, a.��ҳid, a.�Һŵ�, a.��������id, a.ִ�п���id,a.������ĿID, a.������� As ���, b.���ͺ�, a.Id As ҽ��id, Decode(a.�Һŵ�, Null, 2, 1) As ������Դ" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ������ B Where a.Id = b.ҽ��id And a.Id =[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "RevokeOutAdvice", lngRISҽ��ID)
            If Not rsTmp.EOF Then
                Call gobjRis.HISSendAdvice(rsTmp, 1, Val(rsTmp!����ID & ""), 0, rsTmp!�Һŵ� & "", Val(rsTmp!���ͺ� & ""))
            End If
        End If
    End If
End Function

Public Function CheckLISAppAdvice(ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int���� As Integer, ByVal str��� As String, _
    ByVal lng������ĿID As Long, ByVal lng��������ID As Long, ByVal str����ҽ�� As String, ByVal lngִ�п���ID As Long, ByVal lngִ������ As Long, ByVal strժҪ As String) As Boolean
'���ܣ�ҽ�����ݿ�˼�� Zl_Advicecheck
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String
    On Error GoTo errH
    
    strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", int����, lng����ID, lng����ID, int����, 1, str���, lng������ĿID, _
         lng��������ID, str����ҽ��, lngִ�п���ID, lngִ������, 0, 0, strժҪ)
    
    If Not rsTmp.EOF Then
        strMsg = NVL(rsTmp!���)
        If strMsg <> "" Then
            Select Case Val(Split(strMsg, "|")(0))
            Case 1 '��ʾ
                If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    strMsg = "": Exit Function
                End If
            Case 2 '��ֹ
                MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                strMsg = "": Exit Function
            End Select
            strMsg = ""
        End If
    End If
    CheckLISAppAdvice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub DefCommandPlugInPopup(ByRef objBar As Object, ByRef rsBar As ADODB.Recordset)
'���ܣ���ҽ�����Ҽ������˵�
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim objCtl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    '������ť
    rsBar.Filter = "IsInTool=1 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        For i = 1 To rsBar.RecordCount
            Set objControl = objBar.Add(xtpControlButton, rsBar!����ID, rsBar!������)
            objControl.IconId = rsBar!ͼ��ID
            objControl.Parameter = rsBar!������
            objControl.Style = xtpButtonIconAndCaption
            If Val(rsBar!IsGroup) = 1 Then
                objControl.BeginGroup = True
            End If
            rsBar.MoveNext
        Next
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        Set objPopup = objBar.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                objControl.IconId = rsBar!ͼ��ID
                objControl.Parameter = rsBar!������
                objControl.Style = xtpButtonIconAndCaption
                If Val(rsBar!IsGroup) = 1 Then
                    objControl.BeginGroup = True
                End If
                rsBar.MoveNext
            Next
        End With
    End If
End Sub

Public Sub InitCardRsBlood(ByRef rsCard As ADODB.Recordset)
    Set rsCard = New ADODB.Recordset
    With rsCard.Fields
        .Append "��Ѫ����", adInteger '0-��ͨ��1������
        .Append "�ٴ����IDs", adVarChar, 2000 '���ID �����ŷָ�   ������ϼ�¼.ID
        .Append "����", adInteger '0/1
        .Append "��Ѫ����", adVarChar, 500
        .Append "��ѪĿ��", adVarChar, 500
        .Append "��Ѫ����", adInteger
        .Append "������Ѫʷ", adInteger
        .Append "������Ѫ��Ӧʷ", adInteger
        .Append "��Ѫ���ɼ�����ʷ", adInteger
        .Append "�в����", adVarChar, 10 '1/1 ��ʾ:1��1��
        .Append "��Ѫ������", adInteger
        .Append "�Ƿ�ǩ��ͬ����", adInteger, 1, adFldIsNullable
        .Append "�Ƿ�������", adInteger, 1, adFldIsNullable
        .Append "Ԥ����Ѫ����", adVarChar, 500
        .Append "Ѫ��", adInteger
        .Append "RHD", adInteger
        .Append "��Ѫ��ĿID", adBigInt
        .Append "��Ѫִ�п���ID", adBigInt
        .Append "Ԥ����Ѫ��", adDouble
        .Append "��Ѫ;����ĿID", adBigInt
        .Append "��Ѫ;��ִ�п���ID", adBigInt
        .Append "����", adVarChar, 2000  '��Ѫ����¼����ٴ����E��ҽ��������
        .Append "��ע", adVarChar, 2000
        .Append "��Ѫ��������", adVarChar, 500
        .Append "�������ID", adBigInt
        .Append "�ٴ��������", adVarChar, 4000 '������ʾ�������
        .Append "�����", adVarChar, 4000 '��Ŀ�����
        .Append "������Ŀ", adVarChar, 4000 '��Ѫ������Ŀ ��ʽ����ĿID,������,Ѫ��,RH;��ĿID,������,Ѫ��,RH...... (�����뵥ʹ��)
        .Append "����������ĿSQL", adVarChar, 2000
        .Append "������ĿSQL", adVarChar, 2000
        .Append "��Ϲ�����ϢSQL", adVarChar, 2000
        .Append "������ĿSQL", adVarChar, 2000
    End With
    rsCard.CursorLocation = adUseClient
    rsCard.LockType = adLockOptimistic
    rsCard.CursorType = adOpenStatic
    rsCard.Open
End Sub

Public Sub InitCardRsOperate(ByRef rsCard As ADODB.Recordset)
    Set rsCard = New ADODB.Recordset
    With rsCard.Fields
        .Append "�ٴ����IDs", adVarChar, 2000
        .Append "�ٴ��������", adVarChar, 4000 '������ʾ�������
        .Append "�������", adInteger 'סԺ�õ�
        .Append "��������ĿID", adBigInt
        .Append "��������ĿIDs", adVarChar, 2000
        .Append "������ĿID", adBigInt
        .Append "����ִ�п���ID", adBigInt
        .Append "����ִ�п���ID", adBigInt
        .Append "��Чʱ��", adVarChar, 100 '��ʼʱ��
        .Append "����ʱ��", adVarChar, 100 '����ʱ��
        .Append "���븽��", adVarChar, 4000
        .Append "�������ID", adBigInt
        .Append "��������Ҫʱ", adVarChar, 2000 '����������Ҫʱ
    End With
    rsCard.CursorLocation = adUseClient
    rsCard.LockType = adLockOptimistic
    rsCard.CursorType = adOpenStatic
    rsCard.Open
End Sub

Public Sub InitCardRsLIS(ByRef rsCard As ADODB.Recordset)
    Set rsCard = New ADODB.Recordset
    With rsCard.Fields
        .Append "�ٴ����IDs", adVarChar, 2000
        .Append "������Ϣ", adVarChar, 4000
    End With
    rsCard.CursorLocation = adUseClient
    rsCard.LockType = adLockOptimistic
    rsCard.CursorType = adOpenStatic
    rsCard.Open
End Sub

Public Function ShowApply���(frmParent As Object, ByVal lngNo As Long, Optional ByVal lngҽ��ID As Long) As Boolean
'���ܣ��鿴������뵥
    Dim rsTmp As ADODB.Recordset
    Dim objAppPages()  As clsApplicationData
    Dim objTmp As New clsApplicationData
    Dim strSQL As String
    Dim intӤ�� As Integer
    
    On Error GoTo errH
    If lngNo = 0 Then
        strSQL = "Select a.�������,a.Ӥ�� From ����ҽ����¼ A where a.Id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ShowPacsApplication", lngҽ��ID)
        If rsTmp.RecordCount = 0 Then
            MsgBox "û���ҵ���ָ���ļ��ҽ����", vbInformation, gstrSysName
            Exit Function
        Else
            lngNo = Val(rsTmp!������� & "")
            intӤ�� = Val(rsTmp!Ӥ�� & "")
        End If
    Else
        strSQL = "Select a.Ӥ�� From ����ҽ����¼ A where a.������� =[1] and rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ShowPacsApplication", lngNo)
        intӤ�� = Val(rsTmp!Ӥ�� & "")
    End If
    If lngNo = 0 Then
        MsgBox "��ҽ��û�ж�Ӧ���뵥��", vbInformation, gstrSysName
        Exit Function
    End If
    
    Set rsTmp = objTmp.MakePacsData(lngNo, objAppPages(), True)
    Call frmPacsApplication.InitComponents(Val(rsTmp!��������id & ""), frmParent)
    ShowApply��� = frmPacsApplication.ShowApplicationForm(Val(rsTmp!����ID & ""), Val(rsTmp!�������� & ""), Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""), lngNo, objAppPages(), intӤ��, False)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�������() As Long
'���ܣ���ȡ�������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "Select ����ҽ����¼_�������.Nextval as ������� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel")
    Get������� = Val(rsTmp!������� & "")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub InitPlugInRs(ByRef rsDataPlugIn As ADODB.Recordset)
'���ܣ�������¼�������� clsPlugIn.AdviceBeforeSend���������
    Set rsDataPlugIn = New ADODB.Recordset
    rsDataPlugIn.Fields.Append "����ID", adBigInt
    rsDataPlugIn.Fields.Append "����ID", adBigInt
    rsDataPlugIn.Fields.Append "�Һŵ�", adVarChar, 30
    rsDataPlugIn.Fields.Append "ҽ��ID", adBigInt
    rsDataPlugIn.Fields.Append "���ID", adBigInt
    rsDataPlugIn.Fields.Append "�շ�ϸĿID", adBigInt
    rsDataPlugIn.Fields.Append "�ֽ�ʱ��", adVarChar, 40000
    rsDataPlugIn.Fields.Append "����", adInteger
    rsDataPlugIn.Fields.Append "����", adDouble
    rsDataPlugIn.Fields.Append "������λ", adVarChar, 100
    rsDataPlugIn.Fields.Append "����", adDouble
    rsDataPlugIn.Fields.Append "������λ", adVarChar, 100
    rsDataPlugIn.Fields.Append "����", adInteger
    rsDataPlugIn.CursorLocation = adUseClient
    rsDataPlugIn.LockType = adLockOptimistic
    rsDataPlugIn.CursorType = adOpenStatic
    rsDataPlugIn.Open
End Sub

Public Function CheckDocEmpowerEx(ByVal lng������ĿID As Long, ByVal strAppend As String) As Boolean
'���ܣ�������Ա�Ƿ����������Ŀ��ִ��Ȩ
'������strAppend=��ǰ���븽�����д�����,��ʽΪ"��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
    Dim strSQL As String, rsTmp As Recordset
    Dim arrItem As Variant, arrSub As Variant
    Dim strItem As String, i As Integer
    Dim lngID As Long
    Dim strDoc As String
    
    On Error GoTo errH
    If strAppend <> "" Then
        strSQL = "select A.ID from ����������Ŀ A,������������ B where a.����id=b.id and b.����='06' and A.������='����ҽ��'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDocEmpowerEx")
        If rsTmp.RecordCount > 0 Then
            lngID = rsTmp!ID
            arrItem = Split(strAppend, "<Split1>")
            For i = 0 To UBound(arrItem)
                arrSub = Split(arrItem(i), "<Split2>")
                If Val(arrSub(2)) = lngID Then
                    If Trim(arrSub(3)) <> "" Then
                        strDoc = Trim(arrSub(3))
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    If strDoc = "" Then strDoc = UserInfo.����
    strSQL = "Select Count(*) as Ȩ�� From ��Ա����Ȩ�� A,��Ա�� B Where A.��Աid = B.ID And B.����=[1] And A.������Ŀid = [2] And A.��¼���� = 2"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", strDoc, lng������ĿID)
    CheckDocEmpowerEx = Val(rsTmp!Ȩ�� & "") > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPlugInBar(ByVal lngģ�� As Long, ByVal int���� As Integer, rsBar As ADODB.Recordset) As String
'���ܣ���֯��Ҳ����Ĳ˵�����ť
    Dim strFunc As String
    Dim strXML As String
    Call CreatePlugInOK(lngģ��, int����)
    If gobjPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    strFunc = gobjPlugIn.GetFuncNames(glngSys, lngģ��, int����, strXML)
    Call zlPlugInErrH(err, "GetFuncNames")
    err.Clear: On Error GoTo 0
    Call MakePlugInBar(strFunc, strXML, rsBar)
    GetPlugInBar = strFunc
End Function

Private Sub MakePlugInBar(ByVal strFunc As String, ByVal strXML As String, rsBar As ADODB.Recordset)
'���ܣ���֯�˵������ؼ�¼���У�ע����ϰ汾�ļ��ݴ���
'������strFunc �ϰ汾�����д���strXML��������Ϣ�Ĺ��ܴ�
    Dim strM As String
    Dim strB As String
    Dim strP As String
    Dim strTag As String
    Dim i As Long
    Dim strTmp As String
    Dim lngS As Long, lngE As Long
    Dim rsBarFuncID As ADODB.Recordset
    If strXML = "" And strFunc = "" Then Exit Sub
    If strXML = "" And strFunc <> "" Then
        '������ǰ�ϰ汾�ķ�ʽ
        Call InitPlugInRsBar(rsBar)
        Call AddPlugInBarRs(rsBar, strFunc, 1)
        Call AddPlugInBarRs(rsBar, strFunc, 2)
        Call AddPlugInBarRs(rsBar, strFunc, 3)
        Call SetPlugInBar(rsBar, 1)
        Exit Sub
    End If
    
    On Error GoTo errH
    strXML = Trim(strXML)
    '�ݶ�Ϊ200����չ���ܲ������ֹ��ѭ��
    For i = 0 To 200
        lngS = InStr(strXML, "<")
        lngE = InStr(strXML, ">")
        strTag = Mid(strXML, lngS + 1, lngE - lngS - 1)
        If strTag = "menubar" Then
            lngS = lngE
            lngE = InStr(strXML, "</menubar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strM = strM & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "toolbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</toolbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strB = strB & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "popbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</popbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strP = strP & "," & strTmp
            strXML = Mid(strXML, lngE + 9)
        End If
        If strXML = "" Then
            Exit For
        End If
    Next
    If strM = "" Then Exit Sub
    strM = Mid(strM, 2)
    strB = Mid(strB, 2)
    strP = Mid(strP, 2)

    Call InitPlugInRsBar(rsBar)
    Call AddPlugInBarRs(rsBar, strM, 1)
    Call AddPlugInBarRs(rsBar, strB, 2)
    Call AddPlugInBarRs(rsBar, strP, 3)
    Call SetPlugInBar(rsBar, 2)
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddPlugInBarRs(ByRef rsBar As ADODB.Recordset, ByVal strFunc As String, ByVal intType As Integer)
'���ܣ������ܴ�ת��Ϊ��¼����ʽ
'������strFunc ���ܴ���intType ���ܰ�ť������һ�� 1-�˵�����2-��������3-�����
    Dim varFunc As Variant
    Dim i As Long
    Dim strFuncName As String
    Dim blnFirstTool As Boolean
    If strFunc = "" Then Exit Sub
    varFunc = Split(strFunc, ",")
    With rsBar
        For i = 0 To UBound(varFunc)
            strFuncName = varFunc(i)
            .AddNew
            !BarType = intType
            If InStr(strFuncName, "Auto:") > 0 Then
                !IsAuto = 1
                strFuncName = Replace(strFuncName, "Auto:", "")
            Else
                !IsAuto = 0
            End If
            
            If InStr(strFuncName, "InTool:") > 0 Then
                !IsInTool = 1
                strFuncName = Replace(strFuncName, "InTool:", "")
            Else
                !IsInTool = 0
            End If
            If InStr(strFuncName, "|:") > 0 Then
                !IsGroup = 1
                strFuncName = Replace(strFuncName, "|:", "")
            Else
                !IsGroup = 0
                If Not blnFirstTool And !IsInTool = 1 Then
                    '��һ��������ť��ʾ�ָ���
                    blnFirstTool = True
                    !IsGroup = 1
                End If
            End If
            !������ = strFuncName
            !�˵��� = strFuncName
            .Update
        Next
    End With
End Sub

Private Function SetPlugInBar(ByRef rsBar As ADODB.Recordset, ByVal lngV As Long) As String
'���ܣ����书��ID���Ӳ˵����
'������lngV �汾��1-�ϰ棬2-�°�
'���أ��ַ�������ǰ�Ͱ汾��ʽ�Ĺ��ܴ�
    Dim i As Long
    '���书��ID��ͼ��ID
    With rsBar
        .Filter = 0
        If .EOF Then Exit Function
        .MoveFirst
        For i = 1 To .RecordCount
            !��� = i
            !����ID = conMenu_Tool_PlugIn_Item + i
            !ͼ��ID = conMenu_Tool_PlugIn_Item
            If lngV = 1 Then
                !IsInTool = 0
                !IsGroup = 0
            End If
            .Update
            .MoveNext
        Next
    End With
    Call SetPlugInBarKey(rsBar, 1, lngV)
    Call SetPlugInBarKey(rsBar, 2, lngV)
    Call SetPlugInBarKey(rsBar, 3, lngV)
    rsBar.Filter = 0
End Function

Private Sub SetPlugInBarKey(rsBar As ADODB.Recordset, ByVal intType As Integer, ByVal lngV As Long)
'���ܣ��趨���
'������lngV �汾��1-�ϰ棬2-�°� intType ���ܰ�ť������һ�� 1-�˵�����2-��������3-�����
    Dim i As Long
    With rsBar
        .Filter = "IsInTool=0 and BarType=" & intType
        If .RecordCount = 1 And lngV = 2 Then
            '���ֻ��һ����Ҳ��Ϊ������ť
            !IsInTool = 1
            .Update
        Else
            For i = 1 To .RecordCount
                If i <= 35 Then
                    If i <= 9 Then
                        !�˵��� = !�˵��� & "(&" & i & ")"
                    Else
                        !�˵��� = !�˵��� & "(&" & Chr(55 + i) & ")"
                    End If
                    .Update
                    .MoveNext
                Else
                    Exit For
                End If
            Next
        End If
        
        .Filter = "IsInTool=1 and BarType=" & intType
        For i = 1 To .RecordCount
            If i <= 35 Then
                If i <= 9 Then
                    !�˵��� = !�˵��� & "(&" & i & ")"
                Else
                    !�˵��� = !�˵��� & "(&" & Chr(55 + i) & ")"
                End If
                .Update
                .MoveNext
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitPlugInRsBar(rsBar As ADODB.Recordset)
    Set rsBar = New ADODB.Recordset
    rsBar.Fields.Append "���", adBigInt '��������
    rsBar.Fields.Append "����ID", adBigInt '�˵���ť Control.ID
    rsBar.Fields.Append "ͼ��ID", adBigInt
    rsBar.Fields.Append "������", adVarChar, 1000 'ȥ���ؼ���֮��� ���� ���������ϵİ�ť����
    rsBar.Fields.Append "�˵���", adVarChar, 1000 '�˵���/�Ҽ��˵� ����
    rsBar.Fields.Append "IsAuto", adInteger '�Ƿ��Զ�ִ�й���
    rsBar.Fields.Append "IsGroup", adInteger '�Ƿ�ָ���
    rsBar.Fields.Append "IsInTool", adInteger '�Ƿ������ʾ
    rsBar.Fields.Append "BarType", adInteger '1-�˵�����2����������3��������
    rsBar.CursorLocation = adUseClient
    rsBar.LockType = adLockOptimistic
    rsBar.CursorType = adOpenStatic
    rsBar.Open
End Sub

Public Sub Make��ִ����Ϣ(ByVal strSendDate As String)
'���ܣ�ҽ�����ͺ������ִ����Ϣ����
'������strSendDate ����ʱ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL As Variant
    Dim blnTrans As Boolean
    
    On Error GoTo errH
    strSQL = "select b.����id,b.��ҳid,b.������Դ,c.��ǰ����id as ����ID,c.��Ժ����id as ����ID,a.ִ�в���id" & vbNewLine & _
        "from (select max(a.ҽ��id) as ҽ��ID,a.ִ�в���id from  ����ҽ������ A where a.ִ��״̬ = 0 And Exists" & vbNewLine & _
        "(Select 1 From ��������˵�� X Where x.����id = a.ִ�в���id And x.��������='����') And a.����ʱ�� =[1] group by a.ִ�в���id) a," & vbNewLine & _
        " ����ҽ����¼ B,������ҳ c Where a.ҽ��id = b.Id and b.����id=c.����id and b.��ҳid=c.��ҳid"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", CDate(strSendDate))
    
    If Not rsTmp.EOF Then
        arrSQL = Array()
        For i = 1 To rsTmp.RecordCount
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_ҵ����Ϣ�嵥_insert(" & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!����ID & "," & rsTmp!����ID & "," & rsTmp!������Դ & _
                ",'�д�ִ�е�ҽ����','0010','ZLHIS_CIS_034','" & strSendDate & "',1,0,null,'" & rsTmp!ִ�в���ID & "')"
            rsTmp.MoveNext
        Next
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "mdlCISKernel"
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Get���뵥��ز���()
'���ܣ���ȡ���뵥��ز���ֵ
    Dim str���� As String
    Dim str���� As String
    Dim varTmp As Variant
    str���� = zlDatabase.GetPara(238, glngSys, , "11|11|11|11|1")
    str���� = zlDatabase.GetPara(260, glngSys, , "11")
    varTmp = Split(str����, "|")
    gstrOutUseApp = Mid(varTmp(0), 1, 1) & Mid(varTmp(1), 1, 1) & Mid(varTmp(2), 1, 1) & Mid(varTmp(3), 1, 1)
    gstrInUseApp = Mid(varTmp(0), 2, 1) & Mid(varTmp(1), 2, 1) & Mid(varTmp(2), 2, 1) & Mid(varTmp(3), 2, 1) & Mid(varTmp(4), 1, 1)
    gblnOut���� = Mid(str����, 1, 1) = "1"
    gblnIn���� = Mid(str����, 2, 1) = "1"
End Sub

Public Function GetDiag�������(ByVal str���IDs As String) As String
'���ܣ����ָ����ϵ��������
'���أ�str���=������ϵ���������ַ���
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim str��� As String
    
    strSQL = "Select  A.ID,a.������� From ������ϼ�¼ A " & _
        " Where NVL(A.�������,1) = 1 And a.id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetDiag�������", str���IDs)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            str��� = str��� & "," & rsTmp!�������
            rsTmp.MoveNext
        Loop
        str��� = Mid(str���, 2)
    End If
    GetDiag������� = str���
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDataRISԤԼ(ByVal strIDs As String) As ADODB.Recordset
'���ܣ���ȡָ����Χҽ����RIS�в���ԤԼ��Ϣ������
'������strIDs ��ҽ��ID��
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "select b.ҽ��id as ID,b.ԤԼid from RIS���ԤԼ b where b.ҽ��id in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)"
    Set GetDataRISԤԼ = zlDatabase.OpenSQLRecord(strSQL, "GetDataRISԤԼ", strIDs)
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckPathInItem(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str������ĿIDs As String, ByRef str���� As String, ByVal lng�׶�Id As Long, ByVal bln��ҩ�䷽ As Boolean, ByVal byt��Ч As Byte, Optional ByVal bln��ҩ As Boolean) As Long
'���ܣ�����ٴ�·�����ˣ���ǰ�����ҽ����һ��������Ŀ���Ƿ��ǵ�ǰ�׶ε�·������Ŀ������ǣ��򷵻���ĿID
'      �����ҽ�ִ��һ�ε���Ŀ������ʱ�ض��Ѳ���������Ӿ͵���·������Ŀ��
'������str������ĿIDs= 'ҩƷ������ҩ;�����÷����巨����Ѫ����;��,���鲻���ɼ���ʽ��������������������������鲻����λ����
'      lng�׶�ID=���ڼ��ϲ�·������Ŀ�Ƿ�ƥ��
'      bln��ҩ�䷽=��ҩ�䷽�����������ݲ������õ������޸ĵ���ҩ��������
'      byt��Ч =ҽ����Ч
'���أ�·����ĿID�ͷ�������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str֤��IDs As String
    Dim lng�Ķ���ҩζ�� As Long, dbl��ҩζ�� As Double, lng��ҩζ�� As Long
    Dim i As Long
    Dim arrTmp As Variant, blnTmp As Boolean
    Dim blnƥ����Ч As Boolean
    Dim bln��ƥ�� As Boolean
        Dim blnҩƷ������ͬ����·���� As Boolean
    Dim strҩƷ����ids As String
    
    str���� = ""
    If str������ĿIDs = "0" Then
    '����¼���ҽ���̶�����·������Ŀ
        CheckPathInItem = 0
    Else
        'Wm_Concat�����ڷ����е�����������,����10.2.0.5�з���ֵ���ͱ仯��
'        strSQL = "Select ����,·����Ŀid" & vbNewLine & _
'                "From (Select ����,·����Ŀid, ��id, Wmsys.Wm_Concat(������Ŀid) ������Ŀids" & vbNewLine & _
'                "       From (Select Rownum, c.·����Ŀid, d.������Ŀid, b.����, Decode(����Ҫ��,1,Nvl(d.���id,d.id),0) ��id" & vbNewLine & _
'                "              From �����ٴ�·�� A, �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D" & vbNewLine & _
'                "              Where a.����id = [1] And a.��ҳid = [2] And b.�׶�id = a.��ǰ�׶�id And b.Id = c.·����Ŀid And c.ҽ������id = d.Id And b.ִ�з�ʽ<>4" & vbNewLine & _
'                "              Order By b.����, b.��Ŀ���, d.���)" & vbNewLine & _
'                "       Group By ����,·����Ŀid,��id)" & vbNewLine & _
'                "Where ������Ŀids = [3]"
        blnƥ����Ч = CBool(zlDatabase.GetPara("ƥ��ʱ��Ч��ͬ��·������Ŀ", glngSys, p�ٴ�·��Ӧ��, "0"))
        lng��ҩζ�� = Val(zlDatabase.GetPara("��ҩ�䷽�����޸ĵ���ҩζ������", glngSys, p�ٴ�·��Ӧ��, "30"))
        bln��ƥ�� = Val(zlDatabase.GetPara("ҩƷҽ����ƥ��Ϊ·������Ŀ", glngSys, p�ٴ�·��Ӧ��, "0")) = 1
                blnҩƷ������ͬ����·���� = Val(zlDatabase.GetPara("ҩƷҽ����ͬ���಻��·����ҽ��", glngSys, p�ٴ�·��Ӧ��, "0")) = 1
        If blnҩƷ������ͬ����·���� And bln��ҩ Then
            strSQL = "Select f_List2str(Cast(Collect(To_Char(����id)) As t_Strlist)) As ҩƷ����ids" & vbNewLine & _
                    "From ������ĿĿ¼" & vbNewLine & _
                    "Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPathInItem", str������ĿIDs)
            If rsTmp.RecordCount > 0 Then
                strҩƷ����ids = rsTmp!ҩƷ����IDs & ""
            End If
        End If
        '��ҩ;��������Ϊ��ȡ��ͬ���õ�ԭ��ʵ��ʹ��ʱ���Ƕ���ĸ�ҩ;��������ֻ�ж�ҩƷ��ͬ����
        If Not bln��ҩ�䷽ Then
            strSQL = "Select ����, ·����Ŀid,��Ч,������Ŀids,ҩƷ����IDs" & vbNewLine & _
                    "From (Select ����, ·����Ŀid, ��id,��Ч,f_List2str(Cast(Collect(To_Char(������Ŀid)) As t_Strlist)) As ������Ŀids,f_List2str(Cast(Collect(To_Char(ҩƷ����ID)) As t_Strlist)) As ҩƷ����IDs" & vbNewLine & _
                    "       From (Select c.·����Ŀid, b.����, d.������Ŀid, Nvl(d.���id, d.Id) ��id,d.��Ч,d.���,e.����ID as ҩƷ����ID" & vbNewLine & _
                    "              From �����ٴ�·�� A, �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D,������ĿĿ¼ E" & vbNewLine & _
                    "              Where a.����id = [1] And a.��ҳid = [2] And b.�׶�id = [4] And b.Id = c.·����Ŀid And c.ҽ������id = d.Id And D.������ĿID = E.ID And " & vbNewLine & _
                    "                    (b.ִ�з�ʽ <> 4 or b.ִ�з�ʽ = 4 And Not Exists(Select 1 From ����·��ִ�� E Where e.·����¼ID=a.id And a.��ǰ�׶�id=e.�׶�id and b.id=e.��Ŀid))" & vbNewLine & _
                    "                    And Not Exists(Select 1 From ������ĿĿ¼ E Where D.������ĿID = E.ID And E.��� = 'E' And  E.�������� In('2','3','4','6'))" & vbNewLine & _
                    "                    And Not Exists(Select 1 From ������ĿĿ¼ E Where D.������ĿID = E.ID And E.��� In('G','F','D') And D.���ID<>0 )" & vbNewLine & _
                    "              Order By b.����, b.��Ŀ���, d.���)" & vbNewLine & _
                    "       Group By ����, ·����Ŀid, ��id,��Ч)" & vbNewLine & _
                    IIF(blnҩƷ������ͬ����·���� And bln��ҩ, IIF(InStr(strҩƷ����ids, ",") > 0, " Where instr(ҩƷ����IDs,',')>0 ", "Where (ҩƷ����IDs = [6] or instr(','||ҩƷ����IDs||',',','||[6]||',')>0)"), _
                        IIF(InStr(str������ĿIDs, ",") > 0, " Where instr(������Ŀids,',')>0 ", "Where (������Ŀids = [3] or instr(','||������Ŀids||',',','||[3]||',')>0)")) & IIF(blnƥ����Ч, " And ��Ч =[5]", "")
        
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPathInItem", lng����ID, lng��ҳID, str������ĿIDs, lng�׶�Id, byt��Ч, strҩƷ����ids)
            
            If blnҩƷ������ͬ����·���� And bln��ҩ Then
                If InStr(strҩƷ����ids, ",") > 0 Then
                    '�����Ŀ�ж�ʱ���������Ŀ������˳�����������һ����·�������ôһ�����·�����
                    arrTmp = Split(strҩƷ����ids, ",")
                    Do While Not rsTmp.EOF
                        blnTmp = True
                        For i = 0 To UBound(arrTmp)
                            If InStr("," & rsTmp!ҩƷ����IDs & ",", "," & arrTmp(i) & ",") = 0 Then
                                blnTmp = False
                                Exit For
                            End If
                        Next
                        If blnTmp Then
                            CheckPathInItem = rsTmp!·����ĿID
                            str���� = rsTmp!����
                            GoTo FuncEnd
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                Else
                    '����ж��·����Ŀ����ֻȡ��һ��
                    If rsTmp.RecordCount > 0 Then
                        CheckPathInItem = rsTmp!·����ĿID
                        str���� = rsTmp!����
                        If Not blnƥ����Ч Then
                            '���ڶ������¸�����Ч����ƥ��,����ƥ��������ĿID����Ч����ͬ�ģ�Ϊ�˱�֤��ͬһ�׶�,ͬһҩƷ,��ͬ��Ŀ,��Ч��һ�µ������,������ƥ�䣩
                            For i = 1 To rsTmp.RecordCount
                                If rsTmp!��Ч & "" = byt��Ч & "" Then
                                    CheckPathInItem = rsTmp!·����ĿID
                                    str���� = rsTmp!����
                                    Exit For
                                End If
                                rsTmp.MoveNext
                            Next
                        End If
                    End If
                End If
            Else
                If InStr(str������ĿIDs, ",") > 0 Then
                    '�����Ŀ�ж�ʱ���������Ŀ������˳�����������һ����·�������ôһ�����·�����
                    arrTmp = Split(str������ĿIDs, ",")
                    Do While Not rsTmp.EOF
                        blnTmp = True
                        For i = 0 To UBound(arrTmp)
                            If InStr("," & rsTmp!������Ŀids & ",", "," & arrTmp(i) & ",") = 0 Then
                                blnTmp = False
                                Exit For
                            End If
                        Next
                        If blnTmp Then
                            CheckPathInItem = rsTmp!·����ĿID
                            str���� = rsTmp!����
                            GoTo FuncEnd
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                Else
                    '����ж��·����Ŀ����ֻȡ��һ��
                    If rsTmp.RecordCount > 0 Then
                        CheckPathInItem = rsTmp!·����ĿID
                        str���� = rsTmp!����
                        If Not blnƥ����Ч Then
                            '���ڶ������¸�����Ч����ƥ��,����ƥ��������ĿID����Ч����ͬ�ģ�Ϊ�˱�֤��ͬһ�׶�,ͬһҩƷ,��ͬ��Ŀ,��Ч��һ�µ������,������ƥ�䣩
                            For i = 1 To rsTmp.RecordCount
                                If rsTmp!��Ч & "" = byt��Ч & "" Then
                                    CheckPathInItem = rsTmp!·����ĿID
                                    str���� = rsTmp!����
                                    Exit For
                                End If
                                rsTmp.MoveNext
                            Next
                        End If
                    End If
                End If
            End If
        Else
            'ƥ��ʱ�ſ������ϵ�֤��
            str֤��IDs = Get֤��IDs(lng����ID, lng��ҳID)
            strSQL = "Select  ����, ·����Ŀid, ��id, f_List2str(Cast(Collect(To_Char(������Ŀid)) As t_Strlist)) As ������Ŀids" & vbNewLine & _
                    "From (Select c.·����Ŀid, b.����, d.������Ŀid, Nvl(d.���id, d.Id) ��id, d.���" & vbNewLine & _
                    "       From �����ٴ�·�� A, �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D" & vbNewLine & _
                    "       Where a.����id = [1] And a.��ҳid = [2] And b.�׶�id = [3] And b.Id = c.·����Ŀid And c.ҽ������id = d.Id And" & vbNewLine & _
                    "             (b.ִ�з�ʽ <> 4 Or b.ִ�з�ʽ = 4 And Not Exists" & vbNewLine & _
                    "              (Select 1 From ����·��ִ�� E Where e.·����¼id = a.Id And a.��ǰ�׶�id = e.�׶�id And b.Id = e.��Ŀid)) And Exists" & vbNewLine & _
                    "        (Select 1 From ������ĿĿ¼ E Where d.������Ŀid = e.Id And e.��� = '7') " & vbNewLine & _
                    IIF(str֤��IDs <> "", " And (Instr(',' || [4] || ',', ',' || d.�����Ŀid || ',') > 0 Or d.�����Ŀid Is Null)", "") & vbNewLine & _
                    "       Order By b.����, b.��Ŀ���, d.���)" & vbNewLine & _
                    "Group By ����, ·����Ŀid, ��id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPathInItem", lng����ID, lng��ҳID, lng�׶�Id, str֤��IDs)
            Do While Not rsTmp.EOF
                If rsTmp!������Ŀids & "" <> "" Then
                    '����Ķ�����ҩ
                    dbl��ҩζ�� = (UBound(Split(rsTmp!������Ŀids & "", ",")) + 1) * lng��ҩζ�� / 100
                    lng�Ķ���ҩζ�� = 0
                    '���ң��䷽�����ҩ
                    For i = 0 To UBound(Split(str������ĿIDs, ","))
                        If InStr("," & rsTmp!������Ŀids & ",", "," & Split(str������ĿIDs, ",")(i) & ",") = 0 Then
                            lng�Ķ���ҩζ�� = lng�Ķ���ҩζ�� + 1
                        End If
                    Next
                    '�����䷽�е���ҩ������ȱ�ٵ�
                    If rsTmp!������Ŀids & "" <> "" Then
                        For i = 0 To UBound(Split(rsTmp!������Ŀids & "", ","))
                            If InStr("," & str������ĿIDs & ",", "," & Split(rsTmp!������Ŀids & "", ",")(i) & ",") = 0 Then
                                lng�Ķ���ҩζ�� = lng�Ķ���ҩζ�� + 1
                            End If
                        Next
                    End If
                    '���������ķ�Χ֮�ڣ���ƥ��ɹ����������ƥ��
                    If lng�Ķ���ҩζ�� <= dbl��ҩζ�� Then
                        CheckPathInItem = rsTmp!·����ĿID
                        str���� = rsTmp!����
                        Exit Do
                    End If
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
FuncEnd:
    If CheckPathInItem = 0 And bln��ƥ�� = True Then
        
        strSQL = "Select ����, ·����Ŀid,��Ч,������Ŀids,ID" & vbNewLine & _
                    "From (Select ����, ·����Ŀid, ��id,��Ч,f_List2str(Cast(Collect(To_Char(������Ŀid)) As t_Strlist)) As ������Ŀids,ID" & vbNewLine & _
                    "       From (Select c.·����Ŀid, b.����, d.������Ŀid, Nvl(d.���id, d.Id) ��id,d.��Ч,d.���,A.ID" & vbNewLine & _
                    "              From �����ٴ�·�� A, �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D" & vbNewLine & _
                    "              Where a.����id = [1] And a.��ҳid = [2] And b.�׶�id = [4] And b.Id = c.·����Ŀid And c.ҽ������id = d.Id And" & vbNewLine & _
                    "                    (b.ִ�з�ʽ <> 4 or b.ִ�з�ʽ = 4 And Not Exists(Select 1 From ����·��ִ�� E Where e.·����¼ID=a.id And a.��ǰ�׶�id=e.�׶�id and b.id=e.��Ŀid))" & vbNewLine & _
                    "                    And Exists(Select 1 From ������ĿĿ¼ E Where D.������ĿID = E.ID And E.��� In('5','6') And D.���ID<>0 )" & vbNewLine & _
                    "              Order By b.����, b.��Ŀ���, d.���)" & vbNewLine & _
                    "       Group By ����, ·����Ŀid, ��id,��Ч,ID) A Where exists (select 1 from ����·��ִ�� E where E.·����¼ID=A.ID And E.�׶�ID=[4] and E.��ĿID=A.·����ĿID)" & _
                    " and exists(select 1 from ������ĿĿ¼ where ��� in('5','6') and ID In (select /*+cardinality(A,10)*/ column_value from Table(f_Str2list([3]) ) A))"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPathInItem", lng����ID, lng��ҳID, str������ĿIDs, lng�׶�Id, byt��Ч)
        Do While Not rsTmp.EOF
            CheckPathInItem = rsTmp!·����ĿID
            str���� = rsTmp!����
            
            Exit Function
        Loop
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get֤��IDs(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���ȡ�ò���֤��IDs�����ŷָ�
    Dim strSQL As String, rsTmp As Recordset
    Dim str֤��IDs As String
    
    strSQL = "Select ֤��ID From ������ϼ�¼ Where ����id = [1] And ��ҳid = [2] And NVL(�������,1) = 1 And ֤��id Is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get֤��IDs", lng����ID, lng��ҳID)
    Do While Not rsTmp.EOF
        str֤��IDs = str֤��IDs & "," & rsTmp!֤��id
        rsTmp.MoveNext
    Loop
    Get֤��IDs = Mid(str֤��IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiPathInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByRef str·����Ŀ���� As String = "-1") As ADODB.Recordset
'���ܣ���ȡ·�����˵�ǰ·����Ϣ
'���أ�str����=��ǰ�������һ��·����Ŀ�����ķ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsRet As ADODB.Recordset
    Dim blnDo As Boolean
    
    blnDo = str·����Ŀ���� <> "-1"
    str·����Ŀ���� = ""
    strSQL = "Select a.·����¼id, a.��ǰ�׶�id, a.��ǰ����, a.��ʼ����, b.����, b.����" & vbNewLine & _
            "From (Select a.Id As ·����¼id, a.��ǰ�׶�id, a.��ǰ����, Max(b.Id) ִ��id, Min(c.����) As ��ʼ����" & vbNewLine & _
            "       From �����ٴ�·�� A, ����·��ִ�� B, ����·��ִ�� C" & vbNewLine & _
            "       Where a.Id = b.·����¼id And a.Id = c.·����¼id And b.�׶�id + 0 = a.��ǰ�׶�id And b.���� = a.��ǰ���� And a.״̬ = 1 And" & vbNewLine & _
            "             a.����id = [1] And a.��ҳid = [2]" & vbNewLine & _
            "       Group By a.Id, a.��ǰ�׶�id, a.��ǰ����) A, ����·��ִ�� B" & vbNewLine & _
            "Where a.ִ��id = b.Id"

    On Error GoTo errH
    Set rsRet = zlDatabase.OpenSQLRecord(strSQL, "���˵�ǰ·����Ϣ", lng����ID, lng��ҳID)
    If rsRet.RecordCount > 0 And blnDo Then
        str·����Ŀ���� = "" & rsRet!����
        
        '�������������ҽ������Ŀ����ȡҽ������Ŀ�ķ���
        strSQL = "Select ����" & vbNewLine & _
                "From ����·��ִ��" & vbNewLine & _
                "Where ID = (Select Max(ID)" & vbNewLine & _
                "            From ����·��ִ�� A" & vbNewLine & _
                "            Where a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3] And Exists (Select 1 From ����·��ҽ�� B Where a.Id = b.·��ִ��id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˵�ǰ·����Ϣ", Val(rsRet!·����¼id), Val(rsRet!��ǰ�׶�ID), CDate(rsRet!����))
        If rsTmp.RecordCount > 0 Then
            str·����Ŀ���� = "" & rsTmp!����
        End If
    End If
    Set GetPatiPathInfo = rsRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiPathAppend(ByVal lng��¼ID As Long, ByVal dat���� As Date) As ADODB.Recordset
'����:�������ڻ�ȡ��Ч�׶κ�����
    Dim strSQL      As String
    
    strSQL = "Select �׶�id,����" & vbNewLine & _
            "From (Select a.�׶�id, a.����, a.�Ǽ�ʱ��" & vbNewLine & _
            "       From ����·��ִ�� A" & vbNewLine & _
            "       Where a.·����¼id = [1] And a.���� = [2]" & vbNewLine & _
            "       Order By a.�Ǽ�ʱ�� Desc)" & vbNewLine & _
            "Where Rownum < 2"
        
    On Error GoTo errH
    Set GetPatiPathAppend = zlDatabase.OpenSQLRecord(strSQL, "GetPatiPathAppend", lng��¼ID, dat����)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPathOutItemID(ByVal lng·����¼Id As Long, ByVal DatAddDate As Date, Optional ByVal bytFunc As Byte) As Long
'���ܣ���ȡ�ղ���ӵ�·������Ŀ��ִ��ID
'����:bytFunc=0 סԺ�ٴ�·�� 1-�����ٴ�·��
'���أ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    If bytFunc = 1 Then
        strSQL = "Select Max(b.Id) ִ��id" & vbNewLine & _
                    "       From ��������·��ִ�� B" & vbNewLine & _
                    "       Where b.·����¼id = [1] And b.�Ǽ�ʱ�� = [2] And Nvl(b.��ĿID,0) = 0"
    Else
        strSQL = "Select Max(b.Id) ִ��id" & vbNewLine & _
                "       From ����·��ִ�� B" & vbNewLine & _
                "       Where b.·����¼id = [1] And b.�Ǽ�ʱ�� = [2] And Nvl(b.��ĿID,0) = 0"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPathOutItemID", lng·����¼Id, DatAddDate)
    If rsTmp.RecordCount > 0 Then
        GetPathOutItemID = Val("" & rsTmp!ִ��Id)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MakePriceRecord���뵥(ByVal strType As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strAdvice As String, ByVal str��Ŀ���� As String, ByVal str�ѱ� As String, ByVal lng��������ID As Long, ByRef rsPrice As ADODB.Recordset) As Boolean
'���ܣ�����վ�����浥��ʹ�����뵥�������뵥ʱ�����ɶ�Ӧ������ҽ���ķ�����ϸ��¼�����ж��߼�ͬ����/סԺҽ���༭����MakePriceRecord����
'������strAdvice ҽ����ʱ��SQL��str��Ŀ���ң�������ĿID:ִ�п���ID,...��lng��ҳID ��������ﲡ������� �Һ�ID
'      strType ���÷� ��λ���ֱ�ʾ��һλ��ʾ���뵥 1�����飬2����飬3����Ѫ��4��������5������ڶ�λ��ʾ�������סԺ 1�����2��סԺ
'���أ��мƼ����ݼ�¼�����ݲŷ���True
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str�����շ� As String
    Dim lngִ�п���ID As Long, blnLoad As Boolean
    Dim dbl���� As Double, dbl��� As Double, dblʵ�� As Double
    Dim str��Ŀ As String, blnDo As Boolean
    Dim lng���ID As Long, lng����ID As Long
    Dim int���� As String
    
    int���� = Mid(strType, 2, 1)
    
    On Error GoTo errH
    
    str�����շ� = "Select c.������Ŀid, c.�շ���Ŀid, c.��鲿λ, c.��鷽��, c.��������, c.�շ�����, c.���ж���, c.������Ŀ, c.�շѷ�ʽ, c.���ÿ���id,c.top From (" & _
            "Select /*+cardinality(D,10)*/ Distinct C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,C.���ÿ���id" & _
            " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
            " From �����շѹ�ϵ C,Table(f_Num2list2([1])) D Where C.������ĿID=D.c1" & _
            "      And (C.���ÿ���ID is Null or C.���ÿ���ID = D.c2 And C.������Դ = 1)" & _
            " ) c Where Nvl(c.���ÿ���id, 0) = c.Top"
                
    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
        " Select A.���,A.�������,C.��� as �շ����,B.�շ���ĿID as �շ�ϸĿID,D.������ĿID," & _
        " Decode(A.����,0,1,A.����)*B.�շ����� as ����,Decode(C.�Ƿ���,1,D.ȱʡ�۸�,D.�ּ�) as ����," & _
        " C.�Ƿ���,C.���ηѱ�,A.ִ�п���ID, a.��������,d.�����շ���" & _
        " From (" & strAdvice & ") A,(" & str�����շ� & ") B,�շ���ĿĿ¼ C,�շѼ�Ŀ D,������ĿĿ¼ E,��Ѫ������ F" & _
        " Where a.������Ŀid = b.������Ŀid And (a.���id Is Null And a.ִ�б�� In (1, 2) And b.�������� = 1 Or" & _
        "            a.�걾��λ = b.��鲿λ And a.��鷽�� = b.��鷽�� And Nvl(b.��������, 0) = 0 Or" & _
        "            a.��鷽�� Is Null And Nvl(b.��������, 0) = 0 And b.��鲿λ Is Null And b.��鷽�� Is Null) And a.������Ŀid = e.Id And" & vbNewLine & _
        "            e.�Թܱ��� = f.����(+) And (Nvl(b.�շѷ�ʽ, 0) = 1 And c.��� = '4' And b.�շ���Ŀid = f.����id Or" & vbNewLine & _
        "            Not (Nvl(b.�շѷ�ʽ, 0) = 1 And c.��� = '4' And f.����id Is Not Null)) And Nvl(a.�Ƽ�����, 0) = 0 And" & vbNewLine & _
        "            Nvl(a.ִ������, 0) Not In (0, 5) And b.�շ���Ŀid = c.Id And b.�շ���Ŀid = d.�շ�ϸĿid And" & vbNewLine & _
        "            ((Sysdate Between d.ִ������ And d.��ֹ����) Or (Sysdate >= d.ִ������ And d.��ֹ���� Is Null)) And" & vbNewLine & _
        "            (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.������� In (1, 3) And" & vbNewLine & _
        "            (c.վ�� = '-' Or c.վ�� Is Null)"

    strSQL = "Select a.���,a.�������,a.�շ����,a.�շ�ϸĿid,a.������Ŀid,a.����,a.����,a.�Ƿ���,a.���ηѱ�,a.ִ�п���id,a.��������,a.�����շ��� From (" & strSQL & ") A Order by a.���,a.������ĿID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "MakePriceRecord���뵥", str��Ŀ����)
    If Not rsTmp.EOF Then
        '��ʼ����¼��
        Set rsPrice = New ADODB.Recordset
        With rsPrice
            .Fields.Append "����ID", adBigInt
            .Fields.Append "��ҳID", adBigInt, , adFldIsNullable
            .Fields.Append "�շ����", adVarChar, 1
            .Fields.Append "�շ�ϸĿID", adBigInt
            .Fields.Append "����", adDouble
            .Fields.Append "����", adDouble
            .Fields.Append "ʵ�ս��", adDouble
            .Fields.Append "������", adVarChar, 100, adFldIsNullable
            .Fields.Append "��������", adVarChar, 100, adFldIsNullable
            
            If int���� = 1 Then
                .Fields.Append "����ID", adBigInt, , adFldIsNullable
                .Fields.Append "���ID", adBigInt, , adFldIsNullable
            End If
            
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .CursorType = adOpenStatic
            .Open
        End With
        '���������ϸ
        dblʵ�� = 0: blnDo = True
        Do While Not rsTmp.EOF
            'ִ�п���
            If blnDo Then
                lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
            End If
            
            '����
            dbl���� = Format(NVL(rsTmp!����, 0), gstrDecPrice) '�������ȡ��ȱʡ�۸�
   
            '���
            dbl��� = CCur(NVL(rsTmp!����, 0) * dbl����)
            If NVL(rsTmp!��������, 0) = 1 Then
                dbl��� = dbl��� * NVL(rsTmp!�����շ���, 100) / 100
            End If
            dbl��� = Format(dbl���, gstrDec)
            
            If NVL(rsTmp!���ηѱ�, 0) = 0 And str�ѱ� <> "" Then
                dbl��� = ActualMoney(str�ѱ�, rsTmp!������ĿID, dbl���, rsTmp!�շ�ϸĿID, lngִ�п���ID, NVL(rsTmp!����, 0))
            End If
            
            dblʵ�� = dblʵ�� + dbl���
            
            '��Ŀ�仯ʱ����
            str��Ŀ = rsTmp!��� & "," & rsTmp!�շ�ϸĿID
            blnDo = False: rsTmp.MoveNext
            If Not rsTmp.EOF Then
                If rsTmp!��� & "," & rsTmp!�շ�ϸĿID <> str��Ŀ Then blnDo = True
            Else
                blnDo = True
            End If
            rsTmp.MovePrevious
            
            If blnDo Then
                rsPrice.AddNew
                rsPrice!����ID = lng����ID
                rsPrice!��ҳID = lng��ҳID
                rsPrice!�շ���� = rsTmp!�շ����
                rsPrice!�շ�ϸĿID = rsTmp!�շ�ϸĿID
                rsPrice!���� = NVL(rsTmp!����, 0)
                rsPrice!���� = dbl����
                rsPrice!ʵ�ս�� = dbl���
                rsPrice!������ = UserInfo.����
                rsPrice!�������� = CStr(Sys.RowValue("���ű�", lng��������ID, "����"))
'                If lng����ID <> 0 Then mrsPrice!����id = lng����ID
'                If lng���ID <> 0 Then mrsPrice!���id = lng���ID
                rsPrice.Update
                dblʵ�� = 0
            End If
            
            rsTmp.MoveNext
        Loop
        
        rsPrice.MoveFirst
        MakePriceRecord���뵥 = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub FuncClinicPay(frmMain As Object, ByVal lng����ID As Long, ByVal strNO As String)
'���ܣ����֧��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str����ҽ��IDs As String
    Dim blnʹ��Ԥ�� As Boolean
    
    On Error GoTo errH
    
    If gobjSquareCard Is Nothing Then
        On Error Resume Next
        Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        If Not gobjSquareCard Is Nothing Then
            If gobjSquareCard.zlInitComponents(frmMain, p����ҽ���´�, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set gobjSquareCard = Nothing
            End If
        End If
    End If
    
    If Not gobjSquareCard Is Nothing Then
        strSQL = "Select f_List2str(Cast(Collect(a.ҽ����� || '') As t_Strlist)) As ҽ��ids" & vbNewLine & _
            "From ����ҽ����¼ B, ������ü�¼ A" & vbNewLine & _
            "Where a.����id =[1] And a.��¼����=1 And a.��¼״̬ = 0 And a.ҽ�����=b.Id And b.ҽ��״̬=8 And b.�Һŵ� =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, strNO)
        If Not IsNull(rsTmp!ҽ��ids) Then
            blnʹ��Ԥ�� = Val(zlDatabase.GetPara("���֧������ʹ��Ԥ����", glngSys, p����ҽ���´�, "1")) = 1
            str����ҽ��IDs = rsTmp!ҽ��ids & ""
            Call gobjSquareCard.zlSquareAffirm(frmMain, p����ҽ���´�, GetInsidePrivs(p����ҽ���´�), lng����ID, 0, False, 1, , str����ҽ��IDs, , , blnʹ��Ԥ��)
        Else
            MsgBox "��ǰ����û�пɽ����ҽ�����á�", vbInformation, gstrSysName
        End If
    Else
        MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function InitObjPublicDrug() As Boolean
    If gobjPublicDrug Is Nothing Then
        On Error Resume Next
        Set gobjPublicDrug = CreateObject("zlPublicDrug.clsPublicDrug")
        If Not gobjPublicDrug Is Nothing Then
            Call gobjPublicDrug.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
        End If
        err.Clear: On Error GoTo 0
    End If
    InitObjPublicDrug = Not gobjPublicDrug Is Nothing
End Function

Public Function FuncLisRptFileView(frmMain As Object, ByVal lngҽ��ID As Long) As Boolean
'���ܣ���LIS�����ļ��鿴
'���أ��ɹ��򿪣�true
    Dim strFile As String
    Dim lngRetu As Long, strInfo As String
    Dim objFile As New FileSystemObject
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng����ID As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    strSQL = "select a.id,a.����,a.������ from ҽ���������� a,����ҽ������ b where a.id=b.����id and b.ҽ��id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
    
    If rsTmp.EOF Then
        MsgBox "��ҽ��û�в��������ļ���", vbInformation, gstrSysName:
        Screen.MousePointer = 0
        Exit Function
    End If
    
    If rsTmp.RecordCount = 1 Then
        lng����ID = Val(rsTmp!ID & "") '���ֻ��һ����ֱ�Ӵ�
    Else
        strSQL = "select a.id,b.����״̬ as ��¼״̬,null as ������,c.����,c.����,c.�Ա�,a.������ as ����,d.���� as ִ�п���," & vbNewLine & _
            " a.����ʱ�� as ��¼ʱ��,a.����,a.������" & vbNewLine & _
            " from ҽ���������� a,����ҽ������ b,����ҽ����¼ c,���ű� d" & vbNewLine & _
            " where a.id=b.����id and b.ҽ��id=c.id and c.ִ�п���id=d.id and c.id=[1]"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
        Screen.MousePointer = 0
        lng����ID = frmCardSel.ShowMe(rsTmp, frmMain)
        rsTmp.Filter = "ID=" & lng����ID
        Screen.MousePointer = 11
    End If
    
    If lng����ID <> 0 Then
        strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & rsTmp!������ & "." & IIF(Val(rsTmp!���� & "") = 0, "pdf", "html")
        If objFile.FileExists(strFile) Then objFile.DeleteFile strFile, True
        
        strFile = Sys.ReadLob(glngSys, 22, lng����ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Function
        End If
        
        lngRetu = ShellExecute(frmMain.hwnd, "open", strFile, "", "", SW_SHOWNORMAL)
        If lngRetu <= 32 Then
            Select Case lngRetu
            Case 2: strInfo = "����Ĺ���"
            Case 29: strInfo = "����ʧ��"
            Case 30: strInfo = "����Ӧ�ó�ʽæµ��..."
            Case 31: strInfo = "û�й����κ�Ӧ�ó�ʽ"
            Case Else: strInfo = "�޷�ʶ��Ĵ���"
            End Select
            MsgBox "�ļ���ʱ����" & vbCrLf & vbCrLf & vbTab & strInfo, vbExclamation, gstrSysName
        Else
            '�ɹ��򿪺���Ϊ����
            strSQL = "Zl_������ļ�¼_Insert(" & lngҽ��ID & ",null," & lng����ID & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISKernel")
            FuncLisRptFileView = True
        End If
    End If
    
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'���ܣ�֧�ֹ��ֵĹ���
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '���¹�
            zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '���Ϲ�
            zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
End Function




Public Function GetTsPrivs(ByVal lngMdl As Long) As String
    '��ȡ����ҽ��Ȩ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim i As Integer
    On Error GoTo errH

    
    If lngMdl = p����ҽ���´� Then
        If gstrTsPrivsMZ <> "" Then
            GetTsPrivs = gstrTsPrivsMZ
            Exit Function
        Else
            strSQL = "Select ��������ҽ��Ȩ�� as Ȩ�� From ��Ա�� Where ID=[1]"
        End If
    Else
        If gstrTsPrivsZY <> "" Then
            GetTsPrivs = gstrTsPrivsZY
            Exit Function
        Else
            strSQL = "Select סԺ����ҽ��Ȩ�� as Ȩ�� From ��Ա�� Where ID=[1]"
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get��������", UserInfo.ID)
    If rsTmp.RecordCount <> 0 Then
        strTmp = IIF(rsTmp!Ȩ�� & "" = "", "0000", rsTmp!Ȩ�� & "")
    End If
    
    '��֯Ȩ���ַ���
    If strTmp = "0000" Then
        GetTsPrivs = ";"
    Else
        GetTsPrivs = IIF(Mid(strTmp, 1, 1) = 1, ";�´ﶾ��ҩ��", "")
        GetTsPrivs = GetTsPrivs & IIF(Mid(strTmp, 2, 1) = 1, ";�´�����ҩ��", "")
        GetTsPrivs = GetTsPrivs & IIF(Mid(strTmp, 3, 1) = 1, ";�´ﾫ��ҩ��", "")
        GetTsPrivs = GetTsPrivs & IIF(Mid(strTmp, 4, 1) = 1, ";�´����ҩ��", "")
        GetTsPrivs = GetTsPrivs & ";"
    End If
    
    '��������סԺȨ������
    If lngMdl = p����ҽ���´� Then
        gstrTsPrivsMZ = GetTsPrivs
    Else
        gstrTsPrivsZY = GetTsPrivs
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�Զ������뵥(ByVal int���� As Integer, ByRef str�Զ������뵥IDs As String) As Boolean
'���ܣ���ȡ���úõ��Զ������뵥
    Dim strSQL As String, rsTmp As Recordset
    Dim strReturn As String
    
    If str�Զ������뵥IDs <> "" Then Exit Function
    
    strSQL = "Select a.Id, a.����" & vbNewLine & _
                "From �����ļ��б� A, �Զ������뵥�ļ� B" & vbNewLine & _
                "Where a.Id = b.�ļ�id And a.��ʽ = 1 and exists(select 1 from ��������Ӧ�� C Where A.Id=C.�����ļ�ID And C.Ӧ�ó���=[1])" & vbNewLine & _
                "Group By a.Id, a.����" & vbNewLine & _
                "Having Count(1) = 3" '3���ļ����е�ʱ�����ʾ����������޷�ʹ��
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get�Զ������뵥", 2)
    Do While Not rsTmp.EOF
        strReturn = strReturn & "|" & rsTmp!ID & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    str�Զ������뵥IDs = Mid(strReturn, 2)
    Get�Զ������뵥 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiOutPathInfo(ByVal lng����ID As Long, ByVal lng����id As Long, Optional ByRef str·����Ŀ���� As String = "-1") As ADODB.Recordset
'���ܣ���ȡ·�����˵�ǰ·����Ϣ
'���أ�str����=��ǰ�������һ��·����Ŀ�����ķ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsRet As ADODB.Recordset
    Dim blnDo As Boolean
    
    blnDo = str·����Ŀ���� <> "-1"
    str·����Ŀ���� = ""
    strSQL = "Select a.·����¼id, a.��ǰ�׶�id, a.��ǰ����, b.����, b.����" & vbNewLine & _
            "From (Select a.Id As ·����¼id, a.��ǰ�׶�id, a.��ǰ����, Max(b.Id) ִ��id" & vbNewLine & _
            "       From ��������·�� A, ��������·��ִ�� B, ��������·����¼ C" & vbNewLine & _
            "       Where C.·����¼ID=A.ID And a.Id = b.·����¼id And b.�׶�id = a.��ǰ�׶�id And b.���� = a.��ǰ���� And a.״̬ =1 And A.����ID =[1] and A.����ID=[2]" & vbNewLine & _
            "       Group By a.Id, a.��ǰ�׶�id, a.��ǰ����) A, ��������·��ִ�� B" & vbNewLine & _
            "Where a.ִ��id = b.Id"
    On Error GoTo errH
    Set rsRet = zlDatabase.OpenSQLRecord(strSQL, "���˵�ǰ·����Ϣ", lng����ID, lng����id)
    If rsRet.RecordCount > 0 And blnDo Then
        str·����Ŀ���� = "" & rsRet!����
        
        '�������������ҽ������Ŀ����ȡҽ������Ŀ�ķ���
        strSQL = "Select ����" & vbNewLine & _
                "From ��������·��ִ��" & vbNewLine & _
                "Where ID = (Select Max(ID)" & vbNewLine & _
                "            From ��������·��ִ�� A" & vbNewLine & _
                "            Where a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3] And Exists (Select 1 From ��������·��ҽ�� B Where a.Id = b.·��ִ��id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˵�ǰ·����Ϣ", Val(rsRet!·����¼id), Val(rsRet!��ǰ�׶�ID), CDate(rsRet!����))
        If rsTmp.RecordCount > 0 Then
            str·����Ŀ���� = "" & rsTmp!����
        End If
    End If
    Set GetPatiOutPathInfo = rsRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub FuncViewDrugExplain(ByVal lngҩƷID As Long, objParent As Object)
'���ܣ��鿴ҩƷ˵����
    Dim objDrugExplain As New frmDrugExplain
    
    If lngҩƷID = 0 Then
        MsgBox "��ǰҩƷδ������´���ܲ鿴˵���顣", vbInformation, gstrSysName
        Exit Sub
    End If
    objDrugExplain.ShowMe lngҩƷID, objParent
End Sub

Public Function InitObjBlood(Optional ByVal blnMsg As Boolean = True) As Boolean
'�ж����Ѫ�ⲿ��Ϊ�վͳ�ʼ��
    If gobjPublicBlood Is Nothing Then
        On Error Resume Next
        Set gobjPublicBlood = CreateObject("zlPublicBlood.clsPublicBlood")
        If Not gobjPublicBlood Is Nothing Then
            If gobjPublicBlood.zlInitCommon(gcnOracle, gstrDBUser) = False Then
                Set gobjPublicBlood = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    InitObjBlood = Not gobjPublicBlood Is Nothing
    If InitObjBlood = False And blnMsg = True Then
        MsgBox "Ѫ�⹫������[zlPublicBlood]����ʧ�ܣ�����Ӱ����Ѫ���̼�����ʹ�ã�����˲����Ƿ�����Լ��Ƿ���ȷע�ᣡ", vbInformation, gstrSysName
    End If
End Function

Public Function GetBloodCapacity(ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strDate As String, ByVal bln����24H�� As Boolean, _
    Optional ByVal intBaby As Integer, Optional ByVal lngActiveID As Long = 0) As Double
'���ܣ����㲡��24Сʱ��Ѫ���򱾴ξ�����Ѫ����
'������int����--1 ����;2-סԺ
'         lng����ID��������ݱ�ʶID
'         lng����ID---int����=1Ϊ�Һ�ID��int����=2Ϊ��ҳID
'         strDate---ʱ��
'         bln����24H��---true����24����,false���ξ�������
'         intBaby---Ӥ�����
'         lngActiveID--ҽ��ID����Ϊ0���������޳���ҽ��
    Dim rsTemp As New Recordset
    Dim dblNum As Double
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    If int���� = 2 Then
        strSQL = _
            " Select Id, ����ʱ��, ������, ��Ѫʱ��" & vbNewLine & _
            " From (With ҽ����¼ As (Select Decode(Nvl(c.ҽ��id, 0), 0, b.������Ŀid, c.������Ŀid) ������Ŀid, b.Id, b.����ʱ��," & vbNewLine & _
            "                            Decode(Nvl(c.ҽ��id, 0), 0, b.�ܸ�����, c.������) ������," & vbNewLine & _
            "                            Nvl(To_Char(b.����ʱ��, 'YYYY-MM-DD HH24:MI'), b.�걾��λ) As ��Ѫʱ��" & vbNewLine & _
            "                     From ��Ѫ������Ŀ c, ������ĿĿ¼ p, ����ҽ����¼ q, ����ҽ����¼ b" & vbNewLine & _
            "                     Where c.ҽ��id(+) = b.Id And p.Id = q.������Ŀid And (p.�������� = 9 Or p.�������� = 8 And p.ִ�з��� = 0) And" & vbNewLine & _
            "                           q.���id = b.Id And q.������� = 'E' And b.����id = [1] And b.��ҳid = [2] And Nvl(b.Ӥ��, 0) = [3] And" & vbNewLine & _
            "                           b.������� = 'K' And b.ҽ��״̬ Not In (-1, 2, 4))" & vbNewLine & _
            "        Select b.Id, b.����ʱ��, b.������ * Decode(Upper(a.���㵥λ), 'ML', 1, Nvl(a.����ϵ��, 1)) ������, b.��Ѫʱ��" & vbNewLine & _
            "        From ������ĿĿ¼ a, ҽ����¼ b" & vbNewLine & _
            "        Where a.Id = b.������Ŀid)"
    ElseIf int���� = 1 Then
        strSQL = _
            " Select Id, ����ʱ��, ������, ��Ѫʱ��" & vbNewLine & _
            " From (With ҽ����¼ As (Select Decode(Nvl(c.ҽ��id, 0), 0, b.������Ŀid, c.������Ŀid) ������Ŀid, b.Id, b.����ʱ��," & vbNewLine & _
            "                            Decode(Nvl(c.ҽ��id, 0), 0, b.�ܸ�����, c.������) ������," & vbNewLine & _
            "                            Nvl(To_Char(b.����ʱ��, 'YYYY-MM-DD HH24:MI'), b.�걾��λ) As ��Ѫʱ��" & vbNewLine & _
            "                     From ��Ѫ������Ŀ c, ������ĿĿ¼ p, ����ҽ����¼ q, ����ҽ����¼ b, ���˹Һż�¼ d" & vbNewLine & _
            "                     Where c.ҽ��id(+) = b.Id And p.Id = q.������Ŀid And (p.�������� = 9 Or p.�������� = 8 And p.ִ�з��� = 0) And" & vbNewLine & _
            "                           q.���id = b.Id And q.������� = 'E' And b.������� = 'K' And b.ҽ��״̬ Not In (-1, 2, 4) And b.�Һŵ� = d.No And" & vbNewLine & _
            "                           d.����id = [1] And d.Id = [2])" & vbNewLine & _
            "        Select b.Id, b.����ʱ��, b.������ * Decode(Upper(a.���㵥λ), 'ML', 1, Nvl(a.����ϵ��, 1)) ������, b.��Ѫʱ��" & vbNewLine & _
            "        From ������ĿĿ¼ a, ҽ����¼ b" & vbNewLine & _
            "        Where a.Id = b.������Ŀid)"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ѫ��", lng����ID, lng����ID, intBaby)
    If Not rsTemp.BOF Then
        rsTemp.MoveFirst
        dblNum = 0
        Do While Not rsTemp.EOF
            If lngActiveID <> Val(rsTemp!ID & "") Then
                If bln����24H�� = False Then
                    dblNum = dblNum + Val("" & rsTemp("������"))
                Else
                    If rsTemp("��Ѫʱ��") & "" <> "" Then
                        If CDate(rsTemp("��Ѫʱ��")) > CDate(Format(strDate, "yyyy-MM-dd HH:mm:ss")) - 1 And CDate(rsTemp("��Ѫʱ��")) <= CDate(Format(strDate, "yyyy-MM-dd HH:mm:ss")) Then dblNum = dblNum + Val("" & rsTemp("������"))
                    Else
                        If CDate(rsTemp("����ʱ��")) > CDate(Format(strDate, "yyyy-MM-dd HH:mm:ss")) - 1 And CDate(rsTemp("����ʱ��")) <= CDate(Format(strDate, "yyyy-MM-dd HH:mm:ss")) Then dblNum = dblNum + Val("" & rsTemp("������"))
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    GetBloodCapacity = dblNum
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetBloodVerifyState(ByVal int������Դ As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str����ʱ�� As String, ByVal dblTotalByML As Double, _
    Optional ByVal int���� As Integer, Optional ByVal int��Ѫ As Integer, Optional ByVal intBaby As Integer, Optional ByVal lngActiveID As Long) As String
'���ܣ�������Ѫ���ģʽ�����¼��㱸Ѫҽ�������״̬(��Ѫ���������)��Ŀǰֻ֧��������Ѫ��ϵͳ�����������뵥��ʽ�´�����
'������
'   int������Դ��1-���2-סԺ
'   lng����ID������ID
'   lng����ID��int������Դ=1Ϊ�Һ�ID��int������Դ=2Ϊ��ҳID
'   str����ʱ�䣺��Ѫ�����Ԥ����Ѫʱ�䣬���ڼ���24Сʱ����ʹ��
'   dblTotalByML��������Ѫ����������ע�⣺����Ϊ������ML��λ����
'   int�������Ƿ��ǽ���ҽ��������ҽ���������-0���ǽ�����1��������
'   int��Ѫ���Ƿ��Ǳ�Ѫҽ����ֻ�б�Ѫҽ������ˣ�0����Ѫ ��1����Ѫ
'   intBaby��Ӥ�����
'   lngActiveID��ҽ��ID��ע�⣺��Ϊ0ʱ�����μ����24Сʱ��������������ҽ����Ӧ������������Ҫ��������뵥�޸�ʱ�޳�֮ǰ������
    Dim strPrivs As String
    Dim dbl24h�� As Double
    Dim str���״̬ As String
    
    On Error GoTo ErrHand
    str���״̬ = GetBloodState(int����, int��Ѫ)
    If str���״̬ = "1" And gblnѪ��ϵͳ = True Then '���״̬�Ѿ��������Ǳ�Ѫ����
        ' ����ҽ���Ȩ�޵ģ���ֱ��ͨ��;����ҽ��Ȩ�޵�<1600��ֱ��ͨ��,��������˹���
        strPrivs = GetInsidePrivs(p��Ѫ��˹���)
        If gbln��Ѫ����������� = True Then
            If InStr(";" & strPrivs & ";", ";ҽ���;") = 0 Then
                '�޸�ʱ����Ҫ�޳���ҽ������GetBloodCapacity���ų���
                dbl24h�� = GetBloodCapacity(int������Դ, lng����ID, lng����ID, str����ʱ��, True, intBaby, lngActiveID)
                
                If InStr(";" & strPrivs & ";", ";������;") <> 0 Then   'ֻ���п�����Ȩ��
                     If dbl24h�� < 1600 - dblTotalByML Then 'С��1600ֱ��ͨ��
                         str���״̬ = 4
                     Else
                         str���״̬ = 7 '����1600������ҽ�����˻���
                     End If
                Else
                    '��û��ҽ���Ҳû�п�����Ȩ��
                    If dbl24h�� < 800 - dblTotalByML Then 'С��800�����������ϼ�ҽʦ��ֱ�ӹ�,���������󻷽�
                        If UserInfo.רҵ����ְ�� = "����ҽʦ" Or UserInfo.רҵ����ְ�� = "������ҽʦ" Or UserInfo.רҵ����ְ�� = "����ҽʦ" Then
                            str���״̬ = 4
                        End If
                    Else
                        'С��1600�ģ����������ϼ�ҽʦ����������˷�����,���������󻷽�
                        If dbl24h�� < 1600 - dblTotalByML And (UserInfo.רҵ����ְ�� = "����ҽʦ" Or UserInfo.רҵ����ְ�� = "������ҽʦ" Or UserInfo.רҵ����ְ�� = "����ҽʦ") Then
                            str���״̬ = 7
                        End If
                    End If
                End If
            Else
                str���״̬ = 4
            End If
        Else
            str���״̬ = IIF(strPrivs <> "", 4, 1)
        End If
    End If
    GetBloodVerifyState = str���״̬
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetBloodTotalByML(ByVal strItems As String) As Double
'���ܣ�����������Ŀ��������ȡ��Ѫҽ��������(ML)
'������strItems:������Ŀ��Ϣ,��ʽ��������ĿID,������;������ĿID,������
    Dim i As Integer
    Dim arrItem
    Dim strIDs As String
    Dim objItem As New Collection
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim dblNum As Double, dblTotal As Double
    
    If strItems = "" Then Exit Function
    On Error GoTo ErrHand
    arrItem = Split(strItems, ";")
    For i = 0 To UBound(arrItem)
        If InStr(1, "," & strIDs & ",", "," & Split(arrItem(i), ",")(0) & ",") = 0 Then
            strIDs = strIDs & "," & Split(arrItem(i), ",")(0)
            objItem.Add Val(Split(arrItem(i), ",")(1)), "_" & Split(arrItem(i), ",")(0)
        End If
    Next
    If Left(strIDs, 1) = "," Then strIDs = Mid(strIDs, 2)
    strSQL = "Select /*+ CARDINALITY(b 10) */" & vbNewLine & _
        "  a.id,a.����, a.���㵥λ, a.����ϵ��" & vbNewLine & _
        " From ������ĿĿ¼ a, Table(f_Num2list([1])) b" & vbNewLine & _
        " Where a.Id = b.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetBloodTotalByML", strIDs)
    Do While Not rsTemp.EOF
        dblNum = Val(objItem("_" & rsTemp!ID))
        If UCase(rsTemp!���㵥λ & "") <> "ML" Then
            dblNum = dblNum * Val(NVL(rsTemp!����ϵ��, 1))
        End If
        dblTotal = dblTotal + dblNum
        rsTemp.MoveNext
    Loop
    GetBloodTotalByML = dblTotal
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetԭҺƤ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String) As ADODB.Recordset
'���ܣ������˻�ȡԭҺƤ�Ե�ҽ����Ŀ
'���������ﲡ�˴� str�Һŵ���               lng����ID, lng��ҳID ��Ϊ0
'      סԺ���˴��� lng����ID, lng��ҳID �� str�Һŵ� ���մ�
'���أ���¼����ʽ
'˵������  ����ҽ������.�걾�������� ������ҩƷ�к͹���ʵ��ҽ���У�ʹ���� ҩƷ�е�ҽ��ID��
'      ���� һ��ԭҺ����ʵ����Ŀ���շѶ����ж�����ҩƷ����(ҩƷA)�������´��Ƥ�Բ��ұ�ǽ��Ϊ���ԣ��������ٴη���ҩƷҽ��Aʱ����-1
'      һ����-1������ϱ�Ǻ������Ͳ����ټ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim rsƤ�� As ADODB.Recordset
    Dim str��� As String
    
    On Error GoTo errH
    
    'ԭҺƤ�ԣ����Ϊ���ԣ��걾�������� ������
    If str�Һŵ� = "" Then
        strSQL = "Select b.ҽ��Id as Ƥ��ҽ��ID,nvl(b.�걾��������,0) as ���,max(e.id) as ҩƷID" & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C, �����շѹ�ϵ D, �շ���ĿĿ¼ E" & vbNewLine & _
            "Where a.Ƥ�Խ�� = '(-)' And a.Id = b.ҽ��id And a.������Ŀid = c.Id And c.��� = 'E' And c.�������� = '1' And c.ִ�з��� = 5" & vbNewLine & _
            "And  c.Id=d.������Ŀid And d.�շ���Ŀid = e.Id And e.��� In ('5', '6') And a.����ID =[1] and a.��ҳID=[2]" & vbNewLine & _
            "group by b.ҽ��id,b.�걾��������"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    Else
        strSQL = "Select b.ҽ��Id as Ƥ��ҽ��ID,nvl(b.�걾��������,0) as ���,max(e.id) as ҩƷID" & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C, �����շѹ�ϵ D, �շ���ĿĿ¼ E" & vbNewLine & _
            "Where a.Ƥ�Խ�� = '(-)' And a.Id = b.ҽ��id And a.������Ŀid = c.Id And c.��� = 'E' And c.�������� = '1' And c.ִ�з��� = 5" & vbNewLine & _
            "And  c.Id=d.������Ŀid And d.�շ���Ŀid = e.Id And e.��� In ('5', '6') And a.�Һŵ� =[1]" & vbNewLine & _
            "group by b.ҽ��id,b.�걾��������"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str�Һŵ�)
    End If
    Set rsƤ�� = zlDatabase.CopyNewRec(rsTmp)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!��� & "") <> 0 Then
                str��� = str��� & "," & Val(rsTmp!��� & "")
            End If
            rsTmp.MoveNext
        Next
        
        If str��� <> "" Then
            strSQL = "select a.id from ����ҽ����¼ a,����ҽ������ b where a.id=b.ҽ��ID and a.id in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X) group by a.id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", Mid(str���, 2))
            str��� = ""
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    str��� = str��� & "," & rsTmp!ID '�Ѿ����������˵�
                    rsTmp.MoveNext
                Next
            End If
            For i = 1 To rsƤ��.RecordCount
                If Val(rsƤ��!��� & "") <> 0 Then
                    If InStr(str��� & ",", "," & rsƤ��!��� & ",") = 0 Then
                        rsƤ��!��� = 0
                    End If
                End If
                rsƤ��.MoveNext
            Next
            rsƤ��.MoveFirst
        End If
    End If
    Set GetԭҺƤ�� = rsƤ��
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetԭҺƤ��ҩƷ(ByVal lng��Ŀid As Long) As Long
'���ܣ���ȡԭҺƤ�Զ�Ӧ��ҩƷ��Ϣ
'������lng��ĿID ԭҺƤ����ĿID
'���أ��շѶ����е� ҩƷ�շ���ĿID�����ID��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Max(b.Id) As ҩƷid" & vbNewLine & _
        "From �����շѹ�ϵ A, �շ���ĿĿ¼ B" & vbNewLine & _
        "Where a.�շ���Ŀid = b.Id And b.��� In ('5', '6') And a.������Ŀid =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid)
    
    If Not rsTmp.EOF Then
        GetԭҺƤ��ҩƷ = Val(rsTmp!ҩƷID & "")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BloodApplyPrintCheck(ByVal lngҽ��ID As Long, ByVal int���� As Integer, ByVal int���� As Integer, ByVal intģʽ As Integer) As Boolean
'����˵��:���뵥Ԥ����ӡʱ��ҽ����������ݽ��м�飬��������ʾ����������
' --���˵����
' ----���ó���_in=1-����,2-סԺ
' ----��������_In=1-��Ѫ���뵥;2-ȡѪ֪ͨ��(����ҽԺ�����������Ϳ���)
' ----��Ѫ����_In=0-��ͨ��Ѫ;1-������Ѫ(����ҽԺ������Ѫ�����̶ȿ���)
' ----ģʽ_in:0=Ԥ���ǵ��ã�1-��ӡ�ǵ���
' --�������أ�"������|��ʾ��Ϣ",������=0-����,1-ѯ����ʾ,2-��ֹ��������Ϊ0ʱ�����践����ʾ��Ϣ���ָ�����
'���أ�TRUE����,False��ֹ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strMsg As String
    
    On Error GoTo ErrHand
    strSQL = "Select Zl1_Fun_BloodApplyPrint([1],[2],[3],[4]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Zl1_Fun_BloodApplyPrint", lngҽ��ID, int����, int����, intģʽ)
    If Not rsTmp.EOF Then
        strMsg = NVL(rsTmp!���)
        If strMsg <> "" Then
            Select Case Val(Split(strMsg, "|")(0))
            Case 1 '��ʾ
                If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    strMsg = "": Exit Function
                End If
            Case 2 '��ֹ
                MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                strMsg = "": Exit Function
            End Select
            strMsg = ""
        End If
    End If
                
    BloodApplyPrintCheck = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetApplyCustom(ByVal lng��Ŀid As Long) As Long
'���ܣ���ȡ�Զ������뵥���ļ�ID
'������lng��ĿID ������ĿID
'���أ��Զ������뵥���ļ�ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.�����ļ�ID" & vbNewLine & _
        "From ��������Ӧ�� A, �����ļ��б� B" & vbNewLine & _
        "Where a.�����ļ�id = b.Id And b.���� = 7 And Nvl(b.��ʽ, 0) = 1 And a.������Ŀid =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid)
    
    If Not rsTmp.EOF Then
        GetApplyCustom = Val(rsTmp!�����ļ�ID & "")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get����ҽ��IDs(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strDepts As String) As String
'���ܣ���ȡ����ָ�����ҵĻ���ҽ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String
    
    On Error GoTo errH

    strSQL = "Select a.id From ����ҽ����¼ A, ������ĿĿ¼ B Where a.������Ŀid = b.Id And a.������� = 'Z' And b.�������� = '7' And a.ҽ��״̬ = 8 And a.����id =[1] And a.��ҳid =[2] And a.��������id In (Select /*+cardinality(x,10)*/ x.Column_Value  From Table(f_Num2list([3])) X )"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID, strDepts)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strTmp = strTmp & "," & rsTmp!ID
            rsTmp.MoveNext
        Next
        Get����ҽ��IDs = Mid(strTmp, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckCanSendAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngSpecialAdviceID As Long, ByVal intBabyNum As Integer) As Boolean
'���ܣ���ֹ����ǰ����Ŀ�ʼִ��ʱ�����Ƿ��п��Է��͵ĳ���ҽ��
'������strSpecialAdviceIDs=����ҽ��IDת��ҽ��
'      intBabyNum=Ӥ����š�
'���أ�true ���ڣ�false ������
'˵�����ú���ִ��ʱֻ����ĸ������ҽ��ʱ����Զ�Ӥ��ҽ���ļ�飬���ͬʱ������������
'      ����  ����ҽ������Һҽ��ֻ����������

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strEnd As String
    Dim rsSend As ADODB.Recordset
    Dim i As Long
    Dim strPause As String
    Dim datBegin As Date
    Dim datEnd As Date
    Dim str�״�ʱ�� As String
    Dim strĩ��ʱ�� As String
    Dim str�ֽ�ʱ�� As String
    Dim lng���� As Long
    Dim lng��ID As Long
    
    
    On Error GoTo errH
    
    If Val(zlDatabase.GetPara("����δ����ҽ��ʱ��ֹ����ת��ҽ��", glngSys, pסԺҽ������, 0)) = 0 Then Exit Function
    
    '�����жϳ����ǲ����Ѿ�ֹͣ��У��/���� ת��ҽ��ʱ���Զ�ͣ����ͣ��ʱ����ת��ҽ����[��ʼִ��ʱ��]
    '����ҽ���ǲ��ǿ��Է��ͣ���ֹʱ��Ϊת��ҽ���Ŀ�ʼִ��ʱ�䣬ҽ����Ч���ǳ�������У��״̬
    strSQL = "Select to_char(b.��ʼִ��ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���� From ����ҽ����¼ b Where b.id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ҽ��", lngSpecialAdviceID)
    If Not rsTmp.EOF Then
        strEnd = rsTmp!���� & ""
        '������ҩ����
        strSQL = "select a.id, a.���ID,Nvl(A.���ID,A.ID) as ��ID,a.�������,a.ҽ��״̬,a.��ʼִ��ʱ��,a.�ϴ�ִ��ʱ��,a.ִ��ʱ�䷽��,a.Ƶ�ʴ���,a.Ƶ�ʼ��,a.�����λ,a.ִ����ֹʱ��" & _
            " from ����ҽ����¼ a,������ĿĿ¼ E where a.������ĿID=e.id and a.����id=[2] and a.��ҳid=[3] and nvl(a.Ӥ��,0)=[4]" & _
            " And Nvl(A.ҽ��״̬,0) Not IN(-1,1,2,4) And A.��ʼִ��ʱ��<=[1] And (A.�ϴ�ִ��ʱ��<[1] Or A.�ϴ�ִ��ʱ�� is NULL)" & _
            " And (A.ִ����ֹʱ��>A.�ϴ�ִ��ʱ�� Or A.ִ����ֹʱ�� is NULL Or A.�ϴ�ִ��ʱ�� Is NULL)" & _
            " And (A.ִ����ֹʱ��>A.��ʼִ��ʱ�� Or A.ִ����ֹʱ�� is NULL) And A.ҽ����Ч=0" & _
            " And Not(Nvl(a.�������,'����')='H' And E.��������='1' And E.ִ��Ƶ��=2)" & _
            " And Not(Nvl(a.�������,'����')='Z' And E.�������� IN ('4','14','9','10','12'))" & _
            " And Not (a.������� = 'E' And a.���id Is Not Null And e.�������� = '3') And Nvl(a.�������, '����') Not In ('5', '6', '7') And Not Exists (Select ID From ����ҽ����¼ X Where ������� In ('5', '6', '7') And x.���id = a.id)" & _
            " And a.��ʼִ��ʱ�� Is Not Null And Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3 And Nvl(a.ִ��Ƶ��, '��') <> '��Ҫʱ' And Nvl(a.ִ��Ƶ��, '��') <> '��Ҫʱ'" & _
            " order by a.���"
        Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "���ҽ��", CDate(strEnd), lng����ID, lng��ҳID, intBabyNum)
        For i = 1 To rsSend.RecordCount
            If IsNull(rsSend!���ID) Then
                strPause = GetAdvicePause(rsSend!ID)
                '��ǰҽ���ķ��ͼ���ʱ���
                datBegin = rsSend!��ʼִ��ʱ��
                If Not IsNull(rsSend!�ϴ�ִ��ʱ��) Then
                    If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                        datBegin = DateAdd("s", 1, rsSend!�ϴ�ִ��ʱ��)
                    Else
                        datBegin = Calc�����ڿ�ʼʱ��(rsSend!��ʼִ��ʱ��, rsSend!�ϴ�ִ��ʱ��, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                        If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                    End If
                End If
                datEnd = CDate(strEnd)
                If Not IsNull(rsSend!ִ����ֹʱ��) Then
                    If rsSend!ִ����ֹʱ�� < CDate(strEnd) Then
                        datEnd = rsSend!ִ����ֹʱ��
                    End If
                End If
                '����ֽ�ʱ�估����
                If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                    lng���� = Calc�����Գ�������(datBegin, datEnd, Format(NVL(rsSend!�ϴ�ִ��ʱ��), "yyyy-MM-dd HH:mm:ss"), Format(NVL(rsSend!ִ����ֹʱ��), "yyyy-MM-dd HH:mm:ss"), strPause)
                    If lng���� <> 0 Then
                        CheckCanSendAdvice = True
                        Exit Function
                    End If
                Else
                    'ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ
                    str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, datEnd, strPause, NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ, rsSend!��ʼִ��ʱ��)
                    If str�ֽ�ʱ�� <> "" Then
                        CheckCanSendAdvice = True
                        Exit Function
                    End If
                End If
            End If
            rsSend.MoveNext
        Next
                    
        'ҩƷ����
        strSQL = "select a.id,a.���ID,Nvl(A.���ID,A.ID) as ��ID,a.�������,a.ҽ��״̬,a.��ʼִ��ʱ��,a.�ϴ�ִ��ʱ��,a.ִ��ʱ�䷽��,a.Ƶ�ʴ���,a.Ƶ�ʼ��,a.�����λ,a.ִ����ֹʱ��" & _
            " from ����ҽ����¼ a,������ĿĿ¼ E where a.������ĿID=e.id and a.����id=[2] and a.��ҳid=[3] and nvl(a.Ӥ��,0)=[4]" & _
            " And Nvl(A.ҽ��״̬,0) Not IN(-1,1,2,4) And A.��ʼִ��ʱ��<=[1] And (A.�ϴ�ִ��ʱ��<[1] Or A.�ϴ�ִ��ʱ�� is NULL)" & _
            " And (A.ִ����ֹʱ��>A.�ϴ�ִ��ʱ�� Or A.ִ����ֹʱ�� is NULL Or A.�ϴ�ִ��ʱ�� Is NULL)" & _
            " And (A.ִ����ֹʱ��>A.��ʼִ��ʱ�� Or A.ִ����ֹʱ�� is NULL) And A.ҽ����Ч=0" & _
            " And a.������� In ('5', '6', '7') And a.��ʼִ��ʱ�� Is Not Null And Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3 And Nvl(a.ִ��Ƶ��, '��') <> '��Ҫʱ' And Nvl(a.ִ��Ƶ��, '��') <> '��Ҫʱ'" & _
            " order by a.���"
        Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "���ҽ��", CDate(strEnd), lng����ID, lng��ҳID, intBabyNum)
        For i = 1 To rsSend.RecordCount
            If lng��ID <> Val(rsSend!��ID & "") Then
                lng��ID = Val(rsSend!��ID & "")
                strPause = GetAdvicePause(rsSend!ID)
                '��ǰҽ���ķ��ͼ���ʱ���
                datBegin = rsSend!��ʼִ��ʱ��
                If Not IsNull(rsSend!�ϴ�ִ��ʱ��) Then
                    If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                        datBegin = DateAdd("s", 1, rsSend!�ϴ�ִ��ʱ��) '"������"����Ŀ
                    Else
                        datBegin = Calc�����ڿ�ʼʱ��(rsSend!��ʼִ��ʱ��, rsSend!�ϴ�ִ��ʱ��, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                        If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                    End If
                End If
                datEnd = CDate(strEnd)
                If Not IsNull(rsSend!ִ����ֹʱ��) Then
                    If rsSend!ִ����ֹʱ�� < CDate(strEnd) Then
                        datEnd = rsSend!ִ����ֹʱ��
                    End If
                End If
                '����ֽ�ʱ�估����
                If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                    lng���� = Calc�����Գ�������(datBegin, datEnd, Format(NVL(rsSend!�ϴ�ִ��ʱ��), "yyyy-MM-dd HH:mm:ss"), Format(NVL(rsSend!ִ����ֹʱ��), "yyyy-MM-dd HH:mm:ss"), strPause)
                    If lng���� <> 0 Then
                        CheckCanSendAdvice = True
                        Exit Function
                    End If
                Else
                    'ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ
                    str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, datEnd, strPause, NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ, rsSend!��ʼִ��ʱ��)
                    If str�ֽ�ʱ�� <> "" Then
                        CheckCanSendAdvice = True
                        Exit Function
                    End If
                End If
            End If
            rsSend.MoveNext
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SvrԤԼ��Ժȡ������(ByVal lng�Һ�ID As Long)
'���ܣ�����ԤԼ���ķ���ȡ��סԺ����
    Dim strJsIn As String
    Dim strJsOut As String
    Dim strErr As String
   
    strJsIn = "{""input_in"":{""rgst_id"": """ & lng�Һ�ID & """}}"
    Call Sys.NewSystemSvr("ԤԼ����", "סԺ����ȡ��", strJsIn, strJsOut, strErr)
    
    If strErr <> "" Then
        MsgBox "ԤԼ��Ժȡ������:" & strErr, vbInformation, gstrSysName
    End If
End Sub

Private Function GetParaURL(ByVal strSysName As String, ByVal strServiceName As String) As String
'����:
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strUrl As String
    
    On Error GoTo errH
    strSQL = "Select �����ַ From ������������Ŀ¼ Where ϵͳ��ʶ = [1] And �������� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strSysName, strServiceName)
    If Not rsTmp.EOF Then strUrl = Trim(rsTmp!�����ַ & "")
    GetParaURL = strUrl
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get��Һ��ҽ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intType As Integer) As String
'���ܣ���ȡָ�����˵�ĳ��������Һҽ��
'������intType ���ó��ϣ�0 - һ�㷢�ʹ��ڣ�1 - ��Һ���ʹ���
'˵����Ӫ�����Ա�ҩ/�ȵ��ء���ȡҩ����Ժ��ҩ
'���أ���Ҫ�ӷ��͵�ҽ�����ų�����  ҽ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strIDs As String
    Dim str��Ч As String
    Dim lngPar��Ч As Long
    Dim strPar��ҩ As String
    Dim str�Ա� As String
    Dim str�Ա���Һ As String
    Dim str��ȡ As String
    Dim str��Ժ As String
    Dim strӪ�� As String
    Dim str���� As String '��ͨ��������Һҽ��
    Dim strAll As String
    Dim blnOK As Boolean
    Dim strNoInIDs As String
    Dim strTmpALL As String
    Dim varTmp As Variant
    Dim str���� As String
    
    On Error GoTo errH
    
    lngPar��Ч = Val(zlDatabase.GetPara("ҽ������", glngSys, p��Һ��������, "1")) - 1
    If lngPar��Ч = -1 Then
        str��Ч = ",0,1,"
    Else
        str��Ч = lngPar��Ч
    End If
     
    strPar��ҩ = zlDatabase.GetPara("��Һ��ҩ;��", glngSys, p��Һ��������)
    
    'һ�㷢�ʹ��ں���Һ���ʹ������߹���ҽ��ǡ���෴�����ھ��䷢�ľͲ���һ�㷢������ʾ����ȡ����ֵ�ж�ʱ�������intType��
    str���� = ""
    blnOK = Val(zlDatabase.GetPara("�Ա�ҩ��������������", glngSys, p��Һ��������, "0")) = intType
    str���� = str���� & IIF(blnOK, "1", "0")
    blnOK = Val(zlDatabase.GetPara("��ȡҩ��������������", glngSys, p��Һ��������, "0")) = intType
    str���� = str���� & IIF(blnOK, "1", "0")
    blnOK = Val(zlDatabase.GetPara("��Ժ��ҩ��������������", glngSys, p��Һ��������, "0")) = intType
    str���� = str���� & IIF(blnOK, "1", "0")
         
    strSQL = "Select a.id,a.���id,a.ִ������ As ҩƷִ������, a.ִ�б�� As ҩƷִ�б��, b.ִ������ As ��ҩִ������, b.ִ�б�� As ��ҩִ�б��," & vbNewLine & _
        " c.�������� as ��ҩ��������,c.ִ�з��� as ��ҩִ�з���,c.ִ�б�� as ��ҩִ�б��,nvl(a.ִ�п���id,0) as ҩƷִ�п���id,c.id as ��ҩ��Ŀid,d.ҩƷid as ��Һ�Ա�,a.ҽ����Ч" & vbNewLine & _
        " From ����ҽ����¼ A, ����ҽ����¼ B,������ĿĿ¼ c,��Һ�Ա�ҩ�嵥 d" & vbNewLine & _
        " Where a.���id = b.Id and b.������Ŀid=c.id and a.�շ�ϸĿid=d.ҩƷid(+) And a.������� In ('5', '6')" & vbNewLine & _
        " And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3 and nvl(a.ҽ��״̬,0)<>4 and a.����id=[1] and a.��ҳid=[2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get��Һ��ҽ��", lng����ID, lng��ҳID)
    
    For i = 1 To rsTmp.RecordCount
        blnOK = True
        
        If blnOK Then
            '��ҩ;����Һ��
            If Val(rsTmp!��ҩ�������� & "") <> 2 Or Val(rsTmp!��ҩִ�з��� & "") <> 1 Then
                blnOK = False
            End If
        End If
        
        If blnOK Then
            '�����ҩ;������
            If strPar��ҩ <> "" Then
                If InStr("," & strPar��ҩ & ",", "," & rsTmp!��ҩ��Ŀid & ",") = 0 Then
                    blnOK = False
                End If
            End If
        End If
        
        If blnOK Then
            '������Ч����
            If InStr("," & str��Ч & ",", "," & rsTmp!ҽ����Ч & ",") = 0 Then
                blnOK = False
            End If
        End If
        
        If blnOK Then
            '������Ҳ���
            If InStr("," & gstr��Һ�������� & ",0,", "," & rsTmp!ҩƷִ�п���id & ",") = 0 Then
                blnOK = False
            End If
        End If
        
        If blnOK Then
            If intType = 0 Then
                'һ�㷢�ʹ���
                '����ľ���ҽ�����������ٴ���
                If Val(rsTmp!ҩƷִ�п���id & "") > 0 Then
                    If InStr("," & strTmpALL & ",", "," & rsTmp!���ID & ",") = 0 Then
                        strTmpALL = strTmpALL & "," & rsTmp!���ID
                    End If
                End If
            End If
            
            
            If Val(rsTmp!��ҩִ������ & "") <> 5 And Val(rsTmp!ҩƷִ������ & "") = 5 And Val(rsTmp!ҩƷִ�б�� & "") = 2 Then
                '��ȡҩ
                If InStr("," & strAll & ",", "," & rsTmp!���ID & ",") = 0 Then
                    str��ȡ = str��ȡ & "," & rsTmp!���ID
                    strAll = strAll & "," & rsTmp!���ID
                End If
            ElseIf Val(rsTmp!��ҩִ������ & "") <> 5 And Val(rsTmp!ҩƷִ������ & "") = 5 And Val(rsTmp!ҩƷִ�б�� & "") <> 2 Then
                '�Ա�ҩ--���һ��ϸ�֣�һ���Ա�ҩ���þ����Ա�ҩ
                '��Һ�Ա�
                If Val(rsTmp!��Һ�Ա� & "") = 0 Then
                    If InStr("," & strAll & ",", "," & rsTmp!���ID & ",") = 0 Then
                        str�Ա� = str�Ա� & "," & rsTmp!���ID
                        strAll = strAll & "," & rsTmp!���ID
                    End If
                Else
                    If InStr("," & strAll & ",", "," & rsTmp!���ID & ",") = 0 Then
                        str�Ա���Һ = str�Ա���Һ & "," & rsTmp!���ID
                        strAll = strAll & "," & rsTmp!���ID
                    End If
                End If
            ElseIf Val(rsTmp!��ҩִ������ & "") = 5 And Val(rsTmp!ҩƷִ������ & "") <> 5 Then
                '��Ժ��ҩ
                If InStr("," & strAll & ",", "," & rsTmp!���ID & ",") = 0 Then
                    str��Ժ = str��Ժ & "," & rsTmp!���ID
                    strAll = strAll & "," & rsTmp!���ID
                End If
            ElseIf Val(rsTmp!��ҩִ�б�� & "") = 2 Then
                'Ӫ��
                If InStr("," & strAll & ",", "," & rsTmp!���ID & ",") = 0 Then
                    strӪ�� = strӪ�� & "," & rsTmp!���ID
                    strAll = strAll & "," & rsTmp!���ID
                End If
            End If
        End If
        rsTmp.MoveNext
    Next
    
    
    If strTmpALL <> "" Then
        varTmp = Split(strTmpALL, ",")
        strTmpALL = ""
        For i = 0 To UBound(varTmp)
            If InStr("," & str�Ա� & str�Ա���Һ & str��ȡ & str��Ժ & strӪ�� & ",", "," & varTmp(i) & ",") = 0 Then
                strTmpALL = strTmpALL & "," & varTmp(i)
            End If
        Next
    End If
    
    'Ӫ����̶��ų�
    '���������ò������Ա�ҩʱ��  str�Ա���Һ ��Ҫ�����²������ų�
    '����Ϊ���Ա�ҩ����ȡҩ����Ժ��ҩ
    Select Case str����
    Case "000"
        If str�Ա� <> "" Then
            strNoInIDs = strNoInIDs & str�Ա�
        End If
        If str��ȡ <> "" Then
            strNoInIDs = strNoInIDs & str��ȡ
        End If
        If str��Ժ <> "" Then
            strNoInIDs = strNoInIDs & str��Ժ
        End If
    Case "001"
        If str�Ա� <> "" Then
            strNoInIDs = strNoInIDs & str�Ա�
        End If
        If str��ȡ <> "" Then
            strNoInIDs = strNoInIDs & str��ȡ
        End If
'        If str��Ժ <> "" Then
'            strNoInIDs = strNoInIDs & str��Ժ
'        End If
    Case "010"
        If str�Ա� <> "" Then
            strNoInIDs = strNoInIDs & str�Ա�
        End If
'        If str��ȡ <> "" Then
'            strNoInIDs = strNoInIDs & str��ȡ
'        End If
        If str��Ժ <> "" Then
            strNoInIDs = strNoInIDs & str��Ժ
        End If
    Case "011"
        If str�Ա� <> "" Then
            strNoInIDs = strNoInIDs & str�Ա�
        End If
'        If str��ȡ <> "" Then
'            strNoInIDs = strNoInIDs & str��ȡ
'        End If
'        If str��Ժ <> "" Then
'            strNoInIDs = strNoInIDs & str��Ժ
'        End If
    Case "100"
'        If str�Ա� <> "" Then
'            strNoInIDs = strNoInIDs & str�Ա�
'        End If
        If str��ȡ <> "" Then
            strNoInIDs = strNoInIDs & str��ȡ
        End If
        If str��Ժ <> "" Then
            strNoInIDs = strNoInIDs & str��Ժ
        End If
    Case "101"
'        If str�Ա� <> "" Then
'            strNoInIDs = strNoInIDs & str�Ա�
'        End If
        If str��ȡ <> "" Then
            strNoInIDs = strNoInIDs & str��ȡ
        End If
'        If str��Ժ <> "" Then
'            strNoInIDs = strNoInIDs & str��Ժ
'        End If
    Case "110"
'        If str�Ա� <> "" Then
'            strNoInIDs = strNoInIDs & str�Ա�
'        End If
'        If str��ȡ <> "" Then
'            strNoInIDs = strNoInIDs & str��ȡ
'        End If
        If str��Ժ <> "" Then
            strNoInIDs = strNoInIDs & str��Ժ
        End If
    Case "111"
'        If str�Ա� <> "" Then
'            strNoInIDs = strNoInIDs & str�Ա�
'        End If
'        If str��ȡ <> "" Then
'            strNoInIDs = strNoInIDs & str��ȡ
'        End If
'        If str��Ժ <> "" Then
'            strNoInIDs = strNoInIDs & str��Ժ
'        End If
    End Select
    
    
    'һ�㷢�ʹ��ڹ̶��ų�  Ӫ�� ���Ա���Һ����Һ�Ա�ҩ�嵥��������
    If intType = 0 Then
        If strӪ�� <> "" Then
            strNoInIDs = strNoInIDs & strӪ��
        End If
        
        If str�Ա���Һ <> "" Then
            strNoInIDs = strNoInIDs & str�Ա���Һ
        End If
        
        If strTmpALL <> "" Then
            strNoInIDs = strNoInIDs & strTmpALL
        End If
    End If
    
    strNoInIDs = Mid(strNoInIDs, 2)
    Get��Һ��ҽ�� = strNoInIDs
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowSQLSelectCIS(frmParent As Object, ByVal strSQL As String, ByVal strDetail As String, bytStyle As Byte, _
    ByVal strTitle As String, ByVal blnĩ�� As Boolean, _
    ByVal strSeek As String, ByVal strNote As String, _
    ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, _
    ByVal blnNoneWin As Boolean, ByVal X As Long, _
    ByVal Y As Long, ByVal txtH As Long, _
    ByRef Cancel As Boolean, ByVal blnMultiOne As Boolean, _
    ByVal blnSearch As Boolean, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ��๦��ѡ����,ʹ��ADO.Command��,����ʹ��[x]����
'������
'     frmParent=��ʾ�ĸ�����
'     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
'     strDetail  ��ϸ��SQL
'     bytStyle=ѡ�������
'       Ϊ0ʱ:�б���:ID,��
'       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
'       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
'     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
'     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
'     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
'             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
'             bytStyle=1ʱ,�����Ǳ��������
'     strNote=ѡ������˵������
'     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
'     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
'     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
'     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
'     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
'     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
'     arrInput=��Ӧ�ĸ���SQL����ֵ,��˳����,����Ϊ��ȷ���͡����У�
'               ��ʽΪ��"bytSize=?"��ʾ���������С(0-С����,1-������;С����Ϊ9����,������Ϊ12����),Ĭ��С���塣
'               ��ʽΪ��ColSet:...ʱ��ʾ�п�����,ColSet��ʽ:�п�����|����1,���1;����2,���2.....|������ʾ|������
'               ��ʽΪ��HeadCap=SQL����1,�б�չʾ����1;SQL����2,�б�չʾ����2������Ŀ�����ֹ�ָ��SQL�����б���չʾ���ƣ�һ�����ڱ��������У����ǲ��ı��е�Key
'               ��ʽΪ��MultiCheckReturn=0,1����ѡʱֻ���ع�ѡ�У����ڶ�ѡ��ȷ��Ĭ�Ϸ��ص�ǰ���������Ӹò������ƣ��ÿ������ú󣬲�֧��Ĭ���еķ��أ������Ծ�֧��˫�����Զ����ء�
'               ��ʽΪ��HideNullCols=0,1;�Ƿ�����SQl�е�null as д������
'���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
'˵����
'     1.ID���ϼ�ID����Ϊ�ַ�������
'     2.ĩ�����ֶβ�Ҫ����ֵ
'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Dim frmNew As New frmPubSel
    Dim arrPar() As Variant
    arrPar = arrInput
    Set ShowSQLSelectCIS = frmNew.ShowSelect(frmParent, strSQL, strDetail, bytStyle, strTitle, blnĩ��, strSeek, strNote, blnShowSub, _
                                        blnShowRoot, blnNoneWin, X, Y, txtH, Cancel, blnMultiOne, blnSearch, False, arrPar)
End Function

Public Sub LisInfoTrans(ByVal strLIS As String, ByRef rsData As ADODB.Recordset, ByRef rsSub As ADODB.Recordset)
'����:LIS�������뵥��Ϣת��,��LIS���뵥���������ַ�����Ϣת���ɼ�¼��
'����:strLis ���뵥��Ϣ
'     rsData ����,��¼����ʽ
'     rsSub �μ���Ϣ,��ҪΪ������ĿĿ¼�����Ϣ
    Dim arrTmp As Variant
    Dim arrTmp1 As Variant
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim str���� As String
    Dim lng��� As Long
    Dim varRowҽ�� As Variant
    Dim strҽ�� As String
    Dim strALL��ĿIDs As String
    Dim str���� As String
    
    On Error GoTo errH
    
    '��¼��,һ�м�¼����һ��ҽ��
    Set rsData = New ADODB.Recordset
    rsData.Fields.Append "���", adBigInt
    rsData.Fields.Append "ҽ��", adVarChar, 4000
    rsData.Fields.Append "���", adBigInt
    rsData.Fields.Append "�Ƿ�����", adInteger
    rsData.Fields.Append "ʱ��ID", adVarChar, 18
    rsData.Fields.Append "ʱ������", adVarChar, 400
    
    rsData.Fields.Append "�ɼ�����ID", adBigInt
    rsData.Fields.Append "�ɼ���ĿID", adBigInt
    rsData.Fields.Append "ִ�п���ID", adBigInt
    rsData.Fields.Append "������ĿID", adBigInt
    rsData.Fields.Append "��ʼִ��ʱ��", adVarChar, 40
    rsData.Fields.Append "�걾", adVarChar, 400
    rsData.Fields.Append "����", adVarChar, 4000
    rsData.Fields.Append "����", adVarChar, 4000
    rsData.Fields.Append "����", adBigInt
 
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    arrTmp = Split(strLIS, "<Split B>")
    For i = 0 To UBound(arrTmp)
        strҽ�� = arrTmp(i)
        varRowҽ�� = Split(strҽ��, "<Split A>")
        '�������뵥�������ĵļ���ҽ��ֻ��һ��������ĿID
        'varRowҽ��(8)��ͨ������룬�±�ֻ��8������������Ϣ��varRowҽ��(9)��
        If UBound(varRowҽ��) > 8 Then
            str���� = varRowҽ��(9)
            arrTmp1 = Split(str����, "<split2>")
            If Val(arrTmp1(0)) = 1 Then '�ж��Ƿ������ظ�����ҽ����
                For j = 1 To UBound(arrTmp1)
                    varRowҽ�� = Split(arrTmp1(j), "<split3>")
                    lng��� = lng��� + 1
                    rsData.AddNew
                    rsData!��� = lng���
                    rsData!��� = i + 1
                    rsData!ҽ�� = strҽ��
                    rsData!�Ƿ����� = 1
                    rsData!ʱ��ID = Val(varRowҽ��(0))
                    rsData!ʱ������ = varRowҽ��(1)
                    Call LisInfoTransToRs(rsData!ҽ�� & "", rsData, strALL��ĿIDs)
                    rsData.Update
                Next
            Else
                lng��� = lng��� + 1
                rsData.AddNew
                rsData!��� = lng���
                rsData!ҽ�� = strҽ��
                Call LisInfoTransToRs(rsData!ҽ�� & "", rsData, strALL��ĿIDs)
                rsData.Update
            End If
        Else
            lng��� = lng��� + 1
            rsData.AddNew
            rsData!��� = lng���
            rsData!ҽ�� = strҽ��
            Call LisInfoTransToRs(rsData!ҽ�� & "", rsData, strALL��ĿIDs)
            rsData.Update
        End If
    Next
    rsData.MoveFirst
    Set rsSub = Get������Ŀ��¼(0, Mid(strALL��ĿIDs, 2))
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LisInfoTransToRs(ByVal strLIS As String, ByRef rsData As ADODB.Recordset, ByRef strALL��ĿIDs As String)
'����:LIS������Ϣ��ӵ���¼����
'˵��:���ڱ�ģ��� LisInfoTrans �����е���
    Dim varRowҽ�� As Variant
 
    varRowҽ�� = Split(strLIS, "<Split A>")
    
    rsData!�ɼ�����ID = Val(varRowҽ��(0))
    rsData!ִ�п���ID = Val(varRowҽ��(1))
    rsData!��ʼִ��ʱ�� = varRowҽ��(2)
    rsData!�걾 = varRowҽ��(3)
    rsData!���� = varRowҽ��(4)
    rsData!���� = varRowҽ��(5)
    rsData!���� = Val(varRowҽ��(6))
    rsData!�ɼ���ĿID = Val(varRowҽ��(7))
    rsData!������ĿID = Val(varRowҽ��(8)) '���뵥�������ĵļ���ҽ��ֻ��һ��������ĿID��������һ���ɼ������
    
    If InStr("," & strALL��ĿIDs & ",", "," & rsData!�ɼ���ĿID & ",") = 0 Then
        strALL��ĿIDs = strALL��ĿIDs & "," & rsData!�ɼ���ĿID
    End If
    
    If InStr("," & strALL��ĿIDs & ",", "," & rsData!������ĿID & ",") = 0 Then
        strALL��ĿIDs = strALL��ĿIDs & "," & rsData!������ĿID
    End If
      
End Sub

Public Function IsLis������Ŀ(ByVal lngMod As Long, ByVal lng��Ŀid As Long, ByRef strErr As String) As Boolean
'���ܣ���鵱ǰ������Ŀ�Ƿ���������Ŀ
'������lngMod ģ��ţ�lng��ĿID ������Ŀid
'      strErr ���Σ�������Ϣ
'˵����Ҫ����LIS�����нӿڲ����ڵ����
    Dim blnTmp As Boolean
 
    strErr = ""
    Call InitObjLis(lngMod)
    If Not gobjLIS Is Nothing Then
        On Error Resume Next
        blnTmp = gobjLIS.IsToleranceItem(lng��Ŀid, strErr)
        If Not blnTmp And strErr <> "" Then
            blnTmp = False
        End If
        If 438 = err.Number Then
            blnTmp = False
        End If
    End If
    IsLis������Ŀ = blnTmp
End Function

Public Function GetAdviceFeeKind(lngAdviceID As Long) As Byte
'���ܣ�����ҽ��ID��ȡ�������͵ķ��õ��ݵ����ʣ�1=������ã�2=סԺ����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    GetAdviceFeeKind = 2
    strSQL = "Select ��¼����,������� From ����ҽ������ Where ҽ��ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceFeeKind", lngAdviceID)
    If rsTmp.RecordCount > 0 Then
        If rsTmp!��¼���� = 1 Or rsTmp!��¼���� = 2 And Val("" & rsTmp!�������) = 1 Then
            GetAdviceFeeKind = 1
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
