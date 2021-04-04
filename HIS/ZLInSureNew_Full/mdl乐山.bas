Attribute VB_Name = "mdl��ɽ"
Option Explicit
Public Declare Sub LS_ErrMessage Lib "SIHisInterface.dll" Alias "GetErrorMessage" (ErrorMsg As TStringOfChar)
Public Declare Function LS_UserLogin Lib "SIHisInterface.dll" Alias "UserLogin" (UserCode As TStringOfChar, PWD As TStringOfChar) As Byte
Public Declare Function LS_ChangePwd Lib "SIHisInterface.dll" Alias "ChangeUserPwd" (OldPwd As TStringOfChar, NewPWD As TStringOfChar) As Byte
Public Declare Sub LS_UserLogout Lib "SIHisInterface.dll" Alias "UserLogout" ()
Public Declare Function LS_ConnectServer Lib "SIHisInterface.dll" Alias "ConnectServer" (ServerName As TStringOfChar) As Byte
Public Declare Sub LS_DisConnectServer Lib "SIHisInterface.dll" Alias "DisConnectServer" ()

'��ȡ�α�����Ϣ
Public Declare Function LS_GetPersonInfo Lib "SIHisInterface.dll" Alias "GetPersonInfo" (PInfo As �����Ϣ) As Byte
'��Ժ�Ǽ�
Public Declare Function LS_InHospitalRegister Lib "SIHisInterface.dll" Alias "InBedRegster" (InBedRegInfo As סԺ��Ϣ) As Byte
'��ȡ��Ժ�Ǽ���Ϣ
Public Declare Function LS_GetInHospitalRegInfo Lib "SIHisInterface.dll" Alias "GetInBedRegInfo" (InBedRegID As TStringOfChar) As Byte
'¼��ҩƷ����
Public Declare Function LS_AddDrug Lib "SIHisInterface.dll" Alias "AddDrug" (DrugInfo As ҩƷ��Ϣ) As Byte
'¼�����Ʒ���
Public Declare Function LS_AddDiag Lib "SIHisInterface.dll" Alias "AddDiag" (DiagInfo As ������Ϣ) As Byte
'¼�������ʩ����
Public Declare Function LS_AddService Lib "SIHisInterface.dll" Alias "AddServiceItem" (ServiceItemInfo As ������ʩ��Ϣ) As Byte
'���������ϸ
Public Declare Function LS_SaveDetail Lib "SIHisInterface.dll" Alias "InBedRegApplyUpdates" (InBedRegID As TStringOfChar) As Byte
'סԺ����Ԥ����
Public Declare Function LS_PreBalance Lib "SIHisInterface.dll" Alias "NewInBedBill" (InBedBillInfo As סԺ������Ϣ) As Byte
'סԺ���ý���
Public Declare Function LS_Balance Lib "SIHisInterface.dll" Alias "SaveInBedBill" (InBedBillInfo As סԺ������Ϣ) As Byte

'----��������ҵ��----
Public Declare Function LS_ExamBill Lib "SIHisInterface.dll" Alias "NewExamBill" (TexamBillInfo As ������㵥) As Byte
'¼������ҩƷ����
Public Declare Function LS_ExamAddDrug Lib "SIHisInterface.dll" Alias "AddExamDrug" (DrugInfo As ҩƷ��Ϣ) As Byte
'¼���������Ʒ���
Public Declare Function LS_ExamAddDiag Lib "SIHisInterface.dll" Alias "AddExamDiag" (DiagInfo As ������Ϣ) As Byte
'¼�����������ʩ����
Public Declare Function LS_ExamAddServiceItem Lib "SIHisInterface.dll" Alias "AddExamServiceItem" (ServiceItemInfo As ������ʩ��Ϣ) As Byte
'����Ԥ����
Public Declare Function LS_ExamPreBalance Lib "SIHisInterface.dll" Alias "ExamBillReCalculate" (TexamBillInfo As ������㵥) As Byte
'����Ԥ����
Public Declare Function LS_ExamBalance Lib "SIHisInterface.dll" Alias "SaveExamBill" (TexamBillInfo As ������㵥) As Byte
'--------------------

'ȫ�ֱ�����
Private Const mstr��Ժ���� As String = "��Ժ����"
Private Const mstr��;�ݽ��� As String = "��;�ݽ���"
Private Const mstrתԺ���� As String = "תԺ����"

'���������Ϣ����
Private Const ��Ժ���ұ�� = 0
Private Const ��Ժ�������� = 1
Private Const ��Ժ������� = 2
Private Const ��Ժ�������� = 3
Private Const ��Ժ������� = 4
Private Const ��Ժ�������� = 5
Private Const סԺҽʦ = 6
Private Const סԺ�� = 7
Private Const ��Ժ��� = 8
Private Const ��Ժ��� = 9
Private Const ��Ժ���� = 10
Private Const ��Ժ��ʽ = 11

Public Type TStringOfChar
    Data As String * 100
End Type
Public Type �����Ϣ                   'TPersonInfo
    '��������Ϊ��������
    PSN_ID              As Long      'ҽ�Ʋα�ID��
    PSN_No              As Long      '�α��˱���
    PSN_NAME            As String * 100 '�α�������
    Sex                 As String * 100 '�Ա�
    IDCARD              As String * 100 '���֤����
    PSN_STS             As String * 100 '�α���״̬
    PSN_TYP             As String * 100 '��Ա���
    UNIT_CODE           As String * 100 '��λ����
    UNIT_NAME           As String * 100 '��λ����
    OFFICAL_TYP         As String * 100 '����Ա���
    HAI_TYP             As String * 100 '����ҽ������
    ACCT_STS            As String * 100 'ҽ���˻�״̬
    HI_ACCT_PWD         As String * 100 'ҽ���ʻ�����
    SILL_PAY_AMT_TOTAL  As Single       '���ڽ����������⼲��֧�����
    SILL_YR_FUND_AMT    As Single       '��������ͳ�����֧�����
    YR_FUND_AMT         As Single       '����ͳ�����֧�����
    HAI_YR_HIGH_AMT     As Single       '���ڲ���߶�֧�����
    HAI_YR_INBED_AMT    As Single       '���ڲ���סԺ����֧�����
    GZ_CUR_AMT          As Single       '�����˻����
    YR_INBED_CNT        As Long      '����סԺ����
    CARD_NO             As String * 100 'IC����
End Type
Private Type סԺ��Ϣ                   'TInBedRegInfo
    PSN_ID              As Long      'ҽ�Ʋα���ID��
    INBED_SILL_ID       As Long      'סԺ���ⲡ��ID��������
    INBED_NO            As String * 100 'סԺ��
    INBED_EXAM          As String * 100 '��Ժ���
    INBED_EXAM_ICD10_NO As String * 100 '��Ժ���ICD10����
    INBED_DEPT          As String * 100 '��Ժ����
    '��������Ϊ��������
    INBED_REG_ID        As String * 100 'סԺ�Ǽ�ID
    INBED_DT            As String * 100 '��Ժʱ�䣬¼������
End Type
Private Type ҩƷ��Ϣ               'TDrugInfo
    INBED_REG_ID    As String * 100 'סԺ�Ǽ�ID
    RECEIPT_DT      As String * 100 '�շ�ʱ��
    DRUG_CATALOG_ID As String * 100 'ҩƷ�������ID
    DRUG_INFO       As String * 100 'ҩƷ��Ϣ
    UNIT_PRC        As Single       '����
    SRVC_CNT        As Single       '����
    COST_PRC        As Single       '�ɱ�����
    DRUG_TYP        As String * 100 'ҩ�����
    DRUG_SPEC       As String * 100 'ҩ����
    PRODUCE_FACTORY As String * 100 '��������
    '��������Ϊ��������
    FEE_ITEM_TYP    As String * 100 '������Ŀ����
    FEE_TYP         As String * 100 '��������
    PART_PUB_AMT    As Single       '���ֹ��ѽ��
    PART_SELF_AMT   As Single       '�����Էѽ��
    PUB_PAY_AMT     As Single       '���ѽ��
    SELF_PAY_AMT    As Single       '�Էѽ��
    SELF_PAY_PCT    As Single       '�Էѱ���
    MAX_RETAIL_PRC  As Single       '������ۼ�
    DRUG_SPC_FLAG   As Long         '������ҩ��־
End Type
Private Type ������Ϣ               'TDiagInfo
    INBED_REG_ID    As String * 100 'סԺ�Ǽ�ID
    RECEIPT_DT      As String * 100 '�շ�ʱ��
    DIAG_CATALOG_ID As String * 100 '������Ŀ�������ID
    DIAG_ITEM_NAME  As String * 100 '������Ŀ����
    UNIT_PRC        As Single       '����
    SRVC_CNT        As Single       '����
    '��������Ϊ��������
    FEE_ITEM_TYP    As String * 100 '������Ŀ����
    FEE_TYP         As String * 100 '��������
    PART_PUB_AMT    As Single       '���ֹ��ѽ��
    PART_SELF_AMT   As Single       '�����Էѽ��
    PUB_PAY_AMT     As Single       '���ѽ��
    SELF_PAY_AMT    As Single       '�Էѽ��
    SELF_PAY_PCT    As Single       '�Էѱ���
    MAX_RETAIL_PRC  As Single       '������ۼ�
End Type
Private Type ������ʩ��Ϣ           'TServiceItemInfo
    INBED_REG_ID    As String * 100 'סԺ�Ǽ�ID
    RECEIPT_DT      As String * 100 '�շ�ʱ��
    SRVC_ITEM_ID    As String * 100 '����ҽ�Ʊ��շ�����ʩ��׼
    SRVC_NAME       As String * 100 '������ʩ����
    UNIT_PRC        As Single       '����
    SRVC_CNT        As Single       '����
    '��������Ϊ��������
    FEE_ITEM_TYP    As String * 100 '������Ŀ����
    FEE_TYP         As String * 100 '��������
    PART_PUB_AMT    As Single       '���ֹ��ѽ��
    PART_SELF_AMT   As Single       '�����Էѽ��
    PUB_PAY_AMT     As Single       '���ѽ��
    SELF_PAY_AMT    As Single       '�Էѽ��
    SELF_PAY_PCT    As Single       '�Էѱ���
    MAX_RETAIL_PRC  As Single       '������ۼ�
End Type
Private Type סԺ������Ϣ                   'TInBedBillInfo
    INBED_REG_ID        As String * 100     'סԺ�Ǽ�ID
    EXAM_TYP            As String * 100     '�������
    INBED_STL_TYP       As String * 100     'סԺ���ʷ�ʽ
    OUTBED_EXAM         As String * 100     '��Ժ���
    OUTBED_EXAM_ICD10_NO As String * 100    '��Ժ���ICD10����
    OUTBED_DEPT         As String * 100     '��Ժ����
    ILL_TRS_STS         As String * 100     '����ת��(������������)
    INBED_DOCTOR        As String * 100     '�ܴ�ҽ��
    OUTBED_DT           As String * 100     '��Ժʱ��
    '��������Ϊ��������
    INBED_DAY_CNT       As Long          'סԺ����
    FEE_STL_LOC         As String * 100     '���ý���ص�
    EXAM_ADDR           As String * 100     '����ص�
    INBED_STL_BILL_ID   As String * 100     'סԺ���ʵ�id
    INBED_STL_BILL_NO   As String * 100     'סԺ���ʵ���
    PART_PUB_AMT        As Single           '���ֹ��ѽ��
    PART_SELF_AMT       As Single           '�����Էѽ��
    PUB_PAY_AMT         As Single           '���ѽ��
    SELF_PAY_AMT        As Single           '�Էѽ��
    INBED_FUND_AMT      As Single           'סԺͳ��֧�����
    INBED_ACCT_AMT      As Single           'סԺ����֧�����
    CASH_PAY_AMT        As Single           '�ֽ�֧�����
    HAI_INBED_SBS_AMT   As Single           '����סԺ����֧�����
    HAI_INBED_AMT       As Single           '����סԺ֧�����
    HAI_INBED_REPAY_AMT As Single           '����סԺ�ٴ�֧�����
    HAI_INBED_HIGH_AMT  As Single           '����סԺ�߶�֧�����
    OFFICAL_HIGH_AMT    As Single           '����Ա�߶��֧�����
    OFFICAL_INBED_AMT   As Single           '����ԱסԺ����֧�����
    OFFICAL_ACCT_AMT    As Single           '����Ա���ʲ���֧�����
End Type
'----��������ҵ��----
Private Type ������㵥
    PSN_ID           As Long                'ҽ�Ʋα���ID��
    EXAM_TYP         As String * 100        '�������
    EXAM_DEPT        As String * 100        '�������
    EXAM_DOCTOR      As String * 100        '����ҽ��
    '��������Ϊ��������
    FEE_STL_LOC      As String * 100        '���ý���ص�
    EXAM_ADDR        As String * 100        '����ص�
    EXAM_STL_BILL_ID As String * 100        '������ʵ�id
    EXAM_STL_BILL_NO As String * 100        '������ʵ���
    PART_PUB_AMT     As Single              '���ֹ��ѽ��
    PART_SELF_AMT    As Single              '�����Էѽ��
    PUB_PAY_AMT      As Single              '���ѽ��
    SELF_PAY_AMT     As Single              '�Էѽ��
    EXAM_ACCT_AMT    As Single              '�������֧�����
    CASH_PAY_AMT     As Single              '�����ֽ�֧�����
End Type
'--------------------

Private Type ������Ϣ
    ˳��� As TStringOfChar
    �ܷ��� As Currency
    �ֽ� As Currency
    �����ʻ� As Currency
    ҽ������ As Currency
    ������� As Currency
    ����Ա���� As Currency
    ҽ���ܷ��� As Currency
End Type
Public gPersonInfo_��ɽ As �����Ϣ
Public gInBedRegInfo_��ɽ As סԺ��Ϣ
Public gDrugInfo_��ɽ As ҩƷ��Ϣ
Public gDiagInfo_��ɽ As ������Ϣ
Public gServiceItemInfo_��ɽ As ������ʩ��Ϣ
Public gInBedBillInfo_��ɽ As סԺ������Ϣ
Private gtypBalance As ������Ϣ
Private gExamBill As ������㵥

Private glngInterface_��ɽ As Long
Private gstrErrMsg_��ɽ As TStringOfChar          '������Ϣ
Public gbytReturn_��ɽ As Byte                '0-����;����ֵ��������

Public Function ҽ����ʼ��_��ɽ() As Boolean
    Dim strServer As TStringOfChar
    On Error GoTo errHand
    
    If glngInterface_��ɽ <> 0 Then ҽ����ʼ��_��ɽ = True: Exit Function
    strServer = GetServerInfo
    If strServer.Data = "" Then Exit Function
    
    '���ӷ�����
    gbytReturn_��ɽ = LS_ConnectServer(strServer)
    If GetErrInfo_��ɽ Then Exit Function
    
    '��¼����(ʧ����Ͽ����Ӳ��˳�)
    If Not frm��¼����.LoginCenter(TYPE_��ɽ, True) Then
        Call ҽ����ֹ_��ɽ
        Exit Function
    End If
    glngInterface_��ɽ = 1
    
    ҽ����ʼ��_��ɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
  Resume
    End If
End Function

Public Function ҽ����ֹ_��ɽ() As Boolean
    On Error Resume Next
    If glngInterface_��ɽ = 0 Then
        ҽ����ֹ_��ɽ = True
        Exit Function
    End If

    '����Ա�˳�
    Call LS_UserLogout
    '���ӷ�����
    Call LS_DisConnectServer
    glngInterface_��ɽ = 0
    
    ҽ����ֹ_��ɽ = True
End Function

Public Function ҽ������_��ɽ() As Boolean
    With frmSet��ɽ
        ҽ������_��ɽ = .ShowME
    End With
End Function

Public Function GetErrInfo_��ɽ() As Boolean
    If gbytReturn_��ɽ = 1 Then Exit Function
    Call LS_ErrMessage(gstrErrMsg_��ɽ)
    MsgBox gstrErrMsg_��ɽ.Data, vbInformation, gstrSysName
    GetErrInfo_��ɽ = True
End Function

Private Function GetServerInfo() As TStringOfChar
    '��ȡ��������ַ
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡ��������ַ���˿ڼ��������('��������ַ','�������˿ں�','��������ڳ���')
    gstrSQL = " Select ������,����ֵ From ���ղ���" & _
              " Where ����=[1] And ������ = '��������ַ'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������ƻ�IP��ַ", TYPE_��ɽ)
    
    With rsTemp
        If .RecordCount = 0 Then Exit Function
        GetServerInfo.Data = Nvl(!����ֵ)
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������_��ɽ(ByVal strҽ���� As String) As Currency
    '����: ֱ�Ӷ������ڽ��
    '����: �Ƿ����
    '����: ���ظ����ʻ����
    Dim rsAccount As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select Nvl(�ʻ����,0) �ʻ���� From �����ʻ� " & _
              " Where ����=[1] And ҽ����=[2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "���ظ����ʻ����", TYPE_��ɽ, strҽ����)
    
    �������_��ɽ = rsAccount!�ʻ����
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �����������_��ɽ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    'cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��,������
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    On Error GoTo errHand
    Dim intType As Integer
    Dim str����ʱ�� As String
    Dim rs���� As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim dbl���� As Double
    
    If Nvl(rs��ϸ!������) = "" Then
        MsgBox "����ҽ������Ϊ�գ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '����δ���ݿ�������,ֻ�и��ݿ�����,��ȡ��������(ֻ��ȡ�ٴ�����)
    Call DebugTool("��ȡ����ҽ�����ڿ���")
    gstrSQL = "SELECT C.���� AS �������� " & _
             " FROM ������Ա A,��������˵�� B,���ű� C " & _
             " WHERE A.��ԱID= " & _
             "     (SELECT ID FROM ��Ա�� WHERE ����=[1]) " & _
             " AND A.����ID=B.����ID AND A.����ID=C.ID AND B.��������='�ٴ�' AND ������� IN (1,3) " & _
             " AND ROWNUM<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", CStr(rs��ϸ!������))
    If rsTemp.EOF Then
        MsgBox "��ҽ���������κ��ٴ����ң�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����µ�������㵥
    Call DebugTool("�����µ�������㵥")
    With gExamBill
        .PSN_ID = gPersonInfo_��ɽ.PSN_ID       'ҽ�Ʋα���ID
        .EXAM_TYP = ""                          '�������
        .EXAM_DEPT = rsTemp!��������            '�������
        .EXAM_DOCTOR = Nvl(rs��ϸ!������)       '����ҽ��
    End With
    gbytReturn_��ɽ = LS_ExamBill(gExamBill)
    If GetErrInfo_��ɽ Then Exit Function
    
    Call DebugTool("��ȡ����ʱ�估���˻�����Ϣ")
    str����ʱ�� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    Call ��ȡ���˻�����Ϣ(rs��ϸ!����ID)
    
    '��ȡ���մ���
    Call DebugTool("��ȡ���մ���")
    gstrSQL = "Select ID,���� From ����֧������ Where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ������", TYPE_��ɽ)
    
    '�ϴ���ϸ
    Call DebugTool("׼����ϸ�ϴ�")
    gtypBalance.�ܷ��� = 0
    With rs��ϸ
        Do While Not .EOF
            gtypBalance.�ܷ��� = gtypBalance.�ܷ��� + Nvl(!ʵ�ս��, 0)
            If rs��ϸ!ʵ�ս�� = 0 Then
                .MoveNext
                If .EOF Then
                    Exit Do
                End If
            End If
            '�ϴ���ϸ
            intType = 1
            rs����.Filter = "ID=" & !����֧������ID
            If rs����.RecordCount <> 0 Then
                If rs����!���� = "����" Then intType = 2
                If rs����!���� = "����" Then intType = 3
            End If
            rs����.Filter = 0
            
            Select Case intType
            Case 1
                Call DebugTool("ȡҩƷ��Ϣ")
                gstrSQL = "Select A.���,A.����,B.���� AS ����,D.��Ŀ���� AS ҽ����Ŀ����,E.���� AS ϸĿ����" & _
                         " From ҩƷĿ¼ A,ҩƷ���� B,ҩƷ��Ϣ C,����֧����Ŀ D,�շ�ϸĿ E " & _
                         " Where A.ҩ��ID=C.ҩ��ID And C.����=B.���� And A.ҩƷID=" & !�շ�ϸĿID & _
                         " And D.����=[1] And E.ID=D.�շ�ϸĿID And D.�շ�ϸĿID=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��Ϣ", TYPE_��ɽ, CLng(!�շ�ϸĿID))
                
                With gDrugInfo_��ɽ
                    .INBED_REG_ID = ""
                    .RECEIPT_DT = str����ʱ��
                    .DRUG_CATALOG_ID = rsTemp!ҽ����Ŀ����
                    .DRUG_INFO = rsTemp!ϸĿ����
                    dbl���� = rs��ϸ!ʵ�ս�� / rs��ϸ!����
                    '������:2005-07-06 ������۾��ȳ���2λС������������1�����۴�ʵ�ս�
                    If Round(dbl���� * 100) <> dbl���� * 100 Then
                        '������:2005-06-02�޸ģ����ȡ����ֵ,���ⵥ�۳��ָ��������
                        .UNIT_PRC = Format(Abs(rs��ϸ!ʵ�ս��), "#####0.00;-#####0.00;0;")
                        If rs��ϸ!ʵ�ս�� <= 0 Then
                          .SRVC_CNT = -1
                        Else
                          .SRVC_CNT = 1
                        End If
                    Else
                        .UNIT_PRC = Format(rs��ϸ!ʵ�ս�� / rs��ϸ!����, "#####0.00;-#####0.00;0;")
                        .SRVC_CNT = rs��ϸ!����
                    End If
                    .COST_PRC = 0
                    .DRUG_TYP = Nvl(rsTemp!����)
                    .DRUG_SPEC = Nvl(rsTemp!���)
                    .PRODUCE_FACTORY = Nvl(rsTemp!����)
                    .DRUG_SPC_FLAG = 0
                End With
            Case 2
                Call DebugTool("ȡ������Ϣ")
                gstrSQL = "Select D.��Ŀ���� AS ҽ����Ŀ����,E.���� AS ϸĿ����" & _
                         " From ����֧����Ŀ D,�շ�ϸĿ E " & _
                         " Where D.����=[1] And E.ID=D.�շ�ϸĿID And D.�շ�ϸĿID=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��Ϣ", TYPE_��ɽ, CLng(!�շ�ϸĿID))
                
                With gDiagInfo_��ɽ
                    .INBED_REG_ID = ""
                    .RECEIPT_DT = str����ʱ��
                    .DIAG_CATALOG_ID = rsTemp!ҽ����Ŀ����
                    .DIAG_ITEM_NAME = rsTemp!ϸĿ����
                    dbl���� = rs��ϸ!ʵ�ս�� / rs��ϸ!����
                    '������:2006-11-06 ������۾��ȳ���2λС������������1�����۴�ʵ�ս�
                    If Round(dbl���� * 100) <> dbl���� * 100 Then
                        '������:2006-11-06�޸ģ����ȡ����ֵ,���ⵥ�۳��ָ��������
                        .UNIT_PRC = Format(Abs(rs��ϸ!ʵ�ս��), "#####0.00;-#####0.00;0;")
                        If rs��ϸ!ʵ�ս�� <= 0 Then
                          .SRVC_CNT = -1
                        Else
                          .SRVC_CNT = 1
                        End If
                    Else
                        .UNIT_PRC = Format(rs��ϸ!ʵ�ս�� / rs��ϸ!����, "#####0.00;-#####0.00;0;")
                        .SRVC_CNT = rs��ϸ!����
                    End If
                End With
            Case 3
                Call DebugTool("ȡ������ʩ��Ϣ")
                gstrSQL = "Select D.��Ŀ���� AS ҽ����Ŀ����,E.���� AS ϸĿ����" & _
                         " From ����֧����Ŀ D,�շ�ϸĿ E " & _
                         " Where D.����=[1] And E.ID=D.�շ�ϸĿID And D.�շ�ϸĿID=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��Ϣ", TYPE_��ɽ, CLng(!�շ�ϸĿID))
                
                With gServiceItemInfo_��ɽ
                    .INBED_REG_ID = ""
                    .RECEIPT_DT = str����ʱ��
                    .SRVC_ITEM_ID = rsTemp!ҽ����Ŀ����
                    .SRVC_NAME = rsTemp!ϸĿ����
                    dbl���� = rs��ϸ!ʵ�ս�� / rs��ϸ!����
                    '������:2006-11-06 ������۾��ȳ���2λС������������1�����۴�ʵ�ս�
                    If Round(dbl���� * 100) <> dbl���� * 100 Then
                        '������:2006-11-06�޸ģ����ȡ����ֵ,���ⵥ�۳��ָ��������
                        .UNIT_PRC = Format(Abs(rs��ϸ!ʵ�ս��), "#####0.00;-#####0.00;0;")
                        If rs��ϸ!ʵ�ս�� <= 0 Then
                          .SRVC_CNT = -1
                        Else
                          .SRVC_CNT = 1
                        End If
                    Else
                        .UNIT_PRC = Format(rs��ϸ!ʵ�ս�� / rs��ϸ!����, "#####0.00;-#####0.00;0;")
                        .SRVC_CNT = rs��ϸ!����
                    End If
                End With
            End Select
            
            Call DebugTool("�ϴ���ϸ")
            If Not UploadDetail(intType, False) Then Exit Function
            .MoveNext
        Loop
    End With
    
    'Ԥ����
    '�����ʻ�֧����:EXAM_ACCT_AMT,�������޸�,�����֧�ָ����ʻ�,��֧��ҽ������
    Call DebugTool("����Ԥ����")
    gbytReturn_��ɽ = LS_ExamPreBalance(gExamBill)
    If GetErrInfo_��ɽ Then Exit Function
    
    '�������ݸ�ֵ
    Call DebugTool("��ȡ�������֧����")
    With gtypBalance
        .�����ʻ� = gExamBill.EXAM_ACCT_AMT
        .������� = 0
        .����Ա���� = 0
        .ҽ������ = 0
        .�ֽ� = .�ܷ��� - .�����ʻ�
    End With
    
    str���㷽ʽ = "�����ʻ�;" & gtypBalance.�����ʻ� & ";0"
    �����������_��ɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������_��ɽ(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    'cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ���������ڽ��㣬����ʹ��Ȧ��ӿڣ����������Ǯ���������ڣ������ӿ��ڽ��
    '���������Ҫͨ��������������ȡ����Ȧ�����ǽӿڷ��أ���Ҫ�޸�
    On Error GoTo errHand
    Dim lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim datCurr As Date
    
    '�������
    Call DebugTool("��ʽ����")
    gbytReturn_��ɽ = LS_ExamBalance(gExamBill)
    If GetErrInfo_��ɽ Then Exit Function
    
    Call DebugTool("��ȡ����ID")
    gstrSQL = "Select ����ID,�Ǽ�ʱ�� From ������ü�¼ Where ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID)
    lng����ID = rsTemp!����ID
    datCurr = rsTemp!�Ǽ�ʱ��
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim cur�����ۼ� As Currency, cur�������� As Currency, curͳ���޶� As Currency
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��ɽ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, curͳ���޶�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_��ɽ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & curͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    Call DebugTool("���汣�ս����¼")
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��ɽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gtypBalance.�ܷ��� & "," & gtypBalance.�ֽ� & "," & 0 & "," & gtypBalance.ҽ������ & "," & gtypBalance.ҽ������ & "," & _
        gtypBalance.������� & "," & gtypBalance.����Ա���� & "," & gtypBalance.�����ʻ� & ",'" & TrimTsChar(gtypBalance.˳���.Data) & "',null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������������")
    
    �������_��ɽ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_��ɽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    'cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    On Error GoTo errHand
    
    Err.Raise 9000, gstrSysName, "��ɽҽ����֧�������˷ѣ�����ҽ��������ϵ!"
    ����������_��ɽ = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_��ɽ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim str˳��� As String, strID As String
    Dim arrPatient
    
    On Error GoTo errHand
    
    arrPatient = Split(��ȡ���������Ϣ(lng����ID, lng��ҳID), "||")
    'д�������
    With gInBedRegInfo_��ɽ
        .PSN_ID = gPersonInfo_��ɽ.PSN_ID                           'סԺ�α�ID��
        .INBED_SILL_ID = 0                                          'סԺ���ⲡ��ID��������
        .INBED_NO = arrPatient(סԺ��)                              'סԺ��
        .INBED_EXAM = Split(arrPatient(��Ժ���), "|")(0)           '��Ժ���
        .INBED_EXAM_ICD10_NO = Split(arrPatient(��Ժ���), "|")(1)  '��Ժ���ICD10����
        .INBED_DEPT = arrPatient(��Ժ��������)                          '��Ժ����
    End With
    
    '������Ժ�Ǽǽӿ�
    gbytReturn_��ɽ = LS_InHospitalRegister(gInBedRegInfo_��ɽ)
    If GetErrInfo_��ɽ Then Exit Function
    
    '���¸����ʻ��е���Ϣ
    str˳��� = TrimTsChar(gInBedRegInfo_��ɽ.INBED_REG_ID)
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��ɽ & ",'˳���','''" & str˳��� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժҵ�����к�")
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��ɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")

    ��Ժ�Ǽ�_��ɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_��ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ��
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    '����������Ժ
    On Error GoTo errHand
    
    MsgBox "��֧�ֳ�Ժ�Ǽǳ���������ҽ���ӿ�����ϵ��", vbInformation, gstrSysName
    ��Ժ�Ǽǳ���_��ɽ = False
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_��ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHand
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ��
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false

    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��ɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")

    ��Ժ�Ǽ�_��ɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽǳ���_��ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rs����          As ADODB.Recordset
    On Error GoTo errHand
'    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��ɽ & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    MsgBox "���ڸõ�����֧����Ժ�Ǽǳ��������β���ֻ��Ա������ݳ�����ȡ�����ĵǼ��뵽ҽ�����ȡ������ϵҽ���ӿ��̣�", vbInformation, gstrSysName
    gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
    gstrSQL = "select id from סԺ���ü�¼ where ����id = [1] and ��ҳid = [2]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "������иô�סԺ���ü�¼", lng����ID, lng��ҳID)
    If Not rs����.EOF Then
        Do While Not rs����.EOF
        '��������ٴβ���ɱ��ղ���ʱ�������������ϴ�,�ô���Ϊ0
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rs����!ID & ",null,null,null,null" & ",0,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���²��˷��ü�¼�����м�¼�ϴ���־Ϊ0")
            rs����.MoveNext
        Loop
    End If
    ��Ժ�Ǽǳ���_��ɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function סԺ�������_��ɽ(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim lng��ҳID As Long
    Dim bln��Ժ���� As Boolean
    Dim str��¼���� As String, str��¼״̬ As String, strNO As String
    Dim arrPatient
    Dim rsTemp As New ADODB.Recordset
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    On Error GoTo errHand
    
    '��ȡ��ҳID
    gstrSQL = "Select סԺ���� ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    '��ȡ�ܷ���
    gtypBalance.�ܷ��� = 0
    With rsExse
        Do While Not .EOF
            gtypBalance.�ܷ��� = gtypBalance.�ܷ��� + Nvl(!���, 0)
            '�ϴ���ϸ
            If !��ҳID = lng��ҳID And Nvl(!�Ƿ��ϴ�, 0) = 0 And (strNO <> !NO Or str��¼���� <> !��¼���� Or str��¼״̬ <> !��¼״̬) Then
                strNO = !NO
                str��¼���� = !��¼����
                str��¼״̬ = !��¼״̬
                If Not �ϴ�����_��ɽ(str��¼����, str��¼״̬, strNO) Then Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    Call ��ȡ���˻�����Ϣ(lng����ID)
    arrPatient = Split(��ȡ���������Ϣ(lng����ID, lng��ҳID), "||")
    bln��Ժ���� = ҽ�������Ѿ���Ժ(lng����ID)
    
    'д�������
    '������(2006-05-15):��Ժ��ʽ��������������ת��δ������תԺ����������������ת��δ����Ϊ����ת�鷽ʽ�����תԺ��ʽΪ�������򼲲�ת��Ĭ��Ϊ��ת��
    With gInBedBillInfo_��ɽ
        .INBED_REG_ID = gtypBalance.˳���.Data
        .EXAM_TYP = ""
        .INBED_STL_TYP = IIf(bln��Ժ����, IIf(arrPatient(��Ժ��ʽ) = "תԺ", mstrתԺ����, mstr��Ժ����), mstr��;�ݽ���)
        .OUTBED_EXAM = Split(arrPatient(��Ժ���), "|")(0)
        .OUTBED_EXAM_ICD10_NO = Split(arrPatient(��Ժ���), "|")(1)
        .OUTBED_DEPT = arrPatient(��Ժ��������)
        .ILL_TRS_STS = IIf(bln��Ժ����, IIf(arrPatient(��Ժ��ʽ) = "����", "����", arrPatient(��Ժ��ʽ)), "δ��")
        .INBED_DOCTOR = arrPatient(סԺҽʦ)
        .OUTBED_DT = IIf(bln��Ժ����, arrPatient(��Ժ����), "")
    End With
    gbytReturn_��ɽ = LS_PreBalance(gInBedBillInfo_��ɽ)
    If GetErrInfo_��ɽ Then Exit Function

    Call Get������Ϣ
    
    '��ʾ�α����˵�סԺ�����Ϣ
    If Format(gtypBalance.�ܷ���, "#0.00") <> Format(gtypBalance.ҽ���ܷ���, "#0.00") Then
        MsgBox "�òα����˵�ҽ���ܷ��ã���" & Format(gtypBalance.ҽ���ܷ���, "#0.00") & "Ԫ     " & vbCrLf & _
               "      ���ڲ�ϵͳ�ܷ��ã���" & Format(gtypBalance.�ܷ���, "#0.00") & "Ԫ     " & vbCrLf & _
               " ��һ��.����ϵ����Ա�����ٽ���!", vbInformation, gstrSysName
    End If
           
    סԺ�������_��ɽ = "�����ʻ�;" & gtypBalance.�����ʻ� & ";0"
    סԺ�������_��ɽ = סԺ�������_��ɽ & "|ҽ������;" & gtypBalance.ҽ������ & ";0"
    סԺ�������_��ɽ = סԺ�������_��ɽ & "|�������;" & gtypBalance.������� & ";0"
    סԺ�������_��ɽ = סԺ�������_��ɽ & "|����Ա����;" & gtypBalance.����Ա���� & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_��ɽ(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim cur�ʻ�֧�� As Currency
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
  '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
    On Error GoTo errHand
    Call ��ȡ���˻�����Ϣ(lng����ID)
    
    '��ȡ���θ����ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=[2]" & _
        " And A.���㷽ʽ in ('�����ʻ�') And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���θ����ʻ�֧����", lng����ID, TYPE_��ɽ)
    cur�ʻ�֧�� = 0
    If Not rsTemp.EOF Then
        cur�ʻ�֧�� = rsTemp!�����ʻ�
    End If
    
    'ֱ�ӵ��ý���ӿڣ���Ϊ��������Ѿ���д����ڲ���
    gbytReturn_��ɽ = LS_Balance(gInBedBillInfo_��ɽ)
    If GetErrInfo_��ɽ Then Exit Function
    
    Call Get������Ϣ(cur�ʻ�֧��)
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʼ�¼�����ϴ���־")
    
    '��д���ս����¼
    '�����Ը����=�������;�����Ը����=����Ա����
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��ɽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gtypBalance.�ܷ��� & "," & gtypBalance.�ֽ� & "," & 0 & "," & gtypBalance.ҽ������ & "," & gtypBalance.ҽ������ & "," & _
        gtypBalance.������� & "," & gtypBalance.����Ա���� & "," & cur�ʻ�֧�� & ",'" & TrimTsChar(gtypBalance.˳���.Data) & "',null,null,'" & TrimTsChar(gInBedBillInfo_��ɽ.INBED_STL_BILL_NO) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")
    סԺ����_��ɽ = True
    
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_��ɽ(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    Err.Raise 9000, gstrSysName, "��֧��סԺ����������뵽ҽ�����İ��� "
    סԺ�������_��ɽ = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��ݱ�ʶ_��ɽ(Optional bytType As Byte, Optional lng����ID As Long) As String
'    ���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'    ������bytType-ʶ�����ͣ�0-���1-סԺ
'����:     �ջ���Ϣ��
'    ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'    2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'    3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    '��֧��סԺ
    ��ݱ�ʶ_��ɽ = frmIdentify��ɽ.GetPatient(bytType, lng����ID)
End Function

Private Function ��ȡ���������Ϣ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    Dim str��Ժ���ұ�� As String, str��Ժ�������� As String, str��Ժ������� As String
    Dim str��Ժ�������� As String, str��Ժ������� As String, str��Ժ�������� As String
    Dim strסԺҽʦ As String, strסԺ�� As String, str��Ժ��� As String
    Dim str��Ժ��� As String, str��Ժ���� As String, str��Ժ��ʽ As String
    Dim rsTemp As New ADODB.Recordset
'    ��ȡ���������Ϣ (����סԺ����||��Ժ���ұ��||��Ժ��������||��Ժ�������||��Ժ��������||��Ժ�������||סԺ��||��Ժ���||��Ժ���)
    
'    ��ȡ��Ժ�����Ϣ
    gstrSQL = "select C.���� ��Ժ���ұ��,C.���� ��Ժ��������,B.���� ��Ժ�������,B.���� ��Ժ��������, " & _
             " A.��Ժ���� ��Ժ�������,D.���� ��Ժ��������,F.��λ����,E.סԺ�� סԺ��,A.סԺҽʦ,to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,A.��Ժ��ʽ " & _
             " from ������ҳ A,���ű� B,���ű� C,���ű� D,������Ϣ E, " & _
             " (Select D.���� ��λ����,F.����,F.����ID,F.����ID  From ��λ�ȼ� D ,��λ״����¼ F Where F.�ȼ�ID=D.���) F " & _
             " Where A.��Ժ����ID=B.ID(+) And A.��Ժ����ID=C.ID(+) And A.��Ժ����ID=D.ID(+) And A.����ID=E.����ID ANd A.����ID=[1] And A.��ҳID=[2]" & _
             " And A.��Ժ����=F.����(+) And F.����ID(+)=A.��Ժ����ID And F.����ID(+)=A.��Ժ����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ�����Ϣ", lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
        str��Ժ���ұ�� = Nvl(rsTemp!��Ժ���ұ��)
        str��Ժ�������� = Nvl(rsTemp!��Ժ��������)
        str��Ժ������� = Nvl(rsTemp!��Ժ�������)
        str��Ժ�������� = Nvl(rsTemp!��Ժ��������)
        str��Ժ������� = Nvl(rsTemp!��Ժ�������)
        str��Ժ�������� = Nvl(rsTemp!��Ժ��������)
        strסԺҽʦ = Nvl(rsTemp!סԺҽʦ)
        str��Ժ���� = Nvl(rsTemp!��Ժ����)
        str��Ժ��ʽ = Nvl(rsTemp!��Ժ��ʽ)
        strסԺ�� = Nvl(rsTemp!סԺ��)
    End If
    
'    ��ȡ���Ժ��ϣ����|�������룩
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, False, True)
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, False, True)
    ��ȡ���������Ϣ = str��Ժ���ұ�� & "||" & str��Ժ�������� & "||" & _
                    str��Ժ������� & "||" & str��Ժ�������� & "||" & str��Ժ������� & "||" & _
                    str��Ժ�������� & "||" & strסԺҽʦ & "||" & strסԺ�� & "||" & str��Ժ��� & _
                    "||" & str��Ժ��� & "||" & str��Ժ���� & "||" & str��Ժ��ʽ
End Function

Private Sub ��ȡ���˻�����Ϣ(ByVal lng����ID As Long)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ˳��� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˵�סԺ��ˮ��", lng����ID, TYPE_��ɽ)
    
    gtypBalance.˳���.Data = Nvl(rsTemp!˳���)
End Sub

Private Function �Ƿ�ҽ������(ByVal lng����ID As Long) As Boolean
    Dim rsInsure As New ADODB.Recordset
    
    '��鱾���Ƿ���ҽ�������Ժ
    gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1] And A.��ҳID=B.סԺ���� And A.����=[2]"
    Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ҽ������", lng����ID, TYPE_��ɽ)
    �Ƿ�ҽ������ = (rsInsure!Records = 1)
End Function

Private Sub Get������Ϣ(Optional ByVal cur�ʻ�֧�� As Currency = 0)
    '����Ԥ�������㷵�ص�ֵ����ʾ������Ϣ�����ڸ����ʻ��ǽӿڷ��صģ����Ʋ������޸ģ�
    With gtypBalance
'        PART_PUB_AMT      : Single;         //���ֹ��ѽ��
'        PART_SELF_AMT     : Single;         //�����Էѽ��
'        PUB_PAY_AMT       : Single;         //���ѽ��
'        SELF_PAY_AMT      : Single;         //�Էѽ��

'        INBED_FUND_AMT      As Single           'סԺͳ��֧�����
'        INBED_ACCT_AMT      As Single           'סԺ����֧�����
'        CASH_PAY_AMT        As Single           '�ֽ�֧�����
'        HAI_INBED_SBS_AMT   As Single           '����סԺ����֧�����
'        HAI_INBED_AMT       As Single           '����סԺ֧�����
'        HAI_INBED_REPAY_AMT As Single           '����סԺ�ٴ�֧�����
'        HAI_INBED_HIGH_AMT  As Single           '����סԺ�߶�֧�����
'        OFFICAL_HIGH_AMT    As Single           '����Ա�߶��֧�����
'        OFFICAL_INBED_AMT   As Single           '����ԱסԺ����֧�����
'        OFFICAL_ACCT_AMT    As Single           '����Ա���ʲ���֧�����

        .�����ʻ� = IIf(cur�ʻ�֧�� = 0, gInBedBillInfo_��ɽ.INBED_ACCT_AMT, cur�ʻ�֧��)
        .������� = gInBedBillInfo_��ɽ.HAI_INBED_SBS_AMT + gInBedBillInfo_��ɽ.HAI_INBED_AMT + _
                    gInBedBillInfo_��ɽ.HAI_INBED_REPAY_AMT + gInBedBillInfo_��ɽ.HAI_INBED_HIGH_AMT
        .ҽ������ = gInBedBillInfo_��ɽ.INBED_FUND_AMT
        .����Ա���� = gInBedBillInfo_��ɽ.OFFICAL_HIGH_AMT + _
                      gInBedBillInfo_��ɽ.OFFICAL_INBED_AMT + gInBedBillInfo_��ɽ.OFFICAL_ACCT_AMT
        .ҽ���ܷ��� = gInBedBillInfo_��ɽ.PART_PUB_AMT + gInBedBillInfo_��ɽ.PART_SELF_AMT + _
                      gInBedBillInfo_��ɽ.PUB_PAY_AMT + gInBedBillInfo_��ɽ.SELF_PAY_AMT
        If cur�ʻ�֧�� <> 0 Then
            .�ֽ� = .�ܷ��� - .ҽ������ - .������� - .�����ʻ� - .����Ա����
        End If
    End With
End Sub

Public Function �ϴ�����_��ɽ(ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal str���ݺ� As String) As Boolean
    Dim intType As Integer
    Dim int������Ŀ As Integer          '�����ù���վ�������������������Ŀʱ,������Ŀ����Ϊ������Ŀ�ϴ�
    Dim int����Ϊ������Ŀ As Integer    '��������Ŀ,�����Ƿ�������Ϊ������Ŀ
    Dim lng����ID As Long
    Dim blnInsure As Boolean, blnUpload As Boolean, blnTrans As Boolean
    Dim dbl���� As Double
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim gcn�ϴ� As New ADODB.Connection
    On Error GoTo errHand
    
    Call DebugTool("��ȡ������ϸ")
    gstrSQL = " Select A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') �Ǽ�ʱ��,A.�շ����," & _
              " A.������ ҽ��,B.���� ��������,A.�շ�ϸĿID,D.���� ϸĿ����,C.��Ŀ���� ҽ����Ŀ����,C.ҽ������,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�,A.ժҪ" & _
              " From סԺ���ü�¼ A,���ű� B,�շ�ϸĿ D," & _
              "     (Select A.*,B.���� ҽ������ From ����֧����Ŀ A,����֧������ B " & _
              "     Where A.����=B.���� And A.����ID=B.ID And A.����=" & TYPE_��ɽ & ") C,�����ʻ� G " & _
              " Where A.��¼����=[1] And A.��¼״̬=[2] And A.NO=[3]" & _
              " And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And A.�շ�ϸĿID=D.ID And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0" & _
              " And A.����ID=G.����ID And G.����=[4]" & _
              " Order by A.NO,A.����ID"
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", int����, int״̬, str���ݺ�, TYPE_��ɽ)
    
    Set gcn�ϴ� = GetNewConnection
    With rsExse
        Do While Not .EOF
            If lng����ID <> !����ID Then
                '�ύ����
                If lng����ID <> 0 And blnInsure Then
                    Call DebugTool("���浥��")
                    gbytReturn_��ɽ = LS_SaveDetail(gtypBalance.˳���)
                    If GetErrInfo_��ɽ Then
                        gcn�ϴ�.RollbackTrans
                        �ϴ�����_��ɽ = True
                        Exit Function
                    End If
                    gcn�ϴ�.CommitTrans
                    blnTrans = False
                End If
            End If
            
            '�жϵ�ǰ�����Ƿ񱾴���ҽ����ݵǼ�
            If lng����ID <> !����ID Then
                Call DebugTool("�ж��Ƿ�Ϊҽ������")
                blnInsure = �Ƿ�ҽ������(!����ID)
            End If
            If blnInsure Then
                If lng����ID <> !����ID Then
                    Call DebugTool("��ȡ�ò��˻�����Ϣ��סԺ��Ϣ")
                    lng����ID = !����ID
                    Call ��ȡ���˻�����Ϣ(lng����ID)
                    gbytReturn_��ɽ = LS_GetInHospitalRegInfo(gtypBalance.˳���)
                    gcn�ϴ�.BeginTrans
                    blnTrans = True
                    If GetErrInfo_��ɽ Then
                        gcn�ϴ�.RollbackTrans
                        �ϴ�����_��ɽ = True
                        Exit Function
                    End If
                End If
                
                '�ϴ���ϸ
                intType = 1
                If !ҽ������ = "����" Then intType = 2
                If !ҽ������ = "����" Then intType = 3
                
'                Call DebugTool("�жϿɷ���Ϊ������Ŀ")
'                int����Ϊ������Ŀ = 0
'                gstrSQL = "Select ��ע From ����֧����Ŀ Where ����=" & TYPE_��ɽ & " And �շ�ϸĿID=" & !�շ�ϸĿID
'                Call OpenRecordset(rsTemp, "�ж�ҽ����Ŀ�Ƿ�������Ϊ������Ŀ")
'                If Not rsTemp.EOF Then
'                    If Not IsNull(rsTemp!��ע) Then
'                        If UBound(Split(rsTemp!��ע, "|")) >= 2 Then
'                            int����Ϊ������Ŀ = Val(Split(rsTemp!��ע, "|")(2))
'                        End If
'                    End If
'                End If
                
                '�жϱ�����ϸ�Ƿ���Ϊ������Ŀ����Ϊ¼�������ϸʱ�����Ѿ������ÿ����Ŀ�Ƿ�������Ϊ������Ŀ���˴�����Ҫ�ٴ��жϣ�ֻҪժҪΪ1��˵����Ϊ������Ŀ�ϴ���
                int������Ŀ = IIf(Nvl(rsExse!ժҪ) = "1", 1, 0)
                
                '������(2005-10-18) �ж��Ƿ�����ҽ�����룬���δ���ã��������ʾ
                If IsNull(rsExse!ҽ����Ŀ����) Then
                   MsgBox "��Ŀ:" & rsExse!ϸĿ���� & "[���ۣ�" & rsExse!��� / rsExse!���� & "Ԫ]δ����ҽ������!"
                   �ϴ�����_��ɽ = True
                   Exit Function
                End If
                
                Select Case intType
                Case 1
                    Call DebugTool("ȡҩƷ��Ϣ")
                    gstrSQL = "select A.���,A.����,B.���� ����  " & _
                             " from ҩƷĿ¼ A,ҩƷ���� B,ҩƷ��Ϣ C " & _
                             " Where A.ҩ��ID=C.ҩ��ID And C.����=B.���� And A.ҩƷID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��Ϣ", CLng(!�շ�ϸĿID))
                    
                    With gDrugInfo_��ɽ
                        .INBED_REG_ID = gtypBalance.˳���.Data
                        .RECEIPT_DT = Format(rsExse!�Ǽ�ʱ��, "yyyy-MM-dd")
                        .DRUG_CATALOG_ID = rsExse!ҽ����Ŀ����
                        .DRUG_INFO = rsExse!ϸĿ����
                        dbl���� = rsExse!��� / rsExse!����
                        '������:2005-07-06 ������۾��ȳ���2λС������������1�����۴�ʵ�ս�
                        If Round(dbl���� * 100) <> dbl���� * 100 Then
                            '������:2005-06-02�޸ģ����ȡ����ֵ,���ⵥ�۳��ָ��������
                            .UNIT_PRC = Format(Abs(rsExse!���), "#####0.00;-#####0.00;0;")
                            If rsExse!��� <= 0 Then
                              .SRVC_CNT = -1
                            Else
                              .SRVC_CNT = 1
                            End If
                        Else
                            .UNIT_PRC = Format(rsExse!��� / rsExse!����, "#####0.00;-#####0.00;0;")
                            .SRVC_CNT = rsExse!����
                        End If
                        .COST_PRC = 0
                        .DRUG_TYP = Nvl(rsTemp!����)
                        .DRUG_SPEC = Nvl(rsTemp!���)
                        .PRODUCE_FACTORY = Nvl(rsTemp!����)
                        .DRUG_SPC_FLAG = int������Ŀ
                    End With
                Case 2
                    Call DebugTool("ȡ������Ϣ")
                    With gDiagInfo_��ɽ
                        .INBED_REG_ID = gtypBalance.˳���.Data
                        .RECEIPT_DT = Format(rsExse!�Ǽ�ʱ��, "yyyy-MM-dd")
                        .DIAG_CATALOG_ID = rsExse!ҽ����Ŀ����
                        .DIAG_ITEM_NAME = rsExse!ϸĿ����
                        dbl���� = rsExse!��� / rsExse!����
                        '������:2006-11-06 ������۾��ȳ���2λС������������1�����۴�ʵ�ս�
                        If Round(dbl���� * 100) <> dbl���� * 100 Then
                           '������:2006-11-06�޸ģ����ȡ����ֵ,���ⵥ�۳��ָ��������
                            .UNIT_PRC = Format(Abs(rsExse!���), "#####0.00;-#####0.00;0;")
                            If rsExse!��� <= 0 Then
                              .SRVC_CNT = -1
                            Else
                              .SRVC_CNT = 1
                            End If
                        Else
                            .UNIT_PRC = Format(rsExse!��� / rsExse!����, "#####0.00;-#####0.00;0;")
                            .SRVC_CNT = rsExse!����
                        End If
                    End With
                Case 3
                    Call DebugTool("ȡ������ʩ��Ϣ")
                    With gServiceItemInfo_��ɽ
                        .INBED_REG_ID = gtypBalance.˳���.Data
                        .RECEIPT_DT = Format(rsExse!�Ǽ�ʱ��, "yyyy-MM-dd")
                        .SRVC_ITEM_ID = rsExse!ҽ����Ŀ����
                        .SRVC_NAME = rsExse!ϸĿ����
                        dbl���� = rsExse!��� / rsExse!����
                        '������:2006-11-06 ������۾��ȳ���2λС������������1�����۴�ʵ�ս�
                        If Round(dbl���� * 100) <> dbl���� * 100 Then
                           '������:2006-11-06�޸ģ����ȡ����ֵ,���ⵥ�۳��ָ��������
                            .UNIT_PRC = Format(Abs(rsExse!���), "#####0.00;-#####0.00;0;")
                            If rsExse!��� <= 0 Then
                              .SRVC_CNT = -1
                            Else
                              .SRVC_CNT = 1
                            End If
                        Else
                            .UNIT_PRC = Format(rsExse!��� / rsExse!����, "#####0.00;-#####0.00;0;")
                            .SRVC_CNT = rsExse!����
                        End If
                    End With
                End Select
                
                Call DebugTool("�ϴ���ϸ")
                If Not UploadDetail(intType) Then
                    gcn�ϴ�.RollbackTrans
                    �ϴ�����_��ɽ = True
                    Exit Function
                End If
                
                '���ϱ��
                Call DebugTool("���ϴ����")
                '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
                'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
                gcn�ϴ�.Execute gstrSQL, , adCmdStoredProc
                Call DebugTool("�ϴ����SQL��" & gstrSQL)
                
                blnUpload = True
            End If
            .MoveNext
        Loop
        
       
        Call DebugTool("�����ϴ���ϸ")
        If blnUpload And blnInsure Then
            gbytReturn_��ɽ = LS_SaveDetail(gtypBalance.˳���)
            If GetErrInfo_��ɽ Then
                gcn�ϴ�.RollbackTrans
                �ϴ�����_��ɽ = True
                Exit Function
            End If
            Call DebugTool("�ύ����")
            gcn�ϴ�.CommitTrans
            blnTrans = False
        End If
    End With
    
    �ϴ�����_��ɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    �ϴ�����_��ɽ = True
    If blnTrans Then gcn�ϴ�.RollbackTrans
End Function

Private Function UploadDetail(Optional ByVal intType As Integer = 1, Optional ByVal blnסԺ As Boolean = True) As Boolean
    '�ϴ�������ϸ
    'intType:1-ҩƷ;2-����;3-����
    If blnסԺ Then
        Select Case intType
        Case 1
            gbytReturn_��ɽ = LS_AddDrug(gDrugInfo_��ɽ)
        Case 2
            gbytReturn_��ɽ = LS_AddDiag(gDiagInfo_��ɽ)
        Case 3
            gbytReturn_��ɽ = LS_AddService(gServiceItemInfo_��ɽ)
        End Select
    Else
        Select Case intType
        Case 1
            gbytReturn_��ɽ = LS_ExamAddDrug(gDrugInfo_��ɽ)
        Case 2
            gbytReturn_��ɽ = LS_ExamAddDiag(gDiagInfo_��ɽ)
        Case 3
            gbytReturn_��ɽ = LS_ExamAddServiceItem(gServiceItemInfo_��ɽ)
        End Select
    End If
    If GetErrInfo_��ɽ Then Exit Function
    UploadDetail = True
End Function

Private Function TrimTsChar(ByVal strData As Variant) As String
    TrimTsChar = Replace(Replace(strData, " ", ""), Chr(0), "")
End Function
