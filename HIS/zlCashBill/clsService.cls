VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'һ�������ӿ�
'   1.zlInitCommon-��ʼ���ӿ�
'������������ط���ӿ�
 
'   2.
'�ġ��ڲ��ӿڼ�����
'   1.GetServiceCall-��ȡ�������
'   2.GetCardTypeRecStru-������ȡ�Ŀ������Ϣ��
'����:���˺�
'����:2019-11-25 11:11:39
'*********************************************************************************************************************************************
Private mlngErrNum As Long, mstrSource As String, mstrErrMsg As String
Private mobjServiceCall As Object
Private mlngModule As Long
Private mrsCardType As ADODB.Recordset  '�����

Public Function zlInitCommon(ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngModule-ģ���
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngModule = lngModule
    If GetServiceCall(mobjServiceCall) = False Then Set mobjServiceCall = Nothing: Exit Function
    zlInitCommon = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetServiceCall(ByRef objServiceCall_Out As Object, Optional blnShowErrMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������
    '����:objServiceCall_Out-���ع����������
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    If Not mobjServiceCall Is Nothing Then Set objServiceCall_Out = mobjServiceCall: GetServiceCall = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If Err <> 0 Then
        mstrErrMsg = "������zlServiceCall����ʧ������ϵͳ����Ա��ϵ���ָ��ò�����"
        If blnShowErrMsg Then
            MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
            Err = 0: On Error GoTo 0
        Else
            Err.Raise Err.Number, Err.Source, mstrErrMsg: Exit Function
        End If
        Exit Function
    End If
    
    On Error GoTo errHandle
    If mobjServiceCall.InitService(gcnOracle, gstrDBUser, glngSys, glngModul) = False Then
        
        Set mobjServiceCall = Nothing: Exit Function
    End If
    Set objServiceCall_Out = mobjServiceCall
    GetServiceCall = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zl_PatiSvr_GetCardTypes(ByRef rsCardType_Out As ADODB.Recordset, _
    Optional bln�������� As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ�����
    '���:
    '����:rsCardType_Out-���صĿ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 16:53:48
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    
    On Error GoTo errHandle
    
    If Not mrsCardType Is Nothing Then
        mrsCardType.Filter = 0
        If mrsCardType.RecordCount <> 0 Then mrsCardType.MoveFirst
        Set rsCardType_Out = mrsCardType: zl_PatiSvr_GetCardTypes = True: Exit Function
    End If
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ŀ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_PatiSvr_GetCardTypes
    'input
    '    cardtype_id C       �����id:NULL��ʾ���������ID����
    '    query_type  N   1   ��ѯ����:0-������Ϣ;1-������Ϣ(����:id,���룬����,���ų���,ǰ׺�ı�,�Ƿ�����,���㷽ʽ,�Ƿ�ȫ��,�Ƿ�����)
    '    cert_cardtype   N       ֻ��ȡ֤����Ϊ������ҽ�ƿ����:1-ֻ��ȡ�Ƿ�֤��=1��ҽ�ƿ�;0-ȫ����ȡ
    '    dffective_cardtype  N       ֻ��ȡ��Ч�Ŀ����:1-ֻ��ȡ��Ч�Ŀ����;0-ȫ����ȡ
    
    strJson = strJson & "" & GetJsonNodeString("cardtype_id", "", Json_num)
    strJson = strJson & "," & GetJsonNodeString("query_type", "", Json_num)
    strJson = strJson & "," & GetJsonNodeString("cert_cardtype", 0, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("dffective_cardtype", 1, Json_num, True)
    strJson = "{""input"":{" & strJson & "}}"
    
    'output
    '    cardtype_id N   1   ID
    '    cardtype_code   C   1   ����
    '    cardtype_name   C   1   ����
    '    cardtype_stname C   1   ����
    '    prefix_text C   1   ǰ׺�ı�
    '    cardno_len  N   1   ���ų���
    '    default    N   1   ȱʡ��־
    '    fixed N   1   �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
    '    strict   N   1   �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
    '    self_make N   1   �Ƿ�����:1-�ǵ�;0-����
    '    exist_account  N   1   �Ƿ�����ʻ�:1-�����ʻ�;0-�������˻�
    '    allow_return_cash    N   1   �Ƿ�����:1-����;0-������
    '    must_all_return   N   1   �Ƿ�ȫ��:1-����ȫ��;0-����������
    '    component   C   1   ����
    '    memo    C   1   ��ע
    '    spec_item   C   1   �ض���Ŀ
    '    blnc_mode   C   1   ���㷽ʽ
    '    blnc_nature N   1   ��������
    '    cardno_pwdtxt   C   1   ��������:���Ŵӵڼ�λ���ڼ�λ��ʾ����,��ʽΪ:S-N:S��ʾ�ӵڼ�λ��ʼ,���ڼ�λ����.����:3-10,��ʾ��3λ��10λ������*��ʾ:12********3323��Ҫ����Ӧ��ͬ����ҽ�ƿ�
    '    allow_repeat_use N   1   �Ƿ��ظ�ʹ��:1-����;0-������
    '    enabled    N   1   �Ƿ�����:1-������;0-δ����
    '    pwd_len N   1   ���볤��
    '    pwd_len_limit   N   1   ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
    '    pwd_rule    N   1   �������:��-���ֺ��ַ����;1-��Ϊ�������
    '    allow_vaguefind    N   1   �Ƿ�ģ������:1-֧��ģ������;0-��֧��
    '    pwd_require    N   1   ������������:0-������;1-������,����;2-�������ֹ;ȱʡΪ������
    '    default_pwd  N   1   �Ƿ�ȱʡ����:1-������֤��N(�����볤��Ϊ׼)λ��Ϊȱʡ����;0-��ȱʡ����
    '    allow_makecard N   1   �Ƿ��ƿ�:1-��;0-��
    '    allow_sendcard N   1   �Ƿ񷢿�:1-��;0-��
    '    allowwritecard    N   1   �Ƿ�д��:1-��;0-��
    '    insurance_type  N   1   ����
    '    sendcard_nature N   1   ��������:0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�������������ſ���������ʾ;ȱʡΪ0
    '    allow_transfer N   1   �Ƿ�ת�ʼ�����:1-֧��ת�ʼ�����;0-��֧��
    '    readcard_nature C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
    '    keyboard_mode   N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ��������;1-ʹ������������ ,2-ʹ���ַ�������
    '    advsend_buildqrcode N   1   �Ƿ�ҽ�����͵����������ɽӿ�:1-���͵������ɶ�ά��ӿ�;0-������
    '    holding_pay   N   1   �Ƿ�ֿ�����:1-��;0-��
    '    cert_cardtype    N   1   �Ƿ�֤�����͵�ҽ�ƿ�:0-���ǣ�1-��
    '    verfycard    N   1   �Ƿ��˿��鿨
    '    sendcard_sign   N   1   ��������:0��NULL-����ʱ�����ű���ﵽ���ų���;1-����ʱ����������С�ڵ��ڿ��ų���,����ʱ��С�ڿ��ų���ʱ������ʾ����Ա;2-����ʱ����������С�ڵ��ڿ��ų���,С��ʱ����ʾ����Ա��
    '    enterkey_enabled N   1   �豸�Ƿ����ûس�:ҽ�ƿ���Ӧ��ˢ���豸�Ƿ������˻س�����������˻س����򿨺ų���Ĭ������һλ�����λس�
    '    def_return_cash N   1   �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
    '    balalone N   1   �Ƿ��������:1-��������;0-�Ƕ�������
    '    discern_rule    N   1   ����ʶ�����:1-ȫ��ת��Ϊ��д;0-�����ִ�Сд
    '    def_valid_time  C   1   ȱʡ��Чʱ��:NULLʱ����ʾ������;�ǿ�ʱ����ʽΪ:ʱ��+��λ(�죬��),���磺3��,3��
    '    scanpay  N   1   �Ƿ�֧��ɨ�븶:�Ƿ�֧��ɨ�븶,֧��ʱ������á�zlReadQRCode������

    strServiceName = "zl_PatiSvr_GetCardTypes"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    Set cllData = objServiceCall.GetJsonListValue("output.type_list")
    If cllData Is Nothing Then Exit Function
    If cllData.Count = 0 Then Exit Function
  
    If cllData Is Nothing Then Exit Function
    If cllData.Count = 0 Then Exit Function
    
    'output
    '    cardtype_id N   1   ID
    '    cardtype_code   C   1   ����
    '    cardtype_name   C   1   ����
    '    cardtype_stname C   1   ����
    '    prefix_text C   1   ǰ׺�ı�
    '    cardno_len  N   1   ���ų���
    '    default    N   1   ȱʡ��־
    '    fixed N   1   �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
    '    strict   N   1   �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
    '    self_make N   1   �Ƿ�����:1-�ǵ�;0-����
    '    exist_account  N   1   �Ƿ�����ʻ�:1-�����ʻ�;0-�������˻�
    '    allow_return_cash    N   1   �Ƿ�����:1-����;0-������
    '    must_all_return   N   1   �Ƿ�ȫ��:1-����ȫ��;0-����������
    '    component   C   1   ����
    '    memo    C   1   ��ע
    '    spec_item   C   1   �ض���Ŀ
    '    blnc_mode   C   1   ���㷽ʽ
    '    blnc_nature N   1   ��������
    '    cardno_pwdtxt   C   1   ��������:���Ŵӵڼ�λ���ڼ�λ��ʾ����,��ʽΪ:S-N:S��ʾ�ӵڼ�λ��ʼ,���ڼ�λ����.����:3-10,��ʾ��3λ��10λ������*��ʾ:12********3323��Ҫ����Ӧ��ͬ����ҽ�ƿ�
    '    allow_repeat_use N   1   �Ƿ��ظ�ʹ��:1-����;0-������
    '    enabled    N   1   �Ƿ�����:1-������;0-δ����
    '    pwd_len N   1   ���볤��
    '    pwd_len_limit   N   1   ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
    '    pwd_rule    N   1   �������:��-���ֺ��ַ����;1-��Ϊ�������
    '    allow_vaguefind    N   1   �Ƿ�ģ������:1-֧��ģ������;0-��֧��
    '    pwd_require    N   1   ������������:0-������;1-������,����;2-�������ֹ;ȱʡΪ������
    '    default_pwd  N   1   �Ƿ�ȱʡ����:1-������֤��N(�����볤��Ϊ׼)λ��Ϊȱʡ����;0-��ȱʡ����
    '    allow_makecard N   1   �Ƿ��ƿ�:1-��;0-��
    '    allow_sendcard N   1   �Ƿ񷢿�:1-��;0-��
    '    allow_writecard    N   1   �Ƿ�д��:1-��;0-��
    '    insurance_type  N   1   ����
    '    sendcard_nature N   1   ��������:0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�������������ſ���������ʾ;ȱʡΪ0
    '    allow_transfer N   1   �Ƿ�ת�ʼ�����:1-֧��ת�ʼ�����;0-��֧��
    '    readcard_nature C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
    '    keyboard_mode   N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ��������;1-ʹ������������ ,2-ʹ���ַ�������
    '    advsend_buildqrcode N   1   �Ƿ�ҽ�����͵����������ɽӿ�:1-���͵������ɶ�ά��ӿ�;0-������
    '    holding_pay   N   1   �Ƿ�ֿ�����:1-��;0-��
    '    cert_cardtype    N   1   �Ƿ�֤�����͵�ҽ�ƿ�:0-���ǣ�1-��
    '    verfycard    N   1   �Ƿ��˿��鿨
    '    sendcard_sign   N   1   ��������:0��NULL-����ʱ�����ű���ﵽ���ų���;1-����ʱ����������С�ڵ��ڿ��ų���,����ʱ��С�ڿ��ų���ʱ������ʾ����Ա;2-����ʱ����������С�ڵ��ڿ��ų���,С��ʱ����ʾ����Ա��
    '    enterkey_enabled N   1   �豸�Ƿ����ûس�:ҽ�ƿ���Ӧ��ˢ���豸�Ƿ������˻س�����������˻س����򿨺ų���Ĭ������һλ�����λس�
    '    def_return_cash N   1   �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
    '    balalone N   1   �Ƿ��������:1-��������;0-�Ƕ�������
    '    discern_rule    N   1   ����ʶ�����:1-ȫ��ת��Ϊ��д;0-�����ִ�Сд
    '    def_valid_time  C   1   ȱʡ��Чʱ��:NULLʱ����ʾ������;�ǿ�ʱ����ʽΪ:ʱ��+��λ(�죬��),���磺3��,3��
    '    scanpay  N   1   �Ƿ�֧��ɨ�븶:�Ƿ�֧��ɨ�븶,֧��ʱ������á�zlReadQRCode������
    If GetCardTypeRecStru(rsCardType_Out) = False Then
        Set mrsCardType = Nothing: Set rsCardType_Out = Nothing:
     
        Exit Function
    End If
    
    For i = 1 To cllData.Count
        Set cllTemp = cllData(i)
        With rsCardType_Out
        
            If Val(Nvl(cllTemp("_enabled"))) = 1 And _
                (bln�������� And (Val(cllTemp("_self_make")) = 1 Or Val(cllTemp("_allow_sendcard")) = 1) Or bln�������� = False) Then
                .AddNew
                    !ID = cllTemp("_cardtype_id")
                    !���� = cllTemp("_cardtype_code")
                    !���� = cllTemp("_cardtype_name")
                    !���� = cllTemp("_cardtype_stname")
                    
                    !ǰ׺�ı� = cllTemp("_prefix_text")
                    !���ų��� = cllTemp("_cardno_len")
                    !ȱʡ��־ = cllTemp("_default")
                    !�Ƿ�̶� = cllTemp("_fixed")
                    
                    !�Ƿ��ϸ���� = cllTemp("_strict")
                    !�Ƿ����� = cllTemp("_self_make")
                    !�Ƿ�����ʻ� = cllTemp("_exist_account")
                    !�Ƿ����� = cllTemp("_allow_return_cash")
                    !�Ƿ�ȱʡ���� = cllTemp("_def_return_cash")
                    !�Ƿ�ȫ�� = cllTemp("_must_all_return")
                    !���� = cllTemp("_component")
                    !��ע = cllTemp("_memo")
                    
                    !�ض���Ŀ = cllTemp("_spec_item")
                    !���㷽ʽ = cllTemp("_blnc_mode")
                    !�������� = cllTemp("_cardno_pwdtxt")
                      
                    !�Ƿ��ظ�ʹ�� = cllTemp("_allow_repeat_use")
                    !�Ƿ����� = cllTemp("_enabled")
                    
                    !���볤�� = cllTemp("_pwd_len")
                    !���볤������ = cllTemp("_pwd_len_limit")
                    !������� = cllTemp("_pwd_rule")
                    !������������ = cllTemp("_pwd_require")
                    
                    !�Ƿ�ģ������ = cllTemp("_allow_vaguefind")
                    !�Ƿ�ȱʡ���� = cllTemp("_default_pwd")
                    
                    
                    
                    !�Ƿ��ƿ� = cllTemp("_allow_makecard")
                    !�Ƿ񷢿� = cllTemp("_allow_sendcard")
                    !�Ƿ�д�� = cllTemp("_allow_writecard")
                    !�������� = cllTemp("_sendcard_sign")
                    
                    !�������� = cllTemp("_blnc_nature")
                    !���� = cllTemp("_insurance_type")
                    !�������� = cllTemp("_sendcard_nature")
                    !�Ƿ�ת�ʼ����� = cllTemp("_allow_transfer")
                    !�������� = cllTemp("_readcard_nature")
                    !���̿��Ʒ�ʽ = cllTemp("_keyboard_mode")
                    
                    !�Ƿ�ֿ����� = cllTemp("_holding_pay")
                    !�Ƿ�֤�� = cllTemp("_cert_cardtype")
                    
                    !���͵��ýӿ� = cllTemp("_advsend_buildqrcode")
                    !�Ƿ��˿��鿨 = cllTemp("_verfycard")
                                    
                    !�Ƿ�������� = cllTemp("_balalone")
                    !ȱʡ��Чʱ�� = cllTemp("_def_valid_time")
                    !����ʶ����� = cllTemp("_discern_rule")
                    !�Ƿ�֧��ɨ�븶 = cllTemp("_scanpay")
                    !�Ƿ����ûس� = cllTemp("_enterkey_enabled")
                .Update
            End If
       End With
    Next
    If rsCardType_Out.RecordCount <> 0 Then rsCardType_Out.MoveFirst
    Set mrsCardType = rsCardType_Out
    zl_PatiSvr_GetCardTypes = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetCardTypeRecStru(ByRef rsCardType As ADODB.Recordset) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ṹ
     '����:rsCardType-���صļ�¼���ṹ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-24 18:10:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsCardType = New ADODB.Recordset
    With rsCardType
        If .State = 1 Then .Close
        'adBigInt
        .Fields.Append "ID", adVarNumeric, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        
        .Fields.Append "ǰ׺�ı�", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "���ų���", adSmallInt, 20, adFldIsNullable
        .Fields.Append "ȱʡ��־", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�̶�", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ��ϸ����", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�����", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�����ʻ�", adSmallInt, , adFldIsNullable
        
        .Fields.Append "�Ƿ�����", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�ȱʡ����", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�ȫ��", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ��ظ�ʹ��", adSmallInt, , adFldIsNullable
        .Fields.Append "��������", adSmallInt, , adFldIsNullable
        .Fields.Append "���볤��", adSmallInt, , adFldIsNullable
        .Fields.Append "���볤������", adSmallInt, , adFldIsNullable
        .Fields.Append "�������", adSmallInt, , adFldIsNullable
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��ע", adLongVarChar, 300, adFldIsNullable
        .Fields.Append "�ض���Ŀ", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "���㷽ʽ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�Ƿ�����", adSmallInt, , adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        
        .Fields.Append "������������", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�ȱʡ����", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�ģ������", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ��ƿ�", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�д��", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ񷢿�", adSmallInt, , adFldIsNullable
        .Fields.Append "��������", adSmallInt, , adFldIsNullable
        
        
        .Fields.Append "��������", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�֤��", adSmallInt, , adFldIsNullable
        
        
        .Fields.Append "�Ƿ�ת�ʼ�����", adSmallInt, , adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�Ƿ�ֿ�����", adSmallInt, , adFldIsNullable
        .Fields.Append "���͵��ýӿ�", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ��˿��鿨", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ��������", adSmallInt, , adFldIsNullable
        .Fields.Append "ȱʡ��Чʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ʶ�����", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ�֧��ɨ�븶", adSmallInt, , adFldIsNullable
        .Fields.Append "����", adSmallInt, , adFldIsNullable
        .Fields.Append "���̿��Ʒ�ʽ", adSmallInt, , adFldIsNullable
        .Fields.Append "�Ƿ����ûس�", adSmallInt, , adFldIsNullable
    
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    GetCardTypeRecStru = True
End Function
 

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mrsCardType = Nothing
    Set mobjServiceCall = Nothing
    
End Sub