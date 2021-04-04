Attribute VB_Name = "mdlPubServerPati"
Option Explicit
'*********************************************************************************************************************************************
'����:�����漰���ò��˵���ط���
'�ӿ�˵��:
'    1.zl_PatiSvr_GetPatiInfsByRange-������ȡ������Ϣ����
'    2.zl_PatiSvr_GetCardTypes-��ȡҽ�ƿ������Ϣ��
'    3.zl_PatiSvr_GetPatiID:��ȡָ�������Ĳ���Ids
'    4.zl_PatiSvr_GetPatiInfo:��ȡ������Ϣ��ϸ����ӿ�
'    5.zl_PatiSvr_GetPatiExtendInfo:��ȡ������Ϣ�ӱ���Ϣ����ӿ�
'����:���˺�
'����:2019*10*31 14:47:18
'*********************************************************************************************************************************************
Public gobjServiceCall  As Object
 

Private mlngErrNum As Long, mstrSource As String, mstrErrMsg As String
Private Function GetServiceCall(ByRef objServiceCall_Out As Object, Optional blnShowErrMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������
    '����:objServiceCall_Out-���ع����������
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    If Not gobjServiceCall Is Nothing Then Set objServiceCall_Out = gobjServiceCall: GetServiceCall = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
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
    If gobjServiceCall.InitService(gcnOracle, gstrDBUser, glngSys, glngModul) = False Then
        
        Set gobjServiceCall = Nothing: Exit Function
    End If
    Set objServiceCall_Out = gobjServiceCall
    GetServiceCall = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function


Public Function zl_PatiSvr_GetPatiInfsByRange(ByVal intQueryStatus As Integer, ByVal cllFilter As Collection, _
    ByRef cllPatiInfos_out As Collection, Optional ByVal str����Ids As String, Optional ByRef str����IDs As String, _
    Optional ByVal blnExpendInfo As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ��
    '���:intQueryStatus-��ѯ����(0-������;1-��Ժ ;2-���Ｐ��Ժ)
    '     cllFilter-��������
    '     str����Ids-����ID
    '     rsPatiPage-��ҳ��Ϣ
    '     str����IDs-��ǰ����Ids
    '����:cllPatiInfos_out-���ص����ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    Set cllPatiInfos_out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
   
    'zl_PatiSvr_GetPatiInfsByRange
    '  --����:��ȡ������Ϣ
    '  --��Σ�Json_In:��ʽ
    '  --    input
    '  --      query_type        N 1 0����ѯ������Ϣ��1����ѯ������Ϣ+��չ��Ϣ
    '  --      pati_ids          C   ����IDs:����ö���
    '  --      pati_name         C   ����:���Դ�%�ֺű������ƥ��
    '  --      pati_sex          C   �Ա�,����������������Ч
    '  --      birthdate_start   C   ��ʼ��������
    '  --      birthdate_end     C   ��ֹ��������
    '  --      outpatient_num    C   �����
    '  --      pati_idcard       C   ���֤��
    '  --      fee_category      C   �ѱ�
    '  --      pati_sex          C   �Ա�
    '  --      pati_area         C   ����
    '  --      insurance_num     C   ҽ����
    '  --      vcard_no          C   ���￨��
    '  --      iccard_no         C   Ic����
    '  --      wardarea_ids      C   ����ids������ö���
    '  --      qurey_Max         N   ��ѯ������¼����Ϊ0��NULLʱ��ʾ������
    '  --      qrspt_statu       N   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
    '  --      visit_star_time   C   ���￪ʼʱ��:yyyy-mm-dd hh24:mi:ss
    '  --      visit_end_time    C   �������ʱ��:yyyy-mm-dd hh24:mi:ss
    '  --      create_start_time C   ��ʼ�Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '  --      create_end_time   C   ��ֹ�Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '  --      module            N   ģ���:����Zl_Custom_Patiids_Get(�������֤���ز���id)����ʱ�贫��
    '  --      only_ctorg_pati   N   ֻ��ѯ��Լ��λ�Ĳ���
    '  --      ctt_unit_id       N   ��ͬ��λid,ֻ��ѯֻ��ѯ��Լ��λ�Ĳ���ʱ��Ч
    '  --      default_cardtype_id N   ȱʡ�����id
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", IIf(blnExpendInfo, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("qrspt_statu", intQueryStatus, Json_num)
    For i = 1 To cllFilter.count
        Select Case cllFilter(i)(0)
        Case "�Ǽ�ʱ��"
            strJson = strJson & "," & GetJsonNodeString("create_start_time", cllFilter(i)(1), Json_Text)
            strJson = strJson & "," & GetJsonNodeString("create_end_time", cllFilter(i)(2), Json_Text)
        Case "����ʱ��"
            strJson = strJson & "," & GetJsonNodeString("visit_start_time", cllFilter(i)(1), Json_Text)
            strJson = strJson & "," & GetJsonNodeString("visit_end_time", cllFilter(i)(2), Json_Text)
        Case "����ID"
            lng����ID = cllFilter(i)(1)
        Case "����"
            strJson = strJson & "," & GetJsonNodeString("pati_name", cllFilter(i)(1), Json_Text)
        Case "���￨��"
            strJson = strJson & "," & GetJsonNodeString("vcard_no", cllFilter(i)(1), Json_Text)
        Case "�����"
            strJson = strJson & "," & GetJsonNodeString("outpatient_num", Trim(cllFilter(i)(1)), Json_Text)
        Case "ҽ����"
            strJson = strJson & "," & GetJsonNodeString("insurance_num", cllFilter(i)(1), Json_Text)
        Case "���֤��"
            strJson = strJson & "," & GetJsonNodeString("pati_idcard", cllFilter(i)(1), Json_Text)
        Case "IC����"
            strJson = strJson & "," & GetJsonNodeString("iccard_no", cllFilter(i)(1), Json_Text)
        Case "�Ա�"
            strJson = strJson & "," & GetJsonNodeString("pati_sex", cllFilter(i)(1), Json_Text)
        Case "����"
            strJson = strJson & "," & GetJsonNodeString("pati_area", cllFilter(i)(1), Json_Text)
        Case "�ѱ�"
            strJson = strJson & "," & GetJsonNodeString("fee_category", cllFilter(i)(1), Json_Text)
        Case "��ѯ������¼��", "����¼��", "����"
            strJson = strJson & "," & GetJsonNodeString("qurey_Max", Val(cllFilter(i)(1)), Json_num, True)
        Case "סԺ��"
            If gobjOneDataObject.zlGetPatiIDFromInpatientNum(cllFilter(i)(1), lng����ID) = False Then Exit Function
        Case "ȱʡ�����ID"
            strJson = strJson & "," & GetJsonNodeString("default_cardtype_id", cllFilter(i)(1), Json_Text)
        Case "����Լ��λ����"
            strJson = strJson & "," & GetJsonNodeString("only_ctorg_pati", cllFilter(i)(1), Json_num)
        Case "��ͬ��λID"
            strJson = strJson & "," & GetJsonNodeString("ctt_unit_id", cllFilter(i)(1), Json_num)
        Case "����"
            strJson = strJson & "," & GetJsonNodeString("occasion", cllFilter(i)(1), Json_num)
        End Select
    Next
    
    If str����IDs <> "" Then
        strJson = strJson & "," & GetJsonNodeString("wardarea_ids", str����IDs, Json_Text)
    End If
    
    If lng����ID <> 0 Then
        If InStr("," & str����Ids & ",", "," & lng����ID & ",") = 0 Then str����Ids = IIf(str����Ids <> "", ",", "") & lng����ID
    End If
    If str����Ids <> "" Then
        strJson = strJson & "," & GetJsonNodeString("pati_ids", str����Ids, Json_Text)
    End If
    
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "zl_PatiSvr_GetPatiInfsByRange"
    If objServiceCall.CallService(strServiceName, strJson, , "zl_PatiSvr_GetPatiInfsByRange", glngModul) = False Then Exit Function
    
    '  --����      json
    '  --output
    '  -- code                   N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  -- message                C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  -- pati_list[]                ������Ϣ�б�
    '  --   pati_id              N 1 ����id
    '  --   pati_pageid          N 1 ��ҳid��������Ϣ.��ҳID
    '  --   pati_name            C 1 ����
    '  --   pati_sex             C 1 �Ա�
    '  --   pati_age             C 1 ����
    '  --   pati_birthdate       C 1 �������ڣ�yyyy-mm-dd hh24:mi:ss
    '  --   fee_category         C 1 �ѱ�
    '  --   outpatient_num       C 1 �����
    '  --   inpatient_num        C 1 סԺ��
    '  --   inp_times            N 1 סԺ����
    '  --   pati_nation          C 1 ����
    '  --   pati_idcard          C 1 ���֤��
    '  --   vcard_no             C 1 ���￨��
    '  --   phone_number         C 1 �ֻ���
    '  --   pati_education       C 1 ѧ��
    '  --   ocpt_name            C 1 ְҵ
    '  --   pati_identity        C 1 ���
    '  --   country_name         C 1 ����
    '  --   pat_home_addr        C 1 ��ͥ��ַ
    '  --   pati_area            C 1 ����
    '  --   emp_name             C 1 ������λ����
    '  --   pati_bed             C 1 ��ǰ����
    '  --   is_inhspt            N 1 �Ƿ���Ժ��1-��Ժ��0-����Ժ
    '  --   pati_type            C 1 ��������(��ͨ��ҽ��������)
    '  --   insurance_type       C 1 ����
    '  --   pati_wardarea_id     N 1 ��ǰ����id
    '  --   pati_wardarea_name   C 1 ��ǰ��������
    '  --   pati_dept_id         N 1 ��ǰ����id
    '  --   pati_dept_name       C 1 ��ǰ��������
    '  --   adta_time            C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
    '  --   adtd_time            C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
    '  --   create_time          C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '  --   medc_card_no         C   ҽ�ƿ���
    Set cllData = objServiceCall.GetJsonListValue("output.pati_list")
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    Set cllPatiInfos_out = cllData
    zl_PatiSvr_GetPatiInfsByRange = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_GetCardTypes(ByRef cllCardTypes_out As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ�����
    '���:
    '����:cllCardTypes_out-���صĿ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 16:53:48
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    
    On Error GoTo errHandle
    
    Set cllCardTypes_out = New Collection
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
    strJson = strJson & "," & GetJsonNodeString("dffective_cardtype", 0, Json_num, True)
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
    '    must_all_return   N   1   �Ƿ�ȫ��:1-����ȫ��;0-��������
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
    '    default_pwd  N   1   �Ƿ�ȱʡ����:1-�����֤��N(�����볤��Ϊ׼)λ��Ϊȱʡ����;0-��ȱʡ����
    '    allow_makecard N   1   �Ƿ��ƿ�:1-��;0-��
    '    allow_sendcard N   1   �Ƿ񷢿�:1-��;0-��
    '    allowwritecard    N   1   �Ƿ�д��:1-��;0-��
    '    insurance_type  N   1   ����
    '    sendcard_nature N   1   ��������:0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0
    '    allow_transfer N   1   �Ƿ�ת�ʼ�����:1-֧��ת�ʼ�����;0-��֧��
    '    readcard_nature C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
    '    keyboard_mode   N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
    '    advsend_buildqrcode N   1   �Ƿ�ҽ�����͵����������ɽӿ�:1-���͵������ɶ�ά��ӿ�;0-������
    '    holding_pay   N   1   �Ƿ�ֿ�����:1-��;0-��
    '    cert_cardtype    N   1   �Ƿ�֤�����͵�ҽ�ƿ�:0-���ǣ�1-��
    '    verfycard    N   1   �Ƿ��˿��鿨
    '    sendcard_sign   N   1   ��������:0��NULL-����ʱ�����ű���ﵽ���ų���;1-����ʱ��������С�ڵ��ڿ��ų���,����ʱ��С�ڿ��ų���ʱ������ʾ����Ա;2-����ʱ��������С�ڵ��ڿ��ų���,С��ʱ����ʾ����Ա��
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
    If cllData.count = 0 Then Exit Function
    Set cllCardTypes_out = cllData
    zl_PatiSvr_GetCardTypes = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_GetPatiID(ByVal cllFindCons As Collection, ByVal cllOtherFindCons As Collection, _
    ByRef cllPatiDatas_Out As Collection, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg As String, _
    Optional ByVal bln���ʹ��ʱ�� As Boolean = True, Optional ByVal bln���ͣ�û��ʧ As Boolean = True, _
    Optional ByRef intCardStatus As Integer, Optional ByVal blnNotReturnFalse As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ID��Ϣ
    '       cllFindCons-��������(array(�ӵ�����,�ӵ�ֵ))
    '                �ӵ����ư���:�����ID,����,��ά��,������,������)
    '       cllOtherFindCons-������������:array(��ѯ������,��ѯ������)
    '                   ��ѯ������:��:�����,���￨�ţ����֤�ŵ�
    '       blnNotShowErrMsg-����ʾ�������ʾ��Ϣ
    '      bln���ʹ��ʱ��-������������Ч
    '      bln���ͣ�û��ʧ-������������Ч
    '      blnNotReturnFalse-�����񷵻صļ���Ϊ��ʱ��������false
    '����:strErrMsg-���صĴ�����Ϣ
    '        lng����ID-���صĲ���ID
    '        cllPatiDatas_Out-���ز�����Ϣ����
    '����:���ҳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-19 09:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim strCardTypes  As String, strComminuty As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    Dim strOthers As String
    Dim strCardNo As String
    Dim strRQCode As String
     On Error GoTo errHandle
     
    Set cllPatiDatas_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
 
    'zl_PatiSvr_GetPatiID
    '  --����:����ָ��������ȡ������Ϣ�Ĳ���ID
    '  --��Σ�Json_In:��ʽ
    '  --    input
    '  --          card_find             C
    '  --              cardtype_id       N  1  ҽ�ƿ����ID:=0ʱ����ʾģ������
    '  --              card_no           C  1  ����
    '  --              qrcode            C     ��ά��
    '  --              is_check_usetime  N  1  �Ƿ���ʹ��ʱ��:1-���;0-�����
    '  --              is_check_stop     N  1  �Ƿ���ͣ�û��ʧ:1-���;0-�����
    '  --          comminuty_find        C
    '  --            comminuty_num       N  1  �������
    '  --            comminuty_code      C     ������
    '  --          other_cons_find       C
    '  --            find_name           C  1  ���ҵ�����
    '  --            find_text           C  1  ���ҵ��ı�
    strJson = ""
    If Not cllFindCons Is Nothing Then
        strCardTypes = "": strComminuty = ""
        For i = 1 To cllFindCons.count
            Select Case cllFindCons(i)(0)
            Case "�����ID"
                strCardTypes = strCardTypes & "," & GetJsonNodeString("cardtype_id", Val(cllFindCons(i)(1)), Json_num)
            Case "����"
                strCardTypes = strCardTypes & "," & GetJsonNodeString("card_no", cllFindCons(i)(1), Json_Text)
                strCardNo = cllFindCons(i)(1)
            Case "��ά��"
                If cllFindCons(i)(1) <> "" Then
                    strCardTypes = strCardTypes & "," & GetJsonNodeString("qrcode", cllFindCons(i)(1), Json_Text)
                    strRQCode = cllFindCons(i)(1)
                End If
            Case "�������", "����"
                strComminuty = strComminuty & "," & GetJsonNodeString("comminuty_num", cllFindCons(i)(1), Json_num)
            Case "������"
                strComminuty = strComminuty & "," & GetJsonNodeString("comminuty_code", cllFindCons(i)(1), Json_Text)
            End Select
        Next
        
        If strCardTypes <> "" Then
            strCardTypes = strCardTypes & "," & GetJsonNodeString("is_check_usetime", IIf(bln���ʹ��ʱ��, 1, 0), Json_num)
            strCardTypes = strCardTypes & "," & GetJsonNodeString("is_check_stop", IIf(bln���ͣ�û��ʧ, 1, 0), Json_num)
            strCardTypes = Mid(strCardTypes, 2)
            strCardTypes = "," & GetNodeString("card_find") & ":{" & strCardTypes & "}"
        End If
        
        If strComminuty <> "" Then
            strComminuty = Mid(strComminuty, 2)
            strComminuty = "," & GetNodeString("comminuty_find") & ":{" & strComminuty & "}"
        End If
        strJson = strJson & strCardTypes & strComminuty
    End If
    
    If Not cllOtherFindCons Is Nothing Then
        strOthers = ""
        For i = 1 To cllOtherFindCons.count
            strOthers = strOthers & ",{" & GetJsonNodeString("find_name", cllOtherFindCons(i)(0), Json_Text)
            strOthers = strOthers & "," & GetJsonNodeString("find_text", cllOtherFindCons(i)(1), Json_Text) & "}"
        Next
        If strOthers <> "" Then
            strOthers = Mid(strOthers, 2)
            strJson = strJson & "," & GetNodeString("other_cons_find") & ":" & strOthers & ""
        End If
    End If
    
    If strJson = "" Then
        strJson = strJson & "," & GetNodeString("card_find") & ":{" & GetJsonNodeString("cardtype_id", 0, Json_num) & "}"
    End If
    
    strJson = Mid(strJson, 2)
    strJson = "{""input"":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetPatiID"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    'output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    pati_list[] C   1   �����б�,ģ������ʱ�����ܴ��ڶ��
    '        cardtype_id N   1   �����ID
    '        pati_id N   1   ����ID:δ�ҵ�ʱҲ�ɹ�������0
    '        card_pwd    C   1   ����
    '        pati_pageid N       ��ҳID
    '        card_status N   1   ��ǰ��״̬��0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��;3-ʧЧ��������ҽ�ƿ���Ϣ.��ֹʹ��ʱ�䵽��ʱ���ظ�״̬����������ʹ�ã�

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If blnNotReturnFalse Then zl_PatiSvr_GetPatiID = intReturn = 1
    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        intCardStatus = Val(NVL(objServiceCall.GetJsonNodeValue("output.card_status")))
        If strErrMsg = "" Then
            If strCardNo = "" Then
                strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
            Else
                strErrMsg = "δ�ҵ�����Ϊ" & strCardNo & "�Ĳ��ˣ�����ÿ��Ƿ���Ч����"
            End If
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    intCardStatus = objServiceCall.GetJsonNodeValue("output.card_status")
    Set cllData = objServiceCall.GetJsonListValue("output.pati_list")
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    Set cllPatiDatas_Out = cllData
    zl_PatiSvr_GetPatiID = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_PatiSvr_GetPatiInfo(ByVal lng����ID As Long, _
    ByVal cllOtherFindCons As Collection, ByRef cllPatiDatas_Out As Collection, _
    Optional ByVal int��ѯ���� As Integer = 0, _
    Optional ByVal bln�������� As Boolean, _
    Optional ByVal bln��������ҩ�� As Boolean, _
    Optional ByVal bln����������Ϣ As Boolean, _
    Optional ByVal bln��������Ϣ As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg As String, _
    Optional ByVal bln����ҽ������ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ϸ��Ϣ����ӿ�
    '���:cllOtherFindCons-������������(array(��ѯ����,��ѯֵ)
    '             ��ѯ����:����IDS,����,�Ա�,�������ڵ�,��query_cons_list[]�б��е���������
    '      int��ѯ����-0-����;1-����+��ϵ��;2-����
    '����:cllPatiDatas_Out-���ز�����Ϣ��
    '
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 18:02:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Collection, cllExpend As Collection, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strJsonTemp As String
    
    On Error GoTo errHandle
    
    Set cllPatiDatas_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    If Not cllOtherFindCons Is Nothing Then
        For i = 1 To cllOtherFindCons.count
            Select Case UCase(cllOtherFindCons(i)(0))
            Case "����IDS"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_ids", cllOtherFindCons(i)(1), Json_Text)
            Case "����"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_name", cllOtherFindCons(i)(1), Json_Text)
            Case "�����"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("outpatient_num", Trim(cllOtherFindCons(i)(1)), Json_Text)
            Case "���֤��", "�������֤", "���֤"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_idcard", cllOtherFindCons(i)(1), Json_Text)
            Case "��ϵ�����֤"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("contacts_idcard", cllOtherFindCons(i)(1), Json_Text)
            Case "ҽ����", "ҽ��֤��"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_num", cllOtherFindCons(i)(1), Json_Text)
            Case "ҽ�ƿ����ID", "�����ID"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(cllOtherFindCons(i)(1)), Json_num, True)
            Case "����"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("card_no", cllOtherFindCons(i)(1), Json_Text)
            Case "��ά��"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("qrcode", cllOtherFindCons(i)(1), Json_Text)
            Case "IC����", "IC", "IC��"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("iccard_no", cllOtherFindCons(i)(1), Json_Text)
            Case "��ѯסԺ״̬", "סԺ״̬"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("qrspt_statu", Val(cllOtherFindCons(i)(1)), Json_num, True)
            Case "�ֻ���", "�ֻ�"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("phone_number", cllOtherFindCons(i)(1), Json_Text)
            Case "���￨��", "���￨"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("visit_card", cllOtherFindCons(i)(1), Json_Text)
            Case Else
                strErrMsg = "Ŀǰ�ݲ���֧�ְ����Ϊ��" & UCase(cllOtherFindCons(i)(0)) & "�������Ҳ��ˣ�"
                If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            End Select
        Next
        If strJsonTemp <> "" Then strJsonTemp = Mid(strJsonTemp, 2)
    End If
    If lng����ID = 0 And strJsonTemp = "" Then
        strErrMsg = "����Ч�Ĳ�ѯ���������飡"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
     
    'zl_PatiSvr_GetPatiInfo
    '  --����:��ȡ������Ϣ
    '  --��Σ�Json_In:��ʽ
    '  --    input
    '  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
    '  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
    '  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
    '  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
    '  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
    '  --      query_immune      N 1 �Ƿ����������:1-����;0-������
    '  --      query_insurance_pwd C  �Ƿ����ҽ������:1-����;0-������
    '  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
    '  --        pati_ids        C   ����IDs:����ö���
    '  --        pati_name       C   ����:���Դ�%�ֺű������ƥ��
    '  --        outpatient_num  C   �����
    '  --        pati_idcard     C   ���֤��
    '  --        contacts_idcard C   ��ϵ�����֤��
    '  --        cardtype_id     N   ҽ�ƿ����ID
    '  --        medc_card_name  N   ҽ�ƿ�����
    '  --        card_no         C   ����
    '  --        qrcode          C   ��ά��
    '  --        iccard_no       C   Ic����
    '  --        visit_card      C   ���￨��
    '  --        insurance_num   C   ҽ����
    '  --        qrspt_statu     C   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
    '  --        phone_number    C   �ֻ���
    '  --        pati_bed        C   ��ǰ����
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_type", int��ѯ����, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_family", IIf(bln��������, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_drug", IIf(bln��������ҩ��, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_immune", IIf(bln����������Ϣ, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_card", IIf(bln��������Ϣ, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_insurance_pwd", IIf(bln����ҽ������, 1, 0), Json_num)
    strJson = strJson & "," & GetNodeString("query_cons_list") & ":{" & strJsonTemp & "}"
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetPatiInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    
    '����            json    ����    ����+��ϵ�� ����
    'output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�  ��  ��  ��
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ ��  ��  ��
    '    pati_list[]         ������Ϣ�б�    ��  ��  ��
    '    pati_id N   1   ����id  ��  ��  ��
    '    pati_pageid N   1   ��ҳid��������Ϣ.��ҳID ��  ��  ��
    '    pati_name   C   1   ����    ��  ��  ��
    '    pati_sex    C   1   �Ա�    ��  ��  ��
    '    pati_age    C   1   ����    ��  ��  ��
    '    pati_birthdate  C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss ��  ��  ��
    '    fee_category    C   1   �ѱ�    ��  ��  ��
    '    outpatient_num  C   1   �����  ��  ��  ��
    '    inpatient_num   C   1   סԺ��
    '    mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����    ��  ��  ��
    '    mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����    ��  ��  ��
    '    pati_nation C   1   ����    ��  ��
    '    insurance_num   C   1   ҽ����  ��  ��  ��
    '    pati_idcard C   1   ���֤��    ��  ��  ��
    '    vcard_no    C   1   ���￨��            ��
    '    iccard_no   C   1   Ic����          ��
    '    health_num  C   1   ������          ��
    '    pati_education  C   1   ѧ��            ��
    '    ocpt_name   C   1   ְҵ            ��
    '    pati_identity   C   1   ���            ��
    '    ntvplc_name C   1   ����            ��
    '    country_name    C   1   ����            ��
    '    pati_marital_cstatus    C   1   ����״��            ��
    '    pat_home_addr   C   1   ��ͥ��ַ    ��  ��  ��
    '    pat_home_phno   C   1   ��ͥ�绰    ��  ��  ��
    '    pat_home_postcode   C   1   ��ͥ��ַ�ʱ�            ��
    '    pati_area   C   1   ����            ��
    '    pati_birthplace C   1   �����ص�    ��  ��  ��
    '    pat_hous_addr   C   1   ���ڵ�ַ            ��
    '    pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�            ��
    '    emp_name    C   1   ������λ����            ��
    '    emp_phno    C   1   ��λ�绰            ��
    '    emp_postcode    C   1   ��λ�ʱ�            ��
    '    emp_bank_name   C   1   ��λ������          ��
    '    ctt_unit_id N   1   ��ͬ��λID          ��
    '    phone_number    C   1   �ֻ���  ��  ��  ��
    '    pati_bed    C   1   ��ǰ����    ��  ��  ��
    '    pati_type   C   1   ��������(��ͨ��ҽ��������)          ��
    '    insurance_type  C   1   ����    ��  ��  ��
    '    pati_wardarea_id    N   1   ��ǰ����id          ��
    '    pati_wardarea_name  C   1   ��ǰ��������            ��
    '    pati_dept_id    N   1   ��ǰ����id          ��
    '    pati_dept_name  C   1   ��ǰ��������            ��
    '    adta_time   C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss          ��
    '    adtd_time   C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss          ��
    '    contacts_name   C   1   ��ϵ������      ��  ��
    '    contacts_relation   C   1   ��ϵ�˹�ϵ      ��  ��
    '    contacts_idcard C   1   ��ϵ�����֤��      ��  ��
    '    contacts_addr   C   1   ��ϵ�˵�ַ      ��  ��
    '    contacts_phno   C   1   ��ϵ�˵绰      ��  ��
    '    pat_grdn_name   C   1   �໤��          ��
    '    cert_no_other   C   1   ����֤��            ��
    '    is_inhspt   C   1   �Ƿ���Ժ:1-��Ժ ;0-����Ժ   ��  ��  ��
    '    pati_show_color N   1   ������ʾ��ɫ            ��
    '    visit_room  C   1   ��������            ��
    '    visit_statu N   1   ����״̬            ��
    '    visit_time  C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss          ��
    '    create_time C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss          ��
    '    pati_email           C   1   email
    '    pati_qq              C   1   qq
    '    card_captcha         C   1  ����֤��
    '    insurance_pwd        C       ҽ������
    '    family_list[]   C   1   ������Ա:���˼���() query_family=1����
    '        family_id   N   1   ����id  query_family=1
    '        family_relation C   1   ��ϵ
    '    drug_list[] C   1   ����ҩ���б�    query_drug=1ʱ����
    '        pat_algc_cadn_id    N   1   ����ҩƷID
    '        pat_algc_cadn   C   1   ����ҩ������
    '        allergy_info    C   1   ��ÿҩ�ﷴӦ
    '    immune_list[]   C   1   ���������б�    query_immune=1ʱ����
    '        vaccinate_time  C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '        vaccinate_name  C   1   ��������
    '    card_list[] C   1   ����ҽ�ƿ���Ϣ�б�(��������д����˿����ID�ģ��򷵻ظÿ����Ŀ���Ϣ)  query_card=1ʱ����
    '        cardtype_id N   1   ҽ�ƿ����ID
    '        card_no C   1   ����
    '        card_pwd    C   1   ����
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg = "" Then
            strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllData = objServiceCall.GetJsonListValue("output.pati_list")
    If bln�������� Or bln��������ҩ�� Or bln����������Ϣ Or bln��������Ϣ Then
        For i = 1 To cllData.count
            If bln�������� Then
                cllData(i).Remove "_family_list"
                Set cllExpend = objServiceCall.GetJsonListValue("output.pati_list[" & i - 1 & "].family_list")
                If Not cllExpend Is Nothing Then
                    cllData(i).Add cllExpend, "_family_list"
                End If
            End If
            
            If bln��������ҩ�� Then
                cllData(i).Remove "_drug_list"
                Set cllExpend = objServiceCall.GetJsonListValue("output.pati_list[" & i - 1 & "].drug_list")
                If Not cllExpend Is Nothing Then
                    cllData(i).Add cllExpend, "_drug_list"
                End If
            End If
            
            If bln����������Ϣ Then
                cllData(i).Remove "_immune_list"
                Set cllExpend = objServiceCall.GetJsonListValue("output.pati_list[" & i - 1 & "].immune_list")
                If Not cllExpend Is Nothing Then
                    cllData(i).Add cllExpend, "_immune_list"
                End If
            End If
            
            If bln��������Ϣ Then
                cllData(i).Remove "_card_list"
                Set cllExpend = objServiceCall.GetJsonListValue("output.pati_list[" & i - 1 & "].card_list")
                If Not cllExpend Is Nothing Then
                    cllData(i).Add cllExpend, "_card_list"
                End If
            End If
        Next
    End If
    'Set clldata = objServiceCall.GetJsonListValue("output.pati_list[0].drug_list")
    
'    If cllData Is Nothing Then
'        strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
'        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
'    If cllData.count = 0 Then
'        strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
'        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
    
    Set cllPatiDatas_Out = cllData
    zl_PatiSvr_GetPatiInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_GetPatiExtendInfo(ByVal lng����ID As Long, ByVal str��Ϣ���� As String, ByRef cllPatiData_Out As Collection, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg As String, Optional ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ�ӱ���Ϣ����ӿ�
    '���:str��Ϣ����-����ö��ŷ���,�磺ҽѧ��ʾ,��ϵ��2,��ϵ��3��
    '
    '����:cllPatiData_Out-���ز��˴ӱ���Ϣ���ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 20:10:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    
     
    On Error GoTo errHandle
    
   
    Set cllPatiData_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
       
    'zl_PatiSvr_GetPatiExtendInfo
    'input
    '    pati_id N   1   ����id
    '    info_names  C   1   ��Ϣ��������ö���
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("info_names", str��Ϣ����, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("visit_id", lng����ID, Json_num, True)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetPatiExtendInfo"
   
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    'output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   "Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ,ʧ��ʱ���ؾ���Ĵ�����Ϣ ""
    '    slave_list[]    C       �ӱ�����Ϣ�б�
    '       info_name   C   1   ��Ϣ��
    '        info_value  N   1   ��Ϣֵ
    '        visit_id        n 1 ����ID
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg = "" Then
            strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllPatiData_Out = objServiceCall.GetJsonListValue("output.slave_list")
    zl_PatiSvr_GetPatiExtendInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Patisvr_GetPatiCardInfo(ByVal strCardTypeIDs As String, ByVal str����ID As String, _
                Optional ByVal intQueryType As Integer = 1, Optional ByVal blnOnlyCardTypeID As Boolean, _
                Optional ByVal strCardTypes As String, Optional ByRef cllData As Collection, _
                Optional ByVal blnCertCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ӹ���������м���ָ�����˳�����Ч���Ŀ����
    '���: strCardTypeIDs ��������𣬶���ö��ŷָ�
    '      intQueryType:��ѯ��������:0-ֻ��ȡ����ID,1-ֻ��ȡ�����ID;2-�������˻�����Ϣ;3-����
    '����:���ز��˳�����Ч���Ŀ���𣬶���ö��ŷָ�
    '����:���˺�
    '����:2018-12-03 15:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllTemp As Collection
    Dim objServiceCall As Object
    Dim strCards As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
   
    'Zl_Patisvr_Getpaticardinfo
    'input
    '    pati_ids       C   1   ����id�������Ӣ�ĵĶ��ŷָ�
    '    cardtype_ids   C   1   �����IDs,����ö��ŷ���
    '    card_no        C       ����
    '    cert_cardtype  N   1   ֻ��ȡ֤����Ϊ������ҽ�ƿ����:1-ֻ��ȡ�Ƿ�֤��=1��ҽ�ƿ�;0-ȫ����ȡ
    '    query_type     N   1   ��ѯ��������:0-ֻ��ȡ����ID,1-ֻ��ȡ�����ID;2-�������˻�����Ϣ;3-����
    '    dffective_cardtype  N       ֻ��ȡ��Ч�Ŀ����:1-ֻ��ȡ��Ч�Ŀ����;0-ȫ����ȡ
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_ids", str����ID, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cardtype_ids", strCardTypeIDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("query_type", intQueryType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("dffective_cardtype", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cert_cardtype", IIf(blnCertCard, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "Zl_Patisvr_Getpaticardinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    'output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   "Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ
    '    ʧ��ʱ���ؾ���Ĵ�����Ϣ ""
    '    card_list[] C       ���˿���Ϣ�б�
    '        pati_id N   1   ����id
    '        pati_name   C       ����
    '        pati_sex    C       �Ա�
    '        pati_age    C       ����
    '        pati_birthdate  C       �������ڣ�yyyy-mm-dd hh24:mi:ss
    '        outpatient_num  C       �����
    '        pati_idcard C       ���֤��
    '        cardtype_id N   1   �����ID
    '        card_no C   1   ����
    '        card_qrcode C   1   ��ά��
    '        card_passwod    C   1   ����
    '        cardtype_name   C   1   ���������
    '        cardtype_cardlen    N   1   ���ų���
    '        card_statu  N   1   ״̬:0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
    '        loscard_creator C   1   ��ʧ��
    '        loscard_time    C   1   ��ʧʱ��:yyyy-mm-dd hh24:mi:ss
    '        loscard_mode    C   1   ��ʧ��ʽ
    '        sendcard_oper   C   1   ������
    '        end_time    C   1   ��ֹʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
    Set cllData = objServiceCall.GetJsonListValue("output.card_list")
    If cllData Is Nothing Then Exit Function
    
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        If InStr(strCards & ",", "," & cllTemp("_cardtype_id") & ",") = 0 Then
            strCards = strCards & "," & cllTemp("_cardtype_id")
        End If
    Next
    strCardTypes = Mid(strCards, 2)
    Zl_Patisvr_GetPatiCardInfo = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function


Public Function zl_PatiSvr_GetInsureByPatiID(lng����ID As Long, Optional ByRef int����_Out As Integer, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ҽ�������Ƿ����δ�����
    '���:lng����ID
    '     blnNotShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:int����_Out-����
    '     strErrMsg_out-���صĴ�����Ϣֵ
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-05 16:40:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData  As Collection, cllTemp  As Collection
    Dim objServiceCall As Object
    Dim intReturn As Integer
 
    
    On Error GoTo ErrH
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
     
    'zl_PatiSvr_GetInsureByPatiID
    '    ��� json
    '    input
    '    pati_id N   1   ����id
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_ids", lng����ID, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetInsureByPatiID"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, Not blnNotShowErrMsg) = False Then Exit Function
    '    ���� json
    '    output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    insurance_type  N   1   ����
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
                strErrMsg_Out = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    int����_Out = objServiceCall.GetJsonNodeValue("output.insurance_type")
    zl_PatiSvr_GetInsureByPatiID = True
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function


Public Function zl_PatiSvr_CheckOutNoIsExist(ByVal lng����ID As Long, ByVal str����� As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���������Ƿ�����ʹ��
    ' ��� : str�����-������������
    ' ���� : blnUsedByOther:T:������ʹ��
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim intReturn As Integer
    
    On Error GoTo ErrHand
    '��ȡ���÷���ӿ�
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpatient_num", str�����, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_CheckOutNoIsExist", strJson, , , glngModul, Not blnNotShowErrMsg) = False Then Exit Function
    '    ���� json
    '    output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    insurance_type  N   1   ����
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    blnUsedByOther = Val(objServiceCall.GetJsonNodeValue("output.isexist")) <> 0
    
    zl_PatiSvr_CheckOutNoIsExist = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_PhoneNumberExist(ByVal lng����ID As Long, ByVal str�ֻ��� As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ����ֻ����Ƿ�����ʹ��
    ' ��� :
    ' ���� : blnUsedByOther:T:������ʹ��
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim intReturn As Integer
    
    On Error GoTo ErrHand
    '��ȡ���÷���ӿ�
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("phone_number", str�ֻ���, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_PhoneNumberExist", strJson, , , glngModul, Not blnNotShowErrMsg) = False Then Exit Function
    '    ���� json
    '    output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    insurance_type  N   1   ����
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    blnUsedByOther = Val(objServiceCall.GetJsonNodeValue("output.exist")) <> 0
    
    zl_PatiSvr_PhoneNumberExist = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_CheckInsNoIsExist(ByVal strҽ���� As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���ҽ�����Ƿ�����ʹ��
    ' ��� :
    ' ���� : blnUsedByOther:T:������ʹ��
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim intReturn As Integer
    
    On Error GoTo ErrHand
    '��ȡ���÷���ӿ�
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = GetJsonNodeString("insurance_num", strҽ����, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_CheckInsNoIsExist", strJson, , , glngModul, Not blnNotShowErrMsg) = False Then Exit Function
    '    ���� json
    '    output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    insurance_type  N   1   ����
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    blnUsedByOther = Val(objServiceCall.GetJsonNodeValue("output.exist")) <> 0
    
    zl_PatiSvr_CheckInsNoIsExist = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Patisvr_GetPatiFamilyMember(ByVal byt��ѯ���� As Byte, ByVal lng����ID As Long, _
    ByRef str����IDs As String, Optional ByRef rs������Ϣ As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID����ȡ�ò��˵ļ�����Ա��Ϣ
    '���:
    '   byt��ѯ����=��ѯ���ͣ�0-ֻ���ؼ�����Ա����id��1-��ѯ������Ա�Ļ�����Ϣ
    '����:
    '   str����IDs=������Ա����id,���Ӣ�Ķ��ŷָ�
    '   rs������Ϣ=������Ա�Ļ�����Ϣ,�� ��ѯ����=1 ����Ч,�ֶΣ�����ID,��ϵ,����,�Ա�,����,��������,����,���֤��
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, cllData As Collection
    Dim cllTemp As Collection, i As Integer
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    str����IDs = ""
    If byt��ѯ���� = 1 Then
        Set rs������Ϣ = New ADODB.Recordset
        With rs������Ϣ.fields
            .Append "����ID", adBigInt, 18, adFldIsNullable
            .Append "��ϵ", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "�Ա�", adVarChar, 4, adFldIsNullable
            .Append "����", adVarChar, 20, adFldIsNullable
            .Append "��������", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 20, adFldIsNullable
            .Append "���֤��", adVarChar, 18, adFldIsNullable
        End With
        rs������Ϣ.CursorLocation = adUseClient
        rs������Ϣ.LockType = adLockOptimistic
        rs������Ϣ.CursorType = adOpenStatic
        rs������Ϣ.Open
    End If
    
    If lng����ID = 0 Then Zl_Patisvr_GetPatiFamilyMember = True: Exit Function
    'Zl_Patisvr_Getpatifamilymember
    '  --���ܣ����ݲ���ID����ȡ�ò��˵ļ�����Ա��Ϣ
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    pati_id              N   1  ����ID
    '  --    query_type           N   1  ��ѯ���ͣ�0-ֻ���ؼ�����Ա����id��1-��ѯ������Ա�Ļ�����Ϣ
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    family_list[]        C       ������Ա:���˼���
    '  --      pati_id            N   1   ��������ID
    '  --      pati_relation      C   1   ��ϵ
    '  --      pati_name          C   1   ����
    '  --      pati_sex           C   1   �Ա�
    '  --      pati_age           C   1   ����
    '  --      pati_birthdate     C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
    '  --      pati_nation        C   1   ����
    '  --      pati_idcard        C   1   ���֤��
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", byt��ѯ����, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_Getpatifamilymember", strJson, , "", glngModul) = False Then Exit Function
    Set cllData = objServiceCall.GetJsonListValue("output.family_list")
    If cllData Is Nothing Then Exit Function
    
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        str����IDs = str����IDs & "," & cllTemp("_pati_id")
        If byt��ѯ���� = 1 Then
            With rs������Ϣ
                .AddNew
                !����ID = cllTemp("_pati_id")
                !��ϵ = cllTemp("_pati_relation")
                !���� = cllTemp("_pati_name")
                !�Ա� = cllTemp("_pati_sex")
                !���� = cllTemp("_pati_birthdate")
                !�������� = cllTemp("_pati_birthdate")
                !���� = cllTemp("_pati_nation")
                !���֤�� = cllTemp("_pati_idcard")
            End With
        End If
    Next
    If byt��ѯ���� = 1 Then
        If rs������Ϣ.RecordCount > 0 Then rs������Ϣ.MoveFirst
    End If
    If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    Zl_Patisvr_GetPatiFamilyMember = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function Zl_Patisvr_UpdateOutpatiState(ByVal lng����ID As Long, ByVal cllUpdateInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���²��˾���״̬��������Ϣ
    '���:
    '   lng����ID=����ID
    '   cllUpdateInfo=������Ϣ����Ա:(�ֻ���(N),�ѱ�(N),��������(N),����״̬(N),����ʱ��(N),�����(N))
    '����:
    '����:ִ�гɹ�����True��ʧ�ܷ���False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lng����ID = 0 Then Exit Function
    If cllUpdateInfo Is Nothing Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Patisvr_Updateoutpatistate
    '  --���ܣ��������ﲡ�˾���״̬
    '  --      �����жϽ���Ǵ��ڣ�����������򲻸��£����߸���Ϊԭֵ��Ŀǰ��ʱδ�õ��������Ի������������չ
    '  --��Σ�Json_In:��ʽ
    '  --input
    '  --    pati_id            N 1 ����id
    '  --    phone_number       C   �����ֻ���
    '  --    fee_category       C   �ѱ�
    '  --    visit_room         C   ���µľ�������
    '  --    visit_status       N   ���µľ���״̬
    '  --    visit_time         C   ���µľ���ʱ��
    '  --    outpatient_num             C   �����
    '  --����: Json_Out,��ʽ����
    '  --    output
    '  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    If CollectionExitsValue(cllUpdateInfo, "�ֻ���") Then
        strJson = strJson & "," & GetJsonNodeString("phone_number", cllUpdateInfo("�ֻ���"), Json_Text)
    End If
    If CollectionExitsValue(cllUpdateInfo, "�ѱ�") Then
        strJson = strJson & "," & GetJsonNodeString("fee_category", cllUpdateInfo("�ѱ�"), Json_Text)
    End If
    If CollectionExitsValue(cllUpdateInfo, "��������") Then
        strJson = strJson & "," & GetJsonNodeString("visit_room", cllUpdateInfo("��������"), Json_Text)
    End If
    If CollectionExitsValue(cllUpdateInfo, "����״̬") Then
        strJson = strJson & "," & GetJsonNodeString("visit_status", cllUpdateInfo("����״̬"), Json_num)
    End If
    If CollectionExitsValue(cllUpdateInfo, "����ʱ��") Then
        strJson = strJson & "," & GetJsonNodeString("visit_time", cllUpdateInfo("����ʱ��"), Json_Text)
    End If
    If CollectionExitsValue(cllUpdateInfo, "�����") Then
        strJson = strJson & "," & GetJsonNodeString("outpatient_num", cllUpdateInfo("�����"), Json_Text)
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_Updateoutpatistate", strJson, , "", glngModul) = False Then Exit Function
    
    Zl_Patisvr_UpdateOutpatiState = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_PatiSvr_GetPatiIdsByRange(ByVal strCondition As String, ByRef strPatiIds As String, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String, _
    Optional ByVal blnFindByFilter As Boolean, Optional ByVal cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ֵ��ȡ���������Ĳ���ID
    '���:
    '   strCondition=�����Ǿ��￨�š����֤�š�IC���š������
    '   blnFindByFilter=True:����������(cllFilter)��ȡ;False:��strCondition��ȡ
    '   cllFilter=��������:Array(Key,Value),Key:��ͬ��λID
    '����:
    '����:ִ�гɹ�����True��ʧ�ܷ���False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intReturn As Integer
    Dim objServiceCall As Object
    Dim strJson As String, i As Long
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'zl_PatiSvr_GetPatiIdsByRange
    '  --��Σ�Json_In:��ʽ
    '  --input
    '  --    query_condition C 1 ��ѯ����
    '  --    ctt_unit_id     N 1 ��ͬ��λID����ѯָ����ͬ��λ�����ﲡ��
    '  --����: Json_Out,��ʽ����
    '  --    output
    '  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --        pati_ids                C   1   ����IDs
    strJson = ""
    If blnFindByFilter = False Then
        strJson = strJson & "" & GetJsonNodeString("query_condition", strCondition, Json_Text)
    Else
        For i = 1 To cllFilter.count
            Select Case cllFilter(i)(0)
            Case "��ͬ��λID"
                strJson = strJson & "" & GetJsonNodeString("ctt_unit_id", cllFilter(i)(1), Json_num)
            End Select
        Next
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_GetPatiIdsByRange", strJson, , "", glngModul) = False Then Exit Function
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strPatiIds = objServiceCall.GetJsonNodeValue("output.pati_ids")
    zl_PatiSvr_GetPatiIdsByRange = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_PatiSvr_GetInputItemLength(ByVal strTableItems As String, ByRef cllColumn As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������Ϣ��ָ���ֶε���󳤶�
    ' ��� : strTableItems����1:��1,��2|��2:��1,��2,��3|..
    ' ���� : cllcolumn (Collect):��Ա(����,�ֶ�,�ֶγ���)
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/21 20:16
    '---------------------------------------------------------------------------------------
    Dim strJson As String, strSubJson As String, i As Long, strServiceName  As String
    Dim varData As Variant, varTmp As Variant
    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If strTableItems = "" Then Exit Function
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Patisvr_Getinsureinfo
    '--input
    '--  item_list[]
    '--  table_name  C   1   ����
    '--  column_name C   1   ����,����ö���
    varData = Split(strTableItems, "|")
    
    For i = 0 To UBound(varData)
        varTmp = Split(varData(i) & ":", ":")
        strSubJson = strSubJson & ",{"
        strSubJson = strSubJson & "" & GetJsonNodeString("table_name", varTmp(0), Json_Text)
        strSubJson = strSubJson & "," & GetJsonNodeString("column_name", varTmp(1), Json_Text)
        strSubJson = strSubJson & "}"
    Next
    strJson = GetNodeString("item_list") & ":[" & Mid(strSubJson, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "zl_PatiSvr_GetInputItemLength"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    Set cllColumn = objServiceCall.GetJsonListValue("output.item_list")
    
    zl_PatiSvr_GetInputItemLength = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function Zl_Patisvr_Getinsureinfo(ByVal lng����ID As Long, ByRef insure As Integer, Optional insureName As String, _
                                                              Optional strҽ���� As String, Optional strҽ������ As String, Optional str���� As String, _
                                                              Optional str�Ǽ�ʱ�� As String, Optional lng����id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���id�����࣬��ȡ���˵ı�����Ϣ
    '���: insure :����
    '����:���ز���ҽ����Ϣ
    '����:
    '����:2019-11-20 19:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object
    Dim bln�������� As Boolean
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    bln�������� = insure = 0
    'Zl_Patisvr_Getinsureinfo
    'input
    '     pati_id        N   1      ����id
    '     insure_type N           ���ࣺδ��������ʱ�����ݲ���id��ѯ������Ϣ�е����ࣻ��������ʱ�����ݲ���id�������ѯ�����ʻ��е����ࡢ�������ơ�ҽ���š��Ǽ�ʱ�䡢ҽ������
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("insure_type", insure, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "Zl_Patisvr_Getinsureinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    ' output
    '    insure_type C   1   ����
    '    insure_name C   1   ��������
    '    insure_no   C   1   ҽ����
    '    card_no        C       ����
    '    pati_create_time    C   1   ���˵ĵǼ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '    insure_pwd  C   1   ҽ������
    insure = Val(NVL(objServiceCall.GetJsonNodeValue("output.insure_type")))
    If Not bln�������� Then
        insureName = objServiceCall.GetJsonNodeValue("output.insure_name")
        strҽ���� = objServiceCall.GetJsonNodeValue("output.insure_no")
        strҽ������ = objServiceCall.GetJsonNodeValue("output.insure_pwd")
        str���� = objServiceCall.GetJsonNodeValue("output.card_no")
        str�Ǽ�ʱ�� = objServiceCall.GetJsonNodeValue("output.pati_create_time")
        lng����id = Val(NVL(objServiceCall.GetJsonNodeValue("output.dz_type_id")))
    End If

    Zl_Patisvr_Getinsureinfo = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function Zl_Patisvr_Getdeptfrombad(ByRef str����IDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��λ״����¼��Ŀ���id
    '����: str����IDs :����id�ַ���,�ö��ŷָ�
    '����:��ȡ����ids�ɹ�����True,���򷵻�False
    '����:����
    '����:2019-12-25 19:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strServiceName As String
    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    strServiceName = "Zl_Patisvr_Getdeptfrombad"
    If objServiceCall.CallService(strServiceName, "", , "", glngModul) = False Then Exit Function
    
    'Zl_Patisvr_Getdeptfrombad
    ' output
    ' --  code                  C    1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    ' --  message               C    1  Ӧ����Ϣ
    ' --  dept_ids              C    1  ����ids��������ö��ŷָ�
    str����IDs = objServiceCall.GetJsonNodeValue("output.dept_ids")

    Zl_Patisvr_Getdeptfrombad = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function Zl_Patisvr_Checkdepositno(ByVal lng����ID As Long, ByRef strNo As String, Optional intOcc As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�Ԥ��NO�Ƿ����"���˽����쳣��¼"��
    '����:strNo-Ԥ�����ݺ�
    '        intOcc-���˽����쳣��¼.�����龰
    '����:�����NO�Ŵ���"���˽����쳣��¼"�з���true,���򷵻�False
    '����:����
    '����:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Patisvr_Checkdepositerrorno
    '    input
    '    --   pati_id              N 1 ����ID
    '    --   bill_nos             C 1 ����Ԥ����¼.NO
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", strNo, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Patisvr_Checkdepositerrorno"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
     ' --  output
      '  --    code              N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
       ' --    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
       ' --    bill_nos          C 1 ��Ч��Nos,����ö��ŷָ�
       ' --    occasion          N 1 ���ϣ�1-ҽ�ƿ�����;2-������Ϣ�Ǽǣ����ֻ��һ��NO����Ч��
    strNo = objServiceCall.GetJsonNodeValue("output.bill_nos")
    intOcc = Val(objServiceCall.GetJsonNodeValue("output.occasion"))
    Zl_Patisvr_Checkdepositno = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_Patisvr_Calc_Age(ByVal lng����ID As Long, ByVal str�������� As String, Optional ByVal str�������� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݳ������ڼ��㲡������
    '���:str��������-ָ���ļ������������
    '����:
    '����:���������
    '����:���ϴ�
    '����:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
'    --  input
    '  --    pati_id            N 1 ����ID
    '  --    birthdate          N 0 ��������
    '  --    calc_date          N 0 ��������
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    age                C 1  ����:1�����ڣ�XСʱ[X����],1����1�����ڣ�X��[XСʱ],1����1�����ڣ�X��[X��],1������ͯ�������ޣ�X��[X��],>=��ͯ�������ޣ�X��
    '  --                            ˵��:1�����ڣ���ָ����������24Сʱ��;1�����ڣ���ָ������㣻����7.8�ճ�����8.8�ղ���1��;1�����ڣ�Ҳ�Ƕ�����㡣;�����ڡ�����ָ��<����
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    If IsDate(str��������) Then
        strJson = strJson & "," & GetJsonNodeString("birthdate", str��������, Json_Text)
    End If
    If IsDate(str��������) Then
        strJson = strJson & "," & GetJsonNodeString("calc_date", str��������, Json_Text)
    End If
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "zl_Patisvr_Calc_Age"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    zl_Patisvr_Calc_Age = objServiceCall.GetJsonNodeValue("output.age")
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_Patisvr_GetPatiPhoto(ByVal lng����ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '���:
    '����:
    '����:���ز�����ƬBase64
    '����:���ϴ�
    '����:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
'    --  input
    '  --    pati_id            N 1 ����ID
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    pati_photo         C 1 ����:base64
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Patisvr_GetPatiPhoto"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    zl_Patisvr_GetPatiPhoto = objServiceCall.GetJsonNodeValue("output.pati_photo")
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_Patisvr_SavePatiPhoto(ByVal lng����ID As Long, ByVal strPatiPhoto As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���没����Ƭ
    '���:
    '����:
    '����:���ز�����ƬBase64
    '����:���ϴ�
    '����:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
'    --  input
    '  --    pati_id            N 1 ����ID.
    '  --    pati_photo         C 1 ����:base64
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_photo", strPatiPhoto, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "zl_Patisvr_SavePatiPhoto"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    zl_Patisvr_SavePatiPhoto = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function


Public Function zl_Patisvr_DeletePatiPhoto(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��������Ƭ
    '���:
    '����:
    '����:
    '����:���ϴ�
    '����:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
'    --  input
    '  --    pati_id            N 1 ����ID.
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "zl_Patisvr_DeletePatiPhoto"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    zl_Patisvr_DeletePatiPhoto = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_PatiSvr_GetPatiAddrssInfo(ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
                ByVal str��ַ��� As String, ByRef cllAddrList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���˽ṹ����ַ��Ϣ
    ' ��� : str��ַ���:��ѯ�ĵ�ַ���1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ��Ϊ0ʱ��ʾ��ѯ�������͵ĵ�ַ��Ϣ
    '        ����ö��ŷָ������磺"3,4"
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim cllData As Collection
    
    On Error GoTo ErrHand
    Set cllAddrList = New Collection
    
    '��ȡ���÷���ӿ�
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng��ҳId, Json_num)
    If IsNumeric(str��ַ���) Then
        strJson = strJson & "," & GetJsonNodeString("addr_type", Val(str��ַ���), Json_num)
    Else
        strJson = strJson & "," & GetJsonNodeString("addr_types", str��ַ���, Json_Text)
    End If

    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_GetPatiAddrssInfo", strJson, , , glngModul) = False Then Exit Function
    Set cllAddrList = objServiceCall.GetJsonListValue("output.addr_list")
    
    zl_PatiSvr_GetPatiAddrssInfo = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

