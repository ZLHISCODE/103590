Attribute VB_Name = "mdlSplitService"
Option Explicit

Public Function zlSplitService_AdviceIsExist(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '���ݴ���ҽ��id���ش��ڷ��ü�¼��ҽ��id
    'strInput��ҽ��id,ҽ��id...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", strInput, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_ExseSvr_AdviceIsExist"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_ExseSvr_AdviceIsExist��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.advice_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_AdviceIsExist = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetRemainMoney(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOut As String) As Boolean
    'ȡ�������
    'strInput������id,����id...
    'strOut: ���,������
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
    'Zl_Exsesvr_Getremainmoney
'  ---------------------------------------------------------------------------
'  --input      ��ȡ���˷������
'  --  pati_id                 N  1  ����ID
'  --  pati_pageid             N  1  ��ҳID
'  --  insure_account_balance  N  1  ҽ���˻����
'  --  pati_ids                C  0  ������ҳ�ؼ���Ϣƴ��������ID:��ҳID,....
'  --  query_type              N  0  ��ѯ��ʽ 1-������ѯ������2-������ѯ������������ò�����Ϣ
'  --output
'  --  code                    C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message                 C  1  Ӧ����Ϣ
'  --  remain_money            N     ʣ���
'  --  guarantee_money         N     ������
'  --  expected_money          N     Ԥ�����
'  --  prepay_money            N  0  Ԥ�����
'  --  item_list[]����������������Ϣʱ�ŷ��أ����б���Բ�����
'  --       pati_id            N 1 ����id
'  --       pati_pageid        N 1 ��ҳid
'  --       remain_money       N 1 ʣ���
'  --       guarantee_money    N 1 ������
'  --       pati_type          C 1 ���ò���
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageid", Split(strInput, ",")(1), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_Getremainmoney"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Exsesvr_Getremainmoney��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")

    strOut = objServiceCall.GetJsonNodeValue("output.pati_type")
    strOut = strOut & "," & objServiceCall.GetJsonNodeValue("output.remain_money")
    strOut = strOut & "," & objServiceCall.GetJsonNodeValue("output.guarantee_money")
        
    zlSplitService_GetRemainMoney = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiPageWarnScheme(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOut As String) As Boolean
    'ȡ��������
    'strInput������id,��ҳid...
    'strOut: ��������
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
'  --��ȡ���˻�ȡ���˷������
'  ---------------------------------------------------------------------------
'  --input      ��ȡ���˷������
'  --  pati_id  N  1  ����ID
'  --  pati_pageid  N  1  ��ҳID
'  --output
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  "Ӧ����Ϣ��
'  --  pati_type  C    ���ò�������
'  --  remain_money    N    ʣ���
'  --  guarantee_money   N   ������
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageid", Split(strInput, ",")(1), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatpagewarnscheme"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Cissvr_Getpatpagewarnscheme��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")

    strOut = objServiceCall.GetJsonNodeValue("output.pati_type")
        
    zlSplitService_GetPatiPageWarnScheme = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiPage(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String, Optional ByRef colOutListBaby As Collection, _
    Optional ByVal bln����λ��Ϣ As Boolean, Optional ByRef colOutListBad As Collection) As Boolean
    'ȡ������ҳ��Ϣ
    'strInput������id:��ҳid,...
    '����:
    '   colOutListBad ��λ��Ϣ���ϣ���Ա������ID,����ID,��������,����,�������,��������=Key(_����ID_����ID_����)
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim colTmp As New Collection, colBaby As New Collection
    Dim i As Integer, n As Integer
    Dim colBads As Collection
    
'---------------------------------------------------------------------------
'Zl_Cissvr_Getpatipageinfo
'  --����:��ȡ������ҳ�����Ϣ
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      query_type          C 1 ��ѯ����:0-������Ϣ;1-������Ϣ��չ;2-��ȡ��ҳ
'  --      pati_pageids        C 1 ������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
'  --      is_babyinfo         N 1 �Ƿ����Ӥ����Ϣ:1-����;0-������
'  --      is_transdeptinfo    N 1 �Ƿ����ת����Ϣ:1-����;0-������
'  --      is_lastpage         N 1 �Ƿ�ȡ���һ��סԺ
'  --      pati_natures        C 1 ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�������ŷָ�������Ϊ����
'  --      rgst_id             N 1 �Һ�ID,���ݹҺ�ID��ѯ
'  --      is_badinfo          N 1 �Ƿ������λ��Ϣ:1-����;0-������
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    pati_count            N 1 ��ѯ�Ĳ�����Ϣ����
'  --    page_list[]             1 ������
'  --      pati_id             N 1 ����id
'  --      pati_pageid         N 1 ��ҳid
'  --      pati_name           C 1 ����
'  --      pati_sex            C 1 �Ա�
'  --      pati_age            C 1 ����
'  --      inpatient_num       C 1 סԺ��
'  --      fee_category        C 1 �ѱ�
'  --      mdlpay_mode_name    C 1 ҽ�Ƹ��ʽ����
'  --      mdlpay_mode_code    C 1 ҽ�Ƹ��ʽ����
'  --      pati_bed            C 1 ��ǰ����
'  --      pati_type           C 1 ��������(��ͨ��ҽ��������)
'  --      pati_show_color     N    ����������ɫ
'  --      pati_education      C 1 ѧ��
'  --      ocpt_name           C 1 ְҵ
'  --      country_name        C 1 ����
'  --      pati_marital_cstatus  C 1 ����״��
'  --      pati_nature         N 1 ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
'  --      audit_sign          N 1 ��˱�־:������ҳ.��˱�־
'  --      si_inp_status       N 1 סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)
'  --      pati_wardarea_id    N 1 ��ǰ����id
'  --      pati_wardarea_name  C 1 ��ǰ��������
'  --      pati_dept_id        N 1 ��ǰ����id
'  --      pati_dept_name      C 1 ��ǰ��������
'  --      adta_time           C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --      adtd_time           C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --      insurance_type      N 1 ����
'  --      rgst_id             N 1 �Һ�id
'  --      catalog_date        C 1 ��Ŀ����:yyyy-mm-dd hh24:mi:ss
'  --      in_objective        C 1 סԺĿ��
'  --      reg_name            C 1 �Ǽ���
'  --      reg_date            C 1 סԺ�Ǽ�ʱ��
'  --      pat_rsdpscn         C 1 סԺҽʦ
'  --      pati_desc           C 1 ���˱�ע
'  --      baby_list[]           1 Ӥ����Ϣ��[����]
'  --        pati_id           N 1 ����id
'  --        pati_pageid       N 1 ��ҳid
'  --        baby_num          N 1 Ӥ�����
'  --        baby_name         C 1 Ӥ������
'  --        baby_sex          C 1 Ӥ���Ա�
'  --        baby_date         C 1 ����ʱ��
'  --      trans_list[]        C   ת���б���Ϣ
'  --        start_reason      C 1 ��ʼԭ��
'  --        start_time        C 1 ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --        dept_name         C 1 ��������
'  --      badinfo_list[]              ��λ��Ϣ��[����]
'  --        wardarea_id       N 1 ����id
'  --        wardarea_name     C 1 ��������
'  --        bed_no            C 1 ����
'  --        bed_class_code    C 1 �������
'  --        bed_class_name    C 1 ��������
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageids", strInput, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_babyinfo", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_transdeptinfo", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_lastpage", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_badinfo", IIf(bln����λ��Ϣ, 1, 0), 1)
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatipageinfo"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Cissvr_Getpatipageinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.page_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    Set colOutListBaby = New Collection
    Set colOutListBad = New Collection
    
    'Ӥ�����ݵĲ���ID,��ҳid��ͬ��Ӥ����Ų�ͬ��Ҫô����key��Ҫô�� ����id+��ҳid+Ӥ����� ��Ϊkey
    For n = 1 To colOutlist.count
        Set colBaby = objServiceCall.GetJsonListValue("output.page_list[" & n - 1 & "].baby_list")
        For i = 1 To colBaby.count
            Set colTmp = New Collection
            colTmp.Add colBaby(i)("_pati_id"), "_pati_id"
            colTmp.Add colBaby(i)("_pati_pageid"), "_pati_pageid"
            colTmp.Add colBaby(i)("_baby_num"), "_baby_num"
            colTmp.Add colBaby(i)("_baby_name"), "_baby_name"
            colTmp.Add colBaby(i)("_baby_sex"), "_baby_sex"
            colTmp.Add colBaby(i)("_baby_date"), "_baby_date"
            colOutListBaby.Add colTmp, "_" & colBaby(i)("_pati_id") & "_" & colBaby(i)("_pati_pageid") & "_" & colBaby(i)("_baby_num")
        Next
        
        If bln����λ��Ϣ Then
            '  --      badinfo_list[]              ��λ��Ϣ��[����]
            '  --        wardarea_id       N 1 ����id
            '  --        wardarea_name     C 1 ��������
            '  --        bed_no            C 1 ����
            '  --        bed_class_code    C 1 �������
            '  --        bed_class_name    C 1 ��������
            Set colBads = objServiceCall.GetJsonListValue("output.page_list[" & n - 1 & "].badinfo_list")
            For i = 1 To colBads.count
                Set colTmp = New Collection
                colTmp.Add colOutlist(n)("_pati_id"), "����ID"
                colTmp.Add colBads(i)("_wardarea_id"), "����ID"
                colTmp.Add colBads(i)("_wardarea_name"), "��������"
                colTmp.Add colBads(i)("_bed_no"), "����"
                colTmp.Add colBads(i)("_bed_class_code"), "�������"
                colTmp.Add colBads(i)("_bed_class_name"), "��������"
                'colOutListBad ��λ��Ϣ���ϣ���Ա������ID,����ID,��������,����,�������,��������=Key(_����ID_����ID_����)
                colOutListBad.Add colTmp, "_" & colOutlist(n)("_pati_id") & "_" & colBads(i)("_wardarea_id") & "_" & colBads(i)("_bed_no")
            Next
        End If
    Next
    
    zlSplitService_GetPatiPage = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiName(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal colInput As Collection, _
    ByRef colPati As Collection, Optional ByVal strKey As String, Optional ByVal intQueryType As Integer = 3) As Boolean
    '���ڰ�һ������������ѯ������Ϣ
    'Ŀǰ֧�ֵĲ�ѯ������һ�㶼�ǰ�����һ�ֲ�ѯ������ID������ţ����������￨�ţ�ҽ���ţ�����
    'colInput����ѯ������ϣ�Json��input���ڵ���ΪԪ�ص�KEYֵ������ĳԪ��Ϊ�ձ�ʾ�ýڵ�ֵΪ��
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    'Zl_Patisvr_Getpatiinfo
'  ---------------------------------------------------------------------------
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;2-����
'  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
'  --      query_immune      N 1 �Ƿ����������:1-����;0-������
'  --      query_insurance_pwd C  �Ƿ����ҽ������:1-����;0-������
'  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
'  --        pati_ids        C   ����IDs:����ö���
'  --        pati_name       C   ����:���Դ�%�ֺű������ƥ��
'  --        outpatient_num  C   �����
'  --        inpatient_num   C   סԺ��
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
'  --����      json
'  --output
'  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    pati_list[]                 ������Ϣ�б�
'  --    pati_id             N   1   ����id
'  --    pati_pageid         N   1   ��ҳid��������Ϣ.��ҳID
'  --    pati_name           C   1   ����
'  --    pati_sex            C   1   �Ա�
'  --    pati_age            C   1   ����
'  --    pati_birthdate      C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   1   �ѱ�
'  --    outpatient_num      C   1   �����
'  --    inpatient_num       C   1   סԺ��
'  --    mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����
'  --    mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����
'  --    pati_nation         C   1   ����
'  --    insurance_num       C   1   ҽ����
'  --    pati_idcard         C   1   ���֤��
'  --    vcard_no            C   1   ���￨��
'  --    iccard_no           C   1   Ic����
'  --    health_num          C   1   ������
'  --    inp_times           N   1   סԺ����
'  --    pati_education      C   1   ѧ��
'  --    ocpt_name           C   1   ְҵ
'  --    pati_identity       C   1   ���
'  --    ntvplc_name         C   1   ����
'  --    country_name        C   1   ����
'  --    pati_marital_cstatus    C   1   ����״��
'  --    pat_home_addr           C   1   ��ͥ��ַ
'  --    pat_home_phno           C   1   ��ͥ�绰
'  --    pat_home_postcode   C   1   ��ͥ��ַ�ʱ�
'  --    pati_area           C   1   ����
'  --    pati_birthplace     C   1   �����ص�
'  --    pat_hous_addr       C   1   ���ڵ�ַ
'  --    pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�
'  --    emp_name            C   1   ������λ����
'  --    emp_phno            C   1   ��λ�绰
'  --    emp_postcode        C   1   ��λ�ʱ�
'  --    emp_bank_name       C   1   ��λ������
'  --    emp_bank_accnum     C   1   ��λ�ʺ�
'  --    emp_addr             C   1   ��λ��ַ
'  --    ctt_unit_id         N   1   ��ͬ��λID
'  --    phone_number        C   1   �ֻ���
'  --    pati_bed            C   1   ��ǰ����
'  --    pati_type           C   1   ��������(��ͨ��ҽ��������)
'  --    insurance_type      C   1   ����
'  --    insurance_name      C   1   ��������
'  --    pati_wardarea_id    N   1   ��ǰ����id
'  --    pati_wardarea_name  C   1   ��ǰ��������
'  --    pati_dept_id        N   1   ��ǰ����id
'  --    pati_dept_name      C   1   ��ǰ��������
'  --    adta_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --    contacts_name       C   1   ��ϵ������
'  --    contacts_relation   C   1   ��ϵ�˹�ϵ
'  --    contacts_idcard     C   1   ��ϵ�����֤��
'  --    contacts_addr       C   1   ��ϵ�˵�ַ
'  --    contacts_phno       C   1   ��ϵ�˵绰
'  --    pat_grdn_name       C   1   �໤��
'  --    cert_no_other       C   1   ����֤��
'  --    is_inhspt            C   1   �Ƿ���Ժ:1-��Ժ ;0-����Ժ
'  --    pati_show_color      N   1   ������ʾ��ɫ
'  --    visit_room           C   1   ��������
'  --    visit_statu          N   1   ����״̬
'  --    visit_time           C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    create_time          C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    pati_email           C   1   email
'  --    pati_qq              C   1   qq
'  --    card_captcha         C   1  ����֤��
'  --    insurance_pwd        C       ҽ������
'  --    family_list[]        C   1   ������Ա:���˼���() query_family=1����
'  --        family_id        N   1   ����id  query_family=1
'  --        family_relation  C   1   ��ϵ
'  --    drug_list[]          C   1   ����ҩ���б�    query_drug=1ʱ����
'  --        pat_algc_cadn_id N   1   ����ҩƷID
'  --        pat_algc_cadn    C   1   ����ҩ������
'  --        allergy_info     C   1   ��ÿҩ�ﷴӦ
'  --    immune_list[]        C   1   ���������б�    query_immune=1ʱ����
'  --        vaccinate_time   C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --        vaccinate_name   C   1   ��������
'  --    card_list[]          C   1   ����ҽ�ƿ���Ϣ�б�(��������д����˿����ID�ģ��򷵻ظÿ����Ŀ���Ϣ)  query_card=1ʱ����
'  --        cardtype_id      N   1   ҽ�ƿ����ID
'  --        card_no          C   1   ����
'  --        card_pwd         C   1   ����
'  ---------------------------------------------------------------------------

    '���
    StrJson_In = ""
    If Not IsNull(colInput("pati_id")) Then
        StrJson_In = GetJsonNodeString("pati_id", colInput("pati_id"), 1)
    Else
        StrJson_In = GetJsonNodeString("pati_id", 0, 1)
    End If
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", intQueryType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_card", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_family", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_drug", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_immune", 0, 1)
    
    If IsNull(colInput("pati_id")) Then
        '���ܰ�����һ�ַ�ʽ��ѯ������ţ����������￨�ţ�ҽ���ţ����ţ����֤��
        If Not IsNull(colInput("outpatient_num")) Then strJson_List = strJson_List & "," & GetJsonNodeString("outpatient_num", colInput("outpatient_num"), 0)
        If Not IsNull(colInput("pati_name")) Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_name", colInput("pati_name"), 0)
        If Not IsNull(colInput("pati_vcard_no")) Then strJson_List = strJson_List & "," & GetJsonNodeString("visit_card", colInput("pati_vcard_no"), 0)
        If ExistsColObject(colInput, "insurance_num") Then
            If Not IsNull(colInput("insurance_num")) Then strJson_List = strJson_List & "," & GetJsonNodeString("insurance_num", colInput("insurance_num"), 0)
        End If
        If Not IsNull(colInput("pati_bed")) Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_bed", colInput("pati_bed"), 0)
        If ExistsColObject(colInput, "pati_idcard") Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_idcard", nvl(colInput("pati_idcard")), 0)
        
        strJson_List = strJson_List & "," & GetJsonNodeString("qrspt_statu", 2, 1)
        
        strJson_List = Mid(strJson_List, 2)
    End If
  
    strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
    
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    
    strService = "Zl_Patisvr_Getpatiinfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Patisvr_Getpatiinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colPati = objServiceCall.GetJsonListValue("output.pati_list", strKey)
    
    If colPati Is Nothing Then Exit Function
    
    zlSplitService_GetPatiName = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiId(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutPut As String) As Boolean
    'ȡ������Ϣ������סԺ�ţ�����+����
    'strInput��סԺ�� �� ����id|���ţ������Ƿ��С�|��������
    'strOutPut��������Ϣ������ID
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
'  ---------------------------------------------------------------------------
'Zl_Cissvr_Getpatiid
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --input
'  --   wardarea_id          N 1 ��ǰ����id
'  --   pati_bed             C 1 ��ǰ����
'  --   inpatient_num        C 1 סԺ��
'  --output
'  --    code                N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    pati_id             N 1 ����ID:δ�ҵ�ʱҲ�ɹ�������0
'  --    pati_pageid         N   ��ҳID
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    
    If InStr(strInput, "|") > 0 Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("wardarea_id", Split(strInput, "|")(0), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_bed", Split(strInput, "|")(1), 0)
    Else
        StrJson_In = StrJson_In & "" & GetJsonNodeString("inpatient_num", strInput, 0)
    End If
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatiid"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Patisvr_Getpatiinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������
    strOutPut = objServiceCall.GetJsonNodeValue("output.pati_id")
    
    zlSplitService_GetPatiId = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetBillOperControls(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutPut As String) As Boolean
    '��ȡ���ݲ�����������
    'strInput����Աid
    'strOutPut���Ƿ����|ʱ������|�Ƿ��������˵���|�������
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
    'Zl_Exsesvr_Getbillopercontrols
'  ---------------------------------------------------------------------------
'  --����:��ȡ���ݲ�����������
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    bill_type  N  1  ��������:1-�Һŵ���,2-�շѵ�,3-���۵�,4-�������,5-סԺ����,6-Ԥ����,7-���ʵ���,8-���￨,9-����
'  --    operator_id  N  1  ��ԱID
'
'  --����: Json_Out,��ʽ����
'  --   output
'  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message  C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    is_exist    N  1  ���ڿ�������:1-����;0-������
'  --    time_limit  N  1  0(NULL)-������,n-n����
'  --    other_bill  N  1  �Ƿ�������������ݽ��в���
'  --    uplimit_money  N  1  �������
'
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    
    StrJson_In = StrJson_In & "" & GetJsonNodeString("bill_type", 9, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_id", strInput, 1)
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_Getbillopercontrols"

    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Exsesvr_Getbillopercontrols��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    strOutPut = objServiceCall.GetJsonNodeValue("output.is_exist") & "|" & _
        objServiceCall.GetJsonNodeValue("output.time_limit") & "|" & _
        objServiceCall.GetJsonNodeValue("output.other_bill") & "|" & _
        objServiceCall.GetJsonNodeValue("output.uplimit_money")
    
    zlSplitService_GetBillOperControls = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetPatiByRange(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '����Χ���Ҳ�����Ϣ
    'strInput�� ��ѯ������Ŀǰ��������ID����
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    'Zl_Cissvr_Getpatpageinfbyrange
'  ---------------------------------------------------------------------------
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    query_type          N 1 ��ѯ����:0-����;1-������չ
'  --    wararea_ids         C   ����ids:����ö���
'  --    dept_ids            C   ����IDs:���ö���
'  --    pati_ids            C   ����ids:����ö��ŷ���
'  --    pati_pageIds        C   ��ҳIDs:����id:��ҳid,��
'  --    adta_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adta_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   �ѱ�
'  --    inp_status          N   סԺ״̬:0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ
'  --    pati_natures        C   �������ʣ�����ö��ŷ�0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�NULL-��ʾ������
'  --    pati_name           C   ����:���Դ�%�ֺű������ƥ��
'  --    nodeno              C   վ����
'  --    change_dept_pati    N   �Ƿ��ѯת�Ʋ���
'  --����      json
'  --output
'  -- code                   N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  -- message                C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --   page_list[]          ������  ��  ��
'  --    pati_id             N    ����id  ��  ��
'  --    pati_pageid         N    ��ҳid  ��  ��
'  --    pati_name           C    ����  ��  ��
'  --    pati_sex            C    �Ա�  ��  ��
'  --    pati_age            C    ����  ��  ��
'  --    inpatient_num       C    סԺ��  ��  ��
'  --    pati_bed            C    ��Ժ����  ��  ��
'  --    insurance_type      N    ����  ��  ��
'  --    fee_category        C    �ѱ�  ��  ��
'  --    pati_type           C    ��������(��ͨ,ҽ��,����)  ��  ��
'  --    adta_time           C    ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
'  --    adtd_time           C    ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
'  --    si_inp_status       N    סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)  ��  ��
'  --    pati_nature         N    ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
'  --    pati_wardarea_id    N    ��ǰ����id
'  --    pati_wardarea_name  C    ��ǰ��������
'  --    pati_dept_id        N    ��ǰ����id
'  --    pati_dept_name      C    ��ǰ��������
'  --    mdlpay_mode_name    C    ҽ�Ƹ��ʽ����
'  --    mdlpay_mode_code    C    ҽ�Ƹ��ʽ����
'  --    pat_rsdpscn         C    סԺҽʦ
'  --    pati_desc           C    ���˱�ע
'  --    catalog_date        C    ��Ŀ����:yyyy-mm-dd hh24:mi:ss
'  --    create_pati         C    �Ǽ���
'  --    in_objective     C    סԺĿ��
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
'  --    query_type          N 1 ��ѯ����:0-����;1-������չ
'  --    wararea_ids         C   ����ids:����ö���
'  --    dept_ids            C   ����IDs:���ö���
'  --    pati_ids            C   ����ids:����ö��ŷ���
'  --    pati_pageIds        C   ��ҳIDs:����id:��ҳid,��
'  --    adta_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adta_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   �ѱ�
'  --    inp_status          N   סԺ״̬:0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ
'  --    pati_natures        C   �������ʣ�����ö��ŷ�0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�NULL-��ʾ������
'  --    pati_name           C   ����:���Դ�%�ֺű������ƥ��
'  --    nodeno              C   վ����
'  --    change_dept_pati    N   �Ƿ��ѯת�Ʋ���
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("wararea_ids", Val(strInput), 0)

    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatpageinfbyrange"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Cissvr_Getpatpageinfbyrange��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.page_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetPatiByRange = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetPati(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    'ȡ������Ϣ
    'strInput����Ŀ����;��Ŀ����
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
        
'  ---------------------------------------------------------------------------
'Zl_Patisvr_Getpatiinfo
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
'  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
'  --      query_immune      N 1 �Ƿ����������:1-����;0-������
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
'  --        pati_bed        C   ����
'  --����      json
'  --output
'  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    pati_list[]                 ������Ϣ�б�
'  --    pati_id             N   1   ����id
'  --    pati_pageid         N   1   ��ҳid��������Ϣ.��ҳID
'  --    pati_name           C   1   ����
'  --    pati_sex            C   1   �Ա�
'  --    pati_age            C   1   ����
'  --    pati_birthdate      C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   1   �ѱ�
'  --    outpatient_num      C   1   �����
'  --    inpatient_num       C   1   סԺ��
'  --    mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����
'  --    mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����
'  --    pati_nation         C   1   ����
'  --    insurance_num       C   1   ҽ����
'  --    pati_idcard         C   1   ���֤��
'  --    vcard_no            C   1   ���￨��
'  --    iccard_no           C   1   Ic����
'  --    health_num          C   1   ������
'  --    inp_times           N   1   סԺ����
'  --    pati_education      C   1   ѧ��
'  --    ocpt_name           C   1   ְҵ
'  --    pati_identity       C   1   ���
'  --    ntvplc_name         C   1   ����
'  --    country_name        C   1   ����
'  --    pati_marital_cstatus    C   1   ����״��
'  --    pat_home_addr           C   1   ��ͥ��ַ
'  --    pat_home_phno           C   1   ��ͥ�绰
'  --    pat_home_postcode   C   1   ��ͥ��ַ�ʱ�
'  --    pati_area           C   1   ����
'  --    pati_birthplace     C   1   �����ص�
'  --    pat_hous_addr       C   1   ���ڵ�ַ
'  --    pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�
'  --    emp_name            C   1   ������λ����
'  --    emp_phno            C   1   ��λ�绰
'  --    emp_postcode        C   1   ��λ�ʱ�
'  --    emp_bank_name       C   1   ��λ������
'  --    emp_bank_accnum     C   1   ��λ�ʺ�
'  --    emp_addr             C   1   ��λ��ַ
'  --    ctt_unit_id         N   1   ��ͬ��λID
'  --    phone_number        C   1   �ֻ���
'  --    pati_bed            C   1   ��ǰ����
'  --    pati_type           C   1   ��������(��ͨ��ҽ��������)
'  --    insurance_type      C   1   ����
'  --    insurance_name      C   1   ��������
'  --    pati_wardarea_id    N   1   ��ǰ����id
'  --    pati_wardarea_name  C   1   ��ǰ��������
'  --    pati_dept_id        N   1   ��ǰ����id
'  --    pati_dept_name      C   1   ��ǰ��������
'  --    adta_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --    contacts_name       C   1   ��ϵ������
'  --    contacts_relation   C   1   ��ϵ�˹�ϵ
'  --    contacts_idcard     C   1   ��ϵ�����֤��
'  --    contacts_addr       C   1   ��ϵ�˵�ַ
'  --    contacts_phno       C   1   ��ϵ�˵绰
'  --    pat_grdn_name       C   1   �໤��
'  --    cert_no_other       C   1   ����֤��
'  --    is_inhspt            C   1   �Ƿ���Ժ:1-��Ժ ;0-����Ժ
'  --    pati_show_color      N   1   ������ʾ��ɫ
'  --    visit_room           C   1   ��������
'  --    visit_statu          N   1   ����״̬
'  --    visit_time           C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    create_time          C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    pati_email           C   1   email
'  --    pati_qq              C   1   qq
'  --    card_captcha         C   1  ����֤��
'  --    family_list[]        C   1   ������Ա:���˼���() query_family=1����
'  --        family_id        N   1   ����id  query_family=1
'  --        family_relation  C   1   ��ϵ
'  --    drug_list[]          C   1   ����ҩ���б�    query_drug=1ʱ����
'  --        pat_algc_cadn_id N   1   ����ҩƷID
'  --        pat_algc_cadn    C   1   ����ҩ������
'  --        allergy_info     C   1   ��ÿҩ�ﷴӦ
'  --    immune_list[]        C   1   ���������б�    query_immune=1ʱ����
'  --        vaccinate_time   C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --        vaccinate_name   C   1   ��������
'  --    card_list[]          C   1   ����ҽ�ƿ���Ϣ�б�(��������д����˿����ID�ģ��򷵻ظÿ����Ŀ���Ϣ)  query_card=1ʱ����
'  --        cardtype_id      N   1   ҽ�ƿ����ID
'  --        card_no          C   1   ����
'  --        card_pwd         C   1   ����
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
'  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
'  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
'  --      query_immune      N 1 �Ƿ����������:1-����;0-������
'  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
'  --        pati_ids        C   ����IDs:����ö���
    
    '���
    StrJson_In = ""
    
    If Split(strInput, ";")(0) = "����id" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ";")(1), 1)
    Else
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", 0, 1)
    End If
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", 3, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_card", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_family", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_drug", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_immune", 0, 1)
    
    Select Case Split(strInput, ";")(0)
        Case "����ids"
            strJson_List = GetJsonNodeString("pati_ids", Split(strInput, ";")(1), 0)
        Case "ҽ����"
            strJson_List = GetJsonNodeString("insurance_num", Split(strInput, ";")(1), 0)
    End Select
    
    If strJson_List <> "" Then
        strJson_List = strJson_List & "," & GetJsonNodeString("qrspt_statu", 2, 1)
    End If
    strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
    
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    strService = "Zl_Patisvr_Getpatiinfo"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Patisvr_Getpatiinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.pati_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetPati = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetCardTypes(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection) As Boolean
    '��ȡҽ�ƿ��������
    'strInput��Ŀǰ������������ѯ���е�ҽ�ƿ����
    'strOutPut������ID�����룬����
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    'Zl_Patisvr_Getcardtypes
'  ---------------------------------------------------------------------------
'  --����:��ȡҽ�ƿ��������
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      cardtype_id          N   �����id:NULL��ʾ���������ID����
'  --      query_type           N 1 ��ѯ����:0-������Ϣ;1-������Ϣ(����:id,���룬����,���ų���,ǰ׺�ı�,�Ƿ�����,���㷽ʽ,�Ƿ�ȫ��,�Ƿ�����)
'  --      cert_cardtype        N   ֻ��ȡ֤����Ϊ������ҽ�ƿ����:1-ֻ��ȡ�Ƿ�֤��=1��ҽ�ƿ�;0-ȫ����ȡ
'  --      dffective_cardtype   N   ֻ��ȡ��Ч�Ŀ����:1-ֻ��ȡ��Ч�Ŀ����;0-ȫ����ȡ
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                   N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message                C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    type_list[]            C   1   ֧�ֵĿ�����б�
'  --        cardtype_id        N   1   ID
'  --        cardtype_code      C   1   ����
'  --        cardtype_name      C   1   ����
'  --        cardtype_stname    C   1   ����
'  --        prefix_text        C   1   ǰ׺�ı�
'  --        cardno_len         N   1   ���ų���
'  --        default            N   1   ȱʡ��־
'  --        fixed              N   1   �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
'  --        strict             N   1   �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
'  --        self_make          N   1   �Ƿ�����:1-�ǵ�;0-����
'  --        exist_account      N   1   �Ƿ�����ʻ�:1-�����ʻ�;0-�������˻�
'  --        allow_return_cash  N   1   �Ƿ�����:1-����;0-������
'  --        must_all_return    N   1   �Ƿ�ȫ��:1-����ȫ��;0-��������
'  --        component          C   1   ����
'  --        memo               C   1   ��ע
'  --        spec_item          C   1   �ض���Ŀ
'  --        blnc_mode          C   1   ���㷽ʽ
'  --        blnc_nature        N   1   ��������
'  --        cardno_pwdtxt      C   1   ��������:���Ŵӵڼ�λ���ڼ�λ��ʾ����,��ʽΪ:S-N:S��ʾ�ӵڼ�λ��ʼ,���ڼ�λ����.����:3-10,��ʾ��3λ��10λ������*��ʾ:12********3323��Ҫ����Ӧ��ͬ����ҽ�ƿ�
'  --        allow_repeat_use   N   1   �Ƿ��ظ�ʹ��:1-����;0-������
'  --        enabled            N   1   �Ƿ�����:1-������;0-δ����
'  --        pwd_len            N   1   ���볤��
'  --        pwd_len_limit      N   1   ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
'  --        pwd_rule           N   1   �������:��-���ֺ��ַ����;1-��Ϊ�������
'  --        allow_vaguefind    N   1   �Ƿ�ģ������:1-֧��ģ������;0-��֧��
'  --        pwd_require        N   1   ������������:0-������;1-������,����;2-�������ֹ;ȱʡΪ������
'  --        default_pwd        N   1   �Ƿ�ȱʡ����:1-�����֤��N(�����볤��Ϊ׼)λ��Ϊȱʡ����;0-��ȱʡ����
'  --        allow_makecard     N   1   �Ƿ��ƿ�:1-��;0-��
'  --        allow_sendcard     N   1   �Ƿ񷢿�:1-��;0-��
'  --        allow_writcard     N   1   �Ƿ�д��:1-��;0-��
'  --        insurance_type     N   1   ����
'  --        insurance_name     C   1   ��������
'  --        sendcard_nature    N   1   ��������:0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0
'  --        allow_transfer     N   1   �Ƿ�ת�ʼ�����:1-֧��ת�ʼ�����;0-��֧��
'  --        readcard_nature    C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
'  --        keyboard_mode      N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
'  --        advsend_buildqrcode N   1   �Ƿ�ҽ�����͵����������ɽӿ�:1-���͵������ɶ�ά��ӿ�;0-������
'  --        holding_pay         N   1   �Ƿ�ֿ�����:1-��;0-��
'  --        cert_cardtype       N   1   �Ƿ�֤�����͵�ҽ�ƿ�:0-���ǣ�1-��
'  --        verfycard           N   1   �Ƿ��˿��鿨
'  --        sendcard_sign       N   1   ��������:0��NULL-����ʱ�����ű���ﵽ���ų���;1-����ʱ��������С�ڵ��ڿ��ų���,����ʱ��С�ڿ��ų���ʱ������ʾ����Ա;2-����ʱ��������С�ڵ��ڿ��ų���,С��ʱ����ʾ����Ա��
'  --        enterkey_enabled    N   1   �豸�Ƿ����ûس�:ҽ�ƿ���Ӧ��ˢ���豸�Ƿ������˻س�����������˻س����򿨺ų���Ĭ������һλ�����λس�
'  --        def_return_cash     N   1   �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
'  --        balalone            N   1   �Ƿ��������:1-��������;0-�Ƕ�������
'  --        discern_rule        N   1   ����ʶ�����:1-ȫ��ת��Ϊ��д;0-�����ִ�Сд
'  --        def_valid_time      C   1   ȱʡ��Чʱ��:NULLʱ����ʾ������;�ǿ�ʱ����ʽΪ:ʱ��+��λ(�죬��),���磺3��,3��
'  --        scanpay             N   1   �Ƿ�֧��ɨ�븶:�Ƿ�֧��ɨ�븶,֧��ʱ������á�zlReadQRCode������
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle

'  --      cardtype_id          N   �����id:NULL��ʾ���������ID����
'  --      query_type           N 1 ��ѯ����:0-������Ϣ;1-������Ϣ(����:id,���룬����,���ų���,ǰ׺�ı�,�Ƿ�����,���㷽ʽ,�Ƿ�ȫ��,�Ƿ�����)
'  --      cert_cardtype        N   ֻ��ȡ֤����Ϊ������ҽ�ƿ����:1-ֻ��ȡ�Ƿ�֤��=1��ҽ�ƿ�;0-ȫ����ȡ
'  --      dffective_cardtype   N   ֻ��ȡ��Ч�Ŀ����:1-ֻ��ȡ��Ч�Ŀ����;0-ȫ����ȡ

    '���
    StrJson_In = ""
    
    StrJson_In = StrJson_In & "" & GetJsonNodeString("cardtype_id", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", 1, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("cert_cardtype", 0, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("dffective_cardtype", 0, 0)

    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_PatiSvr_GetCardTypes"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_PatiSvr_GetCardTypes��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.type_list")
    
    If colOutlist Is Nothing Then zlSplitService_GetCardTypes = False: Exit Function
    
    zlSplitService_GetCardTypes = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallAccountInsert(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strAccInsert As String) As Boolean
    '������ȡ�������ô���ѵȣ���������/סԺ��¼
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String, strJson_detail As String
    Dim varList As Variant, varNos As Variant
    Dim strPrePati As String, strPati As String, strPatiJson As String
    
    If strAccInsert = "" Then Exit Function
    '����Zl_Exsesvr_Newbill
    ' ---------------------------------------------------------------------------
    '  --���ܣ����ﲡ��סԺ���˷���ҽ�����ɷ��õ���
    '  --��Σ�Json_In:��ʽ
    '  --input
    '  --  pati_list[] �����б���������ʱ�����޸ýڵ�
    '  --    billtype                                            N 1 ����,1-�շѵ���2-���ʵ�
    '  --    pati_source                                         N 1 ��Դ��1-���2-סԺ
    '  --    pati_id                                             N 1 ����id
    '  --    pati_pageid                                         N 1 ��ҳid
    '  --    baby_num                                            N 1 Ӥ����
    '  --    sgin_no                                             N 1 ��ʶ�ţ�����ţ�סԺ��
    '  --    bed_num                                             C 1 ����
    '  --    pati_name                                           C 1 ����
    '  --    pati_sex                                            C 1 �Ա�
    '  --    pati_age                                            C 1 ����
    '  --    fee_category                                        C 1 �ѱ�
    '  --    overtime_sign                                       N 1 �Ӱ��־
    '  --    pati_deptid                                         N 1 ���˿���id
    '  --    pati_wardarea_id                                    N 1 ���˲���id
    '  --    operator_name                                       C 1 ����Ա����
    '  --    operator_code                                       C 1 ����Ա���
    '  --    outpati_tag                                         N 1 �����־
    '  --    rgst_id                                             N 1 ����id
    '  --    emg_sign                                            N 1 �Ƿ���
    '  --    item_list[]��ϸ�б�
    '  --        fee_id                                        N 1 ����id
    '  --        fee_no                                        C 1 No
    '  --        serial_num                                    N 1 ���
    '  --        charge_tag                                    N 1 ����
    '  --        placer                                        C 1 ������
    '  --        plcdept_id                                    N 1 ��������id
    '  --        sub_serial_num                                N 1 ��������
    '  --        fitem_id                                      N 1 �շ�ϸĿid
    '  --        item_type                                     C 1 �շ����
    '  --        unit                                          C 1 ���㵥λ
    '  --        pharmacy_window                               C 1 ��ҩ����
    '  --        packages_num                                  N 1 ����
    '  --        send_num                                      N 1 ����
    '  --        ext_mark                                      N 1 ���ӱ�־
    '  --        exe_deptid                                    N 1 ִ�в���id
    '  --        price_ftrnum                                  N 1 �۸񸸺�
    '  --        income_item_id                                N 1 ������Ŀid
    '  --        receipt_name                                  C 1 �վݷ�Ŀ
    '  --        price                                         N 1 ��׼����
    '  --        fee_amrcvb                                    N 1 Ӧ�ս��
    '  --        fee_ampaib                                    N 1 ʵ�ս��
    '  --        happen_time                                   C 1 ����ʱ��
    '  --        create_time                                   C 1 �Ǽ�ʱ��
    '  --        memo                                          C 1 ����ժҪ
    '  --        order_id                                      N 1 ҽ�����
    '  --        exe_properties                                N 1 ִ������
    '  --        decoction_method                              C 1 �巨
    '  --        morphology                                    C 1 ��ҩ��̬
    '  --        bakstuff_batch                                N 1 ����
    '  --        insurance                                     N 1 ������Ŀ��
    '  --        insure_id                                     N 1 ���մ���id
    '  --        insure_code                                   C 1 ���ձ���
    '  --        fee_type                                      C 1 ��������
    '  --        si_manp_money                                 N 1 ͳ����
    '  --        synchro                                       N 1 ����ͬ����־
    '  --        effective_time                                N 1 ��Ч
    '  --        receipt_issecret                              N 1 ����
    '  --        takedept_id                                   N 1 ��ҩ����id
    '  --        group_id                                      N 0 ҽ��С��id
    '  --����: Json_Out,��ʽ����
    '  --output
    '  --  code                                                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --  message                                             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  ---------------------------------------------------------------------------
    
    '����/סԺ���ʼ�¼_insert
    '����rcp_no 0���ü�¼.NO, pati_id 1����ID, sgin_no 2��־��, pati_name 3����, pati_sex 4�Ա�, pati_age 5����,
             '      fee_category 6�ѱ�, pati_wardarea_id 7���˲���id, pati_deptid 8���˿���id, bill_deptid  9��������id,
             '      placer 10������, operator_name 11����Ա����, operator_code 12����Ա���,
             '      happen_time 13����ʱ��, create_time 14�Ǽ�ʱ��, outpati_flag 15�����־
                         
    '��ϸ��serial_num 16���, baby_num 17Ӥ����, fitem_id 18�շ�ϸĿid, item_type 19�շ����, unit 20���㵥λ
    '      packages_num 21����, send_num 22����, income_item_id 23������Ŀid, receipt_name 24�վݷ�Ŀ, price 25��׼����, fee_amrcvb 26Ӧ�ս��,
    '      fee_ampaib 27ʵ�ս��, exe_deptid 28ִ�в���id, pati_source 29������Դ, pati_pageid 30��ҳid, bed_num 31����
    
    varList = Split(strAccInsert, "|")
    For i = 0 To UBound(varList)
        varNos = Split(varList(i), ",")
        
        If strPrePati <> varNos(1) Then
            If strJson_List <> "" Then
                strPatiJson = IIf(strPatiJson = "", "", strPatiJson & ",") & "{" & strPati & ",""item_list"":[" & strJson_List & "]}"
                strJson_List = ""
            End If
            
            strPrePati = varNos(1)
            
            'ȡ����Ϣ
            strPati = ""
            strPati = strPati & "" & GetJsonNodeString("billtype", 2, 1)
            strPati = strPati & "," & GetJsonNodeString("pati_source", varNos(29), 1)
            strPati = strPati & "," & GetJsonNodeString("pati_id", varNos(1), 1)
            strPati = strPati & "," & GetJsonNodeString("pati_pageid", varNos(30), 1)
            strPati = strPati & "," & GetJsonNodeString("baby_num", varNos(17), 1)
      
            strPati = strPati & "," & GetJsonNodeString("sgin_no", varNos(2), 1)
            strPati = strPati & "," & GetJsonNodeString("bed_num", varNos(31), 0)
            strPati = strPati & "," & GetJsonNodeString("pati_name", varNos(3), 0)
            strPati = strPati & "," & GetJsonNodeString("pati_sex", varNos(4), 0)
            strPati = strPati & "," & GetJsonNodeString("pati_age", varNos(5), 0)
            
            strPati = strPati & "," & GetJsonNodeString("fee_category", IIf(varNos(6) = "", "��ͨ", varNos(6)), 0)
            strPati = strPati & "," & GetJsonNodeString("overtime_sign", 0, 1)
            strPati = strPati & "," & GetJsonNodeString("pati_deptid", varNos(8), 1)
            strPati = strPati & "," & GetJsonNodeString("pati_wardarea_id", varNos(7), 1)
            strPati = strPati & "," & GetJsonNodeString("operator_name", varNos(11), 0)
    
            strPati = strPati & "," & GetJsonNodeString("operator_code", varNos(12), 0)
            strPati = strPati & "," & GetJsonNodeString("outpati_tag", varNos(15), 1)
            strPati = strPati & "," & GetJsonNodeToNull("rgst_id")
            strPati = strPati & "," & GetJsonNodeToNull("emg_sign")
        End If
        
        'ȡ��ϸ��Ϣ
        strJson_detail = ""
        '  --        fee_id                                        N 1 ����id
        '  --        fee_no                                        C 1 No
        '  --        serial_num                                    N 1 ���
        '  --        charge_tag                                    N 1 ����
        '  --        placer                                        C 1 ������
        strJson_detail = strJson_detail & "" & GetJsonNodeToNull("fee_id")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fee_no", varNos(0), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("serial_num", varNos(16), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("charge_tag")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("placer", "", 0)
        
        '  --        plcdept_id                                    N 1 ��������id
        '  --        sub_serial_num                                N 1 ��������
        '  --        fitem_id                                      N 1 �շ�ϸĿid
        '  --        item_type                                     C 1 �շ����
        '  --        unit                                          C 1 ���㵥λ
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("plcdept_id")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("sub_serial_num")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fitem_id", varNos(18), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("item_type", varNos(19), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("unit", IIf(varNos(20) = "", "��", varNos(20)), 0)
        
        '  --        pharmacy_window                               C 1 ��ҩ����
        '  --        packages_num                                  N 1 ����
        '  --        send_num                                      N 1 ����
        '  --        ext_mark                                      N 1 ���ӱ�־
        '  --        exe_deptid                                    N 1 ִ�в���id
        strJson_detail = strJson_detail & "," & GetJsonNodeString("pharmacy_window", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("packages_num", varNos(21), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("send_num", varNos(22), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("ext_mark")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("exe_deptid", varNos(28), 1)
                
        '  --        price_ftrnum                                  N 1 �۸񸸺�
        '  --        income_item_id                                N 1 ������Ŀid
        '  --        receipt_name                                  C 1 �վݷ�Ŀ
        '  --        price                                         N 1 ��׼����
        '  --        fee_amrcvb                                    N 1 Ӧ�ս��
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("price_ftrnum")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("income_item_id", varNos(23), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("receipt_name", varNos(24), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("price", varNos(25), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fee_amrcvb", varNos(26), 1)
        
        '  --        fee_ampaib                                    N 1 ʵ�ս��
        '  --        happen_time                                   C 1 ����ʱ��
        '  --        create_time                                   C 1 �Ǽ�ʱ��
        '  --        memo                                          C 1 ����ժҪ
        '  --        order_id                                      N 1 ҽ�����
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fee_ampaib", varNos(27), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("happen_time", varNos(13), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("create_time", varNos(14), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("memo", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("order_id")
        
        '  --        exe_properties                                N 1 ִ������
        '  --        decoction_method                              C 1 �巨
        '  --        morphology                                    C 1 ��ҩ��̬
        '  --        bakstuff_batch                                N 1 ����
        '  --        insurance                                     N 1 ������Ŀ��
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("exe_properties")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("decoction_method", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("morphology", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("bakstuff_batch")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("insurance")
        
        '  --        insure_id                                     N 1 ���մ���id
        '  --        insure_code                                   C 1 ���ձ���
        '  --        fee_type                                      C 1 ��������
        '  --        si_manp_money                                 N 1 ͳ����
        '  --        synchro                                       N 1 ����ͬ����־
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("insure_id")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("insure_code", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fee_type", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("si_manp_money")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("synchro")
        
        '  --        effective_time                                N 1 ��Ч
        '  --        receipt_issecret                              N 1 ����
        '  --        takedept_id                                   N 1 ��ҩ����id
        '  --        group_id                                      N 0 ҽ��С��id
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("effective_time")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("receipt_issecret")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("takedept_id")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("group_id")
        
        strJson_List = IIf(strJson_List = "", "", strJson_List & ",") & "{" & strJson_detail & "}"
    Next
    
    If strJson_List <> "" Then
        strPatiJson = IIf(strPatiJson = "", "", strPatiJson & ",") & "{" & strPati & ",""item_list"":[" & strJson_List & "]}"
    End If
    StrJson_In = "{""input"":{""pati_list"":[" & strPatiJson & "]}}"
        
    '���÷���
    If objServiceCall.CallService("Zl_Exsesvr_Newbill", StrJson_In, , "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Newbill��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    If strJson_Out = "0" Then
        MsgBox objServiceCall.GetJsonNodeValue("output.message"), vbInformation, gstrSysName
        Exit Function
    End If

    zlSplitService_CallAccountInsert = True
End Function



Public Function zlSplitService_CallAccountDel_Check(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef strMsg As String, Optional ByRef intִ��״̬ As Integer = 1) As Boolean
    '����/סԺ��¼���˼��
    'strInput:������Ҫ����θ�ʽ��no,�ѽ��ֹ����(1),ҽ����ֹ��������(0),����״̬(1),������Դ|���,��������;���,��������...|����id,�ѷ�����;����id,�ѷ�����|����id,��˱�־,סԺ״̬,������Ŀ����;����id,��˱�־,סԺ״̬,������Ŀ����...
    '         һ����"|"������1������Ķ�����������";"��","�ָ�
    Dim arrPart As Variant
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String, strJson_Part As String
    Dim strService As String
    Dim strPart1 As String, strPart2 As String, strPart3 As String, strPart4 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String, strJson_In_Part4 As String
    
On Error GoTo errHandle
    
    '����Zl_ExseSvr_DelBill_Check ����/סԺ��¼ͨ�÷���
'  ---------------------------------------------------------------------------
'  --���ܣ����ָ������ָ�����н�������
'  --��Σ�Json_In:��ʽ
'  --  input
'  --      fee_no                  C   1   ���õ��ݺ�
'  --      fee_bill_type           N   1   ��������:2-������ʵ�,3-�Զ����ʵ�
'  --      balance_ban_writeoffs   N   1   �ѽ��ֹ����:����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
'  --      part_ban_writeoffs      N   1   ��ֹ��������:1-������0-����
'  --      fee_origin              N   1   ������Դ��1-������ʣ�2-סԺ���ʣ�
'  --      item_list[]             ���������б�
'  --          serial_num          N   1   ���
'  --          quantity            N   1   ��������(Ϊ��ʱ�������ֱ������)
'  --      excute_list[]           ������ִ���б�(ҩƷ�����ķ���),��ʹ��ִ����Ϊ0ҲҪ����
'  --          fee_id              N   1   ����ID
'  --          sended_num          N   1   �ѷ�����
'  --      advice_excute_list[]    ������ִ���б�(ҽ������),��ʹ��ִ����Ϊ0ҲҪ����
'  --          advice_id           N   1   ҽ��ID
'  --          fee_item_id         N   1   �շ�ϸĿID
'  --          execute_num         N   1   ��ִ����
'  --      pati_list[]             ������Ϣ���������Щ���˵ķ���
'  --          pati_id             N   1   ����ID
'  --          fee_audit_status    N   1   ������˱�־:0���-δ���;1-����˻�ʼ���(��ϲ���:������˷�ʽ������);2-������,��Ͻ���Ȩ��[��ֹδ��˲��˽���]���й������
'  --          si_inp_status       N   1   סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ
'  --          catalog_date        C   0   ������Ŀ���ڣ�yyyy-mm-dd hh24:mi:ss
'  --����: Json_Out,��ʽ����
'  --  output
'  --      code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --      message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --      item_list[]                         ���������б�
'  --          serial_num          N   1   ���
'  --          quantity            N   1   ��������
'  --          execute_tag         N   1   ִ��״̬��0-δִ��;1-��ִ��;2-����ִ��
'  ---------------------------------------------------------------------------

    If strInput = "" Then zlSplitService_CallAccountDel_Check = True: Exit Function
    
    strJson_List = ""
    
    arrPart = Split(strInput, "|")
    
    strPart1 = arrPart(0)
    strPart2 = arrPart(1)
    strPart3 = arrPart(2)
    strPart4 = arrPart(3)
        
    strJson_In_Part1 = ""
    strJson_In_Part1 = strJson_In_Part1 & "" & GetJsonNodeString("fee_no", Split(strPart1, ",")(0), 0)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("fee_bill_type", 2, 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("balance_ban_writeoffs", Split(strPart1, ",")(1), 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("part_ban_writeoffs", Split(strPart1, ",")(2), 1)
'    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("oper_type", Split(strPart1, ",")(3), 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("fee_origin", Split(strPart1, ",")(4), 1)

    strJson_List = ""
    arrPart = Split(strPart2, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("serial_num", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("quantity", Split(arrPart(i), ",")(1), 1)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_In_Part2 = ",""item_list"":[" & strJson_List & "]"
    
    strJson_List = ""
    arrPart = Split(strPart3, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("fee_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("sended_num", Split(arrPart(i), ",")(1), 1)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_In_Part3 = ",""excute_list"":[" & strJson_List & "]"
    
    strJson_List = ""
    arrPart = Split(strPart4, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("pati_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("fee_audit_status", Split(arrPart(i), ",")(1), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("si_inp_status", Split(arrPart(i), ",")(2), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("catalog_date", Split(arrPart(i), ",")(3), 0)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_In_Part4 = ",""pati_list"":[" & strJson_List & "]"
    
    '����
    StrJson_In = "{""input"":{" & strJson_In_Part1 & strJson_In_Part2 & strJson_In_Part3 & strJson_In_Part4 & "}}"
    
    strService = "Zl_ExseSvr_DelBill_Check"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_ExseSvr_DelBill_Check��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        strMsg = objServiceCall.GetJsonNodeValue("output.message")
        zlSplitService_CallAccountDel_Check = False
        Exit Function
    End If
    
    intִ��״̬ = objServiceCall.GetJsonNodeValue("output.item_list[0].execute_tag")

    zlSplitService_CallAccountDel_Check = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallCancelAccCheck(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String) As Boolean
    '���˺˲�
    'strInput:�˲���,�˲�ʱ��,�������|����id,����ʱ��;����id,����ʱ��;...
    Dim arrPart As Variant
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String, strJson_Part As String
    Dim strService As String
    Dim strPart1 As String, strPart2 As String
    
On Error GoTo errHandle
    
    '����Zl_Exsesvr_Cancelacc_Check
'  ---------------------------------------------------------------------------
'  --input     �������ʼ�¼�˲�
'  --  check_people  C 1 �˲���
'  --  check_time    D 1 �˲�ʱ��
'  --  request_type  N   �������0-δִ��;1-��ִ��;��ҩƷ�����Ĺ̶���Ϊ0
'  --  rcpdtl_list     [����]ÿ��������ϸ��Ϣ
'  --    rcpdtl_id     N 1 ������ϸid(����id)
'  --    request_time  D 1 ����ʱ��
'  --output
'  --  code         C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message      C 1 Ӧ����Ϣ��
'  ---------------------------------------------------------------------------

    If strInput = "" Then zlSplitService_CallCancelAccCheck = True: Exit Function
    
    strJson_List = ""
    
    strPart1 = Split(strInput, "|")(0)
        
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("check_people", Split(strPart1, ",")(0), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("check_time", Split(strPart1, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("request_type", 1, 1)

    strJson_List = ""
    strPart2 = Split(strInput, "|")(1)
    arrPart = Split(strPart2, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("rcpdtl_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("request_time", Split(arrPart(i), ",")(1), 0)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
       
    '����
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    
    strService = "Zl_Exsesvr_Cancelacc_Check"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Cancelacc_Check��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallCancelAccCheck = False
        Exit Function
    End If

    zlSplitService_CallCancelAccCheck = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlSplitService_CallAccountDel(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strAccDel As String) As Boolean
    '��ҩ/��ҩ���÷�������/סԺ��¼����
    'strAccDel:����Ա����,����Ա���,����ʱ��||������Դ;��¼����;���õ��ݺ�;��Ŵ�;����״̬|������Դ;��¼����;���õ��ݺ�;��Ŵ�;����״̬...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim strPart As String
    
    On Error GoTo errHandle
    
    '����Zl_Exsesvr_Delbill ����/סԺ��¼ͨ�÷���
    
'  ---------------------------------------------------------------------------
'  --���ܣ�����ҽ�����ϣ�סԺҽ�����˷��ͣ�ɾ�����õ���
'  --��Σ�Json_In:��ʽ
'  --input
'  --   operator_name         C  1 ����Ա���������ʵ�ɾ��ʱ���롿
'  --   operator_code         C  1 ����Ա��š����ʵ�ɾ��ʱ���롿
'  --   operator_time         C  1 ����ʱ��:yyyy-mm-dd hh:mi:ss�����ʵ�ɾ��ʱ���롿
'  --   del_list  ֱ��ɾ���ĵ����б�
'  --             fee_source          N 1 ������Դ:1-������ü�¼;2-סԺ���ü�¼
'  --             fee_bill_type       N 1 ��¼���ʣ�1-�շѵ���2-���ʵ�
'  --             fee_no              C 1 ���õ��ݺ�
'  --             serial_num          C 1 ��Ŵ������ʵ���ʽ: ���1:����:ִ��״̬1,���2:����2:ִ��״̬2,...���շѵ���ʽ��1,2,3,4,5...
'  --             oper_status         N 1 ����״̬��סԺ���ʵ�ɾʱ�Ŵ��룬0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
'  --����: Json_Out,��ʽ����
'  --output
'  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  ---------------------------------------------------------------------------
    
    If strAccDel = "" Then zlSplitService_CallAccountDel = True: Exit Function
    
    strPart = Split(strAccDel, "||")(0)
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("operator_name", Split(strPart, ",")(0), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_code", Split(strPart, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_time", Split(strPart, ",")(2), 0)
    
    strPart = Split(strAccDel, "||")(1)
    strJson_List = ""
    If strPart <> "" Then
        arrInput = Split(strPart, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_source", Split(arrInput(i), ";")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_bill_type", Split(arrInput(i), ";")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_no", Split(arrInput(i), ";")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_num", Split(arrInput(i), ";")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_status", Split(arrInput(i), ";")(4), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""del_list"":[" & strJson_List & "]"
    End If
    
    '����
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    
    strService = "Zl_Exsesvr_Delbill"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Delbill��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallAccountDel = False
        Exit Function
    End If

    zlSplitService_CallAccountDel = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallAccountVerify(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strMain As String, ByVal strInput As String) As Boolean
    '���ܣ���ҩ�������˻��ۼ��˵�
    '��Σ�
    '   strMain������ʱ��,����Ա����,����Ա���
    '   strInput��������Դ|���ݺ�|����ID|���s||������Դ|���ݺ�|����ID|���s||...
    Dim arrInput As Variant, arrItem As Variant, arrMain As Variant
    Dim i As Integer, StrJson_In As String
    Dim strJson_bill As String, strJson_item As String, strJson_List As String
    
    'Zl_Exsesvr_Billverify
    '  --���ܣ����õ������
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    operator_time         C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '  --    operator_name         C 1 ����Ա����
    '  --    operator_code         C 1 ����Ա���
    '  --    item_list
    '  --        fee_source        N 1 ������Դ:1-����;2-סԺ
    '  --        fee_no            C 1 ���õ��ݺ�
    '  --        serial_nums       C 0 ��Ŵ���������ʾ���ŵ���
    '  --        pharmacy_window   C 0 ��ҩ���ڣ�������ԴΪ����ʱ���룬��ʽ���ⷿID1:��ҩ����1,�ⷿID2:��ҩ����2,....
    '  --        pati_id           N 0 ����id��������ԴΪסԺ�Ұ��������ʱ����(��Ҫ��Լ��ʱ�)
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    If strInput = "" Then Exit Function
    
    '������Դ|���ݺ�|����ID|���s||������Դ|���ݺ�|����ID|���s||...
    arrInput = Split(strInput, "||")
    strJson_List = ""
    
    For i = 0 To UBound(arrInput)
        arrItem = Split(arrInput(i), "|")
        
        strJson_item = ""
        strJson_item = strJson_item & "" & GetJsonNodeString("fee_source", arrItem(0), 1)
        strJson_item = strJson_item & "," & GetJsonNodeString("fee_no", arrItem(1), 0)
        strJson_item = strJson_item & "," & GetJsonNodeString("pati_id", arrItem(2), 1)
        strJson_item = strJson_item & "," & GetJsonNodeString("serial_nums", arrItem(3), 0)
        
        strJson_List = IIf(strJson_List = "", "", strJson_List & ",") & "{" & strJson_item & "}"
    Next
    
    arrMain = Split(strMain, ",")
    strJson_bill = ""
    strJson_bill = strJson_bill & "" & GetJsonNodeString("operator_time", arrMain(0), 0)
    strJson_bill = strJson_bill & "," & GetJsonNodeString("operator_name", arrMain(1), 0)
    strJson_bill = strJson_bill & "," & GetJsonNodeString("operator_code", arrMain(2), 0)
    
    StrJson_In = "{""input"":{" & strJson_bill & ",""item_list"":[" & strJson_List & "]" & "}}"
        
    '���÷���
    If objServiceCall.CallService("Zl_Exsesvr_Billverify", StrJson_In, , "", lngMode, False, , , , True) = False Then
        MsgBox "���á�Zl_Exsesvr_Billverify��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    zlSplitService_CallAccountVerify = True
End Function

Public Function zlSplitService_CallAdviceIsInvalid(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, strOut As String, Optional ByRef colOutlist As Collection) As Boolean
    '��֯ҽ���������ݣ�ҽ��id,ҽ��id...
    '����ֵ�������ϵ�ҽ��ID����ҽ��id,ҽ��id...
    Dim arrLongString As Variant
    Dim i As Integer, n As Integer
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String, strAdvices As String
       
    'zl_CisSvr_AdviceIsInvalid
'  ---------------------------------------------------------------------------
'  --���ܣ�����ҽ��ID��ѯҽ��״̬
'  --��Σ�Json_In:��ʽ
'  --input
'  --   advice_ids           C  1  ���ҽ��ID����,�ָ�
'  --����: Json_Out,��ʽ����
'  --output
'  --    code                 N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message              C  1  Ӧ����Ϣ
'  --    advice_ids           C  1  ҽ��ID�������ϵģ�
'  ---------------------------------------------------------------------------
    strOut = ""
    Set colOutlist = New Collection
    
    arrLongString = GetArrayByStr(strInput, 3900, ",")
    For i = 0 To UBound(arrLongString)
        '���
        StrJson_In = ""
        StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", arrLongString(i), 0)
        StrJson_In = "{""input"":{" & StrJson_In & "}}"
        strService = "zl_CisSvr_AdviceIsInvalid"
    
        '���÷���
        If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
            MsgBox "���á�zl_CisSvr_AdviceIsInvalid��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '���س���
        strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
        
        '��������
        strAdvices = objServiceCall.GetJsonNodeValue("output.advice_ids")
        
        If strAdvices <> "" Then
            For n = 0 To UBound(Split(strAdvices, ","))
                colOutlist.Add strAdvices
            Next
            strOut = IIf(strOut = "", "", strOut & ",") & strAdvices
        End If
    Next

    zlSplitService_CallAdviceIsInvalid = True
End Function



Public Function zlSplitService_CallAuditContent(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strMain As String, ByVal strInput As String) As Boolean
    '����ҽ�����
    'strMain�������
    'strInput��ҽ��������ݣ���ʽ������ ID1,����1,˵��1||ID2,����2,˵��2��
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    '����zl_CISSvr_AuditDrugOrder
    '  ---------------------------------------------------------------------------
    '  --���      json
    '  --input     ����ҽ�����
    '  --  auditor        C  1  �����
    '  --  audit_content  C  1  ҽ��������ݣ���ʽ������ID1,����1,˵��1||ID2,����2,˵��2��
    '  --����      json
    '  --output
    '  --  code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --  message C 1 Ӧ����Ϣ��
    '  ---------------------------------------------------------------------------

    
    If strInput = "" Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("auditor", strMain, 0)
    strJson = strJson & "," & GetJsonNodeString("audit_content", strInput, 0)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "zl_CISSvr_AuditDrugOrder"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�zl_CISSvr_AuditDrugOrder��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallAuditContent = False
        Exit Function
    End If

    zlSplitService_CallAuditContent = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallCancelAccAudit(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strAccAudit As String, ByVal strAccReAudit As String, ByVal strAccDel As String, ByVal strExeSta As String, ByVal strCheck As String) As Boolean
    '��ҩͬʱ����ʱ���õ��������񣺺ϲ�����������ˣ�����������ˣ�����/סԺ��¼ɾ�������·��ü�¼״̬�����ʼ��ȹ���
    '�ϲ�Ϊһ�����������Ч���Ʒ�ҩ/��ҩ����
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim arrList As Variant, arrPart As Variant
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    Dim n As Integer
    Dim strJson_Part As String, strCheckJson_In As String
    
    On Error GoTo errHandle
    
    '1. �����������
    'strAccAudit��������ϸid(����id),����ʱ��,�����,���ʱ��,�������״̬|...
'    rcpdtl_list         [����]ÿ��������ϸ��Ϣ
'        rcpdtl_id   N   1   ������ϸid(����id)
'        request_time    D   1   ����ʱ��
'        auditor C   1   �����
'        audit_time  D   1   ���ʱ��
'        cancel_status   N   1   �������״̬
'        auto_stuff_return   N       �Զ����ϣ�Ĭ�ϴ�1��
'        request_type    N       �������Ĭ�ϴ�1��

    
    strJson_List = ""
    If strAccAudit <> "" Then
        arrInput = Split(strAccAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("auditor", Split(arrInput(i), ",")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("audit_time", Split(arrInput(i), ",")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("cancel_status", Split(arrInput(i), ",")(4), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '2. �����������룺������ϸid(����id),����ʱ��,�����,���ʱ��,�������״̬|...
'    retrial_rcpdtl_list            [����]ÿ��������ϸ��Ϣ
'        rcpdtl_id   N   1   ������ϸid(����id)
'        request_time    D   1   ����ʱ��
'        auditor         C   1   �����
'        audit_time  D   1   ���ʱ��
'        oper_type   N   1   ��������:0-��˾ܾ� 1-ȡ���ܾ�
'        auto_stuff_return   N       �Զ����ϣ�Ĭ�ϴ�1��
'        request_type    N       �������Ĭ�ϴ�1��

    
    strJson_List = ""
    If strAccReAudit <> "" Then
        arrInput = Split(strAccReAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("auditor", Split(arrInput(i), ",")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("audit_time", Split(arrInput(i), ",")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_type", Split(arrInput(i), ",")(4), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""retrial_rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '3. ����/סԺ��¼ɾ���� no;�������(��ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������δ��˵���);����Ա����;����Ա����;��¼����;����״̬;��Һ���;�Ǽ�ʱ��,������Դ|...
'  --  rcp_list         [����]ÿ�����˴�����Ϣ
'  --    rcp_no              N  1  ����no
'  --    serial_nums         C  1  ��ʽΪ�����1:����1:ִ��״̬1,���2:����2:ִ��״̬2,...���n:����n:ִ��״̬n  ��:1:2:1,2:10:1,3:2:1
'  --    operator_code       C  1  ����Ա����
'  --    operator_name       C  1  ����Ա����
'  --    fee_properties      N     ��¼����
'  --    operator_status     N     ����״̬��0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
'  --    pivas_flag          N     �Ƿ���Һ��ҩ��飺0-ҽ�����ã������ҩƷ�Ƿ������Һ��ҩ���ģ�1-��ҽ�����ã����ҩƷ�Ƿ������ҩ����
'  --    create_time         D     �Ǽ�ʱ��
'  --    fee_origin          N     ������Դ(1-������ã�2-סԺ����)
    
    strJson_List = ""
    If strAccDel <> "" Then
        arrInput = Split(strAccDel, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ";")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), ";")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_code", Split(arrInput(i), ";")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_name", Split(arrInput(i), ";")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_properties", Split(arrInput(i), ";")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), ";")(5), 1)
            strJson = strJson & "," & GetJsonNodeString("pivas_flag", Split(arrInput(i), ";")(6), 1)
            strJson = strJson & "," & GetJsonNodeString("create_time", Split(arrInput(i), ";")(7), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_origin", Split(arrInput(i), ";")(8), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '4. ���·���ִ��״̬������ID��;ִ��״̬|����ID��;ִ��״̬...
'    bill_status_list            [����]���·���ִ��״̬������Ϣ(ִ��״̬��ͬ��ƴ��һ������id��)
'        detail_ids  C   1   ������ϸid��(����id��),֧�ֶ��id���á�,���ָ�
'        exe_status  N   1   ִ��״̬(0-δִ��;1-��ȫִ��;2-����ִ��)
    
    strJson_List = ""
    If strExeSta <> "" Then
        If InStr(strExeSta, "||") > 0 Then
            strExeSta = Split(strExeSta, "||")(1)
        End If
        
        If strExeSta <> "" Then
            '��֯���,detail_ids�ڵ�ֵ���ܳ�������Ҫ���ⲿ����strExeStaֵʱ�Ƚ��зֽ�
            arrInput = Split(strExeSta, "|")
            
            For i = 0 To UBound(arrInput)
                strJson = ""
                strJson = strJson & "" & GetJsonNodeString("detail_ids", Split(arrInput(i), ";")(0), 0)
                strJson = strJson & "," & GetJsonNodeString("exe_status", Split(arrInput(i), ";")(1), 1)
                 
                If strJson_List = "" Then
                    strJson_List = "{" & strJson & "}"
                Else
                    strJson_List = strJson_List & "," & "{" & strJson & "}"
                End If
            Next
            
            strJson_List = ",""bill_status_list"":[" & strJson_List & "]"
        End If
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '5.���ʼ��
'  --  feecheck_list    [����]�������ʼ������Ҫ��Ϣ
'  --    fee_no                     C   1   ���õ��ݺ�
'  --    balance_ban_writeoffs      N   1   �ѽ��ֹ����:����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
'  --    part_ban_writeoffs         N   1   ��ֹ��������:1-������0-����
'  --    oper_type                  N 1 ����״̬_In:0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
'  --    fee_origin                 N 1            ������Դ��1-������ʣ�2-סԺ���ʣ�
'  --    item_list[]                 ���������б�
'  --        serial_num              N   1   ���
'  --        quantity                N   1   ��������(Ϊ��ʱ�������ֱ������)
'  --    excute_list[]               ҩƷ����������Ӧ��ִ���б�
'  --        fee_id                  N   1   ����ID
'  --        sended_num              N   1   �ѷ�����

    'strCheck ��ʽ��no,�ѽ��ֹ����(1),ҽ����ֹ��������(0),����״̬(1),������Դ|���,��������;���,��������...|����id,�ѷ�����;����id,�ѷ�����...||...
    arrList = Split(strCheck, "||")
    
    For n = 0 To UBound(arrList)
        arrPart = Split(arrList(n), "|")
        
        strPart1 = arrPart(0)
        strPart2 = arrPart(1)
        strPart3 = arrPart(2)
            
        strJson_In_Part1 = ""
        strJson_In_Part1 = strJson_In_Part1 & "" & GetJsonNodeString("fee_no", Split(strPart1, ",")(0), 0)
        strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("balance_ban_writeoffs", Split(strPart1, ",")(1), 1)
        strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("part_ban_writeoffs", Split(strPart1, ",")(2), 1)
        strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("oper_type", Split(strPart1, ",")(3), 1)
        strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("fee_origin", Split(strPart1, ",")(4), 1)
    
        strJson_List = ""
        arrPart = Split(strPart2, ";")
        For i = 0 To UBound(arrPart)
            strJson_Part = ""
            strJson_Part = strJson_Part & "" & GetJsonNodeString("serial_num", Split(arrPart(i), ",")(0), 1)
            strJson_Part = strJson_Part & "," & GetJsonNodeString("quantity", zlStr.FormatEx(Split(arrPart(i), ",")(1), 5, False), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson_Part & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
            End If
        Next
        strJson_In_Part2 = ",""item_list"":[" & strJson_List & "]"
        
        strJson_List = ""
        arrPart = Split(strPart3, ";")
        For i = 0 To UBound(arrPart)
            strJson_Part = ""
            strJson_Part = strJson_Part & "" & GetJsonNodeString("fee_id", Split(arrPart(i), ",")(0), 1)
            strJson_Part = strJson_Part & "," & GetJsonNodeString("sended_num", Split(arrPart(i), ",")(1), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson_Part & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
            End If
        Next
        strJson_In_Part3 = ",""excute_list"":[" & strJson_List & "]"
        
        If strCheckJson_In = "" Then
            strCheckJson_In = "{" & strJson_In_Part1 & strJson_In_Part2 & strJson_In_Part3 & "}"
        Else
            strCheckJson_In = strCheckJson_In & "," & "{" & strJson_In_Part1 & strJson_In_Part2 & strJson_In_Part3 & "}"
        End If
    Next
    
    strCheckJson_In = ",""feecheck_list"":[" & strCheckJson_In & "]"
    StrJson_In = StrJson_In & strCheckJson_In
     
    '����
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "zl_ExseSvr_BillChargeOffExtend"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�zl_ExseSvr_BillChargeOffExtend��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallCancelAccAudit = False
        Exit Function
    End If
    
    zlSplitService_CallCancelAccAudit = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_WriteOff(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strAccAudit As String) As Boolean
    '��ҩͬʱ����ʱ���õ��������񣺺ϲ�����������ˣ�����������ˣ�����/סԺ��¼ɾ�������·��ü�¼״̬�����ʼ��ȹ���
    '�ϲ�Ϊһ�����������Ч���Ʒ�ҩ/��ҩ����
    'strBase��������Ϣ����ʽ ������Դ,����Ա����,����Ա����,����ʱ��
    'strAccAudit�������б���ʽ ����ID,����ʱ��,��������,�������,��������,�ѷ�����|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim arrList As Variant, arrPart As Variant
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    Dim n As Integer
    Dim strJson_Part As String, strCheckJson_In As String
    
    On Error GoTo errHandle
    
'    'Zl_Exsesvr_Drugwriteoff
'      ---------------------------------------------------------------------------
'  --���ܣ�ҩƷ�����ķ�������(�������ͨ��������ܾ���ȡ���ܾ�)
'  --��Σ�Json_In:��ʽ
'  --input
'  --  fee_origin            N  1  ������Դ��1-���2-סԺ��
'  --  operator_code         C  1  ����Ա����
'  --  operator_name         C  1  ����Ա����
'  --  operator_time         C     ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --  rcpdtl_list                 [����]ÿ��������ϸ��Ϣ
'  --    rcpdtl_id           N  1  ������ϸid(����id)
'  --    request_time        D  1  ����ʱ��
'  --    oper_type           N  1  ��������:0-���ͨ��;1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
'  --    request_type        N  1  �������Ĭ�ϴ�1��
'  --    quantity            N  1  ��������
'  --    sended_num          N  1  �ѷ�����
'
'  --����: Json_Out,��ʽ����
'  --output
'  --   code                          C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --   message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  ---------------------------------------------------------------------------
    
    '1.������Ϣ
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_origin", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_code", Split(strBase, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_name", Split(strBase, ",")(2), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_time", Split(strBase, ",")(3), 0)
    
    '2.������Ϣ
    strJson_List = ""
    If strAccAudit <> "" Then
        arrInput = Split(strAccAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_type", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("request_type", Split(arrInput(i), ",")(3), 1)
            strJson = strJson & "," & GetJsonNodeString("quantity", Split(arrInput(i), ",")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("sended_num", Split(arrInput(i), ",")(5), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
     
    '����
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Drugwriteoff"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Drugwriteoff��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_WriteOff = False
        Exit Function
    End If
    
    zlSplitService_WriteOff = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_WriteOffCheck(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strAccAudit As String, ByVal strPatiList As String, ByRef strCheckMsg As String) As Boolean
    '��ҩͬʱ����ʱ���õ������������ʼ��
    '�ϲ�Ϊһ�����������Ч���Ʒ�ҩ/��ҩ����
    'strBase��������Ϣ����ʽ ��ֹ��������,������Դ
    'strAccAudit�������б���ʽ  ����ID,����ʱ��,��������,�������,��������,�ѷ�����|...
    'strPatiList�������б���ʽ  ����id,������˱�־,סԺ״̬,������Ŀ����|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim arrList As Variant, arrPart As Variant
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    Dim n As Integer
    Dim strJson_Part As String, strCheckJson_In As String
    Dim colOutlist As New Collection
    
    On Error GoTo errHandle
''    'Zl_Exsesvr_Drugwriteoff_Check
'  ---------------------------------------------------------------------------
'  --���ܣ�ҩƷ�����ķ���������˼��
'  --��Σ�Json_In:��ʽ
'  --input      ҩƷ��������ǰ���
'  --  part_ban_writeoffs    N  1  ��ֹ��������:0-����;1-������������(�����ŵ��ݵĲ��ֻ�ĳ�ʵĲ���)
'  --  fee_origin            N  1  ������Դ:1-���2-סԺ
'  --  rcpdtl_list[]               ���������б�
'  --    oper_type           N  1  ��������:0-���ͨ�� 1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
'  --    rcpdtl_id           N  1  ������ϸID(����ID)
'  --    request_time        D     ����ʱ��
'  --    request_type        N     �������ȱʡΪ1
'  --    quantity            N  1  ����������Ϊ���nullʱ,������ID��������ֱ������
'  --    sended_num          N  1  �ѷ�����
'  --  pati_list[]                 ������Ϣ
'  --    pati_id             N     ����ID,ΪNULL��0ʱ����ʾ���ŵ���
'  --    fee_audit_status    N     ������˱�־:0���-δ���;1-����˻�ʼ���;2-������,��Ͻ���Ȩ��
'  --    si_inp_status       N     סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ
'  --    catalog_date        C     ������Ŀ���ڣ�yyyy-mm-dd hh24:mi:ss
'  --����: Json_Out,��ʽ����
'  --output
'  --   code                          C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --   message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --  tip_list[]  C  1  ��ʾ�б�:��Ҫ�ǿ��ܴ��ڶ����ʾѯ�ʷ�ʽ���������б�,��ֹʱ������һ����Ϣ
'  --    tip_mode  C  1  ���Ʒ�ʽ:1-��ʾѯ��;2-��ֹ
'  --    tip_message  C  1  ��ʾ��Ϣ
'  ---------------------------------------------------------------------------
    
    '1.������Ϣ
    StrJson_In = StrJson_In & "" & GetJsonNodeString("part_ban_writeoffs", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_origin", Split(strBase, ",")(1), 1)
    
    '2.������Ϣ
    strJson_List = ""
    If strAccAudit <> "" Then
        arrInput = Split(strAccAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""            '�����б�����ID,����ʱ��,��������,�������,��������,�ѷ�����|...
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_type", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("request_type", Split(arrInput(i), ",")(3), 1)
            strJson = strJson & "," & GetJsonNodeString("quantity", Split(arrInput(i), ",")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("sended_num", Split(arrInput(i), ",")(5), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '3.������Ϣ
    strJson_List = ""
    If strPatiList <> "" Then
        arrInput = Split(strPatiList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""            '����id,������˱�־,סԺ״̬,������Ŀ����|...
            strJson = strJson & "" & GetJsonNodeString("pati_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_audit_status", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("si_inp_status", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("catalog_date", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""pati_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
     
    '����
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Drugwriteoff_Check"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Drugwriteoff_Check��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        '�������ʧ��ʱ
        strCheckMsg = objServiceCall.GetJsonNodeValue("output.message")
        zlSplitService_WriteOffCheck = False
        Exit Function
    Else
        '������óɹ�������ҵ������
        Set colOutlist = objServiceCall.GetJsonListValue("output.tip_list")
        If colOutlist.count > 0 Then
            '���Ʒ�ʽ,��ʾ��Ϣ
            strCheckMsg = colOutlist(1)(1) & "," & colOutlist(1)(2)
        End If
    End If
    
    zlSplitService_WriteOffCheck = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetCloseAccount(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal colInput As Collection, _
    ByRef colCloseAccount As Collection) As Boolean
    'ȡ���ʼ�¼
    'colInput����ѯ������ϣ�Json��input���ڵ���ΪԪ�ص�KEYֵ������ĳԪ��Ϊ�ձ�ʾ�ýڵ�ֵΪ��
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    'Zl_Exsesvr_Getchargeoffinfo
    '  --����������ѯ����������Ϣ
    '  --���      json
    '  --  input      ����������ѯ����������Ϣ
    '  --    audit_dept_id       N    ��˲���ID(ҩ��)
    '  --    request_begin_time  D    ���뿪ʼʱ��
    '  --    request_end_time    D    �������ʱ��
    '  --    audit_begin_time    D    ��˿�ʼʱ��
    '  --    audit_end_time      D    ��˽���ʱ��
    '  --    cancel_status       N  1 ״̬
    '  --    request_dept_id     N    ���벿��ID
    '  --    request_operator    C    ������
    '  --    pati_id             N    ����ID
    '  --    cancel_condition    C    ��������
    '  --    cancel_check        N    �˲飨ѡ�����������������Ҫ�˲顿ʱ���룬0-δ�˲� 1-�Ѻ˲飩
    '  --    rcpdtl_id          C     ������ϸid,[����]��[1,2,3]
    '  --    request_dept_ids   C     ���벿��id��������������ѯ
    '  --    item_ids           C     �շ�ϸĿid��,����������ѯ
    '  --    request_type       N     �������-1-������;0-δִ��;1-��ִ��
    '  --����      json
    '  -- output
    '  --   code     C  1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --   message  C  1   Ӧ����Ϣ��
    '  --   fee_cancel_list      [����]����������ÿ���������ʼ�¼
    '  --     rcpdtl_id          N    ������ϸid(����id)
    '  --     request_type       N    �������
    '  --     item_id            N    �շ�ϸĿid
    '  --     request_dept_id    N    ���벿��id
    '  --     request_dept       C    ���벿��
    '  --     audit_dept_id      N    ��˲���id
    '  --     quantity           N    ����
    '  --     request_operator   C    ������
    '  --     request_time       D    ����ʱ��
    '  --     auditor            C    �����
    '  --     audit_time         D    ���ʱ��
    '  --     cancel_status      N    ״̬
    '  --     cancel_reason      C    ����ԭ��
    '  --     checker            C    �˲���
    '  --     price_retail       N    ���ۼ�
    '  --     advice_id          N    ҽ��id
    '  --     pati_id            N    ����ID
    '  --     pati_name          C    ��������
    '  --     inpatient_num      C    סԺ��
    '  --     pati_pageid        N    ��ҳid
    StrJson_In = ""
    If ExistsColObject(colInput, "audit_dept_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_dept_id", colInput("audit_dept_id"), 1)
    If ExistsColObject(colInput, "request_begin_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_begin_time", colInput("request_begin_time"), 0)
    If ExistsColObject(colInput, "request_end_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_end_time", colInput("request_end_time"), 0)
    If ExistsColObject(colInput, "audit_begin_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_begin_time", colInput("audit_begin_time"), 0)
    If ExistsColObject(colInput, "audit_end_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_end_time", colInput("audit_end_time"), 0)
    
    If ExistsColObject(colInput, "cancel_status") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_status", colInput("cancel_status"), 1)
    If ExistsColObject(colInput, "request_dept_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_dept_id", colInput("request_dept_id"), 1)
    If ExistsColObject(colInput, "request_operator") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_operator", colInput("request_operator"), 0)
    If ExistsColObject(colInput, "pati_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_id", colInput("pati_id"), 1)
    If ExistsColObject(colInput, "cancel_condition") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_condition", colInput("cancel_condition"), 0)
    
    If ExistsColObject(colInput, "cancel_check") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_check", colInput("cancel_check"), 1)
    If ExistsColObject(colInput, "rcpdtl_id") Then StrJson_In = StrJson_In & "," & """rcpdtl_id"":" & "[" & colInput("rcpdtl_id") & "]"
    If ExistsColObject(colInput, "request_dept_ids") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_dept_ids", colInput("request_dept_ids"), 0)
    If ExistsColObject(colInput, "item_ids") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("item_ids", colInput("item_ids"), 0)
    If ExistsColObject(colInput, "request_type") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_type", colInput("request_type"), 1)
    
    If StrJson_In = "" Then Exit Function
    
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_ExseSvr_GetChargeOffInfo"
    
    '����
'    StrJson_In = "{""input"":{""audit_dept_id"":305,""cancel_status"":0,""rcpdtl_id"":[4528923,4528923]}}"
'    StrJson_In = "{""input"":{""audit_dept_id"":305,""cancel_status"":0,""rcpdtl_id"":[1]}}"
'    strService = "Zl_ExseSvr_GetChargeOffInfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_ExseSvr_GetChargeOffInfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colCloseAccount = objServiceCall.GetJsonListValue("output.fee_cancel_list")
    
    If colCloseAccount Is Nothing Then Exit Function
    
    zlSplitService_GetCloseAccount = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlSplitService_GetFee(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String, Optional ByVal strInputByNO As String) As Boolean
    'ȡ������Ϣ
    'strInput������id,����id...
    'strInputByNO��no,��¼����|...
    Dim StrJson_In As String
    Dim strService As String
    
    On Error GoTo errHandle
    'Zl_Exsesvr_GetBillDetailInfo
    '  -------------------------------------------------------------------------------------------------
    '  --���ܣ���ȡ��ҩƷ��ҩҵ����صķ�����Ϣ����Ҫ���ڽ�����ʾ
    '  --��Σ�json��ʽ
    '  --Input
    '  --   fee_ids    C     ����id��֧�ֶ��id����ʽ�� ����id,����id,��
    '  --   bill_nos   C     ����no,��¼���ʣ���ʽ: no,��¼����|,...
    '  --���Σ�json��ʽ
    '  --Json_Out
    '  --fee_list      [����]ÿ������ID��Ϣ
    '  --  bill_prop           N    ��¼����:1-�շѵ�;2-���ʵ�;3-�Զ����ʵ�;4-�Һŵ�;5-���￨;6-Ԥ����
    '  --  bill_no             C    ���ݺ�
    '  --  fee_id              N    ������ϸid(����id)
    '  --  fee_num             N    ���
    '  --  iden_id             N    ��ʶ��
    '  --  pati_bed            C    ����
    '  --  fee_ampaid          N    ʵ�ս��
    '  --  packages_num        N    ����
    '  --  quantity            N    ����
    '  --  placer              C    ������
    '  --  operator_code       C    ����Ա���
    '  --  operator_name       C    ����Ա����
    '  --  create_time         D    �Ǽ�ʱ��
    '  --  happen_time         D    ����ʱ��
    '  --  rcp_type            N    �������(������NO��˵��1-��ҩ��2-��ҩ��3-���)
    '  --  fee_type            C    �ѱ�
    '  --  rec_status          N    ��¼״̬
    '  --  register_id         N    �Һ�id
    '  --  register_no         C    �Һ�NO
    '  --  register_time       D    �ҺŵǼ�ʱ��
    '  --  income_item_id      N    ������Ŀid
    '  --  fee_origin          N    ������Դ(1-������ã�2-סԺ����)
    '  --  bill_deptid         N    ��������id
    '  --  order_id            N    ҽ��ID
    '  --  fee_item_id         N    �շ�ϸĿid
    '  --  fee_status          N    ����״̬
    '  -------------------------------------------------------------------------------------------------
    
    If strInput = "" And strInputByNO = "" Then Exit Function
    
    '���
    StrJson_In = ""
    If strInput <> "" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_ids", strInput, 0)
    ElseIf strInputByNO <> "" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("bill_nos", strInputByNO, 0)
    End If
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_GetBillDetailInfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, , "", lngMode) = False Then Exit Function
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.fee_list", strKey)
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetFee = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiDiagnose(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    'ȡ�����Ϣ
    'strInput������id,��ҳid...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    If strInput = "" Then Exit Function
    
    '���
    StrJson_In = ""
    If strInput <> "" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("visit_id", Split(strInput, ",")(1), 1)
    End If
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_CisSvr_GetPatiDiagnose"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_CisSvr_GetPatiDiagnose��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.diagnose_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetPatiDiagnose = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetDiagInfo(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    'ȡ�����Ϣ������
    'strInput��ҽ��id,ҽ��id...
    '         ��������ҽ��id
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    'Zl_Cissvr_Getdiaginfo
'  ---------------------------------------------------------------------------
'  --����:��ȡ���������Ϣ
'  --��Σ�Json_In:��ʽ
'  -- input
'  --   advice_ids           C 1  ҽ��ids,ҽ��idƴ��
'  --   query_type           N 1 ��ѯ��ʽ1-��ָ��������ѯ,2-��������id,��ҳid��ѯ���
'  --   pati_info            C 0  ����id��������Ϣ
'  --     pati_id            N 1 ����id
'  --     pati_pageid        N 1 ��ҳid
'  --     diag_types         C 1  �������:0-��������,1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
'  --                            ����Ϊ���������ͣ��ö��ŷ���,��:2,12
'  --     rec_source         N 1 ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����;NULL-��������
'  --     diag_num           N 1 ��ϴ���:NULL��ʾ��������
'  --     code_type          C 1  �������:ICD-11�ı���������Ϊ'E',��ʱ��ʾ��ȡICD-10��
'  --     input_num          C 1  ¼�����:������ICD-11����¼�����ϵ�¼�����
'  --     rec_sources        C 1 ��¼��Դƴ��
'
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                N  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    diag_list     [����]
'  --      diag_type         N 1 �������:1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;
'  --                                     12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
'  --      diag_num          N 1 ������
'  --      code_num          N 1 �������
'  --      dz_id             N 1 ����ID
'  --      dz_code           C 1 ��������
'  --      diag_note         C 1 �������
'  --      recoder           C 1 ��¼��
'  --      rec_time          C 1 ��¼ʱ��:yyyy-mm-dd hh24:mi:ss
'  --      adtd_rsn          C 1 ��Ժ���:��������ת��δ��������������
'  --      diag_id           N 1 ���id
'  --      diag_rec_id       N 1 ��ϼ�¼ID:������ϼ�¼.ID
'  --      diag_doubt        N 1 �Ƿ�����
'  --      advice_id         N   ҽ��id(����ҽ��ids��ѯʱ�ŷ���)
'  --      advice_main_id    N   ��ҽ��id(����ҽ��ids��ѯʱ�ŷ���)
'  --      advice_related_id N   ���id(����ҽ��ids��ѯʱ�ŷ���)
'  --      rec_source        N   ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����;NULL-��������
'  ---------------------------------------------------------------------------
  
    On Error GoTo errHandle
    
    If strInput = "" Then Exit Function
    
    '���
    StrJson_In = ""
    If strInput <> "" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", strInput, 0)
    End If
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getdiaginfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Cissvr_Getdiaginfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.diag_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetDiagInfo = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_GetAccWarnLine(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOut As String) As Boolean
    '��ȡ���˱�����
    'strInput�����һ���id,��������
    'strOut: ��������,����ֵ,������־1,������־2,������־3
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim colOutlist As New Collection
    
    'Zl_Exsesvr_Getwarnline
'  ---------------------------------------------------------------------------
'  --���ܣ���ȡ���ʱ����ߣ����˲��������ſ�Ƿ�Ѳ���
'  --��Σ�Json_In:��ʽ
'  --  input
'  --     pati_scheme  C 1 ���ò���
'  --     wardarea_id  N 1 ����id
'  --     query_type   N 1 ��ѯ��ʽ
'  --                     0-������ ����id / ���ò��� ���ң�����һ��ֵ
'  --                     1-������id ���ң������б�
'  --                     2-��ȡ���б���������
'  --                     3-���ݲ���id�����ò��˲��ң����ر�������������ֵ��������־
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --      alarm_value       N 1 ����ֵ
'  --      item_list[]
'  --        pati_scheme     C 1 ���ò���
'  --        alarm_way       N 1 ��������
'  --        alarm_value     N 1 ����ֵ
'  --        alarm_one       C 1 ������־1
'  --        alarm_two       C 1 ������־2
'  --        alarm_three     C 1 ������־3
'  --        wardarea_id     N 1 ����id
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_scheme", Split(strInput, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("wardarea_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", 3, 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_Getwarnline"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Exsesvr_Getwarnline��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    Set colOutlist = objServiceCall.GetJsonListValue("output.item_list")
    
    If Not colOutlist Is Nothing Then
        If colOutlist.count > 0 Then
            strOut = colOutlist(1)("_pati_scheme")
            strOut = strOut & "," & colOutlist(1)("_alarm_way")
            strOut = strOut & "," & colOutlist(1)("_alarm_value")
            strOut = strOut & "," & colOutlist(1)("_alarm_one")
            strOut = strOut & "," & colOutlist(1)("_alarm_two")
            strOut = strOut & "," & colOutlist(1)("_alarm_three")
        End If
    End If
        
    zlSplitService_GetAccWarnLine = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetTodayMoney(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByVal bln������ As Boolean, ByRef strOut As String) As Boolean
    '��ȡ���˱�����
    'strInput������id
    'strOut: �ܶ�
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
     'Zl_Exsesvr_Getpatitotalmoney
'  ---------------------------------------------------------------------------
'  --����:���ݲ���ID,��ҳID��ҽ��id����ȡӦ�ա�ʵ���ܶ�
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    pati_source N 1 ������Դ:0-����;1-����;2-סԺ
'  --    pati_id N 1 ����ID
'  --    visit_id  N   ����ID:סԺʱ��������ҳid,�����ݴ�NULL
'  --    advice_ids  C   ҽ��ids:����ö��ŷ���
'  --    today_fee N   �Ƿ��շ���:1-�ǵ�;0-������
'  --    price_tag N   ���۱�־:0-������;1-�������۵�;2-��ͳ�ƻ��۵�
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code        C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message     C  1  "Ӧ����Ϣ��
'  --    fee_amrcvb  N  1  Ӧ�ս��
'  --    fee_ampaib  N  1  ʵ�ս��
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_source", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_id", strInput, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("today_fee", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("price_tag", IIf(bln������, 0, 1), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_Getpatitotalmoney"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Exsesvr_Getpatitotalmoney��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")

    strOut = objServiceCall.GetJsonNodeValue("output.fee_ampaib")
        
    zlSplitService_GetTodayMoney = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetFeeNO(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByVal lngҩ��id As Long, ByRef colOutlist As Collection) As Boolean
    'ȡ����NO������Ϣ
    'strInput����¼����,NO|...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    'Zl_Exsesvr_Getbillinfo
    '  --���ܣ���ȡ��ҩƷ��ҩҵ����صķ�����Ϣ����Ҫ���ڽ�����ʾ
    '  --��Σ�json��ʽ
    '  --Input
    '  --   pharmacy_id���ⷿid
    '  --   fee_nos������no��֧�ֶ��no����ʽ�� ��¼����,no,��
    '  --���Σ�json��ʽ
    '  --Json_Out
    '  --  code          C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --  message       C   1   Ӧ����Ϣ��
    '  --  fee_list      C       [����]ÿ������NO��Ϣ
    '  --    fee_properties      N ��¼����
    '  --    bill_no             C ����no
    '  --    real_amount         N ʵ�ս��
    '  --    rcp_type            N �շ����(������NO��˵��1-��ҩ��2-��ҩ��3-���)
    '  --    iden_id             C ��ʶ��
    '  --    placer              C ������
    '  --    bill_deptid         N ��������id
    '  --    create_time         D �Ǽ�ʱ��
    '  --    pati_bed            C ��ǰ����
    '  --    operator_name       C ����Ա����
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pharmacy_id", lngҩ��id, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_nos", strInput, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "zl_ExseSvr_GetBillInfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_ExseSvr_GetBillInfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.fee_list")
    If colOutlist Is Nothing Then Exit Function

    zlSplitService_GetFeeNO = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetNOByInvoice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
     ByRef strOutPut As String) As Boolean
    'ͨ��Ʊ�ݺ�ȡ����NO
    'strInput��Ʊ�ݺ�
    'strOutput��NO1,NO2...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim n As Integer
    
    'zl_ExseSvr_GetNoByInvoice
'  -------------------------------------------------------------------------------------------------
'  --���ܣ���Ʊ�ݺŷ�ҩ����ҩ��ͨ��¼�뷢Ʊ�Ż�ȡ��Ӧ��ҩƷ����NO
'  --��Σ�json��ʽ
'  --Input
'  --   invc_no  C  1  Ʊ�ݺ�
'  --���Σ�json��ʽ
'  --Json_Out
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  "Ӧ����Ϣ�� �ɹ�ʱ���ش���No��[����] ʧ��ʱ���ؾ���Ĵ�����Ϣ"
'  --  rcp_nos  C  1 �������ݺţ�����ö��ŷָ�
'  -------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("invc_no", strInput, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "zl_ExseSvr_GetNoByInvoice"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_ExseSvr_GetNoByInvoice��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    strOutPut = objServiceCall.GetJsonNodeValue("output.rcp_nos")

    zlSplitService_GetNOByInvoice = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetNextNO(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef strNos_Out As String, Optional ByVal lngCount As Long = 1) As Boolean
    '��ȡ����ҵ��NO
    '��Σ�
    '   strInput�����|����ID
    '   lngCount = ��ȡNO����
    '���Σ�
    '   strNos_Out = NO,������ŷָ�
    Dim StrJson_In As String, strJson_Out As String
    
    On Error GoTo errHandle
    'Zl_Exsesvr_Getnextno
    '  --���ܣ����ܣ������ض���������µĺ���
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    item_num            N   1   ��Ŀ���
    '  --    dept_id             N   0   ����ID
    '  --    quantity            N   0   ����no�ŵĸ��������ֻȡһ���òβ����򶼴�0
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    next_no             C   1   ��һ������,quantity>1 ʱ����ʾȡ������ݺ�,�ö��ŷ���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("item_num", Val(Split(strInput, ",")(0)), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("dept_id", Val(Split(strInput, ",")(1)), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("quantity", lngCount, 1)
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    If objServiceCall.CallService("zl_ExseSvr_GetNextNo", StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_ExseSvr_GetNextNo����ȡ���õ��ݺ�ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    strNos_Out = objServiceCall.GetJsonNodeValue("output.next_no")
    If strNos_Out = "" Then
        MsgBox "���á�zl_ExseSvr_GetNextNo����ȡ���õ��ݺ�ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    zlSplitService_GetNextNO = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetAllergy(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
   ByRef colOutlist As Collection) As Boolean
    'ȡ����ҩ�������¼
    'strInput������id,��ʶid
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    'Zl_Cissvr_Getpatiallergyinfo
'    -------------------------------------------------------------------------------------------------
'    --���ܣ���ȡ���˹�����Ϣ
'--input      ��ȡ���˹�����Ϣ
'--  pati_id  N  1  ����id
'--  visit_id  N    ��ʶ�ţ��Һ�id���������ҳid��סԺ��
'--output
'--  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'--  message  C  1  "Ӧ����Ϣ��
'--  allergy_list  C    ������Ϣ��[����]
'--     drug_name  C    ҩ������
'--     allergy_time  D    ����ʱ��
'--     allergy_info  C    ������Ӧ
'    -------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("visit_id", Split(strInput, ",")(1), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "Zl_Cissvr_Getpatiallergyinfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Cissvr_Getpatiallergyinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.allergy_list")
    
    If colOutlist Is Nothing Then Exit Function

    zlSplitService_GetAllergy = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPativitalsigns(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
   ByRef colOutlist As Collection) As Boolean
    'ȡ��������������Ϣ
    'strInput������id,��ʶid,�����־
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
'  --��ȡ��������������Ϣ
'  ---------------------------------------------------------------------------
'  --input      ��ȡ��������������Ϣ
'  --  pati_id  N  1  ����ID
'  --  visit_id  N  1  ����id �����ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;
'  --  outpati_flag  N    �����־��1-���2-סԺ
'  --output
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  "Ӧ����Ϣ��
'  --  pativital_list      ������Ϣ��������Ŀ����ֵ����λ��[����]
'  --     pativital_item  C    ��Ŀ
'  --     pativital_value  C    ֵ
'  --     pativital_unit  C    ��λ
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("visit_id", Split(strInput, ",")(1), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("outpati_flag", Split(strInput, ",")(2), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "Zl_Cissvr_Getpativitalsigns"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Cissvr_Getpativitalsigns��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.pativital_list")
    
    If colOutlist Is Nothing Then Exit Function

    zlSplitService_GetPativitalsigns = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetAdvice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String, Optional ByVal bytQueryType As Byte = 0) As Boolean
    'ȡҽ����Ϣ
    'strInput��ҽ��id,ҽ��id...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
            
'      --����:��ȡҽ����Ϣ
'  --��Σ�Json_In:��ʽ
'  -- input
'  --   query_type                   N 0 ��ѯ���ͣ�0:��ѯ������Ϣ��1:��ѯ������Ϣ+��չ��Ϣ
'  --   advice_ids                   C 0 ���ҽ��ID��������ҩ����Ҳ��������ҽ������ҩ;����,�á�,���ָ�
'  --   rgst_no                      C 0 �Һŵ���:�Һŵ�����ID��ҽ��ID�ش�����һ������
'  --   pati_id                      N 0 ����ID:�Һŵ�����ID��ҽ��ID�ش�����һ������
'  --   pati_pageid                  N 0 ��ҳId
  
    On Error GoTo errHandle
    
    If strInput = "" Then zlSplitService_GetAdvice = True: Exit Function
  
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", bytQueryType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("advice_ids", strInput, 0)
'    strJson_In = strJson_In & "," & GetJsonNodeString("advice_ids", strPatis, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_CisSvr_GetAdviceInfo"
    
    '����
    'StrJson_In = "{""input"":{""advice_id"":""521""}}"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_CisSvr_GetAdviceInfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.advice_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetAdvice = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetAdviceSend(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '��֯��ϸ���ݼ���ҽ��������Ϣ
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", strInput, 0)
'    strJson_In = strJson_In & "," & GetJsonNodeString("pati_list", strPatis, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_CisSvr_GetAdviceSendInfo"

    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_CisSvr_GetAdviceSendInfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.advice_send_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function

    zlSplitService_GetAdviceSend = True
End Function

Public Function zlSplitService_CheckAdviceaffirm(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutPut As String, ByRef strErrMsg_Out As String) As Boolean
    '����ָ������ID����ҳID���Һ�ID��ҽ��������Ϣ
    'strInput������id,��ҳID,�Һ�ID,�Һŵ���|...
    'strOutPut��ҽ��id,���ͺ�|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colOutlist As Collection, colOrderList As Collection
    Dim arrOrder As Variant
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    'Zl_CISSvr_GetAffirmErrorData
    ' ---------------------------------------------------------------------------
    '  --���ܣ��ٴ�ҽ��ִ�����ʱ�Զ���˷��ã��쳣��δ��ҩƷ�����Ľ����շ�ȷ�ϣ���Դ����쳣���ݻ�ȡ����
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --      pati_list[]���˹ؼ���Ϣ�����ڻ�ȡҽ��
    '  --           pati_id                    N 1 ����id
    '  --           pati_pageid                N 1 ��ҳid��סԺ���˴��룬���ﴫ0
    '  --           rgst_id                    N 1 �Һ�id�����ﲡ�˴��룬סԺ���˴���
    '  --           rgst_no                    C 1 �Һŵ���
    '  --����: Json_Out,��ʽ����
    '  --   output:
    '  --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --     pati_bill_list[]
    '  --         pati_id                      N 1 ����id
    '  --         pati_pageid                  N 0 ��ҳid��סԺ���˴��룬���ﴫ0
    '  --         rgst_id                      N 0 �Һ�id�����ﲡ�˴��룬סԺ���˴���
    '  --         rgst_no                      C 0 �Һŵ���
    '  --         order_ids                    C 1 ���쳣������ҽ��idƴ��
    '  --         fee_nos                      C 1 ���쳣�����е��ݺ�ƴ��
    '  --         order_list[]ҽ��������Ϣ�б�
    '  --             send_no                  N 1 ���ͺ�
    '  --             advice_id                N 1 ҽ��id
    '  --             fee_no                   C 1 ���ݺ�
    '  --             bill_prop                N 1 ��¼����
    '  --             outpati_account          N 1 �Ƿ�������� 0-����������ʣ�1-���������
    '  --             pati_source              N 1 ������Դ 1-����ҽ����2-סԺҽ��
     
    If strInput = "" Then zlSplitService_CheckAdviceaffirm = True: Exit Function
    
    arrInput = Split(strInput, "|")
    For i = 0 To UBound(arrInput)
        StrJson_In = ""
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(arrInput(i), ",")(0), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageid", Split(arrInput(i), ",")(1), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("rgst_id", Split(arrInput(i), ",")(2), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("rgst_no", Split(arrInput(i), ",")(3), 0)
        
        If strJson_List = "" Then
            strJson_List = "{" & StrJson_In & "}"
        Else
            strJson_List = strJson_List & "," & "{" & StrJson_In & "}"
        End If
    Next
    StrJson_In = """pati_list"":[" & strJson_List & "]"
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_CISSvr_GetAffirmErrorData"
    
    If objServiceCall.CallService(strService, StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    
    Set colOutlist = objServiceCall.GetJsonListValue("output.pati_bill_list")
    'ֻ��Ҫ�������ݣ�ҽ��id,���ͺ�|...
    For i = 0 To colOutlist.count - 1
        'ѭ��ȡ�ӽڵ�����
        Set colOrderList = objServiceCall.GetJsonListValue("output.pati_bill_list[" & i & "].order_list")
        
        For Each arrOrder In colOrderList
            strOutPut = IIf(strOutPut = "", "", strOutPut & "|") & arrOrder("_advice_id") & "," & arrOrder("_send_no")
        Next
    Next

    zlSplitService_CheckAdviceaffirm = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_CheckErrData(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strType As String, ByVal strInputByid As String, ByVal strInputByNO As String, _
    ByRef colOutlist As Collection, ByRef colOutExpenseList As Collection, ByRef strErrMsg_Out As String) As Boolean
    '��ҩ/��ҩ���÷��񣺼������쳣״̬�������շѡ����״̬
    'strInputByid��1.������id���м�飬����id,����id...
    'strInputByNO: 2.������NO���м�飬��������:no1,no2|...
    '���Σ�1.���update_drug_status=2 then ��Ҫ���´��������շ�/���״̬���ٵ��÷��÷�����¶Է��ļǷ�ͬ��״̬
    '      2.���rcp_no_new<>"" then ��Ҫ���´�����NO������ID��
    '      3.���fee_status=2 then ���ܽ��з�ҩ����ҩ�Ȳ���
    Dim arrInput As Variant, i As Integer
    Dim strJson As String, StrJson_In As String, strJson_List As String
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    'Zl_Exsesvr_Checkerrordata
    '  --���ܣ����ݷ���NO�����ID����շѼ����쳣��Ϣ
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --      fee_type              C   1 �������'4'-���ģ�'5,6,7'-ҩƷ
    '  --      rcpdtl_ids            C   1 ������ϸids,����ö��ŷָ�
    '  --      bill_list[]                  ���飬����NO��Ϣ
    '  --         billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
    '  --         rcp_nos                C   1 ����Nos,����ö��ŷָ�
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --     code                   N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --     message                C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --     billid_list[]                   ������ID����ʱ����id�б�
    '  --        rcpdtl_id           N   1 ������ϸid
    '  --        fee_status          N   1 ����״̬�� 0-����,1-����
    '  --        cancel_status       N   1 ����״̬:0-����״̬,1-����ͬ����־�쳣
    '  --        update_status       N   1 �Ƿ�ͬ��״̬:0-����״̬,1-δ����ҩƷ/���ļ���״̬
    '  --     billno_list[]                 ��NO����ʱ����NO�б�
    '  --         billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
    '  --         rcp_no                 C   1 ����no
    '  --         fee_status             N   1 ����״̬������շ�ʱ,0-δ�շ�,1-���շ�,2-�쳣�շ�;��Լ���ʱ,0-����,1-����
    '  --         cancel_status          N   1 ����״̬:0-����״̬,1-����ͬ����־�쳣
    '  --         update_drug_status     N   1 �Ƿ�ͬ��״̬:0-����״̬,2-δ����ҩƷ/�����շ�״̬
    '  --     expense_list[]               ��ҩƷ����
    '  --         billtype               N   1 (ԭʼ)��������:1-�շѴ���;2-���ʴ���
    '  --         rcp_no                 C   1 (ԭʼ)����no
    '  --         rcpdtl_id              N   1 (ԭʼ)������ϸid
    '  --         rcp_no_new             C   1 �����ɵĴ���NO
    '  --         rcpdtl_id_new          N   1 �����ɴ�����ϸid
    '  --         pati_pageid            N   1  ��ҳID
    If strInputByid = "" And strInputByNO = "" Then zlSplitService_CheckErrData = True: Exit Function
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_type", strType, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("rcpdtl_ids", strInputByid, 0)
    If strInputByNO <> "" Then
        arrInput = Split(strInputByNO, "|")
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("billtype", Split(arrInput(i), ":")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_nos", Split(arrInput(i), ":")(1), 0)
            strJson_List = IIf(strJson_List = "", "", strJson_List & ",") & "{" & strJson & "}"
        Next
        strJson_List = ",""bill_list"":[" & strJson_List & "]"
    End If
    
    '����
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"

    If objServiceCall.CallService("Zl_Exsesvr_CheckErrorData", StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    If strInputByid <> "" Then
        Set colOutlist = objServiceCall.GetJsonListValue("output.billid_list", "rcpdtl_id")
    Else
       Set colOutlist = objServiceCall.GetJsonListValue("output.billno_list", "billtype,rcp_no")
    End If
    Set colOutExpenseList = objServiceCall.GetJsonListValue("output.expense_list")
    zlSplitService_CheckErrData = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_CallUpdateSynchrosign(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal intType As Integer, ByVal strInputByid As String, ByRef strErrMsg_Out As String) As Boolean
    '���·��üǷ�ͬ�����
    '������NO���и���
    'intType��0-���¼Ƿ�ͬ����־��1-����ת��ͬ����־
    'strInputByid��������ϸID�����Ӣ�Ķ��ŷָ�
    Dim StrJson_In As String
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    If strInputByid = "" Then zlSplitService_CallUpdateSynchrosign = True: Exit Function
    '����Zl_Exsesvr_Sync_Update
    '  --���ܣ������շ�ͬ����־
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    sign_type           N 1 ��־���ͣ�0-�Ƿ�ͬ����־,1-ת��ͬ����־
    '  --    detail_ids  C  1  ������ϸid��(����id��),֧�ֶ��id���á�,���ָ�
    '  --    bill_list[]
    '  --      billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
    '  --      rcp_no                 C   1 ����No
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --     code                   N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --     message                C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("sign_type", intType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("detail_ids", strInputByid, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Sync_Update", StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    
    zlSplitService_CallUpdateSynchrosign = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_CisUpdateSyncState(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal intType As Integer, ByVal strInput As String, ByRef strErrMsg_Out As String) As Boolean
    '����ҽ��ͬ�����¼����
    'intType��1-���䣬2-ҩƷ ��3-����
    'strInput��ҽ��id,���ͺ�|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    '����Zl_CisSvr_UpdateSyncState
    '  ---------------------------------------------------------------------------
    '  --���ܣ�ͬ�����¼����
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --      order_list[]
    '  --          order_id          N 1 ҽ��id
    '  --          send_no           N 1 ���ͺ�
    '  --          sign_type         N 1 ���ñ��¼�����ͣ�
    '  --                                  ˵����1-���������¼
    '  --                                        2-��� ����ҩƷͬ�����
    '  --                                        3-��� ��������ͬ�����
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  ---------------------------------------------------------------------------
    
    If strInput = "" Then zlSplitService_CisUpdateSyncState = True: Exit Function
    
    If strInput <> "" Then
        arrInput = Split(strInput, "|")
        
        For i = 0 To UBound(arrInput)
            
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("order_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("send_no", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("sign_type", intType, 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
    End If
    StrJson_In = IIf(StrJson_In = "", "", StrJson_In & ",") & """order_list"":[" & strJson_List & "]"
    
    '����
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_CisSvr_UpdateSyncState"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    
    zlSplitService_CisUpdateSyncState = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_UpdateExseInfo(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strItemList As String, ByVal strDeptchangeList As String, ByVal strDelNoList As String) As Boolean
    '��ҩ/��ҩ���÷��񣺸��·���ִ�в��š�ִ���ˡ���ҩ���ڼ�ִ��״̬����Ϣ
    'strBase��������Ϣ����ʽ ������Դ,����Ա����,����Ա����,����ʱ��
    'strItemList�����·���ID�б���ʽ ����ID,��ִ������(��ҩ����),ִ����,ִ��ʱ��,��ҩ����|...
    '             ��ʱ�⼸��ֵ����ҩƷ�����ж���ѯ�������루�򰴹��ܴӽ�������֯������ı䷢ҩ���ڵȣ�
    'strDeptchangeList������ִ�в��� ��ʽ ����id,ԭִ�в���id,��ִ�в���id|...
    'strDelNoList����ҩ�Զ����� ��ʽ NO;���(���1:����:ִ��״̬1,���2:����2:ִ��״̬2,...);����״̬|...
    Dim arrInput As Variant, strPart As String
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim strִ���� As String, strִ��ʱ�� As String
    
    On Error GoTo errHandle
    
    'Zl_Exsesvr_Updateexeinfo
'    ---------------------------------------------------------------------------
'  --���ܣ����·���ִ�в��š�ִ���ˡ���ҩ���ڼ�ִ��״̬����Ϣ
'  --��Σ�Json_In:��ʽ
'  --input
'  --  fee_origin            N  1  ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
'  --  operator_code         C     ����Ա����
'  --  operator_name         C     ����Ա����
'  --  operator_time         C     ����ʱ��
'  --  item_list                   ���б����ִ�������Ϣ�������б�ʱͬʱ��Ҫ����fee_origin
'  --    fee_id              C  1  ����id
'  --    exe_nums            N  1  ��ִ������:Ϊ0��ʾ��δִ��
'  --    exe_people          C     ִ����:����ִ�л���ȫִ��ʱ����Ҫ���룬������ʱ����operator_nameΪ׼
'  --    exe_time            D     ִ��ʱ��:yyyy-mm-dd hh24:mi:ss,:����ִ�л���ȫִ��ʱ����Ҫ���룬������ʱ����"create_time"Ϊ׼
'  --    pharmacy_window     C     ��ҩ����:ҩƷ��������Ч,�޴˽ӵ㣬������·�ҩ����
'  --  deptchange_List       C  1  ִ�п��ұ����Ϣ�б�
'  --    fee_id              C  1  ����id
'  --    exe_old_deptid      N     ԭִ�п���ID
'  --    exe_deptid          N  1  ִ�в���id
'  --  delrcp_list           C     [����]ȡ����Һʱ����Ҫͬ������
'  --    rcp_no              C  1  ����no
'  --    serial_nums         C  1  ��ʽ: ���1:����:ִ��״̬1,���2:����2:ִ��״̬2,...
'  --    operator_status     N     ����״̬��0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
'  --����: Json_Out,��ʽ����
'  --output
'  --   code                 C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --   message              C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  ---------------------------------------------------------------------------

    
    '1.������Ϣ
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_origin", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_code", Split(strBase, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_name", Split(strBase, ",")(2), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_time", Split(strBase, ",")(3), 0)
        
    '2.���·���id��Ӧ�ķ�����Ϣ
    strJson_List = ""
    If strItemList <> "" Then
        arrInput = Split(strItemList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_nums", Split(arrInput(i), ",")(1), 1)
            
            'ִ�������Ϊ���򲻴���ýڵ�
            If Split(arrInput(i), ",")(2) <> "" Then
                If Split(arrInput(i), ",")(2) = "null" Then
                    '����ģ�����������null������մ�������ȡ����ҩʱ���ִ����
                    strJson = strJson & "," & GetJsonNodeString("exe_people", "", 0)
                Else
                    strJson = strJson & "," & GetJsonNodeString("exe_people", Split(arrInput(i), ",")(2), 0)
                End If
            End If
            
            'ִ��ʱ�����Ϊ���򲻴���ýڵ�
            If Split(arrInput(i), ",")(3) <> "" Then
                strJson = strJson & "," & GetJsonNodeString("exe_time", Split(arrInput(i), ",")(3), 0)
            End If
            
            '��ҩ�������Ϊ���򲻴��ýڵ�
            If Split(arrInput(i), ",")(4) <> "" Then
                strJson = strJson & "," & GetJsonNodeString("pharmacy_window", Split(arrInput(i), ",")(4), 0)
            End If
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""item_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '3.�������·���id��Ӧ��ִ�в���id
    strJson_List = ""
    If strDeptchangeList <> "" Then
        arrInput = Split(strDeptchangeList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_old_deptid", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_deptid", Split(arrInput(i), ",")(2), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""deptchange_List"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '4.������������
    strJson_List = ""
    If strDelNoList <> "" Then
        arrInput = Split(strDelNoList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ";")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), ";")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), ";")(2), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""delrcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '����
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Updateexeinfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Updateexeinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_UpdateExseInfo = False
        Exit Function
    End If

    zlSplitService_UpdateExseInfo = True

    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallExseData(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strExeDept As String, ByVal strSendWin As String, ByVal strExePeople As String, _
    ByVal strExeSta As String, Optional ByVal strAccDel As String) As Boolean
    '��ҩ/��ҩ���÷��񣺸��·����ֶΣ�ִ�в��ţ���ҩ���ڣ�ִ����, ִ��״̬�����ݴ��θ��¶�Ӧ����
    Dim arrInput As Variant, strPart As String
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim strִ���� As String, strִ��ʱ�� As String
    
    On Error GoTo errHandle
    
    'Zl_ExseSvr_BillInforUpdate ͨ�÷���
    '1. ����ִ�в���
    '2. ���·�ҩ����
    '3. ����ִ����
    '4. ����ִ��״̬
    
    '1. ����ִ�в���
    '����ִ�в��ţ���ǰҩ��id,��������,no,ԭҩ��id,�����־,��������|...
    'strExeDept��306,1,T0000312,310,2,2018-03-28 13:05:22|...
    'bill_dept_list          [����]���·���ִ�в���������Ϣ
    '    pharmacy_id N   1   ��ǰҩ��id
    '    bill_type   C   1   ��������(1-�շѴ���,2-���˴���)
    '    rcp_no  C   1   NO
    '    pharmacy_id_old C   1   ԭҩ��id
    '    outpati_flag    N       �����־:1-���2-סԺ
    '    write_time  N       ��������
    
    strJson_List = ""
    If strExeDept <> "" Then
        arrInput = Split(strExeDept, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("pharmacy_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("bill_type", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_no", Split(arrInput(i), ",")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("pharmacy_id_old", Split(arrInput(i), ",")(3), 1)
            strJson = strJson & "," & GetJsonNodeString("outpati_flag", Split(arrInput(i), ",")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("write_time", Split(arrInput(i), ",")(5), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_dept_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '2. ���·�ҩ���ڣ�no,��������,ִ�в���id,��ҩ����|...
    'bill_win_list           [����]���·��÷�ҩ����������Ϣ
    '    rcp_no  C   1   NO
    '    fee_properties  N   1   ��¼����
    '    fee_exe_deptid  N   1   ִ�в���id
    '    pharmacy_window C   1   ��ҩ����
    
    strJson_List = ""
    If strSendWin <> "" Then
        arrInput = Split(strSendWin, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ",")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_properties", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_exe_deptid", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("pharmacy_window", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_win_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '3. ���·���ִ���ˣ�no,��������,ִ�в���id,ִ����|...
    'bill_people_list            [����]���·���ִ����������Ϣ
    '    rcp_no  C   1   NO
    '    bill_type   N   1   ��������(1-�շѣ�2-���ˣ�
    '    exe_deptid  N   1   ִ�в���id
    '    exe_people  C   1   ִ����
    
    strJson_List = ""
    If strExePeople <> "" Then
        arrInput = Split(strExePeople, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ",")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("bill_type", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_deptid", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_people", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_people_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '4. ���·���ִ��״̬��ִ����;��ҩʱ��||����ID��;ִ��״̬|����ID��;ִ��״̬
    'strExeSta������;2018-09-10 10:43:22||6341798,6341800,634180;1|6341818;2
    '    bill_status_list            [����]���·���ִ��״̬������Ϣ(ִ��״̬��ͬ��ƴ��һ������id��)
    '        detail_ids  C   1   ������ϸid��(����id��),֧�ֶ��id���á�,���ָ�
    '        exe_status  N   1   ִ��״̬(0-δִ��;1-��ȫִ��;2-����ִ��)
    '        exe_people  C   1   ִ����(��ҩ��������գ�ȫ��ִ���˻��Ϊnull,�����˲��޸�)
    '        give_time   C   1   ��ҩʱ��(��ҩ��������գ���ִ����ͬ��)
    
    strJson_List = ""
    If strExeSta <> "" Then
        strִ���� = Split(Split(strExeSta, "||")(0), ";")(0)
        strִ��ʱ�� = Split(Split(strExeSta, "||")(0), ";")(1)
        strExeSta = Split(strExeSta, "||")(1)
        
        If strExeSta <> "" Then
            '��֯���,detail_ids�ڵ�ֵ���ܳ�������Ҫ���ⲿ����strExeStaֵʱ�Ƚ��зֽ�
            arrInput = Split(strExeSta, "|")
            
            For i = 0 To UBound(arrInput)
                strJson = ""
                strJson = strJson & "" & GetJsonNodeString("detail_ids", Split(arrInput(i), ";")(0), 0)
                strJson = strJson & "," & GetJsonNodeString("exe_status", Split(arrInput(i), ";")(1), 1)
                strJson = strJson & "," & GetJsonNodeString("exe_people", strִ����, 0)
                strJson = strJson & "," & GetJsonNodeString("give_time", strִ��ʱ��, 0)
                
                If strJson_List = "" Then
                    strJson_List = "{" & strJson & "}"
                Else
                    strJson_List = strJson_List & "," & "{" & strJson & "}"
                End If
            Next
            
            strJson_List = ",""bill_status_list"":[" & strJson_List & "]"
        End If
    End If
    StrJson_In = StrJson_In & strJson_List

    '5. ����/סԺ��¼ɾ���� ����Ա����,����Ա���,����ʱ��||������Դ;��¼����;���õ��ݺ�;��Ŵ�;����״̬|������Դ;��¼����;���õ��ݺ�;��Ŵ�;����״̬...
    '  --  bill_rcp_list     [����]��ɾ�����۵��ĵ�����Ϣ
    '  --    rcp_no          N   1     ����no
    '  --    serial_nums     C   1     ��ʽ��"1,3,5,7,8",��"1:2:33456,3:2,5:2,7:2,8:2",ð��ǰ������ֱ�ʾ�к�,�м�����ֱ�ʾ�˵�����,��������ֱ�ʾ��ҩ��¼��ID,Ŀǰ�����������ʱ�Ŵ���
    '  --    operator_code   C   1     ����Ա����
    '  --    operator_name   C   1     ����Ա����
    '  --    fee_properties  N         ��¼����
    '  --    operator_status N         ����״̬��0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
    '  --    create_time     D         �Ǽ�ʱ��
    '  --    fee_origin      N         ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
    
    strJson_List = ""
    If strAccDel <> "" Then
        strPart = Split(strAccDel, "||")(0)
        arrInput = Split(Split(strAccDel, "||")(1), "|")
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_origin", Split(arrInput(i), ";")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_properties", Split(arrInput(i), ";")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_no", Split(arrInput(i), ";")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), ";")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), ";")(4), 1)
            
            strJson = strJson & "," & GetJsonNodeString("operator_code", Split(strPart, ",")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_name", Split(strPart, ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("create_time", Split(strPart, ",")(2), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_rcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    If StrJson_In = "" Then zlSplitService_CallExseData = True: Exit Function
    
    '����
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_ExseSvr_BillInforUpdate"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_ExseSvr_BillInforUpdate��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallExseData = False
        Exit Function
    End If

    zlSplitService_CallExseData = True

    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_CallRetunDrug(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strAccDel As String, ByVal strExeSta As String) As Boolean
    '��ҩͬʱ������۵�ʱ���õ�������������/סԺ��¼ɾ�������·��ü�¼״̬�ȹ���
    '�ϲ�Ϊһ�����������Ч���Ʒ�ҩ/��ҩ����
    'strAccDel��rcp_no|serial_nums|operator_code|operator_name|fee_properties|operator_status|create_time|fee_origin||...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    'Zl_ExseSvr_BillInforUpdate ͨ�÷���
    
    '1. ����/סԺ��¼ɾ���� no;�������(��ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������δ��˵���);����Ա����;����Ա����;��¼����;����״̬;�Ǽ�ʱ��;������Դ|...
'  --  bill_rcp_list     [����]��ɾ�����۵��ĵ�����Ϣ
'  --    rcp_no          N   1     ����no
'  --    serial_nums     C   1     ��ʽ��"1,3,5,7,8",��"1:2:33456,3:2,5:2,7:2,8:2",ð��ǰ������ֱ�ʾ�к�,�м�����ֱ�ʾ�˵�����,��������ֱ�ʾ��ҩ��¼��ID,Ŀǰ�����������ʱ�Ŵ���
'  --    operator_code   C   1     ����Ա����
'  --    operator_name   C   1     ����Ա����
'  --    fee_properties  N         ��¼����
'  --    operator_status N         ����״̬��0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
'  --    create_time     D         �Ǽ�ʱ��
'  --    fee_origin      N         ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
    
    strJson_List = ""
    If strAccDel <> "" Then
        arrInput = Split(strAccDel, "||")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), "|")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), "|")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_code", Split(arrInput(i), "|")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_name", Split(arrInput(i), "|")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_properties", Split(arrInput(i), "|")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), "|")(5), 1)
            strJson = strJson & "," & GetJsonNodeString("create_time", Split(arrInput(i), "|")(6), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_origin", Split(arrInput(i), "|")(7), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_rcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '2. ���·���ִ��״̬������ID��;ִ��״̬|����ID��;ִ��״̬...
'  --  bill_status_list   [����]���·���ִ��״̬������Ϣ(ִ��״̬��ͬ��ƴ��һ������id��)
'  --    detail_ids      C   1     ������ϸid��(����id��),֧�ֶ��id���á�,���ָ�
'  --    exe_status      N   1     ִ��״̬
'  --    exe_people      N   1     ִ����
'  --    give_time       C   1     ��ҩʱ��
    
    strJson_List = ""
    If strExeSta <> "" Then
        strExeSta = Split(strExeSta, "||")(1)
        
        If strExeSta <> "" Then
            '��֯���,detail_ids�ڵ�ֵ���ܳ�������Ҫ���ⲿ����strExeStaֵʱ�Ƚ��зֽ�
            arrInput = Split(strExeSta, "|")
            
            For i = 0 To UBound(arrInput)
                strJson = ""
                strJson = strJson & "" & GetJsonNodeString("detail_ids", Split(arrInput(i), ";")(0), 0)
                strJson = strJson & "," & GetJsonNodeString("exe_status", Split(arrInput(i), ";")(1), 1)
                 
                If strJson_List = "" Then
                    strJson_List = "{" & strJson & "}"
                Else
                    strJson_List = strJson_List & "," & "{" & strJson & "}"
                End If
            Next
            
            strJson_List = ",""bill_status_list"":[" & strJson_List & "]"
        End If
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '����
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_ExseSvr_BillInforUpdate"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_ExseSvr_BillInforUpdate��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallRetunDrug = False
        Exit Function
    End If
    
    zlSplitService_CallRetunDrug = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetJsonNodeString(ByVal strNodeName As String, ByVal strValue As String, Optional ByVal intType As Integer, _
    Optional ByVal blnZeroToNull As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡJson�ӵ㴮
    '���:strNodeName-�ӵ���
    '     strValue-ֵ
    '     intType-����:0-�ַ�;1-����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-09 18:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    
    strJson = Chr(34) & strNodeName & Chr(34)
    
    If intType = 0 Then
        strJson = strJson & ":" & Chr(34) & zlStr.ToJsonStr(strValue) & Chr(34)
    Else
        If strValue = "" Or (blnZeroToNull And Val(strValue) = 0) Then
            strJson = strJson & ":null"
        Else
            strJson = strJson & ":" & IIf(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
        
    GetJsonNodeString = strJson
End Function


Public Function GetJsonNodeToNull(ByVal strNodeName As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ڵ�ֵ��Ϊnull��ֻ���������
    '���:strNodeName-�ӵ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    
    strJson = Chr(34) & strNodeName & Chr(34)
    strJson = strJson & ":null"
    
    GetJsonNodeToNull = strJson
End Function

Public Function zlSplitService_CallIsCloseAcc(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef inState As Integer) As Boolean
    '��ҩ/��ҩ���÷��񣺲�ѯ�Ƿ��ѽ���
    'strInput��������Դ|No
    'inState�����ؽ���״̬
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    'Zl_Exsesvr_Getfeebalancestate
'  ---------------------------------------------------------------------------
'  --����:���ݵ��ݺ���Ϣ����ȡ���ݶ�Ӧ�Ľ���״̬
'  --��Σ�Json_In:��ʽ
'  --input
'  --    query_mode  N 1 ��ѯ��ʽ:0-�������;1-סԺ����
'  --    bill_nos  C 1 ���ݺ�
'  --����: Json_Out,��ʽ����
'  -- output
'  --    code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    state N 1 ״̬:-1-�����ڼ��ʵ���;0-δ����;1-���ֽ���;2-ȫ������
'
'  ---------------------------------------------------------------------------
    
    If strInput = "" Then Exit Function

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_mode", Split(strInput, "|")(0), 1)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", Split(strInput, "|")(1), 0)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "Zl_Exsesvr_Getfeebalancestate"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Getfeebalancestate��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallIsCloseAcc = False
        Exit Function
    Else
        inState = Val(objServiceCall.GetJsonNodeValue("output.state"))
    End If

    zlSplitService_CallIsCloseAcc = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CheckExeItemValied(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef strOutPut As String) As Boolean
    '�����ƺ���㷽ʽ�����ִ����Ŀ�ĺϷ���
    'strInput������id|�Һ�id|�շ����
    'strOutPut������־|��Ϣ����
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    '����Zl_Exsesvr_CheckExeItemValied
'  -------------------------------------------------------------------------------------------------
'  --���ܣ������ƺ���㷽ʽ�����ִ����Ŀ�ĺϷ���
'  --input
'  --  pati_id      N   1   ����id
'  --  register_id   N   1   �Һ�id
'  --  receipt_type  C   1   �շ����
'  --output
'  --  code          C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message       C   1   Ӧ����Ϣ��
'  --  check_flag   N   0   ����־��0-������Ϸ���1-���� ��2-�ܾ�
'  --  check_msg    C   0   ���ѻ�ܾ���������ʾ
'  -------------------------------------------------------------------------------------------------
    
    If strInput = "" Then Exit Function

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", Split(strInput, "|")(0), 1)
    strJson = strJson & "," & GetJsonNodeString("register_id", Split(strInput, "|")(1), 1)
    strJson = strJson & "," & GetJsonNodeString("receipt_type", Split(strInput, "|")(2), 0)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "Zl_Exsesvr_CheckExeItemValied"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�zlSplitService_CheckExeItemValied��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CheckExeItemValied = False
        Exit Function
    Else
        strOutPut = objServiceCall.GetJsonNodeValue("output.check_flag") & "|" & objServiceCall.GetJsonNodeValue("output.check_msg")
    End If

    zlSplitService_CheckExeItemValied = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_CallSetWindows(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String) As Boolean
    '��ҩ���ڵ������÷������÷���ҩƷ���ݵķ�ҩ����
    'strInput���ⷿid,�ɴ���,�´���|��������,no;��������,no...
    'dblSumAmount�����ؽ��ʽ��
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim strWins As String, strNOs As String
    
    On Error GoTo errHandle
    
    '����Zl_Exsesvr_Setsendwin
'  --���ܣ����÷���ҩƷ���ݵķ�ҩ����
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    pharmacy_id              N   1  �ⷿid
'  --    pharmacy_window_old      C   1  �ɷ�ҩ����
'  --    pharmacy_window_new      C   1  �·�ҩ����
'  --    bill_list[]
'  --      billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
'  --      rcp_no                 C   1 ����No
'  --����: Json_Out,��ʽ����
'  --  output
'  --     code                   N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --     message                C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    
    If strInput = "" Then Exit Function
    
    '�����б�
    strWins = Split(strInput, "|")(0)
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pharmacy_id", Split(strWins, ",")(0), 1)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_window_old", Split(strWins, ",")(1), 0)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_window_new", Split(strWins, ",")(2), 0)
    
    '�����б�
    strNOs = Split(strInput, "|")(1)
    arrInput = Split(strNOs, ";")
    
    For i = 0 To UBound(arrInput)
        StrJson_In = ""
        StrJson_In = StrJson_In & "" & GetJsonNodeString("billtype", Split(arrInput(i), ",")(0), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("rcp_no", Split(arrInput(i), ",")(1), 0)
            
        If strJson_List = "" Then
            strJson_List = "{" & StrJson_In & "}"
        Else
            strJson_List = strJson_List & "," & "{" & StrJson_In & "}"
        End If
    Next
    
    strJson_List = """bill_list"":[" & strJson_List & "]"

    StrJson_In = "{""input"":{" & strJson & "," & strJson_List & "}}"
    
    strService = "Zl_Exsesvr_Setsendwin"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Setsendwin��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallSetWindows = False
        Exit Function
    End If

    zlSplitService_CallSetWindows = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallPatiIsOut(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal lngPaitId As Long, ByVal lngPageId As Long, ByRef intOutSign As Integer) As Boolean
    '��ҩ/��ҩ���÷��񣺲�ѯ�����Ƿ��ѳ�Ժ
    '������Ϣ������id����ҳid
    'intOutSign��0-δ��Ժ��1-��Ժ
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    
    On Error GoTo errHandle
    If lngPaitId = 0 Then Exit Function
    
    'Zl_Cissvr_Patiisout
    '  --���ܣ���ѯ�����Ƿ��Ѿ���Ժ
    '  --input      ��ѯ�����Ƿ��Ѿ���Ժ
    '  --  pati_id               N  1  ����id
    '  --  pati_pageid           N  1  ��ҳid
    '  --  query_type            N  1  ��ѯ���ͣ�0-�������˲�ѯ��1-�������������ѯ
    '  --  pati_pageids          C  1  ��ʽ������ID:��ҳID,����ID:��ҳID,...
    '  --output
    '  --  code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --  message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --  pati_outsign          N  1  ��Ժ��ǣ�0-δ��Ժ��1-��Ժ��query_type=0ʱ����
    '  --  item_list[]           ��Ժ����б�query_type=1ʱ����
    '  --    pati_id             N  1  ����id
    '  --    pati_outsign        N  1  ��Ժ��ǣ�0-δ��Ժ��1-��Ժ
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lngPaitId, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lngPageId, 1)
    StrJson_In = "{""input"":{" & strJson & "}}"
    
    '���÷���
    If objServiceCall.CallService("zl_CisSvr_PatiIsOut", StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�zl_CisSvr_PatiIsOut��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    If strJson_Out = "0" Then
        zlSplitService_CallPatiIsOut = False
        Exit Function
    Else
        intOutSign = Val(objServiceCall.GetJsonNodeValue("output.pati_outsign"))
    End If

    zlSplitService_CallPatiIsOut = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallPatiIsOutByBatch(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strPaitPages As String, ByRef cllPatiIsOut As Collection) As Boolean
    '������ѯ�����Ƿ��ѳ�Ժ
    '���:
    '   strPaitPages=������Ϣ������id:��ҳid,����id:��ҳid,...
    '����:
    '   cllPatiIsOut=���˳�Ժ���(����ID,��Ժ��־)�����У���Ժ��־��0-δ��Ժ��1-��Ժ
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    
    On Error GoTo errHandle
    If strPaitPages = "" Then Exit Function
    
    Set cllPatiIsOut = New Collection
    'Zl_Cissvr_Patiisout
    '  --���ܣ���ѯ�����Ƿ��Ѿ���Ժ
    '  --input      ��ѯ�����Ƿ��Ѿ���Ժ
    '  --  pati_id               N  1  ����id
    '  --  pati_pageid           N  1  ��ҳid
    '  --  query_type            N  1  ��ѯ���ͣ�0-�������˲�ѯ��1-�������������ѯ
    '  --  pati_pageids          C  1  ��ʽ������ID:��ҳID,����ID:��ҳID,...
    '  --output
    '  --  code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --  message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --  pati_outsign          N  1  ��Ժ��ǣ�0-δ��Ժ��1-��Ժ��query_type=0ʱ����
    '  --  item_list[]           ��Ժ����б�query_type=1ʱ����
    '  --    pati_id             N  1  ����id
    '  --    pati_outsign        N  1  ��Ժ��ǣ�0-δ��Ժ��1-��Ժ
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 1, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", strPaitPages, 0)
    StrJson_In = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_CisSvr_PatiIsOut", StrJson_In, , "", lngMode, False) = False Then
        MsgBox "���á�zl_CisSvr_PatiIsOut��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    If strJson_Out = "0" Then
        MsgBox objServiceCall.GetJsonNodeValue("output.message"), vbInformation, gstrSysName
        zlSplitService_CallPatiIsOutByBatch = False: Exit Function
    End If
    
    Set cllPatiIsOut = objServiceCall.GetJsonListValue("output.item_list", "pati_id")
    If cllPatiIsOut Is Nothing Then Exit Function

    zlSplitService_CallPatiIsOutByBatch = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CheckExistNoSendStuff(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal lng����id As Long, ByVal lng�ⷿID As Long, ByRef blnExist As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ĳ������ָ���ⷿ�Ƿ����δ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim StrJson_In As String
    
    On Error GoTo ErrHandler
    'Zl_Stuffsvr_Checkpatiexecute
    '  --���ܣ����ݲ�����Ϣ���ȡδ��������
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --     check_type         N 1 ��鷽ʽ:0-�����²���ֵ���м�飻1-��������ID�ͷ��Ͽⷿ���м��
    '  --     pati_id            N 1 ����ID
    '  --     pati_pageid        N 1 ��ҳID
    '  --     baby_num           N 1 Ӥ�����:-1��ʾ������;0-ĸ�׵�;>0����Ӥ������
    '  --     fee_source         N 1 ������Դ:1-����;2-סԺ;4-���
    '  --     stuff_nos             �������ݺţ������磺["A0001","A0002"]
    '  --     warehouse_id       N 0 �ⷿID
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code                    N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    isexist                 N 1 �Ƿ����: 1-����;0-������
    '  --    stuff_notsend_infor     C 1 δ������Ϣ,isexist=1ʱ����
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("check_type", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_id", lng����id, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("warehouse_id", lng�ⷿID, 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Stuffsvr_Checkpatiexecute", StrJson_In, "", "", lngMode) = False Then Exit Function
    
    blnExist = Val(nvl(objServiceCall.GetJsonNodeValue("output.isexist"))) = 1
    
    zlSplitService_CheckExistNoSendStuff = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetRequestCancel(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection) As Boolean
    'ȡ���������¼
    'strInput������id,����id...
    Dim StrJson_In As String
    
    'Zl_Exsesvr_Getrequestcancel
    '  --��ѯ���������¼
    '  --���      json
    '  --  input      ��ѯ�Ƿ�������������¼
    '  --    query_type          N 1 ��ѯ��ʽ:0-���ݷ���ID��ѯ
    '  --    rcpdtl_id           C 0 ������ϸid,[����]��[1,2,3]
    '  --    request_type        N 0 �������
    '  --    cancel_status       N 1 ����״̬
    '  --����      json
    '  -- output
    '  --   code     C  1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --   message  C  1   Ӧ����Ϣ��
    '  --   fee_cancel_list      [����]����������ÿ���������ʼ�¼
    '  --     rcpdtl_id          N    ������ϸid(����id)
    '  --     apply_type         N    �������:��ҩƷ��������Ч:0-δִ��;1-��ִ��;��ҩƷ�����Ĺ̶���Ϊ0
    '  --     apply_time         N    ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '  --     aplnt_name         N    ������
    '  --     apply_dept_id      N    ���벿��id
    '  --     apply_dept_name    N    ���벿������
    '  --     audit_dept_id      N    ��˲���id;
    '  --     audit_dept_name    N    ��˲�������
    '  --     bill_no            N    ���õ��ݺ�
    '  --     item_id            N    �շ�ϸĿid
    '  --     item_name          N    �շ���Ŀ����
    '  --     quantity           N    ����
    
    On Error GoTo ErrHandler
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 0, 1)
    StrJson_In = StrJson_In & ",""rcpdtl_id"":" & "[" & strInput & "]"
    StrJson_In = StrJson_In & "," & GetJsonNodeString("request_type", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_status", 0, 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getrequestcancel", StrJson_In, "", "", lngMode) = False Then Exit Function
    Set colOutlist = objServiceCall.GetJsonListValue("output.fee_cancel_list")
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetRequestCancel = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetExseBillByTime(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal colQueryCons As Collection, ByRef strBill_Out As String, ByRef strErrMsg_Out As String) As Boolean
    '��ʱ�䷶Χ��ȡ���õ���
    '��Σ�
    '   colQueryCons = ��ѯ��������Ա(Key)����ѯ��ʽ,������Դ,��ʼʱ��,����ʱ��,ִ�в���IDS,����ִ�в���IDS
    '                           ���У�������Դ��0-������;1-����;2-סԺ
    '����:
    '   strBill_Out = ������Ϣ����ʽ������1:NO,����1:NO,...�����У����ݣ�8-�շѴ�����ҩ��9-���ʵ�������ҩ��10-���ʱ�����ҩ
    Dim StrJson_In As String
    Dim colOutlist As Collection, colTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    strBill_Out = "": strErrMsg_Out = ""
    'Zl_Exsesvr_Getbillbytime
    '  --���ܣ���ʱ�䷶Χ��ȡ���õ���
    '  --��Σ�json��ʽ
    '  --  input
    '  --    query_type          N 0 ��ѯ��ʽ:0-��ȡҩƷҽ�����õ���
    '  --    fee_source          N 1 ������Դ:0-������;1-����;2-סԺ
    '  --    start_time          C 1 ��ʼʱ�䣬��ʽ��yyyy-mm-dd hh24:mi:ss
    '  --    end_time            C 1 ����ʱ�䣬��ʽ��yyyy-mm-dd hh24:mi:ss
    '  --    exe_deptids         C 0 ִ�в���ID������ö�Ӣ�ĺŷָ�
    '  --    excp_exe_deptids    C 0 ��������ִ�в���ID������ö�Ӣ�ĺŷָ�
    '  --���Σ�json��ʽ
    '  --  output
    '  --    code                C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message             C 1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    bill_nos            C 1 ������Ϣ:��ʽ����������1:NO1,��������2:NO2,...
    '  --                            ���У���������: 1-�շѴ���;2-���ʵ�����;3-���ʱ���
     
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", colQueryCons("��ѯ��ʽ"), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_source", colQueryCons("������Դ"), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("start_time", Format(colQueryCons("��ʼʱ��"), "yyyy-MM-dd HH:mm:ss"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("end_time", Format(colQueryCons("����ʱ��"), "yyyy-MM-dd HH:mm:ss"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("exe_deptids", colQueryCons("ִ�в���IDS"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("excp_exe_deptids", colQueryCons("����ִ�в���IDS"), 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getbillbytime", StrJson_In, "", "", lngMode, False, , , , True) = False Then Exit Function
    
    strBill_Out = objServiceCall.GetJsonNodeValue("output.bill_nos")
    
    If strBill_Out <> "" Then
        '��������ת����8-�շѴ�����ҩ��9-���ʵ�������ҩ��10-���ʱ�����ҩ
        strBill_Out = "," & strBill_Out
        strBill_Out = Replace(strBill_Out, ",1:", ",8:")
        strBill_Out = Replace(strBill_Out, ",2:", ",9:")
        strBill_Out = Replace(strBill_Out, ",3:", ",10:")
        strBill_Out = zlStr.TrimEx(strBill_Out, ",")
    End If
    
    zlSplitService_GetExseBillByTime = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

