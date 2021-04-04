Attribute VB_Name = "mdlPubCardSquare"
Option Explicit
Public grs���ѿ��ӿ� As ADODB.Recordset

Public Function GetConsumerCardTypes(ByRef rsSquareType_Out As ADODB.Recordset, Optional ByVal blnOnlyStart As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���»�ȡ���ѿ�Ŀ¼���ݼ�
    '���:
    '����:rsSquareType_Out-�������ѿ����ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-21 15:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long
    Dim cllData As Collection, cllTemp As Collection

    On Error GoTo errHandle
    If zlGetCardTypeRecStru(rsSquareType_Out) = False Then
        Set grs���ѿ��ӿ� = Nothing: Exit Function
    End If
        
    If zl_ExseSvr_GetConsumerCardType(cllData, blnOnlyStart) = False Then Set grs���ѿ��ӿ� = Nothing: Exit Function
    '    type_list[] C   1   ֧�ֵĿ�����б�
    '        cardtype_id   N   1   id
    '        cardtype_num  N   1   ���
    '        cardtype_name C   1   ����
    '        cardtype_stname   C   1   ����
    '        prefix_text C   1   ǰ׺�ı�
    '        cardno_len  N   1   ���ų���
    '        default    N   1   ȱʡ��־
    '        fixed N   1   �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
    '        strict   N   1   �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
    '        self_make N   1   �Ƿ�����:1-�ǵ�;0-����
    '        allow_return_cash    N   1   �Ƿ�����:1-����;0-������
    '        must_all_return   N   1   �Ƿ�ȫ��:1-����ȫ��;0-��������
    '        specpati N   1   �ض�����
    '        component   C   1   ����
    '        memo    C   1   ��ע
    '        blnc_mode   C   1   ���㷽ʽ
    '        blnc_nature N   1   ��������
    '        pwdtxt   N   1   �Ƿ�����
    '        enabled    N   1   �Ƿ�����:1-������;0-δ����
    '        pwd_len N   1   ���볤��
    '        pwd_len_limit   N   1   ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
    '        pwd_rule    N   1   �������:��-���ֺ��ַ����;1-��Ϊ�������
    '        readcard_nature C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
    '        keyboard_mode   N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
    '        def_return_cash N   1   �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
                
 
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        With rsSquareType_Out
            .AddNew
                !id = cllTemp("_cardtype_id")
                !���� = cllTemp("_cardtype_num")
                !���� = cllTemp("_cardtype_name")
                !���� = cllTemp("_cardtype_stname")
                
                !ǰ׺�ı� = cllTemp("_prefix_text")
                !���ų��� = cllTemp("_cardno_len")
                !ȱʡ��־ = cllTemp("_default")
                !�Ƿ�̶� = cllTemp("_fixed")
                
                !�Ƿ��ϸ���� = cllTemp("_strict")
                !�Ƿ����� = cllTemp("_self_make")
                '!�Ƿ�����ʻ� = cllTemp("_exist_account")
                !�Ƿ����� = cllTemp("_allow_return_cash")
                !�Ƿ�ȱʡ���� = cllTemp("_def_return_cash")
                !�Ƿ�ȫ�� = cllTemp("_must_all_return")
                !���� = cllTemp("_component")
                !��ע = cllTemp("_memo")
                
                '!�ض���Ŀ = cllTemp("_spec_item")
                !���㷽ʽ = cllTemp("_blnc_mode")
                !�������� = cllTemp("_blnc_nature")
                If Val(cllTemp("_pwdtxt")) = 1 Then
                    !�������� = 1
                End If
                
               ' !�Ƿ��ظ�ʹ�� = cllTemp("_allow_repeat_use")
                !�Ƿ����� = cllTemp("_enabled")
                
                !���볤�� = cllTemp("_pwd_len")
                !���볤������ = cllTemp("_pwd_len_limit")
                !������� = cllTemp("_pwd_rule")
                '!������������ = cllTemp("_pwd_require")
                '!�Ƿ�ģ������ = cllTemp("_allow_vaguefind")
                '!�Ƿ�ȱʡ���� = cllTemp("_default_pwd")
                '!�Ƿ��ƿ� = cllTemp("_allow_makecard")
                '!�Ƿ񷢿� = cllTemp("_allow_sendcard")
                '!�Ƿ�д�� = cllTemp("_allow_writecard")
                '!�������� = cllTemp("_sendcard_sign")
                '!���� = cllTemp("_insurance_type")
                '!�������� = cllTemp("_sendcard_nature")
                '!�Ƿ�ת�ʼ����� = cllTemp("_transfer")
                
                !�������� = cllTemp("_readcard_nature")
                !���̿��Ʒ�ʽ = cllTemp("_keyboard_mode")
                
                '!�Ƿ�ֿ����� = cllTemp("_holding_pay")
                '!�Ƿ�֤�� = cllTemp("_cert_cardtype")
                '
                '!���͵��ýӿ� = cllTemp("_advsend_buildqrcode")
                '!�Ƿ��˿��鿨 = cllTemp("_verfycard")
                '
                '!�Ƿ�������� = cllTemp("_balalone")
                '!ȱʡ��Чʱ�� = cllTemp("_def_valid_time")
                '!����ʶ����� = cllTemp("_discern_rule")
                '!�Ƿ�֧��ɨ�븶 = cllTemp("_scanpay")
                '!�Ƿ����ûس� = cllTemp("_enterkey_enabled")
            .Update
       End With
    Next
    If rsSquareType_Out.RecordCount <> 0 Then rsSquareType_Out.MoveFirst
    GetConsumerCardTypes = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGet���ѿ��ӿ�(Optional ByVal blnOnlyStart As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ѿ��ӿ�
    '���:blnOnlyStart-�Ƿ����ȡ���õ����ѿ�
    '����:���˺�
    '����:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
      
    If Not grs���ѿ��ӿ� Is Nothing Then
        If grs���ѿ��ӿ�.State = 1 Then
            grs���ѿ��ӿ�.Filter = 0
            If grs���ѿ��ӿ�.RecordCount = 0 Then grs���ѿ��ӿ�.MoveFirst
            Set zlGet���ѿ��ӿ� = grs���ѿ��ӿ�
            Exit Function
        End If
    End If
    If GetConsumerCardTypes(grs���ѿ��ӿ�, blnOnlyStart) = False Then Set grs���ѿ��ӿ� = Nothing: Exit Function
    Set zlGet���ѿ��ӿ� = grs���ѿ��ӿ�
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

 

